use chrono::{ Local, NaiveTime };
use std::sync::Arc;
use tokio::sync::mpsc;
use tokio_util::sync::CancellationToken;

use crate::api_client::ApiClient;
use crate::broadcast::WorkerEvent;
use crate::config::AppConfig;
use crate::heartbeat;
use crate::job_runner;
use crate::state::{ AppState, SchedulerPhase, TrayCommand };


/// Run the main scheduler loop.
///
/// This is the core state machine that drives the worker:
/// `OFF_SHIFT → IDLE → POLLING → EXECUTING → COOLDOWN → IDLE`
///
/// @param state - shared application state
/// @param api - HTTP client for the Express server
/// @param shutdown - token to signal graceful shutdown
/// @param tray_rx - receiver for commands from the tray UI
pub async fn run(
    state: Arc<AppState>,
    api: Arc<ApiClient>,
    shutdown: CancellationToken,
    mut tray_rx: mpsc::Receiver<TrayCommand>,
) {
    tracing::info!( "Scheduler started" );

    let mut paused = false;
    let mut run_once = false;
    let mut run_job_id: Option<u32> = None;

    loop {
        // Check for shutdown
        if shutdown.is_cancelled() {
            tracing::info!( "Scheduler shutting down" );
            break;
        }

        // Drain tray commands (non-blocking)
        while let Ok( cmd ) = tray_rx.try_recv() {
            if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await {
                break;
            }
        }

        if shutdown.is_cancelled() {
            break;
        }

        // If paused, sleep until unpaused or shutdown
        if paused {
            tokio::select! {
                _ = tokio::time::sleep( std::time::Duration::from_millis( 500 ) ) => {}
                _ = shutdown.cancelled() => { break; }
                Some( cmd ) = tray_rx.recv() => {
                    if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await { break; }
                }
            }
            continue;
        }

        // Handle specific job claim (bypasses shift check, type filtering, etc.)
        if let Some( job_id ) = run_job_id.take() {
            let config = state.config.read().await.clone();
            set_phase( &state, SchedulerPhase::Polling ).await;

            match api.claim_specific_job( job_id, &config.worker.id ).await {
                Ok( Some( job ) ) => {
                    tracing::info!( "Claimed specific job #{} ({})", job.id, job.job_type );
                    set_connected( &state, true ).await;
                    execute_job( &job, &config, &state, &api, &shutdown ).await;
                }
                Ok( None ) => {
                    tracing::warn!( "Job #{} is not pending (already claimed or completed)", job_id );
                    set_connected( &state, true ).await;
                }
                Err( e ) => {
                    tracing::warn!( "Failed to claim job #{}: {}", job_id, e );
                    set_connected( &state, false ).await;
                }
            }
            continue;
        }

        // Read current config
        let config = state.config.read().await.clone();

        // Check if we're within the shift window (run_once and force_on_shift bypass shift check)
        let forced = *state.force_on_shift.read().await;
        if !run_once && !forced && ( !config.schedule.enabled || !is_within_shift( &config ) ) {
            set_phase( &state, SchedulerPhase::OffShift ).await;
            tokio::select! {
                _ = tokio::time::sleep( std::time::Duration::from_secs( 30 ) ) => {}
                _ = shutdown.cancelled() => { break; }
                Some( cmd ) = tray_rx.recv() => {
                    if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await { break; }
                }
            }
            continue;
        }

        // We're in shift — set to idle, then poll
        set_phase( &state, SchedulerPhase::Idle ).await;

        // Wait poll interval before checking for jobs (skip wait on run_once)
        // SSE job:created events wake us early via poll_notify
        if !run_once {
            tokio::select! {
                _ = tokio::time::sleep( std::time::Duration::from_secs( config.timing.poll_interval_secs ) ) => {}
                _ = state.poll_notify.notified() => {
                    tracing::info!( "Woken by SSE job event, polling immediately" );
                }
                _ = shutdown.cancelled() => { break; }
                Some( cmd ) = tray_rx.recv() => {
                    if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await { break; }
                    continue;
                }
            }
        }

        if shutdown.is_cancelled() {
            break;
        }

        // Filter types: only claim types that are both enabled AND not paused
        let paused_types = state.paused_types.read().await.clone();
        let effective_types: Vec<String> = config.worker.types.iter()
            .filter( |t| !paused_types.contains( t ) )
            .cloned()
            .collect();

        if effective_types.is_empty() {
            tracing::debug!( "All enabled types are paused, skipping poll" );
            if run_once {
                tracing::info!( "Run-once cycle complete (no active types)" );
                run_once = false;
            }
            continue;
        }

        // Poll for a job
        set_phase( &state, SchedulerPhase::Polling ).await;
        set_connected( &state, true ).await;

        let claim_result = api.claim_job( &config.worker.id, &effective_types ).await;

        match claim_result {
            Ok( None ) => {
                // No jobs available
                if run_once {
                    tracing::info!( "Run-once: queue empty" );
                    run_once = false;
                } else {
                    tracing::debug!( "No pending jobs" );
                }
            }
            Ok( Some( job ) ) => {
                tracing::info!( "Claimed job #{} ({}) for {}", job.id, job.job_type, job.sunday_date );
                execute_job( &job, &config, &state, &api, &shutdown ).await;

                // Cooldown before next poll
                set_phase( &state, SchedulerPhase::Cooldown ).await;
                tokio::select! {
                    _ = tokio::time::sleep( std::time::Duration::from_secs( config.timing.job_delay_secs ) ) => {}
                    _ = shutdown.cancelled() => { break; }
                    Some( cmd ) = tray_rx.recv() => {
                        if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await { break; }
                    }
                }
            }
            Err( e ) => {
                tracing::warn!( "Failed to poll for jobs: {}", e );
                set_connected( &state, false ).await;
                if run_once {
                    run_once = false;
                }
                // Back off a bit on network errors
                tokio::select! {
                    _ = tokio::time::sleep( std::time::Duration::from_secs( 10 ) ) => {}
                    _ = shutdown.cancelled() => { break; }
                    Some( cmd ) = tray_rx.recv() => {
                        if handle_command( cmd, &mut paused, &mut run_once, &mut run_job_id, &state, &shutdown ).await { break; }
                    }
                }
            }
        }

        // Clear run_once flag after one poll cycle
        if run_once {
            tracing::info!( "Run-once cycle complete" );
            run_once = false;
        }
    }

    set_phase( &state, SchedulerPhase::OffShift ).await;
    tracing::info!( "Scheduler stopped" );
}


/// Handle a single tray command.
///
/// Returns true if the scheduler should shut down.
///
/// @param cmd - the command to handle
/// @param paused - mutable pause state
/// @param run_once - mutable run-once flag
/// @param state - shared application state
/// @param shutdown - cancellation token
async fn handle_command(
    cmd: TrayCommand,
    paused: &mut bool,
    run_once: &mut bool,
    run_job_id: &mut Option<u32>,
    state: &AppState,
    shutdown: &CancellationToken,
) -> bool {
    match cmd {
        TrayCommand::TogglePause => {
            *paused = !*paused;
            if *paused {
                tracing::info!( "Scheduler paused by user" );
                set_phase( state, SchedulerPhase::Paused ).await;
            } else {
                tracing::info!( "Scheduler resumed by user" );
            }
        }
        TrayCommand::Quit => {
            tracing::info!( "Quit command received" );
            shutdown.cancel();
            return true;
        }
        TrayCommand::OpenConfig => {
            let config = state.config.read().await;
            let url = format!( "http://localhost:{}", config.web_ui.port );
            tracing::info!( "Opening config UI: {}", url );
            let _ = open::that( &url );
        }
        TrayCommand::ReloadConfig => {
            match crate::config::load_config() {
                Ok( new_config ) => {
                    *state.config.write().await = new_config;
                    tracing::info!( "Configuration reloaded from disk" );
                    state.emit( WorkerEvent::ConfigReloaded );
                }
                Err( e ) => {
                    tracing::error!( "Failed to reload config: {}", e );
                }
            }
        }
        TrayCommand::RunOnce => {
            tracing::info!( "Run-once requested" );
            *run_once = true;
        }
        TrayCommand::ToggleForceOnShift => {
            let mut forced = state.force_on_shift.write().await;
            *forced = !*forced;
            tracing::info!( "Force on-shift: {}", *forced );
            state.emit( WorkerEvent::ForceOnShift { enabled: *forced } );
        }
        TrayCommand::RunJob( job_id ) => {
            tracing::info!( "Run-job requested: #{}", job_id );
            *run_job_id = Some( job_id );
        }
        TrayCommand::TogglePauseJobType( ref job_type ) => {
            let mut paused_types = state.paused_types.write().await;
            let is_paused = if paused_types.contains( job_type ) {
                paused_types.retain( |t| t != job_type );
                tracing::info!( "Unpaused job type: {}", job_type );
                false
            } else {
                paused_types.push( job_type.clone() );
                tracing::info!( "Paused job type: {}", job_type );
                true
            };
            state.emit( WorkerEvent::TypePauseChange {
                job_type: job_type.clone(),
                paused: is_paused,
            } );
        }
    }
    false
}


/// Update the connection state and emit a broadcast event.
async fn set_connected( state: &AppState, connected: bool ) {
    let current = *state.connected.read().await;
    if current != connected {
        *state.connected.write().await = connected;
        state.emit( WorkerEvent::ConnectionChange { connected } );
    }
}


/// Emit a stats update event.
async fn emit_stats( state: &AppState ) {
    let completed = *state.jobs_completed.read().await;
    let failed = *state.jobs_failed.read().await;
    state.emit( WorkerEvent::StatsUpdate {
        jobs_completed: completed,
        jobs_failed: failed,
    } );
}


/// Check if the current time falls within the configured shift window.
///
/// Supports overnight shifts where `shift_start > shift_end`
/// (e.g., 22:00–06:00 means active from 10 PM to 6 AM).
fn is_within_shift( config: &AppConfig ) -> bool {
    let now = Local::now().time();

    let start = match NaiveTime::parse_from_str( &config.schedule.shift_start, "%H:%M" ) {
        Ok( t ) => t,
        Err( _ ) => {
            tracing::error!( "Invalid shift_start time: {}", config.schedule.shift_start );
            return false;
        }
    };

    let end = match NaiveTime::parse_from_str( &config.schedule.shift_end, "%H:%M" ) {
        Ok( t ) => t,
        Err( _ ) => {
            tracing::error!( "Invalid shift_end time: {}", config.schedule.shift_end );
            return false;
        }
    };

    if start <= end {
        // Normal window: e.g., 02:00–06:00
        now >= start && now < end
    } else {
        // Overnight window: e.g., 22:00–06:00
        now >= start || now < end
    }
}


/// Execute a claimed job: heartbeat, run, report result, clear state.
async fn execute_job(
    job: &crate::api_client::Job,
    config: &crate::config::AppConfig,
    state: &AppState,
    api: &Arc<ApiClient>,
    shutdown: &CancellationToken,
) {
    *state.current_job_id.write().await = Some( job.id );
    *state.current_job_type.write().await = Some( job.job_type.clone() );
    set_phase( state, SchedulerPhase::Executing ).await;
    state.emit( WorkerEvent::JobClaimed {
        job_id: job.id,
        job_type: job.job_type.clone(),
        sunday_date: job.sunday_date.clone(),
    } );

    // Start heartbeat task
    let heartbeat_token = CancellationToken::new();
    let heartbeat_handle = tokio::spawn( heartbeat::run(
        Arc::clone( api ),
        job.id,
        config.worker.id.clone(),
        config.timing.heartbeat_interval_secs,
        heartbeat_token.clone(),
    ) );

    // Execute the job
    let result = job_runner::execute(
        job,
        config,
        Arc::clone( api ),
        shutdown.clone(),
    ).await;

    // Stop heartbeat
    heartbeat_token.cancel();
    let _ = heartbeat_handle.await;

    // Report result to server
    match result {
        Ok( output ) => {
            if let Err( e ) = api.complete_job(
                job.id,
                &config.worker.id,
                output.output_path,
                output.metadata,
            ).await {
                tracing::error!( "Failed to report job completion: {}", e );
            } else {
                tracing::info!( "Job #{} completed successfully", job.id );
                *state.jobs_completed.write().await += 1;
                state.emit( WorkerEvent::JobCompleted { job_id: job.id } );
                emit_stats( state ).await;
            }
        }
        Err( e ) => {
            let error_msg = format!( "{:#}", e );
            if let Err( e2 ) = api.fail_job(
                job.id,
                &config.worker.id,
                &error_msg,
            ).await {
                tracing::error!( "Failed to report job failure: {}", e2 );
            } else {
                tracing::warn!( "Job #{} failed: {}", job.id, error_msg );
                *state.jobs_failed.write().await += 1;
                state.emit( WorkerEvent::JobFailed {
                    job_id: job.id,
                    error: error_msg,
                } );
                emit_stats( state ).await;
            }
        }
    }

    // Clear current job
    *state.current_job_id.write().await = None;
    *state.current_job_type.write().await = None;
    state.emit( WorkerEvent::JobCleared );
}


/// Update the scheduler phase in shared state and emit a broadcast event.
async fn set_phase( state: &AppState, phase: SchedulerPhase ) {
    let current = *state.phase.read().await;
    if current != phase {
        *state.phase.write().await = phase;
        tracing::debug!( "Phase: {} → {}", current, phase );
        state.emit( WorkerEvent::PhaseChange { phase: phase.to_string() } );
    }
}

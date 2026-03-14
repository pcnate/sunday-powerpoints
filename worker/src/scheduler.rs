use chrono::{ Local, NaiveTime };
use std::sync::Arc;
use tokio::sync::mpsc;
use tokio_util::sync::CancellationToken;

use crate::api_client::ApiClient;
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

    loop {
        // Check for shutdown
        if shutdown.is_cancelled() {
            tracing::info!( "Scheduler shutting down" );
            break;
        }

        // Drain tray commands (non-blocking)
        while let Ok( cmd ) = tray_rx.try_recv() {
            if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await {
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
                    if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await { break; }
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
                    if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await { break; }
                }
            }
            continue;
        }

        // We're in shift — set to idle, then poll
        set_phase( &state, SchedulerPhase::Idle ).await;

        // Wait poll interval before checking for jobs
        tokio::select! {
            _ = tokio::time::sleep( std::time::Duration::from_secs( config.timing.poll_interval_secs ) ) => {}
            _ = shutdown.cancelled() => { break; }
            Some( cmd ) = tray_rx.recv() => {
                if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await { break; }
                continue;
            }
        }

        if shutdown.is_cancelled() {
            break;
        }

        // Poll for a job
        set_phase( &state, SchedulerPhase::Polling ).await;
        *state.connected.write().await = true;

        let claim_result = api.claim_job( &config.worker.id, &config.worker.types ).await;

        match claim_result {
            Ok( None ) => {
                // No jobs available
                tracing::debug!( "No pending jobs" );
                *state.connected.write().await = true;
            }
            Ok( Some( job ) ) => {
                tracing::info!(
                    "Claimed job #{} ({}) for {}",
                    job.id, job.job_type, job.sunday_date
                );

                *state.current_job_id.write().await = Some( job.id );
                *state.current_job_type.write().await = Some( job.job_type.clone() );
                set_phase( &state, SchedulerPhase::Executing ).await;

                // Start heartbeat task
                let heartbeat_token = CancellationToken::new();
                let heartbeat_handle = tokio::spawn( heartbeat::run(
                    Arc::clone( &api ),
                    job.id,
                    config.worker.id.clone(),
                    config.timing.heartbeat_interval_secs,
                    heartbeat_token.clone(),
                ) );

                // Execute the job
                let result = job_runner::execute(
                    &job,
                    &config,
                    Arc::clone( &api ),
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
                        }
                    }
                }

                // Clear current job
                *state.current_job_id.write().await = None;
                *state.current_job_type.write().await = None;

                // Cooldown before next poll
                set_phase( &state, SchedulerPhase::Cooldown ).await;
                tokio::select! {
                    _ = tokio::time::sleep( std::time::Duration::from_secs( config.timing.job_delay_secs ) ) => {}
                    _ = shutdown.cancelled() => { break; }
                    Some( cmd ) = tray_rx.recv() => {
                        if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await { break; }
                    }
                }
            }
            Err( e ) => {
                tracing::warn!( "Failed to poll for jobs: {}", e );
                *state.connected.write().await = false;
                // Back off a bit on network errors
                tokio::select! {
                    _ = tokio::time::sleep( std::time::Duration::from_secs( 10 ) ) => {}
                    _ = shutdown.cancelled() => { break; }
                    Some( cmd ) = tray_rx.recv() => {
                        if handle_command( cmd, &mut paused, &mut run_once, &state, &shutdown ).await { break; }
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
        }
    }
    false
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


/// Update the scheduler phase in shared state.
async fn set_phase( state: &AppState, phase: SchedulerPhase ) {
    let current = *state.phase.read().await;
    if current != phase {
        *state.phase.write().await = phase;
        tracing::debug!( "Phase: {} → {}", current, phase );
    }
}

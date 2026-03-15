use chrono::{ DateTime, Utc };
use serde::{ Deserialize, Serialize };
use tokio::sync::{ Notify, RwLock };

use crate::broadcast::{ EventSender, StatusSnapshot, WorkerEvent };
use crate::config::AppConfig;


/// Current scheduler phase.
#[derive( Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize )]
pub enum SchedulerPhase {
    /// Outside the configured shift window.
    OffShift,
    /// Within shift, waiting to poll.
    Idle,
    /// Actively polling the server for a job.
    Polling,
    /// Executing a claimed job.
    Executing,
    /// Brief cooldown between job completions.
    Cooldown,
    /// Manually paused by user.
    Paused,
}


impl std::fmt::Display for SchedulerPhase {
    fn fmt( &self, f: &mut std::fmt::Formatter<'_> ) -> std::fmt::Result {
        match self {
            SchedulerPhase::OffShift => write!( f, "Off Shift" ),
            SchedulerPhase::Idle => write!( f, "Idle" ),
            SchedulerPhase::Polling => write!( f, "Polling" ),
            SchedulerPhase::Executing => write!( f, "Executing" ),
            SchedulerPhase::Cooldown => write!( f, "Cooldown" ),
            SchedulerPhase::Paused => write!( f, "Paused" ),
        }
    }
}


/// Commands sent from the tray UI to the async scheduler.
#[derive( Debug, Clone )]
pub enum TrayCommand {
    /// Toggle between paused and active.
    TogglePause,
    /// Gracefully shut down the worker.
    Quit,
    /// Open the config web UI in the default browser.
    OpenConfig,
    /// Reload configuration from disk.
    ReloadConfig,
    /// Run exactly one job cycle then stop.
    RunOnce,
    /// Toggle force on-shift (bypass shift window).
    ToggleForceOnShift,
    /// Toggle pause for a specific job type (runtime only, not persisted).
    TogglePauseJobType( String ),
    /// Claim and execute a specific job by ID.
    RunJob( u32 ),
}


/// Snapshot of the tray icon state for the UI thread.
#[derive( Debug, Clone )]
pub struct TrayState {
    pub phase: SchedulerPhase,
    pub connected: bool,
    pub current_job_id: Option<u32>,
    pub current_job_type: Option<String>,
    pub jobs_completed: u32,
    pub jobs_failed: u32,
    pub uptime_secs: u64,
    pub force_on_shift: bool,
    pub active_types: Vec<String>,
    pub paused_types: Vec<String>,
}


impl Default for TrayState {
    fn default() -> Self {
        Self {
            phase: SchedulerPhase::OffShift,
            connected: false,
            current_job_id: None,
            current_job_type: None,
            jobs_completed: 0,
            jobs_failed: 0,
            uptime_secs: 0,
            force_on_shift: false,
            active_types: Vec::new(),
            paused_types: Vec::new(),
        }
    }
}


/// Shared application state accessible by all async tasks.
pub struct AppState {
    pub config: RwLock<AppConfig>,
    pub phase: RwLock<SchedulerPhase>,
    pub connected: RwLock<bool>,
    pub current_job_id: RwLock<Option<u32>>,
    pub current_job_type: RwLock<Option<String>>,
    pub jobs_completed: RwLock<u32>,
    pub jobs_failed: RwLock<u32>,
    pub started_at: DateTime<Utc>,
    pub force_on_shift: RwLock<bool>,
    pub paused_types: RwLock<Vec<String>>,
    pub event_tx: EventSender,
    /// Notified by the SSE client when a new job is created on the server.
    pub poll_notify: Notify,
}


impl AppState {
    /// Create a new AppState with the given configuration and event sender.
    pub fn new( config: AppConfig, event_tx: EventSender ) -> Self {
        Self {
            config: RwLock::new( config ),
            phase: RwLock::new( SchedulerPhase::OffShift ),
            connected: RwLock::new( false ),
            current_job_id: RwLock::new( None ),
            current_job_type: RwLock::new( None ),
            jobs_completed: RwLock::new( 0 ),
            jobs_failed: RwLock::new( 0 ),
            started_at: Utc::now(),
            force_on_shift: RwLock::new( false ),
            paused_types: RwLock::new( Vec::new() ),
            event_tx,
            poll_notify: Notify::new(),
        }
    }


    /// Broadcast an event to all WebSocket clients.
    ///
    /// Silently drops if there are no active receivers.
    pub fn emit( &self, event: WorkerEvent ) {
        let _ = self.event_tx.send( event );
    }


    /// Take a snapshot of the current state for the tray UI.
    pub async fn snapshot( &self ) -> TrayState {
        let config = self.config.read().await;

        TrayState {
            phase: *self.phase.read().await,
            connected: *self.connected.read().await,
            current_job_id: *self.current_job_id.read().await,
            current_job_type: self.current_job_type.read().await.clone(),
            jobs_completed: *self.jobs_completed.read().await,
            jobs_failed: *self.jobs_failed.read().await,
            uptime_secs: ( Utc::now() - self.started_at ).num_seconds() as u64,
            force_on_shift: *self.force_on_shift.read().await,
            active_types: config.worker.types.clone(),
            paused_types: self.paused_types.read().await.clone(),
        }
    }


    /// Build a full status snapshot for WebSocket clients.
    pub async fn status_snapshot( &self ) -> StatusSnapshot {
        let config = self.config.read().await;

        StatusSnapshot {
            phase: self.phase.read().await.to_string(),
            connected: *self.connected.read().await,
            current_job_id: *self.current_job_id.read().await,
            current_job_type: self.current_job_type.read().await.clone(),
            jobs_completed: *self.jobs_completed.read().await,
            jobs_failed: *self.jobs_failed.read().await,
            uptime_secs: ( Utc::now() - self.started_at ).num_seconds() as u64,
            worker_id: config.worker.id.clone(),
            server_url: config.server.url.clone(),
            schedule_enabled: config.schedule.enabled,
            shift_start: config.schedule.shift_start.clone(),
            shift_end: config.schedule.shift_end.clone(),
            force_on_shift: *self.force_on_shift.read().await,
            paused_types: self.paused_types.read().await.clone(),
            active_types: config.worker.types.clone(),
        }
    }
}

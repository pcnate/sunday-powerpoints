use serde::Serialize;
use tokio::sync::broadcast;


/// Events broadcast to WebSocket clients.
///
/// Serialized as `{"type": "event_name", "data": {...}}` using serde's
/// adjacently tagged representation. Unit variants omit the `data` field.
#[derive( Debug, Clone, Serialize )]
#[serde( tag = "type", content = "data", rename_all = "snake_case" )]
pub enum WorkerEvent {
    /// Full status snapshot (sent on connect and lag recovery).
    StatusSnapshot( StatusSnapshot ),
    /// Scheduler phase changed.
    PhaseChange { phase: String },
    /// Connection to Express server changed.
    ConnectionChange { connected: bool },
    /// A job was claimed by this worker.
    JobClaimed { job_id: u32, job_type: String, sunday_date: String },
    /// Current job completed successfully.
    JobCompleted { job_id: u32 },
    /// Current job failed.
    JobFailed { job_id: u32, error: String },
    /// Current job cleared (no longer executing).
    JobCleared,
    /// Stats updated.
    StatsUpdate { jobs_completed: u32, jobs_failed: u32 },
    /// Force-on-shift toggled.
    ForceOnShift { enabled: bool },
    /// Job type pause state changed.
    TypePauseChange { job_type: String, paused: bool },
    /// Config was reloaded — UI should re-fetch GET /api/config.
    ConfigReloaded,
}


/// Full status snapshot matching the GET /api/status response shape.
#[derive( Debug, Clone, Serialize )]
pub struct StatusSnapshot {
    pub phase: String,
    pub connected: bool,
    pub current_job_id: Option<u32>,
    pub current_job_type: Option<String>,
    pub jobs_completed: u32,
    pub jobs_failed: u32,
    pub uptime_secs: u64,
    pub worker_id: String,
    pub server_url: String,
    pub schedule_enabled: bool,
    pub shift_start: String,
    pub shift_end: String,
    pub force_on_shift: bool,
    pub paused_types: Vec<String>,
    pub active_types: Vec<String>,
}


pub type EventSender = broadcast::Sender<WorkerEvent>;

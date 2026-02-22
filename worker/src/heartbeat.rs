use std::sync::Arc;
use tokio_util::sync::CancellationToken;

use crate::api_client::ApiClient;


/// Run a periodic heartbeat for a job until cancelled.
///
/// Sends heartbeats at the configured interval. If the server
/// responds with `continue: false`, the heartbeat task logs a
/// warning (the scheduler will handle cancellation via its own
/// shutdown token).
///
/// @param api - HTTP client for the Express server
/// @param job_id - the job being processed
/// @param worker_id - this worker's identifier
/// @param interval_secs - seconds between heartbeats
/// @param cancel - token to stop the heartbeat loop
pub async fn run(
    api: Arc<ApiClient>,
    job_id: u32,
    worker_id: String,
    interval_secs: u64,
    cancel: CancellationToken,
) {
    let interval = std::time::Duration::from_secs( interval_secs );

    loop {
        tokio::select! {
            _ = tokio::time::sleep( interval ) => {}
            _ = cancel.cancelled() => {
                tracing::debug!( "Heartbeat for job #{} stopped", job_id );
                return;
            }
        }

        match api.heartbeat( job_id, &worker_id, None ).await {
            Ok( true ) => {
                tracing::debug!( "Heartbeat OK for job #{}", job_id );
            }
            Ok( false ) => {
                tracing::warn!( "Server says stop for job #{} (job may be cancelled)", job_id );
                // The scheduler will handle cancellation; we just log here
            }
            Err( e ) => {
                tracing::warn!( "Heartbeat failed for job #{}: {} (will retry)", job_id, e );
                // Don't stop — network errors are transient.
                // The server's stale recovery handles prolonged gaps.
            }
        }
    }
}

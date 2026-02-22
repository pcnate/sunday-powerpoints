use anyhow::Result;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// STUB: Video alignment runner.
///
/// In production, this will run FFmpeg scene detection on the main
/// camera MP4 and return scene change timestamps. The Express server
/// will then generate a Kdenlive project with those markers.
///
/// @param job - the video-alignment job
/// @param config - worker configuration (ffmpeg_path, etc.)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    tracing::info!(
        "STUB: Would run FFmpeg scene detection on {}",
        job.input_path
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "STUB: Running ffmpeg scene detection (tool: {})",
                config.tools.ffmpeg_path
            ),
        },
    ] ).await;

    // Simulate work
    tokio::select! {
        _ = tokio::time::sleep( std::time::Duration::from_secs( 5 ) ) => {}
        _ = shutdown.cancelled() => {
            anyhow::bail!( "Job cancelled during execution" );
        }
    }

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: "STUB: Scene detection complete (fake data)".to_string(),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: None,
        metadata: Some( serde_json::json!({
            "scene_changes": [ 0.0, 15.2, 42.7, 78.3, 120.5 ],
            "stub": true
        }) ),
    } )
}

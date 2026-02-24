use anyhow::Result;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// STUB: Kdenlive transcode runner.
///
/// In production, this will spawn `melt <input.kdenlive> -consumer
/// avformat:<output.mp4>` to render the edited Kdenlive project into
/// a production MP4.  Returns the production MP4 output path on completion.
///
/// @param job - the transcode job
/// @param config - worker configuration (melt_path, etc.)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    let output_path = job.output_path.as_deref().unwrap_or( "unknown" );

    tracing::info!(
        "STUB: Would run melt transcode on {} → {}",
        job.input_path,
        output_path
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "STUB: Running {} {} -consumer avformat:{}",
                config.tools.melt_path,
                job.input_path,
                output_path
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
            message: format!( "STUB: Transcode complete → {}", output_path ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( output_path.to_string() ),
        metadata: Some( serde_json::json!({
            "melt_path": config.tools.melt_path,
            "stub": true
        }) ),
    } )
}

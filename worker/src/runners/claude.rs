use anyhow::Result;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// STUB: Claude Code processing runner.
///
/// In production, this will spawn `claude --print --dangerously-skip-permissions
/// "<prompt>"` in the Sunday folder directory. The prompt template is stored
/// on the server and passed via job metadata.
///
/// @param job - the claude-processing job
/// @param config - worker configuration (claude_path)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    tracing::info!(
        "STUB: Would run Claude Code on {} → {}",
        job.input_path,
        job.output_path.as_deref().unwrap_or( "unknown output" )
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "STUB: Running {} on {}",
                config.tools.claude_path,
                job.input_path
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

    let output_path = job.output_path.clone().unwrap_or_else( || {
        format!( "{}/{} Notes.txt", job.sunday_date, job.sunday_date )
    } );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!( "STUB: Claude processing complete → {}", output_path ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( output_path ),
        metadata: Some( serde_json::json!({
            "prompt_template": "sermon-notes",
            "stub": true
        }) ),
    } )
}

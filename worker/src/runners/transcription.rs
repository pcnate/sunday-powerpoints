use anyhow::Result;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// STUB: Whisper transcription runner.
///
/// In production, this will spawn `whisper <input.mp4> --model large-v3
/// --device cuda` and stream stdout/stderr to the server log endpoint.
/// Returns the VTT output path on completion.
///
/// @param job - the transcription job
/// @param config - worker configuration (whisper_path, whisper_model)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    tracing::info!(
        "STUB: Would run Whisper transcription on {}",
        job.input_path
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "STUB: Running {} --model {} on {}",
                config.tools.whisper_path,
                config.tools.whisper_model,
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

    // Generate the expected VTT output path
    let vtt_path = job.input_path.replace( ".mp4", ".vtt" );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!( "STUB: Transcription complete → {}", vtt_path ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( vtt_path.clone() ),
        metadata: Some( serde_json::json!({
            "vtt_path": vtt_path,
            "model": config.tools.whisper_model,
            "stub": true
        }) ),
    } )
}

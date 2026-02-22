use anyhow::Result;
use std::sync::Arc;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::runners;


/// Output produced by a successful job execution.
pub struct JobOutput {
    pub output_path: Option<String>,
    pub metadata: Option<serde_json::Value>,
}


/// Execute a job by dispatching to the appropriate runner.
///
/// @param job - the claimed job to execute
/// @param config - current worker configuration
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn execute(
    job: &Job,
    config: &AppConfig,
    api: Arc<ApiClient>,
    shutdown: CancellationToken,
) -> Result<JobOutput> {
    // Send a log entry indicating the job has started
    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "Worker '{}' starting {} job for {}",
                config.worker.id, job.job_type, job.sunday_date
            ),
        },
    ] ).await;

    let result = match job.job_type.as_str() {
        "video-alignment" => runners::alignment::run( job, config, &api, &shutdown ).await,
        "transcription" => runners::transcription::run( job, config, &api, &shutdown ).await,
        "claude-processing" => runners::claude::run( job, config, &api, &shutdown ).await,
        other => {
            anyhow::bail!( "Unknown job type: {}", other );
        }
    };

    // Log completion or failure
    match &result {
        Ok( _ ) => {
            let _ = api.send_logs( job.id, vec![
                LogEntry {
                    level: "info".to_string(),
                    message: "Job completed successfully".to_string(),
                },
            ] ).await;
        }
        Err( e ) => {
            let _ = api.send_logs( job.id, vec![
                LogEntry {
                    level: "error".to_string(),
                    message: format!( "Job failed: {:#}", e ),
                },
            ] ).await;
        }
    }

    result
}

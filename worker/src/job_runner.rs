use anyhow::Result;
use std::sync::Arc;
use std::time::Duration;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::runners;


/// Output produced by a successful job execution.
pub struct JobOutput {
    pub output_path: Option<String>,
    pub metadata: Option<serde_json::Value>,
}


/// Resolve a job's relative input_path to an absolute path.
///
/// Job paths are stored relative to `%OneDriveConsumer%` (the OneDrive root),
/// so any machine with OneDrive can resolve them without extra configuration.
/// Falls back to `output_directory` config if the env var is not set.
fn resolve_input_path( config: &AppConfig, relative_path: &str ) -> String {
    let normalized = relative_path.replace( '/', "\\" );

    // Already absolute (e.g. starts with drive letter or UNC)
    if normalized.len() >= 2 && normalized.as_bytes()[ 1 ] == b':'
        || normalized.starts_with( "\\\\" )
    {
        return normalized;
    }

    // Resolve OneDrive root: prefer env var, fall back to configured output_directory
    let onedrive_root = std::env::var( "OneDriveConsumer" )
        .unwrap_or_else( |_| {
            let configured = config.paths.output_directory.trim_end_matches( [ '/', '\\' ] );
            configured.to_string()
        });

    if onedrive_root.is_empty() {
        tracing::warn!( "OneDriveConsumer env var not set and no output_directory configured" );
        return normalized;
    }

    format!( "{}\\{}", onedrive_root.trim_end_matches( [ '/', '\\' ] ), normalized )
}


/// Resolve a job's relative output_path to an absolute path.
fn resolve_output_path( config: &AppConfig, relative_path: &str ) -> String {
    resolve_input_path( config, relative_path )
}


/// Check whether a file has Windows OneDrive placeholder attributes.
/// Returns true if the file is NOT ready (still a cloud placeholder).
#[cfg( target_os = "windows" )]
fn is_onedrive_placeholder( path: &str ) -> bool {
    use std::os::windows::fs::MetadataExt;

    const FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS: u32 = 0x00400000;
    const FILE_ATTRIBUTE_RECALL_ON_OPEN: u32 = 0x00040000;

    match std::fs::metadata( path ) {
        Ok( meta ) => {
            let attrs = meta.file_attributes();
            ( attrs & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS ) != 0
                || ( attrs & FILE_ATTRIBUTE_RECALL_ON_OPEN ) != 0
        }
        Err( _ ) => false,
    }
}


/// Non-Windows stub — files are always "ready".
#[cfg( not( target_os = "windows" ) )]
fn is_onedrive_placeholder( _path: &str ) -> bool {
    false
}


/// Trigger OneDrive to download a cloud-only placeholder by reading from it.
///
/// On Windows, files with FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS are recalled
/// (downloaded) when their data is accessed. Simply opening a handle is not
/// enough — we must perform an actual read to trigger hydration.
#[cfg( target_os = "windows" )]
fn trigger_onedrive_hydration( path: &str ) {
    use std::io::Read;
    match std::fs::File::open( path ) {
        Ok( mut f ) => {
            let mut buf = [ 0u8; 1 ];
            match f.read_exact( &mut buf ) {
                Ok( () ) => tracing::info!( "Triggered OneDrive hydration for {}", path ),
                Err( e ) => tracing::warn!( "Read failed (hydration may still start): {}", e ),
            }
        }
        Err( e ) => tracing::warn!( "Could not open file to trigger hydration: {}", e ),
    }
}


/// Non-Windows stub — no-op.
#[cfg( not( target_os = "windows" ) )]
fn trigger_onedrive_hydration( _path: &str ) {}


/// Wait until a file is fully synced from OneDrive.
///
/// If the file is an OneDrive placeholder (cloud-only), triggers hydration
/// by opening it for read, then polls until the download completes.
/// Checks for placeholder attributes and file size stability.
/// Polls every 30 seconds, sending heartbeats to keep the job alive.
/// Returns an error if the file doesn't exist.
async fn wait_for_file_ready(
    path: &str,
    api: &ApiClient,
    job_id: u32,
    worker_id: &str,
    shutdown: &CancellationToken,
) -> Result<()> {
    // Check file exists
    if !std::path::Path::new( path ).exists() {
        anyhow::bail!( "Input file does not exist: {}", path );
    }

    // If it's a placeholder, trigger hydration immediately
    if is_onedrive_placeholder( path ) {
        let _ = api.send_logs( job_id, vec![
            LogEntry {
                level: "info".to_string(),
                message: format!( "File is cloud-only, triggering OneDrive download: {}", path ),
            },
        ] ).await;
        trigger_onedrive_hydration( path );
    }

    let mut logged_waiting = false;

    loop {
        // Check OneDrive placeholder attributes
        if !is_onedrive_placeholder( path ) {
            // Check file size stability (read, wait 2s, read again)
            let size1 = std::fs::metadata( path )
                .map( |m| m.len() )
                .unwrap_or( 0 );

            tokio::select! {
                _ = tokio::time::sleep( Duration::from_secs( 2 ) ) => {}
                _ = shutdown.cancelled() => {
                    anyhow::bail!( "Shutdown requested while waiting for file sync" );
                }
            }

            let size2 = std::fs::metadata( path )
                .map( |m| m.len() )
                .unwrap_or( 0 );

            if size1 == size2 && size1 > 0 {
                // File is stable and ready
                if logged_waiting {
                    let _ = api.send_logs( job_id, vec![
                        LogEntry {
                            level: "info".to_string(),
                            message: "File is now synced and ready".to_string(),
                        },
                    ] ).await;
                }
                return Ok(());
            }
        }

        if !logged_waiting {
            let _ = api.send_logs( job_id, vec![
                LogEntry {
                    level: "info".to_string(),
                    message: format!( "Waiting for OneDrive sync: {}", path ),
                },
            ] ).await;
            logged_waiting = true;
        }

        // Poll every 30 seconds, send heartbeat to keep job alive
        tokio::select! {
            _ = tokio::time::sleep( Duration::from_secs( 30 ) ) => {}
            _ = shutdown.cancelled() => {
                anyhow::bail!( "Shutdown requested while waiting for file sync" );
            }
        }

        let _ = api.heartbeat( job_id, worker_id, None ).await;
    }
}


/// Execute a job by dispatching to the appropriate runner.
///
/// Resolves relative paths to absolute, waits for OneDrive sync,
/// then dispatches to the type-specific runner.
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
    // Resolve relative paths to absolute using worker's output_directory
    let mut resolved_job = job.clone();
    resolved_job.input_path = resolve_input_path( config, &job.input_path );
    if let Some( ref out ) = job.output_path {
        resolved_job.output_path = Some( resolve_output_path( config, out ) );
    }

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

    // Wait for input file to be synced from OneDrive
    wait_for_file_ready(
        &resolved_job.input_path,
        &api,
        job.id,
        &config.worker.id,
        &shutdown,
    ).await?;

    let result = match job.job_type.as_str() {
        "transcode" => runners::transcode::run( &resolved_job, config, &api, &shutdown ).await,
        "transcription" => runners::transcription::run( &resolved_job, config, &api, &shutdown ).await,
        "claude-processing" => runners::claude::run( &resolved_job, config, &api, &shutdown ).await,
        "ffprobe" => runners::ffprobe::run( &resolved_job, config, &api, &shutdown ).await,
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

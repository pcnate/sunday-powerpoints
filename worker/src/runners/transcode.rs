use anyhow::{ Context, Result };
use std::time::{ Duration, Instant };
use tokio::io::{ AsyncBufReadExt, AsyncReadExt, BufReader };
use tokio::process::Command;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// Run the melt transcode for a Kdenlive project.
///
/// Spawns `melt <input.kdenlive> -progress -consumer avformat:<output.mp4>`
/// with configured codec/bitrate settings.  Streams progress to the server
/// log endpoint, handles graceful cancellation, and verifies the output file.
///
/// @param job - the transcode job (input_path = .kdenlive, output_path = .mp4)
/// @param config - worker configuration (melt_path, bitrate settings)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    // --- Validate inputs ---
    let output_path = job.output_path.as_deref()
        .filter( |p| !p.is_empty() )
        .context( "Transcode job is missing output_path" )?;

    let input = std::path::Path::new( &job.input_path );
    if !input.exists() {
        anyhow::bail!( "Input file does not exist: {}", job.input_path );
    }

    let consumer_arg = format!( "avformat:{}", output_path );

    tracing::info!(
        "Starting melt transcode: {} → {}",
        job.input_path,
        output_path
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "& \"{}\" \"{}\" -progress -consumer \"{}\" vcodec=libx264 acodec=aac b={} ab={}",
                config.tools.melt_path,
                job.input_path,
                consumer_arg,
                config.tools.melt_video_bitrate,
                config.tools.melt_audio_bitrate,
            ),
        },
    ] ).await;

    // --- Spawn melt subprocess ---
    let mut child = Command::new( &config.tools.melt_path )
        .arg( &job.input_path )
        .arg( "-progress" )
        .arg( "-consumer" )
        .arg( &consumer_arg )
        .arg( "vcodec=libx264" )
        .arg( "acodec=aac" )
        .arg( format!( "b={}", config.tools.melt_video_bitrate ) )
        .arg( format!( "ab={}", config.tools.melt_audio_bitrate ) )
        .stdout( std::process::Stdio::piped() )
        .stderr( std::process::Stdio::piped() )
        .spawn()
        .context( "Failed to spawn melt process — is melt installed and in PATH?" )?;

    // --- Collect stderr in background ---
    let stderr = child.stderr.take().expect( "stderr was piped" );
    let stderr_handle = tokio::spawn( async move {
        let mut buf = String::new();
        BufReader::new( stderr ).read_to_string( &mut buf ).await.ok();
        buf
    } );

    // --- Stream stdout progress ---
    let stdout = child.stdout.take().expect( "stdout was piped" );
    let reader = BufReader::new( stdout );
    let mut lines = reader.lines();
    let mut last_log = Instant::now();
    let mut last_percentage: Option<f64> = None;

    loop {
        tokio::select! {
            line = lines.next_line() => {
                match line? {
                    Some( text ) => {
                        // Parse percentage from melt progress output
                        if let Some( pct ) = parse_percentage( &text ) {
                            last_percentage = Some( pct );
                        }

                        // Throttle log sends to every 5 seconds
                        if last_log.elapsed() > Duration::from_secs( 5 ) {
                            let msg = if let Some( pct ) = last_percentage {
                                format!( "Transcoding: {:.1}%", pct )
                            } else {
                                text.trim().to_string()
                            };

                            if !msg.is_empty() {
                                let _ = api.send_logs( job.id, vec![
                                    LogEntry {
                                        level: "info".to_string(),
                                        message: msg,
                                    },
                                ] ).await;
                            }

                            last_log = Instant::now();
                        }
                    }
                    None => break, // EOF — process exited
                }
            }
            _ = shutdown.cancelled() => {
                let _ = child.kill().await;
                anyhow::bail!( "Job cancelled — melt process killed" );
            }
        }
    }

    // --- Check exit code ---
    let status = child.wait().await
        .context( "Failed to wait for melt process" )?;

    if !status.success() {
        let stderr_output = stderr_handle.await.unwrap_or_default();
        anyhow::bail!(
            "Melt exited with code {}: {}",
            status.code().unwrap_or( -1 ),
            stderr_output.trim()
        );
    }

    // --- Verify output file exists ---
    if !std::path::Path::new( output_path ).exists() {
        anyhow::bail!( "Melt completed but output file not found: {}", output_path );
    }

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!( "Transcode complete → {}", output_path ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( output_path.to_string() ),
        metadata: Some( serde_json::json!({
            "melt_path": config.tools.melt_path,
            "video_bitrate": config.tools.melt_video_bitrate,
            "audio_bitrate": config.tools.melt_audio_bitrate,
        }) ),
    } )
}


/// Parse a percentage value from a melt progress line.
///
/// Melt outputs lines like `percentage: 42` (integer 0–100).
fn parse_percentage( line: &str ) -> Option<f64> {
    let line = line.trim();
    if let Some( rest ) = line.strip_prefix( "percentage:" ) {
        rest.trim().parse::<f64>().ok()
    } else {
        None
    }
}

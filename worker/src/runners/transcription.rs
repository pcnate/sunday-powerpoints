use anyhow::{ Context, Result };
use std::time::{ Duration, Instant };
use tokio::io::{ AsyncBufReadExt, BufReader };
use tokio::process::Command;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// Run Whisper transcription on a production MP4.
///
/// Spawns `whisper-ctranslate2 <input.mp4> --model <model> --device cuda
/// --output_format vtt --output_dir <dir> --language en`.
/// Streams stderr progress to the server log endpoint,
/// handles graceful cancellation, and returns the VTT output path.
///
/// @param job - the transcription job (input_path = production MP4)
/// @param config - worker configuration (whisper_path, whisper_model, tool_env)
/// @param api - API client for sending logs
/// @param shutdown - cancellation token for graceful shutdown
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    shutdown: &CancellationToken,
) -> Result<JobOutput> {
    // --- Validate inputs ---
    let input = std::path::Path::new( &job.input_path );
    if !input.exists() {
        anyhow::bail!( "Input file does not exist: {}", job.input_path );
    }

    let output_dir = input.parent()
        .context( "Could not determine output directory from input path" )?
        .to_string_lossy()
        .to_string();

    // whisper-ctranslate2 names its output <stem>.vtt
    let stem = input.file_stem()
        .context( "Could not determine file stem from input path" )?
        .to_string_lossy();
    let expected_vtt = format!( "{}\\{}.vtt", output_dir, stem );

    tracing::info!(
        "Starting whisper transcription: {} → {}",
        job.input_path,
        expected_vtt
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "& \"{}\" \"{}\" --model {} --device cuda --output_format vtt --output_dir \"{}\" --compute_type {} --language en",
                config.tools.whisper_path,
                job.input_path,
                config.tools.whisper_model,
                output_dir,
                config.tools.whisper_compute_type,
            ),
        },
    ] ).await;

    // --- Spawn whisper subprocess ---
    let mut cmd = Command::new( &config.tools.whisper_path );
    cmd.arg( &job.input_path )
        .arg( "--model" ).arg( &config.tools.whisper_model )
        .arg( "--device" ).arg( "cuda" )
        .arg( "--output_format" ).arg( "vtt" )
        .arg( "--output_dir" ).arg( &output_dir )
        .arg( "--compute_type" ).arg( &config.tools.whisper_compute_type )
        .arg( "--language" ).arg( "en" )
        .stdout( std::process::Stdio::null() )
        .stderr( std::process::Stdio::piped() );

    // Apply per-tool environment variables (e.g. PYTHONIOENCODING=utf-8)
    if let Some( env_vars ) = config.tools.tool_env.get( "whisper" ) {
        for ( key, value ) in env_vars {
            cmd.env( key, value );
        }
    }

    let mut child = cmd.spawn()
        .context( "Failed to spawn whisper process — is whisper-ctranslate2 installed?" )?;

    // --- Stream stderr progress ---
    let stderr = child.stderr.take().expect( "stderr was piped" );
    let reader = BufReader::new( stderr );
    let mut lines = reader.lines();
    let mut last_log = Instant::now();
    let mut last_progress = String::new();

    loop {
        tokio::select! {
            line = lines.next_line() => {
                match line? {
                    Some( text ) => {
                        let trimmed = text.trim().to_string();
                        if !trimmed.is_empty() {
                            last_progress = trimmed;
                        }

                        // Throttle log sends to every 10 seconds
                        if last_log.elapsed() > Duration::from_secs( 10 ) && !last_progress.is_empty() {
                            let _ = api.send_logs( job.id, vec![
                                LogEntry {
                                    level: "info".to_string(),
                                    message: truncate( &last_progress, 200 ),
                                },
                            ] ).await;
                            last_log = Instant::now();
                        }
                    }
                    None => break, // EOF — process exited
                }
            }
            _ = shutdown.cancelled() => {
                let _ = child.kill().await;
                anyhow::bail!( "Job cancelled — whisper process killed" );
            }
        }
    }

    // --- Check exit code ---
    let status = child.wait().await
        .context( "Failed to wait for whisper process" )?;

    let vtt_exists = std::path::Path::new( &expected_vtt ).exists();

    if !status.success() && !vtt_exists {
        anyhow::bail!(
            "Whisper exited with code {}: {}",
            status.code().unwrap_or( -1 ),
            truncate( last_progress.trim(), 500 )
        );
    }

    if !vtt_exists {
        anyhow::bail!( "Whisper completed but VTT not found: {}", expected_vtt );
    }

    if !status.success() {
        let _ = api.send_logs( job.id, vec![
            LogEntry {
                level: "warn".to_string(),
                message: format!(
                    "Whisper exited with code {} but VTT was created — treating as success",
                    status.code().unwrap_or( -1 )
                ),
            },
        ] ).await;
    }

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!( "Transcription complete → {}", expected_vtt ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( expected_vtt.clone() ),
        metadata: Some( serde_json::json!({
            "vtt_path": expected_vtt,
            "model": config.tools.whisper_model,
        }) ),
    } )
}


/// Truncate a string to at most `max` characters.
fn truncate( s: &str, max: usize ) -> String {
    if s.len() <= max {
        s.to_string()
    } else {
        format!( "{}…", &s[ ..max ] )
    }
}

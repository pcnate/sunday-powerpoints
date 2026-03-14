use anyhow::{ Context, Result };
use std::time::{ Duration, Instant };
use tokio::io::{ AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader };
use tokio::process::Command;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// Default prompt used when the job metadata has no prompt field.
const DEFAULT_PROMPT: &str = "Read the VTT file and provide:

1. Sermon title
2. YouTube timestamp bookmarks for: welcome, songs (with hymn number), \
memory verse (with reference), tithes & offering, scripture reading, \
opening prayer, sermon start, and closing song
3. List all Bible verses referenced during the sermon using \
{book} {chapter} v{verse} format
4. Suggested YouTube tags for the sermon";


/// Run Claude Code processing on a VTT transcription file.
///
/// Reads the VTT file content, combines it with the prompt template
/// (from job metadata or default), pipes the full prompt to
/// `claude --print --dangerously-skip-permissions`, captures stdout,
/// and writes the result to `Summary.txt`.
///
/// @param job - the claude-processing job (input_path = VTT file)
/// @param config - worker configuration (claude_path, tool_env)
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
        anyhow::bail!( "Input VTT file does not exist: {}", job.input_path );
    }

    let output_dir = input.parent()
        .context( "Could not determine output directory from input path" )?;

    // Determine output path: Summary.txt in the sunday folder
    let output_path = job.output_path.clone().unwrap_or_else( || {
        output_dir.join( "Summary.txt" )
            .to_string_lossy()
            .to_string()
    } );

    // --- Read VTT content ---
    let vtt_content = tokio::fs::read_to_string( &job.input_path ).await
        .context( "Failed to read VTT file" )?;

    // --- Build prompt ---
    let instruction = job.metadata.as_ref()
        .and_then( |m| m.get( "prompt" ) )
        .and_then( |v| v.as_str() )
        .unwrap_or( DEFAULT_PROMPT );

    let prompt = format!(
        "{}\n\n<vtt>\n{}\n</vtt>\n\nWrite your response as plain text, not markdown.",
        instruction,
        vtt_content,
    );

    tracing::info!(
        "Starting Claude processing: {} → {}",
        job.input_path,
        output_path,
    );

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "& \"{}\" --print --dangerously-skip-permissions (VTT: {} bytes, prompt piped via stdin)",
                config.tools.claude_path,
                vtt_content.len(),
            ),
        },
    ] ).await;

    // --- Spawn claude subprocess ---
    let mut cmd = Command::new( &config.tools.claude_path );
    cmd.arg( "--print" )
        .arg( "--dangerously-skip-permissions" )
        .stdin( std::process::Stdio::piped() )
        .stdout( std::process::Stdio::piped() )
        .stderr( std::process::Stdio::piped() )
        .current_dir( output_dir );

    // Apply per-tool environment variables
    if let Some( env_vars ) = config.tools.tool_env.get( "claude" ) {
        for ( key, value ) in env_vars {
            cmd.env( key, value );
        }
    }

    let mut child = cmd.spawn()
        .context( "Failed to spawn claude process — is claude CLI installed?" )?;

    // --- Write prompt to stdin, then close ---
    if let Some( mut stdin ) = child.stdin.take() {
        stdin.write_all( prompt.as_bytes() ).await
            .context( "Failed to write prompt to claude stdin" )?;
        drop( stdin );
    }

    // --- Collect stdout in background task ---
    let stdout = child.stdout.take().expect( "stdout was piped" );
    let stdout_handle = tokio::spawn( async move {
        let mut buf = Vec::new();
        let mut reader = BufReader::new( stdout );
        reader.read_to_end( &mut buf ).await?;
        Ok::<Vec<u8>, std::io::Error>( buf )
    } );

    // --- Stream stderr for progress ---
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
                anyhow::bail!( "Job cancelled — claude process killed" );
            }
        }
    }

    // --- Check exit code ---
    let status = child.wait().await
        .context( "Failed to wait for claude process" )?;

    if !status.success() {
        anyhow::bail!(
            "Claude exited with code {}: {}",
            status.code().unwrap_or( -1 ),
            truncate( &last_progress, 500 )
        );
    }

    // --- Capture stdout and write output file ---
    let stdout_bytes = stdout_handle.await
        .context( "stdout collection task panicked" )?
        .context( "Failed to read claude stdout" )?;

    let response = String::from_utf8_lossy( &stdout_bytes );
    let trimmed_response = response.trim();

    if trimmed_response.is_empty() {
        anyhow::bail!( "Claude returned empty output" );
    }

    tokio::fs::write( &output_path, trimmed_response ).await
        .context( format!( "Failed to write output to {}", output_path ) )?;

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "Claude processing complete → {} ({} bytes)",
                output_path,
                trimmed_response.len()
            ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: Some( output_path ),
        metadata: Some( serde_json::json!({
            "prompt_template": "summary",
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

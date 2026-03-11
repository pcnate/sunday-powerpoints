use anyhow::{ Context, Result };
use tokio::process::Command;
use tokio_util::sync::CancellationToken;

use crate::api_client::{ ApiClient, Job, LogEntry };
use crate::config::AppConfig;
use crate::job_runner::JobOutput;


/// Run ffprobe on a video file and return metadata.
///
/// Executes ffprobe to extract duration, codecs, resolution, framerate,
/// audio stream count, bitrate, and file size. Also reads the file's
/// filesystem creation date for later use in purge decisions.
///
/// @param job - the ffprobe job (input_path = video file)
/// @param config - worker configuration (ffprobe_path)
/// @param api - API client for sending logs
/// @param _shutdown - cancellation token (ffprobe is fast, but respected)
pub async fn run(
    job: &Job,
    config: &AppConfig,
    api: &ApiClient,
    _shutdown: &CancellationToken,
) -> Result<JobOutput> {
    let input = std::path::Path::new( &job.input_path );
    if !input.exists() {
        anyhow::bail!( "Input file does not exist: {}", job.input_path );
    }

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "& \"{}\" -v error -show_entries format=duration,size,bit_rate:stream=codec_name,codec_type,width,height,r_frame_rate,channels -of json \"{}\"",
                config.tools.ffprobe_path,
                job.input_path,
            ),
        },
    ] ).await;

    // Run ffprobe
    let output = Command::new( &config.tools.ffprobe_path )
        .arg( "-v" ).arg( "error" )
        .arg( "-show_entries" )
        .arg( "format=duration,size,bit_rate:stream=codec_name,codec_type,width,height,r_frame_rate,channels" )
        .arg( "-of" ).arg( "json" )
        .arg( &job.input_path )
        .output()
        .await
        .context( "Failed to spawn ffprobe process — is ffprobe installed and in PATH?" )?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy( &output.stderr );
        anyhow::bail!( "ffprobe exited with code {}: {}", output.status.code().unwrap_or( -1 ), stderr.trim() );
    }

    let stdout = String::from_utf8_lossy( &output.stdout );
    let probe: serde_json::Value = serde_json::from_str( &stdout )
        .context( "Failed to parse ffprobe JSON output" )?;

    // Extract format info
    let format = &probe[ "format" ];
    let duration_secs: f64 = format[ "duration" ]
        .as_str()
        .and_then( |s| s.parse().ok() )
        .unwrap_or( 0.0 );

    let file_size: u64 = format[ "size" ]
        .as_str()
        .and_then( |s| s.parse().ok() )
        .unwrap_or( 0 );

    let bitrate: u64 = format[ "bit_rate" ]
        .as_str()
        .and_then( |s| s.parse().ok() )
        .unwrap_or( 0 );

    // Convert duration to timecode
    let duration_timecode = secs_to_timecode( duration_secs );

    // Extract stream info
    let streams = probe[ "streams" ].as_array();
    let mut video_codec = String::new();
    let mut audio_codec = String::new();
    let mut width: u32 = 0;
    let mut height: u32 = 0;
    let mut framerate: f64 = 0.0;
    let mut audio_streams: u32 = 0;

    if let Some( streams ) = streams {
        for stream in streams {
            let codec_type = stream[ "codec_type" ].as_str().unwrap_or( "" );
            match codec_type {
                "video" => {
                    if video_codec.is_empty() {
                        video_codec = stream[ "codec_name" ].as_str().unwrap_or( "" ).to_string();
                        width = stream[ "width" ].as_u64().unwrap_or( 0 ) as u32;
                        height = stream[ "height" ].as_u64().unwrap_or( 0 ) as u32;
                        framerate = parse_framerate( stream[ "r_frame_rate" ].as_str().unwrap_or( "0/1" ) );
                    }
                }
                "audio" => {
                    if audio_codec.is_empty() {
                        audio_codec = stream[ "codec_name" ].as_str().unwrap_or( "" ).to_string();
                    }
                    audio_streams += 1;
                }
                _ => {}
            }
        }
    }

    // Read file creation time
    let created_at = std::fs::metadata( &job.input_path )
        .ok()
        .and_then( |m| m.created().ok() )
        .map( |t| {
            let dt: chrono::DateTime<chrono::Utc> = t.into();
            dt.to_rfc3339()
        } )
        .unwrap_or_default();

    let metadata = serde_json::json!({
        "duration_secs": duration_secs,
        "duration_timecode": duration_timecode,
        "framerate": framerate,
        "width": width,
        "height": height,
        "video_codec": video_codec,
        "audio_codec": audio_codec,
        "audio_streams": audio_streams,
        "bitrate": bitrate,
        "file_size": file_size,
        "created_at": created_at,
    });

    let _ = api.send_logs( job.id, vec![
        LogEntry {
            level: "info".to_string(),
            message: format!(
                "ffprobe complete: {}x{} {:.2}fps {} duration={}",
                width, height, framerate, video_codec, duration_timecode
            ),
        },
    ] ).await;

    Ok( JobOutput {
        output_path: None,
        metadata: Some( metadata ),
    } )
}


/// Convert seconds to HH:MM:SS.mmm timecode format.
fn secs_to_timecode( secs: f64 ) -> String {
    let total_ms = ( secs * 1000.0 ) as u64;
    let hours = total_ms / 3_600_000;
    let minutes = ( total_ms % 3_600_000 ) / 60_000;
    let seconds = ( total_ms % 60_000 ) / 1_000;
    let millis = total_ms % 1_000;
    format!( "{:02}:{:02}:{:02}.{:03}", hours, minutes, seconds, millis )
}


/// Parse a framerate fraction string like "30000/1001" into a float.
fn parse_framerate( s: &str ) -> f64 {
    if let Some( ( num, den ) ) = s.split_once( '/' ) {
        let n: f64 = num.trim().parse().unwrap_or( 0.0 );
        let d: f64 = den.trim().parse().unwrap_or( 1.0 );
        if d > 0.0 { n / d } else { 0.0 }
    } else {
        s.parse().unwrap_or( 0.0 )
    }
}

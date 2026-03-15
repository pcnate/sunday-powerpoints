use anyhow::{ Context, Result };
use serde::{ Deserialize, Serialize };
use std::collections::HashMap;
use std::path::PathBuf;


/// Server connection configuration.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct ServerConfig {
    pub url: String,
}


/// Worker identity and capability configuration.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct WorkerConfig {
    pub id: String,
    pub types: Vec<String>,
}


/// Shift schedule configuration.
///
/// If `shift_start` > `shift_end`, the shift crosses midnight
/// (e.g., 22:00–06:00). When `enabled` is false, the worker
/// stays off-shift and only processes jobs via manual "run once".
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct ScheduleConfig {
    #[serde( default = "default_true" )]
    pub enabled: bool,
    pub shift_start: String,
    pub shift_end: String,
}


/// Default value for boolean fields that should be true.
fn default_true() -> bool {
    true
}


/// Polling and heartbeat timing configuration.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct TimingConfig {
    pub poll_interval_secs: u64,
    pub heartbeat_interval_secs: u64,
    pub job_delay_secs: u64,
}


/// File path configuration for resolving job paths.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct PathsConfig {
    #[serde( default = "default_output_directory" )]
    pub output_directory: String,
}


/// Default output directory — resolves %OneDriveConsumer% if available.
fn default_output_directory() -> String {
    std::env::var( "OneDriveConsumer" ).unwrap_or_default()
}


impl Default for PathsConfig {
    fn default() -> Self {
        Self {
            output_directory: default_output_directory(),
        }
    }
}


/// External tool paths and settings.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct ToolsConfig {
    pub whisper_path: String,
    pub whisper_model: String,
    pub claude_path: String,
    pub ffmpeg_path: String,
    #[serde( default = "default_ffprobe_path" )]
    pub ffprobe_path: String,
    #[serde( default = "default_melt_path" )]
    pub melt_path: String,
    #[serde( default = "default_video_bitrate" )]
    pub melt_video_bitrate: String,
    #[serde( default = "default_audio_bitrate" )]
    pub melt_audio_bitrate: String,
    #[serde( default = "default_compute_type" )]
    pub whisper_compute_type: String,
    /// Per-tool environment variables. Outer key is tool name
    /// (whisper, claude, ffprobe, ffmpeg, melt), inner map is env var key-value pairs.
    #[serde( default )]
    pub tool_env: HashMap<String, HashMap<String, String>>,
}


/// Default ffprobe binary path.
fn default_ffprobe_path() -> String {
    "ffprobe".to_string()
}


/// Default melt binary path.
fn default_melt_path() -> String {
    "melt".to_string()
}


/// Default video bitrate for melt transcode.
fn default_video_bitrate() -> String {
    "5000k".to_string()
}


/// Default audio bitrate for melt transcode.
fn default_audio_bitrate() -> String {
    "192k".to_string()
}


/// Default compute type for whisper-ctranslate2.
fn default_compute_type() -> String {
    "float16".to_string()
}


/// Embedded web UI configuration.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct WebUiConfig {
    pub port: u16,
}


/// Top-level application configuration loaded from TOML.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct AppConfig {
    pub server: ServerConfig,
    pub worker: WorkerConfig,
    pub schedule: ScheduleConfig,
    pub timing: TimingConfig,
    pub tools: ToolsConfig,
    #[serde( default )]
    pub paths: PathsConfig,
    pub web_ui: WebUiConfig,
}


impl Default for AppConfig {
    /// Returns a sensible default configuration.
    fn default() -> Self {
        Self {
            server: ServerConfig {
                url: "http://localhost:8080".to_string(),
            },
            worker: WorkerConfig {
                id: default_worker_id(),
                types: vec![
                    "transcription".to_string(),
                    "claude-processing".to_string(),
                ],
            },
            schedule: ScheduleConfig {
                enabled: true,
                shift_start: "02:00".to_string(),
                shift_end: "06:00".to_string(),
            },
            timing: TimingConfig {
                poll_interval_secs: 30,
                heartbeat_interval_secs: 30,
                job_delay_secs: 5,
            },
            tools: ToolsConfig {
                whisper_path: "whisper".to_string(),
                whisper_model: "large-v3".to_string(),
                claude_path: "claude".to_string(),
                ffmpeg_path: "ffmpeg".to_string(),
                ffprobe_path: "ffprobe".to_string(),
                melt_path: "melt".to_string(),
                melt_video_bitrate: "5000k".to_string(),
                melt_audio_bitrate: "192k".to_string(),
                whisper_compute_type: "float16".to_string(),
                tool_env: HashMap::new(),
            },
            paths: PathsConfig::default(),
            web_ui: WebUiConfig {
                port: 9090,
            },
        }
    }
}


/// Get the machine hostname for use as the default worker ID.
fn default_worker_id() -> String {
    std::env::var( "COMPUTERNAME" )
        .or_else( |_| std::env::var( "HOSTNAME" ) )
        .unwrap_or_else( |_| "sunday-worker".to_string() )
        .to_lowercase()
}


/// Get the configuration directory path.
///
/// Returns `%APPDATA%\sunday-worker` on Windows, or
/// `~/.config/sunday-worker` on other platforms.
pub fn config_dir() -> Result<PathBuf> {
    let base = if cfg!( target_os = "windows" ) {
        std::env::var( "APPDATA" )
            .map( PathBuf::from )
            .context( "APPDATA environment variable not set" )?
    } else {
        dirs_fallback()
            .context( "Could not determine config directory" )?
    };

    Ok( base.join( "sunday-worker" ) )
}


/// Fallback for non-Windows config directory.
fn dirs_fallback() -> Option<PathBuf> {
    std::env::var( "HOME" )
        .ok()
        .map( |h| PathBuf::from( h ).join( ".config" ) )
}


/// Get the full path to the configuration file.
pub fn config_path() -> Result<PathBuf> {
    Ok( config_dir()?.join( "config.toml" ) )
}


/// Load configuration from the TOML file.
///
/// If the config directory or file does not exist, creates them
/// with default values and returns the defaults.
pub fn load_config() -> Result<AppConfig> {
    let dir = config_dir()?;
    let path = dir.join( "config.toml" );

    if !dir.exists() {
        std::fs::create_dir_all( &dir )
            .with_context( || format!( "Failed to create config directory: {}", dir.display() ) )?;
        tracing::info!( "Created config directory: {}", dir.display() );
    }

    if !path.exists() {
        let defaults = AppConfig::default();
        save_config( &defaults )?;
        tracing::info!( "Created default config at: {}", path.display() );
        return Ok( defaults );
    }

    let contents = std::fs::read_to_string( &path )
        .with_context( || format!( "Failed to read config file: {}", path.display() ) )?;

    let config: AppConfig = toml::from_str( &contents )
        .with_context( || format!( "Failed to parse config file: {}", path.display() ) )?;

    Ok( config )
}


/// Resolve bare tool names to full paths using the system PATH.
///
/// For each tool path that is a bare name (no path separators),
/// runs `where.exe <name>` to find the full path. Updates the
/// config in place and saves if any paths were resolved.
pub fn resolve_tool_paths( config: &mut AppConfig ) -> Result<()> {
    let tools = [
        ( "whisper", config.tools.whisper_path.clone() ),
        ( "claude", config.tools.claude_path.clone() ),
        ( "ffmpeg", config.tools.ffmpeg_path.clone() ),
        ( "ffprobe", config.tools.ffprobe_path.clone() ),
        ( "melt", config.tools.melt_path.clone() ),
    ];

    let mut changed = false;

    for ( label, current ) in &tools {
        if current.is_empty() || current.contains( '\\' ) || current.contains( '/' ) {
            continue;
        }

        if let Some( resolved ) = resolve_from_path( current ) {
            tracing::info!( "Resolved {} path: {} -> {}", label, current, resolved );
            match *label {
                "whisper" => config.tools.whisper_path = resolved,
                "claude" => config.tools.claude_path = resolved,
                "ffmpeg" => config.tools.ffmpeg_path = resolved,
                "ffprobe" => config.tools.ffprobe_path = resolved,
                "melt" => config.tools.melt_path = resolved,
                _ => {}
            }
            changed = true;
        }
    }

    if changed {
        save_config( config )?;
    }

    Ok(())
}


/// Use `where.exe` to find an executable on the system PATH.
///
/// @param name - bare executable name (e.g. "ffmpeg")
/// @returns full path if found, None otherwise
fn resolve_from_path( name: &str ) -> Option<String> {
    let output = std::process::Command::new( "where" )
        .arg( name )
        .output()
        .ok()?;

    if !output.status.success() {
        return None;
    }

    let stdout = String::from_utf8_lossy( &output.stdout );
    stdout.lines().next()
        .map( |line| line.trim().to_string() )
        .filter( |p| !p.is_empty() )
}


/// Save configuration to the TOML file.
///
/// @param config - The configuration to persist
pub fn save_config( config: &AppConfig ) -> Result<()> {
    let path = config_path()?;

    let contents = toml::to_string_pretty( config )
        .context( "Failed to serialize config to TOML" )?;

    std::fs::write( &path, contents )
        .with_context( || format!( "Failed to write config file: {}", path.display() ) )?;

    tracing::info!( "Config saved to: {}", path.display() );
    Ok(())
}

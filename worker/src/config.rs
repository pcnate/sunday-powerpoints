use anyhow::{ Context, Result };
use serde::{ Deserialize, Serialize };
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
/// (e.g., 22:00–06:00).
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct ScheduleConfig {
    pub shift_start: String,
    pub shift_end: String,
}


/// Polling and heartbeat timing configuration.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct TimingConfig {
    pub poll_interval_secs: u64,
    pub heartbeat_interval_secs: u64,
    pub job_delay_secs: u64,
}


/// External tool paths and settings.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct ToolsConfig {
    pub whisper_path: String,
    pub whisper_model: String,
    pub claude_path: String,
    pub ffmpeg_path: String,
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
                shift_start: "02:00".to_string(),
                shift_end: "06:00".to_string(),
            },
            timing: TimingConfig {
                poll_interval_secs: 30,
                heartbeat_interval_secs: 30,
                job_delay_secs: 10,
            },
            tools: ToolsConfig {
                whisper_path: "whisper".to_string(),
                whisper_model: "large-v3".to_string(),
                claude_path: "claude".to_string(),
                ffmpeg_path: "ffmpeg".to_string(),
            },
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

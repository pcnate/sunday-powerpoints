use axum::{
    extract::{ Path, Query, State as AxumState },
    http::StatusCode,
    response::{ Html, IntoResponse, Json },
    routing::{ get, post, put },
    Router,
};
use serde::{ Deserialize, Serialize };
use std::collections::HashMap;
use std::sync::Arc;
use tokio::sync::mpsc;
use tower_http::cors::CorsLayer;

use crate::config;
use crate::state::{ AppState, TrayCommand };


/// Shared state for the web UI routes.
struct WebState {
    app_state: Arc<AppState>,
    tray_tx: mpsc::Sender<TrayCommand>,
    http_client: reqwest::Client,
}


/// Status response for GET /api/status.
#[derive( Serialize )]
struct StatusResponse {
    phase: String,
    connected: bool,
    current_job_id: Option<u32>,
    current_job_type: Option<String>,
    jobs_completed: u32,
    jobs_failed: u32,
    uptime_secs: u64,
    worker_id: String,
    server_url: String,
    schedule_enabled: bool,
    shift_start: String,
    shift_end: String,
    force_on_shift: bool,
}


/// Config update request for POST /api/config.
#[derive( Deserialize )]
struct ConfigUpdate {
    server_url: Option<String>,
    worker_id: Option<String>,
    types: Option<Vec<String>>,
    schedule_enabled: Option<bool>,
    shift_start: Option<String>,
    shift_end: Option<String>,
    poll_interval_secs: Option<u64>,
    heartbeat_interval_secs: Option<u64>,
    job_delay_secs: Option<u64>,
    web_ui_port: Option<u16>,
    whisper_path: Option<String>,
    whisper_model: Option<String>,
    whisper_compute_type: Option<String>,
    claude_path: Option<String>,
    ffmpeg_path: Option<String>,
    ffprobe_path: Option<String>,
    melt_path: Option<String>,
    melt_video_bitrate: Option<String>,
    melt_audio_bitrate: Option<String>,
    output_directory: Option<String>,
    tool_env: Option<HashMap<String, HashMap<String, String>>>,
}


/// Browse request for POST /api/browse (directory listing).
#[derive( Deserialize )]
struct BrowseRequest {
    path: Option<String>,
    extensions: Option<Vec<String>>,
}


/// A single entry in a directory listing.
#[derive( Serialize )]
struct BrowseEntry {
    name: String,
    is_dir: bool,
    size: u64,
}


/// Browse response for POST /api/browse (directory listing).
#[derive( Serialize )]
struct BrowseResponse {
    path: String,
    entries: Vec<BrowseEntry>,
}


/// Test tool request for POST /api/test-tool.
#[derive( Deserialize )]
struct TestToolRequest {
    command: String,
    args: Option<Vec<String>>,
    env: Option<HashMap<String, String>>,
}


/// Test tool response for POST /api/test-tool.
#[derive( Serialize )]
struct TestToolResponse {
    success: bool,
    output: Option<String>,
    error: Option<String>,
}


/// Query params for GET /api/queue.
#[derive( Deserialize )]
struct QueueQuery {
    status: Option<String>,
    #[serde( rename = "type" )]
    job_type: Option<String>,
    limit: Option<u32>,
}


/// Cancel request for PUT /api/queue/:id/cancel.
#[derive( Serialize )]
struct CancelBody {
    status: String,
}


/// Start the embedded axum web server for configuration.
///
/// @param app_state - shared application state
/// @param tray_tx - sender to push commands to the scheduler
/// @param port - port to listen on
pub async fn run(
    app_state: Arc<AppState>,
    tray_tx: mpsc::Sender<TrayCommand>,
    port: u16,
) {
    let http_client = reqwest::Client::builder()
        .timeout( std::time::Duration::from_secs( 10 ) )
        .build()
        .expect( "Failed to build HTTP client for queue proxy" );

    let shared = Arc::new( WebState { app_state, tray_tx, http_client } );

    let app = Router::new()
        .route( "/", get( index_handler ) )
        .route( "/api/status", get( status_handler ) )
        .route( "/api/config", get( get_config_handler ) )
        .route( "/api/config", post( update_config_handler ) )
        .route( "/api/queue", get( queue_handler ) )
        .route( "/api/queue/{id}", get( job_detail_handler ) )
        .route( "/api/queue/{id}/cancel", put( cancel_job_handler ) )
        .route( "/api/scheduler/pause", post( pause_handler ) )
        .route( "/api/scheduler/run-once", post( run_once_handler ) )
        .route( "/api/scheduler/force-on-shift", post( force_on_shift_handler ) )
        .route( "/api/browse", post( browse_handler ) )
        .route( "/api/test-tool", post( test_tool_handler ) )
        .layer( CorsLayer::permissive() )
        .with_state( shared );

    let addr = format!( "0.0.0.0:{}", port );
    tracing::info!( "Web UI listening on http://localhost:{}", port );

    let listener = tokio::net::TcpListener::bind( &addr ).await
        .expect( "Failed to bind web UI port" );

    axum::serve( listener, app ).await
        .expect( "Web UI server error" );
}


/// GET / — Serve the embedded HTML config page.
async fn index_handler() -> Html<&'static str> {
    Html( INDEX_HTML )
}


/// GET /api/status — Return current worker status as JSON.
async fn status_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> Json<StatusResponse> {
    let snapshot = state.app_state.snapshot().await;
    let config = state.app_state.config.read().await;

    Json( StatusResponse {
        phase: snapshot.phase.to_string(),
        connected: snapshot.connected,
        current_job_id: snapshot.current_job_id,
        current_job_type: snapshot.current_job_type,
        jobs_completed: snapshot.jobs_completed,
        jobs_failed: snapshot.jobs_failed,
        uptime_secs: snapshot.uptime_secs,
        worker_id: config.worker.id.clone(),
        server_url: config.server.url.clone(),
        schedule_enabled: config.schedule.enabled,
        shift_start: config.schedule.shift_start.clone(),
        shift_end: config.schedule.shift_end.clone(),
        force_on_shift: snapshot.force_on_shift,
    } )
}


/// GET /api/config — Return current configuration as JSON.
///
/// If `output_directory` is empty, fills it with the resolved
/// `OneDriveConsumer` environment variable so the UI shows the effective path.
async fn get_config_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> impl IntoResponse {
    let mut config = state.app_state.config.read().await.clone();

    // Populate output_directory with the env var fallback so the UI shows the effective value
    if config.paths.output_directory.is_empty() {
        config.paths.output_directory = std::env::var( "OneDriveConsumer" ).unwrap_or_default();
    }

    Json( config )
}


/// POST /api/config — Update configuration and save to disk.
async fn update_config_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Json( update ): Json<ConfigUpdate>,
) -> impl IntoResponse {
    let mut config = state.app_state.config.write().await;

    // Apply updates
    if let Some( url ) = update.server_url {
        config.server.url = url;
    }
    if let Some( id ) = update.worker_id {
        config.worker.id = id;
    }
    if let Some( types ) = update.types {
        config.worker.types = types;
    }
    if let Some( enabled ) = update.schedule_enabled {
        config.schedule.enabled = enabled;
    }
    if let Some( start ) = update.shift_start {
        config.schedule.shift_start = start;
    }
    if let Some( end ) = update.shift_end {
        config.schedule.shift_end = end;
    }
    if let Some( v ) = update.poll_interval_secs {
        config.timing.poll_interval_secs = v;
    }
    if let Some( v ) = update.heartbeat_interval_secs {
        config.timing.heartbeat_interval_secs = v;
    }
    if let Some( v ) = update.job_delay_secs {
        config.timing.job_delay_secs = v;
    }
    if let Some( v ) = update.web_ui_port {
        config.web_ui.port = v;
    }
    if let Some( v ) = update.whisper_path {
        config.tools.whisper_path = v;
    }
    if let Some( v ) = update.whisper_model {
        config.tools.whisper_model = v;
    }
    if let Some( v ) = update.whisper_compute_type {
        config.tools.whisper_compute_type = v;
    }
    if let Some( v ) = update.claude_path {
        config.tools.claude_path = v;
    }
    if let Some( v ) = update.ffmpeg_path {
        config.tools.ffmpeg_path = v;
    }
    if let Some( v ) = update.ffprobe_path {
        config.tools.ffprobe_path = v;
    }
    if let Some( v ) = update.melt_path {
        config.tools.melt_path = v;
    }
    if let Some( v ) = update.melt_video_bitrate {
        config.tools.melt_video_bitrate = v;
    }
    if let Some( v ) = update.melt_audio_bitrate {
        config.tools.melt_audio_bitrate = v;
    }
    if let Some( v ) = update.output_directory {
        config.paths.output_directory = v;
    }
    if let Some( v ) = update.tool_env {
        config.tools.tool_env = v;
    }

    // Save to disk
    match config::save_config( &config ) {
        Ok( _ ) => {
            tracing::info!( "Config updated and saved" );
            // Signal scheduler to reload
            let _ = state.tray_tx.try_send( TrayCommand::ReloadConfig );
            ( StatusCode::OK, Json( serde_json::json!({ "success": true }) ) )
        }
        Err( e ) => {
            tracing::error!( "Failed to save config: {}", e );
            ( StatusCode::INTERNAL_SERVER_ERROR, Json( serde_json::json!({ "error": format!( "{}", e ) }) ) )
        }
    }
}


/// GET /api/queue — Proxy to Express GET /api/jobs with filtering.
async fn queue_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Query( params ): Query<QueueQuery>,
) -> impl IntoResponse {
    let config = state.app_state.config.read().await;
    let mut url = format!( "{}/api/jobs?", config.server.url.trim_end_matches( '/' ) );

    if let Some( ref status ) = params.status {
        url.push_str( &format!( "status={}&", status ) );
    }
    if let Some( ref job_type ) = params.job_type {
        url.push_str( &format!( "type={}&", job_type ) );
    }
    let limit = params.limit.unwrap_or( 50 );
    url.push_str( &format!( "limit={}", limit ) );

    match state.http_client.get( &url ).send().await {
        Ok( resp ) if resp.status().is_success() => {
            match resp.json::<serde_json::Value>().await {
                Ok( body ) => ( StatusCode::OK, Json( body ) ).into_response(),
                Err( e ) => (
                    StatusCode::BAD_GATEWAY,
                    Json( serde_json::json!({ "error": format!( "Parse error: {}", e ) }) ),
                ).into_response(),
            }
        }
        Ok( resp ) => (
            StatusCode::BAD_GATEWAY,
            Json( serde_json::json!({ "error": format!( "Server returned {}", resp.status() ) }) ),
        ).into_response(),
        Err( e ) => (
            StatusCode::BAD_GATEWAY,
            Json( serde_json::json!({ "error": format!( "Connection failed: {}", e ) }) ),
        ).into_response(),
    }
}


/// GET /api/queue/:id — Proxy to Express GET /api/jobs/:id (job + logs).
async fn job_detail_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Path( id ): Path<u32>,
) -> impl IntoResponse {
    let config = state.app_state.config.read().await;
    let url = format!(
        "{}/api/jobs/{}",
        config.server.url.trim_end_matches( '/' ),
        id
    );

    match state.http_client.get( &url ).send().await {
        Ok( resp ) if resp.status().is_success() => {
            match resp.json::<serde_json::Value>().await {
                Ok( body ) => ( StatusCode::OK, Json( body ) ).into_response(),
                Err( e ) => (
                    StatusCode::BAD_GATEWAY,
                    Json( serde_json::json!({ "error": format!( "Parse error: {}", e ) }) ),
                ).into_response(),
            }
        }
        Ok( resp ) => (
            StatusCode::BAD_GATEWAY,
            Json( serde_json::json!({ "error": format!( "Server returned {}", resp.status() ) }) ),
        ).into_response(),
        Err( e ) => (
            StatusCode::BAD_GATEWAY,
            Json( serde_json::json!({ "error": format!( "Connection failed: {}", e ) }) ),
        ).into_response(),
    }
}


/// PUT /api/queue/:id/cancel — Proxy to Express PUT /api/jobs/:id to cancel.
async fn cancel_job_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Path( id ): Path<u32>,
) -> impl IntoResponse {
    let config = state.app_state.config.read().await;
    let url = format!(
        "{}/api/jobs/{}",
        config.server.url.trim_end_matches( '/' ),
        id
    );

    let body = CancelBody { status: "cancelled".to_string() };

    match state.http_client.put( &url ).json( &body ).send().await {
        Ok( resp ) if resp.status().is_success() => {
            match resp.json::<serde_json::Value>().await {
                Ok( body ) => ( StatusCode::OK, Json( body ) ).into_response(),
                Err( e ) => (
                    StatusCode::BAD_GATEWAY,
                    Json( serde_json::json!({ "error": format!( "Parse error: {}", e ) }) ),
                ).into_response(),
            }
        }
        Ok( resp ) => {
            let status = resp.status();
            let text = resp.text().await.unwrap_or_default();
            (
                StatusCode::from_u16( status.as_u16() ).unwrap_or( StatusCode::BAD_GATEWAY ),
                Json( serde_json::json!({ "error": text }) ),
            ).into_response()
        }
        Err( e ) => (
            StatusCode::BAD_GATEWAY,
            Json( serde_json::json!({ "error": format!( "Connection failed: {}", e ) }) ),
        ).into_response(),
    }
}


/// POST /api/scheduler/pause — Toggle scheduler pause state.
async fn pause_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> impl IntoResponse {
    let _ = state.tray_tx.try_send( TrayCommand::TogglePause );
    Json( serde_json::json!({ "success": true }) )
}


/// POST /api/scheduler/run-once — Request a single job poll cycle.
async fn run_once_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> impl IntoResponse {
    let _ = state.tray_tx.try_send( TrayCommand::RunOnce );
    Json( serde_json::json!({ "success": true }) )
}


/// POST /api/scheduler/force-on-shift — Toggle force on-shift override.
async fn force_on_shift_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> impl IntoResponse {
    let _ = state.tray_tx.try_send( TrayCommand::ToggleForceOnShift );
    // Read the new value after a brief yield to let the scheduler process it
    tokio::time::sleep( std::time::Duration::from_millis( 50 ) ).await;
    let forced = *state.app_state.force_on_shift.read().await;
    Json( serde_json::json!({ "success": true, "force_on_shift": forced }) )
}


/// POST /api/browse — List directory contents or enumerate drives.
///
/// If `path` is None/empty, returns available Windows drive letters.
/// Otherwise lists the directory contents, applying optional extension filters.
async fn browse_handler(
    Json( req ): Json<BrowseRequest>,
) -> impl IntoResponse {
    let path_str = req.path.unwrap_or_default();

    // No path → enumerate Windows drives
    if path_str.is_empty() {
        let mut drives = Vec::new();
        for letter in b'A'..=b'Z' {
            let drive = format!( "{}:\\", letter as char );
            if std::path::Path::new( &drive ).exists() {
                drives.push( BrowseEntry {
                    name: drive,
                    is_dir: true,
                    size: 0,
                } );
            }
        }
        return ( StatusCode::OK, Json( BrowseResponse {
            path: String::new(),
            entries: drives,
        } ) ).into_response();
    }

    // Normalize the path
    let dir = std::path::Path::new( &path_str );
    if !dir.exists() || !dir.is_dir() {
        return ( StatusCode::BAD_REQUEST, Json( serde_json::json!({
            "error": format!( "Invalid directory: {}", path_str )
        }) ) ).into_response();
    }

    let ext_filter: Option<Vec<String>> = req.extensions.map( |exts|
        exts.iter().map( |e| e.to_lowercase() ).collect()
    );

    let read_dir = match std::fs::read_dir( dir ) {
        Ok( rd ) => rd,
        Err( e ) => {
            return ( StatusCode::BAD_REQUEST, Json( serde_json::json!({
                "error": format!( "Cannot read directory: {}", e )
            }) ) ).into_response();
        }
    };

    let mut entries: Vec<BrowseEntry> = Vec::new();
    for entry in read_dir.flatten() {
        let metadata = match entry.metadata() {
            Ok( m ) => m,
            Err( _ ) => continue,
        };

        let name = entry.file_name().to_string_lossy().to_string();
        let is_dir = metadata.is_dir();

        // Apply extension filter to files only
        if !is_dir
            && let Some( ref exts ) = ext_filter
        {
            let file_ext = std::path::Path::new( &name )
                .extension()
                .map( |e| e.to_string_lossy().to_lowercase() )
                .unwrap_or_default();
            if !exts.contains( &file_ext ) {
                continue;
            }
        }

        entries.push( BrowseEntry {
            name,
            is_dir,
            size: metadata.len(),
        } );
    }

    // Sort: directories first, then alphabetical (case-insensitive)
    entries.sort_by( |a, b| {
        b.is_dir.cmp( &a.is_dir )
            .then_with( || a.name.to_lowercase().cmp( &b.name.to_lowercase() ) )
    } );

    ( StatusCode::OK, Json( BrowseResponse {
        path: dir.to_string_lossy().to_string(),
        entries,
    } ) ).into_response()
}


/// POST /api/test-tool — Spawn a command with a 10-second timeout.
async fn test_tool_handler(
    Json( req ): Json<TestToolRequest>,
) -> impl IntoResponse {
    let args = req.args.unwrap_or_default();

    let mut cmd = tokio::process::Command::new( &req.command );
    cmd.args( &args )
        .stdout( std::process::Stdio::piped() )
        .stderr( std::process::Stdio::piped() )
        .kill_on_drop( true );

    if let Some( ref env_vars ) = req.env {
        for ( key, value ) in env_vars {
            cmd.env( key, value );
        }
    }

    let child = cmd.spawn();

    let child = match child {
        Ok( c ) => c,
        Err( e ) => {
            return Json( TestToolResponse {
                success: false,
                output: None,
                error: Some( format!( "Failed to spawn: {}", e ) ),
            } );
        }
    };

    let timeout = tokio::time::timeout(
        std::time::Duration::from_secs( 10 ),
        child.wait_with_output(),
    ).await;

    match timeout {
        Ok( Ok( output ) ) => {
            let stdout = String::from_utf8_lossy( &output.stdout ).to_string();
            let stderr = String::from_utf8_lossy( &output.stderr ).to_string();
            let combined = if stderr.is_empty() {
                stdout
            } else if stdout.is_empty() {
                stderr
            } else {
                format!( "{}\n{}", stdout, stderr )
            };

            Json( TestToolResponse {
                success: output.status.success(),
                output: Some( combined ),
                error: None,
            } )
        }
        Ok( Err( e ) ) => Json( TestToolResponse {
            success: false,
            output: None,
            error: Some( format!( "Process error: {}", e ) ),
        } ),
        Err( _ ) => {
            // child is dropped here, kill_on_drop handles cleanup
            Json( TestToolResponse {
                success: false,
                output: None,
                error: Some( "Command timed out after 10 seconds".to_string() ),
            } )
        }
    }
}


/// Embedded HTML for the config web UI.
///
/// A self-contained single page with inline CSS and JavaScript.
/// No external dependencies or build step needed.
const INDEX_HTML: &str = r#"<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sunday Worker</title>
<style>
  :root {
    --bg: #1e1e2e; --surface: #313244; --text: #cdd6f4;
    --subtext: #a6adc8; --accent: #89b4fa; --green: #a6e3a1;
    --yellow: #f9e2af; --red: #f38ba8; --border: #45475a;
    --surface2: #585b70;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--bg); color: var(--text);
    max-width: 800px; margin: 0 auto; padding: 24px;
  }
  h1 { font-size: 1.5rem; margin-bottom: 8px; }
  h2 { font-size: 1.1rem; margin: 24px 0 12px; color: var(--accent); }
  .tabs {
    display: flex; gap: 4px; margin-bottom: 20px;
    border-bottom: 2px solid var(--border); padding-bottom: 0;
  }
  .tab {
    padding: 8px 20px; background: none; color: var(--subtext);
    border: none; cursor: pointer; font-size: 0.95rem;
    border-bottom: 2px solid transparent; margin-bottom: -2px;
  }
  .tab:hover { color: var(--text); }
  .tab.active {
    color: var(--accent); border-bottom-color: var(--accent);
    font-weight: 600;
  }
  .tab-content { display: none; }
  .tab-content.active { display: block; }
  .status-bar {
    display: flex; gap: 16px; align-items: center;
    padding: 12px 16px; background: var(--surface);
    border-radius: 8px; margin-bottom: 20px;
  }
  .status-dot {
    width: 12px; height: 12px; border-radius: 50%; flex-shrink: 0;
  }
  .dot-green { background: var(--green); }
  .dot-yellow { background: var(--yellow); }
  .dot-gray { background: var(--subtext); }
  .dot-red { background: var(--red); }
  .status-text { flex: 1; }
  .status-label { font-size: 0.85rem; color: var(--subtext); }
  .status-actions { display: flex; gap: 8px; }
  .stats { display: flex; gap: 16px; font-size: 0.85rem; color: var(--subtext); }
  .form-group { margin-bottom: 12px; }
  label {
    display: block; font-size: 0.85rem; color: var(--subtext);
    margin-bottom: 4px;
  }
  input, select {
    width: 100%; padding: 8px 12px;
    background: var(--surface); color: var(--text);
    border: 1px solid var(--border); border-radius: 6px;
    font-size: 0.95rem;
  }
  input:focus, select:focus { outline: none; border-color: var(--accent); }
  .row { display: flex; gap: 12px; }
  .row > .form-group { flex: 1; }
  .checkbox-group {
    display: flex; gap: 16px; padding: 8px 0;
  }
  .checkbox-group label {
    display: flex; align-items: center; gap: 6px;
    cursor: pointer; color: var(--text); font-size: 0.95rem;
  }
  .checkbox-group input[type="checkbox"] {
    width: 16px; height: 16px; accent-color: var(--accent);
  }
  button, .btn {
    padding: 8px 20px; background: var(--accent); color: var(--bg);
    border: none; border-radius: 6px; cursor: pointer;
    font-size: 0.9rem; font-weight: 600;
  }
  button:hover, .btn:hover { opacity: 0.9; }
  .btn-sm {
    padding: 4px 12px; font-size: 0.8rem; border-radius: 4px;
  }
  .btn-danger { background: var(--red); }
  .btn-warn { background: var(--yellow); color: #1e1e2e; }
  .btn-outline {
    background: transparent; border: 1px solid var(--border);
    color: var(--text);
  }
  .btn-outline:hover { border-color: var(--accent); }
  .msg {
    margin-top: 8px; font-size: 0.85rem; padding: 8px;
    border-radius: 4px; display: none;
  }
  .msg-ok { background: rgba(166,227,161,0.15); color: var(--green); display: block; }
  .msg-err { background: rgba(243,139,168,0.15); color: var(--red); display: block; }

  /* Queue table */
  .queue-controls {
    display: flex; gap: 8px; align-items: center;
    margin-bottom: 12px; flex-wrap: wrap;
  }
  .queue-controls select {
    width: auto; padding: 6px 10px; font-size: 0.85rem;
  }
  table {
    width: 100%; border-collapse: collapse; font-size: 0.85rem;
  }
  th {
    text-align: left; padding: 8px 10px;
    color: var(--subtext); border-bottom: 2px solid var(--border);
    font-weight: 600; font-size: 0.8rem; text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  td {
    padding: 8px 10px; border-bottom: 1px solid var(--border);
    vertical-align: middle;
  }
  tr:hover td { background: rgba(137,180,250,0.05); }
  .badge {
    display: inline-block; padding: 2px 8px; border-radius: 10px;
    font-size: 0.75rem; font-weight: 600;
  }
  .badge-pending { background: rgba(137,180,250,0.2); color: var(--accent); }
  .badge-processing { background: rgba(249,226,175,0.2); color: var(--yellow); }
  .badge-completed { background: rgba(166,227,161,0.2); color: var(--green); }
  .badge-failed { background: rgba(243,139,168,0.2); color: var(--red); }
  .badge-cancelled { background: rgba(158,158,158,0.2); color: var(--subtext); }
  .empty-state {
    text-align: center; padding: 32px; color: var(--subtext);
    font-size: 0.9rem;
  }
</style>
</head>
<body>
<h1>Sunday Worker</h1>

<div class="status-bar">
  <div id="dot" class="status-dot dot-gray"></div>
  <div class="status-text">
    <div id="phase">Loading...</div>
    <div class="status-label" id="job-info"></div>
  </div>
  <div class="status-actions">
    <button class="btn btn-sm btn-outline" id="pauseBtn" onclick="togglePause()">Pause</button>
    <button class="btn btn-sm btn-outline" id="forceBtn" onclick="toggleForceOnShift()">Force On-Shift</button>
  </div>
</div>
<div class="stats" style="margin-bottom:20px">
  <span id="completed">0 completed</span>
  <span id="failed">0 failed</span>
  <span id="uptime"></span>
</div>

<div class="tabs">
  <button class="tab active" onclick="switchTab('queue')">Job Queue</button>
  <button class="tab" onclick="switchTab('config')">Configuration</button>
</div>

<!-- Queue Tab -->
<div id="tab-queue" class="tab-content active">
  <div class="queue-controls">
    <select id="queueFilter" onchange="loadQueue()">
      <option value="">All Jobs</option>
      <option value="pending">Pending</option>
      <option value="processing">Processing</option>
      <option value="completed">Completed</option>
      <option value="failed">Failed</option>
      <option value="cancelled">Cancelled</option>
      <option value="pending,processing" selected>Active (Pending + Processing)</option>
    </select>
    <select id="queueTypeFilter" onchange="loadQueue()">
      <option value="">All Types</option>
      <option value="ffprobe">FFprobe</option>
      <option value="transcode">Transcode</option>
      <option value="transcription">Transcription</option>
      <option value="claude-processing">Claude Processing</option>
    </select>
    <button class="btn btn-sm btn-outline" onclick="loadQueue()">Refresh</button>
  </div>
  <div id="queueTable"></div>
</div>

<!-- Config Tab -->
<div id="tab-config" class="tab-content">
  <h2>Server</h2>
  <div class="form-group">
    <label for="server_url">Server URL</label>
    <input id="server_url" type="text">
  </div>

  <h2>Worker</h2>
  <div class="form-group">
    <label for="worker_id">Worker ID</label>
    <input id="worker_id" type="text">
  </div>
  <div class="form-group">
    <label for="output_directory">Output Directory (sermons root folder)</label>
    <input id="output_directory" type="text" placeholder="%OneDriveConsumer%">
  </div>
  <div class="form-group">
    <label>Job Types</label>
    <div class="checkbox-group">
      <label><input type="checkbox" id="type_ffprobe" value="ffprobe"> FFprobe</label>
      <label><input type="checkbox" id="type_transcode" value="transcode"> Transcode</label>
      <label><input type="checkbox" id="type_transcription" value="transcription"> Transcription</label>
      <label><input type="checkbox" id="type_claude" value="claude-processing"> Claude Processing</label>
    </div>
  </div>

  <h2>Schedule</h2>
  <div class="checkbox-group" style="margin-bottom:8px">
    <label><input type="checkbox" id="schedule_enabled"> Enabled</label>
  </div>
  <div class="row" id="shift_row">
    <div class="form-group">
      <label for="shift_start">Shift Start</label>
      <input id="shift_start" type="text" placeholder="HH:MM">
    </div>
    <div class="form-group">
      <label for="shift_end">Shift End</label>
      <input id="shift_end" type="text" placeholder="HH:MM">
    </div>
  </div>
  <div id="run_once_row" style="display:none;margin-bottom:12px">
    <button class="btn btn-sm" onclick="runOnce()">Run Next Job</button>
  </div>

  <h2>Timing</h2>
  <div class="row">
    <div class="form-group">
      <label for="poll_interval">Poll Interval (s)</label>
      <input id="poll_interval" type="number">
    </div>
    <div class="form-group">
      <label for="heartbeat_interval">Heartbeat (s)</label>
      <input id="heartbeat_interval" type="number">
    </div>
    <div class="form-group">
      <label for="job_delay">Job Delay (s)</label>
      <input id="job_delay" type="number">
    </div>
  </div>

  <h2>Tool Paths</h2>
  <div class="row">
    <div class="form-group">
      <label for="whisper_path">Whisper</label>
      <input id="whisper_path" type="text">
    </div>
    <div class="form-group">
      <label for="whisper_model">Model</label>
      <input id="whisper_model" type="text">
    </div>
    <div class="form-group">
      <label for="whisper_compute_type">Compute Type</label>
      <select id="whisper_compute_type">
        <option value="float16">float16</option>
        <option value="float32">float32</option>
        <option value="int8">int8</option>
        <option value="int8_float16">int8_float16</option>
        <option value="int8_float32">int8_float32</option>
        <option value="auto">auto</option>
        <option value="default">default</option>
      </select>
    </div>
  </div>
  <div class="form-group">
    <label for="env_whisper">Whisper Env Vars (KEY=VALUE per line)</label>
    <textarea id="env_whisper" rows="2" style="width:100%;padding:8px 12px;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:6px;font-size:0.85rem;font-family:monospace;resize:vertical" placeholder="PYTHONIOENCODING=utf-8"></textarea>
  </div>
  <div class="row">
    <div class="form-group">
      <label for="claude_path">Claude</label>
      <input id="claude_path" type="text">
    </div>
    <div class="form-group">
      <label for="ffmpeg_path">FFmpeg</label>
      <input id="ffmpeg_path" type="text">
    </div>
    <div class="form-group">
      <label for="ffprobe_path">FFprobe</label>
      <input id="ffprobe_path" type="text">
    </div>
  </div>
  <div class="row">
    <div class="form-group">
      <label for="melt_path">Melt (Kdenlive)</label>
      <input id="melt_path" type="text">
    </div>
    <div class="form-group">
      <label for="melt_video_bitrate">Video Bitrate</label>
      <input id="melt_video_bitrate" type="text" placeholder="5000k">
    </div>
    <div class="form-group">
      <label for="melt_audio_bitrate">Audio Bitrate</label>
      <input id="melt_audio_bitrate" type="text" placeholder="192k">
    </div>
  </div>

  <button onclick="saveConfig()">Save Configuration</button>
  <div id="msg" class="msg"></div>
</div>

<script>
// --- Tabs ---
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t,i) => {
    t.classList.toggle('active', t.textContent.toLowerCase().includes(name));
  });
  document.querySelectorAll('.tab-content').forEach(tc => {
    tc.classList.toggle('active', tc.id === 'tab-' + name);
  });
  if (name === 'queue') loadQueue();
}

// --- Status ---
async function loadStatus() {
  try {
    const r = await fetch('/api/status');
    const s = await r.json();
    const dot = document.getElementById('dot');
    dot.className = 'status-dot ' + (
      !s.connected ? 'dot-red' :
      s.phase === 'Executing' ? 'dot-yellow' :
      ['Idle','Polling','Cooldown'].includes(s.phase) ? 'dot-green' : 'dot-gray'
    );
    document.getElementById('phase').textContent = s.phase;
    document.getElementById('job-info').textContent =
      s.current_job_id ? 'Job #' + s.current_job_id + ' (' + s.current_job_type + ')' : '';
    document.getElementById('completed').textContent = s.jobs_completed + ' completed';
    document.getElementById('failed').textContent = s.jobs_failed + ' failed';

    const hrs = Math.floor(s.uptime_secs / 3600);
    const mins = Math.floor((s.uptime_secs % 3600) / 60);
    document.getElementById('uptime').textContent = 'Uptime: ' + hrs + 'h ' + mins + 'm';

    const btn = document.getElementById('pauseBtn');
    if (s.phase === 'Paused') {
      btn.textContent = 'Resume';
      btn.className = 'btn btn-sm btn-warn';
    } else {
      btn.textContent = 'Pause';
      btn.className = 'btn btn-sm btn-outline';
    }

    const forceBtn = document.getElementById('forceBtn');
    if (s.force_on_shift) {
      forceBtn.textContent = 'Stop Forcing';
      forceBtn.className = 'btn btn-sm btn-warn';
    } else {
      forceBtn.textContent = 'Force On-Shift';
      forceBtn.className = 'btn btn-sm btn-outline';
    }
  } catch(e) { console.error('Status fetch failed', e); }
}

async function togglePause() {
  await fetch('/api/scheduler/pause', { method: 'POST' });
  setTimeout(loadStatus, 300);
}

async function toggleForceOnShift() {
  await fetch('/api/scheduler/force-on-shift', { method: 'POST' });
  setTimeout(loadStatus, 300);
}

// --- Config ---
async function loadConfig() {
  try {
    const r = await fetch('/api/config');
    const c = await r.json();
    document.getElementById('server_url').value = c.server.url;
    document.getElementById('worker_id').value = c.worker.id;

    // Set checkboxes
    const types = c.worker.types || [];
    document.getElementById('type_transcode').checked = types.includes('transcode');
    document.getElementById('type_transcription').checked = types.includes('transcription');
    document.getElementById('type_claude').checked = types.includes('claude-processing');
    document.getElementById('type_ffprobe').checked = types.includes('ffprobe');

    var scheduleEnabled = c.schedule.enabled !== false;
    document.getElementById('schedule_enabled').checked = scheduleEnabled;
    document.getElementById('shift_start').value = c.schedule.shift_start;
    document.getElementById('shift_end').value = c.schedule.shift_end;
    updateScheduleUI(scheduleEnabled);

    document.getElementById('poll_interval').value = c.timing.poll_interval_secs;
    document.getElementById('heartbeat_interval').value = c.timing.heartbeat_interval_secs;
    document.getElementById('job_delay').value = c.timing.job_delay_secs;
    document.getElementById('whisper_path').value = c.tools.whisper_path;
    document.getElementById('whisper_model').value = c.tools.whisper_model;
    document.getElementById('whisper_compute_type').value = c.tools.whisper_compute_type || 'float16';
    document.getElementById('claude_path').value = c.tools.claude_path;
    document.getElementById('ffmpeg_path').value = c.tools.ffmpeg_path;
    document.getElementById('ffprobe_path').value = c.tools.ffprobe_path || '';
    document.getElementById('output_directory').value = (c.paths && c.paths.output_directory) || '';
    document.getElementById('melt_path').value = c.tools.melt_path;
    document.getElementById('melt_video_bitrate').value = c.tools.melt_video_bitrate || '5000k';
    document.getElementById('melt_audio_bitrate').value = c.tools.melt_audio_bitrate || '192k';

    // Load per-tool env vars
    var toolEnv = (c.tools && c.tools.tool_env) || {};
    document.getElementById('env_whisper').value = envMapToText(toolEnv.whisper || {});
  } catch(e) { console.error('Config fetch failed', e); }
}

function envMapToText(map) {
  return Object.entries(map).map(function(e) { return e[0] + '=' + e[1]; }).join('\n');
}

function textToEnvMap(text) {
  var map = {};
  text.split('\n').forEach(function(line) {
    var trimmed = line.trim();
    if (!trimmed || trimmed.indexOf('=') < 0) return;
    var idx = trimmed.indexOf('=');
    var key = trimmed.substring(0, idx).trim();
    var val = trimmed.substring(idx + 1).trim();
    if (key) map[key] = val;
  });
  return map;
}

function updateScheduleUI(enabled) {
  document.getElementById('shift_row').style.opacity = enabled ? '1' : '0.5';
  document.getElementById('shift_start').disabled = !enabled;
  document.getElementById('shift_end').disabled = !enabled;
  document.getElementById('run_once_row').style.display = enabled ? 'none' : 'block';
}

document.addEventListener('DOMContentLoaded', function() {
  var cb = document.getElementById('schedule_enabled');
  if (cb) cb.addEventListener('change', function() { updateScheduleUI(this.checked); });
});

async function runOnce() {
  try {
    await fetch('/api/scheduler/run-once', { method: 'POST' });
  } catch(e) { console.error('Run once failed', e); }
}

function getSelectedTypes() {
  const types = [];
  if (document.getElementById('type_transcode').checked) types.push('transcode');
  if (document.getElementById('type_transcription').checked) types.push('transcription');
  if (document.getElementById('type_claude').checked) types.push('claude-processing');
  if (document.getElementById('type_ffprobe').checked) types.push('ffprobe');
  return types;
}

async function saveConfig() {
  const msg = document.getElementById('msg');
  try {
    const body = {
      server_url: document.getElementById('server_url').value,
      worker_id: document.getElementById('worker_id').value,
      types: getSelectedTypes(),
      schedule_enabled: document.getElementById('schedule_enabled').checked,
      shift_start: document.getElementById('shift_start').value,
      shift_end: document.getElementById('shift_end').value,
      poll_interval_secs: parseInt(document.getElementById('poll_interval').value),
      heartbeat_interval_secs: parseInt(document.getElementById('heartbeat_interval').value),
      job_delay_secs: parseInt(document.getElementById('job_delay').value),
      whisper_path: document.getElementById('whisper_path').value,
      whisper_model: document.getElementById('whisper_model').value,
      whisper_compute_type: document.getElementById('whisper_compute_type').value,
      claude_path: document.getElementById('claude_path').value,
      ffmpeg_path: document.getElementById('ffmpeg_path').value,
      ffprobe_path: document.getElementById('ffprobe_path').value,
      output_directory: document.getElementById('output_directory').value,
      melt_path: document.getElementById('melt_path').value,
      melt_video_bitrate: document.getElementById('melt_video_bitrate').value,
      melt_audio_bitrate: document.getElementById('melt_audio_bitrate').value,
      tool_env: {
        whisper: textToEnvMap(document.getElementById('env_whisper').value),
      },
    };
    const r = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (r.ok) {
      msg.textContent = 'Configuration saved successfully.';
      msg.className = 'msg msg-ok';
    } else {
      const e = await r.json();
      msg.textContent = 'Error: ' + (e.error || 'Unknown');
      msg.className = 'msg msg-err';
    }
  } catch(e) {
    msg.textContent = 'Network error: ' + e.message;
    msg.className = 'msg msg-err';
  }
  setTimeout(function() { msg.className = 'msg'; }, 4000);
}

// --- Queue ---
async function loadQueue() {
  const container = document.getElementById('queueTable');
  const status = document.getElementById('queueFilter').value;
  const type = document.getElementById('queueTypeFilter').value;

  let url = '/api/queue?limit=50';
  if (status) url += '&status=' + status;
  if (type) url += '&type=' + type;

  try {
    const r = await fetch(url);
    if (!r.ok) {
      container.innerHTML = '<div class="empty-state">Could not reach Express server</div>';
      return;
    }
    const data = await r.json();
    const jobs = data.jobs || [];

    if (jobs.length === 0) {
      container.innerHTML = '<div class="empty-state">No jobs found</div>';
      return;
    }

    let html = '<table><thead><tr>';
    html += '<th>ID</th><th>Type</th><th>Date</th><th>Status</th>';
    html += '<th>Worker</th><th>Retry</th><th>Created</th><th></th>';
    html += '</tr></thead><tbody>';

    for (const j of jobs) {
      const badge = 'badge-' + j.status;
      const typeLabel = j.type.replace('video-', 'v-').replace('claude-', 'c-');
      const created = j.created_at ? new Date(j.created_at).toLocaleDateString() : '';
      const canCancel = ['pending', 'processing'].includes(j.status);

      html += '<tr>';
      html += '<td>#' + j.id + '</td>';
      html += '<td>' + typeLabel + '</td>';
      html += '<td>' + j.sunday_date + '</td>';
      html += '<td><span class="badge ' + badge + '">' + j.status + '</span></td>';
      html += '<td>' + (j.worker_id || '-') + '</td>';
      html += '<td>' + j.retry_count + '/' + j.max_retries + '</td>';
      html += '<td>' + created + '</td>';
      html += '<td>';
      if (canCancel) {
        html += '<button class="btn btn-sm btn-danger" onclick="cancelJob(' + j.id + ')">Cancel</button>';
      }
      html += '</td>';
      html += '</tr>';

      if (j.error_message) {
        html += '<tr><td colspan="8" style="padding:4px 10px 8px 32px;color:var(--red);font-size:0.8rem">';
        html += j.error_message.substring(0, 200);
        html += '</td></tr>';
      }
    }

    html += '</tbody></table>';
    html += '<div style="margin-top:8px;font-size:0.8rem;color:var(--subtext)">Showing ' + jobs.length + ' of ' + data.total + ' jobs</div>';
    container.innerHTML = html;
  } catch(e) {
    container.innerHTML = '<div class="empty-state">Error loading queue: ' + e.message + '</div>';
  }
}

async function cancelJob(id) {
  if (!confirm('Cancel job #' + id + '?')) return;
  try {
    const r = await fetch('/api/queue/' + id + '/cancel', { method: 'PUT' });
    if (r.ok) {
      loadQueue();
    } else {
      const e = await r.json();
      alert('Failed to cancel: ' + (e.error || 'Unknown error'));
    }
  } catch(e) {
    alert('Network error: ' + e.message);
  }
}

// --- Init ---
loadConfig();
loadStatus();
loadQueue();
setInterval(loadStatus, 3000);
setInterval(function() {
  if (document.getElementById('tab-queue').classList.contains('active')) {
    loadQueue();
  }
}, 10000);
</script>
</body>
</html>"#;

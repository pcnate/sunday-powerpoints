use axum::{
    extract::{ Path, Query, State as AxumState },
    extract::ws::{ WebSocket, WebSocketUpgrade, Message },
    http::StatusCode,
    response::{ IntoResponse, Json },
    routing::{ get, post, put },
    Router,
};
use serde::{ Deserialize, Serialize };
use std::collections::HashMap;
use std::sync::Arc;
use tokio::sync::{ broadcast, mpsc };
use tower_http::cors::CorsLayer;

use crate::broadcast::WorkerEvent;
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
    paused_types: Vec<String>,
    active_types: Vec<String>,
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
        .route( "/api/scheduler/pause-type/{type_name}", post( pause_type_handler ) )
        .route( "/api/scheduler/run-job/{job_id}", post( run_job_handler ) )
        .route( "/ws", get( ws_handler ) )
        .fallback( fallback_handler )
        .layer( CorsLayer::permissive() )
        .with_state( shared );

    let addr = format!( "0.0.0.0:{}", port );
    tracing::info!( "Web UI listening on http://localhost:{}", port );

    let listener = tokio::net::TcpListener::bind( &addr ).await
        .expect( "Failed to bind web UI port" );

    axum::serve( listener, app ).await
        .expect( "Web UI server error" );
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
        paused_types: snapshot.paused_types,
        active_types: snapshot.active_types,
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


/// POST /api/scheduler/pause-type/:type_name — Toggle pause for a job type.
async fn pause_type_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Path( type_name ): Path<String>,
) -> impl IntoResponse {
    let _ = state.tray_tx.try_send( TrayCommand::TogglePauseJobType( type_name ) );
    Json( serde_json::json!({ "success": true }) )
}


/// POST /api/scheduler/run-job/:job_id — Claim and execute a specific job.
async fn run_job_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    Path( job_id ): Path<u32>,
) -> impl IntoResponse {
    let _ = state.tray_tx.try_send( TrayCommand::RunJob( job_id ) );
    Json( serde_json::json!({ "success": true }) )
}


/// GET /ws — WebSocket upgrade for real-time push events.
async fn ws_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    ws: WebSocketUpgrade,
) -> impl IntoResponse {
    ws.on_upgrade( move |socket| handle_ws( socket, state ) )
}


/// Handle a WebSocket connection: send snapshot, then stream events.
async fn handle_ws( mut socket: WebSocket, state: Arc<WebState> ) {
    // Send initial status snapshot
    let snapshot = state.app_state.status_snapshot().await;
    let event = WorkerEvent::StatusSnapshot( snapshot );
    if let Ok( json ) = serde_json::to_string( &event )
        && socket.send( Message::Text( json.into() ) ).await.is_err()
    {
        return;
    }

    // Subscribe to broadcast channel
    let mut rx = state.app_state.event_tx.subscribe();

    loop {
        tokio::select! {
            result = rx.recv() => {
                match result {
                    Ok( event ) => {
                        if let Ok( json ) = serde_json::to_string( &event )
                            && socket.send( Message::Text( json.into() ) ).await.is_err()
                        {
                            break;
                        }
                    }
                    Err( broadcast::error::RecvError::Lagged( n ) ) => {
                        tracing::warn!( "WebSocket client lagged by {} events, sending snapshot", n );
                        let snapshot = state.app_state.status_snapshot().await;
                        let event = WorkerEvent::StatusSnapshot( snapshot );
                        if let Ok( json ) = serde_json::to_string( &event )
                            && socket.send( Message::Text( json.into() ) ).await.is_err()
                        {
                            break;
                        }
                    }
                    Err( broadcast::error::RecvError::Closed ) => break,
                }
            }
            msg = socket.recv() => {
                match msg {
                    Some( Ok( _ ) ) => {} // Ignore client messages
                    _ => break, // Client disconnected
                }
            }
        }
    }
}


// ─── Static asset serving (release) / dev proxy (debug) ───────────────────────

/// Embedded Angular build assets for release builds.
///
/// Requires `ng build worker-ui` before `cargo build --release`.
#[cfg( not( debug_assertions ) )]
#[derive( rust_embed::Embed )]
#[folder = "../dist/worker-ui/browser/"]
struct Assets;


/// Fallback handler (release): serves embedded Angular build assets.
///
/// Serves files by path with proper MIME types. Unknown paths get index.html
/// for Angular's client-side routing (SPA fallback).
#[cfg( not( debug_assertions ) )]
async fn fallback_handler(
    uri: axum::http::Uri,
) -> axum::response::Response {
    let path = uri.path().trim_start_matches( '/' );
    let path = if path.is_empty() { "index.html" } else { path };

    serve_embedded( path )
        .unwrap_or_else( || {
            serve_embedded( "index.html" )
                .unwrap_or_else( || {
                    axum::response::Response::builder()
                        .status( StatusCode::NOT_FOUND )
                        .body( axum::body::Body::from( "Not found" ) )
                        .unwrap()
                } )
        } )
}


/// Resolve an embedded file to an HTTP response with correct content type.
#[cfg( not( debug_assertions ) )]
fn serve_embedded( path: &str ) -> Option<axum::response::Response> {
    let file = Assets::get( path )?;
    let mime = mime_guess::from_path( path ).first_or_octet_stream();

    Some(
        axum::response::Response::builder()
            .status( StatusCode::OK )
            .header( axum::http::header::CONTENT_TYPE, mime.as_ref() )
            .body( axum::body::Body::from( file.data.into_owned() ) )
            .unwrap()
    )
}


/// Fallback handler (dev): reverse-proxies to Angular dev server on port 4201.
///
/// All API routes are handled by explicit route handlers above, so this
/// fallback only forwards asset/page requests to the Angular dev server.
#[cfg( debug_assertions )]
async fn fallback_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
    req: axum::extract::Request,
) -> axum::response::Response {
    let path = req.uri().path_and_query()
        .map( |pq| pq.as_str().to_string() )
        .unwrap_or_else( || "/".to_string() );
    let url = format!( "http://localhost:4201{}", path );

    match state.http_client.get( &url ).send().await {
        Ok( resp ) => {
            let status = axum::http::StatusCode::from_u16( resp.status().as_u16() )
                .unwrap_or( axum::http::StatusCode::BAD_GATEWAY );
            let mut builder = axum::response::Response::builder().status( status );

            for ( key, value ) in resp.headers() {
                if key != "transfer-encoding" {
                    builder = builder.header( key.as_str(), value.as_bytes() );
                }
            }

            let body = resp.bytes().await.unwrap_or_default();
            builder.body( axum::body::Body::from( body ) )
                .unwrap_or_else( |_| {
                    axum::response::Response::builder()
                        .status( StatusCode::BAD_GATEWAY )
                        .body( axum::body::Body::from( "Proxy error" ) )
                        .unwrap()
                } )
        }
        Err( e ) => {
            tracing::warn!( "Dev proxy error: {}", e );
            axum::response::Response::builder()
                .status( StatusCode::BAD_GATEWAY )
                .body( axum::body::Body::from( format!(
                    "Angular dev server unreachable at localhost:4201: {}", e
                ) ) )
                .unwrap()
        }
    }
}

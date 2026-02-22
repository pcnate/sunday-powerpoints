use axum::{
    extract::{ Path, Query, State as AxumState },
    http::StatusCode,
    response::{ Html, IntoResponse, Json },
    routing::{ get, post, put },
    Router,
};
use serde::{ Deserialize, Serialize };
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
    shift_start: String,
    shift_end: String,
}


/// Config update request for POST /api/config.
#[derive( Deserialize )]
struct ConfigUpdate {
    server_url: Option<String>,
    worker_id: Option<String>,
    types: Option<Vec<String>>,
    shift_start: Option<String>,
    shift_end: Option<String>,
    poll_interval_secs: Option<u64>,
    heartbeat_interval_secs: Option<u64>,
    job_delay_secs: Option<u64>,
    web_ui_port: Option<u16>,
    whisper_path: Option<String>,
    whisper_model: Option<String>,
    claude_path: Option<String>,
    ffmpeg_path: Option<String>,
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
        .route( "/api/queue/{id}/cancel", put( cancel_job_handler ) )
        .route( "/api/scheduler/pause", post( pause_handler ) )
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
        shift_start: config.schedule.shift_start.clone(),
        shift_end: config.schedule.shift_end.clone(),
    } )
}


/// GET /api/config — Return current configuration as JSON.
async fn get_config_handler(
    AxumState( state ): AxumState<Arc<WebState>>,
) -> impl IntoResponse {
    let config = state.app_state.config.read().await;
    Json( config.clone() )
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
    if let Some( v ) = update.claude_path {
        config.tools.claude_path = v;
    }
    if let Some( v ) = update.ffmpeg_path {
        config.tools.ffmpeg_path = v;
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
      <option value="video-alignment">Video Alignment</option>
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
    <label>Job Types</label>
    <div class="checkbox-group">
      <label><input type="checkbox" id="type_alignment" value="video-alignment"> Video Alignment</label>
      <label><input type="checkbox" id="type_transcription" value="transcription"> Transcription</label>
      <label><input type="checkbox" id="type_claude" value="claude-processing"> Claude Processing</label>
    </div>
  </div>

  <h2>Schedule</h2>
  <div class="row">
    <div class="form-group">
      <label for="shift_start">Shift Start</label>
      <input id="shift_start" type="text" placeholder="HH:MM">
    </div>
    <div class="form-group">
      <label for="shift_end">Shift End</label>
      <input id="shift_end" type="text" placeholder="HH:MM">
    </div>
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
  } catch(e) { console.error('Status fetch failed', e); }
}

async function togglePause() {
  await fetch('/api/scheduler/pause', { method: 'POST' });
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
    document.getElementById('type_alignment').checked = types.includes('video-alignment');
    document.getElementById('type_transcription').checked = types.includes('transcription');
    document.getElementById('type_claude').checked = types.includes('claude-processing');

    document.getElementById('shift_start').value = c.schedule.shift_start;
    document.getElementById('shift_end').value = c.schedule.shift_end;
    document.getElementById('poll_interval').value = c.timing.poll_interval_secs;
    document.getElementById('heartbeat_interval').value = c.timing.heartbeat_interval_secs;
    document.getElementById('job_delay').value = c.timing.job_delay_secs;
    document.getElementById('whisper_path').value = c.tools.whisper_path;
    document.getElementById('whisper_model').value = c.tools.whisper_model;
    document.getElementById('claude_path').value = c.tools.claude_path;
    document.getElementById('ffmpeg_path').value = c.tools.ffmpeg_path;
  } catch(e) { console.error('Config fetch failed', e); }
}

function getSelectedTypes() {
  const types = [];
  if (document.getElementById('type_alignment').checked) types.push('video-alignment');
  if (document.getElementById('type_transcription').checked) types.push('transcription');
  if (document.getElementById('type_claude').checked) types.push('claude-processing');
  return types;
}

async function saveConfig() {
  const msg = document.getElementById('msg');
  try {
    const body = {
      server_url: document.getElementById('server_url').value,
      worker_id: document.getElementById('worker_id').value,
      types: getSelectedTypes(),
      shift_start: document.getElementById('shift_start').value,
      shift_end: document.getElementById('shift_end').value,
      poll_interval_secs: parseInt(document.getElementById('poll_interval').value),
      heartbeat_interval_secs: parseInt(document.getElementById('heartbeat_interval').value),
      job_delay_secs: parseInt(document.getElementById('job_delay').value),
      whisper_path: document.getElementById('whisper_path').value,
      whisper_model: document.getElementById('whisper_model').value,
      claude_path: document.getElementById('claude_path').value,
      ffmpeg_path: document.getElementById('ffmpeg_path').value,
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

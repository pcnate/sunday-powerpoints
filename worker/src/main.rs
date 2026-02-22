mod api_client;
mod config;
mod heartbeat;
mod job_runner;
mod runners;
mod scheduler;
mod state;
mod tray;
mod web_ui;

use anyhow::Result;
use std::sync::Arc;
use tokio::sync::mpsc;
use tokio_util::sync::CancellationToken;
use tracing_subscriber::EnvFilter;

use api_client::ApiClient;
use state::{ AppState, TrayCommand };


/// Application entry point.
///
/// Two-thread architecture:
/// - **Main thread**: winit event loop + system tray (win32 message pump)
/// - **Tokio thread**: scheduler, HTTP polling, heartbeat, job execution
///
/// Communication flows through `mpsc` (tray → scheduler) and shared
/// `AppState` with `RwLock` (scheduler → tray).
fn main() -> Result<()> {
    // Initialize tracing with RUST_LOG env filter
    tracing_subscriber::fmt()
        .with_env_filter(
            EnvFilter::try_from_default_env()
                .unwrap_or_else( |_| EnvFilter::new( "info" ) ),
        )
        .init();

    tracing::info!( "sunday-worker starting up" );

    // Load or create configuration
    let config = config::load_config()?;
    let config_path = config::config_path()?;

    tracing::info!( "Config loaded from: {}", config_path.display() );
    tracing::info!( "Server URL: {}", config.server.url );
    tracing::info!( "Worker ID: {}", config.worker.id );
    tracing::info!( "Job types: {:?}", config.worker.types );
    tracing::info!(
        "Shift window: {} - {}",
        config.schedule.shift_start,
        config.schedule.shift_end
    );
    tracing::info!( "Web UI port: {}", config.web_ui.port );

    // Create shared state
    let state = Arc::new( AppState::new( config.clone() ) );

    // Create API client
    let api = Arc::new( ApiClient::new( &config.server.url ) );

    // Shutdown token for graceful termination
    let shutdown = CancellationToken::new();

    // Channel for tray commands → scheduler
    let ( tray_tx, tray_rx ) = mpsc::channel::<TrayCommand>( 32 );

    // Clone shutdown for Ctrl+C handler
    let shutdown_ctrlc = shutdown.clone();

    // Clone tray_tx for the web UI (it also sends TrayCommand::ReloadConfig)
    let web_tray_tx = tray_tx.clone();
    let web_ui_port = config.web_ui.port;

    // Spawn the tokio runtime on a dedicated thread
    let rt_state = Arc::clone( &state );
    let rt_api = Arc::clone( &api );
    let rt_shutdown = shutdown.clone();
    let web_state = Arc::clone( &state );
    std::thread::spawn( move || {
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .build()
            .expect( "Failed to build tokio runtime" );

        rt.block_on( async move {
            // Install Ctrl+C handler
            let ctrlc_shutdown = shutdown_ctrlc.clone();
            tokio::spawn( async move {
                tokio::signal::ctrl_c().await.ok();
                tracing::info!( "Ctrl+C received, shutting down..." );
                ctrlc_shutdown.cancel();
            } );

            // Start the config web UI server (runs in background)
            tokio::spawn( web_ui::run( web_state, web_tray_tx, web_ui_port ) );

            // Check initial connectivity
            let connected = rt_api.health_check().await;
            if connected {
                tracing::info!( "Connected to server" );
                *rt_state.connected.write().await = true;
            } else {
                tracing::warn!( "Could not reach server (will retry during polling)" );
            }

            // Run the scheduler
            scheduler::run(
                rt_state,
                rt_api,
                rt_shutdown,
                tray_rx,
            ).await;
        } );
    } );

    // Main thread: run the system tray event loop
    // This blocks until the user clicks Exit or shutdown is signalled
    tracing::info!( "Starting system tray" );
    tray::run( tray_tx, Arc::clone( &state ), shutdown.clone() )?;

    tracing::info!( "sunday-worker shut down cleanly" );
    Ok(())
}

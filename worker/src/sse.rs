use anyhow::{ Context, Result };
use std::sync::Arc;
use std::time::{ Duration, Instant };
use tokio_util::sync::CancellationToken;

use crate::state::AppState;


/// Run the SSE client, reconnecting on failure.
///
/// Maintains a persistent Server-Sent Events connection to the Express
/// server. The connection itself tracks the worker as online, even when
/// off-shift. Job creation events are logged for potential future use.
pub async fn run(
    state: Arc<AppState>,
    shutdown: CancellationToken,
) {
    let mut retry_delay = Duration::from_secs( 1 );

    loop {
        if shutdown.is_cancelled() { break; }

        let config = state.config.read().await.clone();
        let url = format!(
            "{}/api/jobs/workers/events?worker_id={}&types={}",
            config.server.url.trim_end_matches( '/' ),
            &config.worker.id,
            config.worker.types.join( "," )
        );

        let start = Instant::now();

        match connect( &url, &state, &shutdown ).await {
            Ok(()) => break,
            Err( e ) => {
                *state.connected.write().await = false;

                // Reset backoff if connection lasted > 30s (was working, then dropped)
                if start.elapsed() > Duration::from_secs( 30 ) {
                    retry_delay = Duration::from_secs( 1 );
                }

                tracing::warn!( "SSE: {}, reconnecting in {:?}", e, retry_delay );

                tokio::select! {
                    _ = tokio::time::sleep( retry_delay ) => {}
                    _ = shutdown.cancelled() => { break; }
                }

                retry_delay = ( retry_delay * 2 ).min( Duration::from_secs( 30 ) );
            }
        }
    }

    tracing::info!( "SSE client stopped" );
}


/// Establish an SSE connection and stream events.
///
/// Returns Ok(()) on clean shutdown, Err on connection failure or stream error.
async fn connect(
    url: &str,
    state: &AppState,
    shutdown: &CancellationToken,
) -> Result<()> {
    let client = reqwest::Client::builder()
        .connect_timeout( Duration::from_secs( 10 ) )
        .build()
        .context( "Failed to build SSE client" )?;

    let mut response = client.get( url )
        .header( "Accept", "text/event-stream" )
        .send()
        .await
        .context( "SSE connection failed" )?;

    if !response.status().is_success() {
        anyhow::bail!( "SSE returned status {}", response.status() );
    }

    *state.connected.write().await = true;
    tracing::info!( "SSE connected" );

    let mut buffer = String::new();

    loop {
        // 60s read timeout — server sends keepalive every 30s
        let chunk = tokio::select! {
            result = tokio::time::timeout( Duration::from_secs( 60 ), response.chunk() ) => {
                match result {
                    Ok( r ) => r,
                    Err( _ ) => return Err( anyhow::anyhow!( "SSE read timeout (no data in 60s)" ) ),
                }
            }
            _ = shutdown.cancelled() => return Ok(()),
        };

        match chunk {
            Ok( Some( bytes ) ) => {
                buffer.push_str( &String::from_utf8_lossy( &bytes ) );

                // Process complete lines
                while let Some( pos ) = buffer.find( '\n' ) {
                    let line = buffer[ ..pos ].trim_end_matches( '\r' ).to_string();
                    buffer.drain( ..=pos );
                    process_line( &line );
                }
            }
            Ok( None ) => {
                return Err( anyhow::anyhow!( "SSE stream ended" ) );
            }
            Err( e ) => {
                return Err( anyhow::anyhow!( "SSE stream error: {}", e ) );
            }
        }
    }
}


/// Process a single SSE line.
fn process_line( line: &str ) {
    if line.is_empty() || line.starts_with( ':' ) {
        return;
    }

    if let Some( event ) = line.strip_prefix( "event: " ) {
        tracing::debug!( "SSE event: {}", event );
    }

    if let Some( _data ) = line.strip_prefix( "data: " ) {
        tracing::trace!( "SSE data received" );
    }
}

use anyhow::Result;
use muda::{ Menu, MenuEvent, MenuItem, PredefinedMenuItem };
use std::sync::Arc;
use tokio::sync::mpsc;
use tray_icon::{ TrayIcon, TrayIconBuilder, Icon };
use winit::application::ApplicationHandler;
use winit::event::WindowEvent;
use winit::event_loop::{ ActiveEventLoop, EventLoop };
use winit::window::WindowId;

use crate::state::{ SchedulerPhase, TrayCommand, TrayState };


/// Tray icon color corresponding to scheduler state.
#[derive( Debug, Clone, Copy, PartialEq, Eq )]
enum TrayColor {
    /// Green: in shift, idle or cooldown.
    Green,
    /// Yellow: executing a job.
    Yellow,
    /// Gray: off shift or paused.
    Gray,
    /// Red: disconnected from server.
    Red,
}


/// Map scheduler phase + connection status to a tray color.
fn phase_to_color( state: &TrayState ) -> TrayColor {
    if !state.connected {
        return TrayColor::Red;
    }
    match state.phase {
        SchedulerPhase::Executing => TrayColor::Yellow,
        SchedulerPhase::Idle | SchedulerPhase::Polling | SchedulerPhase::Cooldown => TrayColor::Green,
        SchedulerPhase::OffShift | SchedulerPhase::Paused => TrayColor::Gray,
    }
}


/// Generate a 32x32 RGBA icon with a filled circle of the given color.
fn generate_icon( color: TrayColor ) -> Icon {
    let size = 32u32;
    let ( r, g, b ) = match color {
        TrayColor::Green => ( 76u8, 175, 80 ),    // Material Green 500
        TrayColor::Yellow => ( 255, 193, 7 ),      // Material Amber 500
        TrayColor::Gray => ( 158, 158, 158 ),      // Material Grey 500
        TrayColor::Red => ( 244, 67, 54 ),         // Material Red 500
    };

    let mut rgba = vec![ 0u8; ( size * size * 4 ) as usize ];
    let center = size as f32 / 2.0;
    let radius = center - 2.0;

    for y in 0..size {
        for x in 0..size {
            let dx = x as f32 - center;
            let dy = y as f32 - center;
            let dist = ( dx * dx + dy * dy ).sqrt();

            let idx = ( ( y * size + x ) * 4 ) as usize;
            if dist <= radius {
                // Inside the circle
                rgba[ idx ] = r;
                rgba[ idx + 1 ] = g;
                rgba[ idx + 2 ] = b;
                rgba[ idx + 3 ] = 255;
            } else if dist <= radius + 1.0 {
                // Anti-aliased edge
                let alpha = ( ( radius + 1.0 - dist ) * 255.0 ) as u8;
                rgba[ idx ] = r;
                rgba[ idx + 1 ] = g;
                rgba[ idx + 2 ] = b;
                rgba[ idx + 3 ] = alpha;
            }
            // else: transparent (already 0)
        }
    }

    Icon::from_rgba( rgba, size, size ).expect( "Failed to create tray icon" )
}


/// Build the status text for the tray tooltip.
fn build_tooltip( state: &TrayState ) -> String {
    let mut parts = vec![ format!( "Sunday Worker — {}", state.phase ) ];

    if let Some( ref job_type ) = state.current_job_type {
        parts.push( format!( "Job: #{} ({})", state.current_job_id.unwrap_or( 0 ), job_type ) );
    }

    parts.push( format!(
        "Completed: {} | Failed: {}",
        state.jobs_completed, state.jobs_failed
    ) );

    parts.join( "\n" )
}


/// Menu item IDs.
struct MenuIds {
    status: MenuItem,
    pause_resume: MenuItem,
    open_config: MenuItem,
    quit: MenuItem,
}


/// Run the system tray on the main thread.
///
/// This blocks the main thread with the winit event loop.
/// Communicates with the async scheduler via channels.
///
/// @param tray_tx - sender for commands to the scheduler
/// @param state - shared application state for reading current status
/// @param shutdown - cancellation token signalled on quit
pub fn run(
    tray_tx: mpsc::Sender<TrayCommand>,
    state: Arc<crate::state::AppState>,
    shutdown: tokio_util::sync::CancellationToken,
) -> Result<()> {
    let event_loop = EventLoop::new()?;

    let mut app = TrayApp::new( tray_tx, state, shutdown )?;

    event_loop.run_app( &mut app )?;

    Ok(())
}


/// The winit application handler for the system tray.
struct TrayApp {
    tray_tx: mpsc::Sender<TrayCommand>,
    state: Arc<crate::state::AppState>,
    shutdown: tokio_util::sync::CancellationToken,
    tray_icon: Option<TrayIcon>,
    menu_ids: Option<MenuIds>,
    current_color: TrayColor,
    is_paused: bool,
}


impl TrayApp {
    fn new(
        tray_tx: mpsc::Sender<TrayCommand>,
        state: Arc<crate::state::AppState>,
        shutdown: tokio_util::sync::CancellationToken,
    ) -> Result<Self> {
        Ok( Self {
            tray_tx,
            state,
            shutdown,
            tray_icon: None,
            menu_ids: None,
            current_color: TrayColor::Gray,
            is_paused: false,
        } )
    }


    /// Create the tray icon and menu.
    fn init_tray( &mut self ) -> Result<()> {
        let menu = Menu::new();

        let status = MenuItem::new( "Sunday Worker — Starting...", false, None );
        let separator = PredefinedMenuItem::separator();
        let pause_resume = MenuItem::new( "Pause", true, None );
        let open_config = MenuItem::new( "Open Config", true, None );
        let separator2 = PredefinedMenuItem::separator();
        let quit = MenuItem::new( "Exit", true, None );

        menu.append( &status )?;
        menu.append( &separator )?;
        menu.append( &pause_resume )?;
        menu.append( &open_config )?;
        menu.append( &separator2 )?;
        menu.append( &quit )?;

        let icon = generate_icon( TrayColor::Gray );

        let tray = TrayIconBuilder::new()
            .with_menu( Box::new( menu ) )
            .with_tooltip( "Sunday Worker — Starting..." )
            .with_icon( icon )
            .build()?;

        self.tray_icon = Some( tray );
        self.menu_ids = Some( MenuIds {
            status,
            pause_resume,
            open_config,
            quit,
        } );

        Ok(())
    }


    /// Update the tray icon and menu text based on current state.
    fn update_tray( &mut self, tray_state: &TrayState ) {
        let new_color = phase_to_color( tray_state );

        // Update icon if color changed
        if new_color != self.current_color
            && let Some( ref tray ) = self.tray_icon
        {
            let icon = generate_icon( new_color );
            let _ = tray.set_icon( Some( icon ) );
            self.current_color = new_color;
        }

        // Update tooltip
        let tooltip = build_tooltip( tray_state );
        if let Some( ref tray ) = self.tray_icon {
            let _ = tray.set_tooltip( Some( &tooltip ) );
        }

        // Update menu status text
        if let Some( ref ids ) = self.menu_ids {
            let status_text = format!( "Status: {}", tray_state.phase );
            ids.status.set_text( &status_text );

            // Update pause/resume text
            let pause_text = if tray_state.phase == SchedulerPhase::Paused {
                "Resume"
            } else {
                "Pause"
            };
            ids.pause_resume.set_text( pause_text );
        }
    }
}


impl ApplicationHandler for TrayApp {
    fn resumed( &mut self, _event_loop: &ActiveEventLoop ) {
        // Initialize tray on first resume
        if self.tray_icon.is_none()
            && let Err( e ) = self.init_tray()
        {
            tracing::error!( "Failed to initialize tray: {}", e );
        }
    }


    fn window_event(
        &mut self,
        _event_loop: &ActiveEventLoop,
        _window_id: WindowId,
        _event: WindowEvent,
    ) {
        // No windows — tray only
    }


    fn about_to_wait( &mut self, event_loop: &ActiveEventLoop ) {
        // Check for shutdown
        if self.shutdown.is_cancelled() {
            event_loop.exit();
            return;
        }

        // Process menu events
        while let Ok( event ) = MenuEvent::receiver().try_recv() {
            if let Some( ref ids ) = self.menu_ids {
                if event.id() == ids.quit.id() {
                    tracing::info!( "Exit clicked" );
                    let _ = self.tray_tx.try_send( TrayCommand::Quit );
                    self.shutdown.cancel();
                    event_loop.exit();
                    return;
                } else if event.id() == ids.pause_resume.id() {
                    self.is_paused = !self.is_paused;
                    let _ = self.tray_tx.try_send( TrayCommand::TogglePause );
                } else if event.id() == ids.open_config.id() {
                    let _ = self.tray_tx.try_send( TrayCommand::OpenConfig );
                }
            }
        }

        // Poll the state (non-blocking) to update the tray
        // We create a small tokio runtime just for the state read
        // (this runs on the main thread, so we can't use async directly)
        let state = self.state.clone();
        let tray_state = std::thread::scope( |s| {
            s.spawn( || {
                let rt = tokio::runtime::Builder::new_current_thread()
                    .build()
                    .expect( "mini runtime" );
                rt.block_on( state.snapshot() )
            } ).join().expect( "snapshot thread" )
        } );

        self.update_tray( &tray_state );
    }
}

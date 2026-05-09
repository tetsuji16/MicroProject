mod app;
mod dependency;
mod mspdi;
mod table_view;
mod timeline_view;

use std::path::PathBuf;

fn main() -> eframe::Result<()> {
    let initial_file = std::env::args_os().nth(1).map(PathBuf::from);
    let native_options = eframe::NativeOptions {
        viewport: eframe::egui::ViewportBuilder::default()
            .with_decorations(true)
            .with_inner_size([1366.0, 768.0])
            .with_min_inner_size([1180.0, 720.0])
            .with_title("MicroProject"),
        ..Default::default()
    };

    eframe::run_native(
        "MicroProject",
        native_options,
        Box::new(move |cc| Ok(Box::new(app::GanttApp::new(cc, initial_file.clone())))),
    )
}

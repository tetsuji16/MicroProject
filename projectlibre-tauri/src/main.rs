mod commands;
mod mspdi;
mod session;

use session::{ProjectAppState, ProjectSession};
use std::path::PathBuf;
use std::sync::Mutex;

fn main() {
    let initial_file = std::env::args_os().nth(1).map(PathBuf::from);
    let state = ProjectAppState {
        session: Mutex::new(ProjectSession::new(initial_file)),
    };

    tauri::Builder::default()
        .manage(state)
        .invoke_handler(tauri::generate_handler![
            commands::project_snapshot,
            commands::project_open,
            commands::project_save,
            commands::project_save_as,
            commands::project_upsert_task,
            commands::project_delete_task,
            commands::project_create_task,
            commands::project_upsert_dependency,
            commands::project_delete_dependency,
        ])
        .run(tauri::generate_context!())
        .expect("failed to start MicroProject");
}

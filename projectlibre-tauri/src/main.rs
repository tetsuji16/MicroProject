mod engine;
mod commands;
mod interop;
mod models;
mod storage;

use commands::*;
use storage::{AppState, AppStore};

fn main() {
    let state = AppState::new(AppStore::load_or_default());

    tauri::Builder::default()
        .manage(state)
        .invoke_handler(tauri::generate_handler![
            workspace_snapshot,
            workspace_export_json,
            workspace_export_xml,
            workspace_import_json,
            workspace_import_xml,
            workspace_recalculate,
            workspace_upsert_project,
            workspace_delete_project,
            workspace_upsert_task,
            workspace_delete_task,
            workspace_upsert_dependency,
            workspace_delete_dependency,
            workspace_upsert_resource,
            workspace_delete_resource,
            workspace_upsert_assignment,
            workspace_delete_assignment,
            workspace_upsert_calendar,
            workspace_delete_calendar,
            workspace_capture_baseline
        ])
        .run(tauri::generate_context!("src-tauri/tauri.conf.json"))
        .expect("error while running tauri application");
}

// Lightweight TAURI command stubs for MVP

#[allow(unused_imports)]
use super::models::{Workspace, Project};

#[tauri::command]
fn workspace_snapshot() -> String { "snapshot".to_string() }

#[tauri::command]
fn workspace_export_json() -> String { "export_json".to_string() }

#[tauri::command]
fn workspace_export_xml() -> String { "export_xml".to_string() }

#[tauri::command]
fn workspace_import_json(_data: String) -> String { "import_json".to_string() }

#[tauri::command]
fn workspace_import_xml(_data: String) -> String { "import_xml".to_string() }

#[tauri::command]
fn workspace_recalculate() -> String { "recalculate".to_string() }

#[tauri::command]
fn workspace_upsert_project(name: String) -> String { format!("upsert_project:{}", name) }

#[tauri::command]
fn workspace_delete_project(id: String) -> String { id }

#[tauri::command]
fn workspace_upsert_task(_project_id: String, _task_name: String) -> String { "upsert_task".to_string() }

#[tauri::command]
fn workspace_delete_task(_task_id: String) -> String { "delete_task".to_string() }

#[tauri::command]
fn workspace_upsert_dependency(_task_id: String, _dep: String) -> String { "upsert_dependency".to_string() }

#[tauri::command]
fn workspace_delete_dependency(_dep_id: String) -> String { "delete_dependency".to_string() }

#[tauri::command]
fn workspace_upsert_resource(_name: String) -> String { "upsert_resource".to_string() }

#[tauri::command]
fn workspace_delete_resource(_res_id: String) -> String { "delete_resource".to_string() }

#[tauri::command]
fn workspace_upsert_assignment(_task_id: String, _res_id: String) -> String { "upsert_assignment".to_string() }

#[tauri::command]
fn workspace_delete_assignment(_assn_id: String) -> String { "delete_assignment".to_string() }

#[tauri::command]
fn workspace_upsert_calendar(_name: String) -> String { "upsert_calendar".to_string() }

#[tauri::command]
fn workspace_delete_calendar(_cal_id: String) -> String { "delete_calendar".to_string() }

#[tauri::command]
fn workspace_capture_baseline() -> String { "capture_baseline".to_string() }

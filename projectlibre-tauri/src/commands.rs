use crate::java_bridge::{JavaBridgeState, JavaBridgeStatus};
use projectlibre_tauri_backend::{
    AppState, AssignmentInput, CalendarInput, DependencyInput, ProjectInput, ResourceInput,
    TaskInput, WorkspaceSnapshot,
};
use tauri::State;

pub type CommandResult<T> = Result<T, String>;

#[tauri::command]
pub fn workspace_snapshot(state: State<'_, AppState>) -> CommandResult<WorkspaceSnapshot> {
    let store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_export_json(state: State<'_, AppState>) -> CommandResult<String> {
    let store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.export_json()
}

#[tauri::command]
pub fn workspace_export_xml(state: State<'_, AppState>) -> CommandResult<String> {
    let store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.export_xml()
}

#[tauri::command]
pub fn workspace_import_json(
    state: State<'_, AppState>,
    json: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.import_json(&json)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_import_xml(
    state: State<'_, AppState>,
    xml: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.import_xml(&xml)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_recalculate(state: State<'_, AppState>) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.recalculate_all_and_save()?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_project(state: State<'_, AppState>, input: ProjectInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_project(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_project(state: State<'_, AppState>, project_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_project(&project_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_task(state: State<'_, AppState>, input: TaskInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_task(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_task(state: State<'_, AppState>, task_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_task(&task_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_dependency(state: State<'_, AppState>, input: DependencyInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_dependency(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_dependency(state: State<'_, AppState>, dependency_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_dependency(&dependency_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_resource(state: State<'_, AppState>, input: ResourceInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_resource(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_resource(state: State<'_, AppState>, resource_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_resource(&resource_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_assignment(state: State<'_, AppState>, input: AssignmentInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_assignment(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_assignment(state: State<'_, AppState>, assignment_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_assignment(&assignment_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_calendar(state: State<'_, AppState>, input: CalendarInput) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_calendar(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_calendar(state: State<'_, AppState>, calendar_id: String) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_calendar(&calendar_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_capture_baseline(
    state: State<'_, AppState>,
    project_id: String,
    name: Option<String>,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state.store.lock().map_err(|_| "workspace store lock poisoned".to_string())?;
    store.capture_baseline(&project_id, name)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn java_bridge_status(state: State<'_, JavaBridgeState>) -> CommandResult<JavaBridgeStatus> {
    Ok(state.status())
}

#[tauri::command]
pub fn java_bridge_ping(state: State<'_, JavaBridgeState>) -> CommandResult<serde_json::Value> {
    state.ping()
}

#[tauri::command]
pub fn java_bridge_snapshot(state: State<'_, JavaBridgeState>) -> CommandResult<serde_json::Value> {
    state.snapshot()
}

#[tauri::command]
pub fn java_bridge_open_mpp(
    state: State<'_, JavaBridgeState>,
    path: String,
) -> CommandResult<serde_json::Value> {
    state.open_mpp(&path)
}

#[tauri::command]
pub fn java_bridge_import_mpp(
    state: State<'_, JavaBridgeState>,
    path: String,
) -> CommandResult<serde_json::Value> {
    state.import_mpp(&path)
}

#[tauri::command]
pub fn java_bridge_export_mpp(
    state: State<'_, JavaBridgeState>,
    path: String,
) -> CommandResult<serde_json::Value> {
    state.export_mpp(&path)
}

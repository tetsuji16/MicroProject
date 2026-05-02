use crate::{models::*, storage::AppState};
use tauri::State;

pub type CommandResult<T> = Result<T, String>;

#[tauri::command]
pub fn workspace_snapshot(state: State<'_, AppState>) -> CommandResult<WorkspaceSnapshot> {
    let store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_project(
    state: State<'_, AppState>,
    input: ProjectInput,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_project(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_project(
    state: State<'_, AppState>,
    project_id: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_project(&project_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_task(
    state: State<'_, AppState>,
    input: TaskInput,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_task(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_task(
    state: State<'_, AppState>,
    task_id: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_task(&task_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_dependency(
    state: State<'_, AppState>,
    input: DependencyInput,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_dependency(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_dependency(
    state: State<'_, AppState>,
    dependency_id: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_dependency(&dependency_id)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_upsert_calendar(
    state: State<'_, AppState>,
    input: CalendarInput,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.upsert_calendar(input)?;
    Ok(store.snapshot())
}

#[tauri::command]
pub fn workspace_delete_calendar(
    state: State<'_, AppState>,
    calendar_id: String,
) -> CommandResult<WorkspaceSnapshot> {
    let mut store = state
        .store
        .lock()
        .map_err(|_| "workspace store lock poisoned".to_string())?;
    store.delete_calendar(&calendar_id)?;
    Ok(store.snapshot())
}

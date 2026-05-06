use crate::session::{
    DependencyMutationInput, ProjectAppState, ProjectSnapshot, TaskMutationInput,
};
use std::path::PathBuf;
use tauri::State;

pub type CommandResult<T> = Result<T, String>;

#[tauri::command]
pub fn project_snapshot(state: State<'_, ProjectAppState>) -> CommandResult<ProjectSnapshot> {
    let session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    Ok(session.snapshot())
}

#[tauri::command]
pub fn project_open(
    state: State<'_, ProjectAppState>,
    path: String,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.load_into_current(PathBuf::from(path))
}

#[tauri::command]
pub fn project_save(state: State<'_, ProjectAppState>) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.save()
}

#[tauri::command]
pub fn project_save_as(
    state: State<'_, ProjectAppState>,
    path: String,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.save_as(PathBuf::from(path))
}

#[tauri::command]
pub fn project_upsert_task(
    state: State<'_, ProjectAppState>,
    input: TaskMutationInput,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.upsert_task(input)
}

#[tauri::command]
pub fn project_delete_task(
    state: State<'_, ProjectAppState>,
    uid: u32,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.delete_task(uid)
}

#[tauri::command]
pub fn project_create_task(
    state: State<'_, ProjectAppState>,
    after_uid: Option<u32>,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.create_task(after_uid)
}

#[tauri::command]
pub fn project_upsert_dependency(
    state: State<'_, ProjectAppState>,
    input: DependencyMutationInput,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.upsert_dependency(input)
}

#[tauri::command]
pub fn project_delete_dependency(
    state: State<'_, ProjectAppState>,
    input: DependencyMutationInput,
) -> CommandResult<ProjectSnapshot> {
    let mut session = state
        .session
        .lock()
        .map_err(|_| "project session lock poisoned".to_string())?;
    session.delete_dependency(input)
}

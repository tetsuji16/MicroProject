use serde::{Serialize, Deserialize};
use std::fs;
use std::path::Path;
use uuid::Uuid;

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct Task {
    pub id: String,
    pub name: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub dependencies: Vec<String>,
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct Project {
    pub id: String,
    pub name: String,
    pub tasks: Vec<Task>,
}

#[derive(Serialize, Deserialize, Clone, Debug, Default)]
pub struct Workspace {
    pub projects: Vec<Project>,
}

const WORKSPACE_PATH: &str = "workspace.json";

pub fn new_workspace() -> Workspace {
    Workspace { projects: vec![] }
}

pub fn load_workspace<P: AsRef<Path>>(path: P) -> Result<Workspace, String> {
    let p = path.as_ref();
    if p.exists() {
        let data = fs::read_to_string(p).map_err(|e| e.to_string())?;
        let ws: Workspace = serde_json::from_str(&data).map_err(|e| e.to_string())?;
        Ok(ws)
    } else {
        Ok(new_workspace())
    }
}

pub fn save_workspace<P: AsRef<Path>>(path: P, ws: &Workspace) -> Result<(), String> {
    let data = serde_json::to_string_pretty(ws).map_err(|e| e.to_string())?;
    fs::write(path, data).map_err(|e| e.to_string())?;
    Ok(())
}

pub fn add_project(ws: &mut Workspace, name: &str) -> Project {
    let id = Uuid::new_v4().to_string();
    let proj = Project { id: id.clone(), name: name.to_string(), tasks: vec![] };
    ws.projects.push(proj.clone());
    proj
}

pub fn add_task(project: &mut Project, name: &str) -> Task {
    let id = Uuid::new_v4().to_string();
    let t = Task { id: id, name: name.to_string(), start: None, end: None, dependencies: vec![] };
    project.tasks.push(t.clone());
    t
}

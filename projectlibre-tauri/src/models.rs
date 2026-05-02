use serde::{Deserialize, Serialize};
use std::time::{SystemTime, UNIX_EPOCH};

pub const SCHEMA_VERSION: u32 = 1;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkspaceSnapshot {
    pub schema_version: u32,
    pub projects: Vec<ProjectRecord>,
    pub tasks: Vec<TaskRecord>,
    pub dependencies: Vec<DependencyRecord>,
    pub calendars: Vec<CalendarRecord>,
}

impl Default for WorkspaceSnapshot {
    fn default() -> Self {
        Self::new()
    }
}

impl WorkspaceSnapshot {
    pub fn new() -> Self {
        Self {
            schema_version: SCHEMA_VERSION,
            projects: Vec::new(),
            tasks: Vec::new(),
            dependencies: Vec::new(),
            calendars: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ProjectRecord {
    pub id: String,
    pub name: String,
    pub description: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ProjectInput {
    pub id: Option<String>,
    pub name: String,
    pub description: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TaskRecord {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub description: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub percent_complete: f32,
    pub parent_task_id: Option<String>,
    pub sort_order: i64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TaskInput {
    pub id: Option<String>,
    pub project_id: String,
    pub name: String,
    pub description: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub percent_complete: Option<f32>,
    pub parent_task_id: Option<String>,
    pub sort_order: Option<i64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DependencyRecord {
    pub id: String,
    pub project_id: String,
    pub predecessor_task_id: String,
    pub successor_task_id: String,
    pub relation: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DependencyInput {
    pub id: Option<String>,
    pub project_id: String,
    pub predecessor_task_id: String,
    pub successor_task_id: String,
    pub relation: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalendarRecord {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub timezone: Option<String>,
    pub working_days: Vec<u8>,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalendarInput {
    pub id: Option<String>,
    pub project_id: String,
    pub name: String,
    pub timezone: Option<String>,
    pub working_days: Option<Vec<u8>>,
}

pub fn generate_id(prefix: &str) -> String {
    format!("{prefix}_{}", unix_millis())
}

pub fn now_stamp() -> String {
    unix_millis().to_string()
}

fn unix_millis() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}


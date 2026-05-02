use serde::{Deserialize, Serialize};
use std::time::{SystemTime, UNIX_EPOCH};

pub const SCHEMA_VERSION: u32 = 2;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct WorkspaceSnapshot {
    pub schema_version: u32,
    pub settings: WorkspaceSettings,
    pub projects: Vec<ProjectRecord>,
    pub tasks: Vec<TaskRecord>,
    pub dependencies: Vec<DependencyRecord>,
    pub resources: Vec<ResourceRecord>,
    pub assignments: Vec<AssignmentRecord>,
    pub calendars: Vec<CalendarRecord>,
    pub baselines: Vec<BaselineRecord>,
}

impl WorkspaceSnapshot {
    pub fn new() -> Self {
        Self {
            schema_version: SCHEMA_VERSION,
            settings: WorkspaceSettings::default(),
            projects: Vec::new(),
            tasks: Vec::new(),
            dependencies: Vec::new(),
            resources: Vec::new(),
            assignments: Vec::new(),
            calendars: Vec::new(),
            baselines: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct WorkspaceSettings {
    pub default_calendar_id: Option<String>,
    pub units_per_day: f32,
    pub locale: Option<String>,
}

impl Default for WorkspaceSettings {
    fn default() -> Self {
        Self {
            default_calendar_id: None,
            units_per_day: 8.0,
            locale: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ProjectRecord {
    pub id: String,
    pub name: String,
    pub description: Option<String>,
    pub manager: Option<String>,
    pub status: String,
    pub priority: i32,
    pub calendar_id: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub calculated_start_date: Option<String>,
    pub calculated_finish_date: Option<String>,
    pub notes: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ProjectInput {
    pub id: Option<String>,
    pub name: String,
    pub description: Option<String>,
    pub manager: Option<String>,
    pub status: Option<String>,
    pub priority: Option<i32>,
    pub calendar_id: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub notes: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct TaskRecord {
    pub id: String,
    pub project_id: String,
    pub parent_task_id: Option<String>,
    pub name: String,
    pub description: Option<String>,
    pub outline_level: u32,
    pub wbs: String,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub calculated_start_date: Option<String>,
    pub calculated_finish_date: Option<String>,
    pub duration_hours: Option<f64>,
    pub work_hours: Option<f64>,
    pub percent_complete: f32,
    pub milestone: bool,
    pub constraint_type: String,
    pub calendar_id: Option<String>,
    pub notes: String,
    pub sort_order: i64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct TaskInput {
    pub id: Option<String>,
    pub project_id: String,
    pub parent_task_id: Option<String>,
    pub name: String,
    pub description: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub duration_hours: Option<f64>,
    pub work_hours: Option<f64>,
    pub percent_complete: Option<f32>,
    pub milestone: Option<bool>,
    pub constraint_type: Option<String>,
    pub calendar_id: Option<String>,
    pub notes: Option<String>,
    pub sort_order: Option<i64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct DependencyRecord {
    pub id: String,
    pub project_id: String,
    pub predecessor_task_id: String,
    pub successor_task_id: String,
    pub relation: String,
    pub lag_hours: f64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct DependencyInput {
    pub id: Option<String>,
    pub project_id: String,
    pub predecessor_task_id: String,
    pub successor_task_id: String,
    pub relation: Option<String>,
    pub lag_hours: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ResourceRecord {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub resource_type: String,
    pub max_units: f32,
    pub standard_rate: f64,
    pub overtime_rate: f64,
    pub cost_per_use: f64,
    pub calendar_id: Option<String>,
    pub notes: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ResourceInput {
    pub id: Option<String>,
    pub project_id: String,
    pub name: String,
    pub resource_type: Option<String>,
    pub max_units: Option<f32>,
    pub standard_rate: Option<f64>,
    pub overtime_rate: Option<f64>,
    pub cost_per_use: Option<f64>,
    pub calendar_id: Option<String>,
    pub notes: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct AssignmentRecord {
    pub id: String,
    pub project_id: String,
    pub task_id: String,
    pub resource_id: String,
    pub units: f32,
    pub work_hours: f64,
    pub actual_work_hours: f64,
    pub cost: f64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct AssignmentInput {
    pub id: Option<String>,
    pub project_id: String,
    pub task_id: String,
    pub resource_id: String,
    pub units: Option<f32>,
    pub work_hours: Option<f64>,
    pub actual_work_hours: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct CalendarRecord {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub timezone: Option<String>,
    pub working_days: Vec<u8>,
    pub hours_per_day: f32,
    pub working_hours: Vec<WorkInterval>,
    pub exceptions: Vec<CalendarException>,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct CalendarInput {
    pub id: Option<String>,
    pub project_id: String,
    pub name: String,
    pub timezone: Option<String>,
    pub working_days: Option<Vec<u8>>,
    pub hours_per_day: Option<f32>,
    pub working_hours: Option<Vec<WorkInterval>>,
    pub exceptions: Option<Vec<CalendarException>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct WorkInterval {
    pub start: String,
    pub finish: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct CalendarException {
    pub date: String,
    pub is_working: bool,
    pub intervals: Vec<WorkInterval>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct BaselineRecord {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub captured_at: String,
    pub project: ProjectSnapshot,
    pub tasks: Vec<TaskSnapshot>,
    pub resources: Vec<ResourceSnapshot>,
    pub assignments: Vec<AssignmentSnapshot>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ProjectSnapshot {
    pub id: String,
    pub name: String,
    pub calculated_start_date: Option<String>,
    pub calculated_finish_date: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct TaskSnapshot {
    pub id: String,
    pub name: String,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub percent_complete: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct ResourceSnapshot {
    pub id: String,
    pub name: String,
    pub max_units: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct AssignmentSnapshot {
    pub id: String,
    pub task_id: String,
    pub resource_id: String,
    pub units: f32,
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

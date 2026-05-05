use chrono::{Datelike, Duration as ChronoDuration, NaiveDate};
use roxmltree::{Document, Node};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeSet, HashMap, VecDeque};
use std::env;
use std::fs;
use std::path::PathBuf;
use std::sync::Mutex;
use uuid::Uuid;

pub const SCHEMA_VERSION: u32 = 3;

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
        Self { default_calendar_id: None, units_per_day: 8.0, locale: None }
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
    format!("{prefix}_{}", Uuid::new_v4())
}

pub fn now_stamp() -> String {
    chrono::Utc::now().timestamp_millis().to_string()
}

pub struct AppState {
    pub store: Mutex<AppStore>,
}

impl AppState {
    pub fn new(store: AppStore) -> Self {
        Self { store: Mutex::new(store) }
    }
}

pub struct AppStore {
    path: PathBuf,
    snapshot: WorkspaceSnapshot,
}

impl AppStore {
    pub fn load_or_default() -> Self {
        Self::load().unwrap_or_else(|_| Self::new(default_store_path()))
    }

    pub fn load() -> Result<Self, String> {
        Self::load_from(default_store_path())
    }

    pub fn load_from(path: PathBuf) -> Result<Self, String> {
        if !path.exists() {
            return Ok(Self::new(path));
        }

        let contents = fs::read_to_string(&path)
            .map_err(|error| format!("failed to read store file {path:?}: {error}"))?;
        let mut snapshot = serde_json::from_str::<WorkspaceSnapshot>(&contents)
            .map_err(|error| format!("failed to parse store file {path:?}: {error}"))?;
        if snapshot.schema_version < SCHEMA_VERSION {
            snapshot.schema_version = SCHEMA_VERSION;
        }

        let mut store = Self { path, snapshot };
        store.recalculate_all()?;
        Ok(store)
    }

    fn new(path: PathBuf) -> Self {
        Self { path, snapshot: WorkspaceSnapshot::new() }
    }

    pub fn snapshot(&self) -> WorkspaceSnapshot {
        self.snapshot.clone()
    }

    pub fn export_json(&self) -> Result<String, String> {
        serde_json::to_string_pretty(&self.snapshot).map_err(|error| format!("failed to serialize snapshot: {error}"))
    }

    pub fn export_xml(&self) -> Result<String, String> {
        export_workspace_xml(&self.snapshot)
    }

    pub fn import_json(&mut self, json: &str) -> Result<(), String> {
        self.snapshot = serde_json::from_str(json).map_err(|error| format!("failed to parse snapshot: {error}"))?;
        self.recalculate_all()?;
        self.save()
    }

    pub fn import_xml(&mut self, xml: &str) -> Result<(), String> {
        self.snapshot = import_workspace_xml(xml)?;
        self.recalculate_all()?;
        self.save()
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        rebuild_all(&mut self.snapshot)
    }

    pub fn recalculate_project(&mut self, project_id: &str) -> Result<(), String> {
        rebuild_project(&mut self.snapshot, project_id)
    }

    pub fn recalculate_all_and_save(&mut self) -> Result<(), String> {
        self.recalculate_all()?;
        self.save()
    }

    pub fn upsert_project(&mut self, input: ProjectInput) -> Result<ProjectRecord, String> {
        let name = input.name.trim();
        if name.is_empty() {
            return Err("project name is required".to_string());
        }
        let id = input.id.unwrap_or_else(|| generate_id("project"));
        let now = now_stamp();
        let created_at = self.snapshot.projects.iter().find(|project| project.id == id).map(|project| project.created_at.clone()).unwrap_or_else(|| now.clone());
        let project = ProjectRecord {
            id: id.clone(),
            name: name.to_string(),
            description: input.description,
            manager: input.manager,
            status: input.status.unwrap_or_else(|| "planning".to_string()),
            priority: input.priority.unwrap_or(500),
            calendar_id: input.calendar_id,
            start_date: input.start_date,
            finish_date: input.finish_date,
            calculated_start_date: None,
            calculated_finish_date: None,
            notes: input.notes.unwrap_or_default(),
            created_at,
            updated_at: now,
        };

        upsert_by_id(&mut self.snapshot.projects, project.clone());
        self.recalculate_project(&id)?;
        self.save()?;
        Ok(project)
    }

    pub fn delete_project(&mut self, project_id: &str) -> Result<(), String> {
        let before = self.snapshot.projects.len();
        self.snapshot.projects.retain(|project| project.id != project_id);
        if self.snapshot.projects.len() == before {
            return Err(format!("project not found: {project_id}"));
        }
        self.snapshot.tasks.retain(|task| task.project_id != project_id);
        self.snapshot.dependencies.retain(|dependency| dependency.project_id != project_id);
        self.snapshot.resources.retain(|resource| resource.project_id != project_id);
        self.snapshot.assignments.retain(|assignment| assignment.project_id != project_id);
        self.snapshot.calendars.retain(|calendar| calendar.project_id != project_id);
        self.snapshot.baselines.retain(|baseline| baseline.project_id != project_id);
        self.save()
    }

    pub fn upsert_task(&mut self, input: TaskInput) -> Result<TaskRecord, String> {
        if input.name.trim().is_empty() {
            return Err("task name is required".to_string());
        }
        self.require_project(&input.project_id)?;
        if let Some(parent) = &input.parent_task_id {
            self.require_task(parent, &input.project_id)?;
        }
        let id = input.id.unwrap_or_else(|| generate_id("task"));
        let now = now_stamp();
        let created_at = self.snapshot.tasks.iter().find(|task| task.id == id).map(|task| task.created_at.clone()).unwrap_or_else(|| now.clone());
        let task = TaskRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            parent_task_id: input.parent_task_id,
            name: input.name.trim().to_string(),
            description: input.description,
            outline_level: 0,
            wbs: String::new(),
            start_date: input.start_date,
            finish_date: input.finish_date,
            calculated_start_date: None,
            calculated_finish_date: None,
            duration_hours: input.duration_hours.or(Some(8.0)),
            work_hours: input.work_hours,
            percent_complete: input.percent_complete.unwrap_or(0.0),
            milestone: input.milestone.unwrap_or(false),
            constraint_type: input.constraint_type.unwrap_or_else(|| "ASAP".to_string()),
            calendar_id: input.calendar_id,
            notes: input.notes.unwrap_or_default(),
            sort_order: input.sort_order.unwrap_or(0),
            created_at,
            updated_at: now,
        };
        upsert_by_id(&mut self.snapshot.tasks, task.clone());
        self.recalculate_project(&input.project_id)?;
        self.save()?;
        Ok(task)
    }

    pub fn delete_task(&mut self, task_id: &str) -> Result<(), String> {
        let project_id = self.snapshot.tasks.iter().find(|task| task.id == task_id).map(|task| task.project_id.clone()).ok_or_else(|| format!("task not found: {task_id}"))?;
        self.snapshot.tasks.retain(|task| task.id != task_id);
        self.snapshot.dependencies.retain(|dependency| dependency.predecessor_task_id != task_id && dependency.successor_task_id != task_id);
        self.snapshot.assignments.retain(|assignment| assignment.task_id != task_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_dependency(&mut self, input: DependencyInput) -> Result<DependencyRecord, String> {
        if input.predecessor_task_id == input.successor_task_id {
            return Err("dependency cannot point to the same task".to_string());
        }
        self.require_project(&input.project_id)?;
        self.require_task(&input.predecessor_task_id, &input.project_id)?;
        self.require_task(&input.successor_task_id, &input.project_id)?;
        let id = input.id.unwrap_or_else(|| generate_id("dependency"));
        let now = now_stamp();
        let created_at = self.snapshot.dependencies.iter().find(|dependency| dependency.id == id).map(|dependency| dependency.created_at.clone()).unwrap_or_else(|| now.clone());
        let dependency = DependencyRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            predecessor_task_id: input.predecessor_task_id,
            successor_task_id: input.successor_task_id,
            relation: input.relation.unwrap_or_else(|| "FS".to_string()),
            lag_hours: input.lag_hours.unwrap_or(0.0),
            created_at,
            updated_at: now,
        };
        upsert_by_id(&mut self.snapshot.dependencies, dependency.clone());
        self.recalculate_project(&dependency.project_id)?;
        self.save()?;
        Ok(dependency)
    }

    pub fn delete_dependency(&mut self, dependency_id: &str) -> Result<(), String> {
        let project_id = self.snapshot.dependencies.iter().find(|dependency| dependency.id == dependency_id).map(|dependency| dependency.project_id.clone()).ok_or_else(|| format!("dependency not found: {dependency_id}"))?;
        self.snapshot.dependencies.retain(|dependency| dependency.id != dependency_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_resource(&mut self, input: ResourceInput) -> Result<ResourceRecord, String> {
        if input.name.trim().is_empty() {
            return Err("resource name is required".to_string());
        }
        self.require_project(&input.project_id)?;
        let id = input.id.unwrap_or_else(|| generate_id("resource"));
        let now = now_stamp();
        let created_at = self.snapshot.resources.iter().find(|resource| resource.id == id).map(|resource| resource.created_at.clone()).unwrap_or_else(|| now.clone());
        let resource = ResourceRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            name: input.name.trim().to_string(),
            resource_type: input.resource_type.unwrap_or_else(|| "work".to_string()),
            max_units: input.max_units.unwrap_or(100.0),
            standard_rate: input.standard_rate.unwrap_or(0.0),
            overtime_rate: input.overtime_rate.unwrap_or(0.0),
            cost_per_use: input.cost_per_use.unwrap_or(0.0),
            calendar_id: input.calendar_id,
            notes: input.notes.unwrap_or_default(),
            created_at,
            updated_at: now,
        };
        upsert_by_id(&mut self.snapshot.resources, resource.clone());
        self.recalculate_project(&resource.project_id)?;
        self.save()?;
        Ok(resource)
    }

    pub fn delete_resource(&mut self, resource_id: &str) -> Result<(), String> {
        let project_id = self.snapshot.resources.iter().find(|resource| resource.id == resource_id).map(|resource| resource.project_id.clone()).ok_or_else(|| format!("resource not found: {resource_id}"))?;
        self.snapshot.resources.retain(|resource| resource.id != resource_id);
        self.snapshot.assignments.retain(|assignment| assignment.resource_id != resource_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_assignment(&mut self, input: AssignmentInput) -> Result<AssignmentRecord, String> {
        self.require_project(&input.project_id)?;
        self.require_task(&input.task_id, &input.project_id)?;
        self.require_resource(&input.resource_id, &input.project_id)?;
        let task = self.snapshot.tasks.iter().find(|task| task.id == input.task_id).cloned().ok_or_else(|| format!("task not found: {}", input.task_id))?;
        let resource = self.snapshot.resources.iter().find(|resource| resource.id == input.resource_id).cloned().ok_or_else(|| format!("resource not found: {}", input.resource_id))?;
        let id = input.id.unwrap_or_else(|| generate_id("assignment"));
        let now = now_stamp();
        let created_at = self.snapshot.assignments.iter().find(|assignment| assignment.id == id).map(|assignment| assignment.created_at.clone()).unwrap_or_else(|| now.clone());
        let work_hours = input.work_hours.unwrap_or_else(|| task.duration_hours.unwrap_or(8.0) * f64::from(input.units.unwrap_or(100.0) / 100.0));
        let assignment = AssignmentRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            task_id: input.task_id,
            resource_id: input.resource_id,
            units: input.units.unwrap_or(100.0),
            work_hours,
            actual_work_hours: input.actual_work_hours.unwrap_or(0.0),
            cost: work_hours * resource.standard_rate + resource.cost_per_use,
            created_at,
            updated_at: now,
        };
        upsert_by_id(&mut self.snapshot.assignments, assignment.clone());
        self.recalculate_project(&assignment.project_id)?;
        self.save()?;
        Ok(assignment)
    }

    pub fn delete_assignment(&mut self, assignment_id: &str) -> Result<(), String> {
        let project_id = self.snapshot.assignments.iter().find(|assignment| assignment.id == assignment_id).map(|assignment| assignment.project_id.clone()).ok_or_else(|| format!("assignment not found: {assignment_id}"))?;
        self.snapshot.assignments.retain(|assignment| assignment.id != assignment_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_calendar(&mut self, input: CalendarInput) -> Result<CalendarRecord, String> {
        if input.name.trim().is_empty() {
            return Err("calendar name is required".to_string());
        }
        self.require_project(&input.project_id)?;
        let id = input.id.unwrap_or_else(|| generate_id("calendar"));
        let now = now_stamp();
        let created_at = self.snapshot.calendars.iter().find(|calendar| calendar.id == id).map(|calendar| calendar.created_at.clone()).unwrap_or_else(|| now.clone());
        let calendar = CalendarRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            name: input.name.trim().to_string(),
            timezone: input.timezone,
            working_days: input.working_days.unwrap_or_else(|| vec![1, 2, 3, 4, 5]),
            hours_per_day: input.hours_per_day.unwrap_or(8.0),
            working_hours: input.working_hours.unwrap_or_else(default_working_hours),
            exceptions: input.exceptions.unwrap_or_default(),
            created_at,
            updated_at: now,
        };
        upsert_by_id(&mut self.snapshot.calendars, calendar.clone());
        self.recalculate_project(&calendar.project_id)?;
        self.save()?;
        Ok(calendar)
    }

    pub fn delete_calendar(&mut self, calendar_id: &str) -> Result<(), String> {
        let project_id = self.snapshot.calendars.iter().find(|calendar| calendar.id == calendar_id).map(|calendar| calendar.project_id.clone()).ok_or_else(|| format!("calendar not found: {calendar_id}"))?;
        self.snapshot.calendars.retain(|calendar| calendar.id != calendar_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn capture_baseline(&mut self, project_id: &str, name: Option<String>) -> Result<BaselineRecord, String> {
        self.require_project(project_id)?;
        self.recalculate_project(project_id)?;
        let project = self.snapshot.projects.iter().find(|project| project.id == project_id).cloned().ok_or_else(|| format!("project not found: {project_id}"))?;
        let tasks = self.snapshot.tasks.iter().filter(|task| task.project_id == project_id).map(|task| TaskSnapshot {
            id: task.id.clone(),
            name: task.name.clone(),
            start_date: task.calculated_start_date.clone().or_else(|| task.start_date.clone()),
            finish_date: task.calculated_finish_date.clone().or_else(|| task.finish_date.clone()),
            percent_complete: task.percent_complete,
        }).collect::<Vec<_>>();
        let resources = self.snapshot.resources.iter().filter(|resource| resource.project_id == project_id).map(|resource| ResourceSnapshot {
            id: resource.id.clone(),
            name: resource.name.clone(),
            max_units: resource.max_units,
        }).collect::<Vec<_>>();
        let assignments = self.snapshot.assignments.iter().filter(|assignment| assignment.project_id == project_id).map(|assignment| AssignmentSnapshot {
            id: assignment.id.clone(),
            task_id: assignment.task_id.clone(),
            resource_id: assignment.resource_id.clone(),
            units: assignment.units,
        }).collect::<Vec<_>>();
        let baseline = BaselineRecord {
            id: generate_id("baseline"),
            project_id: project_id.to_string(),
            name: name.unwrap_or_else(|| "Baseline".to_string()),
            captured_at: now_stamp(),
            project: ProjectSnapshot {
                id: project.id.clone(),
                name: project.name.clone(),
                calculated_start_date: project.calculated_start_date.clone().or_else(|| project.start_date.clone()),
                calculated_finish_date: project.calculated_finish_date.clone().or_else(|| project.finish_date.clone()),
            },
            tasks,
            resources,
            assignments,
        };
        self.snapshot.baselines.push(baseline.clone());
        self.save()?;
        Ok(baseline)
    }

    fn require_project(&self, project_id: &str) -> Result<(), String> {
        self.snapshot.projects.iter().any(|project| project.id == project_id).then_some(()).ok_or_else(|| format!("project not found: {project_id}"))
    }

    fn require_task(&self, task_id: &str, project_id: &str) -> Result<(), String> {
        self.snapshot.tasks.iter().any(|task| task.id == task_id && task.project_id == project_id).then_some(()).ok_or_else(|| format!("task not found: {task_id}"))
    }

    fn require_resource(&self, resource_id: &str, project_id: &str) -> Result<(), String> {
        self.snapshot.resources.iter().any(|resource| resource.id == resource_id && resource.project_id == project_id).then_some(()).ok_or_else(|| format!("resource not found: {resource_id}"))
    }

    fn save(&self) -> Result<(), String> {
        if let Some(parent) = self.path.parent() {
            fs::create_dir_all(parent).map_err(|error| format!("failed to create store directory {parent:?}: {error}"))?;
        }
        let serialized = serde_json::to_string_pretty(&self.snapshot).map_err(|error| format!("failed to serialize workspace snapshot: {error}"))?;
        fs::write(&self.path, serialized).map_err(|error| format!("failed to write store file {:?}: {error}", self.path))
    }
}

fn default_store_path() -> PathBuf {
    if let Ok(path) = env::var("MICROPROJECT_STORE_PATH") {
        return PathBuf::from(path);
    }
    dirs::data_local_dir().unwrap_or_else(|| PathBuf::from(".")).join("MicroProject").join("workspace-state.json")
}

fn upsert_by_id<T: Clone>(items: &mut Vec<T>, item: T)
where
    T: HasId,
{
    let id = item.id().to_string();
    if let Some(existing) = items.iter_mut().find(|existing| existing.id() == id) {
        *existing = item;
    } else {
        items.push(item);
    }
}

pub trait HasId {
    fn id(&self) -> &str;
}

impl HasId for ProjectRecord { fn id(&self) -> &str { &self.id } }
impl HasId for TaskRecord { fn id(&self) -> &str { &self.id } }
impl HasId for DependencyRecord { fn id(&self) -> &str { &self.id } }
impl HasId for ResourceRecord { fn id(&self) -> &str { &self.id } }
impl HasId for AssignmentRecord { fn id(&self) -> &str { &self.id } }
impl HasId for CalendarRecord { fn id(&self) -> &str { &self.id } }

fn rebuild_all(snapshot: &mut WorkspaceSnapshot) -> Result<(), String> {
    let project_ids = snapshot.projects.iter().map(|project| project.id.clone()).collect::<Vec<_>>();
    for project_id in project_ids {
        rebuild_project(snapshot, &project_id)?;
    }
    Ok(())
}

fn rebuild_project(snapshot: &mut WorkspaceSnapshot, project_id: &str) -> Result<(), String> {
    let project_idx = snapshot.projects.iter().position(|project| project.id == project_id).ok_or_else(|| format!("project not found: {project_id}"))?;
    let task_indices = snapshot.tasks.iter().enumerate().filter(|(_, task)| task.project_id == project_id).map(|(index, _)| index).collect::<Vec<_>>();
    let task_ids = task_indices.iter().map(|index| snapshot.tasks[*index].id.clone()).collect::<Vec<_>>();
    let children_by_parent = build_children_map(snapshot, project_id);
    let deps_by_successor = build_inbound_dependencies(snapshot, project_id);
    let calendar_lookup = build_calendar_lookup(snapshot, project_id);
    let default_calendar_id = snapshot.settings.default_calendar_id.clone();
    let task_index_map = snapshot.tasks.iter().enumerate().filter(|(_, task)| task.project_id == project_id).map(|(index, task)| (task.id.clone(), index)).collect::<HashMap<_, _>>();
    let order = topo_sort(snapshot, project_id, &task_ids)?;
    let mut scheduled: HashMap<String, (Option<NaiveDate>, Option<NaiveDate>)> = HashMap::new();
    for task_id in order {
        let idx = *task_index_map.get(&task_id).ok_or_else(|| format!("task not found while rebuilding: {task_id}"))?;
        let task = snapshot.tasks[idx].clone();
        let children = children_by_parent.get(&task_id).cloned().unwrap_or_default();
        let calendar = select_calendar(&calendar_lookup, task.calendar_id.as_deref(), default_calendar_id.as_deref());
        if !children.is_empty() {
            let mut child_start = None;
            let mut child_finish = None;
            for child_id in children {
                if let Some((start, finish)) = scheduled.get(&child_id) {
                    child_start = min_date(child_start, *start);
                    child_finish = max_date(child_finish, *finish);
                }
            }
            let start = child_start.or_else(|| parse_optional_date(task.start_date.as_deref())).or_else(|| parse_optional_date(snapshot.projects[project_idx].start_date.as_deref()));
            let finish = child_finish.or_else(|| parse_optional_date(task.finish_date.as_deref())).or_else(|| start.map(|s| compute_finish(s, task.duration_hours.unwrap_or(8.0), calendar.clone())));
            let outline = outline_level(snapshot, project_id, &task_id);
            let wbs = compute_wbs(snapshot, project_id, &task_id, &children_by_parent)?;
            let task_mut = &mut snapshot.tasks[idx];
            task_mut.outline_level = outline;
            task_mut.wbs = wbs;
            task_mut.calculated_start_date = start.map(format_date);
            task_mut.calculated_finish_date = finish.and_then(|f| Some(format_date(f)));
            if task_mut.start_date.is_none() { task_mut.start_date = task_mut.calculated_start_date.clone(); }
            if task_mut.finish_date.is_none() { task_mut.finish_date = task_mut.calculated_finish_date.clone(); }
            scheduled.insert(task_id.clone(), (start, finish));
            continue;
        }
        let mut candidate_start = parse_optional_date(task.start_date.as_deref()).or_else(|| parse_optional_date(snapshot.projects[project_idx].start_date.as_deref()));
        for dependency in deps_by_successor.get(&task_id).into_iter().flatten() {
            if let Some((pred_start, pred_finish)) = scheduled.get(&dependency.predecessor_task_id) {
                let dep_start = dependency_candidate_start(dependency.relation.as_str(), dependency.lag_hours, *pred_start, *pred_finish, hours_per_day(calendar.as_ref()), calendar.as_ref())?;
                candidate_start = max_date(candidate_start, dep_start);
            }
        }
        let start = candidate_start.ok_or_else(|| format!("task has no schedulable start date: {task_id}"))?;
        let finish = compute_finish(start, task.duration_hours.unwrap_or(8.0), calendar.clone());
        let outline = outline_level(snapshot, project_id, &task_id);
        let wbs = compute_wbs(snapshot, project_id, &task_id, &children_by_parent)?;
        let task_mut = &mut snapshot.tasks[idx];
        task_mut.outline_level = outline;
        task_mut.wbs = wbs;
        task_mut.calculated_start_date = Some(format_date(start));
        task_mut.calculated_finish_date = Some(format_date(finish));
        if task_mut.start_date.is_none() { task_mut.start_date = task_mut.calculated_start_date.clone(); }
        if task_mut.finish_date.is_none() { task_mut.finish_date = task_mut.calculated_finish_date.clone(); }
        scheduled.insert(task_id.clone(), (Some(start), Some(finish)));
    }
    let mut min_start = None;
    let mut max_finish = None;
    for task in snapshot.tasks.iter().filter(|task| task.project_id == project_id) {
        let start = parse_optional_date(task.calculated_start_date.as_deref()).or_else(|| parse_optional_date(task.start_date.as_deref()));
        let finish = parse_optional_date(task.calculated_finish_date.as_deref()).or_else(|| parse_optional_date(task.finish_date.as_deref()));
        min_start = min_date(min_start, start);
        max_finish = max_date(max_finish, finish);
    }
    let project = &mut snapshot.projects[project_idx];
    project.calculated_start_date = min_start.map(format_date).or_else(|| project.start_date.clone());
    project.calculated_finish_date = max_finish.map(format_date).or_else(|| project.finish_date.clone());
    Ok(())
}

fn build_children_map(snapshot: &WorkspaceSnapshot, project_id: &str) -> HashMap<String, Vec<String>> {
    let mut map: HashMap<String, Vec<(i64, String)>> = HashMap::new();
    for task in snapshot.tasks.iter().filter(|task| task.project_id == project_id) {
        if let Some(parent_id) = &task.parent_task_id {
            map.entry(parent_id.clone()).or_default().push((task.sort_order, task.id.clone()));
        }
    }
    map.into_iter().map(|(parent, mut children)| {
        children.sort_by(|left, right| left.0.cmp(&right.0).then(left.1.cmp(&right.1)));
        (parent, children.into_iter().map(|(_, id)| id).collect())
    }).collect()
}

fn build_inbound_dependencies(snapshot: &WorkspaceSnapshot, project_id: &str) -> HashMap<String, Vec<DependencyRecord>> {
    let mut map = HashMap::new();
    for dependency in snapshot.dependencies.iter().filter(|dependency| dependency.project_id == project_id) {
        map.entry(dependency.successor_task_id.clone()).or_insert_with(Vec::new).push(dependency.clone());
    }
    map
}

fn build_calendar_lookup(snapshot: &WorkspaceSnapshot, project_id: &str) -> HashMap<String, CalendarRecord> {
    snapshot.calendars.iter().filter(|calendar| calendar.project_id == project_id).map(|calendar| (calendar.id.clone(), calendar.clone())).collect()
}

fn topo_sort(snapshot: &WorkspaceSnapshot, project_id: &str, task_ids: &[String]) -> Result<Vec<String>, String> {
    let mut indegree: HashMap<String, usize> = task_ids.iter().map(|id| (id.clone(), 0)).collect();
    let mut edges: HashMap<String, Vec<String>> = HashMap::new();
    for dependency in snapshot.dependencies.iter().filter(|dependency| dependency.project_id == project_id) {
        if indegree.contains_key(&dependency.predecessor_task_id) && indegree.contains_key(&dependency.successor_task_id) {
            edges.entry(dependency.predecessor_task_id.clone()).or_default().push(dependency.successor_task_id.clone());
            *indegree.get_mut(&dependency.successor_task_id).unwrap() += 1;
        }
    }
    let mut queue = VecDeque::new();
    for task_id in task_ids {
        if indegree.get(task_id) == Some(&0) {
            queue.push_back(task_id.clone());
        }
    }
    let mut ordered = Vec::with_capacity(task_ids.len());
    let mut seen = BTreeSet::new();
    while let Some(task_id) = queue.pop_front() {
        if !seen.insert(task_id.clone()) { continue; }
        ordered.push(task_id.clone());
        if let Some(children) = edges.get(&task_id) {
            for child in children {
                if let Some(degree) = indegree.get_mut(child) {
                    *degree -= 1;
                    if *degree == 0 {
                        queue.push_back(child.clone());
                    }
                }
            }
        }
    }
    if ordered.len() != task_ids.len() {
        return Err("dependency cycle detected while rebuilding project schedule".to_string());
    }
    Ok(ordered)
}

fn outline_level(snapshot: &WorkspaceSnapshot, project_id: &str, task_id: &str) -> u32 {
    let mut level = 1;
    let mut current = snapshot.tasks.iter().find(|task| task.project_id == project_id && task.id == task_id).and_then(|task| task.parent_task_id.clone());
    while let Some(parent_id) = current {
        level += 1;
        current = snapshot.tasks.iter().find(|task| task.project_id == project_id && task.id == parent_id).and_then(|task| task.parent_task_id.clone());
    }
    level
}

fn compute_wbs(snapshot: &WorkspaceSnapshot, project_id: &str, task_id: &str, children_by_parent: &HashMap<String, Vec<String>>) -> Result<String, String> {
    let task = snapshot.tasks.iter().find(|task| task.project_id == project_id && task.id == task_id).ok_or_else(|| format!("task not found: {task_id}"))?;
    match &task.parent_task_id {
        Some(parent_id) => {
            let parent_wbs = snapshot.tasks.iter().find(|candidate| candidate.project_id == project_id && candidate.id == *parent_id).map(|candidate| candidate.wbs.clone()).unwrap_or_default();
            let siblings = children_by_parent.get(parent_id).cloned().unwrap_or_default();
            let position = siblings.iter().position(|candidate| candidate == task_id).unwrap_or(0) + 1;
            Ok(if parent_wbs.is_empty() { position.to_string() } else { format!("{parent_wbs}.{position}") })
        }
        None => {
            let roots = snapshot.tasks.iter().filter(|candidate| candidate.project_id == project_id && candidate.parent_task_id.is_none()).cloned().collect::<Vec<_>>();
            let mut roots_sorted = roots;
            roots_sorted.sort_by(|left, right| left.sort_order.cmp(&right.sort_order).then(left.id.cmp(&right.id)));
            let position = roots_sorted.iter().position(|candidate| candidate.id == task_id).unwrap_or(0) + 1;
            Ok(position.to_string())
        }
    }
}

fn parse_optional_date(input: Option<&str>) -> Option<NaiveDate> {
    input.and_then(|value| NaiveDate::parse_from_str(value, "%Y-%m-%d").ok())
}

fn format_date(date: NaiveDate) -> String { date.format("%Y-%m-%d").to_string() }

fn min_date(left: Option<NaiveDate>, right: Option<NaiveDate>) -> Option<NaiveDate> {
    match (left, right) {
        (Some(a), Some(b)) => Some(a.min(b)),
        (Some(a), None) => Some(a),
        (None, Some(b)) => Some(b),
        _ => None,
    }
}

fn max_date(left: Option<NaiveDate>, right: Option<NaiveDate>) -> Option<NaiveDate> {
    match (left, right) {
        (Some(a), Some(b)) => Some(a.max(b)),
        (Some(a), None) => Some(a),
        (None, Some(b)) => Some(b),
        _ => None,
    }
}

fn hours_per_day(calendar: Option<&CalendarRecord>) -> f64 {
    calendar.map(|calendar| calendar.hours_per_day as f64).unwrap_or(8.0).max(1.0)
}

fn select_calendar(
    calendar_lookup: &HashMap<String, CalendarRecord>,
    explicit_calendar_id: Option<&str>,
    default_calendar_id: Option<&str>,
) -> Option<CalendarRecord> {
    explicit_calendar_id.and_then(|calendar_id| calendar_lookup.get(calendar_id).cloned())
        .or_else(|| default_calendar_id.and_then(|calendar_id| calendar_lookup.get(calendar_id).cloned()))
        .or_else(|| calendar_lookup.values().next().cloned())
}

fn is_working_day(calendar: Option<&CalendarRecord>, date: NaiveDate) -> bool {
    let day = date.weekday().number_from_monday() as u8;
    match calendar {
        Some(calendar) => {
            if let Some(exception) = calendar.exceptions.iter().find(|exception| exception.date == format_date(date)) {
                return exception.is_working;
            }
            calendar.working_days.contains(&day)
        }
        None => true,
    }
}

fn next_working_day(calendar: Option<&CalendarRecord>, date: NaiveDate) -> NaiveDate {
    let mut current = date;
    while !is_working_day(calendar, current) {
        current += ChronoDuration::days(1);
    }
    current
}

fn add_work_days(calendar: Option<&CalendarRecord>, date: NaiveDate, work_days: i64) -> NaiveDate {
    let mut current = date;
    let mut remaining = work_days.max(0);
    while remaining > 0 {
        current += ChronoDuration::days(1);
        if is_working_day(calendar, current) {
            remaining -= 1;
        }
    }
    next_working_day(calendar, current)
}

fn compute_finish(start: NaiveDate, duration_hours: f64, calendar: Option<CalendarRecord>) -> NaiveDate {
    if duration_hours <= 0.0 {
        return start;
    }
    let day_hours = hours_per_day(calendar.as_ref());
    let duration_days = (duration_hours / day_hours).ceil() as i64;
    add_work_days(calendar.as_ref(), start, duration_days.saturating_sub(1))
}

fn dependency_candidate_start(
    relation: &str,
    lag_hours: f64,
    predecessor_start: Option<NaiveDate>,
    predecessor_finish: Option<NaiveDate>,
    hours_per_day: f64,
    calendar: Option<&CalendarRecord>,
) -> Result<Option<NaiveDate>, String> {
    let lag_days = (lag_hours / hours_per_day.max(1.0)).ceil() as i64;
    let result = match relation {
        "FS" => predecessor_finish.map(|finish| add_work_days(calendar, finish, lag_days + 1)),
        "SS" => predecessor_start.map(|start| add_work_days(calendar, start, lag_days)),
        "FF" => predecessor_finish.map(|finish| add_work_days(calendar, finish, lag_days)),
        "SF" => predecessor_start.map(|start| add_work_days(calendar, start, lag_days + 1)),
        other => return Err(format!("unsupported dependency relation: {other}")),
    };
    Ok(result)
}

fn default_working_hours() -> Vec<WorkInterval> {
    vec![WorkInterval { start: "09:00".to_string(), finish: "17:00".to_string() }]
}

fn export_workspace_xml(snapshot: &WorkspaceSnapshot) -> Result<String, String> {
    let mut xml = String::new();
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str(&format!("<microproject schema_version=\"{}\">\n", snapshot.schema_version));
    xml.push_str("  <settings");
    write_opt_attr(&mut xml, "default_calendar_id", snapshot.settings.default_calendar_id.as_deref());
    write_attr(&mut xml, "units_per_day", snapshot.settings.units_per_day.to_string());
    write_opt_attr(&mut xml, "locale", snapshot.settings.locale.as_deref());
    xml.push_str(" />\n");
    write_projects(&mut xml, &snapshot.projects);
    write_tasks(&mut xml, &snapshot.tasks);
    write_dependencies(&mut xml, &snapshot.dependencies);
    write_resources(&mut xml, &snapshot.resources);
    write_assignments(&mut xml, &snapshot.assignments);
    write_calendars(&mut xml, &snapshot.calendars);
    write_baselines(&mut xml, &snapshot.baselines);
    xml.push_str("</microproject>\n");
    Ok(xml)
}

fn import_workspace_xml(xml: &str) -> Result<WorkspaceSnapshot, String> {
    let doc = Document::parse(xml).map_err(|error| format!("failed to parse XML snapshot: {error}"))?;
    let root = doc.root_element();
    if root.tag_name().name() != "microproject" {
        return Err("unexpected XML root element; expected <microproject>".to_string());
    }
    let mut snapshot = WorkspaceSnapshot::new();
    if let Some(schema_version) = attr_u32(root, "schema_version")? {
        snapshot.schema_version = schema_version;
    }
    for child in root.children().filter(|node| node.is_element()) {
        match child.tag_name().name() {
            "settings" => snapshot.settings = parse_settings(child)?,
            "projects" => snapshot.projects = parse_projects(child)?,
            "tasks" => snapshot.tasks = parse_tasks(child)?,
            "dependencies" => snapshot.dependencies = parse_dependencies(child)?,
            "resources" => snapshot.resources = parse_resources(child)?,
            "assignments" => snapshot.assignments = parse_assignments(child)?,
            "calendars" => snapshot.calendars = parse_calendars(child)?,
            "baselines" => snapshot.baselines = parse_baselines(child)?,
            _ => {}
        }
    }
    Ok(snapshot)
}

fn parse_settings(node: Node) -> Result<WorkspaceSettings, String> {
    Ok(WorkspaceSettings {
        default_calendar_id: attr_string(node, "default_calendar_id"),
        units_per_day: attr_f32(node, "units_per_day")?.unwrap_or(8.0),
        locale: attr_string(node, "locale"),
    })
}

fn parse_projects(node: Node) -> Result<Vec<ProjectRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "project").map(|child| {
        Ok(ProjectRecord {
            id: attr_required_string(child, "id")?,
            name: attr_required_string(child, "name")?,
            description: attr_string(child, "description"),
            manager: attr_string(child, "manager"),
            status: attr_string(child, "status").unwrap_or_else(|| "planning".to_string()),
            priority: attr_i32(child, "priority")?.unwrap_or(500),
            calendar_id: attr_string(child, "calendar_id"),
            start_date: attr_string(child, "start_date"),
            finish_date: attr_string(child, "finish_date"),
            calculated_start_date: attr_string(child, "calculated_start_date"),
            calculated_finish_date: attr_string(child, "calculated_finish_date"),
            notes: attr_string(child, "notes").unwrap_or_default(),
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_tasks(node: Node) -> Result<Vec<TaskRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "task").map(|child| {
        Ok(TaskRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            parent_task_id: attr_string(child, "parent_task_id"),
            name: attr_required_string(child, "name")?,
            description: attr_string(child, "description"),
            outline_level: attr_u32(child, "outline_level")?.unwrap_or_default(),
            wbs: attr_string(child, "wbs").unwrap_or_default(),
            start_date: attr_string(child, "start_date"),
            finish_date: attr_string(child, "finish_date"),
            calculated_start_date: attr_string(child, "calculated_start_date"),
            calculated_finish_date: attr_string(child, "calculated_finish_date"),
            duration_hours: attr_f64(child, "duration_hours")?,
            work_hours: attr_f64(child, "work_hours")?,
            percent_complete: attr_f32(child, "percent_complete")?.unwrap_or_default(),
            milestone: attr_bool(child, "milestone")?.unwrap_or(false),
            constraint_type: attr_string(child, "constraint_type").unwrap_or_else(|| "ASAP".to_string()),
            calendar_id: attr_string(child, "calendar_id"),
            notes: attr_string(child, "notes").unwrap_or_default(),
            sort_order: attr_i64(child, "sort_order")?.unwrap_or_default(),
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_dependencies(node: Node) -> Result<Vec<DependencyRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "dependency").map(|child| {
        Ok(DependencyRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            predecessor_task_id: attr_required_string(child, "predecessor_task_id")?,
            successor_task_id: attr_required_string(child, "successor_task_id")?,
            relation: attr_string(child, "relation").unwrap_or_else(|| "FS".to_string()),
            lag_hours: attr_f64(child, "lag_hours")?.unwrap_or_default(),
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_resources(node: Node) -> Result<Vec<ResourceRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "resource").map(|child| {
        Ok(ResourceRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            name: attr_required_string(child, "name")?,
            resource_type: attr_string(child, "resource_type").unwrap_or_else(|| "work".to_string()),
            max_units: attr_f32(child, "max_units")?.unwrap_or(100.0),
            standard_rate: attr_f64(child, "standard_rate")?.unwrap_or_default(),
            overtime_rate: attr_f64(child, "overtime_rate")?.unwrap_or_default(),
            cost_per_use: attr_f64(child, "cost_per_use")?.unwrap_or_default(),
            calendar_id: attr_string(child, "calendar_id"),
            notes: attr_string(child, "notes").unwrap_or_default(),
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_assignments(node: Node) -> Result<Vec<AssignmentRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "assignment").map(|child| {
        Ok(AssignmentRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            task_id: attr_required_string(child, "task_id")?,
            resource_id: attr_required_string(child, "resource_id")?,
            units: attr_f32(child, "units")?.unwrap_or(100.0),
            work_hours: attr_f64(child, "work_hours")?.unwrap_or_default(),
            actual_work_hours: attr_f64(child, "actual_work_hours")?.unwrap_or_default(),
            cost: attr_f64(child, "cost")?.unwrap_or_default(),
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_calendars(node: Node) -> Result<Vec<CalendarRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "calendar").map(|child| {
        let working_days = child.children().find(|section| section.is_element() && section.tag_name().name() == "working_days").map(parse_working_days).transpose()?.unwrap_or_else(|| vec![1, 2, 3, 4, 5]);
        let working_hours = child.children().find(|section| section.is_element() && section.tag_name().name() == "working_hours").map(parse_working_hours).transpose()?.unwrap_or_else(default_working_hours);
        let exceptions = child.children().find(|section| section.is_element() && section.tag_name().name() == "exceptions").map(parse_exceptions).transpose()?.unwrap_or_default();
        Ok(CalendarRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            name: attr_required_string(child, "name")?,
            timezone: attr_string(child, "timezone"),
            working_days,
            hours_per_day: attr_f32(child, "hours_per_day")?.unwrap_or(8.0),
            working_hours,
            exceptions,
            created_at: attr_string(child, "created_at").unwrap_or_default(),
            updated_at: attr_string(child, "updated_at").unwrap_or_default(),
        })
    }).collect()
}

fn parse_working_days(node: Node) -> Result<Vec<u8>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "day").map(|child| {
        attr_u8(child, "value")?.ok_or_else(|| "missing working day value".to_string())
    }).collect()
}

fn parse_working_hours(node: Node) -> Result<Vec<WorkInterval>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "interval").map(|child| {
        Ok(WorkInterval { start: attr_required_string(child, "start")?, finish: attr_required_string(child, "finish")? })
    }).collect()
}

fn parse_exceptions(node: Node) -> Result<Vec<CalendarException>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "exception").map(|child| {
        let intervals = child.children().filter(|grandchild| grandchild.is_element() && grandchild.tag_name().name() == "interval").map(|grandchild| {
            Ok(WorkInterval { start: attr_required_string(grandchild, "start")?, finish: attr_required_string(grandchild, "finish")? })
        }).collect::<Result<Vec<_>, String>>()?;
        Ok(CalendarException {
            date: attr_required_string(child, "date")?,
            is_working: attr_bool(child, "is_working")?.unwrap_or(false),
            intervals,
        })
    }).collect()
}

fn parse_baselines(node: Node) -> Result<Vec<BaselineRecord>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "baseline").map(|child| {
        let project = child.children().find(|section| section.is_element() && section.tag_name().name() == "project").ok_or_else(|| "baseline is missing a project snapshot".to_string())?;
        let tasks = child.children().find(|section| section.is_element() && section.tag_name().name() == "tasks").map(parse_baseline_tasks).transpose()?.unwrap_or_default();
        let resources = child.children().find(|section| section.is_element() && section.tag_name().name() == "resources").map(parse_baseline_resources).transpose()?.unwrap_or_default();
        let assignments = child.children().find(|section| section.is_element() && section.tag_name().name() == "assignments").map(parse_baseline_assignments).transpose()?.unwrap_or_default();
        Ok(BaselineRecord {
            id: attr_required_string(child, "id")?,
            project_id: attr_required_string(child, "project_id")?,
            name: attr_required_string(child, "name")?,
            captured_at: attr_string(child, "captured_at").unwrap_or_default(),
            project: ProjectSnapshot {
                id: attr_required_string(project, "id")?,
                name: attr_required_string(project, "name")?,
                calculated_start_date: attr_string(project, "calculated_start_date"),
                calculated_finish_date: attr_string(project, "calculated_finish_date"),
            },
            tasks,
            resources,
            assignments,
        })
    }).collect()
}

fn parse_baseline_tasks(node: Node) -> Result<Vec<TaskSnapshot>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "task").map(|child| {
        Ok(TaskSnapshot {
            id: attr_required_string(child, "id")?,
            name: attr_required_string(child, "name")?,
            start_date: attr_string(child, "start_date"),
            finish_date: attr_string(child, "finish_date"),
            percent_complete: attr_f32(child, "percent_complete")?.unwrap_or_default(),
        })
    }).collect()
}

fn parse_baseline_resources(node: Node) -> Result<Vec<ResourceSnapshot>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "resource").map(|child| {
        Ok(ResourceSnapshot {
            id: attr_required_string(child, "id")?,
            name: attr_required_string(child, "name")?,
            max_units: attr_f32(child, "max_units")?.unwrap_or(100.0),
        })
    }).collect()
}

fn parse_baseline_assignments(node: Node) -> Result<Vec<AssignmentSnapshot>, String> {
    node.children().filter(|child| child.is_element() && child.tag_name().name() == "assignment").map(|child| {
        Ok(AssignmentSnapshot {
            id: attr_required_string(child, "id")?,
            task_id: attr_required_string(child, "task_id")?,
            resource_id: attr_required_string(child, "resource_id")?,
            units: attr_f32(child, "units")?.unwrap_or(100.0),
        })
    }).collect()
}

fn write_projects(xml: &mut String, projects: &[ProjectRecord]) {
    xml.push_str("  <projects>\n");
    for project in projects {
        xml.push_str("    <project");
        write_attr(xml, "id", &project.id);
        write_attr(xml, "name", &project.name);
        write_opt_attr(xml, "description", project.description.as_deref());
        write_opt_attr(xml, "manager", project.manager.as_deref());
        write_attr(xml, "status", &project.status);
        write_attr(xml, "priority", project.priority.to_string());
        write_opt_attr(xml, "calendar_id", project.calendar_id.as_deref());
        write_opt_attr(xml, "start_date", project.start_date.as_deref());
        write_opt_attr(xml, "finish_date", project.finish_date.as_deref());
        write_opt_attr(xml, "calculated_start_date", project.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", project.calculated_finish_date.as_deref());
        write_attr(xml, "notes", &project.notes);
        write_attr(xml, "created_at", &project.created_at);
        write_attr(xml, "updated_at", &project.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </projects>\n");
}

fn write_tasks(xml: &mut String, tasks: &[TaskRecord]) {
    xml.push_str("  <tasks>\n");
    for task in tasks {
        xml.push_str("    <task");
        write_attr(xml, "id", &task.id);
        write_attr(xml, "project_id", &task.project_id);
        write_opt_attr(xml, "parent_task_id", task.parent_task_id.as_deref());
        write_attr(xml, "name", &task.name);
        write_opt_attr(xml, "description", task.description.as_deref());
        write_attr(xml, "outline_level", task.outline_level.to_string());
        write_attr(xml, "wbs", &task.wbs);
        write_opt_attr(xml, "start_date", task.start_date.as_deref());
        write_opt_attr(xml, "finish_date", task.finish_date.as_deref());
        write_opt_attr(xml, "calculated_start_date", task.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", task.calculated_finish_date.as_deref());
        write_opt_attr(xml, "duration_hours", task.duration_hours.map(|v| v.to_string()).as_deref());
        write_opt_attr(xml, "work_hours", task.work_hours.map(|v| v.to_string()).as_deref());
        write_attr(xml, "percent_complete", task.percent_complete.to_string());
        write_attr(xml, "milestone", task.milestone.to_string());
        write_attr(xml, "constraint_type", &task.constraint_type);
        write_opt_attr(xml, "calendar_id", task.calendar_id.as_deref());
        write_attr(xml, "notes", &task.notes);
        write_attr(xml, "sort_order", task.sort_order.to_string());
        write_attr(xml, "created_at", &task.created_at);
        write_attr(xml, "updated_at", &task.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </tasks>\n");
}

fn write_dependencies(xml: &mut String, dependencies: &[DependencyRecord]) {
    xml.push_str("  <dependencies>\n");
    for dependency in dependencies {
        xml.push_str("    <dependency");
        write_attr(xml, "id", &dependency.id);
        write_attr(xml, "project_id", &dependency.project_id);
        write_attr(xml, "predecessor_task_id", &dependency.predecessor_task_id);
        write_attr(xml, "successor_task_id", &dependency.successor_task_id);
        write_attr(xml, "relation", &dependency.relation);
        write_attr(xml, "lag_hours", dependency.lag_hours.to_string());
        write_attr(xml, "created_at", &dependency.created_at);
        write_attr(xml, "updated_at", &dependency.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </dependencies>\n");
}

fn write_resources(xml: &mut String, resources: &[ResourceRecord]) {
    xml.push_str("  <resources>\n");
    for resource in resources {
        xml.push_str("    <resource");
        write_attr(xml, "id", &resource.id);
        write_attr(xml, "project_id", &resource.project_id);
        write_attr(xml, "name", &resource.name);
        write_attr(xml, "resource_type", &resource.resource_type);
        write_attr(xml, "max_units", resource.max_units.to_string());
        write_attr(xml, "standard_rate", resource.standard_rate.to_string());
        write_attr(xml, "overtime_rate", resource.overtime_rate.to_string());
        write_attr(xml, "cost_per_use", resource.cost_per_use.to_string());
        write_opt_attr(xml, "calendar_id", resource.calendar_id.as_deref());
        write_attr(xml, "notes", &resource.notes);
        write_attr(xml, "created_at", &resource.created_at);
        write_attr(xml, "updated_at", &resource.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </resources>\n");
}

fn write_assignments(xml: &mut String, assignments: &[AssignmentRecord]) {
    xml.push_str("  <assignments>\n");
    for assignment in assignments {
        xml.push_str("    <assignment");
        write_attr(xml, "id", &assignment.id);
        write_attr(xml, "project_id", &assignment.project_id);
        write_attr(xml, "task_id", &assignment.task_id);
        write_attr(xml, "resource_id", &assignment.resource_id);
        write_attr(xml, "units", assignment.units.to_string());
        write_attr(xml, "work_hours", assignment.work_hours.to_string());
        write_attr(xml, "actual_work_hours", assignment.actual_work_hours.to_string());
        write_attr(xml, "cost", assignment.cost.to_string());
        write_attr(xml, "created_at", &assignment.created_at);
        write_attr(xml, "updated_at", &assignment.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </assignments>\n");
}

fn write_calendars(xml: &mut String, calendars: &[CalendarRecord]) {
    xml.push_str("  <calendars>\n");
    for calendar in calendars {
        xml.push_str("    <calendar");
        write_attr(xml, "id", &calendar.id);
        write_attr(xml, "project_id", &calendar.project_id);
        write_attr(xml, "name", &calendar.name);
        write_opt_attr(xml, "timezone", calendar.timezone.as_deref());
        write_attr(xml, "hours_per_day", calendar.hours_per_day.to_string());
        write_attr(xml, "created_at", &calendar.created_at);
        write_attr(xml, "updated_at", &calendar.updated_at);
        xml.push_str(">\n");
        xml.push_str("      <working_days>\n");
        for day in &calendar.working_days {
            xml.push_str(&format!("        <day value=\"{}\" />\n", day));
        }
        xml.push_str("      </working_days>\n");
        xml.push_str("      <working_hours>\n");
        for interval in &calendar.working_hours {
            xml.push_str("        <interval");
            write_attr(xml, "start", &interval.start);
            write_attr(xml, "finish", &interval.finish);
            xml.push_str(" />\n");
        }
        xml.push_str("      </working_hours>\n");
        xml.push_str("      <exceptions>\n");
        for exception in &calendar.exceptions {
            xml.push_str("        <exception");
            write_attr(xml, "date", &exception.date);
            write_attr(xml, "is_working", exception.is_working.to_string());
            xml.push_str(">\n");
            for interval in &exception.intervals {
                xml.push_str("          <interval");
                write_attr(xml, "start", &interval.start);
                write_attr(xml, "finish", &interval.finish);
                xml.push_str(" />\n");
            }
            xml.push_str("        </exception>\n");
        }
        xml.push_str("      </exceptions>\n");
        xml.push_str("    </calendar>\n");
    }
    xml.push_str("  </calendars>\n");
}

fn write_baselines(xml: &mut String, baselines: &[BaselineRecord]) {
    xml.push_str("  <baselines>\n");
    for baseline in baselines {
        xml.push_str("    <baseline");
        write_attr(xml, "id", &baseline.id);
        write_attr(xml, "project_id", &baseline.project_id);
        write_attr(xml, "name", &baseline.name);
        write_attr(xml, "captured_at", &baseline.captured_at);
        xml.push_str(">\n");
        xml.push_str("      <project");
        write_attr(xml, "id", &baseline.project.id);
        write_attr(xml, "name", &baseline.project.name);
        write_opt_attr(xml, "calculated_start_date", baseline.project.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", baseline.project.calculated_finish_date.as_deref());
        xml.push_str(" />\n");
        xml.push_str("    </baseline>\n");
    }
    xml.push_str("  </baselines>\n");
}

fn write_attr(xml: &mut String, name: &str, value: impl AsRef<str>) {
    xml.push_str(&format!(" {}=\"{}\"", name, escape_attr(value.as_ref())));
}

fn write_opt_attr(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        write_attr(xml, name, value);
    }
}

fn escape_attr(value: &str) -> String {
    value.replace('&', "&amp;").replace('"', "&quot;").replace('<', "&lt;").replace('>', "&gt;")
}

fn attr_string(node: Node, name: &str) -> Option<String> {
    node.attribute(name).map(ToString::to_string)
}

fn attr_required_string(node: Node, name: &str) -> Result<String, String> {
    attr_string(node, name).ok_or_else(|| format!("missing required XML attribute: {name}"))
}

fn attr_bool(node: Node, name: &str) -> Result<Option<bool>, String> {
    attr_optional(node, name, |value| match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        other => Err(format!("invalid boolean value for {name}: {other}")),
    })
}

fn attr_u8(node: Node, name: &str) -> Result<Option<u8>, String> {
    attr_optional(node, name, |value| value.parse::<u8>().map_err(|error| format!("invalid u8 value for {name}: {error}")))
}

fn attr_u32(node: Node, name: &str) -> Result<Option<u32>, String> {
    attr_optional(node, name, |value| value.parse::<u32>().map_err(|error| format!("invalid u32 value for {name}: {error}")))
}

fn attr_i32(node: Node, name: &str) -> Result<Option<i32>, String> {
    attr_optional(node, name, |value| value.parse::<i32>().map_err(|error| format!("invalid i32 value for {name}: {error}")))
}

fn attr_i64(node: Node, name: &str) -> Result<Option<i64>, String> {
    attr_optional(node, name, |value| value.parse::<i64>().map_err(|error| format!("invalid i64 value for {name}: {error}")))
}

fn attr_f32(node: Node, name: &str) -> Result<Option<f32>, String> {
    attr_optional(node, name, |value| value.parse::<f32>().map_err(|error| format!("invalid f32 value for {name}: {error}")))
}

fn attr_f64(node: Node, name: &str) -> Result<Option<f64>, String> {
    attr_optional(node, name, |value| value.parse::<f64>().map_err(|error| format!("invalid f64 value for {name}: {error}")))
}

fn attr_optional<T>(node: Node, name: &str, parser: impl FnOnce(&str) -> Result<T, String>) -> Result<Option<T>, String> {
    match node.attribute(name) {
        Some(value) => parser(value).map(Some),
        None => Ok(None),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xml_round_trip_basic() {
        let mut snapshot = WorkspaceSnapshot::new();
        snapshot.projects.push(ProjectRecord { id: "p1".to_string(), name: "Alpha".to_string(), start_date: Some("2026-05-01".to_string()), ..Default::default() });
        snapshot.tasks.push(TaskRecord { id: "t1".to_string(), project_id: "p1".to_string(), name: "Task".to_string(), duration_hours: Some(8.0), sort_order: 1, ..Default::default() });
        let xml = export_workspace_xml(&snapshot).unwrap();
        let imported = import_workspace_xml(&xml).unwrap();
        assert_eq!(imported.projects.len(), 1);
        assert_eq!(imported.tasks.len(), 1);
    }
}

use crate::mspdi::{
    load_project_document, parse_date_time, save_project_document, GanttDependency, GanttTask,
    ProjectDocument,
};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize)]
pub struct ProjectSnapshot {
    pub path: Option<String>,
    pub name: String,
    pub title: Option<String>,
    pub manager: Option<String>,
    pub start_date: Option<String>,
    pub finish_date: Option<String>,
    pub calendars: Vec<CalendarSnapshot>,
    pub tasks: Vec<TaskSnapshot>,
    pub dependencies: Vec<DependencySnapshot>,
    pub chart_range: ChartRangeSnapshot,
    pub dirty: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct CalendarSnapshot {
    pub name: String,
    pub base_calendar: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct ChartRangeSnapshot {
    pub start: String,
    pub end: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct TaskSnapshot {
    pub uid: u32,
    pub id: u32,
    pub name: String,
    pub outline_level: u32,
    pub summary: bool,
    pub milestone: bool,
    pub critical: bool,
    pub percent_complete: f32,
    pub start_text: String,
    pub finish_text: String,
    pub baseline_start_text: Option<String>,
    pub baseline_finish_text: Option<String>,
    pub duration_text: String,
    pub predecessor_text: String,
    pub notes_text: Option<String>,
    pub resource_names: Option<String>,
    pub calendar_uid: Option<u32>,
    pub constraint_type: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct DependencySnapshot {
    pub predecessor_uid: u32,
    pub successor_uid: u32,
    pub relation: String,
    pub lag_text: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct TaskMutationInput {
    pub uid: u32,
    pub name: String,
    pub outline_level: u32,
    pub summary: bool,
    pub milestone: bool,
    pub critical: bool,
    pub percent_complete: f32,
    pub start_text: String,
    pub finish_text: String,
    pub baseline_start_text: Option<String>,
    pub baseline_finish_text: Option<String>,
    pub duration_text: String,
    pub notes_text: Option<String>,
    pub resource_names: Option<String>,
    pub calendar_uid: Option<u32>,
    pub constraint_type: Option<String>,
}

impl Default for TaskMutationInput {
    fn default() -> Self {
        Self {
            uid: 0,
            name: String::new(),
            outline_level: 1,
            summary: false,
            milestone: false,
            critical: false,
            percent_complete: 0.0,
            start_text: String::new(),
            finish_text: String::new(),
            baseline_start_text: None,
            baseline_finish_text: None,
            duration_text: String::new(),
            notes_text: None,
            resource_names: None,
            calendar_uid: None,
            constraint_type: None,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct DependencyMutationInput {
    pub predecessor_uid: u32,
    pub successor_uid: u32,
    pub relation: String,
    pub lag_text: Option<String>,
}

impl Default for DependencyMutationInput {
    fn default() -> Self {
        Self {
            predecessor_uid: 0,
            successor_uid: 0,
            relation: "FS".to_string(),
            lag_text: Some("0".to_string()),
        }
    }
}

pub struct ProjectAppState {
    pub session: std::sync::Mutex<ProjectSession>,
}

pub struct ProjectSession {
    document: ProjectDocument,
    path: Option<PathBuf>,
    dirty: bool,
}

impl ProjectSession {
    pub fn new(initial_path: Option<PathBuf>) -> Self {
        if let Some(path) = initial_path {
            if let Ok(document) = load_project_document(&path) {
                return Self {
                    document,
                    path: Some(path),
                    dirty: false,
                };
            }
        }

        let fallback = default_sample_path();
        if let Some(path) = fallback {
            if let Ok(document) = load_project_document(&path) {
                return Self {
                    document,
                    path: Some(path),
                    dirty: false,
                };
            }
        }

        Self {
            document: empty_document(),
            path: None,
            dirty: false,
        }
    }

    pub fn snapshot(&self) -> ProjectSnapshot {
        let chart_range = self.document.chart_range();
        ProjectSnapshot {
            path: self.path.as_ref().map(|path| path.display().to_string()),
            name: self.document.name.clone(),
            title: self.document.title.clone(),
            manager: self.document.manager.clone(),
            start_date: self
                .document
                .start_date
                .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string()),
            finish_date: self
                .document
                .finish_date
                .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string()),
            calendars: self
                .document
                .calendars
                .iter()
                .map(|calendar| CalendarSnapshot {
                    name: calendar.name.clone(),
                    base_calendar: calendar.base_calendar,
                })
                .collect(),
            tasks: self
                .document
                .tasks
                .iter()
                .map(TaskSnapshot::from)
                .collect(),
            dependencies: self
                .document
                .dependencies
                .iter()
                .map(DependencySnapshot::from)
                .collect(),
            chart_range: ChartRangeSnapshot {
                start: chart_range.start.format("%Y-%m-%d").to_string(),
                end: chart_range.end.format("%Y-%m-%d").to_string(),
            },
            dirty: self.dirty,
        }
    }

    pub fn load_into_current(&mut self, path: PathBuf) -> Result<ProjectSnapshot, String> {
        let document = load_project_document(&path)?;
        self.document = document;
        self.path = Some(path);
        self.dirty = false;
        Ok(self.snapshot())
    }

    pub fn save(&mut self) -> Result<ProjectSnapshot, String> {
        let path = self
            .path
            .clone()
            .ok_or_else(|| "no project file is currently open".to_string())?;
        save_project_document(&path, &self.document)?;
        self.dirty = false;
        Ok(self.snapshot())
    }

    pub fn save_as(&mut self, path: PathBuf) -> Result<ProjectSnapshot, String> {
        save_project_document(&path, &self.document)?;
        self.path = Some(path);
        self.dirty = false;
        Ok(self.snapshot())
    }

    pub fn upsert_task(&mut self, input: TaskMutationInput) -> Result<ProjectSnapshot, String> {
        let task = self
            .document
            .tasks
            .iter_mut()
            .find(|task| task.uid == input.uid)
            .ok_or_else(|| format!("task {} not found", input.uid))?;
        task.name = input.name;
        task.outline_level = input.outline_level.max(1);
        task.summary = input.summary;
        task.milestone = input.milestone;
        task.critical = input.critical;
        task.percent_complete = input.percent_complete.clamp(0.0, 100.0);
        task.start_text = input.start_text.clone();
        task.finish_text = input.finish_text.clone();
        task.start = parse_date_time(Some(&input.start_text));
        task.finish = parse_date_time(Some(&input.finish_text));
        task.baseline_start = input
            .baseline_start_text
            .as_deref()
            .and_then(|value| parse_date_time(Some(value)));
        task.baseline_finish = input
            .baseline_finish_text
            .as_deref()
            .and_then(|value| parse_date_time(Some(value)));
        task.duration_text = input.duration_text;
        task.notes_text = input.notes_text.filter(|value| !value.trim().is_empty());
        task.resource_names = input
            .resource_names
            .filter(|value| !value.trim().is_empty());
        task.calendar_uid = input.calendar_uid;
        task.constraint_type = input
            .constraint_type
            .filter(|value| !value.trim().is_empty());
        self.dirty = true;
        Ok(self.snapshot())
    }

    pub fn delete_task(&mut self, uid: u32) -> Result<ProjectSnapshot, String> {
        let initial_len = self.document.tasks.len();
        self.document.tasks.retain(|task| task.uid != uid);
        self.document
            .dependencies
            .retain(|dependency| dependency.predecessor_uid != uid && dependency.successor_uid != uid);
        if self.document.tasks.len() == initial_len {
            return Err(format!("task {} not found", uid));
        }
        self.dirty = true;
        Ok(self.snapshot())
    }

    pub fn create_task(&mut self, after_uid: Option<u32>) -> Result<ProjectSnapshot, String> {
        let next_uid = self
            .document
            .tasks
            .iter()
            .map(|task| task.uid)
            .max()
            .unwrap_or(0)
            + 1;
        let next_id = self
            .document
            .tasks
            .iter()
            .map(|task| task.id)
            .max()
            .unwrap_or(0)
            + 1;
        let today = chrono::Local::now().date_naive();
        let start = self
            .document
            .start_date
            .map(|dt| dt.date())
            .unwrap_or(today);
        let finish = start.succ_opt().unwrap_or(start);
        let insert_at = after_uid
            .and_then(|uid| self.document.tasks.iter().position(|task| task.uid == uid))
            .map(|index| index + 1)
            .unwrap_or(self.document.tasks.len());
        self.document.tasks.insert(
            insert_at,
            GanttTask {
                uid: next_uid,
                id: next_id,
                name: "New Task".to_string(),
                outline_level: 1,
                summary: false,
                milestone: false,
                critical: false,
                percent_complete: 0.0,
                start_text: start.format("%Y-%m-%dT08:00:00").to_string(),
                finish_text: finish.format("%Y-%m-%dT17:00:00").to_string(),
                start: start.and_hms_opt(8, 0, 0),
                finish: finish.and_hms_opt(17, 0, 0),
                baseline_start: None,
                baseline_finish: None,
                duration_text: "PT8H0M0S".to_string(),
                predecessor_text: String::new(),
                notes_text: None,
                resource_names: None,
                calendar_uid: None,
                constraint_type: None,
            },
        );
        self.dirty = true;
        Ok(self.snapshot())
    }

    pub fn upsert_dependency(
        &mut self,
        input: DependencyMutationInput,
    ) -> Result<ProjectSnapshot, String> {
        if input.predecessor_uid == 0 || input.successor_uid == 0 {
            return Err("both predecessor_uid and successor_uid are required".to_string());
        }
        let relation = normalize_relation(&input.relation);
        let lag_text = input.lag_text.filter(|value| !value.trim().is_empty());
        if let Some(existing) = self.document.dependencies.iter_mut().find(|dependency| {
            dependency.predecessor_uid == input.predecessor_uid
                && dependency.successor_uid == input.successor_uid
                && dependency.relation == relation
        }) {
            existing.lag_text = lag_text;
        } else {
            self.document.dependencies.push(GanttDependency {
                predecessor_uid: input.predecessor_uid,
                successor_uid: input.successor_uid,
                relation,
                lag_text,
            });
        }
        self.rewrite_predecessor_text();
        self.dirty = true;
        Ok(self.snapshot())
    }

    pub fn delete_dependency(
        &mut self,
        input: DependencyMutationInput,
    ) -> Result<ProjectSnapshot, String> {
        let relation = normalize_relation(&input.relation);
        let before = self.document.dependencies.len();
        self.document.dependencies.retain(|dependency| {
            !(dependency.predecessor_uid == input.predecessor_uid
                && dependency.successor_uid == input.successor_uid
                && dependency.relation == relation)
        });
        self.rewrite_predecessor_text();
        if self.document.dependencies.len() == before {
            return Err("dependency not found".to_string());
        }
        self.dirty = true;
        Ok(self.snapshot())
    }

    fn rewrite_predecessor_text(&mut self) {
        let mut by_successor: HashMap<u32, Vec<&GanttDependency>> = HashMap::new();
        for dependency in &self.document.dependencies {
            by_successor
                .entry(dependency.successor_uid)
                .or_default()
                .push(dependency);
        }

        for task in &mut self.document.tasks {
            let text = by_successor
                .get(&task.uid)
                .map(|deps| {
                    deps.iter()
                        .map(|dependency| {
                            let lag = dependency.lag_text.as_deref().unwrap_or("").trim();
                            if lag.is_empty() {
                                dependency.relation.clone()
                            } else {
                                format!("{} {}", dependency.relation, lag)
                            }
                        })
                        .collect::<Vec<_>>()
                        .join(", ")
                })
                .unwrap_or_default();
            task.predecessor_text = text;
        }
    }
}

impl From<&GanttTask> for TaskSnapshot {
    fn from(task: &GanttTask) -> Self {
        Self {
            uid: task.uid,
            id: task.id,
            name: task.name.clone(),
            outline_level: task.outline_level,
            summary: task.summary,
            milestone: task.milestone,
            critical: task.critical,
            percent_complete: task.percent_complete,
            start_text: task.start_text.clone(),
            finish_text: task.finish_text.clone(),
            baseline_start_text: task
                .baseline_start
                .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string()),
            baseline_finish_text: task
                .baseline_finish
                .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string()),
            duration_text: task.duration_text.clone(),
            predecessor_text: task.predecessor_text.clone(),
            notes_text: task.notes_text.clone(),
            resource_names: task.resource_names.clone(),
            calendar_uid: task.calendar_uid,
            constraint_type: task.constraint_type.clone(),
        }
    }
}

impl From<&GanttDependency> for DependencySnapshot {
    fn from(dependency: &GanttDependency) -> Self {
        Self {
            predecessor_uid: dependency.predecessor_uid,
            successor_uid: dependency.successor_uid,
            relation: dependency.relation.clone(),
            lag_text: dependency.lag_text.clone(),
        }
    }
}

fn normalize_relation(relation: &str) -> String {
    match relation.trim().to_uppercase().as_str() {
        "SS" => "SS".to_string(),
        "FF" => "FF".to_string(),
        "SF" => "SF".to_string(),
        _ => "FS".to_string(),
    }
}

fn default_sample_path() -> Option<PathBuf> {
    let candidate = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("testdata/demo-showcase.xml");
    candidate.exists().then_some(candidate)
}

fn empty_document() -> ProjectDocument {
    let today = chrono::Local::now().date_naive();
    ProjectDocument {
        name: "Untitled Project".to_string(),
        title: Some("Untitled Project".to_string()),
        manager: None,
        start_date: today.and_hms_opt(8, 0, 0),
        finish_date: today.and_hms_opt(17, 0, 0),
        calendars: Vec::new(),
        tasks: Vec::new(),
        dependencies: Vec::new(),
    }
}

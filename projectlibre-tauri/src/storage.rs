use crate::{
    engine,
    interop,
    models::*,
};
use std::{
    env,
    fs,
    path::PathBuf,
    sync::Mutex,
};

pub struct AppState {
    pub store: Mutex<AppStore>,
}

impl AppState {
    pub fn new(store: AppStore) -> Self {
        Self {
            store: Mutex::new(store),
        }
    }
}

pub struct AppStore {
    path: PathBuf,
    snapshot: WorkspaceSnapshot,
}

impl AppStore {
    pub fn load_or_default() -> Self {
        match Self::load() {
            Ok(store) => store,
            Err(error) => {
                eprintln!("MicroProject store load failed: {error}");
                Self::new(default_store_path())
            }
        }
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
        Self {
            path,
            snapshot: WorkspaceSnapshot::new(),
        }
    }

    pub fn snapshot(&self) -> WorkspaceSnapshot {
        self.snapshot.clone()
    }

    pub fn export_json(&self) -> Result<String, String> {
        serde_json::to_string_pretty(&self.snapshot)
            .map_err(|error| format!("failed to serialize snapshot: {error}"))
    }

    pub fn export_xml(&self) -> Result<String, String> {
        interop::export_workspace_xml(&self.snapshot)
    }

    pub fn import_json(&mut self, json: &str) -> Result<(), String> {
        self.snapshot = serde_json::from_str::<WorkspaceSnapshot>(json)
            .map_err(|error| format!("failed to parse snapshot: {error}"))?;
        self.recalculate_all()?;
        self.save()
    }

    pub fn import_xml(&mut self, xml: &str) -> Result<(), String> {
        self.snapshot = interop::import_workspace_xml(xml)?;
        self.recalculate_all()?;
        self.save()
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        engine::rebuild_all(&mut self.snapshot).map(|_| ())
    }

    pub fn recalculate_all_and_save(&mut self) -> Result<(), String> {
        self.recalculate_all()?;
        self.save()
    }

    pub fn recalculate_project(&mut self, project_id: &str) -> Result<(), String> {
        engine::rebuild_project(&mut self.snapshot, project_id).map(|_| ())
    }

    pub fn upsert_project(&mut self, input: ProjectInput) -> Result<ProjectRecord, String> {
        let name = input.name.trim();
        if name.is_empty() {
            return Err("project name is required".to_string());
        }

        let now = now_stamp();
        let id = input.id.unwrap_or_else(|| generate_id("project"));
        let created_at = self
            .snapshot
            .projects
            .iter()
            .find(|project| project.id == id)
            .map(|project| project.created_at.clone())
            .unwrap_or_else(|| now.clone());

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

        if let Some(existing) = self.snapshot.projects.iter_mut().find(|project| project.id == id) {
            *existing = project.clone();
        } else {
            self.snapshot.projects.push(project.clone());
        }

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
        self.snapshot
            .dependencies
            .retain(|dependency| dependency.project_id != project_id);
        self.snapshot
            .resources
            .retain(|resource| resource.project_id != project_id);
        self.snapshot
            .assignments
            .retain(|assignment| assignment.project_id != project_id);
        self.snapshot
            .calendars
            .retain(|calendar| calendar.project_id != project_id);
        self.snapshot
            .baselines
            .retain(|baseline| baseline.project_id != project_id);

        self.save()
    }

    pub fn upsert_task(&mut self, input: TaskInput) -> Result<TaskRecord, String> {
        if input.name.trim().is_empty() {
            return Err("task name is required".to_string());
        }
        self.require_project(&input.project_id)?;
        if let Some(parent_id) = &input.parent_task_id {
            self.require_task(parent_id, &input.project_id)?;
        }

        let id = input.id.unwrap_or_else(|| generate_id("task"));
        let now = now_stamp();
        let created_at = self
            .snapshot
            .tasks
            .iter()
            .find(|task| task.id == id)
            .map(|task| task.created_at.clone())
            .unwrap_or_else(|| now.clone());

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

        if let Some(existing) = self.snapshot.tasks.iter_mut().find(|task| task.id == id) {
            *existing = task.clone();
        } else {
            self.snapshot.tasks.push(task.clone());
        }

        self.recalculate_project(&task.project_id)?;
        self.save()?;
        Ok(task)
    }

    pub fn delete_task(&mut self, task_id: &str) -> Result<(), String> {
        let project_id = self
            .snapshot
            .tasks
            .iter()
            .find(|task| task.id == task_id)
            .map(|task| task.project_id.clone())
            .ok_or_else(|| format!("task not found: {task_id}"))?;

        self.snapshot.tasks.retain(|task| task.id != task_id);
        self.snapshot.dependencies.retain(|dependency| {
            dependency.predecessor_task_id != task_id && dependency.successor_task_id != task_id
        });
        self.snapshot.assignments.retain(|assignment| assignment.task_id != task_id);

        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_dependency(
        &mut self,
        input: DependencyInput,
    ) -> Result<DependencyRecord, String> {
        if input.predecessor_task_id == input.successor_task_id {
            return Err("dependency cannot point to the same task".to_string());
        }

        self.require_project(&input.project_id)?;
        self.require_task(&input.predecessor_task_id, &input.project_id)?;
        self.require_task(&input.successor_task_id, &input.project_id)?;

        let relation = input.relation.unwrap_or_else(|| "FS".to_string());
        let id = input.id.unwrap_or_else(|| generate_id("dependency"));
        let now = now_stamp();
        let created_at = self
            .snapshot
            .dependencies
            .iter()
            .find(|dependency| dependency.id == id)
            .map(|dependency| dependency.created_at.clone())
            .unwrap_or_else(|| now.clone());

        let dependency = DependencyRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            predecessor_task_id: input.predecessor_task_id,
            successor_task_id: input.successor_task_id,
            relation,
            lag_hours: input.lag_hours.unwrap_or(0.0),
            created_at,
            updated_at: now,
        };

        if let Some(existing) = self
            .snapshot
            .dependencies
            .iter_mut()
            .find(|dependency| dependency.id == id)
        {
            *existing = dependency.clone();
        } else {
            self.snapshot.dependencies.push(dependency.clone());
        }

        self.recalculate_project(&dependency.project_id)?;
        self.save()?;
        Ok(dependency)
    }

    pub fn delete_dependency(&mut self, dependency_id: &str) -> Result<(), String> {
        let project_id = self
            .snapshot
            .dependencies
            .iter()
            .find(|dependency| dependency.id == dependency_id)
            .map(|dependency| dependency.project_id.clone())
            .ok_or_else(|| format!("dependency not found: {dependency_id}"))?;

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
        let created_at = self
            .snapshot
            .resources
            .iter()
            .find(|resource| resource.id == id)
            .map(|resource| resource.created_at.clone())
            .unwrap_or_else(|| now.clone());

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

        if let Some(existing) = self.snapshot.resources.iter_mut().find(|resource| resource.id == id) {
            *existing = resource.clone();
        } else {
            self.snapshot.resources.push(resource.clone());
        }

        self.recalculate_project(&resource.project_id)?;
        self.save()?;
        Ok(resource)
    }

    pub fn delete_resource(&mut self, resource_id: &str) -> Result<(), String> {
        let project_id = self
            .snapshot
            .resources
            .iter()
            .find(|resource| resource.id == resource_id)
            .map(|resource| resource.project_id.clone())
            .ok_or_else(|| format!("resource not found: {resource_id}"))?;

        self.snapshot.resources.retain(|resource| resource.id != resource_id);
        self.snapshot.assignments.retain(|assignment| assignment.resource_id != resource_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn upsert_assignment(&mut self, input: AssignmentInput) -> Result<AssignmentRecord, String> {
        self.require_project(&input.project_id)?;
        self.require_task(&input.task_id, &input.project_id)?;
        self.require_resource(&input.resource_id, &input.project_id)?;

        let task = self
            .snapshot
            .tasks
            .iter()
            .find(|task| task.id == input.task_id)
            .cloned()
            .ok_or_else(|| format!("task not found: {}", input.task_id))?;
        let resource = self
            .snapshot
            .resources
            .iter()
            .find(|resource| resource.id == input.resource_id)
            .cloned()
            .ok_or_else(|| format!("resource not found: {}", input.resource_id))?;

        let id = input.id.unwrap_or_else(|| generate_id("assignment"));
        let now = now_stamp();
        let created_at = self
            .snapshot
            .assignments
            .iter()
            .find(|assignment| assignment.id == id)
            .map(|assignment| assignment.created_at.clone())
            .unwrap_or_else(|| now.clone());

        let work_hours = input
            .work_hours
            .or(task.duration_hours.map(|duration| duration * f64::from(input.units.unwrap_or(100.0) / 100.0)))
            .unwrap_or(0.0);
        let actual_work_hours = input.actual_work_hours.unwrap_or(0.0);
        let cost = work_hours * resource.standard_rate + resource.cost_per_use;

        let assignment = AssignmentRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            task_id: input.task_id,
            resource_id: input.resource_id,
            units: input.units.unwrap_or(100.0),
            work_hours,
            actual_work_hours,
            cost,
            created_at,
            updated_at: now,
        };

        if let Some(existing) = self
            .snapshot
            .assignments
            .iter_mut()
            .find(|assignment| assignment.id == id)
        {
            *existing = assignment.clone();
        } else {
            self.snapshot.assignments.push(assignment.clone());
        }

        self.recalculate_project(&assignment.project_id)?;
        self.save()?;
        Ok(assignment)
    }

    pub fn delete_assignment(&mut self, assignment_id: &str) -> Result<(), String> {
        let project_id = self
            .snapshot
            .assignments
            .iter()
            .find(|assignment| assignment.id == assignment_id)
            .map(|assignment| assignment.project_id.clone())
            .ok_or_else(|| format!("assignment not found: {assignment_id}"))?;

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
        let created_at = self
            .snapshot
            .calendars
            .iter()
            .find(|calendar| calendar.id == id)
            .map(|calendar| calendar.created_at.clone())
            .unwrap_or_else(|| now.clone());

        let calendar = CalendarRecord {
            id: id.clone(),
            project_id: input.project_id.clone(),
            name: input.name.trim().to_string(),
            timezone: input.timezone,
            working_days: input.working_days.unwrap_or_else(|| vec![1, 2, 3, 4, 5]),
            hours_per_day: input.hours_per_day.unwrap_or(8.0),
            working_hours: input.working_hours.unwrap_or_else(|| {
                vec![WorkInterval {
                    start: "09:00".to_string(),
                    finish: "17:00".to_string(),
                }]
            }),
            exceptions: input.exceptions.unwrap_or_default(),
            created_at,
            updated_at: now,
        };

        if let Some(existing) = self
            .snapshot
            .calendars
            .iter_mut()
            .find(|calendar| calendar.id == id)
        {
            *existing = calendar.clone();
        } else {
            self.snapshot.calendars.push(calendar.clone());
        }

        self.recalculate_project(&calendar.project_id)?;
        self.save()?;
        Ok(calendar)
    }

    pub fn delete_calendar(&mut self, calendar_id: &str) -> Result<(), String> {
        let project_id = self
            .snapshot
            .calendars
            .iter()
            .find(|calendar| calendar.id == calendar_id)
            .map(|calendar| calendar.project_id.clone())
            .ok_or_else(|| format!("calendar not found: {calendar_id}"))?;

        self.snapshot.calendars.retain(|calendar| calendar.id != calendar_id);
        self.recalculate_project(&project_id)?;
        self.save()
    }

    pub fn capture_baseline(&mut self, project_id: &str, name: Option<String>) -> Result<BaselineRecord, String> {
        self.require_project(project_id)?;
        self.recalculate_project(project_id)?;

        let project = self
            .snapshot
            .projects
            .iter()
            .find(|project| project.id == project_id)
            .cloned()
            .ok_or_else(|| format!("project not found: {project_id}"))?;
        let tasks = self
            .snapshot
            .tasks
            .iter()
            .filter(|task| task.project_id == project_id)
            .map(|task| TaskSnapshot {
                id: task.id.clone(),
                name: task.name.clone(),
                start_date: task.calculated_start_date.clone().or_else(|| task.start_date.clone()),
                finish_date: task.calculated_finish_date.clone().or_else(|| task.finish_date.clone()),
                percent_complete: task.percent_complete,
            })
            .collect::<Vec<_>>();
        let resources = self
            .snapshot
            .resources
            .iter()
            .filter(|resource| resource.project_id == project_id)
            .map(|resource| ResourceSnapshot {
                id: resource.id.clone(),
                name: resource.name.clone(),
                max_units: resource.max_units,
            })
            .collect::<Vec<_>>();
        let assignments = self
            .snapshot
            .assignments
            .iter()
            .filter(|assignment| assignment.project_id == project_id)
            .map(|assignment| AssignmentSnapshot {
                id: assignment.id.clone(),
                task_id: assignment.task_id.clone(),
                resource_id: assignment.resource_id.clone(),
                units: assignment.units,
            })
            .collect::<Vec<_>>();

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
        self.snapshot
            .projects
            .iter()
            .any(|project| project.id == project_id)
            .then_some(())
            .ok_or_else(|| format!("project not found: {project_id}"))
    }

    fn require_task(&self, task_id: &str, project_id: &str) -> Result<(), String> {
        self.snapshot
            .tasks
            .iter()
            .any(|task| task.id == task_id && task.project_id == project_id)
            .then_some(())
            .ok_or_else(|| format!("task not found: {task_id}"))
    }

    fn require_resource(&self, resource_id: &str, project_id: &str) -> Result<(), String> {
        self.snapshot
            .resources
            .iter()
            .any(|resource| resource.id == resource_id && resource.project_id == project_id)
            .then_some(())
            .ok_or_else(|| format!("resource not found: {resource_id}"))
    }

    fn save(&self) -> Result<(), String> {
        if let Some(parent) = self.path.parent() {
            fs::create_dir_all(parent)
                .map_err(|error| format!("failed to create store directory {parent:?}: {error}"))?;
        }

        let serialized = serde_json::to_string_pretty(&self.snapshot)
            .map_err(|error| format!("failed to serialize workspace snapshot: {error}"))?;
        fs::write(&self.path, serialized)
            .map_err(|error| format!("failed to write store file {:?}: {error}", self.path))
    }
}

fn default_store_path() -> PathBuf {
    if let Ok(path) = env::var("MICROPROJECT_STORE_PATH") {
        return PathBuf::from(path);
    }

    dirs::data_local_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("MicroProject")
        .join("workspace-state.json")
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    fn sample_project(store: &mut AppStore, name: &str) -> ProjectRecord {
        store
            .upsert_project(ProjectInput {
                id: None,
                name: name.to_string(),
                description: Some("desc".to_string()),
                manager: Some("Manager".to_string()),
                status: Some("planning".to_string()),
                priority: Some(100),
                calendar_id: None,
                start_date: Some("2026-05-01".to_string()),
                finish_date: Some("2026-05-10".to_string()),
                notes: Some("notes".to_string()),
            })
            .expect("project should be created")
    }

    fn sample_task(store: &mut AppStore, project_id: &str, name: &str) -> TaskRecord {
        store
            .upsert_task(TaskInput {
                id: None,
                project_id: project_id.to_string(),
                parent_task_id: None,
                name: name.to_string(),
                description: None,
                start_date: Some("2026-05-01".to_string()),
                finish_date: Some("2026-05-02".to_string()),
                duration_hours: Some(8.0),
                work_hours: Some(8.0),
                percent_complete: Some(25.0),
                milestone: Some(false),
                constraint_type: Some("ASAP".to_string()),
                calendar_id: None,
                notes: None,
                sort_order: Some(1),
            })
            .expect("task should be created")
    }

    #[test]
    fn project_task_and_dependency_round_trip() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("workspace.json");
        let mut store = AppStore::load_from(path.clone()).expect("store");

        let project = sample_project(&mut store, "Alpha");
        let task_a = sample_task(&mut store, &project.id, "Plan");
        let task_b = sample_task(&mut store, &project.id, "Build");
        let dependency = store
            .upsert_dependency(DependencyInput {
                id: None,
                project_id: project.id.clone(),
                predecessor_task_id: task_a.id.clone(),
                successor_task_id: task_b.id.clone(),
                relation: Some("FS".to_string()),
                lag_hours: Some(0.0),
            })
            .expect("dependency");

        let loaded = AppStore::load_from(path).expect("reloaded store");
        assert_eq!(loaded.snapshot.projects.len(), 1);
        assert_eq!(loaded.snapshot.tasks.len(), 2);
        assert_eq!(loaded.snapshot.dependencies.len(), 1);
        assert_eq!(loaded.snapshot.dependencies[0].id, dependency.id);
    }

    #[test]
    fn deleting_project_cascades_to_children() {
        let dir = tempdir().expect("tempdir");
        let mut store = AppStore::load_from(dir.path().join("workspace.json")).expect("store");
        let project = sample_project(&mut store, "Beta");
        let task = sample_task(&mut store, &project.id, "Task");
        store
            .upsert_calendar(CalendarInput {
                id: None,
                project_id: project.id.clone(),
                name: "Default".to_string(),
                timezone: None,
                working_days: None,
                hours_per_day: None,
                working_hours: None,
                exceptions: None,
            })
            .expect("calendar");
        let resource = store
            .upsert_resource(ResourceInput {
                id: None,
                project_id: project.id.clone(),
                name: "Alice".to_string(),
                resource_type: None,
                max_units: Some(100.0),
                standard_rate: Some(50.0),
                overtime_rate: None,
                cost_per_use: None,
                calendar_id: None,
                notes: None,
            })
            .expect("resource");
        store
            .upsert_assignment(AssignmentInput {
                id: None,
                project_id: project.id.clone(),
                task_id: task.id.clone(),
                resource_id: resource.id.clone(),
                units: Some(100.0),
                work_hours: None,
                actual_work_hours: None,
            })
            .expect("assignment");

        store.delete_project(&project.id).expect("delete project");
        assert!(store.snapshot.projects.is_empty());
        assert!(store.snapshot.tasks.is_empty());
        assert!(store.snapshot.dependencies.is_empty());
        assert!(store.snapshot.calendars.is_empty());
        assert!(store.snapshot.resources.is_empty());
        assert!(store.snapshot.assignments.is_empty());
    }

    #[test]
    fn rejects_invalid_dependency_and_task_without_project() {
        let dir = tempdir().expect("tempdir");
        let mut store = AppStore::load_from(dir.path().join("workspace.json")).expect("store");

        assert!(store
            .upsert_task(TaskInput {
                id: None,
                project_id: "missing".to_string(),
                parent_task_id: None,
                name: "Task".to_string(),
                description: None,
                start_date: None,
                finish_date: None,
                duration_hours: None,
                work_hours: None,
                percent_complete: None,
                milestone: None,
                constraint_type: None,
                calendar_id: None,
                notes: None,
                sort_order: None,
            })
            .is_err());

        let project = sample_project(&mut store, "Gamma");
        let task = sample_task(&mut store, &project.id, "One");

        assert!(store
            .upsert_dependency(DependencyInput {
                id: None,
                project_id: project.id.clone(),
                predecessor_task_id: task.id.clone(),
                successor_task_id: task.id.clone(),
                relation: None,
                lag_hours: None,
            })
            .is_err());
    }

    #[test]
    fn xml_round_trip_through_store() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("workspace.json");
        let mut store = AppStore::load_from(path.clone()).expect("store");
        let project = sample_project(&mut store, "Delta");
        let task_a = sample_task(&mut store, &project.id, "Plan");
        let task_b = sample_task(&mut store, &project.id, "Build");
        store
            .upsert_dependency(DependencyInput {
                id: None,
                project_id: project.id.clone(),
                predecessor_task_id: task_a.id.clone(),
                successor_task_id: task_b.id.clone(),
                relation: Some("FS".to_string()),
                lag_hours: Some(0.0),
            })
            .expect("dependency");
        let xml = store.export_xml().expect("export xml");

        let mut imported = AppStore::load_from(dir.path().join("workspace-import.json")).expect("import store");
        imported.import_xml(&xml).expect("import xml");
        assert_eq!(imported.snapshot.projects.len(), 1);
        assert_eq!(imported.snapshot.tasks.len(), 2);
        assert_eq!(imported.snapshot.dependencies.len(), 1);
        assert!(imported.snapshot.projects[0].calculated_start_date.is_some());
    }
}

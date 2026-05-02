use crate::models::*;
use std::{
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
        let path = default_store_path();
        if !path.exists() {
            return Ok(Self::new(path));
        }

        let contents = fs::read_to_string(&path)
            .map_err(|error| format!("failed to read store file {path:?}: {error}"))?;
        let snapshot = serde_json::from_str::<WorkspaceSnapshot>(&contents)
            .map_err(|error| format!("failed to parse store file {path:?}: {error}"))?;

        Ok(Self { path, snapshot })
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
            start_date: input.start_date,
            finish_date: input.finish_date,
            created_at,
            updated_at: now,
        };

        if let Some(existing) = self.snapshot.projects.iter_mut().find(|project| project.id == id) {
            *existing = project.clone();
        } else {
            self.snapshot.projects.push(project.clone());
        }

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
            .calendars
            .retain(|calendar| calendar.project_id != project_id);

        self.save()
    }

    pub fn upsert_task(&mut self, input: TaskInput) -> Result<TaskRecord, String> {
        if input.name.trim().is_empty() {
            return Err("task name is required".to_string());
        }
        if !self
            .snapshot
            .projects
            .iter()
            .any(|project| project.id == input.project_id)
        {
            return Err(format!("project not found: {}", input.project_id));
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
            project_id: input.project_id,
            name: input.name.trim().to_string(),
            description: input.description,
            start_date: input.start_date,
            finish_date: input.finish_date,
            percent_complete: input.percent_complete.unwrap_or(0.0),
            parent_task_id: input.parent_task_id,
            sort_order: input.sort_order.unwrap_or(0),
            created_at,
            updated_at: now,
        };

        if let Some(existing) = self.snapshot.tasks.iter_mut().find(|task| task.id == id) {
            *existing = task.clone();
        } else {
            self.snapshot.tasks.push(task.clone());
        }

        self.save()?;
        Ok(task)
    }

    pub fn delete_task(&mut self, task_id: &str) -> Result<(), String> {
        let before = self.snapshot.tasks.len();
        self.snapshot.tasks.retain(|task| task.id != task_id);
        if self.snapshot.tasks.len() == before {
            return Err(format!("task not found: {task_id}"));
        }

        self.snapshot.dependencies.retain(|dependency| {
            dependency.predecessor_task_id != task_id && dependency.successor_task_id != task_id
        });

        self.save()
    }

    pub fn upsert_dependency(
        &mut self,
        input: DependencyInput,
    ) -> Result<DependencyRecord, String> {
        if input.predecessor_task_id == input.successor_task_id {
            return Err("dependency cannot point to the same task".to_string());
        }

        let relation = input.relation.unwrap_or_else(|| "FS".to_string());
        let predecessor = self
            .snapshot
            .tasks
            .iter()
            .find(|task| task.id == input.predecessor_task_id)
            .ok_or_else(|| format!("predecessor task not found: {}", input.predecessor_task_id))?;
        let successor = self
            .snapshot
            .tasks
            .iter()
            .find(|task| task.id == input.successor_task_id)
            .ok_or_else(|| format!("successor task not found: {}", input.successor_task_id))?;

        if predecessor.project_id != input.project_id || successor.project_id != input.project_id {
            return Err("dependency tasks must belong to the same project".to_string());
        }

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
            project_id: input.project_id,
            predecessor_task_id: input.predecessor_task_id,
            successor_task_id: input.successor_task_id,
            relation,
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

        self.save()?;
        Ok(dependency)
    }

    pub fn delete_dependency(&mut self, dependency_id: &str) -> Result<(), String> {
        let before = self.snapshot.dependencies.len();
        self.snapshot.dependencies.retain(|dependency| dependency.id != dependency_id);
        if self.snapshot.dependencies.len() == before {
            return Err(format!("dependency not found: {dependency_id}"));
        }
        self.save()
    }

    pub fn upsert_calendar(&mut self, input: CalendarInput) -> Result<CalendarRecord, String> {
        if input.name.trim().is_empty() {
            return Err("calendar name is required".to_string());
        }
        if !self
            .snapshot
            .projects
            .iter()
            .any(|project| project.id == input.project_id)
        {
            return Err(format!("project not found: {}", input.project_id));
        }

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
            project_id: input.project_id,
            name: input.name.trim().to_string(),
            timezone: input.timezone,
            working_days: input.working_days.unwrap_or_else(|| vec![1, 2, 3, 4, 5]),
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

        self.save()?;
        Ok(calendar)
    }

    pub fn delete_calendar(&mut self, calendar_id: &str) -> Result<(), String> {
        let before = self.snapshot.calendars.len();
        self.snapshot.calendars.retain(|calendar| calendar.id != calendar_id);
        if self.snapshot.calendars.len() == before {
            return Err(format!("calendar not found: {calendar_id}"));
        }
        self.save()
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
    dirs::data_local_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("MicroProject")
        .join("workspace-state.json")
}


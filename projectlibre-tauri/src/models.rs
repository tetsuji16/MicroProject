use serde::{Serialize, Deserialize};

#[derive(Serialize, Deserialize, Clone, Debug, Default)]
pub struct Task {
    pub id: String,
    pub name: String,
}

#[derive(Serialize, Deserialize, Clone, Debug, Default)]
pub struct Project {
    pub id: String,
    pub name: String,
    pub tasks: Vec<Task>,
}

#[derive(Serialize, Deserialize, Clone, Debug, Default)]
pub struct Workspace {
    pub projects: Vec<Project>,
}

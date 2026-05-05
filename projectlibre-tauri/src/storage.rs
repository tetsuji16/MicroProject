// Simple in-memory store placeholder
use crate::models::{Workspace, Project};

pub struct AppStore {
    pub workspace: Workspace,
}

impl AppStore {
    pub fn load_or_default() -> Self {
        Self { workspace: Workspace { projects: Vec::new() } }
    }
}

pub struct AppState {
    pub store: AppStore,
}

impl AppState {
    pub fn new(store: AppStore) -> Self { Self { store } }
}

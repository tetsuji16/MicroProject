use crate::mspdi::{GanttDependency, ProjectDocument};
use petgraph::graphmap::DiGraphMap;
use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, Default, PartialEq, Eq, Hash, Serialize, Deserialize)]
struct DependencyWeight {
    relation: String,
    lag_text: Option<String>,
}

#[derive(Clone, Debug, Default)]
pub struct DependencyIndex {
    graph: DiGraphMap<u32, DependencyWeight>,
}

impl DependencyIndex {
    pub fn from_document(document: &ProjectDocument) -> Self {
        let mut graph = DiGraphMap::new();
        for dependency in &document.dependencies {
            graph.add_edge(
                dependency.predecessor_uid,
                dependency.successor_uid,
                DependencyWeight {
                    relation: dependency.relation.clone(),
                    lag_text: dependency.lag_text.clone(),
                },
            );
        }

        Self { graph }
    }

    pub fn contains_dependency(&self, predecessor_uid: u32, successor_uid: u32) -> bool {
        self.graph.contains_edge(predecessor_uid, successor_uid)
    }

    pub fn predecessor_text_for_successor(
        &self,
        document: &ProjectDocument,
        successor_uid: u32,
    ) -> String {
        let mut assignments = self
            .graph
            .all_edges()
            .filter(|(_, target, _)| *target == successor_uid)
            .map(|(predecessor_uid, _, dependency)| {
                let predecessor_id = document
                    .tasks
                    .iter()
                    .find(|task| task.uid == predecessor_uid)
                    .map(|task| task.id)
                    .unwrap_or(predecessor_uid);
                let relation = if dependency.relation.trim().is_empty() {
                    "FS"
                } else {
                    dependency.relation.trim()
                };
                let lag = dependency
                    .lag_text
                    .as_deref()
                    .unwrap_or("")
                    .trim()
                    .to_string();
                (predecessor_id, relation.to_string(), lag)
            })
            .collect::<Vec<_>>();

        assignments.sort_by(|left, right| left.0.cmp(&right.0).then_with(|| left.1.cmp(&right.1)));

        assignments
            .into_iter()
            .map(|(predecessor_id, relation, lag)| {
                if lag.is_empty() {
                    format!("{predecessor_id}{relation}")
                } else {
                    format!("{predecessor_id}{relation} {lag}")
                }
            })
            .collect::<Vec<_>>()
            .join(", ")
    }
}

pub fn dependency_exists(
    document: &ProjectDocument,
    predecessor_uid: u32,
    successor_uid: u32,
) -> bool {
    DependencyIndex::from_document(document).contains_dependency(predecessor_uid, successor_uid)
}

pub fn refresh_predecessor_text_for_task(document: &mut ProjectDocument, successor_uid: u32) {
    let Some(task_index) = document
        .tasks
        .iter()
        .position(|task| task.uid == successor_uid)
    else {
        return;
    };

    let text = DependencyIndex::from_document(document)
        .predecessor_text_for_successor(document, successor_uid);
    document.tasks[task_index].predecessor_text = text;
}

pub fn add_dependency_from_drag(
    document: &mut ProjectDocument,
    predecessor_uid: u32,
    successor_uid: u32,
    relation: String,
    lag_text: Option<String>,
) -> bool {
    if predecessor_uid == successor_uid
        || dependency_exists(document, predecessor_uid, successor_uid)
    {
        return false;
    }

    document.dependencies.push(GanttDependency {
        predecessor_uid,
        successor_uid,
        relation,
        lag_text,
    });
    refresh_predecessor_text_for_task(document, successor_uid);
    true
}

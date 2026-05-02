use crate::models::*;
use chrono::{Datelike, Duration as ChronoDuration, NaiveDate};
use std::collections::{BTreeMap, BTreeSet, HashMap, VecDeque};

#[derive(Debug, Clone, Default)]
pub struct RebuildReport {
    pub scheduled_tasks: usize,
    pub summary_tasks: usize,
}

pub fn rebuild_project(snapshot: &mut WorkspaceSnapshot, project_id: &str) -> Result<RebuildReport, String> {
    let project_idx = snapshot
        .projects
        .iter()
        .position(|project| project.id == project_id)
        .ok_or_else(|| format!("project not found: {project_id}"))?;

    let calendar_lookup = build_calendar_lookup(snapshot, project_id);
    let default_calendar_id = snapshot.settings.default_calendar_id.clone();
    let task_indices = snapshot
        .tasks
        .iter()
        .enumerate()
        .filter(|(_, task)| task.project_id == project_id)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let task_ids = task_indices
        .iter()
        .map(|index| snapshot.tasks[*index].id.clone())
        .collect::<Vec<_>>();

    let children_by_parent = build_children_map(snapshot, project_id);
    let deps_by_successor = build_inbound_dependencies(snapshot, project_id);
    let task_map = snapshot
        .tasks
        .iter()
        .enumerate()
        .filter(|(_, task)| task.project_id == project_id)
        .map(|(index, task)| (task.id.clone(), index))
        .collect::<HashMap<_, _>>();

    let root_start = snapshot.projects[project_idx].start_date.clone();

    let order = topo_sort(snapshot, project_id, &task_ids)?;
    let mut report = RebuildReport::default();
    let mut scheduled: HashMap<String, (Option<NaiveDate>, Option<NaiveDate>)> = HashMap::new();

    for task_id in &order {
        let task_index = task_map
            .get(task_id)
            .ok_or_else(|| format!("task not found while rebuilding: {task_id}"))?;
        let task = snapshot.tasks[*task_index].clone();
        let children = children_by_parent.get(task_id).cloned().unwrap_or_default();
        let is_summary = !children.is_empty();

        if is_summary {
            let mut child_start: Option<NaiveDate> = None;
            let mut child_finish: Option<NaiveDate> = None;
            for child_id in children {
                if let Some((start, finish)) = scheduled.get(&child_id) {
                    child_start = min_date(child_start, *start);
                    child_finish = max_date(child_finish, *finish);
                }
            }

            let (start, finish) = if child_start.is_some() && child_finish.is_some() {
                (child_start, child_finish)
            } else {
                let fallback = parse_optional_date(task.start_date.as_deref())
                    .or_else(|| parse_optional_date(root_start.as_deref()))
                    .or_else(|| parse_optional_date(snapshot.projects[project_idx].start_date.as_deref()))
                    .ok_or_else(|| format!("task has no schedulable start date: {task_id}"))?;
                let duration = task.duration_hours.unwrap_or(0.0);
                let derived_finish = compute_finish(
                    fallback,
                    duration,
                    calendar_lookup.get(task.calendar_id.as_deref().unwrap_or("")).cloned(),
                )?;
                (Some(fallback), Some(derived_finish))
                };

            let outline = outline_level(snapshot, project_id, task_id, &children_by_parent);
            let wbs = compute_wbs(snapshot, project_id, task_id, &children_by_parent)?;
            let task_mut = &mut snapshot.tasks[*task_index];
            task_mut.outline_level = outline;
            task_mut.wbs = wbs;
            task_mut.calculated_start_date = start.map(format_date);
            task_mut.calculated_finish_date = finish.map(format_date);
            if task_mut.start_date.is_none() {
                task_mut.start_date = task_mut.calculated_start_date.clone();
            }
            if task_mut.finish_date.is_none() {
                task_mut.finish_date = task_mut.calculated_finish_date.clone();
            }
            scheduled.insert(task_id.clone(), (start, finish));
            report.summary_tasks += 1;
            continue;
        }

        let calendar = select_calendar(
            &calendar_lookup,
            task.calendar_id.as_deref(),
            default_calendar_id.as_deref(),
        );
        let mut candidate_start = parse_optional_date(task.start_date.as_deref())
            .or_else(|| parse_optional_date(snapshot.projects[project_idx].start_date.as_deref()))
            .or_else(|| parse_optional_date(root_start.as_deref()));

        for predecessor in deps_by_successor.get(task_id).into_iter().flatten() {
            if let Some((pred_start, pred_finish)) = scheduled.get(&predecessor.predecessor_task_id) {
                let dependency_start = dependency_candidate_start(
                    predecessor.relation.as_str(),
                    predecessor.lag_hours,
                    *pred_start,
                    *pred_finish,
                    hours_per_day(calendar.as_ref()),
                    calendar.as_ref(),
                )?;
                candidate_start = max_date(candidate_start, dependency_start);
            }
        }

        let start = candidate_start.ok_or_else(|| format!("task has no schedulable start date: {task_id}"))?;
        let finish = compute_finish(
            start,
            task.duration_hours.unwrap_or(8.0),
            calendar.clone(),
        )?;

        let outline = outline_level(snapshot, project_id, task_id, &children_by_parent);
        let wbs = compute_wbs(snapshot, project_id, task_id, &children_by_parent)?;
        let task_mut = &mut snapshot.tasks[*task_index];
        task_mut.outline_level = outline;
        task_mut.wbs = wbs;
        task_mut.calculated_start_date = Some(format_date(start));
        task_mut.calculated_finish_date = Some(format_date(finish));
        if task_mut.start_date.is_none() {
            task_mut.start_date = task_mut.calculated_start_date.clone();
        }
        if task_mut.finish_date.is_none() {
            task_mut.finish_date = task_mut.calculated_finish_date.clone();
        }

        scheduled.insert(task_id.clone(), (Some(start), Some(finish)));
        report.scheduled_tasks += 1;
    }

    let mut child_coverage: BTreeMap<String, (Option<NaiveDate>, Option<NaiveDate>)> = BTreeMap::new();
    for task in snapshot.tasks.iter().filter(|task| task.project_id == project_id) {
        if let Some(parent_id) = &task.parent_task_id {
            let start = parse_optional_date(task.calculated_start_date.as_deref())
                .or_else(|| parse_optional_date(task.start_date.as_deref()));
            let finish = parse_optional_date(task.calculated_finish_date.as_deref())
                .or_else(|| parse_optional_date(task.finish_date.as_deref()));
            let entry = child_coverage.entry(parent_id.clone()).or_insert((None, None));
            entry.0 = min_date(entry.0, start);
            entry.1 = max_date(entry.1, finish);
        }
    }

    for project in snapshot.projects.iter_mut().filter(|project| project.id == project_id) {
        let (start, finish) = child_coverage
            .get(&project.id)
            .cloned()
            .unwrap_or((None, None));
        project.calculated_start_date = start
            .map(format_date)
            .or_else(|| project.start_date.clone());
        project.calculated_finish_date = finish
            .map(format_date)
            .or_else(|| project.finish_date.clone());
    }

    Ok(report)
}

pub fn rebuild_all(snapshot: &mut WorkspaceSnapshot) -> Result<Vec<RebuildReport>, String> {
    let project_ids = snapshot.projects.iter().map(|project| project.id.clone()).collect::<Vec<_>>();
    let mut reports = Vec::with_capacity(project_ids.len());
    for project_id in project_ids {
        reports.push(rebuild_project(snapshot, &project_id)?);
    }
    Ok(reports)
}

fn build_calendar_lookup(snapshot: &WorkspaceSnapshot, project_id: &str) -> HashMap<String, CalendarRecord> {
    snapshot
        .calendars
        .iter()
        .filter(|calendar| calendar.project_id == project_id)
        .map(|calendar| (calendar.id.clone(), calendar.clone()))
        .collect()
}

fn build_children_map(snapshot: &WorkspaceSnapshot, project_id: &str) -> HashMap<String, Vec<String>> {
    let mut map: HashMap<String, Vec<(i64, String)>> = HashMap::new();
    for task in snapshot.tasks.iter().filter(|task| task.project_id == project_id) {
        if let Some(parent_id) = &task.parent_task_id {
            map.entry(parent_id.clone())
                .or_default()
                .push((task.sort_order, task.id.clone()));
        }
    }

    map.into_iter()
        .map(|(parent, mut children)| {
            children.sort_by(|left, right| left.0.cmp(&right.0).then(left.1.cmp(&right.1)));
            (parent, children.into_iter().map(|(_, id)| id).collect())
        })
        .collect()
}

fn build_inbound_dependencies(
    snapshot: &WorkspaceSnapshot,
    project_id: &str,
) -> HashMap<String, Vec<DependencyRecord>> {
    let mut map: HashMap<String, Vec<DependencyRecord>> = HashMap::new();
    for dependency in snapshot.dependencies.iter().filter(|dependency| dependency.project_id == project_id) {
        map.entry(dependency.successor_task_id.clone())
            .or_default()
            .push(dependency.clone());
    }
    map
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
    for task in snapshot.tasks.iter().filter(|task| task.project_id == project_id) {
        if indegree.get(&task.id) == Some(&0) {
            queue.push_back(task.id.clone());
        }
    }

    let mut ordered = Vec::with_capacity(task_ids.len());
    let mut seen = BTreeSet::new();
    while let Some(task_id) = queue.pop_front() {
        if !seen.insert(task_id.clone()) {
            continue;
        }
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

fn outline_level(
    snapshot: &WorkspaceSnapshot,
    project_id: &str,
    task_id: &str,
    children_by_parent: &HashMap<String, Vec<String>>,
) -> u32 {
    let mut level = 1;
    let mut current = snapshot
        .tasks
        .iter()
        .find(|task| task.project_id == project_id && task.id == task_id)
        .and_then(|task| task.parent_task_id.clone());
    while let Some(parent_id) = current {
        level += 1;
        current = snapshot
            .tasks
            .iter()
            .find(|task| task.project_id == project_id && task.id == parent_id)
            .and_then(|task| task.parent_task_id.clone());
    }
    if children_by_parent.contains_key(task_id) {
        level
    } else {
        level
    }
}

fn compute_wbs(
    snapshot: &WorkspaceSnapshot,
    project_id: &str,
    task_id: &str,
    children_by_parent: &HashMap<String, Vec<String>>,
) -> Result<String, String> {
    let task = snapshot
        .tasks
        .iter()
        .find(|task| task.project_id == project_id && task.id == task_id)
        .ok_or_else(|| format!("task not found: {task_id}"))?;

    match &task.parent_task_id {
        Some(parent_id) => {
            let parent_wbs = snapshot
                .tasks
                .iter()
                .find(|candidate| candidate.project_id == project_id && candidate.id == *parent_id)
                .map(|candidate| candidate.wbs.clone())
                .unwrap_or_default();
            let siblings = children_by_parent.get(parent_id).cloned().unwrap_or_default();
            let position = siblings
                .iter()
                .position(|candidate| candidate == task_id)
                .unwrap_or(0)
                + 1;
            Ok(if parent_wbs.is_empty() {
                position.to_string()
            } else {
                format!("{parent_wbs}.{position}")
            })
        }
        None => {
            let roots = snapshot
                .tasks
                .iter()
                .filter(|candidate| candidate.project_id == project_id && candidate.parent_task_id.is_none())
                .cloned()
                .collect::<Vec<_>>();
            let mut roots_sorted = roots;
            roots_sorted.sort_by(|left, right| left.sort_order.cmp(&right.sort_order).then(left.id.cmp(&right.id)));
            let position = roots_sorted
                .iter()
                .position(|candidate| candidate.id == task_id)
                .unwrap_or(0)
                + 1;
            Ok(position.to_string())
        }
    }
}

fn parse_optional_date(input: Option<&str>) -> Option<NaiveDate> {
    input.and_then(|value| NaiveDate::parse_from_str(value, "%Y-%m-%d").ok())
}

fn format_date(date: NaiveDate) -> String {
    date.format("%Y-%m-%d").to_string()
}

fn min_date(left: Option<NaiveDate>, right: Option<NaiveDate>) -> Option<NaiveDate> {
    match (left, right) {
        (Some(a), Some(b)) => Some(a.min(b)),
        (Some(a), None) => Some(a),
        (None, Some(b)) => Some(b),
        (None, None) => None,
    }
}

fn max_date(left: Option<NaiveDate>, right: Option<NaiveDate>) -> Option<NaiveDate> {
    match (left, right) {
        (Some(a), Some(b)) => Some(a.max(b)),
        (Some(a), None) => Some(a),
        (None, Some(b)) => Some(b),
        (None, None) => None,
    }
}

fn hours_per_day(calendar: Option<&CalendarRecord>) -> f64 {
    calendar.map(|calendar| calendar.hours_per_day as f64).unwrap_or(8.0).max(1.0)
}

fn is_working_day(calendar: Option<&CalendarRecord>, date: NaiveDate) -> bool {
    let day = date.weekday().number_from_monday() as u8;
    match calendar {
        Some(calendar) => {
            if let Some(exception) = calendar
                .exceptions
                .iter()
                .find(|exception| exception.date == format_date(date))
            {
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

fn compute_finish(
    start: NaiveDate,
    duration_hours: f64,
    calendar: Option<CalendarRecord>,
) -> Result<NaiveDate, String> {
    if duration_hours <= 0.0 {
        return Ok(start);
    }
    let day_hours = hours_per_day(calendar.as_ref());
    let duration_days = (duration_hours / day_hours).ceil() as i64;
    Ok(add_work_days(calendar.as_ref(), start, duration_days.saturating_sub(1)))
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

fn select_calendar(
    calendar_lookup: &HashMap<String, CalendarRecord>,
    explicit_calendar_id: Option<&str>,
    default_calendar_id: Option<&str>,
) -> Option<CalendarRecord> {
    if let Some(calendar_id) = explicit_calendar_id {
        if let Some(calendar) = calendar_lookup.get(calendar_id) {
            return Some(calendar.clone());
        }
    }

    if let Some(calendar_id) = default_calendar_id {
        if let Some(calendar) = calendar_lookup.get(calendar_id) {
            return Some(calendar.clone());
        }
    }

    calendar_lookup.values().next().cloned()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn base_snapshot() -> WorkspaceSnapshot {
        let mut snapshot = WorkspaceSnapshot::new();
        snapshot.projects.push(ProjectRecord {
            id: "project_1".to_string(),
            name: "Alpha".to_string(),
            start_date: Some("2026-05-01".to_string()),
            ..ProjectRecord::default()
        });
        snapshot.calendars.push(CalendarRecord {
            id: "cal_1".to_string(),
            project_id: "project_1".to_string(),
            name: "Default".to_string(),
            working_days: vec![1, 2, 3, 4, 5],
            hours_per_day: 8.0,
            ..CalendarRecord::default()
        });
        snapshot
    }

    #[test]
    fn rebuilds_outline_and_dependency_schedule() {
        let mut snapshot = base_snapshot();
        snapshot.tasks.push(TaskRecord {
            id: "task_a".to_string(),
            project_id: "project_1".to_string(),
            name: "A".to_string(),
            sort_order: 1,
            duration_hours: Some(8.0),
            calendar_id: Some("cal_1".to_string()),
            ..TaskRecord::default()
        });
        snapshot.tasks.push(TaskRecord {
            id: "task_b".to_string(),
            project_id: "project_1".to_string(),
            name: "B".to_string(),
            sort_order: 2,
            duration_hours: Some(8.0),
            calendar_id: Some("cal_1".to_string()),
            ..TaskRecord::default()
        });
        snapshot.dependencies.push(DependencyRecord {
            id: "dep_1".to_string(),
            project_id: "project_1".to_string(),
            predecessor_task_id: "task_a".to_string(),
            successor_task_id: "task_b".to_string(),
            relation: "FS".to_string(),
            lag_hours: 0.0,
            ..DependencyRecord::default()
        });

        let report = rebuild_project(&mut snapshot, "project_1").expect("rebuild");
        assert_eq!(report.scheduled_tasks, 2);
        assert_eq!(snapshot.tasks[0].wbs, "1");
        assert_eq!(snapshot.tasks[1].wbs, "2");
        assert_eq!(snapshot.tasks[0].calculated_start_date.as_deref(), Some("2026-05-01"));
        assert_eq!(snapshot.tasks[0].calculated_finish_date.as_deref(), Some("2026-05-01"));
        assert_eq!(snapshot.tasks[1].calculated_start_date.as_deref(), Some("2026-05-04"));
    }

    #[test]
    fn detects_cycles() {
        let mut snapshot = base_snapshot();
        snapshot.tasks.push(TaskRecord {
            id: "task_a".to_string(),
            project_id: "project_1".to_string(),
            name: "A".to_string(),
            duration_hours: Some(8.0),
            ..TaskRecord::default()
        });
        snapshot.tasks.push(TaskRecord {
            id: "task_b".to_string(),
            project_id: "project_1".to_string(),
            name: "B".to_string(),
            duration_hours: Some(8.0),
            ..TaskRecord::default()
        });
        snapshot.dependencies.push(DependencyRecord {
            id: "dep_1".to_string(),
            project_id: "project_1".to_string(),
            predecessor_task_id: "task_a".to_string(),
            successor_task_id: "task_b".to_string(),
            relation: "FS".to_string(),
            ..DependencyRecord::default()
        });
        snapshot.dependencies.push(DependencyRecord {
            id: "dep_2".to_string(),
            project_id: "project_1".to_string(),
            predecessor_task_id: "task_b".to_string(),
            successor_task_id: "task_a".to_string(),
            relation: "FS".to_string(),
            ..DependencyRecord::default()
        });

        assert!(rebuild_project(&mut snapshot, "project_1").is_err());
    }
}

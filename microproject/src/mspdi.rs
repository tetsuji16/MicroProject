use chrono::{NaiveDate, NaiveDateTime};
use quick_xml::de::from_str;
use quick_xml::se::to_string;
use serde::{Deserialize, Serialize};
use std::fs::{read_to_string, write};
use std::path::Path;

#[derive(Debug, Clone)]
pub struct ProjectDocument {
    pub name: String,
    pub title: Option<String>,
    pub manager: Option<String>,
    pub start_date: Option<NaiveDateTime>,
    pub finish_date: Option<NaiveDateTime>,
    pub calendars: Vec<CalendarInfo>,
    pub tasks: Vec<GanttTask>,
    pub dependencies: Vec<GanttDependency>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename = "Project")]
struct ProjectXml {
    #[serde(rename = "@xmlns", default, skip_serializing_if = "Option::is_none")]
    xmlns: Option<String>,
    #[serde(rename = "Name", default, skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(rename = "Title", default, skip_serializing_if = "Option::is_none")]
    title: Option<String>,
    #[serde(rename = "Manager", default, skip_serializing_if = "Option::is_none")]
    manager: Option<String>,
    #[serde(rename = "StartDate", default, skip_serializing_if = "Option::is_none")]
    start_date: Option<String>,
    #[serde(
        rename = "FinishDate",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    finish_date: Option<String>,
    #[serde(rename = "Calendars", default, skip_serializing_if = "Option::is_none")]
    calendars: Option<CalendarsXml>,
    #[serde(rename = "Tasks", default)]
    tasks: TasksXml,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct CalendarsXml {
    #[serde(rename = "Calendar", default)]
    calendars: Vec<CalendarXml>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct CalendarXml {
    #[serde(rename = "Name", default, skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(
        rename = "IsBaseCalendar",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    base_calendar: Option<u8>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct TasksXml {
    #[serde(rename = "Task", default)]
    tasks: Vec<TaskXml>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct TaskXml {
    #[serde(rename = "UID", default)]
    uid: u32,
    #[serde(rename = "ID", default)]
    id: u32,
    #[serde(rename = "Name", default, skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(
        rename = "OutlineLevel",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    outline_level: Option<u32>,
    #[serde(rename = "Summary", default, skip_serializing_if = "Option::is_none")]
    summary: Option<u8>,
    #[serde(rename = "Milestone", default, skip_serializing_if = "Option::is_none")]
    milestone: Option<u8>,
    #[serde(rename = "Critical", default, skip_serializing_if = "Option::is_none")]
    critical: Option<u8>,
    #[serde(
        rename = "PercentComplete",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    percent_complete: Option<f32>,
    #[serde(rename = "Start", default, skip_serializing_if = "Option::is_none")]
    start: Option<String>,
    #[serde(rename = "Finish", default, skip_serializing_if = "Option::is_none")]
    finish: Option<String>,
    #[serde(
        rename = "BaselineStart",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    baseline_start: Option<String>,
    #[serde(
        rename = "BaselineFinish",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    baseline_finish: Option<String>,
    #[serde(rename = "Duration", default, skip_serializing_if = "Option::is_none")]
    duration: Option<String>,
    #[serde(rename = "Notes", default, skip_serializing_if = "Option::is_none")]
    notes: Option<String>,
    #[serde(
        rename = "ResourceNames",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    resource_names: Option<String>,
    #[serde(
        rename = "CalendarUID",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    calendar_uid: Option<u32>,
    #[serde(
        rename = "ConstraintType",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    constraint_type: Option<String>,
    #[serde(rename = "PredecessorLink", default)]
    predecessor_links: Vec<PredecessorLinkXml>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct PredecessorLinkXml {
    #[serde(
        rename = "PredecessorUID",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    predecessor_uid: Option<u32>,
    #[serde(rename = "Type", default, skip_serializing_if = "Option::is_none")]
    relation: Option<String>,
    #[serde(rename = "LinkLag", default, skip_serializing_if = "Option::is_none")]
    lag: Option<String>,
}

#[derive(Debug, Clone)]
pub struct CalendarInfo {
    pub name: String,
    pub base_calendar: bool,
}

#[derive(Debug, Clone)]
pub struct GanttTask {
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
    pub start: Option<NaiveDateTime>,
    pub finish: Option<NaiveDateTime>,
    pub baseline_start: Option<NaiveDateTime>,
    pub baseline_finish: Option<NaiveDateTime>,
    pub duration_text: String,
    pub predecessor_text: String,
    pub notes_text: Option<String>,
    pub resource_names: Option<String>,
    pub calendar_uid: Option<u32>,
    pub constraint_type: Option<String>,
}

#[derive(Debug, Clone)]
pub struct GanttDependency {
    pub predecessor_uid: u32,
    pub successor_uid: u32,
    pub relation: String,
    pub lag_text: Option<String>,
}

#[derive(Debug, Clone, Copy)]
pub struct ChartRange {
    pub start: NaiveDate,
    pub end: NaiveDate,
}

impl ChartRange {
    pub fn days(self) -> i64 {
        (self.end - self.start).num_days().max(1)
    }
}

pub fn load_project_document(path: &Path) -> Result<ProjectDocument, String> {
    let xml = read_to_string(path)
        .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
    let xml = xml.trim_start_matches(|ch: char| ch.is_whitespace() || ch == '\u{feff}');
    let project: ProjectXml =
        from_str(xml).map_err(|error| format!("failed to parse {}: {error}", path.display()))?;
    Ok(convert(project, path))
}

pub fn save_project_document(path: &Path, document: &ProjectDocument) -> Result<(), String> {
    let project = ProjectXml {
        xmlns: Some("http://schemas.microsoft.com/project".to_string()),
        name: Some(document.name.clone()),
        title: document.title.clone(),
        manager: document.manager.clone(),
        start_date: document.start_date.map(format_date_time),
        finish_date: document.finish_date.map(format_date_time),
        calendars: (!document.calendars.is_empty()).then(|| CalendarsXml {
            calendars: document
                .calendars
                .iter()
                .map(|calendar| CalendarXml {
                    name: Some(calendar.name.clone()),
                    base_calendar: Some(if calendar.base_calendar { 1 } else { 0 }),
                })
                .collect(),
        }),
        tasks: TasksXml {
            tasks: document
                .tasks
                .iter()
                .map(|task| TaskXml {
                    uid: task.uid,
                    id: task.id,
                    name: (!task.name.trim().is_empty()).then(|| task.name.clone()),
                    outline_level: Some(task.outline_level.max(1)),
                    summary: Some(if task.summary { 1 } else { 0 }),
                    milestone: Some(if task.milestone { 1 } else { 0 }),
                    critical: Some(if task.critical { 1 } else { 0 }),
                    percent_complete: Some(task.percent_complete),
                    start: (!task.start_text.trim().is_empty()).then(|| task.start_text.clone()),
                    finish: (!task.finish_text.trim().is_empty()).then(|| task.finish_text.clone()),
                    baseline_start: task.baseline_start.map(format_date_time),
                    baseline_finish: task.baseline_finish.map(format_date_time),
                    duration: (!task.duration_text.trim().is_empty())
                        .then(|| task.duration_text.clone()),
                    notes: task
                        .notes_text
                        .as_ref()
                        .and_then(|text| (!text.trim().is_empty()).then(|| text.clone())),
                    resource_names: task
                        .resource_names
                        .as_ref()
                        .and_then(|text| (!text.trim().is_empty()).then(|| text.clone())),
                    calendar_uid: task.calendar_uid,
                    constraint_type: task
                        .constraint_type
                        .as_ref()
                        .and_then(|text| (!text.trim().is_empty()).then(|| text.clone())),
                    predecessor_links: document
                        .dependencies
                        .iter()
                        .filter(|dependency| dependency.successor_uid == task.uid)
                        .map(|dependency| PredecessorLinkXml {
                            predecessor_uid: Some(dependency.predecessor_uid),
                            relation: Some(dependency.relation.clone()),
                            lag: dependency.lag_text.clone(),
                        })
                        .collect(),
                })
                .collect(),
        },
    };

    let mut rendered = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    rendered.push_str(
        &to_string(&project)
            .map_err(|error| format!("failed to serialize {}: {error}", path.display()))?,
    );
    rendered.push('\n');

    write(path, rendered).map_err(|error| format!("failed to write {}: {error}", path.display()))
}

fn convert(project: ProjectXml, path: &Path) -> ProjectDocument {
    let fallback_name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled Project")
        .to_string();
    let name = project.name.unwrap_or_else(|| fallback_name.clone());
    let title = project.title.or_else(|| Some(name.clone()));
    let manager = project.manager;
    let start_date = parse_date_time(project.start_date.as_deref());
    let finish_date = parse_date_time(project.finish_date.as_deref());
    let calendars = project
        .calendars
        .map(|wrapper| {
            wrapper
                .calendars
                .into_iter()
                .map(|calendar| CalendarInfo {
                    name: calendar.name.unwrap_or_else(|| "Calendar".to_string()),
                    base_calendar: calendar.base_calendar.unwrap_or(0) != 0,
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();

    let mut tasks = Vec::new();
    let mut dependencies = Vec::new();

    for task in project.tasks.tasks {
        let start_text = task.start.unwrap_or_default();
        let finish_text = task.finish.unwrap_or_default();
        let baseline_start_text = task.baseline_start.unwrap_or_default();
        let baseline_finish_text = task.baseline_finish.unwrap_or_default();
        let start = parse_date_time(Some(start_text.as_str()));
        let finish = parse_date_time(Some(finish_text.as_str()));
        let baseline_start = parse_date_time(Some(baseline_start_text.as_str()));
        let baseline_finish = parse_date_time(Some(baseline_finish_text.as_str()));
        let predecessor_text = task
            .predecessor_links
            .iter()
            .map(format_predecessor_link_data)
            .collect::<Vec<_>>()
            .join(", ");

        for link in task.predecessor_links {
            if let Some(predecessor_uid) = link.predecessor_uid {
                dependencies.push(GanttDependency {
                    predecessor_uid,
                    successor_uid: task.uid,
                    relation: link.relation.unwrap_or_else(|| "FS".to_string()),
                    lag_text: link.lag,
                });
            }
        }

        tasks.push(GanttTask {
            uid: task.uid,
            id: if task.id == 0 { task.uid } else { task.id },
            name: task.name.unwrap_or_default(),
            outline_level: task.outline_level.unwrap_or(1).max(1),
            summary: task.summary.unwrap_or(0) != 0,
            milestone: task.milestone.unwrap_or(0) != 0,
            critical: task.critical.unwrap_or(0) != 0,
            percent_complete: task.percent_complete.unwrap_or(0.0),
            start_text,
            finish_text,
            start,
            finish,
            baseline_start,
            baseline_finish,
            duration_text: task.duration.unwrap_or_default(),
            predecessor_text,
            notes_text: task.notes,
            resource_names: task.resource_names,
            calendar_uid: task.calendar_uid,
            constraint_type: task.constraint_type,
        });
    }

    ProjectDocument {
        name,
        title,
        manager,
        start_date,
        finish_date,
        calendars,
        tasks,
        dependencies,
    }
}

pub fn parse_date_time(text: Option<&str>) -> Option<NaiveDateTime> {
    let text = text?.trim();
    if text.is_empty() {
        return None;
    }

    const FORMATS: &[&str] = &[
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S%.f",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ];

    for format in FORMATS {
        if let Ok(parsed) = NaiveDateTime::parse_from_str(text, format) {
            return Some(parsed);
        }
    }

    NaiveDate::parse_from_str(text, "%Y-%m-%d")
        .ok()
        .and_then(|date| date.and_hms_opt(0, 0, 0))
}

fn format_predecessor_link_data(link: &PredecessorLinkXml) -> String {
    let relation = link.relation.as_deref().unwrap_or("FS");
    let lag = link.lag.as_deref().unwrap_or("").trim();
    if lag.is_empty() {
        relation.to_string()
    } else {
        format!("{relation} {lag}")
    }
}

fn format_date_time(value: NaiveDateTime) -> String {
    value.format("%Y-%m-%dT%H:%M:%S").to_string()
}

impl ProjectDocument {
    pub fn chart_range(&self) -> ChartRange {
        let mut dates = Vec::new();

        if let Some(start) = self.start_date {
            dates.push(start.date());
        }
        if let Some(finish) = self.finish_date {
            dates.push(finish.date());
        }

        for task in &self.tasks {
            if let Some(start) = task.start {
                dates.push(start.date());
            }
            if let Some(finish) = task.finish {
                dates.push(finish.date());
            }
        }

        let today = chrono::Local::now().date_naive();
        let (start, end) = if dates.is_empty() {
            (today, today.succ_opt().unwrap_or(today))
        } else {
            let min = dates.iter().min().copied().unwrap_or(today);
            let max = dates.iter().max().copied().unwrap_or(today);
            // 終了日に1日の余白を設定（視認性向上のため）
            let start = min;
            let end = max.succ_opt().unwrap_or(max);
            (start, end)
        };

        let end = if end <= start {
            start.succ_opt().unwrap_or(start)
        } else {
            end
        };

        ChartRange { start, end }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::io::Write;
    use tempfile::tempdir;

    #[test]
    fn parses_task_and_dependency_xml() {
        let xml = r#"
        <?xml version="1.0" encoding="UTF-8"?>
        <Project xmlns="http://schemas.microsoft.com/project">
          <Name>Demo</Name>
          <StartDate>2026-05-06T08:00:00</StartDate>
          <FinishDate>2026-05-08T17:00:00</FinishDate>
          <Calendars>
            <Calendar>
              <UID>1</UID>
              <Name>Standard</Name>
              <IsBaseCalendar>1</IsBaseCalendar>
            </Calendar>
          </Calendars>
          <Tasks>
            <Task>
              <UID>1</UID>
              <ID>1</ID>
              <Name>Task A</Name>
              <OutlineLevel>1</OutlineLevel>
              <Summary>0</Summary>
              <Milestone>0</Milestone>
              <PercentComplete>25</PercentComplete>
              <Start>2026-05-06T08:00:00</Start>
              <Finish>2026-05-06T17:00:00</Finish>
              <Duration>PT8H0M0S</Duration>
            </Task>
            <Task>
              <UID>2</UID>
              <ID>2</ID>
              <Name>Task B</Name>
              <OutlineLevel>1</OutlineLevel>
              <Summary>0</Summary>
              <Milestone>0</Milestone>
              <PercentComplete>0</PercentComplete>
              <Start>2026-05-07T08:00:00</Start>
              <Finish>2026-05-07T17:00:00</Finish>
              <Duration>PT8H0M0S</Duration>
              <PredecessorLink>
                <PredecessorUID>1</PredecessorUID>
                <Type>FS</Type>
                <LinkLag>0</LinkLag>
              </PredecessorLink>
            </Task>
          </Tasks>
        </Project>
        "#;

        let mut file = tempfile::NamedTempFile::new().unwrap();
        file.write_all(xml.as_bytes()).unwrap();
        let document = load_project_document(file.path()).unwrap();

        assert_eq!(document.name, "Demo");
        assert_eq!(document.tasks.len(), 2);
        assert_eq!(document.dependencies.len(), 1);
        assert_eq!(document.tasks[1].predecessor_text, "FS 0");
        assert_eq!(
            document.chart_range().start,
            NaiveDate::from_ymd_opt(2026, 5, 6).unwrap()
        );
    }

    #[test]
    fn saves_and_reloads_project_xml() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: Some("Demo Title".to_string()),
            manager: Some("Manager".to_string()),
            start_date: parse_date_time(Some("2026-05-06T08:00:00")),
            finish_date: parse_date_time(Some("2026-05-08T17:00:00")),
            calendars: vec![CalendarInfo {
                name: "Standard".to_string(),
                base_calendar: true,
            }],
            tasks: vec![GanttTask {
                uid: 1,
                id: 1,
                name: "Task A".to_string(),
                outline_level: 1,
                summary: false,
                milestone: false,
                critical: false,
                percent_complete: 25.0,
                start_text: "2026-05-06T08:00:00".to_string(),
                finish_text: "2026-05-06T17:00:00".to_string(),
                start: parse_date_time(Some("2026-05-06T08:00:00")),
                finish: parse_date_time(Some("2026-05-06T17:00:00")),
                baseline_start: parse_date_time(Some("2026-05-05T08:00:00")),
                baseline_finish: parse_date_time(Some("2026-05-05T17:00:00")),
                duration_text: "PT8H0M0S".to_string(),
                predecessor_text: "FS".to_string(),
                notes_text: Some("Notes".to_string()),
                resource_names: Some("Dev".to_string()),
                calendar_uid: Some(1),
                constraint_type: Some("ASAP".to_string()),
            }],
            dependencies: vec![GanttDependency {
                predecessor_uid: 1,
                successor_uid: 1,
                relation: "FS".to_string(),
                lag_text: Some("0".to_string()),
            }],
        };

        let dir = tempdir().unwrap();
        let path = dir.path().join("roundtrip.xml");
        save_project_document(&path, &document).unwrap();

        let raw = fs::read_to_string(&path).unwrap();
        assert!(raw.contains("<Project"));
        assert!(raw.contains("Task A"));
        assert!(raw.contains("Demo Title"));

        let reloaded = load_project_document(&path).unwrap();
        assert_eq!(reloaded.name, "Demo");
        assert_eq!(reloaded.tasks.len(), 1);
        assert_eq!(reloaded.dependencies.len(), 1);
        assert_eq!(reloaded.tasks[0].name, "Task A");
        assert_eq!(reloaded.tasks[0].percent_complete, 25.0);
    }

    #[test]
    fn parses_real_world_project_xml_with_extra_elements() {
        let xml = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Project xmlns="http://schemas.microsoft.com/project">
            <SaveVersion>14</SaveVersion>
            <Name>test</Name>
            <Title>test</Title>
            <Manager>aaaa</Manager>
            <ScheduleFromStart>1</ScheduleFromStart>
            <StartDate>2026-05-06T08:00:00</StartDate>
            <FinishDate>2026-05-21T17:00:00</FinishDate>
            <Calendars>
                <Calendar>
                    <UID>1</UID>
                    <Name>標準</Name>
                    <IsBaseCalendar>1</IsBaseCalendar>
                </Calendar>
            </Calendars>
            <Tasks>
                <Task>
                    <UID>1</UID>
                    <ID>1</ID>
                    <Name>Task A</Name>
                    <Summary>0</Summary>
                    <Milestone>0</Milestone>
                    <PercentComplete>50</PercentComplete>
                    <Start>2026-05-06T08:00:00</Start>
                    <Finish>2026-05-06T17:00:00</Finish>
                    <Duration>PT8H0M0S</Duration>
                </Task>
            </Tasks>
        </Project>
        "#;

        let dir = tempdir().unwrap();
        let path = dir.path().join("real.xml");
        fs::write(&path, xml).unwrap();
        let document = load_project_document(&path).unwrap();

        assert_eq!(document.name, "test");
        assert_eq!(document.manager.as_deref(), Some("aaaa"));
        assert_eq!(document.calendars.len(), 1);
        assert_eq!(document.tasks.len(), 1);
        assert_eq!(document.tasks[0].name, "Task A");
    }
}

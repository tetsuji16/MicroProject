use chrono::{NaiveDate, NaiveDateTime};
use quick_xml::se::to_string;
use serde::Serialize;
use std::fs::read_to_string;
use std::fs::write;
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

#[derive(Debug, Serialize)]
#[serde(rename = "Project")]
struct ProjectXmlOut {
    #[serde(rename = "@xmlns")]
    xmlns: &'static str,
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Title", skip_serializing_if = "Option::is_none")]
    title: Option<String>,
    #[serde(rename = "Manager", skip_serializing_if = "Option::is_none")]
    manager: Option<String>,
    #[serde(rename = "StartDate", skip_serializing_if = "Option::is_none")]
    start_date: Option<String>,
    #[serde(rename = "FinishDate", skip_serializing_if = "Option::is_none")]
    finish_date: Option<String>,
    #[serde(rename = "Calendars", skip_serializing_if = "Option::is_none")]
    calendars: Option<CalendarsXmlOut>,
    #[serde(rename = "Tasks")]
    tasks: TasksXmlOut,
}

#[derive(Debug, Serialize)]
struct CalendarsXmlOut {
    #[serde(rename = "Calendar", default)]
    calendars: Vec<CalendarXmlOut>,
}

#[derive(Debug, Serialize)]
struct CalendarXmlOut {
    #[serde(rename = "Name", skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(rename = "IsBaseCalendar", skip_serializing_if = "Option::is_none")]
    base_calendar: Option<u8>,
}

#[derive(Debug, Serialize)]
struct TasksXmlOut {
    #[serde(rename = "Task", default)]
    tasks: Vec<TaskXmlOut>,
}

#[derive(Debug, Serialize)]
struct TaskXmlOut {
    #[serde(rename = "UID")]
    uid: u32,
    #[serde(rename = "ID")]
    id: u32,
    #[serde(rename = "Name", skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(rename = "OutlineLevel", skip_serializing_if = "Option::is_none")]
    outline_level: Option<u32>,
    #[serde(rename = "Summary", skip_serializing_if = "Option::is_none")]
    summary: Option<u8>,
    #[serde(rename = "Milestone", skip_serializing_if = "Option::is_none")]
    milestone: Option<u8>,
    #[serde(rename = "Critical", skip_serializing_if = "Option::is_none")]
    critical: Option<u8>,
    #[serde(rename = "PercentComplete", skip_serializing_if = "Option::is_none")]
    percent_complete: Option<f32>,
    #[serde(rename = "Start", skip_serializing_if = "Option::is_none")]
    start: Option<String>,
    #[serde(rename = "Finish", skip_serializing_if = "Option::is_none")]
    finish: Option<String>,
    #[serde(rename = "BaselineStart", skip_serializing_if = "Option::is_none")]
    baseline_start: Option<String>,
    #[serde(rename = "BaselineFinish", skip_serializing_if = "Option::is_none")]
    baseline_finish: Option<String>,
    #[serde(rename = "Duration", skip_serializing_if = "Option::is_none")]
    duration: Option<String>,
    #[serde(rename = "Notes", skip_serializing_if = "Option::is_none")]
    notes: Option<String>,
    #[serde(rename = "ResourceNames", skip_serializing_if = "Option::is_none")]
    resource_names: Option<String>,
    #[serde(rename = "CalendarUID", skip_serializing_if = "Option::is_none")]
    calendar_uid: Option<u32>,
    #[serde(rename = "ConstraintType", skip_serializing_if = "Option::is_none")]
    constraint_type: Option<String>,
    #[serde(rename = "PredecessorLink", default)]
    predecessor_links: Vec<PredecessorLinkXmlOut>,
}

#[derive(Debug, Serialize)]
struct PredecessorLinkXmlOut {
    #[serde(rename = "PredecessorUID")]
    predecessor_uid: u32,
    #[serde(rename = "Type")]
    relation: String,
    #[serde(rename = "LinkLag", skip_serializing_if = "Option::is_none")]
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
    let xml = xml.trim_start_matches(char::is_whitespace);
    let doc = roxmltree::Document::parse(xml)
        .map_err(|error| format!("failed to parse {}: {error}", path.display()))?;
    Ok(convert(&doc, path))
}

pub fn save_project_document(path: &Path, document: &ProjectDocument) -> Result<(), String> {
    let xml = ProjectXmlOut {
        xmlns: "http://schemas.microsoft.com/project",
        name: document.name.clone(),
        title: document.title.clone(),
        manager: document.manager.clone(),
        start_date: document.start_date.map(format_date_time),
        finish_date: document.finish_date.map(format_date_time),
        calendars: (!document.calendars.is_empty()).then(|| CalendarsXmlOut {
            calendars: document
                .calendars
                .iter()
                .map(|calendar| CalendarXmlOut {
                    name: Some(calendar.name.clone()),
                    base_calendar: Some(if calendar.base_calendar { 1 } else { 0 }),
                })
                .collect(),
        }),
        tasks: TasksXmlOut {
            tasks: document
                .tasks
                .iter()
                .map(|task| TaskXmlOut {
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
                        .map(|dependency| PredecessorLinkXmlOut {
                            predecessor_uid: dependency.predecessor_uid,
                            relation: dependency.relation.clone(),
                            lag: dependency.lag_text.clone(),
                        })
                        .collect(),
                })
                .collect(),
        },
    };

    let mut rendered = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    rendered.push_str(
        &to_string(&xml)
            .map_err(|error| format!("failed to serialize {}: {error}", path.display()))?,
    );
    rendered.push('\n');

    write(path, rendered).map_err(|error| format!("failed to write {}: {error}", path.display()))
}

fn convert(doc: &roxmltree::Document<'_>, path: &Path) -> ProjectDocument {
    let root = doc
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "Project")
        .unwrap_or_else(|| doc.root_element());

    let fallback_name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled Project")
        .to_string();
    let name = text_child(root, "Name").unwrap_or_else(|| fallback_name.clone());
    let title = text_child(root, "Title").or_else(|| Some(name.clone()));
    let manager = text_child(root, "Manager");
    let start_date = parse_date_time(text_child(root, "StartDate").as_deref());
    let finish_date = parse_date_time(text_child(root, "FinishDate").as_deref());
    let calendars = root
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "Calendars")
        .map(|wrapper| {
            wrapper
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "Calendar")
                .map(|calendar| CalendarInfo {
                    name: text_child(calendar, "Name").unwrap_or_else(|| "Calendar".to_string()),
                    base_calendar: text_child(calendar, "IsBaseCalendar")
                        .and_then(|text| text.parse::<u8>().ok())
                        .unwrap_or(0)
                        != 0,
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();

    let mut tasks = Vec::new();
    let mut dependencies = Vec::new();

    if let Some(task_wrapper) = root
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "Tasks")
    {
        for task in task_wrapper
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "Task")
        {
            let uid = text_child(task, "UID")
                .and_then(|text| text.parse::<u32>().ok())
                .unwrap_or(0);
            let id = text_child(task, "ID")
                .and_then(|text| text.parse::<u32>().ok())
                .unwrap_or(uid);
            let start_text = text_child(task, "Start").unwrap_or_default();
            let finish_text = text_child(task, "Finish").unwrap_or_default();
            let baseline_start_text = text_child(task, "BaselineStart").unwrap_or_default();
            let baseline_finish_text = text_child(task, "BaselineFinish").unwrap_or_default();
            let start = parse_date_time(Some(start_text.as_str()));
            let finish = parse_date_time(Some(finish_text.as_str()));
            let baseline_start = parse_date_time(Some(baseline_start_text.as_str()));
            let baseline_finish = parse_date_time(Some(baseline_finish_text.as_str()));
            let predecessor_links = task
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "PredecessorLink")
                .map(|link| PredecessorLinkData {
                    predecessor_uid: text_child(link, "PredecessorUID")
                        .and_then(|text| text.parse::<u32>().ok()),
                    relation: text_child(link, "Type"),
                    lag: text_child(link, "LinkLag"),
                })
                .collect::<Vec<_>>();
            let predecessor_text = predecessor_links
                .iter()
                .map(format_predecessor_link_data)
                .collect::<Vec<_>>()
                .join(", ");

            for link in predecessor_links {
                if let Some(predecessor_uid) = link.predecessor_uid {
                    dependencies.push(GanttDependency {
                        predecessor_uid,
                        successor_uid: uid,
                        relation: link.relation.unwrap_or_else(|| "FS".to_string()),
                        lag_text: link.lag.clone(),
                    });
                }
            }

            tasks.push(GanttTask {
                uid,
                id,
                name: text_child(task, "Name").unwrap_or_default(),
                outline_level: text_child(task, "OutlineLevel")
                    .and_then(|text| text.parse::<u32>().ok())
                    .unwrap_or(1),
                summary: text_child(task, "Summary")
                    .and_then(|text| text.parse::<u8>().ok())
                    .unwrap_or(0)
                    != 0,
                milestone: text_child(task, "Milestone")
                    .and_then(|text| text.parse::<u8>().ok())
                    .unwrap_or(0)
                    != 0,
                critical: text_child(task, "Critical")
                    .and_then(|text| text.parse::<u8>().ok())
                    .unwrap_or(0)
                    != 0,
                percent_complete: text_child(task, "PercentComplete")
                    .and_then(|text| text.parse::<f32>().ok())
                    .unwrap_or(0.0),
                start_text,
                finish_text,
                start,
                finish,
                baseline_start,
                baseline_finish,
                duration_text: text_child(task, "Duration").unwrap_or_default(),
                predecessor_text,
                notes_text: text_child(task, "Notes"),
                resource_names: text_child(task, "ResourceNames"),
                calendar_uid: text_child(task, "CalendarUID")
                    .and_then(|text| text.parse::<u32>().ok()),
                constraint_type: text_child(task, "ConstraintType"),
            });
        }
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

#[derive(Debug, Clone)]
struct PredecessorLinkData {
    predecessor_uid: Option<u32>,
    relation: Option<String>,
    lag: Option<String>,
}

fn text_child(node: roxmltree::Node<'_, '_>, name: &str) -> Option<String> {
    node.children()
        .find(|child| child.is_element() && child.tag_name().name() == name)
        .and_then(|child| child.text())
        .map(|text| text.trim().to_string())
        .filter(|text| !text.is_empty())
}

fn format_predecessor_link_data(link: &PredecessorLinkData) -> String {
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
            let start = min.pred_opt().unwrap_or(min);
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
            NaiveDate::from_ymd_opt(2026, 5, 5).unwrap()
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

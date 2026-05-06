use chrono::{NaiveDate, NaiveDateTime};
use quick_xml::de::from_reader;
use quick_xml::se::to_string;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::fs::write;
use std::io::BufReader;
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
    #[serde(rename = "PercentComplete", skip_serializing_if = "Option::is_none")]
    percent_complete: Option<f32>,
    #[serde(rename = "Start", skip_serializing_if = "Option::is_none")]
    start: Option<String>,
    #[serde(rename = "Finish", skip_serializing_if = "Option::is_none")]
    finish: Option<String>,
    #[serde(rename = "Duration", skip_serializing_if = "Option::is_none")]
    duration: Option<String>,
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
    pub percent_complete: f32,
    pub start: Option<NaiveDateTime>,
    pub finish: Option<NaiveDateTime>,
    pub duration_text: Option<String>,
    pub predecessor_text: String,
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

#[derive(Debug, Deserialize)]
#[serde(rename = "Project")]
struct ProjectXml {
    #[serde(rename = "Name")]
    name: Option<String>,
    #[serde(rename = "Title")]
    title: Option<String>,
    #[serde(rename = "Manager")]
    manager: Option<String>,
    #[serde(rename = "StartDate")]
    start_date: Option<String>,
    #[serde(rename = "FinishDate")]
    finish_date: Option<String>,
    #[serde(rename = "Calendars")]
    calendars: Option<CalendarsXml>,
    #[serde(rename = "Tasks")]
    tasks: Option<TasksXml>,
}

#[derive(Debug, Deserialize, Default)]
struct CalendarsXml {
    #[serde(rename = "Calendar", default)]
    calendars: Vec<CalendarXml>,
}

#[derive(Debug, Deserialize, Default)]
struct CalendarXml {
    #[serde(rename = "Name")]
    name: Option<String>,
    #[serde(rename = "IsBaseCalendar")]
    base_calendar: Option<u8>,
}

#[derive(Debug, Deserialize, Default)]
struct TasksXml {
    #[serde(rename = "Task", default)]
    tasks: Vec<TaskXml>,
}

#[derive(Debug, Deserialize, Default)]
struct TaskXml {
    #[serde(rename = "UID")]
    uid: Option<u32>,
    #[serde(rename = "ID")]
    id: Option<u32>,
    #[serde(rename = "Name")]
    name: Option<String>,
    #[serde(rename = "OutlineLevel")]
    outline_level: Option<u32>,
    #[serde(rename = "Summary")]
    summary: Option<u8>,
    #[serde(rename = "Milestone")]
    milestone: Option<u8>,
    #[serde(rename = "PercentComplete")]
    percent_complete: Option<f32>,
    #[serde(rename = "Start")]
    start: Option<String>,
    #[serde(rename = "Finish")]
    finish: Option<String>,
    #[serde(rename = "Duration")]
    duration: Option<String>,
    #[serde(rename = "PredecessorLink", default)]
    predecessor_links: Vec<PredecessorLinkXml>,
}

#[derive(Debug, Deserialize, Default)]
struct PredecessorLinkXml {
    #[serde(rename = "PredecessorUID")]
    predecessor_uid: Option<u32>,
    #[serde(rename = "Type")]
    relation: Option<String>,
    #[serde(rename = "LinkLag")]
    lag: Option<String>,
}

pub fn load_project_document(path: &Path) -> Result<ProjectDocument, String> {
    let file = File::open(path)
        .map_err(|error| format!("failed to open {}: {error}", path.display()))?;
    let reader = BufReader::new(file);
    let xml: ProjectXml = from_reader(reader)
        .map_err(|error| format!("failed to parse {}: {error}", path.display()))?;
    Ok(convert(xml, path))
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
                    percent_complete: Some(task.percent_complete),
                    start: task.start.map(format_date_time),
                    finish: task.finish.map(format_date_time),
                    duration: task.duration_text.clone(),
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
        &to_string(&xml).map_err(|error| format!("failed to serialize {}: {error}", path.display()))?,
    );
    rendered.push('\n');

    write(path, rendered).map_err(|error| format!("failed to write {}: {error}", path.display()))
}

fn convert(xml: ProjectXml, path: &Path) -> ProjectDocument {
    let fallback_name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled Project")
        .to_string();
    let name = xml.name.clone().unwrap_or_else(|| fallback_name.clone());
    let title = xml.title.clone().or_else(|| Some(name.clone()));
    let calendars = xml
        .calendars
        .map(|wrapper| {
            wrapper
                .calendars
                .into_iter()
                .map(|calendar| CalendarInfo {
                    name: calendar.name.unwrap_or_else(|| "Calendar".to_string()),
                    base_calendar: calendar.base_calendar.unwrap_or(0) != 0,
                })
                .collect()
        })
        .unwrap_or_default();

    let mut tasks = Vec::new();
    let mut dependencies = Vec::new();

    if let Some(task_wrapper) = xml.tasks {
        for task in task_wrapper.tasks {
            let uid = task.uid.unwrap_or(0);
            let id = task.id.unwrap_or(uid);
            let start = parse_date_time(task.start.as_deref());
            let finish = parse_date_time(task.finish.as_deref());
            let predecessor_text = task
                .predecessor_links
                .iter()
                .map(format_predecessor_link)
                .collect::<Vec<_>>()
                .join(", ");

            for link in task.predecessor_links {
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
                name: task.name.unwrap_or_default(),
                outline_level: task.outline_level.unwrap_or(1),
                summary: task.summary.unwrap_or(0) != 0,
                milestone: task.milestone.unwrap_or(0) != 0,
                percent_complete: task.percent_complete.unwrap_or(0.0),
                start,
                finish,
                duration_text: task.duration,
                predecessor_text,
            });
        }
    }

    ProjectDocument {
        name,
        title,
        manager: xml.manager,
        start_date: parse_date_time(xml.start_date.as_deref()),
        finish_date: parse_date_time(xml.finish_date.as_deref()),
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

fn format_predecessor_link(link: &PredecessorLinkXml) -> String {
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
                percent_complete: 25.0,
                start: parse_date_time(Some("2026-05-06T08:00:00")),
                finish: parse_date_time(Some("2026-05-06T17:00:00")),
                duration_text: Some("PT8H0M0S".to_string()),
                predecessor_text: "FS".to_string(),
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
}

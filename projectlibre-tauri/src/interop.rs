use crate::models::*;
use roxmltree::{Document, Node};

pub fn export_workspace_xml(snapshot: &WorkspaceSnapshot) -> Result<String, String> {
    let mut xml = String::new();
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str(&format!(
        "<microproject schema_version=\"{}\">\n",
        snapshot.schema_version
    ));
    write_settings(&mut xml, &snapshot.settings);
    write_projects(&mut xml, &snapshot.projects);
    write_tasks(&mut xml, &snapshot.tasks);
    write_dependencies(&mut xml, &snapshot.dependencies);
    write_resources(&mut xml, &snapshot.resources);
    write_assignments(&mut xml, &snapshot.assignments);
    write_calendars(&mut xml, &snapshot.calendars);
    write_baselines(&mut xml, &snapshot.baselines);
    xml.push_str("</microproject>\n");
    Ok(xml)
}

pub fn import_workspace_xml(xml: &str) -> Result<WorkspaceSnapshot, String> {
    let doc = Document::parse(xml).map_err(|error| format!("failed to parse XML snapshot: {error}"))?;
    let root = doc.root_element();
    if root.tag_name().name() != "microproject" {
        return Err("unexpected XML root element; expected <microproject>".to_string());
    }

    let mut snapshot = WorkspaceSnapshot::new();
    if let Some(schema_version) = attr_u32(root, "schema_version")? {
        snapshot.schema_version = schema_version;
    }

    for child in root.children().filter(|node| node.is_element()) {
        match child.tag_name().name() {
            "settings" => snapshot.settings = parse_settings(child)?,
            "projects" => snapshot.projects = parse_projects(child)?,
            "tasks" => snapshot.tasks = parse_tasks(child)?,
            "dependencies" => snapshot.dependencies = parse_dependencies(child)?,
            "resources" => snapshot.resources = parse_resources(child)?,
            "assignments" => snapshot.assignments = parse_assignments(child)?,
            "calendars" => snapshot.calendars = parse_calendars(child)?,
            "baselines" => snapshot.baselines = parse_baselines(child)?,
            _ => {}
        }
    }

    Ok(snapshot)
}

fn write_settings(xml: &mut String, settings: &WorkspaceSettings) {
    xml.push_str("  <settings");
    write_opt_attr(xml, "default_calendar_id", settings.default_calendar_id.as_deref());
    write_attr(xml, "units_per_day", settings.units_per_day.to_string());
    write_opt_attr(xml, "locale", settings.locale.as_deref());
    xml.push_str(" />\n");
}

fn write_projects(xml: &mut String, projects: &[ProjectRecord]) {
    xml.push_str("  <projects>\n");
    for project in projects {
        xml.push_str("    <project");
        write_attr(xml, "id", &project.id);
        write_attr(xml, "name", &project.name);
        write_opt_attr(xml, "description", project.description.as_deref());
        write_opt_attr(xml, "manager", project.manager.as_deref());
        write_attr(xml, "status", &project.status);
        write_attr(xml, "priority", project.priority.to_string());
        write_opt_attr(xml, "calendar_id", project.calendar_id.as_deref());
        write_opt_attr(xml, "start_date", project.start_date.as_deref());
        write_opt_attr(xml, "finish_date", project.finish_date.as_deref());
        write_opt_attr(xml, "calculated_start_date", project.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", project.calculated_finish_date.as_deref());
        write_attr(xml, "notes", &project.notes);
        write_attr(xml, "created_at", &project.created_at);
        write_attr(xml, "updated_at", &project.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </projects>\n");
}

fn write_tasks(xml: &mut String, tasks: &[TaskRecord]) {
    xml.push_str("  <tasks>\n");
    for task in tasks {
        xml.push_str("    <task");
        write_attr(xml, "id", &task.id);
        write_attr(xml, "project_id", &task.project_id);
        write_opt_attr(xml, "parent_task_id", task.parent_task_id.as_deref());
        write_attr(xml, "name", &task.name);
        write_opt_attr(xml, "description", task.description.as_deref());
        write_attr(xml, "outline_level", task.outline_level.to_string());
        write_attr(xml, "wbs", &task.wbs);
        write_opt_attr(xml, "start_date", task.start_date.as_deref());
        write_opt_attr(xml, "finish_date", task.finish_date.as_deref());
        write_opt_attr(xml, "calculated_start_date", task.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", task.calculated_finish_date.as_deref());
        write_opt_attr(xml, "duration_hours", task.duration_hours.map(|value| value.to_string()).as_deref());
        write_opt_attr(xml, "work_hours", task.work_hours.map(|value| value.to_string()).as_deref());
        write_attr(xml, "percent_complete", task.percent_complete.to_string());
        write_attr(xml, "milestone", task.milestone.to_string());
        write_attr(xml, "constraint_type", &task.constraint_type);
        write_opt_attr(xml, "calendar_id", task.calendar_id.as_deref());
        write_attr(xml, "notes", &task.notes);
        write_attr(xml, "sort_order", task.sort_order.to_string());
        write_attr(xml, "created_at", &task.created_at);
        write_attr(xml, "updated_at", &task.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </tasks>\n");
}

fn write_dependencies(xml: &mut String, dependencies: &[DependencyRecord]) {
    xml.push_str("  <dependencies>\n");
    for dependency in dependencies {
        xml.push_str("    <dependency");
        write_attr(xml, "id", &dependency.id);
        write_attr(xml, "project_id", &dependency.project_id);
        write_attr(xml, "predecessor_task_id", &dependency.predecessor_task_id);
        write_attr(xml, "successor_task_id", &dependency.successor_task_id);
        write_attr(xml, "relation", &dependency.relation);
        write_attr(xml, "lag_hours", dependency.lag_hours.to_string());
        write_attr(xml, "created_at", &dependency.created_at);
        write_attr(xml, "updated_at", &dependency.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </dependencies>\n");
}

fn write_resources(xml: &mut String, resources: &[ResourceRecord]) {
    xml.push_str("  <resources>\n");
    for resource in resources {
        xml.push_str("    <resource");
        write_attr(xml, "id", &resource.id);
        write_attr(xml, "project_id", &resource.project_id);
        write_attr(xml, "name", &resource.name);
        write_attr(xml, "resource_type", &resource.resource_type);
        write_attr(xml, "max_units", resource.max_units.to_string());
        write_attr(xml, "standard_rate", resource.standard_rate.to_string());
        write_attr(xml, "overtime_rate", resource.overtime_rate.to_string());
        write_attr(xml, "cost_per_use", resource.cost_per_use.to_string());
        write_opt_attr(xml, "calendar_id", resource.calendar_id.as_deref());
        write_attr(xml, "notes", &resource.notes);
        write_attr(xml, "created_at", &resource.created_at);
        write_attr(xml, "updated_at", &resource.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </resources>\n");
}

fn write_assignments(xml: &mut String, assignments: &[AssignmentRecord]) {
    xml.push_str("  <assignments>\n");
    for assignment in assignments {
        xml.push_str("    <assignment");
        write_attr(xml, "id", &assignment.id);
        write_attr(xml, "project_id", &assignment.project_id);
        write_attr(xml, "task_id", &assignment.task_id);
        write_attr(xml, "resource_id", &assignment.resource_id);
        write_attr(xml, "units", assignment.units.to_string());
        write_attr(xml, "work_hours", assignment.work_hours.to_string());
        write_attr(xml, "actual_work_hours", assignment.actual_work_hours.to_string());
        write_attr(xml, "cost", assignment.cost.to_string());
        write_attr(xml, "created_at", &assignment.created_at);
        write_attr(xml, "updated_at", &assignment.updated_at);
        xml.push_str(" />\n");
    }
    xml.push_str("  </assignments>\n");
}

fn write_calendars(xml: &mut String, calendars: &[CalendarRecord]) {
    xml.push_str("  <calendars>\n");
    for calendar in calendars {
        xml.push_str("    <calendar");
        write_attr(xml, "id", &calendar.id);
        write_attr(xml, "project_id", &calendar.project_id);
        write_attr(xml, "name", &calendar.name);
        write_opt_attr(xml, "timezone", calendar.timezone.as_deref());
        write_attr(xml, "hours_per_day", calendar.hours_per_day.to_string());
        write_attr(xml, "created_at", &calendar.created_at);
        write_attr(xml, "updated_at", &calendar.updated_at);
        xml.push_str(">\n");
        xml.push_str("      <working_days>\n");
        for day in &calendar.working_days {
            xml.push_str(&format!("        <day value=\"{}\" />\n", day));
        }
        xml.push_str("      </working_days>\n");
        xml.push_str("      <working_hours>\n");
        for interval in &calendar.working_hours {
            xml.push_str("        <interval");
            write_attr(xml, "start", &interval.start);
            write_attr(xml, "finish", &interval.finish);
            xml.push_str(" />\n");
        }
        xml.push_str("      </working_hours>\n");
        xml.push_str("      <exceptions>\n");
        for exception in &calendar.exceptions {
            xml.push_str("        <exception");
            write_attr(xml, "date", &exception.date);
            write_attr(xml, "is_working", exception.is_working.to_string());
            xml.push_str(">\n");
            for interval in &exception.intervals {
                xml.push_str("          <interval");
                write_attr(xml, "start", &interval.start);
                write_attr(xml, "finish", &interval.finish);
                xml.push_str(" />\n");
            }
            xml.push_str("        </exception>\n");
        }
        xml.push_str("      </exceptions>\n");
        xml.push_str("    </calendar>\n");
    }
    xml.push_str("  </calendars>\n");
}

fn write_baselines(xml: &mut String, baselines: &[BaselineRecord]) {
    xml.push_str("  <baselines>\n");
    for baseline in baselines {
        xml.push_str("    <baseline");
        write_attr(xml, "id", &baseline.id);
        write_attr(xml, "project_id", &baseline.project_id);
        write_attr(xml, "name", &baseline.name);
        write_attr(xml, "captured_at", &baseline.captured_at);
        xml.push_str(">\n");
        xml.push_str("      <project");
        write_attr(xml, "id", &baseline.project.id);
        write_attr(xml, "name", &baseline.project.name);
        write_opt_attr(xml, "calculated_start_date", baseline.project.calculated_start_date.as_deref());
        write_opt_attr(xml, "calculated_finish_date", baseline.project.calculated_finish_date.as_deref());
        xml.push_str(" />\n");
        xml.push_str("      <tasks>\n");
        for task in &baseline.tasks {
            xml.push_str("        <task");
            write_attr(xml, "id", &task.id);
            write_attr(xml, "name", &task.name);
            write_opt_attr(xml, "start_date", task.start_date.as_deref());
            write_opt_attr(xml, "finish_date", task.finish_date.as_deref());
            write_attr(xml, "percent_complete", task.percent_complete.to_string());
            xml.push_str(" />\n");
        }
        xml.push_str("      </tasks>\n");
        xml.push_str("      <resources>\n");
        for resource in &baseline.resources {
            xml.push_str("        <resource");
            write_attr(xml, "id", &resource.id);
            write_attr(xml, "name", &resource.name);
            write_attr(xml, "max_units", resource.max_units.to_string());
            xml.push_str(" />\n");
        }
        xml.push_str("      </resources>\n");
        xml.push_str("      <assignments>\n");
        for assignment in &baseline.assignments {
            xml.push_str("        <assignment");
            write_attr(xml, "id", &assignment.id);
            write_attr(xml, "task_id", &assignment.task_id);
            write_attr(xml, "resource_id", &assignment.resource_id);
            write_attr(xml, "units", assignment.units.to_string());
            xml.push_str(" />\n");
        }
        xml.push_str("      </assignments>\n");
        xml.push_str("    </baseline>\n");
    }
    xml.push_str("  </baselines>\n");
}

fn parse_settings(node: Node) -> Result<WorkspaceSettings, String> {
    Ok(WorkspaceSettings {
        default_calendar_id: attr_string(node, "default_calendar_id"),
        units_per_day: attr_f32(node, "units_per_day")?.unwrap_or(8.0),
        locale: attr_string(node, "locale"),
    })
}

fn parse_projects(node: Node) -> Result<Vec<ProjectRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "project")
        .map(|child| {
            Ok(ProjectRecord {
                id: attr_required_string(child, "id")?,
                name: attr_required_string(child, "name")?,
                description: attr_string(child, "description"),
                manager: attr_string(child, "manager"),
                status: attr_string(child, "status").unwrap_or_else(|| "planning".to_string()),
                priority: attr_i32(child, "priority")?.unwrap_or(500),
                calendar_id: attr_string(child, "calendar_id"),
                start_date: attr_string(child, "start_date"),
                finish_date: attr_string(child, "finish_date"),
                calculated_start_date: attr_string(child, "calculated_start_date"),
                calculated_finish_date: attr_string(child, "calculated_finish_date"),
                notes: attr_string(child, "notes").unwrap_or_default(),
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_tasks(node: Node) -> Result<Vec<TaskRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "task")
        .map(|child| {
            Ok(TaskRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                parent_task_id: attr_string(child, "parent_task_id"),
                name: attr_required_string(child, "name")?,
                description: attr_string(child, "description"),
                outline_level: attr_u32(child, "outline_level")?.unwrap_or_default(),
                wbs: attr_string(child, "wbs").unwrap_or_default(),
                start_date: attr_string(child, "start_date"),
                finish_date: attr_string(child, "finish_date"),
                calculated_start_date: attr_string(child, "calculated_start_date"),
                calculated_finish_date: attr_string(child, "calculated_finish_date"),
                duration_hours: attr_f64(child, "duration_hours")?,
                work_hours: attr_f64(child, "work_hours")?,
                percent_complete: attr_f32(child, "percent_complete")?.unwrap_or_default(),
                milestone: attr_bool(child, "milestone")?.unwrap_or(false),
                constraint_type: attr_string(child, "constraint_type").unwrap_or_else(|| "ASAP".to_string()),
                calendar_id: attr_string(child, "calendar_id"),
                notes: attr_string(child, "notes").unwrap_or_default(),
                sort_order: attr_i64(child, "sort_order")?.unwrap_or_default(),
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_dependencies(node: Node) -> Result<Vec<DependencyRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "dependency")
        .map(|child| {
            Ok(DependencyRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                predecessor_task_id: attr_required_string(child, "predecessor_task_id")?,
                successor_task_id: attr_required_string(child, "successor_task_id")?,
                relation: attr_string(child, "relation").unwrap_or_else(|| "FS".to_string()),
                lag_hours: attr_f64(child, "lag_hours")?.unwrap_or_default(),
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_resources(node: Node) -> Result<Vec<ResourceRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "resource")
        .map(|child| {
            Ok(ResourceRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                name: attr_required_string(child, "name")?,
                resource_type: attr_string(child, "resource_type").unwrap_or_else(|| "work".to_string()),
                max_units: attr_f32(child, "max_units")?.unwrap_or(100.0),
                standard_rate: attr_f64(child, "standard_rate")?.unwrap_or_default(),
                overtime_rate: attr_f64(child, "overtime_rate")?.unwrap_or_default(),
                cost_per_use: attr_f64(child, "cost_per_use")?.unwrap_or_default(),
                calendar_id: attr_string(child, "calendar_id"),
                notes: attr_string(child, "notes").unwrap_or_default(),
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_assignments(node: Node) -> Result<Vec<AssignmentRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "assignment")
        .map(|child| {
            Ok(AssignmentRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                task_id: attr_required_string(child, "task_id")?,
                resource_id: attr_required_string(child, "resource_id")?,
                units: attr_f32(child, "units")?.unwrap_or(100.0),
                work_hours: attr_f64(child, "work_hours")?.unwrap_or_default(),
                actual_work_hours: attr_f64(child, "actual_work_hours")?.unwrap_or_default(),
                cost: attr_f64(child, "cost")?.unwrap_or_default(),
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_calendars(node: Node) -> Result<Vec<CalendarRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "calendar")
        .map(|child| {
            let working_days = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "working_days")
                .map(parse_working_days)
                .transpose()?
                .unwrap_or_else(|| vec![1, 2, 3, 4, 5]);
            let working_hours = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "working_hours")
                .map(parse_working_hours)
                .transpose()?
                .unwrap_or_else(default_working_hours);
            let exceptions = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "exceptions")
                .map(parse_exceptions)
                .transpose()?
                .unwrap_or_default();

            Ok(CalendarRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                name: attr_required_string(child, "name")?,
                timezone: attr_string(child, "timezone"),
                working_days,
                hours_per_day: attr_f32(child, "hours_per_day")?.unwrap_or(8.0),
                working_hours,
                exceptions,
                created_at: attr_string(child, "created_at").unwrap_or_default(),
                updated_at: attr_string(child, "updated_at").unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_baselines(node: Node) -> Result<Vec<BaselineRecord>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "baseline")
        .map(|child| {
            let project = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "project")
                .ok_or_else(|| "baseline is missing a project snapshot".to_string())?;
            let tasks = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "tasks")
                .map(parse_baseline_tasks)
                .transpose()?
                .unwrap_or_default();
            let resources = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "resources")
                .map(parse_baseline_resources)
                .transpose()?
                .unwrap_or_default();
            let assignments = child
                .children()
                .find(|section| section.is_element() && section.tag_name().name() == "assignments")
                .map(parse_baseline_assignments)
                .transpose()?
                .unwrap_or_default();

            Ok(BaselineRecord {
                id: attr_required_string(child, "id")?,
                project_id: attr_required_string(child, "project_id")?,
                name: attr_required_string(child, "name")?,
                captured_at: attr_string(child, "captured_at").unwrap_or_default(),
                project: ProjectSnapshot {
                    id: attr_required_string(project, "id")?,
                    name: attr_required_string(project, "name")?,
                    calculated_start_date: attr_string(project, "calculated_start_date"),
                    calculated_finish_date: attr_string(project, "calculated_finish_date"),
                },
                tasks,
                resources,
                assignments,
            })
        })
        .collect()
}

fn parse_working_days(node: Node) -> Result<Vec<u8>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "day")
        .map(|child| {
            attr_u8(child, "value")?.ok_or_else(|| "missing working day value".to_string())
        })
        .collect()
}

fn parse_working_hours(node: Node) -> Result<Vec<WorkInterval>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "interval")
        .map(|child| {
            Ok(WorkInterval {
                start: attr_required_string(child, "start")?,
                finish: attr_required_string(child, "finish")?,
            })
        })
        .collect()
}

fn parse_exceptions(node: Node) -> Result<Vec<CalendarException>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "exception")
        .map(|child| {
            let intervals = child
                .children()
                .filter(|grandchild| grandchild.is_element() && grandchild.tag_name().name() == "interval")
                .map(|grandchild| {
                    Ok(WorkInterval {
                        start: attr_required_string(grandchild, "start")?,
                        finish: attr_required_string(grandchild, "finish")?,
                    })
                })
                .collect::<Result<Vec<_>, String>>()?;

            Ok(CalendarException {
                date: attr_required_string(child, "date")?,
                is_working: attr_bool(child, "is_working")?.unwrap_or(false),
                intervals,
            })
        })
        .collect()
}

fn parse_baseline_tasks(node: Node) -> Result<Vec<TaskSnapshot>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "task")
        .map(|child| {
            Ok(TaskSnapshot {
                id: attr_required_string(child, "id")?,
                name: attr_required_string(child, "name")?,
                start_date: attr_string(child, "start_date"),
                finish_date: attr_string(child, "finish_date"),
                percent_complete: attr_f32(child, "percent_complete")?.unwrap_or_default(),
            })
        })
        .collect()
}

fn parse_baseline_resources(node: Node) -> Result<Vec<ResourceSnapshot>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "resource")
        .map(|child| {
            Ok(ResourceSnapshot {
                id: attr_required_string(child, "id")?,
                name: attr_required_string(child, "name")?,
                max_units: attr_f32(child, "max_units")?.unwrap_or(100.0),
            })
        })
        .collect()
}

fn parse_baseline_assignments(node: Node) -> Result<Vec<AssignmentSnapshot>, String> {
    node.children()
        .filter(|child| child.is_element() && child.tag_name().name() == "assignment")
        .map(|child| {
            Ok(AssignmentSnapshot {
                id: attr_required_string(child, "id")?,
                task_id: attr_required_string(child, "task_id")?,
                resource_id: attr_required_string(child, "resource_id")?,
                units: attr_f32(child, "units")?.unwrap_or(100.0),
            })
        })
        .collect()
}

fn default_working_hours() -> Vec<WorkInterval> {
    vec![WorkInterval {
        start: "09:00".to_string(),
        finish: "17:00".to_string(),
    }]
}

fn write_attr(xml: &mut String, name: &str, value: impl AsRef<str>) {
    xml.push_str(&format!(" {}=\"{}\"", name, escape_attr(value.as_ref())));
}

fn write_opt_attr(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        write_attr(xml, name, value);
    }
}

fn escape_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn attr_string(node: Node, name: &str) -> Option<String> {
    node.attribute(name).map(ToString::to_string)
}

fn attr_required_string(node: Node, name: &str) -> Result<String, String> {
    attr_string(node, name).ok_or_else(|| format!("missing required XML attribute: {name}"))
}

fn attr_bool(node: Node, name: &str) -> Result<Option<bool>, String> {
    attr_optional(node, name, |value| match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        other => Err(format!("invalid boolean value for {name}: {other}")),
    })
}

fn attr_u8(node: Node, name: &str) -> Result<Option<u8>, String> {
    attr_optional(node, name, |value| value.parse::<u8>().map_err(|error| format!("invalid u8 value for {name}: {error}")))
}

fn attr_u32(node: Node, name: &str) -> Result<Option<u32>, String> {
    attr_optional(node, name, |value| value.parse::<u32>().map_err(|error| format!("invalid u32 value for {name}: {error}")))
}

fn attr_i32(node: Node, name: &str) -> Result<Option<i32>, String> {
    attr_optional(node, name, |value| value.parse::<i32>().map_err(|error| format!("invalid i32 value for {name}: {error}")))
}

fn attr_i64(node: Node, name: &str) -> Result<Option<i64>, String> {
    attr_optional(node, name, |value| value.parse::<i64>().map_err(|error| format!("invalid i64 value for {name}: {error}")))
}

fn attr_f32(node: Node, name: &str) -> Result<Option<f32>, String> {
    attr_optional(node, name, |value| value.parse::<f32>().map_err(|error| format!("invalid f32 value for {name}: {error}")))
}

fn attr_f64(node: Node, name: &str) -> Result<Option<f64>, String> {
    attr_optional(node, name, |value| value.parse::<f64>().map_err(|error| format!("invalid f64 value for {name}: {error}")))
}

fn attr_optional<T>(
    node: Node,
    name: &str,
    parser: impl FnOnce(&str) -> Result<T, String>,
) -> Result<Option<T>, String> {
    match node.attribute(name) {
        Some(value) => parser(value).map(Some),
        None => Ok(None),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_snapshot() -> WorkspaceSnapshot {
        let mut snapshot = WorkspaceSnapshot::new();
        snapshot.settings = WorkspaceSettings {
            default_calendar_id: Some("cal_1".to_string()),
            units_per_day: 7.5,
            locale: Some("ja-JP".to_string()),
        };
        snapshot.projects.push(ProjectRecord {
            id: "project_1".to_string(),
            name: "Alpha".to_string(),
            status: "planning".to_string(),
            start_date: Some("2026-05-01".to_string()),
            notes: "hello".to_string(),
            created_at: "1".to_string(),
            updated_at: "2".to_string(),
            ..ProjectRecord::default()
        });
        snapshot.tasks.push(TaskRecord {
            id: "task_1".to_string(),
            project_id: "project_1".to_string(),
            name: "Task".to_string(),
            duration_hours: Some(8.0),
            percent_complete: 10.0,
            sort_order: 1,
            created_at: "3".to_string(),
            updated_at: "4".to_string(),
            ..TaskRecord::default()
        });
        snapshot.tasks.push(TaskRecord {
            id: "task_1_s".to_string(),
            project_id: "project_1".to_string(),
            name: "Task S".to_string(),
            duration_hours: Some(8.0),
            percent_complete: 0.0,
            sort_order: 2,
            created_at: "3b".to_string(),
            updated_at: "4b".to_string(),
            ..TaskRecord::default()
        });
        snapshot.dependencies.push(DependencyRecord {
            id: "dep_1".to_string(),
            project_id: "project_1".to_string(),
            predecessor_task_id: "task_1".to_string(),
            successor_task_id: "task_1_s".to_string(),
            relation: "FS".to_string(),
            created_at: "5".to_string(),
            updated_at: "6".to_string(),
            ..DependencyRecord::default()
        });
        snapshot.resources.push(ResourceRecord {
            id: "res_1".to_string(),
            project_id: "project_1".to_string(),
            name: "Alice".to_string(),
            standard_rate: 42.0,
            created_at: "7".to_string(),
            updated_at: "8".to_string(),
            ..ResourceRecord::default()
        });
        snapshot.assignments.push(AssignmentRecord {
            id: "asg_1".to_string(),
            project_id: "project_1".to_string(),
            task_id: "task_1".to_string(),
            resource_id: "res_1".to_string(),
            units: 75.0,
            cost: 31.5,
            created_at: "9".to_string(),
            updated_at: "10".to_string(),
            ..AssignmentRecord::default()
        });
        snapshot.calendars.push(CalendarRecord {
            id: "cal_1".to_string(),
            project_id: "project_1".to_string(),
            name: "Default".to_string(),
            timezone: None,
            working_days: vec![1, 2, 3, 4, 5],
            hours_per_day: 8.0,
            working_hours: vec![WorkInterval {
                start: "09:00".to_string(),
                finish: "17:00".to_string(),
            }],
            exceptions: vec![CalendarException {
                date: "2026-05-02".to_string(),
                is_working: false,
                intervals: Vec::new(),
            }],
            created_at: "11".to_string(),
            updated_at: "12".to_string(),
        });
        snapshot.baselines.push(BaselineRecord {
            id: "base_1".to_string(),
            project_id: "project_1".to_string(),
            name: "Baseline 1".to_string(),
            captured_at: "13".to_string(),
            project: ProjectSnapshot {
                id: "project_1".to_string(),
                name: "Alpha".to_string(),
                calculated_start_date: Some("2026-05-01".to_string()),
                calculated_finish_date: Some("2026-05-04".to_string()),
            },
            tasks: vec![TaskSnapshot {
                id: "task_1".to_string(),
                name: "Task".to_string(),
                start_date: Some("2026-05-01".to_string()),
                finish_date: Some("2026-05-02".to_string()),
                percent_complete: 10.0,
            }],
            resources: vec![ResourceSnapshot {
                id: "res_1".to_string(),
                name: "Alice".to_string(),
                max_units: 100.0,
            }],
            assignments: vec![AssignmentSnapshot {
                id: "asg_1".to_string(),
                task_id: "task_1".to_string(),
                resource_id: "res_1".to_string(),
                units: 75.0,
            }],
        });
        snapshot
    }

    #[test]
    fn xml_round_trip_preserves_key_records() {
        let snapshot = sample_snapshot();
        let xml = export_workspace_xml(&snapshot).expect("export xml");
        let imported = import_workspace_xml(&xml).expect("import xml");

        assert_eq!(imported.schema_version, snapshot.schema_version);
        assert_eq!(imported.settings.locale.as_deref(), Some("ja-JP"));
        assert_eq!(imported.projects.len(), 1);
        assert_eq!(imported.tasks.len(), 2);
        assert_eq!(imported.dependencies.len(), 1);
        assert_eq!(imported.resources.len(), 1);
        assert_eq!(imported.assignments.len(), 1);
        assert_eq!(imported.calendars.len(), 1);
        assert_eq!(imported.baselines.len(), 1);
        assert_eq!(imported.projects[0].name, "Alpha");
        assert_eq!(imported.tasks[0].name, "Task");
        assert_eq!(imported.tasks[1].name, "Task S");
        assert_eq!(imported.calendars[0].working_days, vec![1, 2, 3, 4, 5]);
    }
}

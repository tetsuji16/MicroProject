use crate::mspdi::{ChartRange, GanttDependency, GanttTask, ProjectDocument};
use chrono::Datelike;
use eframe::egui::{self, Align, Color32, FontFamily, FontId, Pos2, Rect, RichText, Stroke};
use egui_extras::{Column, TableBuilder};
use std::collections::HashMap;
use std::path::PathBuf;

const BODY_ROW_HEIGHT: f32 = 28.0;
const HEADER_HEIGHT: f32 = 52.0;

pub struct GanttApp {
    document: Option<ProjectDocument>,
    document_path: Option<PathBuf>,
    status: String,
    error: Option<String>,
    zoom_px_per_day: f32,
    dirty: bool,
    selected_task: Option<usize>,
    task_editor: TaskEditorState,
}

#[derive(Clone, Default)]
struct TaskEditorState {
    task_uid: Option<u32>,
    name: String,
    outline_level: u32,
    summary: bool,
    milestone: bool,
    percent_complete: f32,
    start_text: String,
    finish_text: String,
    duration_text: String,
}

impl GanttApp {
    pub fn new(cc: &eframe::CreationContext<'_>, initial_file: Option<PathBuf>) -> Self {
        let mut visuals = egui::Visuals::light();
        visuals.override_text_color = Some(Color32::from_rgb(28, 28, 30));
        visuals.widgets.inactive.bg_fill = Color32::from_rgb(245, 245, 247);
        visuals.widgets.hovered.bg_fill = Color32::from_rgb(235, 239, 247);
        visuals.widgets.active.bg_fill = Color32::from_rgb(222, 232, 248);
        visuals.window_fill = Color32::from_rgb(250, 250, 252);
        cc.egui_ctx.set_visuals(visuals);

        let mut app = Self {
            document: None,
            document_path: None,
            status: "Open an MS Project XML file to begin.".to_string(),
            error: None,
            zoom_px_per_day: 18.0,
            dirty: false,
            selected_task: None,
            task_editor: TaskEditorState::default(),
        };

        if let Some(path) = initial_file {
            app.load_file(path);
        }

        app
    }

    fn load_dialog(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .set_title("Open MS Project XML")
            .add_filter("MS Project XML", &["xml", "mspdi", "pod.xml"])
            .pick_file()
        {
            self.load_file(path);
        }
    }

    fn load_file(&mut self, path: PathBuf) {
        match crate::mspdi::load_project_document(&path) {
            Ok(document) => {
                self.selected_task = if document.tasks.is_empty() { None } else { Some(0) };
                self.status = format!(
                    "Loaded {} task(s) from {}",
                    document.tasks.len(),
                    path.display()
                );
                self.document = Some(document);
                self.document_path = Some(path);
                self.error = None;
                self.dirty = false;
                self.sync_task_editor_with_selection();
            }
            Err(error) => {
                self.error = Some(error.clone());
                self.status = "Failed to load document.".to_string();
            }
        }
    }

    fn save_current_file(&mut self) {
        let Some(path) = self.document_path.clone() else {
            self.save_file_as();
            return;
        };
        self.save_to_path(path);
    }

    fn save_file_as(&mut self) {
        if self.document.is_none() {
            self.status = "Nothing to save.".to_string();
            return;
        }

        if let Some(path) = rfd::FileDialog::new()
            .set_title("Save Project XML")
            .add_filter("MS Project XML", &["xml", "mspdi", "pod.xml"])
            .set_file_name(
                self.document_path
                    .as_ref()
                    .and_then(|path| path.file_name())
                    .and_then(|name| name.to_str())
                    .unwrap_or("project.xml"),
            )
            .save_file()
        {
            self.save_to_path(path);
        }
    }

    fn save_to_path(&mut self, path: PathBuf) {
        let Some(document) = self.document.as_ref() else {
            self.status = "Nothing to save.".to_string();
            return;
        };

        match crate::mspdi::save_project_document(&path, document) {
            Ok(()) => {
                self.status = format!("Saved {}", path.display());
                self.document_path = Some(path);
                self.error = None;
                self.dirty = false;
            }
            Err(error) => {
                self.error = Some(error.clone());
                self.status = "Failed to save document.".to_string();
            }
        }
    }

    fn sync_task_editor_with_selection(&mut self) {
        let Some(document) = self.document.as_ref() else {
            self.task_editor = TaskEditorState::default();
            return;
        };

        let Some(index) = self.selected_task else {
            self.task_editor = TaskEditorState::default();
            return;
        };

        let Some(task) = document.tasks.get(index) else {
            self.task_editor = TaskEditorState::default();
            self.selected_task = None;
            return;
        };

        self.task_editor.sync_from_task(task);
    }

    fn visible_shortcut_open(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::O) && input.modifiers.command)
    }

    fn visible_shortcut_save(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::S) && input.modifiers.command && !input.modifiers.shift)
    }

    fn visible_shortcut_save_as(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::S) && input.modifiers.command && input.modifiers.shift)
    }

    fn visible_shortcut_zoom_in(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::Plus) && input.modifiers.command)
    }

    fn visible_shortcut_zoom_out(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::Minus) && input.modifiers.command)
    }

    fn handle_drop(&mut self, ctx: &egui::Context) {
        let dropped = ctx.input(|input| {
            input
                .raw
                .dropped_files
                .iter()
                .find_map(|file| file.path.clone())
        });

        if let Some(path) = dropped {
            self.load_file(path);
        }
    }

    fn toolbar(&mut self, ui: &mut egui::Ui) {
        ui.horizontal_wrapped(|ui| {
            ui.label(
                RichText::new("ProjectLibre XML Gantt")
                    .font(FontId::new(21.0, FontFamily::Proportional))
                    .strong(),
            );

            if ui.button("Open").clicked() {
                self.load_dialog();
            }

            if ui.button("Reload").clicked() {
                if let Some(path) = self.document_path.clone() {
                    self.load_file(path);
                }
            }

            if ui.button("Save").clicked() {
                self.save_current_file();
            }

            if ui.button("Save As").clicked() {
                self.save_file_as();
            }

            ui.add_space(12.0);
            ui.label("Zoom");
            ui.add(
                egui::Slider::new(&mut self.zoom_px_per_day, 8.0..=40.0)
                    .clamping(egui::SliderClamping::Always)
                    .show_value(true),
            );

            ui.add_space(12.0);
            let mut status = self.status.clone();
            if self.dirty {
                status.push_str("  *unsaved*");
            }
            ui.label(status);

            if let Some(error) = &self.error {
                ui.add_space(12.0);
                ui.colored_label(Color32::from_rgb(170, 40, 40), error);
            }
        });
    }

    fn document_view(&mut self, ui: &mut egui::Ui) {
        egui::Panel::right("task_editor")
            .resizable(true)
            .default_size(360.0)
            .show_inside(ui, |ui| {
                self.task_editor_panel(ui);
            });

        egui::CentralPanel::default().show_inside(ui, |ui| {
            if let Some(document) = &self.document {
                self.document_header(ui, document);
                ui.add_space(8.0);
                self.render_gantt_table(ui, document);
            } else {
                egui::Frame::group(ui.style())
                    .fill(Color32::from_rgb(252, 252, 253))
                    .show(ui, |ui| {
                        ui.vertical_centered(|ui| {
                            ui.add_space(100.0);
                            ui.label(
                                RichText::new("Open an MS Project XML file")
                                    .font(FontId::new(24.0, FontFamily::Proportional))
                                    .strong(),
                            );
                            ui.add_space(8.0);
                            ui.label("The viewer reads Project XML and draws a task table with a Gantt timeline.");
                            ui.add_space(16.0);
                            if ui.button("Choose File").clicked() {
                                self.load_dialog();
                            }
                        });
                    });
            }
        });

        if let Some(index) = ui
            .memory_mut(|memory| memory.data.remove_temp::<usize>(egui::Id::new("task_select")))
        {
            if let Some(task_count) = self.document.as_ref().map(|document| document.tasks.len()) {
                if index < task_count {
                    self.selected_task = Some(index);
                    self.sync_task_editor_with_selection();
                }
            }
        }
    }

    fn document_header(&self, ui: &mut egui::Ui, document: &ProjectDocument) {
        let calendar_count = document.calendars.len();
        let base_calendar_count = document
            .calendars
            .iter()
            .filter(|calendar| calendar.base_calendar)
            .count();
        let calendar_names = document
            .calendars
            .iter()
            .map(|calendar| calendar.name.as_str())
            .take(3)
            .collect::<Vec<_>>()
            .join(", ");
        let task_count = document.tasks.len();
        let dep_count = document.dependencies.len();
        let chart_range = document.chart_range();

        ui.horizontal_wrapped(|ui| {
            ui.label(
                RichText::new(document.title.as_deref().unwrap_or(&document.name))
                    .font(FontId::new(18.0, FontFamily::Proportional))
                    .strong(),
            );
            ui.separator();
            ui.label(format!("Tasks: {task_count}"));
            ui.label(format!("Dependencies: {dep_count}"));
            if calendar_count > 0 {
                ui.label(format!(
                    "Calendars: {calendar_count} ({} base) {}",
                    base_calendar_count,
                    if calendar_names.is_empty() {
                        String::new()
                    } else {
                        format!("[{calendar_names}]")
                    }
                ));
            } else {
                ui.label("Calendars: 0");
            }
            ui.label(format!(
                "Range: {} - {}",
                chart_range.start.format("%Y-%m-%d"),
                chart_range.end.format("%Y-%m-%d")
            ));
            if let Some(manager) = &document.manager {
                ui.label(format!("Manager: {manager}"));
            }
        });
    }

    fn render_gantt_table(&self, ui: &mut egui::Ui, document: &ProjectDocument) {
        let range = document.chart_range();
        let chart_width = (range.days() as f32 * self.zoom_px_per_day).max(320.0);
        let task_by_uid = build_task_index(&document.tasks);
        let geometry = std::cell::RefCell::new(Vec::with_capacity(document.tasks.len()));

        egui::ScrollArea::both()
            .auto_shrink([false, false])
            .show(ui, |ui| {
                let table = TableBuilder::new(ui)
                    .striped(true)
                    .resizable(true)
                    .cell_layout(egui::Layout::left_to_right(Align::Center))
                    .column(Column::exact(56.0))
                    .column(Column::exact(56.0))
                    .column(Column::exact(72.0))
                    .column(Column::exact(280.0))
                    .column(Column::exact(76.0))
                    .column(Column::exact(72.0))
                    .column(Column::exact(84.0))
                    .column(Column::exact(116.0))
                    .column(Column::exact(116.0))
                    .column(Column::exact(122.0))
                    .column(Column::exact(110.0))
                    .column(Column::exact(chart_width));

                table
                    .header(HEADER_HEIGHT, |mut header| {
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("ID").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("UID").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Level").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Task").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Flag").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Type").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Duration").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Start").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Finish").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("%").strong());
                            });
                        });
                        header.col(|ui| {
                            ui.centered_and_justified(|ui| {
                                ui.label(RichText::new("Predecessors").strong());
                            });
                        });
                        header.col(|ui| {
                            let rect = ui.max_rect();
                            draw_timescale_header(ui, rect, range, self.zoom_px_per_day);
                        });
                    })
                    .body(|body| {
                        body.rows(BODY_ROW_HEIGHT, document.tasks.len(), |mut row| {
                            let index = row.index();
                            let task = &document.tasks[index];
                            let selected = self.selected_task == Some(index);
                            row.col(|ui| {
                                ui.label(task.id.to_string());
                            });
                            row.col(|ui| {
                                ui.label(task.uid.to_string());
                            });
                            row.col(|ui| {
                                ui.label(task.outline_level.to_string());
                            });
                            row.col(|ui| {
                                let response = draw_task_name(ui, task, selected);
                                if response.clicked() {
                                    ui.memory_mut(|memory| memory.data.insert_temp(egui::Id::new("task_select"), index));
                                }
                            });
                            row.col(|ui| {
                                ui.label(if task.summary {
                                    "Summary"
                                } else if task.milestone {
                                    "Milestone"
                                } else {
                                    "Task"
                                });
                            });
                            row.col(|ui| {
                                ui.label(if task.summary {
                                    "Summary"
                                } else if task.milestone {
                                    "Milestone"
                                } else {
                                    "Task"
                                });
                            });
                            row.col(|ui| {
                                ui.label(task.duration_text.as_deref().unwrap_or(""));
                            });
                            row.col(|ui| {
                                ui.label(format_option_datetime(task.start));
                            });
                            row.col(|ui| {
                                ui.label(format_option_datetime(task.finish));
                            });
                            row.col(|ui| {
                                ui.label(format!("{:.0}%", task.percent_complete));
                            });
                            row.col(|ui| {
                                ui.label(task.predecessor_text.clone());
                            });
                            row.col(|ui| {
                                let rect = ui.max_rect();
                                let painter = ui.painter();
                                paint_chart_row_background(painter, rect, range, self.zoom_px_per_day);
                                let geom = draw_task_bar(
                                    painter,
                                    rect,
                                    task,
                                    range,
                                    self.zoom_px_per_day,
                                );
                                geometry.borrow_mut().push(geom);
                            });
                        });
                    });

                let rects = geometry.into_inner();
                if !rects.is_empty() {
                    draw_dependencies(
                        ui.painter(),
                        &rects,
                        &task_by_uid,
                        &document.dependencies,
                    );
                }
            });
    }

    fn task_editor_panel(&mut self, ui: &mut egui::Ui) {
        let mut request_new = false;
        let mut request_duplicate = false;
        let mut request_delete = false;
        let mut request_select_first = false;

        ui.vertical(|ui| {
            ui.heading("Edit Task");
            ui.add_space(4.0);

            let Some(document) = self.document.as_mut() else {
                ui.label("Open a Project XML file to edit tasks.");
                return;
            };

            let Some(index) = self.selected_task else {
                ui.label("Select a task in the table to edit it.");
                if !document.tasks.is_empty() && ui.button("Select first task").clicked() {
                    request_select_first = true;
                }
                return;
            };

            if index >= document.tasks.len() {
                ui.label("The selected task is no longer available.");
                return;
            }

            let task_uid = document.tasks[index].uid;
            if self.task_editor.task_uid != Some(task_uid) {
                self.task_editor.sync_from_task(&document.tasks[index]);
            }

            let task = &mut document.tasks[index];
            let mut changed = false;

            ui.label(format!("UID: {}", task.uid));
            ui.label(format!("ID: {}", task.id));
            ui.separator();

            ui.label("Name");
            changed |= ui.text_edit_singleline(&mut self.task_editor.name).changed();

            ui.horizontal(|ui| {
                ui.label("Outline");
                changed |= ui
                    .add(
                        egui::DragValue::new(&mut self.task_editor.outline_level)
                            .range(1..=99)
                            .speed(1.0),
                    )
                    .changed();
            });

            changed |= ui.checkbox(&mut self.task_editor.summary, "Summary task").changed();
            changed |= ui.checkbox(&mut self.task_editor.milestone, "Milestone").changed();

            ui.horizontal(|ui| {
                ui.label("Percent");
                changed |= ui
                    .add(
                        egui::Slider::new(&mut self.task_editor.percent_complete, 0.0..=100.0)
                            .show_value(true),
                    )
                    .changed();
            });

            ui.label("Start");
            changed |= ui.text_edit_singleline(&mut self.task_editor.start_text).changed();
            ui.label("Finish");
            changed |= ui.text_edit_singleline(&mut self.task_editor.finish_text).changed();
            ui.label("Duration");
            changed |= ui.text_edit_singleline(&mut self.task_editor.duration_text).changed();

            if changed {
                task.name = self.task_editor.name.clone();
                task.outline_level = self.task_editor.outline_level.max(1);
                task.summary = self.task_editor.summary;
                task.milestone = self.task_editor.milestone;
                task.percent_complete = self.task_editor.percent_complete.clamp(0.0, 100.0);
                task.duration_text = (!self.task_editor.duration_text.trim().is_empty())
                    .then(|| self.task_editor.duration_text.clone());
                Self::apply_editor_dates(&self.task_editor, task);
                self.dirty = true;
            }

            ui.separator();
            ui.horizontal(|ui| {
                if ui.button("New Task").clicked() {
                    request_new = true;
                }
                if ui.button("Duplicate").clicked() {
                    request_duplicate = true;
                }
                if ui.button("Delete").clicked() {
                    request_delete = true;
                }
            });

            ui.separator();
            ui.label(RichText::new("Read-only dependency links").strong());
            ui.label(task.predecessor_text.as_str());
        });

        if request_select_first {
            if let Some(document) = self.document.as_ref() {
                if let Some(task) = document.tasks.first() {
                    self.selected_task = Some(0);
                    self.task_editor.sync_from_task(task);
                }
            }
        }
        if request_new {
            self.insert_task_after_selected(false);
        }
        if request_duplicate {
            self.insert_task_after_selected(true);
        }
        if request_delete {
            self.delete_selected_task();
        }
    }

    fn apply_editor_dates(editor: &TaskEditorState, task: &mut GanttTask) {
        let start_text = editor.start_text.trim();
        task.start = if start_text.is_empty() {
            None
        } else {
            crate::mspdi::parse_date_time(Some(start_text)).or(task.start)
        };

        let finish_text = editor.finish_text.trim();
        task.finish = if finish_text.is_empty() {
            None
        } else {
            crate::mspdi::parse_date_time(Some(finish_text)).or(task.finish)
        };
    }

    fn insert_task_after_selected(&mut self, duplicate: bool) {
        let inserted_task = {
            let Some(document) = self.document.as_mut() else {
                return;
            };

            let insert_index = self
                .selected_task
                .map(|index| index + 1)
                .unwrap_or(document.tasks.len());
            let next_uid = document
                .tasks
                .iter()
                .map(|task| task.uid)
                .max()
                .unwrap_or(0)
                + 1;
            let next_id = document
                .tasks
                .iter()
                .map(|task| task.id)
                .max()
                .unwrap_or(0)
                + 1;

            let mut task = if duplicate {
                self.selected_task
                    .and_then(|index| document.tasks.get(index).cloned())
                    .unwrap_or_else(|| GanttTask {
                        uid: next_uid,
                        id: next_id,
                        name: String::new(),
                        outline_level: 1,
                        summary: false,
                        milestone: false,
                        percent_complete: 0.0,
                        start: None,
                        finish: None,
                        duration_text: None,
                        predecessor_text: String::new(),
                    })
            } else {
                GanttTask {
                    uid: next_uid,
                    id: next_id,
                    name: String::new(),
                    outline_level: 1,
                    summary: false,
                    milestone: false,
                    percent_complete: 0.0,
                    start: None,
                    finish: None,
                    duration_text: None,
                    predecessor_text: String::new(),
                }
            };

            if duplicate {
                if task.name.trim().is_empty() {
                    task.name = "Copy of task".to_string();
                } else {
                    task.name = format!("{} (Copy)", task.name);
                }
                task.uid = next_uid;
                task.id = next_id;
            }

            document.tasks.insert(insert_index, task);
            self.selected_task = Some(insert_index);
            document.tasks.get(insert_index).cloned()
        };

        if let Some(task) = inserted_task.as_ref() {
            self.task_editor.sync_from_task(task);
        }
        self.dirty = true;
        self.status = if duplicate {
            "Duplicated task.".to_string()
        } else {
            "Created new task.".to_string()
        };
    }

    fn delete_selected_task(&mut self) {
        let next_task = {
            let Some(document) = self.document.as_mut() else {
                return;
            };
            let Some(index) = self.selected_task else {
                return;
            };
            if index >= document.tasks.len() {
                return;
            }

            let deleted_uid = document.tasks[index].uid;
            document.tasks.remove(index);
            document.dependencies.retain(|dependency| {
                dependency.predecessor_uid != deleted_uid && dependency.successor_uid != deleted_uid
            });
            self.selected_task = if document.tasks.is_empty() {
                None
            } else {
                Some(index.min(document.tasks.len() - 1))
            };
            self.selected_task
                .and_then(|index| document.tasks.get(index).cloned())
        };

        if let Some(task) = next_task.as_ref() {
            self.task_editor.sync_from_task(task);
        } else {
            self.task_editor = TaskEditorState::default();
        }
        self.dirty = true;
        self.status = "Deleted task.".to_string();
    }
}

impl eframe::App for GanttApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = ui.ctx();

        if Self::visible_shortcut_open(ctx) {
            self.load_dialog();
        }
        if Self::visible_shortcut_save(ctx) {
            self.save_current_file();
        }
        if Self::visible_shortcut_save_as(ctx) {
            self.save_file_as();
        }
        if Self::visible_shortcut_zoom_in(ctx) {
            self.zoom_px_per_day = (self.zoom_px_per_day + 2.0).min(40.0);
        }
        if Self::visible_shortcut_zoom_out(ctx) {
            self.zoom_px_per_day = (self.zoom_px_per_day - 2.0).max(8.0);
        }

        self.handle_drop(ctx);
        self.toolbar(ui);
        self.document_view(ui);
    }
}

impl TaskEditorState {
    fn sync_from_task(&mut self, task: &GanttTask) {
        self.task_uid = Some(task.uid);
        self.name = task.name.clone();
        self.outline_level = task.outline_level.max(1);
        self.summary = task.summary;
        self.milestone = task.milestone;
        self.percent_complete = task.percent_complete.clamp(0.0, 100.0);
        self.start_text = format_option_datetime(task.start);
        self.finish_text = format_option_datetime(task.finish);
        self.duration_text = task.duration_text.clone().unwrap_or_default();
    }
}

#[derive(Clone)]
struct RowGeometry {
    start_x: f32,
    end_x: f32,
    center_y: f32,
    milestone: bool,
}

fn build_task_index(tasks: &[GanttTask]) -> HashMap<u32, usize> {
    tasks
        .iter()
        .enumerate()
        .map(|(index, task)| (task.uid, index))
        .collect()
}

fn format_option_datetime(value: Option<chrono::NaiveDateTime>) -> String {
    value
        .map(|dt| dt.format("%Y-%m-%d %H:%M").to_string())
        .unwrap_or_default()
}

fn draw_task_name(ui: &mut egui::Ui, task: &GanttTask, selected: bool) -> egui::Response {
    let indent = task.outline_level.saturating_sub(1) as f32 * 14.0;
    ui.horizontal(|ui| {
        ui.add_space(indent);
        let prefix = if task.summary {
            "▣"
        } else if task.milestone {
            "◆"
        } else {
            "•"
        };
        let title = if task.name.trim().is_empty() {
            "[Untitled]"
        } else {
            &task.name
        };
        let text = RichText::new(format!("{prefix} {title}"))
            .strong()
            .color(if task.summary {
                Color32::from_rgb(24, 24, 24)
            } else {
                Color32::from_rgb(40, 54, 92)
            });
        ui.selectable_label(selected, text)
    })
    .inner
}

fn paint_chart_row_background(painter: &egui::Painter, rect: Rect, range: ChartRange, px_per_day: f32) {
    painter.rect_filled(rect, 0.0, Color32::from_rgb(252, 252, 253));
    let days = range.days();
    for day in 0..=days {
        let x = rect.left() + day as f32 * px_per_day;
        let line_color = if day % 7 == 0 {
            Color32::from_rgb(214, 219, 226)
        } else {
            Color32::from_rgb(235, 238, 243)
        };
        painter.line_segment(
            [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
            Stroke::new(1.0, line_color),
        );
    }
}

fn draw_task_bar(
    painter: &egui::Painter,
    rect: Rect,
    task: &GanttTask,
    range: ChartRange,
    px_per_day: f32,
) -> RowGeometry {
    let row_center_y = rect.center().y;
    let bar_height = if task.summary {
        8.0
    } else if task.milestone {
        0.0
    } else {
        12.0
    };
    let start_day = task
        .start
        .map(|dt| dt.date())
        .unwrap_or(range.start);
    let finish_day = task
        .finish
        .map(|dt| dt.date())
        .unwrap_or(start_day);

    let start_offset = (start_day - range.start).num_days().max(0) as f32;
    let finish_offset = (finish_day - range.start).num_days().max(0) as f32 + 1.0;

    let mut start_x = rect.left() + start_offset * px_per_day;
    let mut end_x = rect.left() + finish_offset * px_per_day;
    if end_x < start_x {
        std::mem::swap(&mut start_x, &mut end_x);
    }
    if (end_x - start_x).abs() < 1.0 {
        end_x = start_x + 2.0;
    }

    let base_color = if task.summary {
        Color32::from_rgb(12, 12, 12)
    } else if task.milestone {
        Color32::from_rgb(212, 72, 48)
    } else {
        Color32::from_rgb(94, 144, 255)
    };
    let fill_color = if task.summary {
        Color32::from_rgb(28, 28, 28)
    } else {
        base_color
    };
    let outline = if task.summary {
        Color32::from_rgb(12, 12, 12)
    } else {
        Color32::from_rgb(53, 94, 176)
    };

    if task.milestone {
        let center = Pos2::new(start_x, row_center_y);
        let points = [
            Pos2::new(center.x, center.y - 7.0),
            Pos2::new(center.x + 7.0, center.y),
            Pos2::new(center.x, center.y + 7.0),
            Pos2::new(center.x - 7.0, center.y),
        ];
        painter.add(egui::Shape::convex_polygon(
            points.to_vec(),
            fill_color,
            Stroke::new(1.0, outline),
        ));
    } else if task.summary {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - bar_height / 2.0),
            Pos2::new(end_x, row_center_y + bar_height / 2.0),
        );
        painter.rect_filled(bar_rect, 1.0, fill_color);
        painter.rect_stroke(
            bar_rect,
            1.0,
            Stroke::new(1.2, outline),
            egui::StrokeKind::Inside,
        );
        painter.line_segment(
            [
                Pos2::new(bar_rect.left(), bar_rect.bottom() + 2.0),
                Pos2::new(bar_rect.right(), bar_rect.bottom() + 2.0),
            ],
            Stroke::new(1.5, outline),
        );
    } else {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - bar_height / 2.0),
            Pos2::new(end_x, row_center_y + bar_height / 2.0),
        );
        painter.rect_filled(bar_rect, 2.0, Color32::from_rgb(219, 229, 255));
        painter.rect_filled(
            Rect::from_min_max(
                bar_rect.left_top(),
                Pos2::new(
                    bar_rect.left() + bar_rect.width() * (task.percent_complete / 100.0).clamp(0.0, 1.0),
                    bar_rect.bottom(),
                ),
            ),
            2.0,
            fill_color,
        );
        painter.rect_stroke(
            bar_rect,
            2.0,
            Stroke::new(1.0, outline),
            egui::StrokeKind::Inside,
        );
    }

    RowGeometry {
        start_x,
        end_x,
        center_y: row_center_y,
        milestone: task.milestone,
    }
}

fn draw_timescale_header(
    ui: &mut egui::Ui,
    rect: Rect,
    range: ChartRange,
    px_per_day: f32,
) {
    let painter = ui.painter();
    painter.rect_filled(rect, 0.0, Color32::from_rgb(245, 247, 251));
    painter.rect_stroke(
        rect,
        0.0,
        Stroke::new(1.0, Color32::from_rgb(201, 208, 218)),
        egui::StrokeKind::Inside,
    );

    let month_h = rect.height() * 0.44;
    let month_text_color = Color32::from_rgb(48, 56, 72);
    let day_text_color = Color32::from_rgb(68, 76, 90);
    let days = range.days();

    let mut current_month = range.start.month();
    let mut month_start = 0;
    for day in 0..=days {
        let date = range.start + chrono::Duration::days(day);
        let x = rect.left() + day as f32 * px_per_day;
        let is_month_boundary = day == 0 || date.month() != current_month || day == days;
        if is_month_boundary && day > 0 {
            let label = format!("{}-{:02}", date.year(), current_month);
            let start_x = rect.left() + month_start as f32 * px_per_day;
            let end_x = x;
            let month_rect = Rect::from_min_max(
                Pos2::new(start_x + 4.0, rect.top() + 2.0),
                Pos2::new(end_x - 4.0, rect.top() + month_h - 2.0),
            );
            painter.text(
                month_rect.center(),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(13.0, FontFamily::Proportional),
                month_text_color,
            );
            painter.line_segment(
                [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
                Stroke::new(1.2, Color32::from_rgb(170, 178, 190)),
            );
            current_month = date.month();
            month_start = day;
        }

        if day < days {
            let label = date.format("%d").to_string();
            let day_rect = Rect::from_min_max(
                Pos2::new(x, rect.top() + month_h),
                Pos2::new(x + px_per_day, rect.bottom()),
            );
            painter.text(
                day_rect.center(),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(11.0, FontFamily::Proportional),
                day_text_color,
            );

            if day % 7 == 0 {
                let weekday = match date.weekday().number_from_monday() {
                    1 => "Mon",
                    2 => "Tue",
                    3 => "Wed",
                    4 => "Thu",
                    5 => "Fri",
                    6 => "Sat",
                    _ => "Sun",
                };
                painter.text(
                    Pos2::new(day_rect.center().x, day_rect.bottom() - 2.0),
                    egui::Align2::CENTER_BOTTOM,
                    weekday,
                    FontId::new(9.0, FontFamily::Proportional),
                    Color32::from_rgb(117, 124, 137),
                );
            }
        }
    }
}

fn draw_dependencies(
    painter: &egui::Painter,
    row_geometries: &[RowGeometry],
    task_by_uid: &HashMap<u32, usize>,
    dependencies: &[GanttDependency],
) {
    for dependency in dependencies {
        let Some(&pred_index) = task_by_uid.get(&dependency.predecessor_uid) else {
            continue;
        };
        let Some(&succ_index) = task_by_uid.get(&dependency.successor_uid) else {
            continue;
        };
        let predecessor = &row_geometries[pred_index];
        let successor = &row_geometries[succ_index];

        let start = Pos2::new(predecessor.end_x, predecessor.center_y);
        let end = if successor.milestone {
            Pos2::new(successor.start_x, successor.center_y)
        } else {
            Pos2::new(successor.start_x, successor.center_y)
        };

        let mid_x = (start.x + end.x) / 2.0;
        let path = vec![
            start,
            Pos2::new(mid_x, start.y),
            Pos2::new(mid_x, end.y),
            Pos2::new(end.x - 5.0, end.y),
        ];
        for segment in path.windows(2) {
            painter.line_segment(
                [segment[0], segment[1]],
                Stroke::new(1.0, Color32::from_rgb(96, 96, 96)),
            );
        }
        painter.line_segment(
            [Pos2::new(end.x - 5.0, end.y - 4.0), end],
            Stroke::new(1.0, Color32::from_rgb(96, 96, 96)),
        );
        painter.line_segment(
            [Pos2::new(end.x - 5.0, end.y + 4.0), end],
            Stroke::new(1.0, Color32::from_rgb(96, 96, 96)),
        );
        let label = if let Some(lag) = &dependency.lag_text {
            format!("{} {}", dependency.relation, lag)
        } else {
            dependency.relation.clone()
        };
        painter.text(
            Pos2::new(mid_x + 4.0, end.y - 2.0),
            egui::Align2::LEFT_BOTTOM,
            label,
            FontId::new(10.0, FontFamily::Proportional),
            Color32::from_rgb(86, 86, 86),
        );
    }
}

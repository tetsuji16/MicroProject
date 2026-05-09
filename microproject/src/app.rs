use crate::dependency::{
    add_dependency_from_drag, dependency_exists, refresh_predecessor_text_for_task,
};
use crate::mspdi::{parse_date_time, ChartRange, GanttDependency, GanttTask, ProjectDocument};
use chrono::{Datelike, Duration, NaiveDate};
use eframe::egui::{
    self, Align, Color32, Event, FontData, FontDefinitions, FontFamily, FontId, Pos2, Rect,
    RichText, Sense, Stroke, Vec2,
};
use egui_extras::{Column, TableBuilder};
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use crate::table_view::{
    sheet_fixed_columns_width, SHEET_DATE_WIDTH, SHEET_DURATION_WIDTH, SHEET_NAME_WIDTH,
    SHEET_PERCENT_WIDTH, SHEET_PREDECESSOR_WIDTH, SHEET_ROW_HEADER_WIDTH,
};
use crate::timeline_view::{chart_width as compute_chart_width, table_width as sum_widths};

const BODY_ROW_HEIGHT: f32 = 24.0;
const HEADER_HEIGHT: f32 = 44.0;
const SHEET_CELL_PADDING_X: f32 = 8.0;
const SHEET_CELL_PADDING_Y: f32 = 4.0;
const APP_BG: Color32 = Color32::from_rgb(212, 208, 200);
const PANEL_BG: Color32 = Color32::from_rgb(236, 233, 216);
const SHEET_BG: Color32 = Color32::from_rgb(255, 255, 255);
const HEADER_BG: Color32 = Color32::from_rgb(242, 242, 242);
const BORDER: Color32 = Color32::from_rgb(197, 205, 214);
const GRID_LINE: Color32 = Color32::from_rgb(232, 237, 242);
const GRID_LINE_STRONG: Color32 = Color32::from_rgb(208, 218, 228);
const SHEET_SELECTED_FILL: Color32 = Color32::from_rgb(219, 235, 255);
const SHEET_SUMMARY_FILL: Color32 = Color32::from_rgb(248, 249, 251);
const SHEET_CRITICAL_FILL: Color32 = Color32::from_rgb(255, 246, 243);
const SHEET_ROW_HEADER_FILL: Color32 = Color32::from_rgb(246, 248, 250);
const SHEET_ROW_HEADER_SELECTED_FILL: Color32 = Color32::from_rgb(205, 226, 251);
const SHEET_HEADER_TOP: Color32 = Color32::from_rgb(251, 251, 251);
const SHEET_HEADER_BOTTOM: Color32 = Color32::from_rgb(229, 234, 239);
const SHEET_HEADER_TEXT: Color32 = Color32::from_rgb(67, 74, 83);
const SHEET_SELECTION_BORDER: Color32 = Color32::from_rgb(0, 84, 147);
const SHEET_ACTIVE_BORDER: Color32 = Color32::from_rgb(0, 84, 147);
const TITLE_BAR_BG: Color32 = Color32::from_rgb(10, 36, 106);
const TITLE_BAR_BG_DARK: Color32 = Color32::from_rgb(4, 22, 78);
const TITLE_BAR_TEXT: Color32 = Color32::from_rgb(248, 248, 248);
const OFFICE_GREEN_DARK: Color32 = Color32::from_rgb(24, 93, 31);
const RIBBON_CHROME_BG: Color32 = Color32::from_rgb(236, 233, 216);
const TEXT: Color32 = Color32::from_rgb(30, 30, 30);
const SUBTEXT: Color32 = Color32::from_rgb(78, 78, 78);
const ACCENT: Color32 = Color32::from_rgb(0, 84, 147);
const ACCENT_SOFT: Color32 = Color32::from_rgb(217, 228, 242);
const CRITICAL: Color32 = Color32::from_rgb(194, 60, 40);
const BASELINE: Color32 = Color32::from_rgb(148, 148, 148);
const WEEKEND_BG: Color32 = Color32::from_rgb(248, 248, 248);
const TODAY: Color32 = Color32::from_rgb(242, 171, 61);
const UNDO_STACK_LIMIT: usize = 64;
const APP_SETTINGS_DIR: &str = "microproject";
const APP_SETTINGS_FILE: &str = "settings.txt";
const TASK_BAR_HANDLE_WIDTH: f32 = 7.0;
const TASK_BAR_HANDLE_HEIGHT: f32 = 13.0;
const TASK_BAR_LINK_HANDLE_GAP: f32 = 7.0;

const JAPANESE_FONT_CANDIDATES: &[&str] = &[
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\meiryob.ttc",
    r"C:\Windows\Fonts\YuGothR.ttc",
    r"C:\Windows\Fonts\YuGothM.ttc",
    r"C:\Windows\Fonts\YuGothB.ttc",
    r"C:\Windows\Fonts\NotoSansJP-VF.ttf",
    r"C:\Windows\Fonts\msgothic.ttc",
];

fn configure_fonts(ctx: &egui::Context) {
    let mut fonts = FontDefinitions::default();
    for candidate in JAPANESE_FONT_CANDIDATES {
        if let Ok(bytes) = fs::read(candidate) {
            let name = Path::new(candidate)
                .file_stem()
                .and_then(|value| value.to_str())
                .unwrap_or("japanese_font")
                .to_string();
            fonts
                .font_data
                .insert(name.clone(), FontData::from_owned(bytes).into());
            fonts
                .families
                .entry(FontFamily::Proportional)
                .or_default()
                .insert(0, name.clone());
            fonts
                .families
                .entry(FontFamily::Monospace)
                .or_default()
                .insert(0, name);
        }
    }
    ctx.set_fonts(fonts);
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RibbonIcon {
    Open,
    Save,
    Undo,
    Redo,
    ZoomIn,
    ZoomOut,
    ExpandAll,
    CollapseAll,
    Filter,
    TaskEdit,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum HeaderGlyph {
    Mode,
    Name,
    Duration,
    Date,
    Percent,
    Predecessors,
    Gantt,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SheetColumn {
    Mode,
    Name,
    Duration,
    Start,
    Finish,
    PercentComplete,
    Predecessors,
}

impl SheetColumn {
    const ORDER: [SheetColumn; 7] = [
        SheetColumn::Mode,
        SheetColumn::Name,
        SheetColumn::Duration,
        SheetColumn::Start,
        SheetColumn::Finish,
        SheetColumn::PercentComplete,
        SheetColumn::Predecessors,
    ];

    fn next(self) -> Self {
        let index = Self::ORDER.iter().position(|col| *col == self).unwrap_or(0);
        Self::ORDER[(index + 1).min(Self::ORDER.len() - 1)]
    }

    fn prev(self) -> Self {
        let index = Self::ORDER.iter().position(|col| *col == self).unwrap_or(0);
        Self::ORDER[index.saturating_sub(1)]
    }

    fn first() -> Self {
        Self::ORDER[0]
    }

    fn last() -> Self {
        Self::ORDER[Self::ORDER.len() - 1]
    }

    fn index(self) -> usize {
        Self::ORDER.iter().position(|col| *col == self).unwrap_or(0)
    }

    fn from_index(index: usize) -> Option<Self> {
        Self::ORDER.get(index).copied()
    }

    fn is_editable(self) -> bool {
        !matches!(self, SheetColumn::Mode)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct SheetCellSelection {
    task_index: usize,
    column: SheetColumn,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SheetNavigationDirection {
    Left,
    Right,
    Up,
    Down,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TaskDragKind {
    Move,
    ResizeStart,
    ResizeFinish,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DependencyRelationChoice {
    FS,
    SS,
    FF,
    SF,
}

impl DependencyRelationChoice {
    const ALL: [DependencyRelationChoice; 4] = [
        DependencyRelationChoice::FS,
        DependencyRelationChoice::SS,
        DependencyRelationChoice::FF,
        DependencyRelationChoice::SF,
    ];

    fn label(self) -> &'static str {
        match self {
            DependencyRelationChoice::FS => "FS",
            DependencyRelationChoice::SS => "SS",
            DependencyRelationChoice::FF => "FF",
            DependencyRelationChoice::SF => "SF",
        }
    }
}

#[derive(Clone, Debug)]
enum ChartInteractionKind {
    Task {
        task_uid: u32,
        kind: TaskDragKind,
        pointer_origin_x: f32,
    },
    Dependency {
        source_task_uid: u32,
        pointer_origin: Pos2,
    },
}

#[derive(Clone, Debug)]
struct ChartInteractionState {
    original_document: ProjectDocument,
    dirty_before: bool,
    kind: ChartInteractionKind,
}

#[derive(Clone, Debug)]
struct DependencyPickerState {
    original_document: ProjectDocument,
    source_task_uid: u32,
    target_task_uid: u32,
    relation: DependencyRelationChoice,
    lag_text: String,
    popup_pos: Pos2,
    dirty_before: bool,
}

pub struct GanttApp {
    document: Option<ProjectDocument>,
    document_path: Option<PathBuf>,
    recent_files: Vec<PathBuf>,
    status: String,
    error: Option<String>,
    zoom_px_per_day: f32,
    status_date: NaiveDate,
    status_date_text: String,
    status_date_picker_open: bool,
    status_date_picker_month: NaiveDate,
    filter_text: String,
    show_weekends: bool,
    show_critical_path: bool,
    dirty: bool,
    selected_task: Option<usize>,
    selected_sheet_column: SheetColumn,
    sheet_selection_anchor: Option<SheetCellSelection>,
    sheet_editing_cell: Option<SheetCellSelection>,
    sheet_selection_dragging: bool,
    chart_interaction: Option<ChartInteractionState>,
    dependency_picker: Option<DependencyPickerState>,
    collapsed_task_uids: HashSet<u32>,
    task_editor: TaskEditorState,
    undo_stack: Vec<UndoState>,
    redo_stack: Vec<UndoState>,
}

#[derive(Clone, Default)]
struct TaskEditorState {
    task_uid: Option<u32>,
    name: String,
    outline_level: u32,
    summary: bool,
    milestone: bool,
    critical: bool,
    percent_complete: f32,
    start_text: String,
    finish_text: String,
    baseline_start_text: String,
    baseline_finish_text: String,
    duration_text: String,
    notes_text: String,
    resource_names_text: String,
    calendar_uid_text: String,
    constraint_type_text: String,
}

#[derive(Clone)]
struct UndoState {
    document: ProjectDocument,
    selected_task: Option<usize>,
    selected_sheet_column: SheetColumn,
    sheet_selection_anchor: Option<SheetCellSelection>,
    dirty: bool,
}

impl GanttApp {
    fn default_status_date() -> NaiveDate {
        chrono::Local::now().date_naive()
    }

    fn settings_path() -> Option<PathBuf> {
        dirs::config_dir().map(|dir| dir.join(APP_SETTINGS_DIR).join(APP_SETTINGS_FILE))
    }

    fn load_app_settings() -> Option<AppSettings> {
        let path = Self::settings_path()?;
        let text = fs::read_to_string(path).ok()?;
        parse_app_settings_text(&text)
    }

    fn save_app_settings(&self) {
        let Some(path) = Self::settings_path() else {
            return;
        };
        if let Some(parent) = path.parent() {
            if fs::create_dir_all(parent).is_err() {
                return;
            }
        }

        let mut file = match fs::File::create(path) {
            Ok(file) => file,
            Err(_) => return,
        };
        let _ = writeln!(
            file,
            "status_date={}",
            self.status_date.format("%Y-%m-%d")
        );
    }

    pub fn new(cc: &eframe::CreationContext<'_>, initial_file: Option<PathBuf>) -> Self {
        configure_fonts(&cc.egui_ctx);

        let loaded_settings = Self::load_app_settings();
        let status_date = loaded_settings
            .as_ref()
            .map(|settings| settings.status_date)
            .unwrap_or_else(Self::default_status_date);

        let mut visuals = egui::Visuals::light();
        visuals.override_text_color = Some(TEXT);
        visuals.panel_fill = APP_BG;
        visuals.window_fill = APP_BG;
        visuals.faint_bg_color = PANEL_BG;
        visuals.extreme_bg_color = SHEET_BG;
        visuals.code_bg_color = SHEET_BG;
        visuals.widgets.noninteractive.bg_fill = APP_BG;
        visuals.widgets.noninteractive.fg_stroke.color = TEXT;
        visuals.widgets.inactive.bg_fill = Color32::from_rgb(245, 243, 237);
        visuals.widgets.hovered.bg_fill = Color32::from_rgb(250, 248, 242);
        visuals.widgets.active.bg_fill = Color32::from_rgb(232, 229, 221);
        visuals.widgets.inactive.fg_stroke.color = TEXT;
        visuals.widgets.hovered.fg_stroke.color = TEXT;
        visuals.widgets.active.fg_stroke.color = TEXT;
        visuals.selection.bg_fill = ACCENT_SOFT;
        visuals.hyperlink_color = ACCENT;
        cc.egui_ctx.set_visuals(visuals);

        let mut app = Self {
            document: None,
            document_path: None,
            recent_files: Vec::new(),
            status: "MS Project XML ファイルを開いてください。".to_string(),
            error: None,
            zoom_px_per_day: 18.0,
            status_date,
            status_date_text: status_date.format("%Y-%m-%d").to_string(),
            status_date_picker_open: false,
            status_date_picker_month: first_day_of_month(status_date),
            filter_text: String::new(),
            show_weekends: true,
            show_critical_path: true,
            dirty: false,
            selected_task: None,
            selected_sheet_column: SheetColumn::Name,
            sheet_selection_anchor: None,
            sheet_editing_cell: None,
            sheet_selection_dragging: false,
            chart_interaction: None,
            dependency_picker: None,
            collapsed_task_uids: HashSet::new(),
            task_editor: TaskEditorState::default(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        };

        if let Some(path) = initial_file {
            app.load_file(path);
        }

        app
    }

    fn load_dialog(&mut self) {
        let mut dialog = rfd::FileDialog::new()
            .set_title("MS Project XML を開く")
            .add_filter("MS Project XML", &["xml", "mspdi"]);

        if let Some(desktop) = dirs::desktop_dir() {
            dialog = dialog.set_directory(desktop);
        }

        if let Some(path) = dialog.pick_file() {
            self.load_file(path);
        }
    }

    fn load_file(&mut self, path: PathBuf) {
        match crate::mspdi::load_project_document(&path) {
            Ok(document) => {
                self.register_recent_file(&path);
                let mut document = document;
                for task in &mut document.tasks {
                    task.duration_text = normalize_duration_text(&task.duration_text);
                }
                self.selected_task = if document.tasks.is_empty() {
                    None
                } else {
                    Some(0)
                };
                self.selected_sheet_column = SheetColumn::Name;
                self.sheet_selection_anchor =
                    self.selected_task.map(|task_index| SheetCellSelection {
                        task_index,
                        column: self.selected_sheet_column,
                    });
                self.sheet_editing_cell = None;
                self.document = Some(document);
                self.document_path = Some(path.clone());
                self.error = None;
                self.dirty = false;
                self.chart_interaction = None;
                self.dependency_picker = None;
                self.undo_stack.clear();
                self.redo_stack.clear();
                self.status = format!("読み込みました: {}", path.display());
                self.sync_task_editor_with_selection();
            }
            Err(error) => {
                self.error = Some(error.clone());
                self.status = "ドキュメントの読み込みに失敗しました。".to_string();
            }
        }
    }

    fn register_recent_file(&mut self, path: &Path) {
        self.recent_files.retain(|existing| existing != path);
        self.recent_files.insert(0, path.to_path_buf());
        self.recent_files.truncate(8);
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
            self.status = "保存する内容がありません。".to_string();
            return;
        }

        let mut dialog = rfd::FileDialog::new()
            .set_title("MS Project XML を保存")
            .add_filter("MS Project XML", &["xml", "mspdi"])
            .set_file_name(
                self.document_path
                    .as_ref()
                    .and_then(|path| path.file_name())
                    .and_then(|name| name.to_str())
                    .unwrap_or("project.xml"),
            );

        if let Some(desktop) = dirs::desktop_dir() {
            dialog = dialog.set_directory(desktop);
        }

        if let Some(path) = dialog.save_file() {
            self.save_to_path(path);
        }
    }

    fn save_to_path(&mut self, path: PathBuf) {
        let Some(document) = self.document.as_ref() else {
            self.status = "保存する内容がありません。".to_string();
            return;
        };

        match crate::mspdi::save_project_document(&path, document) {
            Ok(()) => {
                self.register_recent_file(&path);
                self.status = format!("保存しました: {}", path.display());
                self.document_path = Some(path);
                self.error = None;
                self.dirty = false;
            }
            Err(error) => {
                self.error = Some(error.clone());
                self.status = "ドキュメントの保存に失敗しました。".to_string();
            }
        }
    }

    fn record_undo_state_from(
        &mut self,
        document: ProjectDocument,
        selected_task: Option<usize>,
        selected_sheet_column: SheetColumn,
        sheet_selection_anchor: Option<SheetCellSelection>,
        dirty: bool,
    ) {
        self.undo_stack.push(UndoState {
            document,
            selected_task,
            selected_sheet_column,
            sheet_selection_anchor,
            dirty,
        });
        self.redo_stack.clear();

        if self.undo_stack.len() > UNDO_STACK_LIMIT {
            let overflow = self.undo_stack.len() - UNDO_STACK_LIMIT;
            self.undo_stack.drain(0..overflow);
        }
    }

    fn undo_last_change(&mut self) {
        let Some(previous) = self.undo_stack.pop() else {
            self.status = "元に戻せる変更がありません。".to_string();
            return;
        };

        let Some(current_document) = self.document.as_ref().cloned() else {
            return;
        };
        self.redo_stack.push(UndoState {
            document: current_document,
            selected_task: self.selected_task,
            selected_sheet_column: self.selected_sheet_column,
            sheet_selection_anchor: self.sheet_selection_anchor,
            dirty: self.dirty,
        });
        if self.redo_stack.len() > UNDO_STACK_LIMIT {
            let overflow = self.redo_stack.len() - UNDO_STACK_LIMIT;
            self.redo_stack.drain(0..overflow);
        }

        self.document = Some(previous.document);
        self.selected_task = previous.selected_task;
        self.selected_sheet_column = previous.selected_sheet_column;
        self.sheet_selection_anchor = previous.sheet_selection_anchor;
        self.sheet_editing_cell = None;
        self.dirty = previous.dirty;
        self.chart_interaction = None;
        self.dependency_picker = None;
        self.status = "元に戻しました。".to_string();
        self.sync_task_editor_with_selection();
    }

    fn redo_last_change(&mut self) {
        let Some(next) = self.redo_stack.pop() else {
            self.status = "やり直せる変更がありません。".to_string();
            return;
        };

        let Some(current_document) = self.document.as_ref().cloned() else {
            return;
        };
        self.undo_stack.push(UndoState {
            document: current_document,
            selected_task: self.selected_task,
            selected_sheet_column: self.selected_sheet_column,
            sheet_selection_anchor: self.sheet_selection_anchor,
            dirty: self.dirty,
        });
        if self.undo_stack.len() > UNDO_STACK_LIMIT {
            let overflow = self.undo_stack.len() - UNDO_STACK_LIMIT;
            self.undo_stack.drain(0..overflow);
        }

        self.document = Some(next.document);
        self.selected_task = next.selected_task;
        self.selected_sheet_column = next.selected_sheet_column;
        self.sheet_selection_anchor = next.sheet_selection_anchor;
        self.sheet_editing_cell = None;
        self.dirty = next.dirty;
        self.chart_interaction = None;
        self.dependency_picker = None;
        self.status = "やり直しました。".to_string();
        self.sync_task_editor_with_selection();
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

    fn set_status_date(&mut self, date: NaiveDate) {
        self.status_date = date;
        self.status_date_text = date.format("%Y-%m-%d").to_string();
        self.status_date_picker_month = first_day_of_month(date);
        self.save_app_settings();
    }

    fn show_status_date_picker(&mut self, ctx: &egui::Context) {
        if !self.status_date_picker_open {
            return;
        }

        let mut open = self.status_date_picker_open;
        let mut month = self.status_date_picker_month;
        let mut selected_date = None;
        let mut close_requested = false;

        egui::Window::new("基準日")
            .collapsible(false)
            .resizable(false)
            .default_pos(Pos2::new(860.0, 86.0))
            .open(&mut open)
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    if ui.small_button("◀").clicked() {
                        month = prev_month_start(month);
                    }
                    ui.label(RichText::new(status_date_picker_label(month)).strong());
                    if ui.small_button("▶").clicked() {
                        month = next_month_start(month);
                    }
                });

                ui.add_space(4.0);
                egui::Grid::new("status_date_picker_weekdays")
                    .spacing(Vec2::new(4.0, 4.0))
                    .show(ui, |ui| {
                        for (index, label) in ["月", "火", "水", "木", "金", "土", "日"]
                            .into_iter()
                            .enumerate()
                        {
                            let color = if index == 5 {
                                Color32::from_rgb(52, 94, 176)
                            } else if index == 6 {
                                Color32::from_rgb(192, 64, 48)
                            } else {
                                SUBTEXT
                            };
                            ui.label(RichText::new(label).size(10.0).color(color));
                        }
                        ui.end_row();
                    });

                let first_weekday = month.weekday().num_days_from_monday() as usize;
                let month_days = days_in_month(month.year(), month.month());
                let today = chrono::Local::now().date_naive();

                egui::Grid::new("status_date_picker_days")
                    .spacing(Vec2::new(4.0, 4.0))
                    .show(ui, |ui| {
                        for _ in 0..first_weekday {
                            ui.label(" ");
                        }

                        for day in 1..=month_days {
                            let date = NaiveDate::from_ymd_opt(month.year(), month.month(), day)
                                .unwrap_or(month);
                            let is_selected = date == self.status_date;
                            let is_today = date == today;
                            let weekday = date.weekday().num_days_from_monday();
                            let weekend_color = if weekday == 5 {
                                Color32::from_rgb(52, 94, 176)
                            } else if weekday == 6 {
                                Color32::from_rgb(192, 64, 48)
                            } else {
                                TEXT
                            };
                            let text = if is_selected {
                                RichText::new(day.to_string()).strong().color(Color32::WHITE)
                            } else if is_today {
                                RichText::new(day.to_string()).strong().color(TODAY)
                            } else {
                                RichText::new(day.to_string()).color(weekend_color)
                            };
                            let mut button = egui::Button::new(text).min_size(Vec2::new(30.0, 24.0));
                            if is_selected {
                                button = button.fill(ACCENT);
                            } else if is_today {
                                button = button.fill(Color32::from_rgb(255, 242, 218));
                            } else if weekday == 5 || weekday == 6 {
                                button = button.fill(Color32::from_rgb(250, 250, 252));
                            }
                            if ui.add(button).clicked() {
                                selected_date = Some(date);
                            }
                            if (first_weekday + day as usize) % 7 == 0 {
                                ui.end_row();
                            }
                        }
                    });

                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    if ui.button("今日").clicked() {
                        selected_date = Some(chrono::Local::now().date_naive());
                    }
                    if ui.button("閉じる").clicked() {
                        close_requested = true;
                    }
                });
            });

        self.status_date_picker_open = open && !close_requested;
        self.status_date_picker_month = month;
        if let Some(date) = selected_date {
            self.set_status_date(date);
            self.status_date_picker_open = false;
        }
    }

    fn set_sheet_selection(
        &mut self,
        task_index: usize,
        column: SheetColumn,
        extend_selection: bool,
    ) {
        let previous_selection = self.selected_sheet_cell();
        self.selected_task = Some(task_index);
        self.selected_sheet_column = column;
        self.sheet_editing_cell = None;
        if extend_selection {
            if self.sheet_selection_anchor.is_none() {
                self.sheet_selection_anchor =
                    previous_selection.or(Some(SheetCellSelection { task_index, column }));
            }
        } else {
            self.sheet_selection_anchor = Some(SheetCellSelection { task_index, column });
        }
    }

    fn selected_sheet_cell(&self) -> Option<SheetCellSelection> {
        self.selected_task.map(|task_index| SheetCellSelection {
            task_index,
            column: self.selected_sheet_column,
        })
    }

    fn begin_sheet_editing(&mut self) {
        self.sheet_editing_cell = self.selected_sheet_cell();
    }

    fn cancel_sheet_editing(&mut self) {
        self.sheet_editing_cell = None;
    }

    fn apply_task_editor_to_selection(&mut self) {
        let Some(index) = self.selected_task else {
            return;
        };

        let Some(document_snapshot) = self.document.as_ref().cloned() else {
            return;
        };
        let selected_task = self.selected_task;
        let selected_sheet_column = self.selected_sheet_column;
        let sheet_selection_anchor = self.sheet_selection_anchor;
        let dirty_before = self.dirty;

        let changed = {
            let Some(document) = self.document.as_mut() else {
                return;
            };
            apply_task_editor(document, index, &self.task_editor)
        };

        if changed {
            self.record_undo_state_from(
                document_snapshot,
                selected_task,
                selected_sheet_column,
                sheet_selection_anchor,
                dirty_before,
            );
            if let Some(document) = self.document.as_ref() {
                if let Some(task) = document.tasks.get(index) {
                    self.task_editor.sync_from_task(task);
                    self.status = format!("タスクを更新しました: '{}'", task.name);
                    self.dirty = true;
                }
            }
        }
    }

    fn collapse_all_summary_tasks(&mut self) {
        if let Some(document) = self.document.as_ref() {
            self.collapsed_task_uids = document
                .tasks
                .iter()
                .filter(|task| task.summary)
                .map(|task| task.uid)
                .collect();
        }
    }

    fn expand_all_summary_tasks(&mut self) {
        self.collapsed_task_uids.clear();
    }

    fn visible_shortcut_open(ctx: &egui::Context) -> bool {
        ctx.input(|input| input.key_pressed(egui::Key::O) && input.modifiers.command)
    }

    fn visible_shortcut_save(ctx: &egui::Context) -> bool {
        ctx.input(|input| {
            input.key_pressed(egui::Key::S) && input.modifiers.command && !input.modifiers.shift
        })
    }

    fn visible_shortcut_save_as(ctx: &egui::Context) -> bool {
        ctx.input(|input| {
            input.key_pressed(egui::Key::S) && input.modifiers.command && input.modifiers.shift
        })
    }

    fn visible_shortcut_undo(ctx: &egui::Context) -> bool {
        ctx.input(|input| {
            input.key_pressed(egui::Key::Z) && input.modifiers.command && !input.modifiers.shift
        })
    }

    fn visible_shortcut_redo(ctx: &egui::Context) -> bool {
        ctx.input(|input| {
            (input.key_pressed(egui::Key::Y) && input.modifiers.command)
                || (input.key_pressed(egui::Key::Z)
                    && input.modifiers.command
                    && input.modifiers.shift)
        })
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

    fn handle_sheet_clipboard(&mut self, ctx: &egui::Context, visible_indices: &[usize]) {
        if self.document.is_none() || visible_indices.is_empty() {
            return;
        }

        let editing_sheet_cell = self.sheet_editing_cell.is_some();
        let wants_keyboard_input = ctx.egui_wants_keyboard_input();
        let mut pasted_blocks = Vec::new();
        let mut copy_requested = false;

        ctx.input_mut(|input| {
            input.events.retain(|event| match event {
                Event::Copy if !editing_sheet_cell || !wants_keyboard_input => {
                    copy_requested = true;
                    false
                }
                Event::Paste(text) if !editing_sheet_cell || !wants_keyboard_input => {
                    pasted_blocks.push(text.clone());
                    false
                }
                _ => true,
            });
        });

        if copy_requested {
            if let Some(text) = self.copy_sheet_selection_to_tsv(visible_indices) {
                ctx.copy_text(text);
                self.status = "選択範囲をコピーしました。".to_string();
            }
        }

        for text in pasted_blocks {
            self.apply_excel_paste(text, visible_indices);
        }
    }

    fn apply_excel_paste(&mut self, text: String, visible_indices: &[usize]) {
        let Some((start_row, start_column)) = self.paste_start_cell(visible_indices) else {
            return;
        };

        let Some(document_snapshot) = self.document.as_ref().cloned() else {
            return;
        };
        let selected_task = self.selected_task;
        let selected_sheet_column = self.selected_sheet_column;
        let sheet_selection_anchor = self.sheet_selection_anchor;
        let dirty_before = self.dirty;

        let mut row_count = 0usize;
        let mut changed = false;
        {
            let Some(document) = self.document.as_mut() else {
                return;
            };
            let normalized = Self::normalize_excel_clipboard_text(&text);
            let rows = Self::parse_tsv_rows(&normalized);
            if rows.is_empty() {
                return;
            }

            let starts_with_mode = rows
                .first()
                .and_then(|row| row.first())
                .map(|value| Self::is_mode_label(value.trim()))
                .unwrap_or(false);
            let source_offset = if starts_with_mode { 1 } else { 0 };

            for (offset, row) in rows.iter().enumerate() {
                let task_position = start_row + offset;
                let Some(&task_index) = visible_indices.get(task_position) else {
                    break;
                };

                if row.is_empty() || source_offset >= row.len() {
                    continue;
                }

                if Self::apply_excel_row(document, task_index, start_column, &row[source_offset..])
                {
                    changed = true;
                    row_count += 1;
                }
            }
        }

        if changed {
            self.record_undo_state_from(
                document_snapshot,
                selected_task,
                selected_sheet_column,
                sheet_selection_anchor,
                dirty_before,
            );
            self.dirty = true;
            if let Some(index) = self.selected_task {
                if let Some(document) = self.document.as_ref() {
                    if let Some(task) = document.tasks.get(index) {
                        self.task_editor.sync_from_task(task);
                    }
                }
            }
            self.sheet_selection_anchor = self.selected_sheet_cell();
            self.status = if row_count == 1 {
                "Excel の貼り付けを反映しました。".to_string()
            } else {
                format!("Excel の貼り付けを {} 行反映しました。", row_count)
            };
        }
    }

    fn apply_excel_row(
        document: &mut ProjectDocument,
        index: usize,
        start_column: SheetColumn,
        cells: &[&str],
    ) -> bool {
        if document.tasks.get(index).is_none() {
            return false;
        }

        let mut changed = false;
        for (offset, value) in cells.iter().enumerate() {
            let Some(column) = SheetColumn::from_index(start_column.index() + offset) else {
                break;
            };
            if !column.is_editable() {
                continue;
            }
            if Self::apply_sheet_cell_value(document, index, column, value.trim()) {
                changed = true;
            }
        }

        changed
    }

    fn apply_sheet_cell_value(
        document: &mut ProjectDocument,
        index: usize,
        column: SheetColumn,
        value: &str,
    ) -> bool {
        match column {
            SheetColumn::Mode => false,
            SheetColumn::Name => {
                let Some(task) = document.tasks.get_mut(index) else {
                    return false;
                };
                let normalized = value.trim().to_string();
                if task.name != normalized {
                    task.name = normalized;
                    return true;
                }
                false
            }
            SheetColumn::Duration => {
                let Some(task) = document.tasks.get_mut(index) else {
                    return false;
                };
                let normalized = normalize_duration_text(value);
                if task.duration_text != normalized {
                    task.duration_text = normalized;
                    return true;
                }
                false
            }
            SheetColumn::Start => {
                let Some(task) = document.tasks.get_mut(index) else {
                    return false;
                };
                let normalized = value.trim().to_string();
                let parsed = parse_date_time(Some(&normalized));
                if task.start_text != normalized || task.start != parsed {
                    task.start_text = normalized;
                    task.start = parsed;
                    return true;
                }
                false
            }
            SheetColumn::Finish => {
                let Some(task) = document.tasks.get_mut(index) else {
                    return false;
                };
                let normalized = value.trim().to_string();
                let parsed = parse_date_time(Some(&normalized));
                if task.finish_text != normalized || task.finish != parsed {
                    task.finish_text = normalized;
                    task.finish = parsed;
                    return true;
                }
                false
            }
            SheetColumn::PercentComplete => {
                let Some(task) = document.tasks.get_mut(index) else {
                    return false;
                };
                let cleaned = value.trim().trim_end_matches('%').trim();
                if let Ok(percent) = cleaned.parse::<f32>() {
                    let percent = percent.clamp(0.0, 100.0);
                    if (task.percent_complete - percent).abs() > f32::EPSILON {
                        task.percent_complete = percent;
                        return true;
                    }
                }
                false
            }
            SheetColumn::Predecessors => Self::apply_excel_predecessors(document, index, value),
        }
    }

    fn apply_excel_predecessors(document: &mut ProjectDocument, index: usize, value: &str) -> bool {
        let Some(task_uid) = document.tasks.get(index).map(|task| task.uid) else {
            return false;
        };
        let trimmed = value.trim();
        if trimmed.is_empty() {
            let previous = document.dependencies.len();
            document
                .dependencies
                .retain(|dep| dep.successor_uid != task_uid);
            if let Some(task) = document.tasks.get_mut(index) {
                if !task.predecessor_text.is_empty() {
                    task.predecessor_text.clear();
                }
            }
            return previous != document.dependencies.len();
        }

        let previous_text = document
            .tasks
            .get(index)
            .map(|task| task.predecessor_text.clone())
            .unwrap_or_default();
        let mut text_changed = false;
        if let Some(task) = document.tasks.get_mut(index) {
            if task.predecessor_text != trimmed {
                task.predecessor_text = trimmed.to_string();
                text_changed = true;
            }
        }

        let parsed = parse_predecessor_assignments(trimmed);
        if parsed.is_empty() {
            return text_changed;
        }

        let mut resolved_dependencies = Vec::new();
        let mut seen_predecessors = HashSet::new();

        for assignment in parsed {
            let Some(predecessor_uid) =
                find_task_by_id(&document.tasks, assignment.predecessor_id).map(|task| task.uid)
            else {
                continue;
            };
            if !seen_predecessors.insert(predecessor_uid) {
                continue;
            }
            resolved_dependencies.push(GanttDependency {
                predecessor_uid,
                successor_uid: task_uid,
                relation: assignment.relation,
                lag_text: assignment.lag_text,
            });
        }

        if resolved_dependencies.is_empty() {
            return false;
        }

        document
            .dependencies
            .retain(|dep| dep.successor_uid != task_uid);
        document.dependencies.extend(resolved_dependencies);
        refresh_predecessor_text_for_task(document, task_uid);
        if let Some(task) = document.tasks.get(index) {
            if task.predecessor_text != previous_text {
                text_changed = true;
            }
        }

        text_changed
    }

    fn handle_spreadsheet_navigation(&mut self, ctx: &egui::Context, visible_indices: &[usize]) {
        if self.document.is_none() || visible_indices.is_empty() {
            return;
        }

        let wants_keyboard_input = ctx.egui_wants_keyboard_input();
        let mut target_row = self.visible_row_index(visible_indices).unwrap_or(0);
        let mut target_col = self.selected_sheet_column;

        let input = ctx.input(|input| {
            (
                input.key_pressed(egui::Key::ArrowLeft),
                input.key_pressed(egui::Key::ArrowRight),
                input.key_pressed(egui::Key::ArrowUp),
                input.key_pressed(egui::Key::ArrowDown),
                input.key_pressed(egui::Key::Tab),
                input.modifiers.shift,
                input.key_pressed(egui::Key::Enter),
                input.key_pressed(egui::Key::F2),
                input.key_pressed(egui::Key::Escape),
                input.modifiers.command,
            )
        });
        let shift = input.5;
        let ctrl = input.9;

        if input.8 {
            self.cancel_sheet_editing();
            ctx.request_repaint();
            return;
        }

        if input.7 {
            if self.selected_task.is_some() {
                self.begin_sheet_editing();
                ctx.request_repaint();
            }
            return;
        }

        if wants_keyboard_input && !input.4 && !input.6 {
            return;
        }

        let navigation = if input.0 {
            Some(SheetNavigationDirection::Left)
        } else if input.1 {
            Some(SheetNavigationDirection::Right)
        } else if input.2 {
            Some(SheetNavigationDirection::Up)
        } else if input.3 {
            Some(SheetNavigationDirection::Down)
        } else {
            None
        };

        let Some(direction) = navigation else {
            return;
        };

        let (next_row, next_col) = Self::sheet_navigation_target(
            target_row,
            target_col,
            visible_indices.len(),
            direction,
            ctrl,
        );

        let moved = next_row != target_row || next_col != target_col;
        if moved {
            target_row = next_row;
            target_col = next_col;
            self.set_sheet_selection(visible_indices[target_row], target_col, shift);
            ctx.request_repaint();
        }
    }

    fn visible_row_index(&self, visible_indices: &[usize]) -> Option<usize> {
        self.selected_task
            .and_then(|selected| visible_indices.iter().position(|index| *index == selected))
    }

    fn sheet_navigation_target(
        current_row: usize,
        current_col: SheetColumn,
        visible_len: usize,
        direction: SheetNavigationDirection,
        ctrl: bool,
    ) -> (usize, SheetColumn) {
        let last_row = visible_len.saturating_sub(1);

        match direction {
            SheetNavigationDirection::Left => {
                let next_col = if ctrl {
                    SheetColumn::first()
                } else {
                    current_col.prev()
                };
                (current_row, next_col)
            }
            SheetNavigationDirection::Right => {
                let next_col = if ctrl {
                    SheetColumn::last()
                } else {
                    current_col.next()
                };
                (current_row, next_col)
            }
            SheetNavigationDirection::Up => {
                let next_row = if ctrl {
                    0
                } else {
                    current_row.saturating_sub(1)
                };
                (next_row, current_col)
            }
            SheetNavigationDirection::Down => {
                let next_row = if ctrl {
                    last_row
                } else {
                    current_row.saturating_add(1).min(last_row)
                };
                (next_row, current_col)
            }
        }
    }

    fn title_bar(&mut self, ui: &mut egui::Ui) {
        let ctx = ui.ctx().clone();
        egui::Frame::NONE
            .fill(TITLE_BAR_BG)
            .stroke(Stroke::new(1.0, TITLE_BAR_BG_DARK))
            .inner_margin(egui::Margin::symmetric(8, 2))
            .show(ui, |ui| {
                ui.set_height(24.0);
                ui.horizontal(|ui| {
                    let icon_rect = Rect::from_min_size(ui.cursor().min, Vec2::new(16.0, 16.0));
                    let icon_painter = ui.painter_at(icon_rect);
                    icon_painter.rect_filled(icon_rect, 2.0, Color32::from_rgb(232, 236, 248));
                    icon_painter.rect_stroke(
                        icon_rect,
                        2.0,
                        Stroke::new(1.0, Color32::from_rgb(158, 170, 214)),
                        egui::StrokeKind::Inside,
                    );
                    icon_painter.line_segment(
                        [
                            Pos2::new(icon_rect.left() + 3.0, icon_rect.top() + 4.0),
                            Pos2::new(icon_rect.right() - 3.0, icon_rect.top() + 4.0),
                        ],
                        Stroke::new(1.2, Color32::from_rgb(58, 94, 156)),
                    );
                    icon_painter.line_segment(
                        [
                            Pos2::new(icon_rect.left() + 3.0, icon_rect.center().y),
                            Pos2::new(icon_rect.right() - 3.0, icon_rect.center().y),
                        ],
                        Stroke::new(1.2, Color32::from_rgb(58, 94, 156)),
                    );
                    icon_painter.line_segment(
                        [
                            Pos2::new(icon_rect.left() + 3.0, icon_rect.bottom() - 4.0),
                            Pos2::new(icon_rect.right() - 5.0, icon_rect.bottom() - 4.0),
                        ],
                        Stroke::new(1.2, Color32::from_rgb(58, 94, 156)),
                    );
                    ui.add_space(22.0);
                    ui.label(
                        RichText::new(
                            self.document_path
                                .as_ref()
                                .and_then(|path| path.file_stem())
                                .and_then(|name| name.to_str())
                                .unwrap_or("MicroProject"),
                        )
                        .font(FontId::new(12.5, FontFamily::Proportional))
                        .strong()
                        .color(TITLE_BAR_TEXT),
                    );
                    if let Some(path) = &self.document_path {
                        ui.add_space(8.0);
                        ui.label(
                            RichText::new(path.display().to_string())
                                .size(10.0)
                                .color(Color32::from_rgb(224, 231, 245)),
                        );
                    }

                    let drag_width = (ui.available_width() - 140.0).max(120.0);
                    ui.add_sized(
                        [drag_width, 22.0],
                        egui::Label::new(RichText::new("")).sense(Sense::hover()),
                    );

                    ui.with_layout(egui::Layout::right_to_left(Align::Center), |ui| {
                        let maximized =
                            ui.input(|input| input.viewport().maximized.unwrap_or(false));
                        if self.window_button(ui, "×", true).clicked() {
                            ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                        }
                        if self
                            .window_button(ui, if maximized { "❐" } else { "▢" }, false)
                            .clicked()
                        {
                            ctx.send_viewport_cmd(egui::ViewportCommand::Maximized(!maximized));
                        }
                        if self.window_button(ui, "—", false).clicked() {
                            ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(true));
                        }
                        ui.add_space(10.0);
                        ui.label(RichText::new("Ready").size(10.0).color(TITLE_BAR_TEXT));
                    });
                });
            });
    }

    fn window_button(&self, ui: &mut egui::Ui, text: &str, destructive: bool) -> egui::Response {
        let desired = Vec2::new(18.0, 16.0);
        let (rect, response) = ui.allocate_exact_size(desired, Sense::click());
        if ui.is_rect_visible(rect) {
            let painter = ui.painter_at(rect);
            let bg = if response.hovered() {
                if destructive {
                    Color32::from_rgb(164, 38, 38)
                } else {
                    Color32::from_rgb(65, 98, 166)
                }
            } else {
                Color32::from_rgba_unmultiplied(255, 255, 255, 22)
            };
            painter.rect_filled(rect, 1.0, bg);
            painter.rect_stroke(
                rect,
                1.0,
                Stroke::new(1.0, Color32::from_rgba_unmultiplied(255, 255, 255, 72)),
                egui::StrokeKind::Inside,
            );
            painter.text(
                rect.center(),
                egui::Align2::CENTER_CENTER,
                text,
                FontId::new(11.0, FontFamily::Proportional),
                TITLE_BAR_TEXT,
            );
        }
        response
    }

    fn ribbon(&mut self, ui: &mut egui::Ui) {
        egui::Frame::NONE
            .fill(RIBBON_CHROME_BG)
            .stroke(Stroke::new(1.0, BORDER))
            .show(ui, |ui| {
                ui.add_space(1.0);
                ui.horizontal(|ui| {
                    ui.add_space(4.0);
                    egui::MenuBar::new().ui(ui, |ui| {
                        ui.menu_button("ファイル", |ui| {
                            if ui.button("開く...").clicked() {
                                self.load_dialog();
                                ui.close();
                            }
                            if ui.button("保存").clicked() {
                                self.save_current_file();
                                ui.close();
                            }
                            if ui.button("名前を付けて保存").clicked() {
                                self.save_file_as();
                                ui.close();
                            }
                            ui.separator();
                            if ui.button("最近使ったファイル").clicked() {
                                if let Some(path) = self.recent_files.first().cloned() {
                                    self.load_file(path);
                                }
                                ui.close();
                            }
                        });
                        ui.menu_button("編集", |ui| {
                            if ui.button("元に戻す").clicked() {
                                self.undo_last_change();
                                ui.close();
                            }
                            if ui.button("やり直す").clicked() {
                                self.redo_last_change();
                                ui.close();
                            }
                        });
                        ui.menu_button("表示", |ui| {
                            if ui.button("拡大").clicked() {
                                self.zoom_px_per_day = (self.zoom_px_per_day + 2.0).min(40.0);
                                ui.close();
                            }
                            if ui.button("縮小").clicked() {
                                self.zoom_px_per_day = (self.zoom_px_per_day - 2.0).max(8.0);
                                ui.close();
                            }
                            ui.separator();
                            if ui.checkbox(&mut self.show_weekends, "週末を表示").changed() {
                                ui.close();
                            }
                            if ui
                                .checkbox(&mut self.show_critical_path, "クリティカル表示")
                                .changed()
                            {
                                ui.close();
                            }
                        });
                        ui.menu_button("プロジェクト", |ui| {
                            if ui.button("要約を展開").clicked() {
                                self.expand_all_summary_tasks();
                                ui.close();
                            }
                            if ui.button("要約を折りたたむ").clicked() {
                                self.collapse_all_summary_tasks();
                                ui.close();
                            }
                            if ui.button("選択を反映").clicked() {
                                self.apply_task_editor_to_selection();
                                ui.close();
                            }
                        });
                    });

                    ui.add_space(10.0);
                    ui.separator();
                    ui.add_space(6.0);

                    if Self::classic_toolbar_button(ui, RibbonIcon::Open, "開く").clicked() {
                        self.load_dialog();
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::Save, "保存").clicked() {
                        self.save_current_file();
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::Undo, "元に戻す").clicked()
                    {
                        self.undo_last_change();
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::Redo, "やり直す").clicked()
                    {
                        self.redo_last_change();
                    }
                    ui.separator();
                    if Self::classic_toolbar_button(ui, RibbonIcon::ExpandAll, "展開").clicked() {
                        self.expand_all_summary_tasks();
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::CollapseAll, "折りたたむ")
                        .clicked()
                    {
                        self.collapse_all_summary_tasks();
                    }
                    ui.separator();
                    if Self::classic_toolbar_button(ui, RibbonIcon::ZoomOut, "縮小").clicked() {
                        self.zoom_px_per_day = (self.zoom_px_per_day - 2.0).max(8.0);
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::ZoomIn, "拡大").clicked() {
                        self.zoom_px_per_day = (self.zoom_px_per_day + 2.0).min(40.0);
                    }
                    ui.separator();
                    if Self::classic_toolbar_button(ui, RibbonIcon::TaskEdit, "適用").clicked() {
                        self.apply_task_editor_to_selection();
                    }
                    if Self::classic_toolbar_button(ui, RibbonIcon::Filter, "表示切替").clicked()
                    {
                        self.show_critical_path = !self.show_critical_path;
                    }

                    ui.with_layout(egui::Layout::right_to_left(Align::Center), |ui| {
                        let label = self.status_date.format("%Y-%m-%d").to_string();
                        if ui.button(format!("📅 {}", label)).clicked() {
                            self.status_date_picker_open = true;
                            self.status_date_picker_month = first_day_of_month(self.status_date);
                        }
                        ui.label(RichText::new("基準日").size(10.0).color(SUBTEXT));
                        ui.add_space(8.0);
                        ui.add_sized(
                            [162.0, 20.0],
                            egui::TextEdit::singleline(&mut self.filter_text)
                                .hint_text("検索/フィルタ"),
                        );
                        ui.add_space(8.0);
                        ui.label(RichText::new("フィルタ").size(10.0).color(SUBTEXT));
                    });
                });
            });
        self.show_status_date_picker(ui.ctx());
    }

    fn classic_toolbar_button(ui: &mut egui::Ui, icon: RibbonIcon, label: &str) -> egui::Response {
        let desired = Vec2::new(22.0, 22.0);
        let (rect, response) = ui.allocate_exact_size(desired, Sense::click());
        if ui.is_rect_visible(rect) {
            let painter = ui.painter_at(rect);
            let bg = if response.hovered() {
                Color32::from_rgb(251, 243, 189)
            } else {
                Color32::from_rgb(236, 233, 216)
            };
            painter.rect_filled(rect, 0.0, bg);
            painter.rect_stroke(
                rect,
                0.0,
                Stroke::new(
                    1.0,
                    if response.hovered() {
                        Color32::from_rgb(218, 160, 59)
                    } else {
                        Color32::from_rgb(160, 160, 160)
                    },
                ),
                egui::StrokeKind::Inside,
            );
            let icon_rect = Rect::from_min_size(
                Pos2::new(rect.center().x - 7.0, rect.center().y - 7.0),
                Vec2::new(14.0, 14.0),
            );
            Self::draw_colibre_icon(&painter, icon_rect, icon);
        }
        response.on_hover_text(label)
    }

    fn ribbon_group(ui: &mut egui::Ui, title: &str, add: impl FnOnce(&mut egui::Ui)) {
        egui::Frame::group(ui.style())
            .fill(Color32::from_rgb(255, 255, 255))
            .stroke(Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(5, 3))
            .show(ui, |ui| {
                ui.set_min_width(132.0);
                ui.vertical(|ui| {
                    if !title.is_empty() {
                        ui.label(
                            RichText::new(title)
                                .size(10.5)
                                .strong()
                                .color(Color32::from_rgb(69, 69, 69)),
                        );
                        ui.add_space(2.0);
                    }
                    ui.horizontal_wrapped(|ui| {
                        add(ui);
                    });
                });
            });
    }

    fn ribbon_style_tile(
        ui: &mut egui::Ui,
        colors: (Color32, Color32, Color32),
        selected: bool,
    ) -> egui::Response {
        let desired = Vec2::new(34.0, 39.0);
        let (rect, response) = ui.allocate_exact_size(desired, Sense::click());
        if ui.is_rect_visible(rect) {
            let painter = ui.painter_at(rect);
            let bg = if selected {
                Color32::from_rgb(252, 253, 255)
            } else {
                Color32::from_rgb(248, 249, 252)
            };
            painter.rect_filled(rect, 3.0, bg);
            painter.rect_stroke(
                rect,
                3.0,
                Stroke::new(
                    1.0,
                    if response.hovered() {
                        ACCENT
                    } else if selected {
                        Color32::from_rgb(168, 189, 212)
                    } else {
                        BORDER
                    },
                ),
                egui::StrokeKind::Inside,
            );
            let top = rect.top() + 7.0;
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 6.0, top),
                    Pos2::new(rect.right() - 6.0, top),
                ],
                Stroke::new(1.5, colors.1),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 6.0, top + 6.5),
                    Pos2::new(rect.right() - 8.0, top + 6.5),
                ],
                Stroke::new(4.0, colors.0),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 6.0, top + 14.0),
                    Pos2::new(rect.right() - 12.0, top + 14.0),
                ],
                Stroke::new(1.4, colors.2),
            );
            painter.rect_filled(
                Rect::from_center_size(
                    Pos2::new(rect.left() + 20.0, rect.bottom() - 9.0),
                    Vec2::new(8.0, 8.0),
                ),
                0.0,
                colors.2,
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 16.0, rect.bottom() - 13.0),
                    Pos2::new(rect.left() + 24.0, rect.bottom() - 5.0),
                ],
                Stroke::new(1.1, colors.1),
            );
        }
        response.on_hover_text("Gantt style")
    }

    fn ribbon_icon_button(ui: &mut egui::Ui, icon: RibbonIcon, label: &str) -> egui::Response {
        let desired = Vec2::new(66.0, 54.0);
        let (rect, response) = ui.allocate_exact_size(desired, Sense::click());
        if ui.is_rect_visible(rect) {
            let painter = ui.painter_at(rect);
            let bg = if response.hovered() {
                Color32::from_rgb(244, 248, 253)
            } else {
                Color32::from_rgb(249, 251, 254)
            };
            painter.rect_filled(rect, 4.0, bg);
            painter.rect_stroke(
                rect,
                4.0,
                Stroke::new(1.0, if response.hovered() { ACCENT } else { BORDER }),
                egui::StrokeKind::Inside,
            );

            let icon_rect = Rect::from_center_size(
                Pos2::new(rect.center().x, rect.top() + 18.0),
                Vec2::new(20.0, 20.0),
            );
            Self::draw_colibre_icon(&painter, icon_rect, icon);
        }
        response.on_hover_text(label)
    }

    fn ribbon_small_button(ui: &mut egui::Ui, icon: RibbonIcon, label: &str) -> egui::Response {
        let desired = Vec2::new(26.0, 26.0);
        let (rect, response) = ui.allocate_exact_size(desired, Sense::click());
        if ui.is_rect_visible(rect) {
            let painter = ui.painter_at(rect);
            let bg = if response.hovered() {
                Color32::from_rgb(244, 248, 253)
            } else {
                Color32::from_rgb(251, 252, 254)
            };
            painter.rect_filled(rect, 4.0, bg);
            painter.rect_stroke(
                rect,
                4.0,
                Stroke::new(1.0, if response.hovered() { ACCENT } else { BORDER }),
                egui::StrokeKind::Inside,
            );

            let icon_rect = Rect::from_min_size(
                Pos2::new(rect.center().x - 7.0, rect.center().y - 7.0),
                Vec2::new(14.0, 14.0),
            );
            Self::draw_colibre_icon(&painter, icon_rect, icon);
        }
        response.on_hover_text(label)
    }

    fn draw_colibre_icon(painter: &egui::Painter, rect: Rect, icon: RibbonIcon) {
        let blue = Color32::from_rgb(54, 121, 83);
        let blue_dark = Color32::from_rgb(34, 89, 59);
        let green = Color32::from_rgb(55, 145, 78);
        let orange = Color32::from_rgb(230, 145, 54);
        let gray = Color32::from_rgb(112, 124, 140);
        let x1 = rect.left();
        let y1 = rect.top();
        let x2 = rect.right();
        let y2 = rect.bottom();
        let cx = rect.center().x;
        let cy = rect.center().y;

        match icon {
            RibbonIcon::Open => {
                let folder = Rect::from_min_max(
                    Pos2::new(x1 + 2.0, y1 + 6.5),
                    Pos2::new(x2 - 2.0, y2 - 2.0),
                );
                painter.rect_filled(folder, 2.0, blue);
                painter.rect_filled(
                    Rect::from_min_max(
                        Pos2::new(x1 + 2.0, y1 + 4.0),
                        Pos2::new(cx - 1.0, y1 + 8.5),
                    ),
                    1.5,
                    Color32::from_rgb(177, 205, 238),
                );
                painter.line_segment(
                    [Pos2::new(x1 + 6.0, cy), Pos2::new(cx + 3.0, cy)],
                    Stroke::new(1.7, Color32::WHITE),
                );
                painter.line_segment(
                    [Pos2::new(cx + 0.5, cy - 3.5), Pos2::new(cx + 3.5, cy)],
                    Stroke::new(1.7, Color32::WHITE),
                );
                painter.line_segment(
                    [Pos2::new(cx + 0.5, cy + 3.5), Pos2::new(cx + 3.5, cy)],
                    Stroke::new(1.7, Color32::WHITE),
                );
            }
            RibbonIcon::Save => {
                let body = Rect::from_min_max(
                    Pos2::new(x1 + 3.0, y1 + 2.0),
                    Pos2::new(x2 - 3.0, y2 - 2.0),
                );
                painter.rect_filled(body, 2.0, blue_dark);
                painter.rect_filled(
                    Rect::from_min_max(
                        Pos2::new(x1 + 6.0, y1 + 4.0),
                        Pos2::new(x2 - 6.0, y1 + 8.0),
                    ),
                    1.0,
                    Color32::from_rgb(176, 201, 232),
                );
                painter.rect_filled(
                    Rect::from_min_max(
                        Pos2::new(x1 + 7.0, y1 + 11.0),
                        Pos2::new(x2 - 7.0, y2 - 5.0),
                    ),
                    1.5,
                    Color32::from_rgb(239, 244, 250),
                );
                painter.line_segment(
                    [
                        Pos2::new(cx - 0.5, y1 + 12.0),
                        Pos2::new(cx - 0.5, y2 - 6.0),
                    ],
                    Stroke::new(1.6, blue),
                );
                painter.line_segment(
                    [Pos2::new(cx - 3.5, y2 - 9.0), Pos2::new(cx - 0.5, y2 - 6.0)],
                    Stroke::new(1.6, blue),
                );
                painter.line_segment(
                    [Pos2::new(cx + 2.5, y2 - 9.0), Pos2::new(cx - 0.5, y2 - 6.0)],
                    Stroke::new(1.6, blue),
                );
            }
            RibbonIcon::Undo => {
                painter.circle_stroke(rect.center(), 7.5, Stroke::new(1.5, orange));
                painter.line_segment(
                    [Pos2::new(cx - 2.0, cy - 4.0), Pos2::new(cx - 6.0, cy - 1.0)],
                    Stroke::new(1.6, orange),
                );
                painter.line_segment(
                    [Pos2::new(cx - 6.0, cy - 1.0), Pos2::new(cx - 2.0, cy + 2.0)],
                    Stroke::new(1.6, orange),
                );
                painter.line_segment(
                    [Pos2::new(cx - 2.0, cy - 1.0), Pos2::new(cx + 4.0, cy - 1.0)],
                    Stroke::new(1.6, orange),
                );
            }
            RibbonIcon::Redo => {
                painter.circle_stroke(rect.center(), 7.5, Stroke::new(1.5, green));
                painter.line_segment(
                    [Pos2::new(cx + 2.0, cy - 4.0), Pos2::new(cx + 6.0, cy - 1.0)],
                    Stroke::new(1.6, green),
                );
                painter.line_segment(
                    [Pos2::new(cx + 6.0, cy - 1.0), Pos2::new(cx + 2.0, cy + 2.0)],
                    Stroke::new(1.6, green),
                );
                painter.line_segment(
                    [Pos2::new(cx + 2.0, cy - 1.0), Pos2::new(cx - 4.0, cy - 1.0)],
                    Stroke::new(1.6, green),
                );
            }
            RibbonIcon::ZoomIn => {
                painter.circle_stroke(Pos2::new(cx - 1.5, cy - 1.5), 5.5, Stroke::new(1.5, blue));
                painter.line_segment(
                    [Pos2::new(cx + 2.2, cy + 2.2), Pos2::new(x2 - 1.5, y2 - 1.5)],
                    Stroke::new(1.7, blue),
                );
                painter.line_segment(
                    [Pos2::new(cx - 1.5, cy - 4.0), Pos2::new(cx - 1.5, cy + 1.0)],
                    Stroke::new(1.5, blue),
                );
                painter.line_segment(
                    [Pos2::new(cx - 4.0, cy - 1.5), Pos2::new(cx + 1.0, cy - 1.5)],
                    Stroke::new(1.5, blue),
                );
            }
            RibbonIcon::ZoomOut => {
                painter.circle_stroke(Pos2::new(cx - 1.5, cy - 1.5), 5.5, Stroke::new(1.5, blue));
                painter.line_segment(
                    [Pos2::new(cx + 2.2, cy + 2.2), Pos2::new(x2 - 1.5, y2 - 1.5)],
                    Stroke::new(1.7, blue),
                );
                painter.line_segment(
                    [Pos2::new(cx - 4.0, cy - 1.5), Pos2::new(cx + 1.0, cy - 1.5)],
                    Stroke::new(1.5, blue),
                );
            }
            RibbonIcon::ExpandAll => {
                painter.rect_stroke(
                    Rect::from_min_max(
                        Pos2::new(x1 + 4.0, y1 + 4.0),
                        Pos2::new(x2 - 4.0, y2 - 4.0),
                    ),
                    1.0,
                    Stroke::new(1.2, gray),
                    egui::StrokeKind::Inside,
                );
                painter.line_segment(
                    [Pos2::new(x1 + 5.0, cy), Pos2::new(x1 + 9.0, cy)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(x2 - 5.0, cy), Pos2::new(x2 - 9.0, cy)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(cx, y1 + 5.0), Pos2::new(cx, y1 + 9.0)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(cx, y2 - 5.0), Pos2::new(cx, y2 - 9.0)],
                    Stroke::new(1.5, gray),
                );
            }
            RibbonIcon::CollapseAll => {
                painter.rect_stroke(
                    Rect::from_min_max(
                        Pos2::new(x1 + 4.0, y1 + 4.0),
                        Pos2::new(x2 - 4.0, y2 - 4.0),
                    ),
                    1.0,
                    Stroke::new(1.2, gray),
                    egui::StrokeKind::Inside,
                );
                painter.line_segment(
                    [Pos2::new(x1 + 7.0, cy), Pos2::new(cx - 2.0, cy)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(x2 - 7.0, cy), Pos2::new(cx + 2.0, cy)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(cx, y1 + 7.0), Pos2::new(cx, cy - 2.0)],
                    Stroke::new(1.5, gray),
                );
                painter.line_segment(
                    [Pos2::new(cx, y2 - 7.0), Pos2::new(cx, cy + 2.0)],
                    Stroke::new(1.5, gray),
                );
            }
            RibbonIcon::Filter => {
                let points = vec![
                    Pos2::new(x1 + 4.0, y1 + 4.0),
                    Pos2::new(x2 - 4.0, y1 + 4.0),
                    Pos2::new(cx + 2.0, cy + 1.0),
                    Pos2::new(cx + 1.0, y2 - 4.0),
                    Pos2::new(cx - 1.0, y2 - 4.0),
                    Pos2::new(cx - 2.0, cy + 1.0),
                ];
                painter.add(egui::Shape::convex_polygon(
                    points,
                    blue,
                    Stroke::new(1.0, blue_dark),
                ));
            }
            RibbonIcon::TaskEdit => {
                let card = Rect::from_min_max(
                    Pos2::new(x1 + 3.0, y1 + 5.0),
                    Pos2::new(x2 - 5.0, y2 - 4.0),
                );
                painter.rect_filled(card, 2.0, Color32::from_rgb(237, 243, 251));
                painter.rect_stroke(card, 2.0, Stroke::new(1.1, blue), egui::StrokeKind::Inside);
                painter.line_segment(
                    [Pos2::new(x1 + 6.0, y1 + 8.0), Pos2::new(x2 - 8.0, y1 + 8.0)],
                    Stroke::new(1.3, blue),
                );
                painter.line_segment(
                    [Pos2::new(x1 + 6.0, cy), Pos2::new(x2 - 9.0, cy)],
                    Stroke::new(1.2, gray),
                );
                painter.line_segment(
                    [
                        Pos2::new(x1 + 6.0, y2 - 8.0),
                        Pos2::new(x2 - 12.0, y2 - 8.0),
                    ],
                    Stroke::new(1.2, gray),
                );
                painter.line_segment(
                    [
                        Pos2::new(x2 - 10.0, y1 + 9.0),
                        Pos2::new(x2 - 6.0, y1 + 5.0),
                    ],
                    Stroke::new(1.5, orange),
                );
            }
        }
    }

    fn detail_card(ui: &mut egui::Ui, title: &str, add: impl FnOnce(&mut egui::Ui)) {
        egui::Frame::group(ui.style())
            .fill(Color32::from_rgb(255, 255, 255))
            .stroke(Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(5, 4))
            .show(ui, |ui| {
                ui.vertical(|ui| {
                    ui.label(
                        RichText::new(title)
                            .size(11.0)
                            .strong()
                            .color(Color32::from_rgb(59, 80, 112)),
                    );
                    ui.add_space(3.0);
                    add(ui);
                });
            });
    }

    fn document_view(&mut self, ui: &mut egui::Ui) {
        let previous_selection = self.selected_task;

        if self.document.is_none() {
            Self::landing_view(ui, &self.recent_files);
            return;
        }

        let filter_text = self.filter_text.clone();
        let show_weekends = self.show_weekends;
        let show_critical_path = self.show_critical_path;
        let zoom = self.zoom_px_per_day;
        let mut dirty = self.dirty;
        let mut pending_selection = self.selected_task;
        let mut selection_changed = false;

        let (visible_indices, _chart_range) = {
            let document = self.document.as_ref().expect("document exists");
            (
                build_visible_task_indices(
                    &document.tasks,
                    &filter_text,
                    &self.collapsed_task_uids,
                ),
                document.chart_range(),
            )
        };

        self.handle_spreadsheet_navigation(ui.ctx(), &visible_indices);
        self.handle_sheet_clipboard(ui.ctx(), &visible_indices);

        let (document_ref, collapsed_task_uids, _task_editor) = (
            &mut self.document,
            &mut self.collapsed_task_uids,
            &mut self.task_editor,
        );
        let document = document_ref.as_mut().expect("document exists");

        let mut sheet_selection_dragging = self.sheet_selection_dragging;
        let mut chart_task_editor_refresh = false;
        pending_selection = Self::render_gantt_table(
            ui,
            document,
            &visible_indices,
            pending_selection,
            &mut self.selected_sheet_column,
            &mut self.sheet_selection_anchor,
            &mut self.sheet_editing_cell,
            &mut sheet_selection_dragging,
            zoom,
            self.status_date,
            show_weekends,
            show_critical_path,
            &mut dirty,
            &mut self.undo_stack,
            collapsed_task_uids,
            &mut self.chart_interaction,
            &mut chart_task_editor_refresh,
            &mut self.dependency_picker,
            &mut self.status,
        );
        self.sheet_selection_dragging = sheet_selection_dragging;
        if pending_selection != previous_selection {
            selection_changed = true;
        }

        if self.dependency_picker.is_some() {
            self.show_dependency_picker(ui, &mut pending_selection, &mut chart_task_editor_refresh);
        }

        self.dirty = dirty;
        self.selected_task = pending_selection;
        if selection_changed && self.selected_task.is_some() {
            self.sync_task_editor_with_selection();
        } else if chart_task_editor_refresh && self.selected_task.is_some() {
            self.sync_task_editor_with_selection();
        }
    }

    fn show_dependency_picker(
        &mut self,
        ui: &mut egui::Ui,
        pending_selection: &mut Option<usize>,
        chart_task_editor_refresh: &mut bool,
    ) {
        let ctx = ui.ctx().clone();
        let mut apply_requested = false;
        let mut cancel_requested = false;
        let mut open = true;
        {
            let Some(state) = self.dependency_picker.as_mut() else {
                return;
            };
            let source_name =
                find_task_by_uid(&state.original_document.tasks, state.source_task_uid)
                    .map(|task| task.name.clone())
                    .unwrap_or_else(|| "Source".to_string());
            let target_name =
                find_task_by_uid(&state.original_document.tasks, state.target_task_uid)
                    .map(|task| task.name.clone())
                    .unwrap_or_else(|| "Target".to_string());

            egui::Window::new("依存関係を追加")
                .collapsible(false)
                .resizable(false)
                .fixed_pos(state.popup_pos)
                .open(&mut open)
                .show(&ctx, |ui| {
                    ui.label(format!("{source_name} -> {target_name}"));
                    ui.add_space(4.0);

                    ui.horizontal_wrapped(|ui| {
                        for relation in DependencyRelationChoice::ALL {
                            ui.radio_value(&mut state.relation, relation, relation.label());
                        }
                    });

                    ui.add_space(4.0);
                    ui.horizontal(|ui| {
                        ui.label("Lag");
                        ui.text_edit_singleline(&mut state.lag_text);
                    });

                    ui.add_space(8.0);
                    ui.horizontal(|ui| {
                        if ui.button("追加").clicked() {
                            apply_requested = true;
                        }
                        if ui.button("キャンセル").clicked() {
                            cancel_requested = true;
                        }
                    });
                });
        }

        if cancel_requested || !open {
            self.dependency_picker = None;
            self.status = "依存関係の追加を取り消しました。".to_string();
            return;
        }

        if !apply_requested {
            return;
        }

        let Some(state) = self.dependency_picker.take() else {
            return;
        };
        let Some(document) = self.document.as_mut() else {
            return;
        };

        if dependency_exists(document, state.source_task_uid, state.target_task_uid) {
            self.status = "既に同じ依存関係があります。".to_string();
            return;
        }

        let lag_text = {
            let trimmed = state.lag_text.trim().to_string();
            (!trimmed.is_empty()).then_some(trimmed)
        };
        let added = add_dependency_from_drag(
            document,
            state.source_task_uid,
            state.target_task_uid,
            state.relation.label().to_string(),
            lag_text,
        );
        if !added {
            self.status = "依存関係を追加できませんでした。".to_string();
            return;
        }

        push_undo_state(
            &mut self.undo_stack,
            &state.original_document,
            self.selected_task,
            self.selected_sheet_column,
            self.sheet_selection_anchor,
            state.dirty_before,
        );
        self.dirty = true;
        self.status = "依存関係を追加しました。".to_string();
        *pending_selection = find_task_index_by_uid(&document.tasks, state.target_task_uid);
        self.selected_sheet_column = SheetColumn::Predecessors;
        if let Some(task_index) = *pending_selection {
            self.sheet_selection_anchor = Some(SheetCellSelection {
                task_index,
                column: SheetColumn::Predecessors,
            });
        }
        *chart_task_editor_refresh = true;
    }

    fn landing_view(ui: &mut egui::Ui, recent_files: &[PathBuf]) {
        let rect = ui.max_rect();
        ui.painter()
            .rect_filled(rect, 0.0, Color32::from_rgb(250, 252, 255));
        ui.allocate_ui_at_rect(rect, |ui| {
            ui.add_space(36.0);
            ui.vertical_centered(|ui| {
                ui.label(
                    RichText::new("MS Project XML ファイルを開く")
                        .font(FontId::new(18.0, FontFamily::Proportional))
                        .strong(),
                );
                ui.add_space(4.0);
                ui.label("このビューアーはタスク表、ガントチャート、詳細ペインを表示します。");
                ui.add_space(10.0);
                ui.label(RichText::new("最近使ったファイル").strong());
                if recent_files.is_empty() {
                    ui.label(RichText::new("まだありません").color(SUBTEXT));
                } else {
                    for path in recent_files {
                        ui.label(path.display().to_string());
                    }
                }
            });
        });
    }

    fn render_detail_pane(
        ui: &mut egui::Ui,
        document: &mut ProjectDocument,
        selected_task: Option<usize>,
        task_editor: &mut TaskEditorState,
        dirty: &mut bool,
    ) {
        egui::Frame::default()
            .fill(Color32::from_rgb(250, 252, 255))
            .stroke(Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(6, 6))
            .show(ui, |ui| {
                ui.vertical(|ui| {
                    ui.label(
                        RichText::new("タスクのプロパティ")
                            .font(FontId::new(14.0, FontFamily::Proportional))
                            .strong(),
                    );
                    ui.label(RichText::new("Office風のプロパティカード").color(SUBTEXT));
                    ui.add_space(4.0);

                    if let Some(index) = selected_task {
                        let task_snapshot = document.tasks.get(index).cloned();
                        if let Some(task) = task_snapshot.as_ref() {
                            Self::detail_card(ui, "概要", |ui| {
                                ui.label(
                                    RichText::new(&task.name)
                                        .font(FontId::new(16.0, FontFamily::Proportional))
                                        .strong(),
                                );
                                ui.add_space(4.0);
                                ui.horizontal_wrapped(|ui| {
                                    ui.label(format!("UID {}", task.uid));
                                    ui.separator();
                                    ui.label(format!("ID {}", task.id));
                                    ui.separator();
                                    ui.label(format!("モード {}", derive_task_mode(task)));
                                });
                                ui.add_space(2.0);
                                ui.label(format!(
                                    "日付 {} -> {}",
                                    task.start_text.as_str(),
                                    task.finish_text.as_str()
                                ));
                            });

                            ui.add_space(4.0);

                            Self::detail_card(ui, "基本情報", |ui| {
                                ui.label("名前");
                                ui.add(
                                    egui::TextEdit::singleline(&mut task_editor.name)
                                        .desired_width(ui.available_width().max(140.0))
                                        .hint_text("タスク名"),
                                );
                                ui.add_space(4.0);
                                ui.horizontal(|ui| {
                                    ui.label("階層");
                                    ui.add(
                                        egui::DragValue::new(&mut task_editor.outline_level)
                                            .range(1..=99),
                                    );
                                    ui.checkbox(&mut task_editor.summary, "要約");
                                    ui.checkbox(&mut task_editor.milestone, "マイルストーン");
                                });
                                ui.horizontal(|ui| {
                                    ui.checkbox(&mut task_editor.critical, "重要");
                                    ui.label("進捗率");
                                    ui.add(
                                        egui::DragValue::new(&mut task_editor.percent_complete)
                                            .range(0.0..=100.0)
                                            .speed(1.0),
                                    );
                                });
                                ui.label("開始");
                                ui.text_edit_singleline(&mut task_editor.start_text);
                                ui.label("終了");
                                ui.text_edit_singleline(&mut task_editor.finish_text);
                            });

                            ui.add_space(4.0);

                            Self::detail_card(ui, "スケジュール", |ui| {
                                ui.horizontal(|ui| {
                                    ui.vertical(|ui| {
                                        ui.label("ベースライン開始");
                                        ui.text_edit_singleline(
                                            &mut task_editor.baseline_start_text,
                                        );
                                    });
                                    ui.vertical(|ui| {
                                        ui.label("ベースライン終了");
                                        ui.text_edit_singleline(
                                            &mut task_editor.baseline_finish_text,
                                        );
                                    });
                                });
                                ui.add_space(4.0);
                                ui.horizontal(|ui| {
                                    ui.vertical(|ui| {
                                        ui.label("期間");
                                        ui.text_edit_singleline(&mut task_editor.duration_text);
                                    });
                                    ui.vertical(|ui| {
                                        ui.label("カレンダー UID");
                                        ui.text_edit_singleline(&mut task_editor.calendar_uid_text);
                                    });
                                });
                                ui.add_space(4.0);
                                ui.label("制約");
                                ui.text_edit_singleline(&mut task_editor.constraint_type_text);
                            });

                            ui.add_space(4.0);

                            Self::detail_card(ui, "作業", |ui| {
                                ui.label("資源名");
                                ui.text_edit_multiline(&mut task_editor.resource_names_text);
                                ui.add_space(4.0);
                                ui.label("メモ");
                                ui.text_edit_multiline(&mut task_editor.notes_text);
                            });

                            ui.add_space(4.0);

                            Self::detail_card(ui, "関連タスク", |ui| {
                                let relation_badge =
                                    |text: &str, fill: Color32, text_color: Color32| {
                                        RichText::new(text)
                                            .color(text_color)
                                            .background_color(fill)
                                            .strong()
                                    };
                                let dependency_line =
                                    |ui: &mut egui::Ui,
                                     from: &str,
                                     relation: &str,
                                     to: &str,
                                     lag: Option<&String>| {
                                        ui.horizontal_wrapped(|ui| {
                                            ui.label(
                                                RichText::new("•")
                                                    .color(Color32::from_rgb(115, 115, 115)),
                                            );
                                            ui.label(RichText::new(from).strong());
                                            ui.label(RichText::new("→").color(SUBTEXT));
                                            ui.label(RichText::new(to).strong());
                                            ui.label(relation_badge(relation, ACCENT_SOFT, ACCENT));
                                            if let Some(lag) = lag {
                                                ui.label(
                                                    RichText::new(format!("lag {}", lag))
                                                        .color(SUBTEXT),
                                                );
                                            }
                                        });
                                    };

                                let predecessors = document
                                    .dependencies
                                    .iter()
                                    .filter(|dep| dep.successor_uid == task.uid)
                                    .filter_map(|dep| {
                                        find_task_by_uid(&document.tasks, dep.predecessor_uid)
                                            .map(|pred| (dep, pred))
                                    })
                                    .take(8)
                                    .collect::<Vec<_>>();
                                ui.horizontal(|ui| {
                                    ui.label(relation_badge(
                                        "先行",
                                        Color32::from_rgb(232, 241, 252),
                                        ACCENT,
                                    ));
                                    ui.label(RichText::new("タスク").strong());
                                });
                                if predecessors.is_empty() {
                                    ui.label(
                                        RichText::new("先行タスクはありません").color(SUBTEXT),
                                    );
                                } else {
                                    for (dep, pred) in predecessors {
                                        dependency_line(
                                            ui,
                                            &pred.name,
                                            &dep.relation,
                                            &task.name,
                                            dep.lag_text.as_ref(),
                                        );
                                    }
                                }

                                ui.add_space(4.0);

                                let successors = document
                                    .dependencies
                                    .iter()
                                    .filter(|dep| dep.predecessor_uid == task.uid)
                                    .filter_map(|dep| {
                                        find_task_by_uid(&document.tasks, dep.successor_uid)
                                            .map(|succ| (dep, succ))
                                    })
                                    .take(8)
                                    .collect::<Vec<_>>();
                                ui.horizontal(|ui| {
                                    ui.label(relation_badge(
                                        "後続",
                                        Color32::from_rgb(237, 245, 231),
                                        OFFICE_GREEN_DARK,
                                    ));
                                    ui.label(RichText::new("タスク").strong());
                                });
                                if successors.is_empty() {
                                    ui.label(
                                        RichText::new("後続タスクはありません").color(SUBTEXT),
                                    );
                                } else {
                                    for (dep, succ) in successors {
                                        dependency_line(
                                            ui,
                                            &task.name,
                                            &dep.relation,
                                            &succ.name,
                                            dep.lag_text.as_ref(),
                                        );
                                    }
                                }
                            });

                            ui.add_space(4.0);

                            Self::detail_card(ui, "操作", |ui| {
                                ui.horizontal(|ui| {
                                    if Self::ribbon_icon_button(ui, RibbonIcon::TaskEdit, "適用")
                                        .clicked()
                                    {
                                        if apply_task_editor(document, index, task_editor) {
                                            if let Some(task) = document.tasks.get(index) {
                                                task_editor.sync_from_task(task);
                                            }
                                            *dirty = true;
                                        }
                                    }
                                    if Self::ribbon_icon_button(ui, RibbonIcon::Undo, "入力を戻す")
                                        .clicked()
                                    {
                                        if let Some(task) = document.tasks.get(index) {
                                            task_editor.sync_from_task(task);
                                        }
                                    }
                                });
                            });
                        } else {
                            ui.label(RichText::new("選択が範囲外です").color(SUBTEXT));
                        }
                    } else {
                        ui.label(
                            RichText::new("タスクを選択すると、ここで詳細の確認と編集ができます。")
                                .color(SUBTEXT),
                        );
                    }
                });
            });
    }

    fn render_gantt_table(
        ui: &mut egui::Ui,
        document: &mut ProjectDocument,
        visible_indices: &[usize],
        selected_task: Option<usize>,
        selected_sheet_column: &mut SheetColumn,
        selection_anchor: &mut Option<SheetCellSelection>,
        editing_cell: &mut Option<SheetCellSelection>,
        dragging: &mut bool,
        zoom_px_per_day: f32,
        status_date: NaiveDate,
        show_weekends: bool,
        show_critical_path: bool,
        dirty: &mut bool,
        undo_stack: &mut Vec<UndoState>,
        collapsed_task_uids: &mut HashSet<u32>,
        chart_interaction: &mut Option<ChartInteractionState>,
        chart_task_editor_refresh: &mut bool,
        dependency_picker: &mut Option<DependencyPickerState>,
        status: &mut String,
    ) -> Option<usize> {
        let pointer_position = ui.input(|input| input.pointer.interact_pos());
        Self::apply_active_chart_drag_preview(
            document,
            chart_interaction,
            pointer_position,
            zoom_px_per_day,
        );
        let active_drag_task_uid = chart_interaction
            .as_ref()
            .and_then(|state| match state.kind {
                ChartInteractionKind::Task { task_uid, .. } => Some(task_uid),
                ChartInteractionKind::Dependency { .. } => None,
            });
        let range = document.chart_range();
        let fixed_columns_width = sheet_fixed_columns_width();
        let chart_width =
            compute_chart_width(range, zoom_px_per_day, ui.available_width(), fixed_columns_width);
        let filler_rows = ((ui.available_height() / BODY_ROW_HEIGHT).ceil() as usize).max(12);
        let total_rows = visible_indices.len().max(filler_rows);
        let visible_task_uids: Vec<u32> = visible_indices
            .iter()
            .map(|&index| document.tasks[index].uid)
            .collect();
        let task_by_uid = visible_task_uids
            .iter()
            .enumerate()
            .map(|(row_index, &uid)| (uid, row_index))
            .collect::<HashMap<_, _>>();
        let geometry = RefCell::new(Vec::with_capacity(visible_task_uids.len()));
        let selected_uid =
            selected_task.and_then(|index| document.tasks.get(index).map(|task| task.uid));
        let editing_cell_value = *editing_cell;
        let selection_bounds = Self::sheet_selection_bounds_for_render(
            visible_indices,
            selected_task,
            *selected_sheet_column,
            *selection_anchor,
        );

        debug_assert_eq!(
            crate::table_view::build_egui_table_columns(chart_width).len(),
            8
        );
        let widths = [
            SHEET_ROW_HEADER_WIDTH,
            SHEET_NAME_WIDTH,
            SHEET_DURATION_WIDTH,
            SHEET_DATE_WIDTH,
            SHEET_DATE_WIDTH,
            SHEET_PERCENT_WIDTH,
            SHEET_PREDECESSOR_WIDTH,
            chart_width,
        ];
        let table_width = sum_widths(&widths);
        let table_height = HEADER_HEIGHT + total_rows as f32 * BODY_ROW_HEIGHT;
        let (table_rect, _) =
            ui.allocate_exact_size(Vec2::new(table_width, table_height), Sense::hover());
        let painter = ui.painter_at(table_rect);
        let table_left = table_rect.left();
        let chart_left = table_left + fixed_columns_width;
        let header_top = table_rect.top();
        let body_top = header_top + HEADER_HEIGHT;

        painter.rect_filled(table_rect, 0.0, SHEET_BG);
        painter.rect_stroke(
            table_rect,
            0.0,
            Stroke::new(1.0, BORDER),
            egui::StrokeKind::Inside,
        );

        let mut x = table_left;
        let header_cells = [
            ("", HeaderGlyph::Mode, SHEET_ROW_HEADER_WIDTH),
            ("Task Name", HeaderGlyph::Name, SHEET_NAME_WIDTH),
            ("Duration", HeaderGlyph::Duration, SHEET_DURATION_WIDTH),
            ("Start", HeaderGlyph::Date, SHEET_DATE_WIDTH),
            ("Finish", HeaderGlyph::Date, SHEET_DATE_WIDTH),
            ("% Complete", HeaderGlyph::Percent, SHEET_PERCENT_WIDTH),
            (
                "Predecessors",
                HeaderGlyph::Predecessors,
                SHEET_PREDECESSOR_WIDTH,
            ),
        ];
        for (title, glyph, width) in header_cells {
            let rect =
                Rect::from_min_size(Pos2::new(x, header_top), Vec2::new(width, HEADER_HEIGHT));
            paint_sheet_header_cell(&painter, rect, title, glyph);
            x += width;
        }
        let gantt_rect = Rect::from_min_size(
            Pos2::new(x, header_top),
            Vec2::new(chart_width, HEADER_HEIGHT),
        );
        draw_timescale_header(ui, gantt_rect, range, zoom_px_per_day);
        painter.text(
            Pos2::new(gantt_rect.left() + 26.0, gantt_rect.top() + 7.0),
            egui::Align2::LEFT_TOP,
            "Gantt",
            FontId::new(11.0, FontFamily::Proportional),
            Color32::from_rgb(49, 66, 94),
        );
        draw_header_glyph(
            &painter,
            Rect::from_min_size(
                Pos2::new(gantt_rect.left() + 6.0, gantt_rect.top() + 6.0),
                Vec2::new(14.0, 14.0),
            ),
            HeaderGlyph::Gantt,
        );

        let mut pending_selection = selected_task;
        for visible_row in 0..total_rows {
            let row_top = body_top + visible_row as f32 * BODY_ROW_HEIGHT;
            let mut cell_x = table_left;
            if let Some(index) = visible_indices.get(visible_row).copied() {
                let task = document.tasks[index].clone();
                let collapsed = collapsed_task_uids.contains(&task.uid);
                let row_matches_critical = task.critical && show_critical_path;
                let row_base_color = if row_matches_critical {
                    SHEET_CRITICAL_FILL
                } else if task.summary {
                    SHEET_SUMMARY_FILL
                } else {
                    SHEET_BG
                };
                let row_selected = pending_selection == Some(index);

                let mode_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_ROW_HEADER_WIDTH, BODY_ROW_HEIGHT),
                );
                let mode_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Mode));
                paint_sheet_cell(
                    &painter,
                    mode_rect,
                    if row_selected || mode_selected {
                        SHEET_ROW_HEADER_SELECTED_FILL
                    } else {
                        SHEET_ROW_HEADER_FILL
                    },
                    row_selected || mode_selected,
                    false,
                );
                let mode_response = ui.interact(
                    mode_rect,
                    ui.id().with(("sheet_mode", index)),
                    Sense::click_and_drag(),
                );
                let chip_color = if task.summary {
                    Color32::from_rgb(118, 118, 118)
                } else if task.milestone {
                    Color32::from_rgb(194, 60, 40)
                } else if row_matches_critical {
                    CRITICAL
                } else {
                    OFFICE_GREEN_DARK
                };
                painter.text(
                    Pos2::new(mode_rect.left() + 6.0, mode_rect.center().y),
                    egui::Align2::LEFT_CENTER,
                    format!("{:>2}", visible_row + 1),
                    FontId::new(10.0, FontFamily::Proportional),
                    SHEET_HEADER_TEXT,
                );
                let chip_rect = Rect::from_min_size(
                    Pos2::new(mode_rect.left() + 22.0, mode_rect.center().y - 5.5),
                    Vec2::new(11.0, 11.0),
                );
                draw_mode_badge(&painter, chip_rect, chip_color, &task);
                if mode_response.double_clicked() {
                    *selected_sheet_column = SheetColumn::Mode;
                    pending_selection = Some(index);
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Mode,
                    });
                    *editing_cell = None;
                } else if mode_response.clicked() || mode_response.drag_started() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &mode_response,
                        index,
                        SheetColumn::Mode,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_ROW_HEADER_WIDTH;

                let name_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_NAME_WIDTH, BODY_ROW_HEIGHT),
                );
                let name_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Name));
                let name_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Name,
                    });
                paint_sheet_cell(
                    &painter,
                    name_rect,
                    if name_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    name_selected,
                    name_active,
                );
                let name_response = ui.interact(
                    name_rect,
                    ui.id().with(("sheet_name", index)),
                    Sense::click_and_drag(),
                );
                let indent = ((task.outline_level.max(1) - 1) as f32 * 12.0).max(0.0);
                let content_left = name_rect.left() + SHEET_CELL_PADDING_X + indent;
                if task.summary {
                    painter.text(
                        Pos2::new(content_left + 4.0, name_rect.center().y),
                        egui::Align2::LEFT_CENTER,
                        if collapsed { "▶" } else { "▼" },
                        FontId::new(10.0, FontFamily::Proportional),
                        SUBTEXT,
                    );
                }
                if name_active {
                    let mut name = task.name.clone();
                    let edit_rect = Rect::from_min_max(
                        Pos2::new(content_left + 14.0, name_rect.top() + 1.0),
                        Pos2::new(
                            name_rect.right() - SHEET_CELL_PADDING_X,
                            name_rect.bottom() - 1.0,
                        ),
                    );
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut name).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        document.tasks[index].name = name;
                        *dirty = true;
                    }
                } else {
                    painter.text(
                        Pos2::new(
                            content_left + if task.summary { 16.0 } else { 14.0 },
                            name_rect.center().y,
                        ),
                        egui::Align2::LEFT_CENTER,
                        task.name.as_str(),
                        FontId::new(10.5, FontFamily::Proportional),
                        TEXT,
                    );
                }
                if name_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::Name;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Name,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Name,
                    });
                } else if name_response.clicked() {
                    if task.summary
                        && name_response.interact_pointer_pos().is_some_and(|pos| {
                            Rect::from_min_size(
                                Pos2::new(content_left, name_rect.center().y - 6.0),
                                Vec2::new(12.0, 12.0),
                            )
                            .contains(pos)
                        })
                    {
                        if collapsed {
                            collapsed_task_uids.remove(&task.uid);
                        } else {
                            collapsed_task_uids.insert(task.uid);
                        }
                    }
                    Self::update_sheet_selection_from_response(
                        ui,
                        &name_response,
                        index,
                        SheetColumn::Name,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_NAME_WIDTH;

                let duration_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_DURATION_WIDTH, BODY_ROW_HEIGHT),
                );
                let duration_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Duration));
                let duration_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Duration,
                    });
                paint_sheet_cell(
                    &painter,
                    duration_rect,
                    if duration_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    duration_selected,
                    duration_active,
                );
                let duration_response = ui.interact(
                    duration_rect,
                    ui.id().with(("sheet_duration", index)),
                    Sense::click_and_drag(),
                );
                if duration_active {
                    let mut duration = display_duration_text(&task.duration_text);
                    let edit_rect = duration_rect
                        .shrink2(Vec2::new(SHEET_CELL_PADDING_X, SHEET_CELL_PADDING_Y));
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut duration).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        document.tasks[index].duration_text = duration;
                        *dirty = true;
                    }
                    if response.lost_focus() {
                        let normalized =
                            normalize_duration_text(&document.tasks[index].duration_text);
                        if document.tasks[index].duration_text != normalized {
                            push_undo_state(
                                undo_stack,
                                document,
                                selected_task,
                                *selected_sheet_column,
                                *selection_anchor,
                                *dirty,
                            );
                            document.tasks[index].duration_text = normalized;
                            *dirty = true;
                        }
                    }
                }
                if duration_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::Duration;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Duration,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Duration,
                    });
                } else if duration_response.clicked() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &duration_response,
                        index,
                        SheetColumn::Duration,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_DURATION_WIDTH;

                let start_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_DATE_WIDTH, BODY_ROW_HEIGHT),
                );
                let start_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Start));
                let start_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Start,
                    });
                paint_sheet_cell(
                    &painter,
                    start_rect,
                    if start_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    start_selected,
                    start_active,
                );
                let start_response = ui.interact(
                    start_rect,
                    ui.id().with(("sheet_start", index)),
                    Sense::click_and_drag(),
                );
                if start_active {
                    let mut start_text = task.start_text.clone();
                    let edit_rect =
                        start_rect.shrink2(Vec2::new(SHEET_CELL_PADDING_X, SHEET_CELL_PADDING_Y));
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut start_text).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        document.tasks[index].start_text = start_text.clone();
                        document.tasks[index].start = parse_date_time(Some(start_text.as_str()));
                        *dirty = true;
                    }
                } else {
                    painter.text(
                        Pos2::new(
                            start_rect.right() - SHEET_CELL_PADDING_X,
                            start_rect.center().y,
                        ),
                        egui::Align2::RIGHT_CENTER,
                        task.start_text.as_str(),
                        FontId::new(10.0, FontFamily::Proportional),
                        TEXT,
                    );
                }
                if start_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::Start;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Start,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Start,
                    });
                } else if start_response.clicked() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &start_response,
                        index,
                        SheetColumn::Start,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_DATE_WIDTH;

                let finish_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_DATE_WIDTH, BODY_ROW_HEIGHT),
                );
                let finish_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Finish));
                let finish_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Finish,
                    });
                paint_sheet_cell(
                    &painter,
                    finish_rect,
                    if finish_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    finish_selected,
                    finish_active,
                );
                let finish_response = ui.interact(
                    finish_rect,
                    ui.id().with(("sheet_finish", index)),
                    Sense::click_and_drag(),
                );
                if finish_active {
                    let mut finish_text = task.finish_text.clone();
                    let edit_rect =
                        finish_rect.shrink2(Vec2::new(SHEET_CELL_PADDING_X, SHEET_CELL_PADDING_Y));
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut finish_text).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        document.tasks[index].finish_text = finish_text.clone();
                        document.tasks[index].finish = parse_date_time(Some(finish_text.as_str()));
                        *dirty = true;
                    }
                } else {
                    painter.text(
                        Pos2::new(
                            finish_rect.right() - SHEET_CELL_PADDING_X,
                            finish_rect.center().y,
                        ),
                        egui::Align2::RIGHT_CENTER,
                        task.finish_text.as_str(),
                        FontId::new(10.0, FontFamily::Proportional),
                        TEXT,
                    );
                }
                if finish_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::Finish;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Finish,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Finish,
                    });
                } else if finish_response.clicked() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &finish_response,
                        index,
                        SheetColumn::Finish,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_DATE_WIDTH;

                let percent_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_PERCENT_WIDTH, BODY_ROW_HEIGHT),
                );
                let percent_selected = selection_bounds.is_some_and(|bounds| {
                    bounds.contains(visible_row, SheetColumn::PercentComplete)
                });
                let percent_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::PercentComplete,
                    });
                paint_sheet_cell(
                    &painter,
                    percent_rect,
                    if percent_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    percent_selected,
                    percent_active,
                );
                let percent_response = ui.interact(
                    percent_rect,
                    ui.id().with(("sheet_percent", index)),
                    Sense::click_and_drag(),
                );
                if percent_active {
                    let mut percent_text = format!("{:.0}", task.percent_complete);
                    let edit_rect =
                        percent_rect.shrink2(Vec2::new(SHEET_CELL_PADDING_X, SHEET_CELL_PADDING_Y));
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut percent_text).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        if let Ok(value) = percent_text
                            .trim()
                            .trim_end_matches('%')
                            .trim()
                            .parse::<f32>()
                        {
                            document.tasks[index].percent_complete = value.clamp(0.0, 100.0);
                            *dirty = true;
                        }
                    }
                } else {
                    painter.text(
                        Pos2::new(
                            percent_rect.right() - SHEET_CELL_PADDING_X,
                            percent_rect.center().y,
                        ),
                        egui::Align2::RIGHT_CENTER,
                        format!("{:.0}%", task.percent_complete),
                        FontId::new(10.0, FontFamily::Proportional),
                        TEXT,
                    );
                }
                if percent_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::PercentComplete;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::PercentComplete,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::PercentComplete,
                    });
                } else if percent_response.clicked() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &percent_response,
                        index,
                        SheetColumn::PercentComplete,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_PERCENT_WIDTH;

                let predecessor_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(SHEET_PREDECESSOR_WIDTH, BODY_ROW_HEIGHT),
                );
                let predecessor_selected = selection_bounds
                    .is_some_and(|bounds| bounds.contains(visible_row, SheetColumn::Predecessors));
                let predecessor_active = editing_cell_value
                    == Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Predecessors,
                    });
                paint_sheet_cell(
                    &painter,
                    predecessor_rect,
                    if predecessor_selected {
                        SHEET_SELECTED_FILL
                    } else {
                        row_base_color
                    },
                    predecessor_selected,
                    predecessor_active,
                );
                let predecessor_response = ui.interact(
                    predecessor_rect,
                    ui.id().with(("sheet_predecessors", index)),
                    Sense::click_and_drag(),
                );
                if predecessor_active {
                    let mut label = task.predecessor_text.as_str().to_string();
                    let edit_rect = predecessor_rect
                        .shrink2(Vec2::new(SHEET_CELL_PADDING_X, SHEET_CELL_PADDING_Y));
                    let response = ui.put(
                        edit_rect,
                        egui::TextEdit::singleline(&mut label).frame(egui::Frame::NONE),
                    );
                    response.request_focus();
                    if response.changed() {
                        push_undo_state(
                            undo_stack,
                            document,
                            selected_task,
                            *selected_sheet_column,
                            *selection_anchor,
                            *dirty,
                        );
                        document.tasks[index].predecessor_text = label.clone();
                        if Self::apply_excel_predecessors(document, index, &label) {
                            *dirty = true;
                        }
                    }
                } else {
                    let label = if task.predecessor_text.is_empty() {
                        "-".to_string()
                    } else {
                        task.predecessor_text.clone()
                    };
                    painter.text(
                        Pos2::new(
                            predecessor_rect.left() + SHEET_CELL_PADDING_X,
                            predecessor_rect.center().y,
                        ),
                        egui::Align2::LEFT_CENTER,
                        label,
                        FontId::new(10.0, FontFamily::Proportional),
                        SUBTEXT,
                    );
                }
                if predecessor_response.double_clicked() {
                    pending_selection = Some(index);
                    *selected_sheet_column = SheetColumn::Predecessors;
                    *selection_anchor = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Predecessors,
                    });
                    *editing_cell = Some(SheetCellSelection {
                        task_index: index,
                        column: SheetColumn::Predecessors,
                    });
                } else if predecessor_response.clicked() {
                    Self::update_sheet_selection_from_response(
                        ui,
                        &predecessor_response,
                        index,
                        SheetColumn::Predecessors,
                        &mut pending_selection,
                        selected_sheet_column,
                        selection_anchor,
                        dragging,
                    );
                    *editing_cell = None;
                }
                cell_x += SHEET_PREDECESSOR_WIDTH;

                let chart_rect = Rect::from_min_size(
                    Pos2::new(cell_x, row_top),
                    Vec2::new(chart_width, BODY_ROW_HEIGHT),
                );
                paint_sheet_cell(&painter, chart_rect, row_base_color, false, false);
                paint_chart_row_background(
                    &painter,
                    chart_rect,
                    range,
                    zoom_px_per_day,
                    show_weekends,
                );
                let geom = draw_task_bar(
                    &painter,
                    chart_rect,
                    &task,
                    range,
                    zoom_px_per_day,
                    row_selected,
                    show_critical_path,
                    active_drag_task_uid == Some(task.uid),
                );
                let chart_drag_active = chart_interaction.is_some() || dependency_picker.is_some();
                let body_response = ui
                    .interact(
                        geom.body_rect,
                        ui.id().with(("gantt_body", task.uid)),
                        Sense::click_and_drag(),
                    )
                    .on_hover_text("ドラッグでタスクを移動");
                let body_response = body_response.on_hover_cursor(egui::CursorIcon::Grab);
                let start_handle_response = geom.start_handle_rect.map(|rect| {
                    ui.interact(
                        rect,
                        ui.id().with(("gantt_start_handle", task.uid)),
                        Sense::click_and_drag(),
                    )
                    .on_hover_text("開始日を変更")
                    .on_hover_cursor(egui::CursorIcon::ResizeHorizontal)
                });
                let finish_handle_response = geom.finish_handle_rect.map(|rect| {
                    ui.interact(
                        rect,
                        ui.id().with(("gantt_finish_handle", task.uid)),
                        Sense::click_and_drag(),
                    )
                    .on_hover_text("終了日を変更")
                    .on_hover_cursor(egui::CursorIcon::ResizeHorizontal)
                });
                let link_handle_response = ui
                    .interact(
                        geom.link_handle_rect,
                        ui.id().with(("gantt_link_handle", task.uid)),
                        Sense::click_and_drag(),
                    )
                    .on_hover_text("依存関係を作成")
                    .on_hover_cursor(egui::CursorIcon::Crosshair);

                if !chart_drag_active {
                    if body_response.double_clicked() || body_response.clicked() {
                        pending_selection = Some(index);
                        *selected_sheet_column = SheetColumn::Name;
                        *selection_anchor = Some(SheetCellSelection {
                            task_index: index,
                            column: SheetColumn::Name,
                        });
                    }
                    if body_response.drag_started() {
                        Self::begin_chart_task_drag(
                            chart_interaction,
                            document,
                            index,
                            task.uid,
                            TaskDragKind::Move,
                            pointer_position,
                            *dirty,
                        );
                        pending_selection = Some(index);
                        *selected_sheet_column = SheetColumn::Name;
                        *selection_anchor = Some(SheetCellSelection {
                            task_index: index,
                            column: SheetColumn::Name,
                        });
                        *editing_cell = None;
                    }
                    if let Some(response) = start_handle_response.as_ref() {
                        if response.drag_started() {
                            Self::begin_chart_task_drag(
                                chart_interaction,
                                document,
                                index,
                                task.uid,
                                TaskDragKind::ResizeStart,
                                pointer_position,
                                *dirty,
                            );
                            pending_selection = Some(index);
                            *selected_sheet_column = SheetColumn::Name;
                            *selection_anchor = Some(SheetCellSelection {
                                task_index: index,
                                column: SheetColumn::Name,
                            });
                            *editing_cell = None;
                        }
                    }
                    if let Some(response) = finish_handle_response.as_ref() {
                        if response.drag_started() {
                            Self::begin_chart_task_drag(
                                chart_interaction,
                                document,
                                index,
                                task.uid,
                                TaskDragKind::ResizeFinish,
                                pointer_position,
                                *dirty,
                            );
                            pending_selection = Some(index);
                            *selected_sheet_column = SheetColumn::Name;
                            *selection_anchor = Some(SheetCellSelection {
                                task_index: index,
                                column: SheetColumn::Name,
                            });
                            *editing_cell = None;
                        }
                    }
                    if link_handle_response.drag_started() {
                        Self::begin_chart_dependency_drag(
                            chart_interaction,
                            document,
                            index,
                            task.uid,
                            pointer_position,
                            *dirty,
                        );
                        pending_selection = Some(index);
                        *selected_sheet_column = SheetColumn::Predecessors;
                        *selection_anchor = Some(SheetCellSelection {
                            task_index: index,
                            column: SheetColumn::Predecessors,
                        });
                        *editing_cell = None;
                    }
                }
                geometry.borrow_mut().push(geom);
            } else {
                let mut blank_x = table_left;
                for (column_index, width) in widths.iter().copied().enumerate() {
                    let rect = Rect::from_min_size(
                        Pos2::new(blank_x, row_top),
                        Vec2::new(width, BODY_ROW_HEIGHT),
                    );
                    let fill = if column_index == 0 {
                        SHEET_ROW_HEADER_FILL
                    } else {
                        SHEET_BG
                    };
                    paint_sheet_cell(&painter, rect, fill, false, false);
                    if column_index == widths.len() - 1 {
                        paint_chart_row_background(
                            &painter,
                            rect,
                            range,
                            zoom_px_per_day,
                            show_weekends,
                        );
                    }
                    blank_x += width;
                }

                if let Some(task_index) = Self::sheet_row_target_index(visible_indices, visible_row)
                {
                    let row_rect = Rect::from_min_size(
                        Pos2::new(table_left, row_top),
                        Vec2::new(table_width, BODY_ROW_HEIGHT),
                    );
                    let filler_response = ui.interact(
                        row_rect,
                        ui.id().with(("sheet_filler", visible_row)),
                        Sense::click_and_drag(),
                    );
                    if filler_response.clicked() || filler_response.drag_started() {
                        Self::update_sheet_selection_from_response(
                            ui,
                            &filler_response,
                            task_index,
                            *selected_sheet_column,
                            &mut pending_selection,
                            selected_sheet_column,
                            selection_anchor,
                            dragging,
                        );
                        *editing_cell = None;
                    }
                }
            }
        }

        let rects = geometry.borrow().clone();
        let pointer_released = ui.input(|input| !input.pointer.primary_down());
        if pointer_released && chart_interaction.is_some() {
            Self::finish_chart_interaction(
                document,
                chart_interaction,
                &rects,
                pointer_position,
                &mut pending_selection,
                selected_sheet_column,
                selection_anchor,
                dirty,
                undo_stack,
                chart_task_editor_refresh,
                dependency_picker,
                status,
            );
        }

        if !rects.is_empty() {
            draw_dependencies(
                ui.painter(),
                &rects,
                &task_by_uid,
                &document.dependencies,
                selected_uid,
                active_drag_task_uid,
            );

            let status_x = status_date_x(
                status_date,
                range,
                chart_left,
                zoom_px_per_day,
            );
            let line_points = collect_inazuma_progress_points(
                &rects,
                visible_indices,
                &document.tasks,
                range,
                chart_left,
                gantt_rect.right(),
                zoom_px_per_day,
                status_date,
            );
            let stroke_color = Color32::from_rgb(226, 0, 0);
            draw_progress_actual_line(
                ui.painter(),
                &line_points,
                status_x,
                Rect::from_min_max(
                    Pos2::new(gantt_rect.left(), body_top),
                    Pos2::new(gantt_rect.right(), table_rect.bottom()),
                ),
                Stroke::new(2.0, stroke_color),
            );
        }

        if false {
        egui::ScrollArea::both()
            .auto_shrink([false, false])
            .drag_to_scroll(false)
            .show(ui, |ui| {
                let table = TableBuilder::new(ui)
                    .striped(false)
                    .resizable(true)
                    .cell_layout(egui::Layout::left_to_right(Align::Center))
                    .column(Column::exact(42.0))
                    .column(Column::exact(180.0))
                    .column(Column::exact(72.0))
                    .column(Column::exact(92.0))
                    .column(Column::exact(92.0))
                    .column(Column::exact(74.0))
                    .column(Column::exact(122.0))
                    .column(Column::exact(chart_width));

                table
                    .header(HEADER_HEIGHT, |mut header| {
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "", HeaderGlyph::Mode);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "Task Name", HeaderGlyph::Name);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "Duration", HeaderGlyph::Duration);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "Start", HeaderGlyph::Date);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "Finish", HeaderGlyph::Date);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "% Complete", HeaderGlyph::Percent);
                        });
                        header.col(|ui| {
                            paint_sheet_header_cell(ui.painter(), ui.max_rect(), "Predecessors", HeaderGlyph::Predecessors);
                        });
                        header.col(|ui| {
                            let rect = ui.max_rect();
                            ui.painter()
                                .rect_filled(rect, 0.0, Color32::from_rgb(217, 228, 237));
                            ui.painter().rect_stroke(
                                rect,
                                0.0,
                                Stroke::new(1.0, Color32::from_rgb(176, 188, 199)),
                                egui::StrokeKind::Inside,
                            );
                            let title_pos = Pos2::new(rect.left() + 26.0, rect.top() + 7.0);
                            draw_header_glyph(
                                ui.painter(),
                                Rect::from_min_size(
                                    Pos2::new(rect.left() + 6.0, rect.top() + 6.0),
                                    Vec2::new(14.0, 14.0),
                                ),
                                HeaderGlyph::Gantt,
                            );
                            ui.painter().text(
                                title_pos,
                                egui::Align2::LEFT_TOP,
                                "Gantt",
                                FontId::new(11.0, FontFamily::Proportional),
                                Color32::from_rgb(49, 66, 94),
                            );
                            draw_timescale_header(ui, rect, range, zoom_px_per_day);
                        });
                    })
                    .body(|body| {
                        body.rows(BODY_ROW_HEIGHT, total_rows, |mut row| {
                            let visible_row = row.index();
                            if let Some(index) = visible_indices.get(visible_row).copied() {
                                let task = document.tasks[index].clone();
                                let active_column = *selected_sheet_column;
                                let collapsed = collapsed_task_uids.contains(&task.uid);
                                let row_matches_critical = task.critical && show_critical_path;
                                let row_base_color = if row_matches_critical {
                                    SHEET_CRITICAL_FILL
                                } else if task.summary {
                                    SHEET_SUMMARY_FILL
                                } else {
                                    SHEET_BG
                                };

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Mode)
                                    });
                                    let _selected = pending_selection == Some(index)
                                        && active_column == SheetColumn::Mode;
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Mode,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_ROW_HEADER_SELECTED_FILL
                                        } else {
                                            SHEET_ROW_HEADER_FILL
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    ui.horizontal(|ui| {
                                        let chip_color = if task.summary {
                                            Color32::from_rgb(118, 118, 118)
                                        } else if task.milestone {
                                            Color32::from_rgb(194, 60, 40)
                                        } else if row_matches_critical {
                                            CRITICAL
                                        } else {
                                            OFFICE_GREEN_DARK
                                        };
                                        ui.add_space(2.0);
                                        let row_response = ui.add(
                                            egui::Label::new(
                                                RichText::new(format!("{:>2}", visible_row + 1))
                                                    .size(10.0)
                                                    .strong()
                                                    .color(SHEET_HEADER_TEXT),
                                            )
                                            .sense(Sense::click_and_drag()),
                                        );
                                        ui.add_space(4.0);
                                        let (chip_rect, chip_response) = ui.allocate_exact_size(
                                            Vec2::new(11.0, 11.0),
                                            Sense::click(),
                                        );
                                        if ui.is_rect_visible(chip_rect) {
                                            draw_mode_badge(
                                                ui.painter(),
                                                chip_rect,
                                                chip_color,
                                                &task,
                                            );
                                        }
                                        if row_response.double_clicked()
                                            || chip_response.double_clicked()
                                        {
                                            *selected_sheet_column = SheetColumn::Mode;
                                            pending_selection = Some(index);
                                            *selection_anchor = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Mode,
                                            });
                                            *editing_cell = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Mode,
                                            });
                                        } else if row_response.clicked() || chip_response.clicked()
                                        {
                                            Self::update_sheet_selection_from_response(
                                                ui,
                                                if row_response.clicked() {
                                                    &row_response
                                                } else {
                                                    &chip_response
                                                },
                                                index,
                                                SheetColumn::Mode,
                                                &mut pending_selection,
                                                selected_sheet_column,
                                                selection_anchor,
                                                dragging,
                                            );
                                            *editing_cell = None;
                                        }
                });
            });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Name)
                                    });
                                    let _selected = pending_selection == Some(index)
                                        && active_column == SheetColumn::Name;
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Name,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    ui.horizontal(|ui| {
                                        let indent = ((task.outline_level.max(1) - 1) as f32
                                            * 12.0)
                                            .max(0.0);
                                        ui.add_space(indent);
                                        if task.summary {
                                            let arrow = if collapsed { "▶" } else { "▼" };
                                            if ui.small_button(arrow).clicked() {
                                                if collapsed {
                                                    collapsed_task_uids.remove(&task.uid);
                                                } else {
                                                    collapsed_task_uids.insert(task.uid);
                                                }
                                            }
                                        } else {
                                            ui.add_space(16.0);
                                        }
                                        let mut name = task.name.clone();
                                        let response = if active {
                                            ui.add(
                                                egui::TextEdit::singleline(&mut name)
                                                    .desired_width(ui.available_width().max(120.0)),
                                            )
                                        } else {
                                            ui.add(
                                                egui::Label::new(RichText::new(task.name.as_str()))
                                                    .sense(Sense::click_and_drag()),
                                            )
                                        };
                                        if active {
                                            response.request_focus();
                                        }
                                        if response.double_clicked() {
                                            pending_selection = Some(index);
                                            *selected_sheet_column = SheetColumn::Name;
                                            *selection_anchor = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Name,
                                            });
                                            *editing_cell = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Name,
                                            });
                                        } else {
                                            Self::update_sheet_selection_from_response(
                                                ui,
                                                &response,
                                                index,
                                                SheetColumn::Name,
                                                &mut pending_selection,
                                                selected_sheet_column,
                                                selection_anchor,
                                                dragging,
                                            );
                                            if response.clicked() {
                                                *editing_cell = None;
                                            }
                                        }
                                        if response.changed() {
                                            document.tasks[index].name = name;
                                            *dirty = true;
                                        }
                                    });
                                });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Duration)
                                    });
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Duration,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    let mut duration = display_duration_text(&task.duration_text);
                                    let response = if active {
                                        ui.add(
                                            egui::TextEdit::singleline(&mut duration)
                                                .desired_width(ui.available_width().max(90.0)),
                                        )
                                    } else {
                                        ui.add(
                                            egui::Label::new(RichText::new(
                                                task.duration_text.as_str(),
                                            ))
                                            .sense(Sense::click_and_drag()),
                                        )
                                    };
                                    if active {
                                        response.request_focus();
                                    }
                                    if response.double_clicked() {
                                        pending_selection = Some(index);
                                        *selected_sheet_column = SheetColumn::Duration;
                                        *selection_anchor = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Duration,
                                        });
                                        *editing_cell = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Duration,
                                        });
                                    } else {
                                        Self::update_sheet_selection_from_response(
                                            ui,
                                            &response,
                                            index,
                                            SheetColumn::Duration,
                                            &mut pending_selection,
                                            selected_sheet_column,
                                            selection_anchor,
                                            dragging,
                                        );
                                        if response.clicked() {
                                            *editing_cell = None;
                                        }
                                    }
                                    if response.changed() {
                                        document.tasks[index].duration_text = duration;
                                        *dirty = true;
                                    }
                                    if response.lost_focus() {
                                        let normalized = normalize_duration_text(
                                            &document.tasks[index].duration_text,
                                        );
                                        if document.tasks[index].duration_text != normalized {
                                            document.tasks[index].duration_text = normalized;
                                            *dirty = true;
                                        }
                                    }
                                });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Start)
                                    });
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Start,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    let mut start_text = task.start_text.clone();
                                    let response = if active {
                                        ui.add(
                                            egui::TextEdit::singleline(&mut start_text)
                                                .desired_width(ui.available_width().max(110.0)),
                                        )
                                    } else {
                                        ui.add(
                                            egui::Label::new(RichText::new(
                                                task.start_text.as_str(),
                                            ))
                                            .sense(Sense::click_and_drag()),
                                        )
                                    };
                                    if active {
                                        response.request_focus();
                                    }
                                    if response.double_clicked() {
                                        pending_selection = Some(index);
                                        *selected_sheet_column = SheetColumn::Start;
                                        *selection_anchor = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Start,
                                        });
                                        *editing_cell = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Start,
                                        });
                                    } else {
                                        Self::update_sheet_selection_from_response(
                                            ui,
                                            &response,
                                            index,
                                            SheetColumn::Start,
                                            &mut pending_selection,
                                            selected_sheet_column,
                                            selection_anchor,
                                            dragging,
                                        );
                                        if response.clicked() {
                                            *editing_cell = None;
                                        }
                                    }
                                    if response.changed() {
                                        document.tasks[index].start_text = start_text.clone();
                                        document.tasks[index].start =
                                            parse_date_time(Some(start_text.as_str()));
                                        *dirty = true;
                                    }
                                });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Finish)
                                    });
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Finish,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    let mut finish_text = task.finish_text.clone();
                                    let response = if active {
                                        ui.add(
                                            egui::TextEdit::singleline(&mut finish_text)
                                                .desired_width(ui.available_width().max(110.0)),
                                        )
                                    } else {
                                        ui.add(
                                            egui::Label::new(RichText::new(
                                                task.finish_text.as_str(),
                                            ))
                                            .sense(Sense::click_and_drag()),
                                        )
                                    };
                                    if active {
                                        response.request_focus();
                                    }
                                    if response.double_clicked() {
                                        pending_selection = Some(index);
                                        *selected_sheet_column = SheetColumn::Finish;
                                        *selection_anchor = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Finish,
                                        });
                                        *editing_cell = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Finish,
                                        });
                                    } else {
                                        Self::update_sheet_selection_from_response(
                                            ui,
                                            &response,
                                            index,
                                            SheetColumn::Finish,
                                            &mut pending_selection,
                                            selected_sheet_column,
                                            selection_anchor,
                                            dragging,
                                        );
                                        if response.clicked() {
                                            *editing_cell = None;
                                        }
                                    }
                                    if response.changed() {
                                        document.tasks[index].finish_text = finish_text.clone();
                                        document.tasks[index].finish =
                                            parse_date_time(Some(finish_text.as_str()));
                                        *dirty = true;
                                    }
                                });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::PercentComplete)
                                    });
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::PercentComplete,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    let mut percent_text = format!("{:.0}", task.percent_complete);
                                    let response = if active {
                                        ui.add(
                                            egui::TextEdit::singleline(&mut percent_text)
                                                .desired_width(ui.available_width().max(60.0)),
                                        )
                                    } else {
                                        ui.add(
                                            egui::Label::new(RichText::new(format!(
                                                "{:.0}%",
                                                task.percent_complete
                                            )))
                                            .sense(Sense::click_and_drag()),
                                        )
                                    };
                                    if active {
                                        response.request_focus();
                                    }
                                    if response.double_clicked() {
                                        pending_selection = Some(index);
                                        *selected_sheet_column = SheetColumn::PercentComplete;
                                        *selection_anchor = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::PercentComplete,
                                        });
                                        *editing_cell = Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::PercentComplete,
                                        });
                                    } else {
                                        Self::update_sheet_selection_from_response(
                                            ui,
                                            &response,
                                            index,
                                            SheetColumn::PercentComplete,
                                            &mut pending_selection,
                                            selected_sheet_column,
                                            selection_anchor,
                                            dragging,
                                        );
                                        if response.clicked() {
                                            *editing_cell = None;
                                        }
                                    }
                                    if response.changed() {
                                        if let Ok(value) = percent_text
                                            .trim()
                                            .trim_end_matches('%')
                                            .trim()
                                            .parse::<f32>()
                                        {
                                            document.tasks[index].percent_complete =
                                                value.clamp(0.0, 100.0);
                                            *dirty = true;
                                        }
                                    }
                                });

                                row.col(|ui| {
                                    let cell_selected = selection_bounds.is_some_and(|bounds| {
                                        bounds.contains(visible_row, SheetColumn::Predecessors)
                                    });
                                    let active = editing_cell_value
                                        == Some(SheetCellSelection {
                                            task_index: index,
                                            column: SheetColumn::Predecessors,
                                        });
                                    paint_cell(
                                        ui,
                                        Some(if cell_selected {
                                            SHEET_SELECTED_FILL
                                        } else {
                                            row_base_color
                                        }),
                                        cell_selected,
                                        active,
                                    );
                                    let mut label = task.predecessor_text.as_str().to_string();
                                    if active {
                                        let response = ui.add(
                                            egui::TextEdit::singleline(&mut label)
                                                .desired_width(ui.available_width().max(100.0)),
                                        );
                                        response.request_focus();
                                        if response.double_clicked() {
                                            pending_selection = Some(index);
                                            *selected_sheet_column = SheetColumn::Predecessors;
                                            *selection_anchor = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Predecessors,
                                            });
                                            *editing_cell = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Predecessors,
                                            });
                                        } else {
                                            Self::update_sheet_selection_from_response(
                                                ui,
                                                &response,
                                                index,
                                                SheetColumn::Predecessors,
                                                &mut pending_selection,
                                                selected_sheet_column,
                                                selection_anchor,
                                                dragging,
                                            );
                                            if response.clicked() {
                                                *editing_cell = None;
                                            }
                                        }
                                        if response.changed() {
                                            document.tasks[index].predecessor_text = label.clone();
                                            if Self::apply_excel_predecessors(
                                                document, index, &label,
                                            ) {
                                                *dirty = true;
                                            }
                                        }
                                    } else {
                                        if label.is_empty() {
                                            label = "-".to_string();
                                        }
                                        let response = ui.add(
                                            egui::Label::new(RichText::new(label).color(SUBTEXT))
                                                .sense(Sense::click_and_drag()),
                                        );
                                        if response.double_clicked() {
                                            pending_selection = Some(index);
                                            *selected_sheet_column = SheetColumn::Predecessors;
                                            *selection_anchor = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Predecessors,
                                            });
                                            *editing_cell = Some(SheetCellSelection {
                                                task_index: index,
                                                column: SheetColumn::Predecessors,
                                            });
                                        } else {
                                            Self::update_sheet_selection_from_response(
                                                ui,
                                                &response,
                                                index,
                                                SheetColumn::Predecessors,
                                                &mut pending_selection,
                                                selected_sheet_column,
                                                selection_anchor,
                                                dragging,
                                            );
                                            if response.clicked() {
                                                *editing_cell = None;
                                            }
                                        }
                                    }
                                });

                                row.col(|ui| {
                                    paint_cell(ui, Some(row_base_color), false, false);
                                    let rect = ui.max_rect();
                                    let painter = ui.painter();
                                    paint_chart_row_background(
                                        painter,
                                        rect,
                                        range,
                                        zoom_px_per_day,
                                        show_weekends,
                                    );
                                    let geom = draw_task_bar(
                                        painter,
                                        rect,
                                        &task,
                                        range,
                                        zoom_px_per_day,
                                        pending_selection == Some(index),
                                        show_critical_path,
                                        active_drag_task_uid == Some(task.uid),
                                    );
                                    geometry.borrow_mut().push(geom);
                                });
                            }
                        });
                    });

                let rects = geometry.into_inner();
                if !rects.is_empty() {
                    draw_dependencies(
                        ui.painter(),
                        &rects,
                        &task_by_uid,
                        &document.dependencies,
                        selected_uid,
                        active_drag_task_uid,
                    );
                }
            });
        }

        if pending_selection != selected_task {
            *dirty = true;
        }

        pending_selection
    }

    fn begin_chart_task_drag(
        chart_interaction: &mut Option<ChartInteractionState>,
        document: &ProjectDocument,
        _task_index: usize,
        task_uid: u32,
        kind: TaskDragKind,
        pointer_position: Option<Pos2>,
        dirty_before: bool,
    ) {
        let pointer_origin_x = pointer_position.map(|position| position.x).unwrap_or(0.0);
        *chart_interaction = Some(ChartInteractionState {
            original_document: document.clone(),
            dirty_before,
            kind: ChartInteractionKind::Task {
                task_uid,
                kind,
                pointer_origin_x,
            },
        });
    }

    fn begin_chart_dependency_drag(
        chart_interaction: &mut Option<ChartInteractionState>,
        document: &ProjectDocument,
        _task_index: usize,
        task_uid: u32,
        pointer_position: Option<Pos2>,
        dirty_before: bool,
    ) {
        let pointer_origin = pointer_position.unwrap_or(Pos2::new(0.0, 0.0));
        *chart_interaction = Some(ChartInteractionState {
            original_document: document.clone(),
            dirty_before,
            kind: ChartInteractionKind::Dependency {
                source_task_uid: task_uid,
                pointer_origin,
            },
        });
    }

    fn apply_active_chart_drag_preview(
        document: &mut ProjectDocument,
        chart_interaction: &Option<ChartInteractionState>,
        pointer_position: Option<Pos2>,
        px_per_day: f32,
    ) {
        let Some(state) = chart_interaction.as_ref() else {
            return;
        };

        *document = state.original_document.clone();

        let Some(pointer_position) = pointer_position else {
            return;
        };

        let ChartInteractionKind::Task {
            task_uid,
            kind,
            pointer_origin_x,
        } = state.kind
        else {
            return;
        };

        let delta_days =
            ((pointer_position.x - pointer_origin_x) / px_per_day.max(1.0)).round() as i64;
        if delta_days == 0 {
            return;
        }

        let _ = apply_task_drag_preview(
            document,
            &state.original_document,
            task_uid,
            kind,
            delta_days,
        );
    }

    fn finish_chart_interaction(
        document: &mut ProjectDocument,
        chart_interaction: &mut Option<ChartInteractionState>,
        row_geometries: &[TaskBarGeometry],
        pointer_position: Option<Pos2>,
        pending_selection: &mut Option<usize>,
        selected_sheet_column: &mut SheetColumn,
        selection_anchor: &mut Option<SheetCellSelection>,
        dirty: &mut bool,
        undo_stack: &mut Vec<UndoState>,
        chart_task_editor_refresh: &mut bool,
        dependency_picker: &mut Option<DependencyPickerState>,
        status: &mut String,
    ) {
        let Some(state) = chart_interaction.take() else {
            return;
        };

        match state.kind {
            ChartInteractionKind::Task { task_uid, .. } => {
                let original = state.original_document;
                let original_index = find_task_index_by_uid(&original.tasks, task_uid);
                let current_index = find_task_index_by_uid(&document.tasks, task_uid);
                let changed = match (original_index, current_index) {
                    (Some(original_index), Some(current_index)) => {
                        let original_task = &original.tasks[original_index];
                        let current_task = &document.tasks[current_index];
                        original_task.start != current_task.start
                            || original_task.finish != current_task.finish
                            || original_task.start_text != current_task.start_text
                            || original_task.finish_text != current_task.finish_text
                    }
                    _ => false,
                };

                if changed {
                    push_undo_state(
                        undo_stack,
                        &original,
                        *pending_selection,
                        *selected_sheet_column,
                        *selection_anchor,
                        state.dirty_before,
                    );
                    *dirty = true;
                    *chart_task_editor_refresh = pending_selection
                        .and_then(|task_index| document.tasks.get(task_index))
                        .is_some_and(|task| task.uid == task_uid);
                    *status = "タスク位置を更新しました。".to_string();
                }
            }
            ChartInteractionKind::Dependency {
                source_task_uid,
                pointer_origin,
            } => {
                let pointer = pointer_position.unwrap_or(pointer_origin);
                let Some(target_task_uid) = row_geometries
                    .iter()
                    .find(|geometry| {
                        geometry
                            .bar_rect
                            .expand2(Vec2::splat(4.0))
                            .contains(pointer)
                    })
                    .map(|geometry| geometry.task_uid)
                else {
                    *status = "依存先タスクの上でドロップしてください。".to_string();
                    return;
                };

                if source_task_uid == target_task_uid {
                    *status = "自己依存は作成できません。".to_string();
                    return;
                }

                let Some(source_task_index) =
                    find_task_index_by_uid(&document.tasks, source_task_uid)
                else {
                    return;
                };
                let Some(target_task_index) =
                    find_task_index_by_uid(&document.tasks, target_task_uid)
                else {
                    return;
                };

                let source_task = &document.tasks[source_task_index];
                let target_task = &document.tasks[target_task_index];
                if dependency_exists(document, source_task.uid, target_task.uid) {
                    *dependency_picker = None;
                    *status = "同じ依存関係がすでにあります。".to_string();
                    return;
                }

                *dependency_picker = Some(DependencyPickerState {
                    original_document: state.original_document,
                    source_task_uid: source_task.uid,
                    target_task_uid: target_task.uid,
                    relation: DependencyRelationChoice::FS,
                    lag_text: "0".to_string(),
                    popup_pos: Pos2::new(pointer.x + 12.0, pointer.y + 12.0),
                    dirty_before: state.dirty_before,
                });
            }
        }
    }
}

impl eframe::App for GanttApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = ui.ctx();
        let title = self
            .document_path
            .as_ref()
            .and_then(|path| path.file_name())
            .and_then(|name| name.to_str())
            .map(|name| format!("{name} - MicroProject"))
            .unwrap_or_else(|| "MicroProject".to_string());
        ctx.send_viewport_cmd(egui::ViewportCommand::Title(title));

        if Self::visible_shortcut_open(ctx) {
            self.load_dialog();
        }
        if Self::visible_shortcut_save(ctx) {
            self.save_current_file();
        }
        if Self::visible_shortcut_save_as(ctx) {
            self.save_file_as();
        }
        if Self::visible_shortcut_undo(ctx) && !ctx.egui_wants_keyboard_input() {
            self.undo_last_change();
        }
        if Self::visible_shortcut_redo(ctx) && !ctx.egui_wants_keyboard_input() {
            self.redo_last_change();
        }
        if Self::visible_shortcut_zoom_in(ctx) {
            self.zoom_px_per_day = (self.zoom_px_per_day + 2.0).min(40.0);
        }
        if Self::visible_shortcut_zoom_out(ctx) {
            self.zoom_px_per_day = (self.zoom_px_per_day - 2.0).max(8.0);
        }

        self.handle_drop(ctx);
        ui.set_min_size(ui.available_size());
        let full_rect = ui.max_rect();
        ui.painter().rect_filled(full_rect, 0.0, APP_BG);

        egui::Frame::default()
            .fill(APP_BG)
            .inner_margin(egui::Margin::same(0))
            .show(ui, |ui| {
                self.ribbon(ui);
                ui.add_space(1.0);

                egui::Frame::default()
                    .fill(APP_BG)
                    .inner_margin(egui::Margin::symmetric(2, 2))
                    .show(ui, |ui| {
                        self.document_view(ui);
                    });
            });
    }
}

impl TaskEditorState {
    fn sync_from_task(&mut self, task: &GanttTask) {
        self.task_uid = Some(task.uid);
        self.name = task.name.clone();
        self.outline_level = task.outline_level.max(1);
        self.summary = task.summary;
        self.milestone = task.milestone;
        self.critical = task.critical;
        self.percent_complete = task.percent_complete.clamp(0.0, 100.0);
        self.start_text = task.start_text.clone();
        self.finish_text = task.finish_text.clone();
        self.baseline_start_text = task
            .baseline_start
            .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string())
            .unwrap_or_default();
        self.baseline_finish_text = task
            .baseline_finish
            .map(|value| value.format("%Y-%m-%dT%H:%M:%S").to_string())
            .unwrap_or_default();
        self.duration_text = normalize_duration_text(&task.duration_text);
        self.notes_text = task.notes_text.clone().unwrap_or_default();
        self.resource_names_text = task.resource_names.clone().unwrap_or_default();
        self.calendar_uid_text = task
            .calendar_uid
            .map(|value| value.to_string())
            .unwrap_or_default();
        self.constraint_type_text = task.constraint_type.clone().unwrap_or_default();
    }
}

#[derive(Clone)]
struct TaskBarGeometry {
    task_uid: u32,
    start_x: f32,
    end_x: f32,
    center_y: f32,
    bar_rect: Rect,
    body_rect: Rect,
    start_handle_rect: Option<Rect>,
    finish_handle_rect: Option<Rect>,
    link_handle_rect: Rect,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct SheetSelectionBounds {
    start_row: usize,
    end_row: usize,
    start_column: SheetColumn,
    end_column: SheetColumn,
}

impl SheetSelectionBounds {
    fn contains(&self, row_index: usize, column: SheetColumn) -> bool {
        row_index >= self.start_row
            && row_index <= self.end_row
            && column.index() >= self.start_column.index()
            && column.index() <= self.end_column.index()
    }
}

impl GanttApp {
    fn copy_sheet_selection_to_tsv(&self, visible_indices: &[usize]) -> Option<String> {
        let document = self.document.as_ref()?;
        let bounds = self.sheet_selection_bounds(visible_indices)?;
        let mut lines = Vec::new();

        for row_index in bounds.start_row..=bounds.end_row {
            let task_index = *visible_indices.get(row_index)?;
            let task = document.tasks.get(task_index)?;
            let mut cells = Vec::new();

            for column_index in bounds.start_column.index()..=bounds.end_column.index() {
                let column = SheetColumn::from_index(column_index)?;
                cells.push(Self::sheet_cell_text(task, column));
            }

            lines.push(cells.join("\t"));
        }

        Some(lines.join("\r\n"))
    }

    fn paste_start_cell(&self, visible_indices: &[usize]) -> Option<(usize, SheetColumn)> {
        let bounds = self.sheet_selection_bounds(visible_indices)?;
        let start_column = if bounds.start_column == SheetColumn::Mode {
            SheetColumn::Name
        } else {
            bounds.start_column
        };
        Some((bounds.start_row, start_column))
    }

    fn sheet_selection_bounds(&self, visible_indices: &[usize]) -> Option<SheetSelectionBounds> {
        let active = self.selected_sheet_cell()?;
        let anchor = self.sheet_selection_anchor.unwrap_or(active);
        let active_row = visible_indices
            .iter()
            .position(|index| *index == active.task_index)?;
        let anchor_row = visible_indices
            .iter()
            .position(|index| *index == anchor.task_index)?;

        let start_row = active_row.min(anchor_row);
        let end_row = active_row.max(anchor_row);
        let start_column = if active.column.index() <= anchor.column.index() {
            active.column
        } else {
            anchor.column
        };
        let end_column = if active.column.index() >= anchor.column.index() {
            active.column
        } else {
            anchor.column
        };

        Some(SheetSelectionBounds {
            start_row,
            end_row,
            start_column,
            end_column,
        })
    }

    fn sheet_cell_text(task: &GanttTask, column: SheetColumn) -> String {
        match column {
            SheetColumn::Mode => derive_task_mode(task).to_string(),
            SheetColumn::Name => task.name.clone(),
            SheetColumn::Duration => task.duration_text.clone(),
            SheetColumn::Start => task.start_text.clone(),
            SheetColumn::Finish => task.finish_text.clone(),
            SheetColumn::PercentComplete => {
                if (task.percent_complete.fract()).abs() < f32::EPSILON {
                    format!("{:.0}", task.percent_complete)
                } else {
                    format!("{}", task.percent_complete)
                }
            }
            SheetColumn::Predecessors => task.predecessor_text.clone(),
        }
    }

    fn sheet_selection_bounds_for_render(
        visible_indices: &[usize],
        selected_task: Option<usize>,
        selected_sheet_column: SheetColumn,
        selection_anchor: Option<SheetCellSelection>,
    ) -> Option<SheetSelectionBounds> {
        let active = selected_task.map(|task_index| SheetCellSelection {
            task_index,
            column: selected_sheet_column,
        })?;
        let anchor = selection_anchor.unwrap_or(active);
        let active_row = visible_indices
            .iter()
            .position(|index| *index == active.task_index)?;
        let anchor_row = visible_indices
            .iter()
            .position(|index| *index == anchor.task_index)?;

        let start_row = active_row.min(anchor_row);
        let end_row = active_row.max(anchor_row);
        let start_column = if active.column.index() <= anchor.column.index() {
            active.column
        } else {
            anchor.column
        };
        let end_column = if active.column.index() >= anchor.column.index() {
            active.column
        } else {
            anchor.column
        };

        Some(SheetSelectionBounds {
            start_row,
            end_row,
            start_column,
            end_column,
        })
    }

    fn parse_tsv_rows(text: &str) -> Vec<Vec<&str>> {
        text.split('\n')
            .filter(|row| !row.trim().is_empty())
            .map(|row| row.split('\t').collect())
            .collect()
    }

    fn normalize_excel_clipboard_text(text: &str) -> String {
        text.trim_start_matches('\u{feff}')
            .replace("\r\n", "\n")
            .replace('\r', "\n")
    }

    fn is_mode_label(text: &str) -> bool {
        matches!(
            text.trim(),
            "Summary" | "Milestone" | "Auto Scheduled" | "Manually Scheduled"
        )
    }

    fn update_sheet_selection_from_response(
        ui: &egui::Ui,
        response: &egui::Response,
        task_index: usize,
        column: SheetColumn,
        pending_selection: &mut Option<usize>,
        selected_sheet_column: &mut SheetColumn,
        selection_anchor: &mut Option<SheetCellSelection>,
        dragging: &mut bool,
    ) {
        let shift = ui.input(|input| input.modifiers.shift);
        let pointer_down = ui.input(|input| input.pointer.primary_down());
        let editing = ui.ctx().egui_wants_keyboard_input();

        if response.drag_started() {
            *dragging = true;
            *pending_selection = Some(task_index);
            *selected_sheet_column = column;
            *selection_anchor = Some(SheetCellSelection { task_index, column });
            ui.ctx().request_repaint();
        } else if response.clicked() && !*dragging {
            *pending_selection = Some(task_index);
            *selected_sheet_column = column;
            if !shift || selection_anchor.is_none() {
                *selection_anchor = Some(SheetCellSelection { task_index, column });
            }
            ui.ctx().request_repaint();
        }

        if !editing && pointer_down && response.hovered() {
            *pending_selection = Some(task_index);
            *selected_sheet_column = column;
            if *dragging {
                if selection_anchor.is_none() {
                    *selection_anchor = Some(SheetCellSelection { task_index, column });
                }
            } else if selection_anchor.is_none() {
                *selection_anchor = Some(SheetCellSelection { task_index, column });
            }
            ui.ctx().request_repaint();
        }

        if !pointer_down {
            *dragging = false;
        }
    }

    fn sheet_row_target_index(visible_indices: &[usize], visible_row: usize) -> Option<usize> {
        visible_indices
            .get(visible_row)
            .copied()
            .or_else(|| visible_indices.last().copied())
    }
}

fn build_visible_task_indices(
    tasks: &[GanttTask],
    filter_text: &str,
    collapsed_task_uids: &HashSet<u32>,
) -> Vec<usize> {
    let filter = filter_text.trim().to_lowercase();
    if filter.is_empty() && collapsed_task_uids.is_empty() {
        return (0..tasks.len()).collect();
    }

    if filter.is_empty() {
        let mut visible = Vec::new();
        let mut ancestor_stack: Vec<(usize, u32)> = Vec::new();
        for (index, task) in tasks.iter().enumerate() {
            let level = task.outline_level.max(1);
            while ancestor_stack
                .last()
                .is_some_and(|(_, ancestor_level)| *ancestor_level >= level)
            {
                ancestor_stack.pop();
            }
            let hidden = ancestor_stack.iter().any(|(ancestor_index, _)| {
                collapsed_task_uids.contains(&tasks[*ancestor_index].uid)
            });
            if !hidden {
                visible.push(index);
            }
            ancestor_stack.push((index, level));
        }
        return visible;
    }

    let mut visible = HashSet::<usize>::new();
    let mut ancestor_stack: Vec<(usize, u32)> = Vec::new();

    for (index, task) in tasks.iter().enumerate() {
        let level = task.outline_level.max(1);
        while ancestor_stack
            .last()
            .is_some_and(|(_, ancestor_level)| *ancestor_level >= level)
        {
            ancestor_stack.pop();
        }
        if task_matches_filter(task, &filter) {
            visible.insert(index);
            for (ancestor_index, _) in &ancestor_stack {
                visible.insert(*ancestor_index);
            }
        }
        ancestor_stack.push((index, level));
    }

    let mut result = visible.into_iter().collect::<Vec<_>>();
    result.sort_unstable();
    result
}

fn task_matches_filter(task: &GanttTask, filter: &str) -> bool {
    if filter.is_empty() {
        return true;
    }

    let haystacks = [
        task.name.as_str(),
        task.duration_text.as_str(),
        task.predecessor_text.as_str(),
        task.start_text.as_str(),
        task.finish_text.as_str(),
        task.notes_text.as_deref().unwrap_or_default(),
        task.resource_names.as_deref().unwrap_or_default(),
        task.constraint_type.as_deref().unwrap_or_default(),
    ];

    haystacks
        .iter()
        .any(|value| value.to_lowercase().contains(filter))
}

fn derive_task_mode(task: &GanttTask) -> &'static str {
    if task.summary {
        "Summary"
    } else if task.milestone {
        "Milestone"
    } else if task.start.is_some() && task.finish.is_some() {
        "Auto Scheduled"
    } else {
        "Manually Scheduled"
    }
}

#[derive(Debug, Clone)]
struct PredecessorAssignment {
    predecessor_id: u32,
    relation: String,
    lag_text: Option<String>,
}

fn parse_predecessor_assignments(text: &str) -> Vec<PredecessorAssignment> {
    text.split([',', ';'])
        .filter_map(|item| {
            let token = item.trim();
            if token.is_empty() {
                return None;
            }

            let digit_len = token.chars().take_while(|ch| ch.is_ascii_digit()).count();
            if digit_len == 0 {
                return None;
            }

            let predecessor_id = token[..digit_len].parse::<u32>().ok()?;
            let rest = token[digit_len..].trim();
            if rest.is_empty() {
                return Some(PredecessorAssignment {
                    predecessor_id,
                    relation: "FS".to_string(),
                    lag_text: None,
                });
            }

            let relation_len = rest
                .chars()
                .take_while(|ch| ch.is_ascii_alphabetic() || ch.is_ascii_digit())
                .count();
            let relation = if relation_len == 0 {
                "FS".to_string()
            } else {
                rest[..relation_len].trim().to_string()
            };
            let lag = rest[relation_len..].trim();
            let lag_text = (!lag.is_empty()).then(|| lag.to_string());

            Some(PredecessorAssignment {
                predecessor_id,
                relation,
                lag_text,
            })
        })
        .collect()
}

fn normalize_duration_text(text: &str) -> String {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return String::new();
    }

    if let Some(days) = parse_duration_days(trimmed) {
        return format!("{days}日");
    }

    trimmed.to_string()
}

fn parse_duration_days(text: &str) -> Option<u32> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }

    let lower = trimmed.to_ascii_lowercase();

    let parse_prefix_number = |suffix: &str| -> Option<u32> {
        let number = lower.strip_suffix(suffix)?.trim();
        number.parse::<u32>().ok()
    };

    if let Ok(days) = trimmed.parse::<u32>() {
        return Some(days);
    }

    for suffix in ["日", "d", "day", "days"] {
        if let Some(days) = parse_prefix_number(suffix) {
            return Some(days);
        }
    }

    if let Some(months) = parse_prefix_number("month").or_else(|| parse_prefix_number("months")) {
        return Some(months.saturating_mul(20));
    }

    if let Some(years) = parse_prefix_number("year").or_else(|| parse_prefix_number("years")) {
        return Some(years.saturating_mul(365));
    }

    if let Some(body) = lower.strip_prefix("pt") {
        let mut hours = 0.0f32;
        let mut minutes = 0.0f32;
        let mut seconds = 0.0f32;

        let mut remaining = body;
        if let Some((prefix, rest)) = remaining.split_once('h') {
            hours = prefix.trim().parse::<f32>().ok()?;
            remaining = rest;
        }
        if let Some((prefix, rest)) = remaining.split_once('m') {
            minutes = prefix.trim().parse::<f32>().ok()?;
            remaining = rest;
        }
        if let Some((prefix, _rest)) = remaining.split_once('s') {
            seconds = prefix.trim().parse::<f32>().ok()?;
        }

        let total_hours = hours + minutes / 60.0 + seconds / 3600.0;
        if total_hours > 0.0 {
            return Some((total_hours / 8.0).ceil().max(1.0) as u32);
        }
    }

    if let Some(days) = lower
        .strip_prefix('p')
        .and_then(|rest| rest.strip_suffix('d'))
    {
        return days.trim().parse::<u32>().ok();
    }

    None
}

fn display_duration_text(text: &str) -> String {
    normalize_duration_text(text)
}

fn find_task_by_id<'a>(tasks: &'a [GanttTask], id: u32) -> Option<&'a GanttTask> {
    tasks.iter().find(|task| task.id == id)
}

fn apply_task_editor(
    document: &mut ProjectDocument,
    index: usize,
    task_editor: &TaskEditorState,
) -> bool {
    let Some(task) = document.tasks.get_mut(index) else {
        return false;
    };

    task.name = task_editor.name.trim().to_string();
    task.outline_level = task_editor.outline_level.max(1);
    task.summary = task_editor.summary;
    task.milestone = task_editor.milestone;
    task.critical = task_editor.critical;
    task.percent_complete = task_editor.percent_complete.clamp(0.0, 100.0);
    task.start_text = task_editor.start_text.trim().to_string();
    task.finish_text = task_editor.finish_text.trim().to_string();
    task.start = parse_date_time(Some(&task.start_text));
    task.finish = parse_date_time(Some(&task.finish_text));
    task.baseline_start = parse_date_time(Some(&task_editor.baseline_start_text));
    task.baseline_finish = parse_date_time(Some(&task_editor.baseline_finish_text));
    task.duration_text = normalize_duration_text(&task_editor.duration_text);
    task.notes_text =
        (!task_editor.notes_text.trim().is_empty()).then(|| task_editor.notes_text.clone());
    task.resource_names = (!task_editor.resource_names_text.trim().is_empty())
        .then(|| task_editor.resource_names_text.clone());
    task.calendar_uid = task_editor.calendar_uid_text.trim().parse::<u32>().ok();
    task.constraint_type = (!task_editor.constraint_type_text.trim().is_empty())
        .then(|| task_editor.constraint_type_text.clone());
    true
}

fn push_undo_state(
    undo_stack: &mut Vec<UndoState>,
    document: &ProjectDocument,
    selected_task: Option<usize>,
    selected_sheet_column: SheetColumn,
    sheet_selection_anchor: Option<SheetCellSelection>,
    dirty: bool,
) {
    undo_stack.push(UndoState {
        document: document.clone(),
        selected_task,
        selected_sheet_column,
        sheet_selection_anchor,
        dirty,
    });

    if undo_stack.len() > UNDO_STACK_LIMIT {
        let overflow = undo_stack.len() - UNDO_STACK_LIMIT;
        undo_stack.drain(0..overflow);
    }
}

fn paint_sheet_cell(
    painter: &egui::Painter,
    rect: Rect,
    fill: Color32,
    selected: bool,
    active: bool,
) {
    painter.rect_filled(rect, 0.0, fill);
    painter.line_segment(
        [
            Pos2::new(rect.left(), rect.bottom() - 0.5),
            Pos2::new(rect.right(), rect.bottom() - 0.5),
        ],
        Stroke::new(1.0, GRID_LINE),
    );
    painter.line_segment(
        [
            Pos2::new(rect.right() - 0.5, rect.top()),
            Pos2::new(rect.right() - 0.5, rect.bottom()),
        ],
        Stroke::new(1.0, GRID_LINE),
    );

    if active {
        painter.rect_stroke(
            rect.shrink(0.5),
            0.0,
            Stroke::new(1.5, SHEET_ACTIVE_BORDER),
            egui::StrokeKind::Inside,
        );
    } else if selected {
        painter.rect_stroke(
            rect.shrink(0.5),
            0.0,
            Stroke::new(1.0, SHEET_SELECTION_BORDER),
            egui::StrokeKind::Inside,
        );
    } else {
        painter.rect_stroke(
            rect.shrink(0.5),
            0.0,
            Stroke::new(0.75, GRID_LINE_STRONG),
            egui::StrokeKind::Inside,
        );
    }
}

fn paint_sheet_header_cell(painter: &egui::Painter, rect: Rect, title: &str, glyph: HeaderGlyph) {
    painter.rect_filled(rect, 0.0, HEADER_BG);
    painter.rect_filled(
        Rect::from_min_max(rect.min, Pos2::new(rect.right(), rect.center().y)),
        0.0,
        SHEET_HEADER_TOP,
    );
    painter.line_segment(
        [
            Pos2::new(rect.left(), rect.top()),
            Pos2::new(rect.right(), rect.top()),
        ],
        Stroke::new(1.0, Color32::WHITE),
    );
    painter.line_segment(
        [
            Pos2::new(rect.left(), rect.bottom() - 0.5),
            Pos2::new(rect.right(), rect.bottom() - 0.5),
        ],
        Stroke::new(1.0, SHEET_HEADER_BOTTOM),
    );
    painter.line_segment(
        [
            Pos2::new(rect.right() - 0.5, rect.top()),
            Pos2::new(rect.right() - 0.5, rect.bottom()),
        ],
        Stroke::new(1.0, SHEET_HEADER_BOTTOM),
    );
    draw_header_glyph(
        painter,
        Rect::from_min_size(
            Pos2::new(rect.left() + 6.0, rect.top() + 5.0),
            Vec2::new(13.0, 13.0),
        ),
        glyph,
    );
    if !title.is_empty() {
        painter.text(
            Pos2::new(rect.left() + 23.0, rect.center().y - 1.0),
            egui::Align2::LEFT_CENTER,
            title,
            FontId::new(10.5, FontFamily::Proportional),
            SHEET_HEADER_TEXT,
        );
    }
}

fn draw_header_glyph(painter: &egui::Painter, rect: Rect, glyph: HeaderGlyph) {
    let ink = Color32::from_rgb(97, 97, 97);
    let accent = Color32::from_rgb(0, 84, 147);
    let teal = Color32::from_rgb(108, 108, 108);
    let orange = Color32::from_rgb(167, 106, 22);
    let purple = Color32::from_rgb(118, 118, 118);
    let gray = Color32::from_rgb(108, 108, 108);

    match glyph {
        HeaderGlyph::Mode => {
            painter.rect_filled(rect.shrink(1.0), 2.0, Color32::from_rgb(239, 243, 247));
            painter.rect_stroke(
                rect.shrink(1.0),
                2.0,
                Stroke::new(1.0, ink),
                egui::StrokeKind::Inside,
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 3.0, rect.top() + 4.0),
                    Pos2::new(rect.right() - 3.0, rect.top() + 4.0),
                ],
                Stroke::new(1.0, gray),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 3.0, rect.center().y),
                    Pos2::new(rect.right() - 5.0, rect.center().y),
                ],
                Stroke::new(1.5, accent),
            );
            painter.circle_filled(
                Pos2::new(rect.left() + 4.5, rect.bottom() - 4.5),
                1.8,
                orange,
            );
        }
        HeaderGlyph::Name => {
            for i in 0..3 {
                let y = rect.top() + 3.0 + i as f32 * 4.0;
                painter.line_segment(
                    [
                        Pos2::new(rect.left() + 2.0, y),
                        Pos2::new(rect.right() - 2.0, y),
                    ],
                    Stroke::new(if i == 1 { 1.6 } else { 1.0 }, teal),
                );
            }
        }
        HeaderGlyph::Duration => {
            painter.circle_stroke(rect.center(), 5.0, Stroke::new(1.2, orange));
            painter.line_segment(
                [
                    rect.center(),
                    Pos2::new(rect.center().x + 3.0, rect.center().y - 2.0),
                ],
                Stroke::new(1.2, orange),
            );
            painter.line_segment(
                [
                    rect.center(),
                    Pos2::new(rect.center().x, rect.center().y - 4.0),
                ],
                Stroke::new(1.2, orange),
            );
        }
        HeaderGlyph::Date => {
            let body = rect.shrink(1.0);
            painter.rect_filled(body, 1.5, Color32::from_rgb(245, 248, 252));
            painter.rect_stroke(body, 1.5, Stroke::new(1.0, teal), egui::StrokeKind::Inside);
            painter.rect_filled(
                Rect::from_min_size(body.min, Vec2::new(body.width(), 3.0)),
                1.0,
                teal,
            );
            painter.circle_filled(
                Pos2::new(body.left() + 4.0, body.bottom() - 4.0),
                1.4,
                purple,
            );
        }
        HeaderGlyph::Percent => {
            let base = rect.shrink(1.5);
            painter.rect_stroke(base, 1.5, Stroke::new(1.0, gray), egui::StrokeKind::Inside);
            painter.rect_filled(
                Rect::from_min_max(
                    Pos2::new(base.left() + 2.0, base.bottom() - 4.0),
                    Pos2::new(base.left() + 5.0, base.bottom() - 2.0),
                ),
                0.5,
                gray,
            );
            painter.rect_filled(
                Rect::from_min_max(
                    Pos2::new(base.left() + 7.0, base.bottom() - 7.0),
                    Pos2::new(base.left() + 10.0, base.bottom() - 2.0),
                ),
                0.5,
                teal,
            );
            painter.rect_filled(
                Rect::from_min_max(
                    Pos2::new(base.left() + 12.0, base.top() + 4.0),
                    Pos2::new(base.left() + 15.0, base.bottom() - 2.0),
                ),
                0.5,
                orange,
            );
        }
        HeaderGlyph::Predecessors => {
            painter.circle_filled(Pos2::new(rect.left() + 4.0, rect.center().y), 2.0, purple);
            painter.circle_filled(Pos2::new(rect.right() - 4.0, rect.center().y), 2.0, teal);
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 6.0, rect.center().y),
                    Pos2::new(rect.right() - 7.0, rect.center().y),
                ],
                Stroke::new(1.2, ink),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.right() - 9.0, rect.center().y - 2.0),
                    Pos2::new(rect.right() - 7.0, rect.center().y),
                ],
                Stroke::new(1.2, ink),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.right() - 9.0, rect.center().y + 2.0),
                    Pos2::new(rect.right() - 7.0, rect.center().y),
                ],
                Stroke::new(1.2, ink),
            );
        }
        HeaderGlyph::Gantt => {
            painter.rect_stroke(
                rect.shrink(1.0),
                1.5,
                Stroke::new(1.0, gray),
                egui::StrokeKind::Inside,
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 3.0, rect.top() + 4.0),
                    Pos2::new(rect.left() + 10.0, rect.top() + 4.0),
                ],
                Stroke::new(1.4, teal),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 3.0, rect.center().y),
                    Pos2::new(rect.left() + 11.0, rect.center().y),
                ],
                Stroke::new(1.4, accent),
            );
            painter.line_segment(
                [
                    Pos2::new(rect.left() + 3.0, rect.bottom() - 4.0),
                    Pos2::new(rect.left() + 8.0, rect.bottom() - 4.0),
                ],
                Stroke::new(1.4, orange),
            );
        }
    }
}

fn draw_mode_badge(painter: &egui::Painter, rect: Rect, color: Color32, task: &GanttTask) {
    painter.rect_filled(rect, 2.5, Color32::from_rgb(245, 247, 249));
    painter.rect_stroke(rect, 2.5, Stroke::new(1.0, color), egui::StrokeKind::Inside);

    if task.summary {
        painter.line_segment(
            [
                Pos2::new(rect.left() + 3.0, rect.center().y - 2.5),
                Pos2::new(rect.right() - 3.0, rect.center().y - 2.5),
            ],
            Stroke::new(1.2, color),
        );
        painter.line_segment(
            [
                Pos2::new(rect.left() + 3.0, rect.center().y + 2.0),
                Pos2::new(rect.right() - 3.0, rect.center().y + 2.0),
            ],
            Stroke::new(1.2, color),
        );
    } else if task.milestone {
        let center = rect.center();
        painter.add(egui::Shape::convex_polygon(
            vec![
                Pos2::new(center.x, rect.top() + 2.0),
                Pos2::new(rect.right() - 2.0, center.y),
                Pos2::new(center.x, rect.bottom() - 2.0),
                Pos2::new(rect.left() + 2.0, center.y),
            ],
            color,
            Stroke::new(1.0, color),
        ));
    } else if task.critical {
        painter.circle_filled(rect.center(), 3.0, color);
        painter.line_segment(
            [
                Pos2::new(rect.left() + 3.0, rect.top() + 4.0),
                Pos2::new(rect.right() - 3.0, rect.bottom() - 4.0),
            ],
            Stroke::new(1.0, color),
        );
    } else {
        painter.line_segment(
            [
                Pos2::new(rect.left() + 3.0, rect.center().y),
                Pos2::new(rect.right() - 4.0, rect.center().y),
            ],
            Stroke::new(1.3, color),
        );
        painter.line_segment(
            [
                Pos2::new(rect.right() - 5.0, rect.center().y - 2.0),
                Pos2::new(rect.right() - 3.0, rect.center().y),
            ],
            Stroke::new(1.3, color),
        );
        painter.line_segment(
            [
                Pos2::new(rect.right() - 5.0, rect.center().y + 2.0),
                Pos2::new(rect.right() - 3.0, rect.center().y),
            ],
            Stroke::new(1.3, color),
        );
    }
}

// draw_mode_chip was removed - it was only used by task_mode_strip which was deleted

fn paint_chart_row_background(
    painter: &egui::Painter,
    rect: Rect,
    range: ChartRange,
    px_per_day: f32,
    show_weekends: bool,
) {
    painter.rect_filled(rect, 0.0, SHEET_BG);
    let days = range.days();
    let today = chrono::Local::now().date_naive();
    let today_offset = (today - range.start).num_days();
    if today_offset >= 0 && today_offset <= days {
        let x = rect.left() + today_offset as f32 * px_per_day;
        let today_band = Rect::from_min_max(
            Pos2::new((x - 1.0).max(rect.left()), rect.top()),
            Pos2::new((x + 1.0).min(rect.right()), rect.bottom()),
        );
        painter.rect_filled(
            today_band,
            0.0,
            Color32::from_rgba_unmultiplied(242, 171, 61, 26),
        );
        painter.line_segment(
            [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
            Stroke::new(2.0, TODAY),
        );
    }

    for day in 0..=days {
        let date = range.start + Duration::days(day);
        let x = rect.left() + day as f32 * px_per_day;
        if show_weekends {
            let weekday = date.weekday().number_from_monday();
            if weekday == 6 || weekday == 7 {
                let weekend_rect = Rect::from_min_max(
                    Pos2::new(x, rect.top()),
                    Pos2::new(x + px_per_day, rect.bottom()),
                );
                painter.rect_filled(weekend_rect, 0.0, WEEKEND_BG);
            }
        }
        let line_color = if date.day() == 1 {
            Color32::from_rgb(158, 158, 158)
        } else if day % 7 == 0 {
            Color32::from_rgb(188, 188, 188)
        } else {
            Color32::from_rgb(228, 228, 228)
        };
        let line_width = if date.day() == 1 {
            1.4
        } else if day % 7 == 0 {
            1.2
        } else {
            1.0
        };
        painter.line_segment(
            [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
            Stroke::new(line_width, line_color),
        );
    }
}

fn draw_task_bar(
    painter: &egui::Painter,
    rect: Rect,
    task: &GanttTask,
    range: ChartRange,
    px_per_day: f32,
    selected: bool,
    show_critical_path: bool,
    active_drag_preview: bool,
) -> TaskBarGeometry {
    let row_center_y = rect.center().y;
    let start_day = task.start.map(|dt| dt.date()).unwrap_or(range.start);
    let finish_day = task.finish.map(|dt| dt.date()).unwrap_or(start_day);

    let start_offset_f64 = (start_day - range.start).num_days().max(0) as f64;
    let finish_offset_f64 = (finish_day - range.start).num_days().max(0) as f64 + 1.0;
    let px_per_day_f64 = px_per_day as f64;
    let rect_left_f64 = rect.left() as f64;

    let mut start_x = rect_left_f64 + start_offset_f64 * px_per_day_f64;
    let mut end_x = rect_left_f64 + finish_offset_f64 * px_per_day_f64;
    if end_x < start_x {
        std::mem::swap(&mut start_x, &mut end_x);
    }
    let min_width = 1.5;
    if (end_x - start_x).abs() < min_width {
        end_x = start_x + min_width;
    }

    let start_x = start_x as f32;
    let end_x = end_x as f32;

    let is_critical = task.critical && show_critical_path;
    let base_color = if task.summary {
        Color32::from_rgb(55, 55, 55)
    } else if task.milestone {
        if is_critical {
            CRITICAL
        } else {
            Color32::from_rgb(212, 72, 48)
        }
    } else if is_critical {
        CRITICAL
    } else {
        Color32::from_rgb(65, 103, 174)
    };
    let preview_alpha = if active_drag_preview { 170 } else { 255 };
    let accent_alpha = if active_drag_preview { 190 } else { 255 };
    let preview_color = |color: Color32, alpha: u8| {
        Color32::from_rgba_unmultiplied(color.r(), color.g(), color.b(), alpha)
    };

    let baseline_rect = task_baseline_rect(task, range, px_per_day, row_center_y, rect.left());
    if let Some(baseline) = baseline_rect {
        painter.rect_filled(
            baseline,
            0.0,
            preview_color(BASELINE, if active_drag_preview { 190 } else { 255 }),
        );
    }

    let bar_rect = if task.milestone {
        let center = Pos2::new(start_x, row_center_y);
        let points = [
            Pos2::new(center.x, center.y - 7.0),
            Pos2::new(center.x + 7.0, center.y),
            Pos2::new(center.x, center.y + 7.0),
            Pos2::new(center.x - 7.0, center.y),
        ];
        painter.add(egui::Shape::convex_polygon(
            points.to_vec(),
            preview_color(base_color, preview_alpha),
            Stroke::new(1.1, if selected { ACCENT } else { base_color }),
        ));
        Rect::from_center_size(center, Vec2::new(14.0, 14.0))
    } else if task.summary {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - 4.0),
            Pos2::new(end_x, row_center_y + 4.0),
        );
        painter.rect_filled(bar_rect, 1.0, preview_color(base_color, preview_alpha));
        painter.rect_stroke(
            bar_rect,
            1.0,
            Stroke::new(
                if selected { 1.5 } else { 1.1 },
                if selected { ACCENT } else { base_color },
            ),
            egui::StrokeKind::Inside,
        );
        painter.line_segment(
            [
                Pos2::new(bar_rect.left(), bar_rect.bottom() + 2.0),
                Pos2::new(bar_rect.right(), bar_rect.bottom() + 2.0),
            ],
            Stroke::new(1.5, base_color),
        );
        bar_rect
    } else {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - 6.0),
            Pos2::new(end_x, row_center_y + 6.0),
        );
        painter.rect_filled(
            bar_rect,
            1.0,
            preview_color(Color32::from_rgb(227, 232, 242), preview_alpha),
        );
        painter.rect_filled(
            Rect::from_min_max(
                bar_rect.left_top(),
                Pos2::new(
                    bar_rect.left()
                        + bar_rect.width() * (task.percent_complete / 100.0).clamp(0.0, 1.0),
                    bar_rect.bottom(),
                ),
            ),
            2.0,
            preview_color(base_color, accent_alpha),
        );
        painter.rect_stroke(
            bar_rect,
            1.0,
            Stroke::new(
                if selected { 1.5 } else { 1.0 },
                if selected {
                    ACCENT
                } else {
                    Color32::from_rgb(64, 94, 150)
                },
            ),
            egui::StrokeKind::Inside,
        );
        bar_rect
    };

    let mut start_handle_rect = None;
    let mut finish_handle_rect = None;
    let body_rect = if task.milestone {
        bar_rect
    } else {
        let handle_width = TASK_BAR_HANDLE_WIDTH.max(4.0);
        let handle_height = TASK_BAR_HANDLE_HEIGHT.max(10.0);
        let left = (bar_rect.left() + handle_width).min(bar_rect.right());
        let right = (bar_rect.right() - handle_width).max(left);
        let start_rect = Rect::from_center_size(
            Pos2::new(bar_rect.left(), row_center_y),
            Vec2::new(handle_width * 2.0, handle_height),
        );
        let finish_rect = Rect::from_center_size(
            Pos2::new(bar_rect.right(), row_center_y),
            Vec2::new(handle_width * 2.0, handle_height),
        );
        start_handle_rect = Some(start_rect);
        finish_handle_rect = Some(finish_rect);
        Rect::from_min_max(
            Pos2::new(left, bar_rect.top()),
            Pos2::new(right, bar_rect.bottom()),
        )
    };
    let link_handle_rect = Rect::from_center_size(
        Pos2::new(bar_rect.right() + TASK_BAR_LINK_HANDLE_GAP, row_center_y),
        Vec2::new(TASK_BAR_HANDLE_WIDTH * 2.0, TASK_BAR_HANDLE_HEIGHT),
    );

    if let Some(handle_rect) = start_handle_rect {
        let grip_color = if selected {
            preview_color(ACCENT, accent_alpha)
        } else {
            preview_color(Color32::from_rgb(96, 128, 184), preview_alpha)
        };
        painter.rect_filled(
            Rect::from_center_size(handle_rect.center(), Vec2::new(4.0, 12.0)),
            1.0,
            preview_color(Color32::WHITE, preview_alpha),
        );
        painter.line_segment(
            [
                Pos2::new(handle_rect.center().x, handle_rect.top() + 1.0),
                Pos2::new(handle_rect.center().x, handle_rect.bottom() - 1.0),
            ],
            Stroke::new(1.6, grip_color),
        );
        painter.circle_filled(handle_rect.center(), 3.1, grip_color);
    }
    if let Some(handle_rect) = finish_handle_rect {
        let grip_color = if selected {
            preview_color(ACCENT, accent_alpha)
        } else {
            preview_color(Color32::from_rgb(96, 128, 184), preview_alpha)
        };
        painter.rect_filled(
            Rect::from_center_size(handle_rect.center(), Vec2::new(4.0, 12.0)),
            1.0,
            preview_color(Color32::WHITE, preview_alpha),
        );
        painter.line_segment(
            [
                Pos2::new(handle_rect.center().x, handle_rect.top() + 1.0),
                Pos2::new(handle_rect.center().x, handle_rect.bottom() - 1.0),
            ],
            Stroke::new(1.6, grip_color),
        );
        painter.circle_filled(handle_rect.center(), 3.1, grip_color);
    }
    painter.line_segment(
        [
            Pos2::new(bar_rect.right() + 1.5, row_center_y),
            Pos2::new(link_handle_rect.left() - 2.0, row_center_y),
        ],
        Stroke::new(
            1.0,
            preview_color(Color32::from_rgb(128, 128, 128), preview_alpha),
        ),
    );
    painter.circle_filled(
        link_handle_rect.center(),
        3.2,
        if selected {
            preview_color(ACCENT, accent_alpha)
        } else {
            preview_color(Color32::from_rgb(90, 90, 90), preview_alpha)
        },
    );
    painter.circle_stroke(
        link_handle_rect.center(),
        3.2,
        Stroke::new(1.0, preview_color(Color32::WHITE, preview_alpha)),
    );

    TaskBarGeometry {
        task_uid: task.uid,
        start_x,
        end_x,
        center_y: row_center_y,
        bar_rect,
        body_rect,
        start_handle_rect,
        finish_handle_rect,
        link_handle_rect,
    }
}

fn task_baseline_rect(
    task: &GanttTask,
    range: ChartRange,
    px_per_day: f32,
    row_center_y: f32,
    left: f32,
) -> Option<Rect> {
    let (Some(baseline_start), Some(baseline_finish)) = (task.baseline_start, task.baseline_finish)
    else {
        return None;
    };

    let baseline_start_day = baseline_start.date();
    let baseline_finish_day = baseline_finish.date();
    let baseline_start_offset_f64 = (baseline_start_day - range.start).num_days().max(0) as f64;
    let baseline_finish_offset_f64 =
        (baseline_finish_day - range.start).num_days().max(0) as f64 + 1.0;
    let px_per_day_f64 = px_per_day as f64;
    let left_f64 = left as f64;
    let row_center_y_f64 = row_center_y as f64;
    Some(Rect::from_min_max(
        Pos2::new(
            (left_f64 + baseline_start_offset_f64 * px_per_day_f64) as f32,
            (row_center_y_f64 - 7.0) as f32,
        ),
        Pos2::new(
            (left_f64 + baseline_finish_offset_f64 * px_per_day_f64) as f32,
            (row_center_y_f64 - 5.0) as f32,
        ),
    ))
}

fn date_x(date: NaiveDate, range: ChartRange, chart_left: f32, px_per_day: f32) -> f32 {
    let offset_days = (date - range.start).num_days() as f32;
    chart_left + offset_days * px_per_day
}

fn status_date_x(status_date: NaiveDate, range: ChartRange, chart_left: f32, px_per_day: f32) -> f32 {
    date_x(status_date, range, chart_left, px_per_day)
}

fn first_day_of_month(date: NaiveDate) -> NaiveDate {
    NaiveDate::from_ymd_opt(date.year(), date.month(), 1).unwrap_or(date)
}

fn next_month_start(date: NaiveDate) -> NaiveDate {
    let year = if date.month() == 12 {
        date.year() + 1
    } else {
        date.year()
    };
    let month = if date.month() == 12 { 1 } else { date.month() + 1 };
    NaiveDate::from_ymd_opt(year, month, 1).unwrap_or(date)
}

fn prev_month_start(date: NaiveDate) -> NaiveDate {
    let year = if date.month() == 1 {
        date.year() - 1
    } else {
        date.year()
    };
    let month = if date.month() == 1 { 12 } else { date.month() - 1 };
    NaiveDate::from_ymd_opt(year, month, 1).unwrap_or(date)
}

fn days_in_month(year: i32, month: u32) -> u32 {
    let first = NaiveDate::from_ymd_opt(year, month, 1).unwrap();
    let next = if month == 12 {
        NaiveDate::from_ymd_opt(year + 1, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(year, month + 1, 1).unwrap()
    };
    (next - first).num_days() as u32
}

#[derive(Debug, Clone)]
struct AppSettings {
    status_date: NaiveDate,
}

fn parse_app_settings_text(text: &str) -> Option<AppSettings> {
    for line in text.lines() {
        let line = line.trim();
        if let Some(value) = line.strip_prefix("status_date=") {
            if let Ok(status_date) = NaiveDate::parse_from_str(value.trim(), "%Y-%m-%d") {
                return Some(AppSettings { status_date });
            }
        }
    }

    None
}

fn status_date_picker_label(date: NaiveDate) -> String {
    format!("{}年{:02}月", date.year(), date.month())
}

fn task_duration_days(task: &GanttTask) -> Option<f32> {
    let start = task.start?;
    if let Some(days) = parse_duration_days(task.duration_text.as_str()) {
        return Some(days as f32);
    }

    let finish = task.finish?;
    let days = (finish.date() - start.date()).num_days().max(0) as f32;
    Some(days)
}

fn task_progress_x(
    task: &GanttTask,
    geom: &TaskBarGeometry,
    range: ChartRange,
    chart_left: f32,
    chart_right: f32,
    px_per_day: f32,
    full_complete_date: NaiveDate,
) -> Option<f32> {
    let progress_x = if task.percent_complete >= 100.0 {
        date_x(full_complete_date, range, chart_left, px_per_day)
    } else {
        let duration_days = task_duration_days(task)?;
        let progress = (task.percent_complete / 100.0).clamp(0.0, 1.0);
        geom.start_x + duration_days * progress * px_per_day
    };

    Some(progress_x.clamp(chart_left, chart_right))
}

fn collect_inazuma_progress_points(
    row_geometries: &[TaskBarGeometry],
    visible_indices: &[usize],
    tasks: &[GanttTask],
    range: ChartRange,
    chart_left: f32,
    chart_right: f32,
    px_per_day: f32,
    full_complete_date: NaiveDate,
) -> Vec<Pos2> {
    let mut line_points = Vec::with_capacity(row_geometries.len());

    for (geom, &task_index) in row_geometries.iter().zip(visible_indices.iter()) {
        let Some(task) = tasks.get(task_index) else {
            continue;
        };

        if task.summary || task.milestone || task.start.is_none() || task.finish.is_none() {
            continue;
        }

        let Some(progress_x) = task_progress_x(
            task,
            geom,
            range,
            chart_left,
            chart_right,
            px_per_day,
            full_complete_date,
        ) else {
            continue;
        };
        line_points.push(Pos2::new(progress_x, geom.center_y));
    }

    line_points
}

fn draw_progress_actual_line(
    painter: &egui::Painter,
    line_points: &[Pos2],
    status_x: f32,
    chart_body_rect: Rect,
    stroke: Stroke,
) {
    if line_points.is_empty() {
        return;
    }

    let mut path = Vec::with_capacity(line_points.len() + 2);
    path.push(Pos2::new(status_x, chart_body_rect.top()));
    path.extend_from_slice(line_points);
    path.push(Pos2::new(status_x, chart_body_rect.bottom()));

    let clipped_painter = painter.with_clip_rect(chart_body_rect);
    clipped_painter.add(egui::Shape::line(path, stroke));
}

fn paint_cell(ui: &mut egui::Ui, fill: Option<Color32>, selected: bool, active: bool) {
    let rect = ui.max_rect();
    if let Some(fill) = fill {
        paint_sheet_cell(ui.painter(), rect, fill, selected, active);
    }
}

fn draw_timescale_header(ui: &mut egui::Ui, rect: Rect, range: ChartRange, px_per_day: f32) {
    let painter = ui.painter();
    painter.rect_filled(rect, 0.0, HEADER_BG);
    painter.rect_stroke(
        rect,
        0.0,
        Stroke::new(1.0, Color32::from_rgb(118, 124, 132)),
        egui::StrokeKind::Inside,
    );

    let month_h = rect.height() * 0.30;
    let week_h = rect.height() * 0.24;
    let days = range.days();

    let month_rect = Rect::from_min_max(rect.min, Pos2::new(rect.max.x, rect.top() + month_h));
    let week_rect = Rect::from_min_max(
        Pos2::new(rect.left(), rect.top() + month_h),
        Pos2::new(rect.right(), rect.top() + month_h + week_h),
    );
    let day_rect = Rect::from_min_max(
        Pos2::new(rect.left(), rect.top() + month_h + week_h),
        rect.max,
    );
    painter.rect_filled(month_rect, 0.0, Color32::from_rgb(214, 221, 230));
    painter.rect_filled(week_rect, 0.0, Color32::from_rgb(224, 231, 239));
    painter.rect_filled(day_rect, 0.0, Color32::from_rgb(247, 249, 252));

    let mut month_start = 0;
    let mut week_start = 0;
    let mut current_month = range.start.month();
    let mut current_week = range.start.iso_week().week();
    let month_names = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];

    for day in 0..=days {
        let date = range.start + Duration::days(day);
        let x = rect.left() + day as f32 * px_per_day;

        let month_boundary = day == 0 || date.month() != current_month || day == days;
        if month_boundary && day > 0 {
            let label = format!(
                "{} {}",
                month_names[(current_month - 1) as usize],
                date.year()
            );
            let start_x = rect.left() + month_start as f32 * px_per_day;
            let month_rect = Rect::from_min_max(
                Pos2::new(start_x + 4.0, rect.top() + 2.0),
                Pos2::new((x - 4.0).max(start_x + 4.0), rect.top() + month_h - 2.0),
            );
            painter.text(
                month_rect.center(),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(10.5, FontFamily::Proportional),
                Color32::from_rgb(39, 51, 66),
            );
            painter.line_segment(
                [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
                Stroke::new(1.4, Color32::from_rgb(135, 142, 151)),
            );
            current_month = date.month();
            month_start = day;
        }

        let week_boundary = day == 0 || date.weekday() == chrono::Weekday::Mon || day == days;
        if week_boundary && day > 0 {
            let label = format!("W{:02}", current_week);
            let start_x = rect.left() + week_start as f32 * px_per_day;
            let week_rect = Rect::from_min_max(
                Pos2::new(start_x + 4.0, rect.top() + month_h + 2.0),
                Pos2::new(
                    (x - 4.0).max(start_x + 4.0),
                    rect.top() + month_h + week_h - 2.0,
                ),
            );
            painter.text(
                week_rect.center(),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(10.0, FontFamily::Proportional),
                Color32::from_rgb(54, 64, 78),
            );
            week_start = day;
            current_week = date.iso_week().week();
        }

        if day < days {
            let label = date.format("%d").to_string();
            let day_rect = Rect::from_min_max(
                Pos2::new(x, rect.top() + month_h + week_h),
                Pos2::new(x + px_per_day, rect.bottom()),
            );
            if date.weekday().number_from_monday() >= 6 {
                painter.rect_filled(
                    day_rect,
                    0.0,
                    Color32::from_rgba_unmultiplied(92, 116, 154, 10),
                );
            }
            painter.text(
                Pos2::new(day_rect.center().x, day_rect.center().y - 3.0),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(10.0, FontFamily::Proportional),
                Color32::from_rgb(34, 44, 58),
            );

            let weekday = match date.weekday().number_from_monday() {
                1 => "M",
                2 => "T",
                3 => "W",
                4 => "T",
                5 => "F",
                6 => "S",
                _ => "S",
            };
            painter.text(
                Pos2::new(day_rect.center().x, day_rect.bottom() - 2.0),
                egui::Align2::CENTER_BOTTOM,
                weekday,
                FontId::new(8.0, FontFamily::Proportional),
                if date.weekday().number_from_monday() >= 6 {
                    Color32::from_rgb(105, 83, 28)
                } else {
                    Color32::from_rgb(82, 92, 106)
                },
            );
        }
    }
}

fn draw_dependencies(
    painter: &egui::Painter,
    row_geometries: &[TaskBarGeometry],
    task_by_uid: &HashMap<u32, usize>,
    dependencies: &[GanttDependency],
    selected_uid: Option<u32>,
    active_uid: Option<u32>,
) {
    for dependency in dependencies {
        let Some(&pred_index) = task_by_uid.get(&dependency.predecessor_uid) else {
            continue;
        };
        let Some(&succ_index) = task_by_uid.get(&dependency.successor_uid) else {
            continue;
        };
        let Some(predecessor) = row_geometries.get(pred_index) else {
            continue;
        };
        let Some(successor) = row_geometries.get(succ_index) else {
            continue;
        };
        let highlighted = selected_uid == Some(dependency.predecessor_uid)
            || selected_uid == Some(dependency.successor_uid)
            || active_uid == Some(dependency.predecessor_uid)
            || active_uid == Some(dependency.successor_uid);

        let start = Pos2::new(predecessor.end_x, predecessor.center_y);
        let end = Pos2::new(successor.start_x, successor.center_y);
        let mid_x = (start.x + end.x) / 2.0;
        let path = vec![
            start,
            Pos2::new(mid_x, start.y),
            Pos2::new(mid_x, end.y),
            Pos2::new(end.x - 5.0, end.y),
        ];
        let shadow_color = if highlighted {
            Color32::from_rgba_unmultiplied(255, 255, 255, 220)
        } else {
            Color32::from_rgba_unmultiplied(248, 249, 251, 160)
        };
        let stroke = Stroke::new(
            if highlighted { 1.6 } else { 1.1 },
            if highlighted {
                ACCENT
            } else {
                Color32::from_rgb(90, 98, 110)
            },
        );
        for segment in path.windows(2) {
            painter.line_segment(
                [segment[0], segment[1]],
                Stroke::new(stroke.width + 2.0, shadow_color),
            );
        }
        for segment in path.windows(2) {
            painter.line_segment([segment[0], segment[1]], stroke);
        }
        painter.line_segment(
            [Pos2::new(end.x - 5.0, end.y - 4.0), end],
            Stroke::new(stroke.width + 1.5, shadow_color),
        );
        painter.line_segment(
            [Pos2::new(end.x - 5.0, end.y + 4.0), end],
            Stroke::new(stroke.width + 1.5, shadow_color),
        );
        painter.line_segment([Pos2::new(end.x - 5.0, end.y - 4.0), end], stroke);
        painter.line_segment([Pos2::new(end.x - 5.0, end.y + 4.0), end], stroke);
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
            if highlighted {
                ACCENT
            } else {
                Color32::from_rgb(80, 88, 98)
            },
        );
    }
}

fn find_task_by_uid<'a>(tasks: &'a [GanttTask], uid: u32) -> Option<&'a GanttTask> {
    tasks.iter().find(|task| task.uid == uid)
}

fn find_task_index_by_uid(tasks: &[GanttTask], uid: u32) -> Option<usize> {
    tasks.iter().position(|task| task.uid == uid)
}

fn apply_task_drag_preview(
    document: &mut ProjectDocument,
    original_document: &ProjectDocument,
    task_uid: u32,
    kind: TaskDragKind,
    delta_days: i64,
) -> bool {
    let Some(original_index) = find_task_index_by_uid(&original_document.tasks, task_uid) else {
        return false;
    };
    let Some(current_index) = find_task_index_by_uid(&document.tasks, task_uid) else {
        return false;
    };

    let original_task = &original_document.tasks[original_index];
    let Some(current_task) = document.tasks.get_mut(current_index) else {
        return false;
    };

    let Some(original_start) = original_task.start else {
        return false;
    };
    let Some(original_finish) = original_task.finish else {
        return false;
    };

    let (new_start, new_finish) = match kind {
        TaskDragKind::Move => (
            original_start + Duration::days(delta_days),
            original_finish + Duration::days(delta_days),
        ),
        TaskDragKind::ResizeStart => {
            let candidate = original_start + Duration::days(delta_days);
            (candidate.min(original_finish), original_finish)
        }
        TaskDragKind::ResizeFinish => {
            let candidate = original_finish + Duration::days(delta_days);
            (original_start, candidate.max(original_start))
        }
    };

    apply_task_date_change(current_task, new_start, new_finish)
}

fn apply_task_date_change(
    task: &mut GanttTask,
    start: chrono::NaiveDateTime,
    finish: chrono::NaiveDateTime,
) -> bool {
    let start_text = start.format("%Y-%m-%dT%H:%M:%S").to_string();
    let finish_text = finish.format("%Y-%m-%dT%H:%M:%S").to_string();
    let changed = task.start != Some(start)
        || task.finish != Some(finish)
        || task.start_text != start_text
        || task.finish_text != finish_text;

    if changed {
        task.start = Some(start);
        task.finish = Some(finish);
        task.start_text = start_text;
        task.finish_text = finish_text;
    }

    changed
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_task(uid: u32, id: u32, name: &str) -> GanttTask {
        GanttTask {
            uid,
            id,
            name: name.to_string(),
            outline_level: 1,
            summary: false,
            milestone: false,
            critical: false,
            percent_complete: 0.0,
            start_text: String::new(),
            finish_text: String::new(),
            start: None,
            finish: None,
            baseline_start: None,
            baseline_finish: None,
            duration_text: String::new(),
            predecessor_text: String::new(),
            notes_text: None,
            resource_names: None,
            calendar_uid: None,
            constraint_type: None,
        }
    }

    fn sample_status_date_text(date: NaiveDate) -> String {
        date.format("%Y-%m-%d").to_string()
    }

    fn sample_app(
        document: ProjectDocument,
        selected_task: usize,
        column: SheetColumn,
    ) -> GanttApp {
        let status_date = NaiveDate::from_ymd_opt(2026, 5, 9).unwrap();
        GanttApp {
            document: Some(document),
            document_path: None,
            recent_files: Vec::new(),
            status: String::new(),
            error: None,
            zoom_px_per_day: 18.0,
            status_date,
            status_date_text: sample_status_date_text(status_date),
            status_date_picker_open: false,
            status_date_picker_month: first_day_of_month(status_date),
            filter_text: String::new(),
            show_weekends: true,
            show_critical_path: true,
            dirty: false,
            selected_task: Some(selected_task),
            selected_sheet_column: column,
            sheet_selection_anchor: Some(SheetCellSelection {
                task_index: selected_task,
                column,
            }),
            sheet_editing_cell: None,
            sheet_selection_dragging: false,
            chart_interaction: None,
            dependency_picker: None,
            collapsed_task_uids: HashSet::new(),
            task_editor: TaskEditorState::default(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        }
    }

    #[test]
    fn copy_selection_serializes_tsv_rows() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(10, 1, "Task A"), sample_task(11, 2, "Task B")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 1, SheetColumn::Predecessors);
        if let Some(doc) = app.document.as_mut() {
            doc.tasks[0].duration_text = "2d".to_string();
            doc.tasks[0].start_text = "2026-05-01".to_string();
            doc.tasks[0].finish_text = "2026-05-03".to_string();
            doc.tasks[0].percent_complete = 25.0;
            doc.tasks[0].predecessor_text = "1FS".to_string();

            doc.tasks[1].duration_text = "3d".to_string();
            doc.tasks[1].start_text = "2026-05-04".to_string();
            doc.tasks[1].finish_text = "2026-05-06".to_string();
            doc.tasks[1].percent_complete = 50.0;
        }
        app.sheet_selection_anchor = Some(SheetCellSelection {
            task_index: 0,
            column: SheetColumn::Name,
        });

        let text = app
            .copy_sheet_selection_to_tsv(&[0, 1])
            .expect("selection should copy");

        assert_eq!(
            text,
            "Task A\t2d\t2026-05-01\t2026-05-03\t25\t1FS\r\nTask B\t3d\t2026-05-04\t2026-05-06\t50\t"
        );
    }

    #[test]
    fn copy_selection_serializes_rectangular_range_in_selection_order() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(10, 1, "Task A"), sample_task(11, 2, "Task B")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        if let Some(doc) = app.document.as_mut() {
            doc.tasks[0].duration_text = "2d".to_string();
            doc.tasks[0].start_text = "2026-05-01".to_string();
            doc.tasks[0].finish_text = "2026-05-03".to_string();
            doc.tasks[0].percent_complete = 25.0;
            doc.tasks[0].predecessor_text = "1FS".to_string();

            doc.tasks[1].duration_text = "3d".to_string();
            doc.tasks[1].start_text = "2026-05-04".to_string();
            doc.tasks[1].finish_text = "2026-05-06".to_string();
            doc.tasks[1].percent_complete = 50.0;
        }
        app.sheet_selection_anchor = Some(SheetCellSelection {
            task_index: 1,
            column: SheetColumn::Predecessors,
        });
        app.selected_task = Some(0);
        app.selected_sheet_column = SheetColumn::Name;

        let text = app
            .copy_sheet_selection_to_tsv(&[0, 1])
            .expect("selection should copy");

        assert_eq!(
            text,
            "Task A\t2d\t2026-05-01\t2026-05-03\t25\t1FS\r\nTask B\t3d\t2026-05-04\t2026-05-06\t50\t"
        );
    }

    #[test]
    fn paste_skips_mode_column_and_updates_multiple_rows() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![
                sample_task(100, 1, "Predecessor"),
                sample_task(101, 2, "Original A"),
                sample_task(102, 3, "Original B"),
            ],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 1, SheetColumn::Name);
        let visible_indices = vec![0, 1, 2];
        let pasted = "Auto Scheduled\tRenamed A\t5d\t2026-05-01\t2026-05-04\t50\t1FS\r\nAuto Scheduled\tRenamed B\t2d\t2026-05-05\t2026-05-06\t75\t";

        app.apply_excel_paste(pasted.to_string(), &visible_indices);

        let document = app.document.as_ref().expect("document should remain");
        assert_eq!(document.tasks[1].name, "Renamed A");
        assert_eq!(document.tasks[1].duration_text, "5日");
        assert_eq!(document.tasks[1].start_text, "2026-05-01");
        assert_eq!(document.tasks[1].finish_text, "2026-05-04");
        assert_eq!(document.tasks[1].percent_complete, 50.0);
        assert_eq!(document.tasks[1].predecessor_text, "1FS");
        assert_eq!(document.dependencies.len(), 1);
        assert_eq!(document.dependencies[0].predecessor_uid, 100);
        assert_eq!(document.dependencies[0].successor_uid, 101);

        assert_eq!(document.tasks[2].name, "Renamed B");
        assert_eq!(document.tasks[2].duration_text, "2日");
        assert_eq!(document.tasks[2].percent_complete, 75.0);
    }

    #[test]
    fn paste_handles_excel_bom_and_crlf_text() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![
                sample_task(100, 1, "Original A"),
                sample_task(101, 2, "Original B"),
            ],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        let visible_indices = vec![0, 1];
        let pasted = "\u{feff}Auto Scheduled\tRenamed A\t5d\t2026-05-01\t2026-05-04\t50\t1FS\r\nManually Scheduled\tRenamed B\t2d\t2026-05-05\t2026-05-06\t75\t";

        app.apply_excel_paste(pasted.to_string(), &visible_indices);

        let document = app.document.as_ref().expect("document should remain");
        assert_eq!(document.tasks[0].name, "Renamed A");
        assert_eq!(document.tasks[0].duration_text, "5日");
        assert_eq!(document.tasks[0].start_text, "2026-05-01");
        assert_eq!(document.tasks[0].finish_text, "2026-05-04");
        assert_eq!(document.tasks[0].percent_complete, 50.0);
        assert_eq!(document.tasks[0].predecessor_text, "1FS");
        assert_eq!(document.tasks[1].name, "Renamed B");
        assert_eq!(document.tasks[1].duration_text, "2日");
        assert_eq!(document.tasks[1].percent_complete, 75.0);
    }

    #[test]
    fn filler_row_selection_targets_last_visible_task() {
        let visible_indices = vec![4, 6, 9];

        assert_eq!(
            GanttApp::sheet_row_target_index(&visible_indices, 3),
            Some(9)
        );
        assert_eq!(
            GanttApp::sheet_row_target_index(&visible_indices, 20),
            Some(9)
        );
    }

    #[test]
    fn undo_restores_previous_task_editor_changes() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(100, 1, "Original")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        app.sync_task_editor_with_selection();
        app.task_editor.name = "Renamed".to_string();

        app.apply_task_editor_to_selection();

        assert_eq!(
            app.document.as_ref().expect("document should remain").tasks[0].name,
            "Renamed"
        );
        assert!(app.dirty);

        app.undo_last_change();

        let document = app.document.as_ref().expect("document should remain");
        assert_eq!(document.tasks[0].name, "Original");
        assert!(!app.dirty);
        assert_eq!(app.task_editor.name, "Original");
    }

    #[test]
    fn redo_restores_undone_task_editor_changes() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(100, 1, "Original")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        app.sync_task_editor_with_selection();
        app.task_editor.name = "Renamed".to_string();

        app.apply_task_editor_to_selection();
        app.undo_last_change();
        app.redo_last_change();

        let document = app.document.as_ref().expect("document should remain");
        assert_eq!(document.tasks[0].name, "Renamed");
        assert!(app.dirty);
        assert_eq!(app.task_editor.name, "Renamed");
    }

    #[test]
    fn ctrl_arrow_moves_to_visible_edges_without_changing_column_context() {
        assert_eq!(
            GanttApp::sheet_navigation_target(
                2,
                SheetColumn::Name,
                5,
                SheetNavigationDirection::Up,
                true
            ),
            (0, SheetColumn::Name)
        );
        assert_eq!(
            GanttApp::sheet_navigation_target(
                2,
                SheetColumn::Name,
                5,
                SheetNavigationDirection::Down,
                true
            ),
            (4, SheetColumn::Name)
        );
        assert_eq!(
            GanttApp::sheet_navigation_target(
                2,
                SheetColumn::Duration,
                5,
                SheetNavigationDirection::Left,
                true
            ),
            (2, SheetColumn::Mode)
        );
        assert_eq!(
            GanttApp::sheet_navigation_target(
                2,
                SheetColumn::Duration,
                5,
                SheetNavigationDirection::Right,
                true
            ),
            (2, SheetColumn::Predecessors)
        );
    }

    #[test]
    fn navigation_uses_visible_row_positions_for_hidden_rows() {
        let visible_indices = vec![2, 7, 15];
        let (next_row, next_col) = GanttApp::sheet_navigation_target(
            1,
            SheetColumn::Name,
            visible_indices.len(),
            SheetNavigationDirection::Down,
            false,
        );

        assert_eq!(visible_indices[next_row], 15);
        assert_eq!(next_col, SheetColumn::Name);
    }

    #[test]
    fn shift_selection_keeps_anchor_and_expands_from_it() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(1, 1, "Task A"), sample_task(2, 2, "Task B")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        app.sheet_selection_anchor = Some(SheetCellSelection {
            task_index: 0,
            column: SheetColumn::Name,
        });

        app.set_sheet_selection(1, SheetColumn::Duration, true);

        assert_eq!(app.selected_task, Some(1));
        assert_eq!(app.selected_sheet_column, SheetColumn::Duration);
        assert_eq!(
            app.sheet_selection_anchor,
            Some(SheetCellSelection {
                task_index: 0,
                column: SheetColumn::Name,
            })
        );
    }

    #[test]
    fn shift_selection_uses_previous_cell_when_anchor_is_missing() {
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![sample_task(1, 1, "Task A"), sample_task(2, 2, "Task B")],
            dependencies: vec![],
        };
        let mut app = sample_app(document, 0, SheetColumn::Name);
        app.sheet_selection_anchor = None;

        app.set_sheet_selection(1, SheetColumn::Duration, true);

        assert_eq!(app.selected_task, Some(1));
        assert_eq!(app.selected_sheet_column, SheetColumn::Duration);
        assert_eq!(
            app.sheet_selection_anchor,
            Some(SheetCellSelection {
                task_index: 0,
                column: SheetColumn::Name,
            })
        );
    }

    #[test]
    fn status_date_x_tracks_chart_position() {
        let range = ChartRange {
            start: NaiveDate::from_ymd_opt(2026, 5, 1).unwrap(),
            end: NaiveDate::from_ymd_opt(2026, 5, 11).unwrap(),
        };
        let chart_left = 12.0;
        let px_per_day = 20.0;

        assert_eq!(
            status_date_x(range.start, range, chart_left, px_per_day),
            12.0
        );
        assert_eq!(
            status_date_x(
                NaiveDate::from_ymd_opt(2026, 5, 6).unwrap(),
                range,
                chart_left,
                px_per_day
            ),
            112.0
        );
        assert_eq!(
            status_date_x(range.end, range, chart_left, px_per_day),
            212.0
        );
    }

    #[test]
    fn parse_app_settings_text_reads_status_date() {
        let settings = parse_app_settings_text("status_date=2026-05-09\n");
        assert_eq!(
            settings.map(|settings| settings.status_date),
            Some(NaiveDate::from_ymd_opt(2026, 5, 9).unwrap())
        );
    }

    #[test]
    fn task_progress_x_uses_duration_and_falls_back_to_finish_span() {
        let mut task = sample_task(1, 1, "Task A");
        task.start = parse_date_time(Some("2026-05-01T08:00:00"));
        task.finish = parse_date_time(Some("2026-05-05T17:00:00"));
        task.duration_text = "4d".to_string();
        task.percent_complete = 50.0;
        let range = ChartRange {
            start: NaiveDate::from_ymd_opt(2026, 5, 1).unwrap(),
            end: NaiveDate::from_ymd_opt(2026, 5, 11).unwrap(),
        };
        let geom = TaskBarGeometry {
            task_uid: task.uid,
            start_x: 40.0,
            end_x: 120.0,
            center_y: 20.0,
            bar_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(80.0, 12.0)),
            body_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(80.0, 12.0)),
            start_handle_rect: None,
            finish_handle_rect: None,
            link_handle_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(80.0, 12.0)),
        };

        assert_eq!(
            task_progress_x(
                &task,
                &geom,
                range,
                40.0,
                140.0,
                10.0,
                NaiveDate::from_ymd_opt(2026, 5, 9).unwrap()
            ),
            Some(60.0)
        );

        task.duration_text.clear();
        assert_eq!(
            task_progress_x(
                &task,
                &geom,
                range,
                40.0,
                140.0,
                10.0,
                NaiveDate::from_ymd_opt(2026, 5, 9).unwrap()
            ),
            Some(60.0)
        );
    }

    #[test]
    fn task_progress_x_uses_status_date_for_full_complete_tasks_and_clamps_right_edge() {
        let mut task = sample_task(1, 1, "Task A");
        task.start = parse_date_time(Some("2026-05-01T08:00:00"));
        task.finish = parse_date_time(Some("2026-05-12T17:00:00"));
        task.duration_text = "12d".to_string();
        task.percent_complete = 100.0;
        let range = ChartRange {
            start: NaiveDate::from_ymd_opt(2026, 5, 1).unwrap(),
            end: NaiveDate::from_ymd_opt(2026, 5, 11).unwrap(),
        };
        let geom = TaskBarGeometry {
            task_uid: task.uid,
            start_x: 40.0,
            end_x: 160.0,
            center_y: 20.0,
            bar_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(120.0, 12.0)),
            body_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(120.0, 12.0)),
            start_handle_rect: None,
            finish_handle_rect: None,
            link_handle_rect: Rect::from_min_size(Pos2::new(40.0, 14.0), Vec2::new(120.0, 12.0)),
        };

        assert_eq!(
            task_progress_x(
                &task,
                &geom,
                range,
                40.0,
                95.0,
                10.0,
                NaiveDate::from_ymd_opt(2026, 5, 9).unwrap()
            ),
            Some(95.0)
        );
    }

    #[test]
    fn collect_inazuma_progress_points_keeps_visible_row_order_and_skips_unscheduled_tasks() {
        let mut first = sample_task(1, 1, "Task A");
        first.start = parse_date_time(Some("2026-05-01T08:00:00"));
        first.finish = parse_date_time(Some("2026-05-03T17:00:00"));
        first.duration_text = "4d".to_string();
        first.percent_complete = 25.0;

        let mut second = sample_task(2, 2, "Task B");
        second.start = parse_date_time(Some("2026-05-04T08:00:00"));
        second.finish = parse_date_time(Some("2026-05-06T17:00:00"));
        second.duration_text = "8d".to_string();
        second.percent_complete = 75.0;

        let third = sample_task(3, 3, "Task C");

        let geometries = vec![
            TaskBarGeometry {
                task_uid: first.uid,
                start_x: 10.0,
                end_x: 110.0,
                center_y: 20.0,
                bar_rect: Rect::from_min_size(Pos2::new(10.0, 14.0), Vec2::new(100.0, 12.0)),
                body_rect: Rect::from_min_size(Pos2::new(10.0, 14.0), Vec2::new(100.0, 12.0)),
                start_handle_rect: None,
                finish_handle_rect: None,
                link_handle_rect: Rect::from_min_size(Pos2::new(10.0, 14.0), Vec2::new(100.0, 12.0)),
            },
            TaskBarGeometry {
                task_uid: second.uid,
                start_x: 10.0,
                end_x: 110.0,
                center_y: 36.0,
                bar_rect: Rect::from_min_size(Pos2::new(10.0, 30.0), Vec2::new(100.0, 12.0)),
                body_rect: Rect::from_min_size(Pos2::new(10.0, 30.0), Vec2::new(100.0, 12.0)),
                start_handle_rect: None,
                finish_handle_rect: None,
                link_handle_rect: Rect::from_min_size(Pos2::new(10.0, 30.0), Vec2::new(100.0, 12.0)),
            },
            TaskBarGeometry {
                task_uid: third.uid,
                start_x: 10.0,
                end_x: 110.0,
                center_y: 52.0,
                bar_rect: Rect::from_min_size(Pos2::new(10.0, 46.0), Vec2::new(100.0, 12.0)),
                body_rect: Rect::from_min_size(Pos2::new(10.0, 46.0), Vec2::new(100.0, 12.0)),
                start_handle_rect: None,
                finish_handle_rect: None,
                link_handle_rect: Rect::from_min_size(Pos2::new(10.0, 46.0), Vec2::new(100.0, 12.0)),
            },
        ];
        let tasks = vec![first, second, third];
        let visible_indices = vec![0, 1, 2];

        let line_points = collect_inazuma_progress_points(
            &geometries,
            &visible_indices,
            &tasks,
            ChartRange {
                start: NaiveDate::from_ymd_opt(2026, 5, 1).unwrap(),
                end: NaiveDate::from_ymd_opt(2026, 5, 11).unwrap(),
            },
            10.0,
            110.0,
            10.0,
            NaiveDate::from_ymd_opt(2026, 5, 9).unwrap(),
        );

        assert_eq!(line_points.len(), 2);
        assert_eq!(line_points[0], Pos2::new(20.0, 20.0));
        assert_eq!(line_points[1], Pos2::new(70.0, 36.0));
    }
}

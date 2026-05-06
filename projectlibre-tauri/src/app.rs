use crate::mspdi::{parse_date_time, ChartRange, GanttDependency, GanttTask, ProjectDocument};
use chrono::{Datelike, Duration};
use eframe::egui::{
    self, Align, Color32, Event, FontData, FontDefinitions, FontFamily, FontId, Pos2, Rect,
    RichText, Sense, Stroke, Vec2,
};
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};

const BODY_ROW_HEIGHT: f32 = 21.0;
const HEADER_HEIGHT: f32 = 42.0;
const SHEET_ROW_HEADER_WIDTH: f32 = 42.0;
const SHEET_NAME_WIDTH: f32 = 180.0;
const SHEET_DURATION_WIDTH: f32 = 72.0;
const SHEET_DATE_WIDTH: f32 = 92.0;
const SHEET_PERCENT_WIDTH: f32 = 74.0;
const SHEET_PREDECESSOR_WIDTH: f32 = 122.0;
const SHEET_CELL_PADDING_X: f32 = 6.0;
const SHEET_CELL_PADDING_Y: f32 = 2.0;
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
const OFFICE_GREEN: Color32 = Color32::from_rgb(45, 131, 50);
const OFFICE_GREEN_DARK: Color32 = Color32::from_rgb(24, 93, 31);
const OFFICE_PURPLE: Color32 = Color32::from_rgb(150, 82, 195);
const OFFICE_PURPLE_BG: Color32 = Color32::from_rgb(245, 239, 251);
const RIBBON_CHROME_BG: Color32 = Color32::from_rgb(236, 233, 216);
const TAB_BG: Color32 = Color32::from_rgb(236, 233, 216);
const TAB_ACTIVE_BG: Color32 = Color32::from_rgb(250, 250, 245);
const TIMELINE_BORDER: Color32 = Color32::from_rgb(160, 160, 160);
const TIMELINE_FILL: Color32 = Color32::from_rgb(236, 233, 216);
const TEXT: Color32 = Color32::from_rgb(30, 30, 30);
const SUBTEXT: Color32 = Color32::from_rgb(78, 78, 78);
const ACCENT: Color32 = Color32::from_rgb(0, 84, 147);
const ACCENT_SOFT: Color32 = Color32::from_rgb(217, 228, 242);
const CRITICAL: Color32 = Color32::from_rgb(194, 60, 40);
const BASELINE: Color32 = Color32::from_rgb(148, 148, 148);
const WEEKEND_BG: Color32 = Color32::from_rgb(248, 248, 248);
const TODAY: Color32 = Color32::from_rgb(242, 171, 61);
const UNDO_STACK_LIMIT: usize = 64;

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
enum RibbonTab {
    File,
    Task,
    Resource,
    Report,
    Project,
    View,
    Developer,
    Format,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RibbonIcon {
    Open,
    Save,
    Recent,
    Apply,
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

pub struct GanttApp {
    document: Option<ProjectDocument>,
    document_path: Option<PathBuf>,
    recent_files: Vec<PathBuf>,
    status: String,
    error: Option<String>,
    zoom_px_per_day: f32,
    filter_text: String,
    show_weekends: bool,
    show_critical_path: bool,
    dirty: bool,
    selected_task: Option<usize>,
    selected_sheet_column: SheetColumn,
    sheet_selection_anchor: Option<SheetCellSelection>,
    sheet_editing_cell: Option<SheetCellSelection>,
    sheet_selection_dragging: bool,
    collapsed_task_uids: HashSet<u32>,
    task_editor: TaskEditorState,
    undo_stack: Vec<UndoState>,
    redo_stack: Vec<UndoState>,
    active_ribbon_tab: RibbonTab,
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
    pub fn new(cc: &eframe::CreationContext<'_>, initial_file: Option<PathBuf>) -> Self {
        configure_fonts(&cc.egui_ctx);

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
            filter_text: String::new(),
            show_weekends: true,
            show_critical_path: true,
            dirty: false,
            selected_task: None,
            selected_sheet_column: SheetColumn::Name,
            sheet_selection_anchor: None,
            sheet_editing_cell: None,
            sheet_selection_dragging: false,
            collapsed_task_uids: HashSet::new(),
            task_editor: TaskEditorState::default(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            active_ribbon_tab: RibbonTab::File,
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
            let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
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

                if Self::apply_excel_row(document, task_index, start_column, &row[source_offset..]) {
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

        for assignment in parsed {
            let Some(predecessor_uid) =
                find_task_by_id(&document.tasks, assignment.predecessor_id).map(|task| task.uid)
            else {
                continue;
            };
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

        if let Some(task) = document.tasks.get_mut(index) {
            if task.predecessor_text != trimmed {
                task.predecessor_text = trimmed.to_string();
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
        let editing_name = matches!(
            self.sheet_editing_cell,
            Some(SheetCellSelection {
                column: SheetColumn::Name,
                ..
            })
        );

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

        if editing_name && input.4 {
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
            RibbonIcon::Recent => {
                painter.circle_stroke(rect.center(), 7.0, Stroke::new(1.6, gray));
                painter.line_segment(
                    [Pos2::new(cx, cy), Pos2::new(cx, cy - 3.8)],
                    Stroke::new(1.6, gray),
                );
                painter.line_segment(
                    [Pos2::new(cx, cy), Pos2::new(cx + 2.8, cy + 1.8)],
                    Stroke::new(1.6, gray),
                );
            }
            RibbonIcon::Apply => {
                painter.circle_filled(rect.center(), 8.5, green);
                painter.line_segment(
                    [Pos2::new(x1 + 6.0, cy + 0.5), Pos2::new(cx - 0.5, y2 - 5.0)],
                    Stroke::new(2.0, Color32::WHITE),
                );
                painter.line_segment(
                    [Pos2::new(cx - 0.5, y2 - 5.0), Pos2::new(x2 - 5.0, y1 + 6.0)],
                    Stroke::new(2.0, Color32::WHITE),
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
            show_weekends,
            show_critical_path,
            &mut dirty,
            &mut self.undo_stack,
            collapsed_task_uids,
        );
        self.sheet_selection_dragging = sheet_selection_dragging;
        if pending_selection != previous_selection {
            selection_changed = true;
        }

        self.dirty = dirty;
        self.selected_task = pending_selection;
        if selection_changed && self.selected_task.is_some() {
            self.sync_task_editor_with_selection();
        }
    }

    fn landing_view(ui: &mut egui::Ui, recent_files: &[PathBuf]) {
        egui::Frame::default()
            .fill(Color32::from_rgb(250, 252, 255))
            .show(ui, |ui| {
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

    fn document_header(
        ui: &mut egui::Ui,
        document: &ProjectDocument,
        visible_indices: &[usize],
        chart_range: ChartRange,
    ) {
        let calendar_count = document.calendars.len();
        let base_calendar_count = document
            .calendars
            .iter()
            .filter(|calendar| calendar.base_calendar)
            .count();
        let task_count = document.tasks.len();
        let dep_count = document.dependencies.len();
        let visible_count = visible_indices.len();

        egui::Frame::default()
            .fill(Color32::from_rgb(250, 252, 255))
            .stroke(Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(3, 2))
            .show(ui, |ui| {
                ui.horizontal_wrapped(|ui| {
                    ui.label(
                        RichText::new(document.title.as_deref().unwrap_or(&document.name))
                            .font(FontId::new(14.0, FontFamily::Proportional))
                            .strong(),
                    );
                    ui.separator();
                    ui.label(format!("タスク: {task_count}"));
                    ui.label(format!("表示: {visible_count}"));
                    ui.label(format!("依存関係: {dep_count}"));
                    if calendar_count > 0 {
                        ui.label(format!(
                            "カレンダー: {calendar_count} ({base_calendar_count} ベース)"
                        ));
                    } else {
                        ui.label("カレンダー: 0");
                    }
                    ui.label(format!(
                        "期間: {} - {}",
                        chart_range.start.format("%Y-%m-%d"),
                        chart_range.end.format("%Y-%m-%d")
                    ));
                    if let Some(manager) = &document.manager {
                        ui.label(format!("管理者: {manager}"));
                    }
                });
            });
    }

    fn render_toolbar_strip(
        ui: &mut egui::Ui,
        document: &mut ProjectDocument,
        visible_indices: &[usize],
    ) {
        egui::Frame::default()
            .fill(Color32::from_rgb(250, 252, 255))
            .stroke(Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(3, 2))
            .show(ui, |ui| {
                ui.horizontal_wrapped(|ui| {
                    ui.label(RichText::new("タスク一覧").strong());
                    ui.separator();
                    ui.label(format!("行: {}", visible_indices.len()));
                    ui.separator();
                    ui.label(format!("合計: {}", document.tasks.len()));
                    ui.separator();
                    ui.label("要約タスクの矢印で折りたたみ / 展開");
                });
            });
    }

    fn task_mode_strip(&mut self, ui: &mut egui::Ui) {
        let (rect, _) =
            ui.allocate_exact_size(Vec2::new(ui.available_width(), 36.0), Sense::hover());
        let painter = ui.painter_at(rect);
        painter.rect_filled(rect, 0.0, Color32::WHITE);
        painter.rect_stroke(
            rect,
            0.0,
            Stroke::new(1.0, BORDER),
            egui::StrokeKind::Inside,
        );

        let chip_y = rect.center().y + 1.0;
        draw_mode_chip(
            &painter,
            Rect::from_min_size(
                Pos2::new(rect.left() + 10.0, chip_y - 10.0),
                Vec2::new(20.0, 20.0),
            ),
            Color32::from_rgb(198, 68, 68),
            "✕",
        );
        draw_mode_chip(
            &painter,
            Rect::from_min_size(
                Pos2::new(rect.left() + 34.0, chip_y - 10.0),
                Vec2::new(20.0, 20.0),
            ),
            Color32::from_rgb(45, 116, 177),
            "✓",
        );

        let pill_rect = Rect::from_min_size(
            Pos2::new(rect.left() + 60.0, chip_y - 12.0),
            Vec2::new(128.0, 24.0),
        );
        painter.rect_filled(pill_rect, 12.0, Color32::from_rgb(232, 240, 232));
        painter.rect_stroke(
            pill_rect,
            12.0,
            Stroke::new(1.0, Color32::from_rgb(175, 198, 175)),
            egui::StrokeKind::Inside,
        );
        painter.text(
            Pos2::new(pill_rect.left() + 10.0, pill_rect.center().y - 1.0),
            egui::Align2::LEFT_CENTER,
            "↻",
            FontId::new(13.0, FontFamily::Proportional),
            OFFICE_GREEN_DARK,
        );
        painter.text(
            Pos2::new(pill_rect.left() + 24.0, pill_rect.center().y - 1.0),
            egui::Align2::LEFT_CENTER,
            "Auto Scheduled",
            FontId::new(13.0, FontFamily::Proportional),
            Color32::from_rgb(40, 40, 40),
        );
        painter.text(
            Pos2::new(pill_rect.right() - 10.0, pill_rect.center().y - 1.0),
            egui::Align2::RIGHT_CENTER,
            "▾",
            FontId::new(12.0, FontFamily::Proportional),
            Color32::from_rgb(84, 84, 84),
        );
    }

    fn timeline_strip(ui: &mut egui::Ui, _document: &ProjectDocument, chart_range: ChartRange) {
        let start_label = chart_range.start.format("%Y年%m月%d日").to_string();
        let end_label = chart_range.end.format("%Y年%m月%d日").to_string();
        let height = 56.0;
        let (rect, _) =
            ui.allocate_exact_size(Vec2::new(ui.available_width(), height), Sense::hover());
        let painter = ui.painter_at(rect);
        painter.rect_filled(rect, 0.0, TIMELINE_FILL);
        painter.rect_stroke(
            rect,
            0.0,
            Stroke::new(1.0, TIMELINE_BORDER),
            egui::StrokeKind::Inside,
        );

        let inner = rect.shrink2(Vec2::new(10.0, 6.0));
        let label_rect = Rect::from_min_size(inner.min, Vec2::new(70.0, 18.0));
        painter.text(
            label_rect.left_center(),
            egui::Align2::LEFT_CENTER,
            "タイムライン",
            FontId::new(11.0, FontFamily::Proportional),
            TIMELINE_BORDER,
        );
        painter.line_segment(
            [
                Pos2::new(label_rect.right() + 8.0, inner.top() + 9.0),
                Pos2::new(label_rect.right() + 8.0, inner.bottom() - 4.0),
            ],
            Stroke::new(1.0, Color32::from_rgb(201, 214, 198)),
        );

        painter.text(
            Pos2::new(inner.left() + 94.0, inner.top() + 3.0),
            egui::Align2::LEFT_TOP,
            start_label,
            FontId::new(10.0, FontFamily::Proportional),
            Color32::from_rgb(99, 99, 99),
        );
        painter.text(
            Pos2::new(inner.right() - 2.0, inner.top() + 3.0),
            egui::Align2::RIGHT_TOP,
            end_label,
            FontId::new(10.0, FontFamily::Proportional),
            Color32::from_rgb(99, 99, 99),
        );

        let message_rect = Rect::from_min_max(
            Pos2::new(inner.left() + 82.0, inner.top() + 18.0),
            Pos2::new(inner.right() - 10.0, inner.bottom() - 6.0),
        );
        painter.rect_stroke(
            message_rect,
            0.0,
            Stroke::new(1.0, Color32::from_rgb(180, 193, 175)),
            egui::StrokeKind::Inside,
        );
        painter.rect_filled(message_rect, 0.0, Color32::from_rgb(255, 255, 255));
        painter.text(
            message_rect.center(),
            egui::Align2::CENTER_CENTER,
            "日付が設定されたタスクをタイムラインに追加します",
            FontId::new(14.0, FontFamily::Proportional),
            Color32::from_rgb(71, 71, 71),
        );
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
                                ui.label(RichText::new("先行タスク").strong());
                                if predecessors.is_empty() {
                                    ui.label(
                                        RichText::new("先行タスクはありません").color(SUBTEXT),
                                    );
                                } else {
                                    for (dep, pred) in predecessors {
                                        ui.label(format!(
                                            "{} {} (lag {:?})",
                                            pred.name, dep.relation, dep.lag_text
                                        ));
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
                                ui.label(RichText::new("後続タスク").strong());
                                if successors.is_empty() {
                                    ui.label(
                                        RichText::new("後続タスクはありません").color(SUBTEXT),
                                    );
                                } else {
                                    for (dep, succ) in successors {
                                        ui.label(format!(
                                            "{} {} (lag {:?})",
                                            task.name, dep.relation, dep.lag_text
                                        ));
                                        ui.label(format!("  -> {}", succ.name));
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
        show_weekends: bool,
        show_critical_path: bool,
        dirty: &mut bool,
        undo_stack: &mut Vec<UndoState>,
        collapsed_task_uids: &mut HashSet<u32>,
    ) -> Option<usize> {
        let range = document.chart_range();
        let fixed_columns_width = 42.0 + 180.0 + 72.0 + 92.0 + 92.0 + 74.0 + 122.0;
        let chart_content_width = (range.days() as f32 * zoom_px_per_day).max(360.0);
        let chart_width =
            chart_content_width.max((ui.available_width() - fixed_columns_width).max(420.0));
        let filler_rows = ((ui.available_height() / BODY_ROW_HEIGHT).ceil() as usize).max(12);
        let total_rows = visible_indices.len().max(filler_rows);
        let visible_task_indices: Vec<usize> = visible_indices.to_vec();
        let task_by_uid = visible_task_indices
            .iter()
            .enumerate()
            .map(|(row_index, &task_index)| (document.tasks[task_index].uid, row_index))
            .collect::<HashMap<_, _>>();
        let geometry = RefCell::new(Vec::with_capacity(visible_task_indices.len()));
        let selected_uid =
            selected_task.and_then(|index| document.tasks.get(index).map(|task| task.uid));
        let editing_cell_value = *editing_cell;
        let selection_bounds = Self::sheet_selection_bounds_for_render(
            visible_indices,
            selected_task,
            *selected_sheet_column,
            *selection_anchor,
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
        let table_width: f32 = widths.iter().copied().sum();
        let table_height = HEADER_HEIGHT + total_rows as f32 * BODY_ROW_HEIGHT;
        let (table_rect, _) =
            ui.allocate_exact_size(Vec2::new(table_width, table_height), Sense::hover());
        let painter = ui.painter_at(table_rect);
        let table_left = table_rect.left();
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
                        egui::TextEdit::singleline(&mut name)
                            .frame(egui::Frame::NONE)
                            .lock_focus(true),
                    );
                    response.request_focus();
                    if Self::handle_task_name_outline_shortcut(
                        ui,
                        document,
                        index,
                        name_active,
                        selection_bounds,
                        visible_indices,
                        editing_cell,
                        undo_stack,
                        selected_task,
                        *selected_sheet_column,
                        *selection_anchor,
                        dirty,
                    ) {
                        ui.ctx().request_repaint();
                    }
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
                );
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
            }
        }

        let rects = geometry.into_inner();
        if !rects.is_empty() {
            draw_dependencies(
                ui.painter(),
                &rects,
                &task_by_uid,
                &document.dependencies,
                selected_uid,
            );
            let visible_tasks: Vec<&GanttTask> = visible_task_indices
                .iter()
                .map(|&index| &document.tasks[index])
                .collect();
            draw_progress_line(ui.painter(), &rects, &visible_tasks, selected_uid);
        }

        if pending_selection != selected_task {
            *dirty = true;
        }

        return pending_selection;

        #[cfg(any())]
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
                            paint_header_cell(ui, "", HeaderGlyph::Mode);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "Task Name", HeaderGlyph::Name);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "Duration", HeaderGlyph::Duration);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "Start", HeaderGlyph::Date);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "Finish", HeaderGlyph::Date);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "% Complete", HeaderGlyph::Percent);
                        });
                        header.col(|ui| {
                            paint_header_cell(ui, "Predecessors", HeaderGlyph::Predecessors);
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
                                                    .desired_width(ui.available_width().max(120.0))
                                                    .lock_focus(true),
                                            )
                                        } else {
                                            ui.add(
                                                egui::Label::new(RichText::new(task.name.as_str()))
                                                    .sense(Sense::click_and_drag()),
                                            )
                                        };
                                        if active {
                                            response.request_focus();
                                            if Self::handle_task_name_outline_shortcut(
                                                ui,
                                                document,
                                                index,
                                                active,
                                                selection_bounds,
                                                visible_indices,
                                                editing_cell,
                                                undo_stack,
                                                selected_task,
                                                *selected_sheet_column,
                                                *selection_anchor,
                                                dirty,
                                            ) {
                                                ui.ctx().request_repaint();
                                            }
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
                                    );
                                    geometry.borrow_mut().push(geom);
                                });
                            }
                        });
                    });

                #[cfg(any())]
                let rects = geometry.into_inner();
                #[cfg(any())]
                if !rects.is_empty() {
                    draw_dependencies(
                        ui.painter(),
                        &rects,
                        &task_by_uid,
                        &document.dependencies,
                        selected_uid,
                    );
                }
            });

        if pending_selection != selected_task {
            *dirty = true;
        }

        pending_selection
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
struct RowGeometry {
    start_x: f32,
    end_x: f32,
    center_y: f32,
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

    fn handle_task_name_outline_shortcut(
        ui: &egui::Ui,
        document: &mut ProjectDocument,
        index: usize,
        has_focus: bool,
        selection_bounds: Option<SheetSelectionBounds>,
        visible_indices: &[usize],
        editing_cell: &mut Option<SheetCellSelection>,
        undo_stack: &mut Vec<UndoState>,
        selected_task: Option<usize>,
        selected_sheet_column: SheetColumn,
        selection_anchor: Option<SheetCellSelection>,
        dirty: &mut bool,
    ) -> bool {
        if !has_focus {
            return false;
        }

        let tab_pressed = ui.input(|input| input.key_pressed(egui::Key::Tab));
        if !tab_pressed {
            return false;
        }

        let shift = ui.input(|input| input.modifiers.shift);
        let delta = if shift { -1 } else { 1 };
        let selected_task_indices = selection_bounds
            .filter(|bounds| bounds.start_row != bounds.end_row)
            .map(|bounds| Self::selected_task_indices_from_bounds(bounds, visible_indices))
            .filter(|indices| !indices.is_empty() && indices.len() > 1);

        if let Some(task_indices) = selected_task_indices {
            if Self::outline_delta_would_change_tasks(document, &task_indices, delta) {
                push_undo_state(
                    undo_stack,
                    document,
                    selected_task,
                    selected_sheet_column,
                    selection_anchor,
                    *dirty,
                );
                Self::apply_outline_delta_to_tasks(document, &task_indices, delta);
                *dirty = true;
            }
        } else {
            let Some(target_level) = Self::task_outline_level_target(document, index, delta) else {
                *editing_cell = Some(SheetCellSelection {
                    task_index: index,
                    column: SheetColumn::Name,
                });
                return true;
            };

            let Some(task) = document.tasks.get(index) else {
                return true;
            };
            let current = task.outline_level.max(1);
            if target_level != current {
                push_undo_state(
                    undo_stack,
                    document,
                    selected_task,
                    selected_sheet_column,
                    selection_anchor,
                    *dirty,
                );
                if let Some(task) = document.tasks.get_mut(index) {
                    task.outline_level = target_level;
                    *dirty = true;
                }
            }
        }

        *editing_cell = Some(SheetCellSelection {
            task_index: index,
            column: SheetColumn::Name,
        });
        true
    }

    fn task_outline_level_target(
        document: &ProjectDocument,
        index: usize,
        delta: i32,
    ) -> Option<u32> {
        let Some(task) = document.tasks.get(index) else {
            return None;
        };

        let current = task.outline_level.max(1);
        let target = if delta > 0 {
            let previous_limit = if index == 0 {
                1
            } else {
                document.tasks[index - 1].outline_level.max(1) + 1
            };
            (current + 1).min(previous_limit)
        } else {
            current.saturating_sub(1).max(1)
        };

        Some(target)
    }

    fn selected_task_indices_from_bounds(
        bounds: SheetSelectionBounds,
        visible_indices: &[usize],
    ) -> Vec<usize> {
        (bounds.start_row..=bounds.end_row)
            .filter_map(|row_index| visible_indices.get(row_index).copied())
            .collect()
    }

    fn apply_outline_delta_to_tasks(
        document: &mut ProjectDocument,
        task_indices: &[usize],
        delta: i32,
    ) -> bool {
        if task_indices.is_empty() {
            return false;
        }

        let mut changed = false;
        for &index in task_indices {
            let Some(task) = document.tasks.get_mut(index) else {
                continue;
            };
            let current = task.outline_level.max(1);
            let target = if delta > 0 {
                current.saturating_add(delta as u32)
            } else {
                current.saturating_sub(delta.unsigned_abs()).max(1)
            };
            if target != current {
                task.outline_level = target;
                changed = true;
            }
        }

        changed
    }

    fn outline_delta_would_change_tasks(
        document: &ProjectDocument,
        task_indices: &[usize],
        delta: i32,
    ) -> bool {
        task_indices.iter().copied().any(|index| {
            document.tasks.get(index).is_some_and(|task| {
                let current = task.outline_level.max(1);
                let target = if delta > 0 {
                    current.saturating_add(delta as u32)
                } else {
                    current.saturating_sub(delta.unsigned_abs()).max(1)
                };
                target != current
            })
        })
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

fn paint_cell(ui: &mut egui::Ui, fill: Option<Color32>, selected: bool, active: bool) {
    let rect = ui.max_rect();
    paint_sheet_cell(
        &ui.painter_at(rect),
        rect,
        fill.unwrap_or(SHEET_BG),
        selected,
        active,
    );
}

fn paint_header_cell(ui: &mut egui::Ui, title: &str, glyph: HeaderGlyph) {
    let rect = ui.max_rect();
    paint_sheet_header_cell(&ui.painter_at(rect), rect, title, glyph);
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

fn draw_mode_chip(painter: &egui::Painter, rect: Rect, color: Color32, text: &str) {
    painter.rect_filled(rect, 4.0, Color32::from_rgb(251, 252, 253));
    painter.rect_stroke(rect, 4.0, Stroke::new(1.0, color), egui::StrokeKind::Inside);
    painter.text(
        rect.center(),
        egui::Align2::CENTER_CENTER,
        text,
        FontId::new(13.0, FontFamily::Proportional),
        color,
    );
}

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
        painter.line_segment(
            [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
            Stroke::new(1.5, TODAY),
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
                painter.rect_filled(weekend_rect, 0.0, Color32::from_rgb(244, 242, 236));
            }
        }
        let line_color = if day % 7 == 0 {
            Color32::from_rgb(198, 198, 198)
        } else {
            Color32::from_rgb(228, 228, 228)
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
    selected: bool,
    show_critical_path: bool,
) -> RowGeometry {
    let row_center_y = rect.center().y;
    let start_day = task.start.map(|dt| dt.date()).unwrap_or(range.start);
    let finish_day = task.finish.map(|dt| dt.date()).unwrap_or(start_day);

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

    if let (Some(baseline_start), Some(baseline_finish)) =
        (task.baseline_start, task.baseline_finish)
    {
        let baseline_start_day = baseline_start.date();
        let baseline_finish_day = baseline_finish.date();
        let baseline_start_offset = (baseline_start_day - range.start).num_days().max(0) as f32;
        let baseline_finish_offset =
            (baseline_finish_day - range.start).num_days().max(0) as f32 + 1.0;
        let baseline = Rect::from_min_max(
            Pos2::new(
                rect.left() + baseline_start_offset * px_per_day,
                row_center_y - 7.0,
            ),
            Pos2::new(
                rect.left() + baseline_finish_offset * px_per_day,
                row_center_y - 5.0,
            ),
        );
        painter.rect_filled(baseline, 0.0, BASELINE);
    }

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
            base_color,
            Stroke::new(1.1, if selected { ACCENT } else { base_color }),
        ));
    } else if task.summary {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - 4.0),
            Pos2::new(end_x, row_center_y + 4.0),
        );
        painter.rect_filled(bar_rect, 1.0, base_color);
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
    } else {
        let bar_rect = Rect::from_min_max(
            Pos2::new(start_x, row_center_y - 6.0),
            Pos2::new(end_x, row_center_y + 6.0),
        );
        painter.rect_filled(bar_rect, 1.0, Color32::from_rgb(227, 232, 242));
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
            base_color,
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
    };

    RowGeometry {
        start_x,
        end_x,
        center_y: row_center_y,
    }
}

fn progress_point_on_rect(rect: Rect, percent_complete: f32) -> Pos2 {
    let progress = (percent_complete / 100.0).clamp(0.0, 1.0);
    Pos2::new(
        rect.left() + rect.width() * progress,
        rect.center().y,
    )
}

fn progress_line_point(task: &GanttTask, row: &RowGeometry) -> Option<Pos2> {
    if task.start.is_none() || task.finish.is_none() {
        return None;
    }

    if task.milestone {
        return Some(Pos2::new(row.start_x, row.center_y));
    }

    let width = (row.end_x - row.start_x).max(1.0);
    let rect = Rect::from_min_max(
        Pos2::new(row.start_x, row.center_y - 6.0),
        Pos2::new(row.start_x + width, row.center_y + 6.0),
    );
    Some(progress_point_on_rect(rect, task.percent_complete))
}

fn draw_progress_line(
    painter: &egui::Painter,
    row_geometries: &[RowGeometry],
    visible_tasks: &[&GanttTask],
    selected_uid: Option<u32>,
) {
    let mut points = Vec::with_capacity(row_geometries.len());
    let mut selected_point = None;

    for (task, row) in visible_tasks.iter().zip(row_geometries.iter()) {
        let Some(point) = progress_line_point(task, row) else {
            continue;
        };
        if selected_uid == Some(task.uid) {
            selected_point = Some(point);
        }
        points.push(point);
    }

    if points.len() < 2 {
        return;
    }

    let halo = Stroke::new(3.0, Color32::from_rgba_unmultiplied(255, 255, 255, 180));
    let stroke = Stroke::new(1.7, Color32::from_rgb(239, 0, 0));
    painter.add(egui::Shape::line(points.clone(), halo));
    painter.add(egui::Shape::line(points, stroke));

    if let Some(point) = selected_point {
        painter.circle_filled(point, 3.4, ACCENT);
        painter.circle_stroke(point, 3.4, Stroke::new(0.9, Color32::WHITE));
    }
}

fn draw_timescale_header(ui: &mut egui::Ui, rect: Rect, range: ChartRange, px_per_day: f32) {
    let painter = ui.painter();
    painter.rect_filled(rect, 0.0, HEADER_BG);
    painter.rect_stroke(
        rect,
        0.0,
        Stroke::new(1.0, Color32::from_rgb(128, 128, 128)),
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
    painter.rect_filled(month_rect, 0.0, Color32::from_rgb(223, 220, 213));
    painter.rect_filled(week_rect, 0.0, Color32::from_rgb(231, 228, 220));
    painter.rect_filled(day_rect, 0.0, Color32::from_rgb(248, 247, 243));

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
                Color32::from_rgb(53, 53, 53),
            );
            painter.line_segment(
                [Pos2::new(x, rect.top()), Pos2::new(x, rect.bottom())],
                Stroke::new(1.0, Color32::from_rgb(145, 145, 145)),
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
                Color32::from_rgb(64, 64, 64),
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
            painter.text(
                Pos2::new(day_rect.center().x, day_rect.center().y - 3.0),
                egui::Align2::CENTER_CENTER,
                label,
                FontId::new(10.0, FontFamily::Proportional),
                Color32::from_rgb(48, 48, 48),
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
                Color32::from_rgb(88, 88, 88),
            );
        }
    }
}

fn draw_dependencies(
    painter: &egui::Painter,
    row_geometries: &[RowGeometry],
    task_by_uid: &HashMap<u32, usize>,
    dependencies: &[GanttDependency],
    selected_uid: Option<u32>,
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
        let highlighted = selected_uid == Some(dependency.predecessor_uid)
            || selected_uid == Some(dependency.successor_uid);

        let start = Pos2::new(predecessor.end_x, predecessor.center_y);
        let end = Pos2::new(successor.start_x, successor.center_y);
        let mid_x = (start.x + end.x) / 2.0;
        let path = vec![
            start,
            Pos2::new(mid_x, start.y),
            Pos2::new(mid_x, end.y),
            Pos2::new(end.x - 5.0, end.y),
        ];
        let stroke = Stroke::new(
            if highlighted { 1.4 } else { 1.0 },
            if highlighted {
                ACCENT
            } else {
                Color32::from_rgb(96, 96, 96)
            },
        );
        for segment in path.windows(2) {
            painter.line_segment([segment[0], segment[1]], stroke);
        }
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
                Color32::from_rgb(86, 86, 86)
            },
        );
    }
}

fn find_task_by_uid<'a>(tasks: &'a [GanttTask], uid: u32) -> Option<&'a GanttTask> {
    tasks.iter().find(|task| task.uid == uid)
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

    fn sample_app(
        document: ProjectDocument,
        selected_task: usize,
        column: SheetColumn,
    ) -> GanttApp {
        GanttApp {
            document: Some(document),
            document_path: None,
            recent_files: Vec::new(),
            status: String::new(),
            error: None,
            zoom_px_per_day: 18.0,
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
            collapsed_task_uids: HashSet::new(),
            task_editor: TaskEditorState::default(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            active_ribbon_tab: RibbonTab::File,
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
    fn task_outline_level_target_respects_previous_row_limit() {
        let mut first = sample_task(1, 1, "Task A");
        first.outline_level = 1;
        let mut second = sample_task(2, 2, "Task B");
        second.outline_level = 3;
        let document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![first, second],
            dependencies: vec![],
        };

        assert_eq!(
            GanttApp::task_outline_level_target(&document, 1, 1),
            Some(2)
        );
        assert_eq!(
            GanttApp::task_outline_level_target(&document, 1, -1),
            Some(2)
        );
        assert_eq!(
            GanttApp::task_outline_level_target(&document, 0, -1),
            Some(1)
        );
    }

    #[test]
    fn selected_task_indices_from_bounds_collects_visible_rows() {
        let visible_indices = vec![10, 11, 12, 13];
        let bounds = SheetSelectionBounds {
            start_row: 1,
            end_row: 3,
            start_column: SheetColumn::Name,
            end_column: SheetColumn::Name,
        };

        assert_eq!(
            GanttApp::selected_task_indices_from_bounds(bounds, &visible_indices),
            vec![11, 12, 13]
        );
    }

    #[test]
    fn apply_outline_delta_to_tasks_updates_multiple_rows_uniformly() {
        let mut first = sample_task(1, 1, "Task A");
        first.outline_level = 1;
        let mut second = sample_task(2, 2, "Task B");
        second.outline_level = 3;
        let mut third = sample_task(3, 3, "Task C");
        third.outline_level = 2;
        let mut document = ProjectDocument {
            name: "Demo".to_string(),
            title: None,
            manager: None,
            start_date: None,
            finish_date: None,
            calendars: vec![],
            tasks: vec![first, second, third],
            dependencies: vec![],
        };

        assert!(GanttApp::outline_delta_would_change_tasks(&document, &[0, 1, 2], 1));
        assert!(GanttApp::apply_outline_delta_to_tasks(
            &mut document,
            &[0, 1, 2],
            1
        ));
        assert_eq!(document.tasks[0].outline_level, 2);
        assert_eq!(document.tasks[1].outline_level, 4);
        assert_eq!(document.tasks[2].outline_level, 3);
    }

    #[test]
    fn progress_point_is_interpolated_across_bar_width() {
        let rect = Rect::from_min_size(Pos2::new(10.0, 20.0), Vec2::new(120.0, 14.0));

        let start = progress_point_on_rect(rect, 0.0);
        let middle = progress_point_on_rect(rect, 50.0);
        let end = progress_point_on_rect(rect, 100.0);

        assert_eq!(start, Pos2::new(10.0, 27.0));
        assert_eq!(middle, Pos2::new(70.0, 27.0));
        assert_eq!(end, Pos2::new(130.0, 27.0));
    }

    #[test]
    fn progress_line_point_uses_percent_complete_across_row_width() {
        let mut task = sample_task(1, 1, "Task A");
        task.start = parse_date_time(Some("2026-05-01T08:00:00"));
        task.finish = parse_date_time(Some("2026-05-03T17:00:00"));
        task.percent_complete = 50.0;

        let row = RowGeometry {
            start_x: 10.0,
            end_x: 130.0,
            center_y: 27.0,
        };

        assert_eq!(
            progress_line_point(&task, &row),
            Some(Pos2::new(70.0, 27.0))
        );
    }

    #[test]
    fn progress_line_point_uses_milestone_center() {
        let mut task = sample_task(1, 1, "Task A");
        task.start = parse_date_time(Some("2026-05-01T08:00:00"));
        task.finish = parse_date_time(Some("2026-05-01T08:00:00"));
        task.milestone = true;
        task.percent_complete = 80.0;

        let row = RowGeometry {
            start_x: 42.0,
            end_x: 42.0,
            center_y: 27.0,
        };

        assert_eq!(progress_line_point(&task, &row), Some(Pos2::new(42.0, 27.0)));
    }

    #[test]
    fn progress_line_point_skips_unscheduled_tasks() {
        let task = sample_task(1, 1, "Task A");
        let row = RowGeometry {
            start_x: 10.0,
            end_x: 130.0,
            center_y: 27.0,
        };

        assert!(progress_line_point(&task, &row).is_none());
    }
}

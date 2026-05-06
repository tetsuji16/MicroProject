const tauriApi = window.__TAURI__ || {};
const invoke = tauriApi.tauri?.invoke;
const openDialog = tauriApi.dialog?.open;
const saveDialog = tauriApi.dialog?.save;

const state = {
  snapshot: null,
  selectedUid: null,
  collapsed: new Set(),
  filter: "",
  viewMode: "Day",
  showCritical: true,
  showWeekends: true,
  gantt: null,
  resizeObserver: null,
};

const els = {};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  cacheElements();
  bindEvents();
  state.resizeObserver = new ResizeObserver(() => {
    if (state.snapshot) {
      renderGantt();
    }
  });
  state.resizeObserver.observe(els.ganttHost);
  await refreshSnapshot();
}

function cacheElements() {
  const ids = [
    "project-title",
    "sheet-body",
    "gantt",
    "range-label",
    "status-text",
    "selection-text",
    "dirty-state",
    "filter-input",
    "show-weekends",
    "show-critical",
    "open-button",
    "save-button",
    "save-as-button",
    "new-task-button",
    "delete-task-button",
    "today-button",
    "apply-task-button",
    "refresh-button",
    "task-form",
    "task-uid",
    "task-name",
    "task-outline",
    "task-progress",
    "task-summary",
    "task-milestone",
    "task-critical",
    "task-start",
    "task-finish",
    "task-baseline-start",
    "task-baseline-finish",
    "task-duration",
    "task-resources",
    "task-calendar",
    "task-constraint",
    "task-notes",
    "dependency-predecessor",
    "dependency-relation",
    "dependency-lag",
    "add-dependency-button",
    "dependency-list",
  ];

  for (const id of ids) {
    els[id.replace(/-([a-z])/g, (_, ch) => ch.toUpperCase())] = document.getElementById(id);
  }
  els.sheetBody = document.getElementById("sheet-body");
  els.ganttHost = document.getElementById("gantt");
}

function bindEvents() {
  els.filterInput.addEventListener("input", () => {
    state.filter = els.filterInput.value.trim().toLowerCase();
    renderAll();
  });

  els.openButton.addEventListener("click", openProjectDialog);
  els.saveButton.addEventListener("click", saveProject);
  els.saveAsButton.addEventListener("click", saveProjectAs);
  els.newTaskButton.addEventListener("click", createTask);
  els.deleteTaskButton.addEventListener("click", deleteSelectedTask);
  els.todayButton.addEventListener("click", () => state.gantt?.scroll_current?.());
  els.refreshButton.addEventListener("click", refreshSnapshot);
  els.applyTaskButton.addEventListener("click", applyTaskForm);
  els.addDependencyButton.addEventListener("click", addDependencyFromForm);
  els.showCritical.addEventListener("change", () => {
    state.showCritical = els.showCritical.checked;
    renderAll();
  });
  els.showWeekends.addEventListener("change", () => {
    state.showWeekends = els.showWeekends.checked;
    renderAll();
  });

  for (const button of document.querySelectorAll("[data-zoom]")) {
    button.addEventListener("click", () => {
      const mode = button.dataset.zoom;
      state.viewMode = mode;
      renderGantt();
    });
  }

  els.taskForm.addEventListener("submit", (event) => {
    event.preventDefault();
    applyTaskForm();
  });

  for (const input of [
    els.taskName,
    els.taskOutline,
    els.taskProgress,
    els.taskSummary,
    els.taskMilestone,
    els.taskCritical,
    els.taskStart,
    els.taskFinish,
    els.taskBaselineStart,
    els.taskBaselineFinish,
    els.taskDuration,
    els.taskResources,
    els.taskCalendar,
    els.taskConstraint,
    els.taskNotes,
  ]) {
    input.addEventListener("change", () => {
      els.dirtyState.textContent = "Unsaved";
    });
  }
}

async function refreshSnapshot() {
  if (!invoke) {
    setStatus("Tauri bridge is unavailable.");
    return;
  }

  try {
    const snapshot = await invoke("project_snapshot");
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Failed to load project: ${stringifyError(error)}`);
  }
}

function loadSnapshot(snapshot) {
  state.snapshot = snapshot;
  if (!state.selectedUid && snapshot.tasks.length) {
    state.selectedUid = snapshot.tasks[0].uid;
  }
  if (!snapshot.tasks.some((task) => task.uid === state.selectedUid)) {
    state.selectedUid = snapshot.tasks[0]?.uid ?? null;
  }
  els.projectTitle.textContent = snapshot.title || snapshot.name || "MicroProject";
  els.rangeLabel.textContent = `${snapshot.chart_range.start} - ${snapshot.chart_range.end}`;
  setStatus(
    snapshot.path ? `Loaded ${snapshot.path}` : "Loaded in-memory project",
    snapshot.dirty,
  );
  state.viewMode = state.viewMode || "Day";
  renderAll();
}

function renderAll() {
  if (!state.snapshot) {
    return;
  }
  const visibleTasks = getVisibleTasks();
  renderSheet(visibleTasks);
  renderGantt(visibleTasks);
  renderInspector();
  renderStatus();
}

function getVisibleTasks() {
  const snapshot = state.snapshot;
  const visible = [];
  const stack = [];
  for (const task of snapshot.tasks) {
    while (stack.length && stack[stack.length - 1].level >= task.outline_level) {
      stack.pop();
    }
    const blocked = stack.some((item) => item.collapsed);
    const matches = !state.filter || matchesFilter(task, state.filter);
    if (!blocked && matches) {
      visible.push(task);
    }
    stack.push({
      level: task.outline_level,
      collapsed: state.collapsed.has(task.uid),
    });
  }
  return visible;
}

function matchesFilter(task, filter) {
  const haystack = [
    task.uid,
    task.id,
    task.name,
    task.duration_text,
    task.start_text,
    task.finish_text,
    task.predecessor_text,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
  return haystack.includes(filter);
}

function renderSheet(tasks) {
  const selectedUid = state.selectedUid;
  els.sheetBody.innerHTML = "";
  const fragment = document.createDocumentFragment();

  tasks.forEach((task, index) => {
    const row = document.createElement("div");
    row.className = [
      "sheet-row",
      task.summary ? "is-summary" : "",
      state.showCritical && task.critical ? "is-critical" : "",
      task.uid === selectedUid ? "is-selected" : "",
    ]
      .filter(Boolean)
      .join(" ");

    row.addEventListener("click", () => {
      state.selectedUid = task.uid;
      renderAll();
    });

    const nameCell = document.createElement("div");
    nameCell.className = "sheet-cell sheet-name";
    const rowIndex = document.createElement("span");
    rowIndex.className = "sheet-row-index";
    rowIndex.textContent = String(index + 1);
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "sheet-toggle";
    toggle.textContent = task.summary ? (state.collapsed.has(task.uid) ? "+" : "−") : "";
    toggle.title = task.summary ? "Toggle summary" : "";
    toggle.disabled = !task.summary;
    toggle.addEventListener("click", (event) => {
      event.stopPropagation();
      if (state.collapsed.has(task.uid)) {
        state.collapsed.delete(task.uid);
      } else {
        state.collapsed.add(task.uid);
      }
      renderAll();
    });
    const indent = document.createElement("span");
    indent.className = "sheet-indent";
    indent.style.width = `${Math.max(0, task.outline_level - 1) * 12}px`;
    const mode = document.createElement("span");
    mode.className = [
      "sheet-mode",
      task.milestone ? "sheet-mode--milestone" : "",
      task.summary ? "sheet-mode--summary" : "",
    ]
      .filter(Boolean)
      .join(" ");
    mode.textContent = task.milestone ? "◆" : task.summary ? "▦" : "●";
    const label = document.createElement("span");
    label.className = "sheet-label";
    label.textContent = task.name;
    nameCell.append(rowIndex, toggle, indent, mode, label);

    const durationCell = cell(task.duration_text);
    const startCell = cell(task.start_text);
    const finishCell = cell(task.finish_text);
    const percentCell = cell(`${Math.round(task.percent_complete)}%`, "sheet-cell--percent");
    const predCell = cell(task.predecessor_text || "—", "sheet-cell--predecessors");

    row.append(nameCell, durationCell, startCell, finishCell, percentCell, predCell);
    fragment.append(row);
  });

    if (!tasks.length) {
      const empty = document.createElement("div");
      empty.className = "sheet-row";
      empty.textContent = "No tasks match the current filter.";
      fragment.append(empty);
  }

  els.sheetBody.append(fragment);
}

function renderGantt(tasks) {
  if (!tasks.length) {
    els.ganttHost.innerHTML = '<div class="gantt-empty">No visible tasks.</div>';
    state.gantt = null;
    return;
  }

  const visibleIds = new Set(tasks.map((task) => String(task.uid)));
  const dependencies = new Map();
  for (const dependency of state.snapshot.dependencies) {
    if (!visibleIds.has(String(dependency.predecessor_uid))) {
      continue;
    }
    const list = dependencies.get(dependency.successor_uid) || [];
    list.push(String(dependency.predecessor_uid));
    dependencies.set(dependency.successor_uid, list);
  }

  const ganttTasks = tasks.map((task) => {
    const start = dateOnly(task.start_text || state.snapshot.chart_range.start);
    const end = dateOnly(task.finish_text || task.start_text || state.snapshot.chart_range.end);
    return {
      id: String(task.uid),
      name: task.name,
      start,
      end: task.milestone ? start : end,
      progress: Math.round(task.percent_complete),
      dependencies: (dependencies.get(task.uid) || []).join(","),
      custom_class: [
        task.summary ? "task-summary" : "",
        task.milestone ? "task-milestone" : "",
        state.showCritical && task.critical ? "task-critical" : "",
      ]
        .filter(Boolean)
        .join(" "),
    };
  });

  els.ganttHost.innerHTML = "";
  const options = {
    view_mode: state.viewMode,
    view_mode_select: false,
    today_button: false,
    readonly: false,
    readonly_dates: false,
    readonly_progress: false,
    scroll_to: "today",
    show_expected_progress: true,
    popup: () => false,
    on_click: (task) => {
      state.selectedUid = Number(task.id);
      renderAll();
    },
    on_date_change: (task, start, end) => {
      updateTaskFromChart(task.id, {
        start_text: toDateTimeText(start, false),
        finish_text: toDateTimeText(end, true),
      });
    },
    on_progress_change: (task, progress) => {
      updateTaskFromChart(task.id, { percent_complete: progress });
    },
  };

  state.gantt = new Gantt("#gantt", ganttTasks, options);
  setTimeout(() => state.gantt?.scroll_current?.(), 0);
}

function renderInspector() {
  const task = currentTask();
  const tasks = state.snapshot?.tasks || [];
  if (!task) {
    clearForm();
    renderDependencies([]);
    els.selectionText.textContent = "No task selected";
    return;
  }

  els.taskUid.value = String(task.uid);
  els.taskName.value = task.name || "";
  els.taskOutline.value = String(task.outline_level || 1);
  els.taskProgress.value = String(Math.round(task.percent_complete || 0));
  els.taskSummary.checked = Boolean(task.summary);
  els.taskMilestone.checked = Boolean(task.milestone);
  els.taskCritical.checked = Boolean(task.critical);
  els.taskStart.value = task.start_text || "";
  els.taskFinish.value = task.finish_text || "";
  els.taskBaselineStart.value = task.baseline_start_text || "";
  els.taskBaselineFinish.value = task.baseline_finish_text || "";
  els.taskDuration.value = task.duration_text || "";
  els.taskResources.value = task.resource_names || "";
  els.taskCalendar.value = task.calendar_uid == null ? "" : String(task.calendar_uid);
  els.taskConstraint.value = task.constraint_type || "";
  els.taskNotes.value = task.notes_text || "";

  const incoming = state.snapshot.dependencies.filter(
    (dependency) => dependency.successor_uid === task.uid,
  );
  renderDependencies(incoming, tasks, task.uid);
  populateDependencyPredecessorSelect(task.uid);
  els.selectionText.textContent = `Selected UID ${task.uid}`;
}

function renderDependencies(incoming, tasks = state.snapshot.tasks, selectedUid = null) {
  const container = els.dependencyList;
  container.innerHTML = "";
  if (!incoming.length) {
    const empty = document.createElement("div");
    empty.className = "dependency-item";
    empty.textContent = "No predecessors.";
    container.append(empty);
    return;
  }

  for (const dependency of incoming) {
    const item = document.createElement("div");
    item.className = "dependency-item";
    const left = document.createElement("span");
    const pred = tasks.find((task) => task.uid === dependency.predecessor_uid);
    left.textContent = `${pred ? pred.name : dependency.predecessor_uid} -> ${dependency.relation} ${dependency.lag_text || ""}`.trim();
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      invoke("project_delete_dependency", {
        input: {
          predecessor_uid: dependency.predecessor_uid,
          successor_uid: selectedUid,
          relation: dependency.relation,
          lag_text: dependency.lag_text,
        },
      })
        .then(loadSnapshot)
        .catch((error) => setStatus(`Failed to remove dependency: ${stringifyError(error)}`));
    });
    item.append(left, remove);
    container.append(item);
  }
}

function populateDependencyPredecessorSelect(selectedUid) {
  const select = els.dependencyPredecessor;
  select.innerHTML = "";
  for (const task of state.snapshot.tasks) {
    if (task.uid === selectedUid) {
      continue;
    }
    const option = document.createElement("option");
    option.value = String(task.uid);
    option.textContent = `${task.uid}: ${task.name}`;
    select.append(option);
  }
}

async function openProjectDialog() {
  if (!openDialog) {
    return;
  }

  try {
    const selected = await openDialog({
      multiple: false,
      filters: [
        {
          name: "MS Project XML",
          extensions: ["xml", "mspdi"],
        },
      ],
    });

    if (!selected) {
      return;
    }
    const path = Array.isArray(selected) ? selected[0] : selected;
    const snapshot = await invoke("project_open", { path });
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Open failed: ${stringifyError(error)}`);
  }
}

async function saveProject() {
  try {
    const snapshot = await invoke("project_save");
    loadSnapshot(snapshot);
    setStatus("Saved", false);
  } catch (error) {
    setStatus(`Save failed: ${stringifyError(error)}`);
  }
}

async function saveProjectAs() {
  if (!saveDialog) {
    return;
  }

  try {
    const selected = await saveDialog({
      filters: [
        {
          name: "MS Project XML",
          extensions: ["xml", "mspdi"],
        },
      ],
    });

    if (!selected) {
      return;
    }
    const snapshot = await invoke("project_save_as", { path: selected });
    loadSnapshot(snapshot);
    setStatus("Saved as new file", false);
  } catch (error) {
    setStatus(`Save As failed: ${stringifyError(error)}`);
  }
}

async function createTask() {
  try {
    const snapshot = await invoke("project_create_task", { after_uid: state.selectedUid });
    loadSnapshot(snapshot);
    state.selectedUid = snapshot.tasks[snapshot.tasks.length - 1]?.uid ?? state.selectedUid;
    renderAll();
  } catch (error) {
    setStatus(`Create task failed: ${stringifyError(error)}`);
  }
}

async function deleteSelectedTask() {
  const task = currentTask();
  if (!task) {
    return;
  }
  try {
    const snapshot = await invoke("project_delete_task", { uid: task.uid });
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Delete failed: ${stringifyError(error)}`);
  }
}

async function applyTaskForm() {
  const task = currentTask();
  if (!task) {
    return;
  }

  const input = {
    uid: task.uid,
    name: els.taskName.value.trim(),
    outline_level: clampInt(els.taskOutline.value, 1),
    summary: els.taskSummary.checked,
    milestone: els.taskMilestone.checked,
    critical: els.taskCritical.checked,
    percent_complete: clampNumber(els.taskProgress.value, 0, 100),
    start_text: els.taskStart.value.trim(),
    finish_text: els.taskFinish.value.trim(),
    baseline_start_text: blankToNull(els.taskBaselineStart.value),
    baseline_finish_text: blankToNull(els.taskBaselineFinish.value),
    duration_text: els.taskDuration.value.trim(),
    notes_text: blankToNull(els.taskNotes.value),
    resource_names: blankToNull(els.taskResources.value),
    calendar_uid: blankToNull(els.taskCalendar.value),
    constraint_type: blankToNull(els.taskConstraint.value),
  };

  try {
    const snapshot = await invoke("project_upsert_task", { input });
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Save task failed: ${stringifyError(error)}`);
  }
}

async function addDependencyFromForm() {
  const task = currentTask();
  if (!task || !els.dependencyPredecessor.value) {
    return;
  }

  try {
    const snapshot = await invoke("project_upsert_dependency", {
      input: {
        predecessor_uid: Number(els.dependencyPredecessor.value),
        successor_uid: task.uid,
        relation: els.dependencyRelation.value,
        lag_text: blankToNull(els.dependencyLag.value),
      },
    });
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Add dependency failed: ${stringifyError(error)}`);
  }
}

async function updateTaskFromChart(taskId, patch) {
  const task = state.snapshot.tasks.find((item) => String(item.uid) === String(taskId));
  if (!task) {
    return;
  }
  try {
    const snapshot = await invoke("project_upsert_task", {
      input: {
        uid: task.uid,
        name: task.name,
        outline_level: task.outline_level,
        summary: task.summary,
        milestone: task.milestone,
        critical: task.critical,
        percent_complete:
          patch.percent_complete == null ? task.percent_complete : patch.percent_complete,
        start_text: patch.start_text || task.start_text,
        finish_text: patch.finish_text || task.finish_text,
        baseline_start_text: task.baseline_start_text,
        baseline_finish_text: task.baseline_finish_text,
        duration_text: task.duration_text,
        notes_text: task.notes_text,
        resource_names: task.resource_names,
        calendar_uid: task.calendar_uid,
        constraint_type: task.constraint_type,
      },
    });
    loadSnapshot(snapshot);
  } catch (error) {
    setStatus(`Task update failed: ${stringifyError(error)}`);
  }
}

function currentTask() {
  return state.snapshot?.tasks.find((task) => task.uid === state.selectedUid) || null;
}

function clearForm() {
  for (const input of [
    els.taskUid,
    els.taskName,
    els.taskOutline,
    els.taskProgress,
    els.taskStart,
    els.taskFinish,
    els.taskBaselineStart,
    els.taskBaselineFinish,
    els.taskDuration,
    els.taskResources,
    els.taskCalendar,
    els.taskConstraint,
    els.taskNotes,
  ]) {
    input.value = "";
  }
  els.taskSummary.checked = false;
  els.taskMilestone.checked = false;
  els.taskCritical.checked = false;
  els.dependencyList.innerHTML = "";
  els.dependencyPredecessor.innerHTML = "";
  els.selectionText.textContent = "";
}

function renderStatus() {
  const dirty = state.snapshot?.dirty;
  els.dirtyState.textContent = dirty ? "Unsaved" : "Saved";
}

function setStatus(text, dirty = null) {
  els.statusText.textContent = text;
  if (dirty !== null) {
    els.dirtyState.textContent = dirty ? "Unsaved" : "Saved";
  }
}

function cell(text, className = "") {
  const element = document.createElement("div");
  element.className = `sheet-cell ${className}`.trim();
  element.textContent = text;
  return element;
}

function dateOnly(value) {
  if (!value) {
    return state.snapshot?.chart_range.start || todayDate();
  }
  const text = String(value).trim();
  return text.length >= 10 ? text.slice(0, 10) : text;
}

function toDateTimeText(value, endOfDay) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return `${date.toISOString().slice(0, 10)}T${endOfDay ? "17:00:00" : "08:00:00"}`;
}

function todayDate() {
  return new Date().toISOString().slice(0, 10);
}

function blankToNull(value) {
  const text = String(value ?? "").trim();
  if (!text) {
    return null;
  }
  if (/^\d+$/.test(text)) {
    return Number(text);
  }
  return text;
}

function clampInt(value, fallback = 0) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function clampNumber(value, min, max) {
  const parsed = Number.parseFloat(value);
  if (!Number.isFinite(parsed)) {
    return min;
  }
  return Math.max(min, Math.min(max, parsed));
}

function stringifyError(error) {
  if (typeof error === "string") {
    return error;
  }
  if (error && typeof error === "object") {
    return error.message || JSON.stringify(error);
  }
  return String(error);
}

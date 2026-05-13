const SHIFT_LABELS = {
  day: "주",
  night: "야",
  off_night: "비",
  off: "휴",
  leave: "연",
};

const RAW_TO_CODE = {
  0: "day",
  1: "night",
  2: "off_night",
  3: "off",
  4: "leave",
};

const STATE_FILE_FORMAT = "work-scheduler-v3-state";
const STATE_FILE_VERSION = 2;
const TOAST_GAP_PX = 8;
const TOAST_EXIT_MS = 220;
const TOAST_DUPLICATE_WINDOW_MS = 1200;

const state = {
  year: null,
  month: null,
  daysInMonth: 0,
  commonTargetHours: 0,
  workers: [],
  settings: {},
  flags: {
    useIndividualTargets: false,
    useDayOnly: false,
    useWorkPeriod: false,
    showShiftColors: true,
  },
  specialShifts: {},
  uploadedTemplate: null,
  histories: [],
  selectedHistory: 0,
  randomSeed: 0,
  repairEditing: false,
  repairBaseResult: null,
  repairDraftResult: null,
  repairEdits: {},
  guideTimer: null,
  guideRequestId: 0,
  isUploading: false,
  isSolving: false,
  solveAbortController: null,
  activeSolveTaskId: null,
  isDownloading: false,
  sidebarCollapsed: false,
  activeHelpTarget: null,
  helpTip: null,
  toasts: [],
  nextToastId: 1,
  progressTasks: {},
  hasUnsavedChanges: false,
  isReady: false,
};

const $ = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", init);
document.addEventListener(
  "wheel",
  (event) => {
    if (event.target instanceof HTMLSelectElement) {
      event.preventDefault();
      event.target.blur();
    }
  },
  { passive: false },
);

async function init() {
  bindEvents();
  const data = await requestJson("/api/default-state");
  applyDefaults(data.defaults);
  state.workers = data.workers;
  renderAll();
  refreshGuide();
  setStatus("준비됨");
  state.isReady = true;
}

function bindEvents() {
  $("yearInput").addEventListener("change", handleMonthChange);
  $("monthInput").addEventListener("change", handleMonthChange);
  $("commonTargetInput").addEventListener("change", () => {
    markDirty();
    state.commonTargetHours = toInt($("commonTargetInput").value, state.commonTargetHours);
    scheduleGuideRefresh();
  });

  for (const id of [
    "useIndividualTargets",
    "useDayOnly",
    "useWorkPeriod",
    "usePreference",
    "allowLeaveAfterOffNight",
    "allowDoubleNightCycle",
    "showShiftColors",
    "useEmergencyRange",
  ]) {
    $(id).addEventListener("change", () => handleConditionToggleChange(id));
  }

  for (const id of [
    "targetDay",
    "targetNight",
    "minDay",
    "maxDay",
    "minNight",
    "maxNight",
    "emergencyMinDay",
    "emergencyMaxDay",
    "emergencyMinNight",
    "emergencyMaxNight",
  ]) {
    $(id).addEventListener("change", () => {
      markDirty();
      collectSidebar();
      scheduleGuideRefresh();
    });
  }

  $("templateFile").addEventListener("change", uploadTemplate);
  $("removeUploadBtn").addEventListener("click", removeUpload);
  $("addSpecialBtn").addEventListener("click", addSpecialShift);
  $("solveBtn").addEventListener("click", () => solveSchedule(false));
  $("rerollBtn").addEventListener("click", () => solveSchedule(true));
  $("solveCancelBtn").addEventListener("click", () => cancelActiveSolve());
  $("resultSolveCancelBtn").addEventListener("click", () => cancelActiveSolve());
  $("repairEditBtn").addEventListener("click", startRepairEdit);
  $("repairApplyBtn").addEventListener("click", applyRepairEdit);
  $("repairCancelBtn").addEventListener("click", cancelRepairEdit);
  $("resetInputBtn").addEventListener("click", resetInput);
  $("addWorkerBtn").addEventListener("click", addWorker);
  $("saveStateBtn").addEventListener("click", saveInputState);
  $("loadStateBtn").addEventListener("click", () => $("stateFileInput").click());
  $("stateFileInput").addEventListener("change", loadInputState);
  $("downloadBtn").addEventListener("click", downloadExcel);
  $("sidebarToggleBtn").addEventListener("click", toggleSidebar);
  $("historySelect").addEventListener("change", (event) => {
    if (state.repairEditing) cancelRepairEdit(false);
    state.selectedHistory = Number(event.target.value);
    renderHistory();
    renderResult();
  });
  $("resetSettingsBtn").addEventListener("click", () => $("confirmDialog").showModal());
  $("confirmResetSettingsBtn").addEventListener("click", resetSettings);

  $("inputTableWrap").addEventListener("change", handleTableChange);
  $("inputTableWrap").addEventListener("input", handleTableChange);
  $("inputTableWrap").addEventListener("click", handleInputTableClick);
  $("inputTableWrap").addEventListener("keydown", handleInputTableKeydown);
  $("resultTableWrap").addEventListener("change", handleResultTableChange);
  document.addEventListener("mouseover", handleHelpEnter);
  document.addEventListener("focusin", handleHelpEnter);
  document.addEventListener("mouseout", handleHelpLeave);
  document.addEventListener("focusout", handleHelpLeave);
  document.addEventListener("scroll", positionActiveHelpTip, true);
  window.addEventListener("resize", positionActiveHelpTip);
  window.addEventListener("resize", updateToastStack);
  window.addEventListener("beforeunload", handleBeforeUnload);
}

function handleConditionToggleChange(id) {
  markDirty();
  collectSidebar();

  const inputTableOptions = new Set([
    "useIndividualTargets",
    "useDayOnly",
    "useWorkPeriod",
    "usePreference",
  ]);

  if (id === "useWorkPeriod") trimFixedShiftsToWorkPeriods();
  if (id === "useDayOnly" && state.flags.useDayOnly) {
    state.workers.forEach((worker) => pruneDedicatedFixedShifts(worker));
  }

  if (inputTableOptions.has(id)) {
    renderInputTable();
  }

  if (id === "showShiftColors") {
    updateRenderedShiftColors();
  }

  if (state.repairEditing && ["useDayOnly", "useWorkPeriod"].includes(id)) {
    renderResult();
  }

  scheduleGuideRefresh();
}

function markDirty() {
  if (!state.isReady) return;
  state.hasUnsavedChanges = true;
}

function handleBeforeUnload(event) {
  if (!state.hasUnsavedChanges) return;
  event.preventDefault();
  event.returnValue = "";
}

async function handleMonthChange() {
  markDirty();
  const year = toInt($("yearInput").value, state.year);
  const month = toInt($("monthInput").value, state.month);
  const info = await requestJson(`/api/month-info?year=${year}&month=${month}`);
  state.year = info.year;
  state.month = info.month;
  state.daysInMonth = info.days_in_month;
  state.commonTargetHours = info.default_target_hours;
  $("commonTargetInput").value = state.commonTargetHours;
  trimFixedShiftsToMonth();
  renderAll();
  refreshGuide();
}

function applyDefaults(defaults) {
  const info = defaults.month_info;
  state.year = info.year;
  state.month = info.month;
  state.daysInMonth = info.days_in_month;
  state.commonTargetHours = defaults.common_target_hours;
  state.settings = { ...defaults.settings };
  state.specialShifts = { ...state.settings.special_shifts };
  state.flags.useIndividualTargets = Boolean(defaults.use_individual_targets);
  state.flags.useDayOnly = Boolean(defaults.use_day_only);
  state.flags.useWorkPeriod = Boolean(defaults.use_work_period);
  state.flags.showShiftColors = Boolean(defaults.show_shift_colors);
}

function collectSidebar() {
  state.year = toInt($("yearInput").value, state.year);
  state.month = toInt($("monthInput").value, state.month);
  state.commonTargetHours = toInt($("commonTargetInput").value, state.commonTargetHours);
  state.flags.useIndividualTargets = $("useIndividualTargets").checked;
  state.flags.useDayOnly = $("useDayOnly").checked;
  state.flags.useWorkPeriod = $("useWorkPeriod").checked;
  state.flags.showShiftColors = $("showShiftColors").checked;

  state.settings = {
    target_day: toInt($("targetDay").value, 1),
    target_night: toInt($("targetNight").value, 2),
    min_day: toInt($("minDay").value, 1),
    max_day: toInt($("maxDay").value, 2),
    min_night: toInt($("minNight").value, 1),
    max_night: toInt($("maxNight").value, 2),
    use_emergency_range: $("useEmergencyRange").checked,
    emergency_min_day: optionalInt($("emergencyMinDay").value),
    emergency_max_day: optionalInt($("emergencyMaxDay").value),
    emergency_min_night: optionalInt($("emergencyMinNight").value),
    emergency_max_night: optionalInt($("emergencyMaxNight").value),
    use_preference: $("usePreference").checked,
    allow_leave_after_off_night: $("allowLeaveAfterOffNight").checked,
    allow_double_night_cycle: $("allowDoubleNightCycle").checked,
    special_shifts: { ...state.specialShifts },
  };
}

function syncSidebarControls() {
  $("yearInput").value = state.year;
  $("monthInput").value = state.month;
  $("commonTargetInput").value = state.commonTargetHours;
  $("useIndividualTargets").checked = state.flags.useIndividualTargets;
  $("useDayOnly").checked = state.flags.useDayOnly;
  $("useWorkPeriod").checked = state.flags.useWorkPeriod;
  $("usePreference").checked = Boolean(state.settings.use_preference);
  $("allowLeaveAfterOffNight").checked = Boolean(state.settings.allow_leave_after_off_night);
  $("allowDoubleNightCycle").checked = Boolean(state.settings.allow_double_night_cycle);
  $("showShiftColors").checked = state.flags.showShiftColors;
  $("useEmergencyRange").checked = Boolean(state.settings.use_emergency_range);

  $("targetDay").value = state.settings.target_day ?? 1;
  $("targetNight").value = state.settings.target_night ?? 2;
  $("minDay").value = state.settings.min_day ?? 1;
  $("maxDay").value = state.settings.max_day ?? 2;
  $("minNight").value = state.settings.min_night ?? 1;
  $("maxNight").value = state.settings.max_night ?? 2;
  $("emergencyMinDay").value = state.settings.emergency_min_day ?? "";
  $("emergencyMaxDay").value = state.settings.emergency_max_day ?? "";
  $("emergencyMinNight").value = state.settings.emergency_min_night ?? "";
  $("emergencyMaxNight").value = state.settings.emergency_max_night ?? "";
}

function renderAll() {
  syncSidebarControls();
  $("monthLabel").textContent = `${state.year}년 ${state.month}월 | 총 ${state.daysInMonth}일`;
  renderUploadInfo();
  renderSpecialList();
  renderInputTable();
  renderHistory();
  renderResult();
}

function renderUploadInfo() {
  const box = $("uploadInfo");
  if (!state.uploadedTemplate) {
    box.classList.add("hidden");
    $("uploadName").textContent = "";
    return;
  }
  $("uploadName").textContent = state.uploadedTemplate.filename;
  box.classList.remove("hidden");
}

function renderSpecialList() {
  const list = $("specialList");
  list.innerHTML = "";
  for (const [code, hours] of Object.entries(state.specialShifts)) {
    const item = document.createElement("span");
    item.className = "special-item";
    item.innerHTML = `<strong>${escapeHtml(code)}</strong> ${hours}h <button type="button" title="삭제" data-special-delete="${escapeHtml(code)}">×</button>`;
    list.appendChild(item);
  }
  list.querySelectorAll("[data-special-delete]").forEach((button) => {
    button.addEventListener("click", () => {
      markDirty();
      delete state.specialShifts[button.dataset.specialDelete];
      collectSidebar();
      renderSpecialList();
      renderInputTable();
      scheduleGuideRefresh();
    });
  });
}

function renderInputTable() {
  const wrap = $("inputTableWrap");
  const metaColumns = [
    { key: "name", label: "이름" },
    ...(state.settings.use_preference ? [{ key: "preference", label: "선호" }] : []),
    ...(state.flags.useDayOnly ? [{ key: "dedicated_shift", label: "전담" }] : []),
    ...(state.flags.useIndividualTargets ? [{ key: "target_hours", label: "목표" }] : []),
    ...(state.flags.useWorkPeriod
      ? [
          { key: "start_day", label: "시작" },
          { key: "end_day", label: "종료" },
        ]
      : []),
  ];
  const head = [
    `<th class="row-index-head row-index-sticky"></th>`,
    ...metaColumns.map((col, index) => `<th class="${index === 0 ? "name-sticky" : ""} meta-head meta-head-${col.key}">${col.label}</th>`),
    ...range(1, state.daysInMonth).map((day) => `<th class="date-head">${day}</th>`),
  ].join("");

  const body = state.workers
    .map((worker, rowIndex) => {
      const meta = metaColumns
        .map((col, colIndex) => `<td class="${colIndex === 0 ? "name-sticky name-cell" : "meta-cell"}">${renderMetaControl(worker, rowIndex, col.key)}</td>`)
        .join("");
      const days = range(1, state.daysInMonth)
        .map((day) => {
          const outsidePeriod = isOutsideWorkPeriod(worker, day);
          const value = outsidePeriod ? "" : worker.fixed_shifts?.[day] || "";
          return `<td class="day-cell ${outsidePeriod ? "outside-period-cell" : ""}">${renderShiftSelect(rowIndex, day, value, outsidePeriod)}</td>`;
        })
        .join("");
      return `<tr class="input-worker-row"><td class="row-index-cell row-index-sticky input-row-index">${renderRowIndexControl(rowIndex)}</td>${meta}${days}</tr>`;
    })
    .join("");

  wrap.innerHTML = `<table class="schedule-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function renderRowIndexControl(rowIndex) {
  const rowNumber = rowIndex + 1;
  const disabled = state.workers.length <= 1 ? "disabled" : "";
  return `<div class="row-index-content">
    <span class="row-number">${rowNumber}</span>
    <button class="row-delete-btn" type="button" data-delete-row="${rowIndex}" aria-label="${rowNumber}행 삭제" title="${rowNumber}행 삭제" ${disabled}>×</button>
  </div>`;
}

function renderMetaControl(worker, rowIndex, key) {
  const attr = `data-row="${rowIndex}" data-field="${key}"`;
  if (key === "name") {
    return `<input ${attr} value="${escapeHtml(worker.name || "")}" />`;
  }
  if (key === "preference") {
    return `<select ${attr} class="compact-select">
      <option value="" ${!worker.preference ? "selected" : ""}>-</option>
      <option value="day" ${worker.preference === "day" ? "selected" : ""}>주간</option>
      <option value="night" ${worker.preference === "night" ? "selected" : ""}>야간</option>
    </select>`;
  }
  if (key === "dedicated_shift") {
    const value = workerDedicatedShift(worker);
    return `<select ${attr} class="compact-select">
      <option value="" ${!value ? "selected" : ""}>-</option>
      <option value="day" ${value === "day" ? "selected" : ""}>주간</option>
      <option value="night" ${value === "night" ? "selected" : ""}>야간</option>
    </select>`;
  }
  if (key === "target_hours") {
    return `<input ${attr} type="number" min="0" step="8" value="${worker.target_hours ?? state.commonTargetHours}" />`;
  }
  if (key === "start_day") {
    return `<input ${attr} type="number" min="1" max="${state.daysInMonth}" value="${worker.start_day ?? 1}" />`;
  }
  if (key === "end_day") {
    return `<input ${attr} type="number" min="1" max="${state.daysInMonth}" value="${worker.end_day ?? state.daysInMonth}" />`;
  }
  return "";
}

function renderShiftSelect(rowIndex, day, value, disabled = false) {
  const options = shiftOptionsForWorker(rowIndex)
    .map((option) => `<option value="${escapeHtml(option.value)}" ${option.value === value ? "selected" : ""}>${escapeHtml(option.label)}</option>`)
    .join("");
  const disabledAttr = disabled ? ` disabled title="근무 기간 밖"` : "";
  return `<select data-row="${rowIndex}" data-day="${day}" class="compact-select ${disabled ? "outside-period-select" : shiftClass(value)}"${disabledAttr}>${options}</select>`;
}

function shiftOptions() {
  return [
    { value: "", label: "" },
    ...Object.entries(SHIFT_LABELS).map(([value, label]) => ({ value, label })),
    ...Object.keys(state.specialShifts).map((code) => ({ value: code, label: code })),
  ];
}

function shiftOptionsForWorker(rowIndex) {
  const worker = state.workers[rowIndex] || {};
  return shiftOptions().filter((option) => isShiftAllowedForWorker(worker, option.value));
}

function workerDedicatedShift(worker = {}) {
  const value = String(worker.dedicated_shift || "").trim();
  if (value === "day" || value === "night") return value;
  return worker.is_day_only ? "day" : "";
}

function setWorkerDedicatedShift(worker, value) {
  const dedicatedShift = ["day", "night"].includes(value) ? value : "";
  worker.dedicated_shift = dedicatedShift;
  worker.is_day_only = dedicatedShift === "day";
}

function isShiftAllowedForWorker(worker, shift) {
  if (!shift || !(shift in SHIFT_LABELS)) return true;
  const dedicatedShift = workerDedicatedShift(worker);
  if (dedicatedShift === "day") return !["night", "off_night"].includes(shift);
  if (dedicatedShift === "night") return shift !== "day";
  return true;
}

function pruneDedicatedFixedShifts(worker) {
  if (!worker?.fixed_shifts) return false;
  let changed = false;
  for (const [day, shift] of Object.entries(worker.fixed_shifts)) {
    if (!isShiftAllowedForWorker(worker, shift)) {
      delete worker.fixed_shifts[day];
      changed = true;
    }
  }
  return changed;
}

function repairShiftOptions(rowIndex) {
  return shiftOptionsForWorker(rowIndex);
}

function handleTableChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;
  const rowIndex = Number(target.dataset.row);
  if (!Number.isInteger(rowIndex) || !state.workers[rowIndex]) return;
  markDirty();
  const worker = state.workers[rowIndex];

  if (target.dataset.day) {
    const day = Number(target.dataset.day);
    worker.fixed_shifts ||= {};
    if (target.value) {
      worker.fixed_shifts[day] = target.value;
    } else {
      delete worker.fixed_shifts[day];
    }
    target.className = `compact-select ${shiftClass(target.value)}`;
    scheduleGuideRefresh();
    return;
  }

  const field = target.dataset.field;
  if (!field) return;
  if (target.type === "checkbox") {
    worker[field] = target.checked;
  } else if (field === "dedicated_shift") {
    setWorkerDedicatedShift(worker, target.value);
    pruneDedicatedFixedShifts(worker);
    renderInputTable();
  } else if (["target_hours", "start_day", "end_day"].includes(field)) {
    worker[field] = optionalInt(target.value);
    if (event.type === "change" && ["start_day", "end_day"].includes(field)) {
      trimWorkerFixedShiftsToWorkPeriod(worker);
      refreshWorkPeriodCells(rowIndex);
    }
  } else {
    worker[field] = target.value;
  }
  scheduleGuideRefresh();
}

function handleInputTableClick(event) {
  const target = event.target;
  if (!(target instanceof Element)) return;
  if (target instanceof HTMLInputElement && ["target_hours", "start_day", "end_day"].includes(target.dataset.field)) {
    requestAnimationFrame(() => target.select());
    return;
  }
  const button = target.closest("[data-delete-row]");
  if (!(button instanceof HTMLButtonElement)) return;
  const rowIndex = Number(button.dataset.deleteRow);
  removeWorkerAt(rowIndex);
}

function handleInputTableKeydown(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (target.dataset.field !== "name") return;
  if (event.isComposing || !["Enter", "Tab"].includes(event.key)) return;

  event.preventDefault();
  const rowIndex = Number(target.dataset.row);
  if (!Number.isInteger(rowIndex) || !state.workers[rowIndex]) return;
  state.workers[rowIndex].name = target.value;
  const nextRow = event.shiftKey && event.key === "Tab" ? rowIndex - 1 : rowIndex + 1;

  if (nextRow >= state.workers.length) {
    addWorker();
    focusNameInput(nextRow);
    return;
  }

  if (nextRow >= 0) {
    focusNameInput(nextRow);
  }
}

function focusNameInput(rowIndex) {
  requestAnimationFrame(() => {
    const input = $("inputTableWrap").querySelector(`input[data-field="name"][data-row="${rowIndex}"]`);
    if (!input) return;
    input.focus();
    input.select();
  });
}

function toggleSidebar() {
  state.sidebarCollapsed = !state.sidebarCollapsed;
  document.body.classList.toggle("sidebar-collapsed", state.sidebarCollapsed);
  const button = $("sidebarToggleBtn");
  button.textContent = state.sidebarCollapsed ? "›" : "‹";
  button.setAttribute("aria-expanded", String(!state.sidebarCollapsed));
  button.title = state.sidebarCollapsed ? "좌측 바 펼치기" : "좌측 바 접기";
  hideHelpTip();
}

function handleHelpEnter(event) {
  const target = event.target.closest?.(".help");
  if (!target) return;
  showHelpTip(target);
}

function handleHelpLeave(event) {
  const target = event.target.closest?.(".help");
  if (!target || target !== state.activeHelpTarget) return;
  hideHelpTip();
}

function showHelpTip(target) {
  const text = target.dataset.tip;
  if (!text) return;
  if (!state.helpTip) {
    state.helpTip = document.createElement("div");
    state.helpTip.className = "floating-help-tip";
    document.body.appendChild(state.helpTip);
  }
  state.activeHelpTarget = target;
  state.helpTip.textContent = text;
  state.helpTip.classList.remove("hidden");
  positionActiveHelpTip();
}

function hideHelpTip() {
  state.activeHelpTarget = null;
  if (state.helpTip) state.helpTip.classList.add("hidden");
}

function positionActiveHelpTip() {
  if (!state.activeHelpTarget || !state.helpTip || state.helpTip.classList.contains("hidden")) return;
  const targetRect = state.activeHelpTarget.getBoundingClientRect();
  const tipRect = state.helpTip.getBoundingClientRect();
  const margin = 10;
  const viewportWidth = document.documentElement.clientWidth;
  const viewportHeight = document.documentElement.clientHeight;
  const left = clamp(targetRect.left + targetRect.width / 2 - tipRect.width / 2, margin, viewportWidth - tipRect.width - margin);
  let top = targetRect.top - tipRect.height - 8;
  if (top < margin) top = targetRect.bottom + 8;
  top = clamp(top, margin, viewportHeight - tipRect.height - margin);
  state.helpTip.style.left = `${left}px`;
  state.helpTip.style.top = `${top}px`;
}

async function uploadTemplate(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const form = new FormData();
  form.append("file", file);
  if (state.uploadedTemplate?.file_id) {
    form.append("previous_file_id", state.uploadedTemplate.file_id);
  }

  try {
    setUploadBusy(true);
    startProgressTask("upload", "엑셀 업로드 중", 7000);
    setStatus("엑셀 업로드 중");
    const data = await uploadJsonWithProgress("/api/upload-template", form, (percent, label) => {
      setProgressTask("upload", percent, label);
    });
    completeProgressTask("upload", "업로드 완료");
    markDirty();
    state.uploadedTemplate = { file_id: data.file_id, filename: data.filename };
    if (!data.load_result.preserve_existing_workers && Array.isArray(data.load_result.workers) && data.load_result.workers.length > 0) {
      state.workers = data.load_result.workers;
    }
    clearMessage();
    if (data.load_result.warning) {
      showTopToast(`엑셀 업로드 완료. ${data.load_result.warning}`, "warning", 7000);
    } else {
      showTopToast("엑셀 업로드 완료", "success", 3500);
    }
    renderAll();
    refreshGuide();
    setStatus("업로드 완료");
  } catch (error) {
    failProgressTask("upload", "업로드 실패");
    showTopToast(error.message, "error", 7000);
    setStatus("업로드 실패");
  } finally {
    setUploadBusy(false);
    event.target.value = "";
  }
}

async function removeUpload() {
  if (!state.uploadedTemplate) return;
  markDirty();
  await requestJson(`/api/upload-template/${state.uploadedTemplate.file_id}`, { method: "DELETE" });
  state.uploadedTemplate = null;
  renderUploadInfo();
  renderResult();
}

function addSpecialShift() {
  const code = $("specialCodeInput").value.trim();
  const hours = optionalInt($("specialHoursInput").value);
  if (!code || code.length !== 1 || ["주", "야", "비", "휴", "연", "-"].includes(code)) {
    showMessage("기타 근무 약어를 확인해 주세요.", "error");
    return;
  }
  if (hours === null || hours < 0 || hours > 24) {
    showMessage("기타 근무 인정시간을 확인해 주세요.", "error");
    return;
  }
  markDirty();
  state.specialShifts[code] = hours;
  $("specialCodeInput").value = "";
  $("specialHoursInput").value = "";
  collectSidebar();
  renderSpecialList();
  renderInputTable();
  scheduleGuideRefresh();
}

async function solveSchedule(isReroll) {
  if (state.isSolving) return;
  collectSidebar();
  if (isReroll) state.randomSeed += 1;
  const payload = buildSchedulePayload();
  const progressTaskId = isReroll ? "resultSolve" : "solve";
  const controller = new AbortController();
  let guideForToast = null;
  state.solveAbortController = controller;
  state.activeSolveTaskId = progressTaskId;
  setSolveBusy(true);
  startProgressTask(progressTaskId, isReroll ? "근무표 다시 생성 중" : "근무표 생성 중", 42000);
  setStatus("계산 중");
  clearMessage();

  try {
    const guideData = await requestJson("/api/leave-guide", { method: "POST", json: payload, signal: controller.signal });
    if (controller.signal.aborted) return;
    guideForToast = guideData.guide;
    renderGuide(guideData.guide);
    const solveData = await requestJson("/api/solve", { method: "POST", json: payload, signal: controller.signal });
    if (controller.signal.aborted) return;
    if (!solveData.success) throw new Error(solveData.error || "생성에 실패했습니다.");
    completeProgressTask(progressTaskId, "생성 완료");
    markDirty();
    state.histories.unshift({
      label: `${new Date().toLocaleTimeString()} 생성`,
      result: solveData.result,
    });
    state.selectedHistory = 0;
    cancelRepairEdit(false);
    $("rerollBtn").disabled = false;
    $("downloadBtn").disabled = false;
    renderHistory();
    renderResult();
    setStatus("생성 완료");
  } catch (error) {
    if (isAbortError(error)) {
      cancelProgressTask(progressTaskId, "생성 취소");
      setStatus("생성 취소");
      return;
    }
    failProgressTask(progressTaskId, "생성 실패");
    showTopToast(scheduleFailureToastMessage(error.message, guideForToast), "error", 7000);
    setStatus("생성 실패");
  } finally {
    if (state.solveAbortController === controller) {
      state.solveAbortController = null;
      state.activeSolveTaskId = null;
      setSolveBusy(false);
    }
  }
}

function buildSchedulePayload() {
  return {
    year: state.year,
    month: state.month,
    workers: state.workers.map((worker) => {
      const dedicatedShift = state.flags.useDayOnly ? workerDedicatedShift(worker) : "";
      return {
        name: worker.name || "",
        start_day: state.flags.useWorkPeriod ? worker.start_day || 1 : 1,
        end_day: state.flags.useWorkPeriod ? worker.end_day || state.daysInMonth : null,
        is_day_only: dedicatedShift === "day",
        dedicated_shift: dedicatedShift || null,
        target_hours: state.flags.useIndividualTargets ? worker.target_hours ?? state.commonTargetHours : state.commonTargetHours,
        preference: state.settings.use_preference ? worker.preference || null : null,
        prev_month_last_day_night: Boolean(worker.prev_month_last_day_night),
        fixed_shifts: buildFixedShiftsForPayload(worker),
      };
    }),
    settings: {
      ...state.settings,
      special_shifts: { ...state.specialShifts },
    },
    random_seed: state.randomSeed,
  };
}

function buildFixedShiftsForPayload(worker) {
  if (!state.flags.useWorkPeriod) return { ...(worker.fixed_shifts || {}) };
  const filtered = {};
  for (const [dayText, shift] of Object.entries(worker.fixed_shifts || {})) {
    const day = Number(dayText);
    if (!isOutsideWorkPeriod(worker, day)) filtered[dayText] = shift;
  }
  return filtered;
}

function isOutsideWorkPeriod(worker, day) {
  if (!state.flags.useWorkPeriod) return false;
  const { start, end } = workerPeriodBounds(worker);
  return day < start || day > end;
}

function workerPeriodBounds(worker) {
  const start = Number.isInteger(worker.start_day) ? worker.start_day : 1;
  const end = Number.isInteger(worker.end_day) ? worker.end_day : state.daysInMonth;
  return { start, end };
}

function trimWorkerFixedShiftsToWorkPeriod(worker) {
  if (!state.flags.useWorkPeriod) return;
  worker.fixed_shifts = buildFixedShiftsForPayload(worker);
}

function trimFixedShiftsToWorkPeriods() {
  if (!state.flags.useWorkPeriod) return;
  for (const worker of state.workers) {
    trimWorkerFixedShiftsToWorkPeriod(worker);
  }
}

function refreshWorkPeriodCells(rowIndex) {
  const worker = state.workers[rowIndex];
  if (!worker) return;
  for (const day of range(1, state.daysInMonth)) {
    const select = $("inputTableWrap").querySelector(`select[data-row="${rowIndex}"][data-day="${day}"]`);
    if (!select) continue;
    const cell = select.closest(".day-cell");
    const outsidePeriod = isOutsideWorkPeriod(worker, day);
    if (outsidePeriod) {
      select.value = "";
      select.disabled = true;
      select.title = "근무 기간 밖";
      select.className = "compact-select outside-period-select";
      cell?.classList.add("outside-period-cell");
    } else {
      select.disabled = false;
      select.title = "";
      select.className = `compact-select ${shiftClass(select.value)}`;
      cell?.classList.remove("outside-period-cell");
    }
  }
}

async function resetInput() {
  markDirty();
  const data = await requestJson("/api/reset-input", {
    method: "POST",
    json: { workers: state.workers },
  });
  state.workers = data.workers;
  renderInputTable();
  scheduleGuideRefresh();
  clearMessage();
}

async function resetSettings(event) {
  event.preventDefault();
  markDirty();
  const data = await requestJson("/api/reset-settings", {
    method: "POST",
    json: {},
  });
  applyDefaults(data.defaults);
  trimFixedShiftsToMonth();
  collectSidebar();
  renderAll();
  refreshGuide();
  $("confirmDialog").close();
}

function addWorker() {
  markDirty();
  state.workers.push({
    name: "",
    start_day: 1,
    end_day: null,
    is_day_only: false,
    dedicated_shift: "",
    target_hours: null,
    preference: null,
    prev_month_last_day_night: false,
    fixed_shifts: {},
  });
  renderInputTable();
  scheduleGuideRefresh();
}

function removeWorkerAt(rowIndex) {
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= state.workers.length) return;
  if (state.workers.length <= 1) return;
  markDirty();
  state.workers.splice(rowIndex, 1);
  renderInputTable();
  scheduleGuideRefresh();
}

function saveInputState() {
  collectSidebar();
  collectInputTableState();
  const snapshot = buildStateSnapshot();
  const blob = new Blob([JSON.stringify(snapshot, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = stateSnapshotFilename();
  link.click();
  URL.revokeObjectURL(url);
  state.hasUnsavedChanges = false;
  setStatus("입력 상태 저장 완료");
  showMessage("현재 입력표와 설정을 저장했습니다.", "success");
}

async function loadInputState(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    setStatus("입력 상태 불러오는 중");
    const snapshot = parseStateSnapshot(await file.text());
    const info = await requestJson(`/api/month-info?year=${snapshotYear(snapshot)}&month=${snapshotMonth(snapshot)}`);
    applyStateSnapshot(snapshot, info);
    renderAll();
    refreshGuide();
    state.hasUnsavedChanges = false;
    clearMessage();
    setStatus("입력 상태 불러오기 완료");
    if (snapshot.uploadedTemplate?.filename) {
      showTopToast("입력 상태를 불러왔습니다. 엑셀 원본 파일은 저장 파일에 포함되지 않으므로 필요하면 다시 업로드해 주세요.", "success", 7000);
    } else {
      showTopToast("입력 상태를 불러왔습니다.", "success", 3500);
    }
  } catch (error) {
    showTopToast(error.message, "error", 7000);
    setStatus("입력 상태 불러오기 실패");
  } finally {
    event.target.value = "";
  }
}

function buildStateSnapshot() {
  const options = {
    useIndividualTargets: state.flags.useIndividualTargets,
    useDayOnly: state.flags.useDayOnly,
    useWorkPeriod: state.flags.useWorkPeriod,
    usePreference: Boolean(state.settings.use_preference),
    allowLeaveAfterOffNight: Boolean(state.settings.allow_leave_after_off_night),
    allowDoubleNightCycle: Boolean(state.settings.allow_double_night_cycle),
    showShiftColors: state.flags.showShiftColors,
    useEmergencyRange: Boolean(state.settings.use_emergency_range),
  };

  return {
    format: STATE_FILE_FORMAT,
    version: STATE_FILE_VERSION,
    saved_at: new Date().toISOString(),
    year: state.year,
    month: state.month,
    commonTargetHours: state.commonTargetHours,
    common_target_hours: state.commonTargetHours,
    basic: {
      year: state.year,
      month: state.month,
      commonTargetHours: state.commonTargetHours,
      common_target_hours: state.commonTargetHours,
    },
    flags: { ...state.flags },
    options,
    settings: {
      ...state.settings,
      special_shifts: { ...state.specialShifts },
    },
    specialShifts: { ...state.specialShifts },
    workers: state.workers.map((worker) => ({
      name: worker.name || "",
      start_day: worker.start_day ?? 1,
      end_day: state.flags.useWorkPeriod ? worker.end_day ?? state.daysInMonth : worker.end_day ?? null,
      is_day_only: workerDedicatedShift(worker) === "day",
      dedicated_shift: workerDedicatedShift(worker) || null,
      target_hours: state.flags.useIndividualTargets ? worker.target_hours ?? state.commonTargetHours : worker.target_hours ?? null,
      preference: worker.preference || null,
      prev_month_last_day_night: Boolean(worker.prev_month_last_day_night),
      fixed_shifts: { ...(worker.fixed_shifts || {}) },
    })),
    randomSeed: state.randomSeed,
    uploadedTemplate: state.uploadedTemplate
      ? { filename: state.uploadedTemplate.filename || "" }
      : null,
  };
}

function collectInputTableState() {
  const controls = $("inputTableWrap").querySelectorAll("[data-row]");
  controls.forEach((control) => {
    if (!(control instanceof HTMLInputElement || control instanceof HTMLSelectElement)) return;
    const rowIndex = Number(control.dataset.row);
    const worker = state.workers[rowIndex];
    if (!Number.isInteger(rowIndex) || !worker) return;

    if (control.dataset.day) {
      const day = Number(control.dataset.day);
      if (!Number.isInteger(day) || control.disabled) return;
      worker.fixed_shifts ||= {};
      if (control.value) {
        worker.fixed_shifts[day] = control.value;
      } else {
        delete worker.fixed_shifts[day];
      }
      return;
    }

    const field = control.dataset.field;
    if (!field) return;
    if (control.type === "checkbox") {
      worker[field] = control.checked;
    } else if (field === "dedicated_shift") {
      setWorkerDedicatedShift(worker, control.value);
      pruneDedicatedFixedShifts(worker);
    } else if (["target_hours", "start_day", "end_day"].includes(field)) {
      worker[field] = optionalInt(control.value);
    } else {
      worker[field] = control.value;
    }
  });
}

function parseStateSnapshot(text) {
  let snapshot;
  try {
    snapshot = JSON.parse(text);
  } catch (error) {
    throw new Error("저장 파일을 읽을 수 없습니다. JSON 파일인지 확인해 주세요.");
  }
  if (!snapshot || snapshot.format !== STATE_FILE_FORMAT) {
    throw new Error("이 앱에서 내보낸 입력 상태 파일이 아닙니다.");
  }
  if (Number(snapshot.version) > STATE_FILE_VERSION) {
    throw new Error("현재 앱보다 새 버전에서 만든 저장 파일입니다. 앱을 업데이트한 뒤 다시 시도해 주세요.");
  }
  if (!Number.isInteger(Number(snapshotYear(snapshot))) || !Number.isInteger(Number(snapshotMonth(snapshot)))) {
    throw new Error("저장 파일의 연도/월 값이 올바르지 않습니다.");
  }
  if (!Array.isArray(snapshot.workers) || snapshot.workers.length === 0) {
    throw new Error("저장 파일에 근무자 입력표가 없습니다.");
  }
  return snapshot;
}

function applyStateSnapshot(snapshot, monthInfo) {
  state.year = Number(monthInfo.year);
  state.month = Number(monthInfo.month);
  state.daysInMonth = Number(monthInfo.days_in_month);
  state.commonTargetHours = toInt(snapshotCommonTargetHours(snapshot), monthInfo.default_target_hours);
  state.specialShifts = normalizeSpecialShifts(snapshot.specialShifts || snapshot.settings?.special_shifts || {});
  const options = snapshot.options || {};
  state.settings = normalizeSettings({ ...options, ...(snapshot.settings || snapshot.toSettings || {}) });
  state.flags = normalizeFlags(snapshot.flags || options || {}, options, state.settings);
  state.settings.special_shifts = { ...state.specialShifts };
  state.workers = snapshot.workers.map((worker) => normalizeWorkerSnapshot(worker));
  state.uploadedTemplate = null;
  state.histories = [];
  state.selectedHistory = 0;
  state.randomSeed = toInt(snapshot.randomSeed, 0);
  trimFixedShiftsToMonth();
  if (state.flags.useWorkPeriod) trimFixedShiftsToWorkPeriods();
}

function normalizeFlags(flags = {}, options = {}, settings = {}) {
  return {
    useIndividualTargets: boolFrom(flags.useIndividualTargets ?? flags.use_individual_targets ?? options.useIndividualTargets ?? options.use_individual_targets),
    useDayOnly: boolFrom(flags.useDayOnly ?? flags.use_day_only ?? options.useDayOnly ?? options.use_day_only),
    useWorkPeriod: boolFrom(flags.useWorkPeriod ?? flags.use_work_period ?? options.useWorkPeriod ?? options.use_work_period),
    showShiftColors: boolFrom(flags.showShiftColors ?? flags.show_shift_colors ?? options.showShiftColors ?? options.show_shift_colors, true),
  };
}

function normalizeSettings(settings = {}) {
  return {
    target_day: toInt(settings.target_day, 1),
    target_night: toInt(settings.target_night, 2),
    min_day: toInt(settings.min_day, 1),
    max_day: toInt(settings.max_day, 2),
    min_night: toInt(settings.min_night, 1),
    max_night: toInt(settings.max_night, 2),
    use_emergency_range: boolFrom(settings.use_emergency_range ?? settings.useEmergencyRange),
    emergency_min_day: nullableInt(settings.emergency_min_day),
    emergency_max_day: nullableInt(settings.emergency_max_day),
    emergency_min_night: nullableInt(settings.emergency_min_night),
    emergency_max_night: nullableInt(settings.emergency_max_night),
    use_preference: boolFrom(settings.use_preference ?? settings.usePreference),
    allow_leave_after_off_night: boolFrom(settings.allow_leave_after_off_night ?? settings.allowLeaveAfterOffNight),
    allow_double_night_cycle: boolFrom(settings.allow_double_night_cycle ?? settings.allowDoubleNightCycle),
    special_shifts: { ...state.specialShifts },
  };
}

function snapshotYear(snapshot) {
  return snapshot.year ?? snapshot.basic?.year;
}

function snapshotMonth(snapshot) {
  return snapshot.month ?? snapshot.basic?.month;
}

function snapshotCommonTargetHours(snapshot) {
  return snapshot.commonTargetHours ?? snapshot.common_target_hours ?? snapshot.basic?.commonTargetHours ?? snapshot.basic?.common_target_hours;
}

function boolFrom(value, fallback = false) {
  if (value === undefined || value === null) return fallback;
  if (typeof value === "string") return ["true", "1", "yes", "on"].includes(value.toLowerCase());
  return Boolean(value);
}

function normalizeSpecialShifts(specialShifts = {}) {
  const normalized = {};
  for (const [code, hours] of Object.entries(specialShifts)) {
    const token = String(code || "").trim();
    const parsedHours = nullableInt(hours);
    if (token.length !== 1 || parsedHours === null || parsedHours < 0 || parsedHours > 24) continue;
    if (["주", "야", "비", "휴", "연", "-"].includes(token)) continue;
    normalized[token] = parsedHours;
  }
  return normalized;
}

function normalizeWorkerSnapshot(worker = {}) {
  const startDay = nullableInt(worker.start_day);
  const endDay = nullableInt(worker.end_day);
  const dedicatedShift = ["day", "night"].includes(worker.dedicated_shift)
    ? worker.dedicated_shift
    : Boolean(worker.is_day_only)
      ? "day"
      : "";
  const normalized = {
    name: String(worker.name || ""),
    start_day: clamp(startDay ?? 1, 1, state.daysInMonth),
    end_day: endDay === null ? null : clamp(endDay, 1, state.daysInMonth),
    is_day_only: dedicatedShift === "day",
    dedicated_shift: dedicatedShift,
    target_hours: nullableInt(worker.target_hours),
    preference: ["day", "night"].includes(worker.preference) ? worker.preference : null,
    prev_month_last_day_night: Boolean(worker.prev_month_last_day_night),
    fixed_shifts: normalizeFixedShifts(worker.fixed_shifts || {}),
  };
  pruneDedicatedFixedShifts(normalized);
  return normalized;
}

function normalizeFixedShifts(fixedShifts = {}) {
  const allowedValues = new Set(["", ...Object.keys(SHIFT_LABELS), ...Object.keys(state.specialShifts)]);
  const normalized = {};
  for (const [dayText, shiftValue] of Object.entries(fixedShifts)) {
    const day = Number(dayText);
    const value = String(shiftValue || "").trim();
    if (!Number.isInteger(day) || day < 1 || day > state.daysInMonth) continue;
    if (!value || !allowedValues.has(value)) continue;
    normalized[day] = value;
  }
  return normalized;
}

function stateSnapshotFilename() {
  const stamp = new Date()
    .toISOString()
    .slice(0, 19)
    .replaceAll(":", "")
    .replace("T", "_");
  return `근무표입력_${state.year}_${state.month}_${stamp}.json`;
}

function renderGuide(guide) {
  const panel = $("guidePanel");
  panel.classList.remove("hidden");
  const creditedDays = Math.round((guide.non_to_credit_hours / 8) * 10) / 10;
  const needsExtraLeave = Number(guide.suggested_leave_days || 0) > 0 || Number(guide.shortage_hours || 0) > 0;
  const capacityLabel = guide.uses_emergency_range ? "최대 배치(예외)" : "최대 배치";
  const html = `
    <div class="guide-grid">
      ${metric("총 목표", `${guide.total_target_hours}h`, {
        tip: "현재 입력된 전체 근무자의 목표 근무시간 합계입니다.",
      })}
      ${metric("연가 인정", `${guide.leave_credit_hours}h`, {
        tip: "직접 지정한 연가의 인정시간 합계입니다.",
      })}
      ${metric("기타 인정", `${guide.special_credit_hours}h`, {
        tip: "직접 지정한 기타 근무의 인정시간 합계입니다.",
      })}
      ${metric("현재 인정일", `${creditedDays}일`, {
        tip: "연가 인정시간과 기타 근무 인정시간을 8시간 단위의 일수로 환산한 값입니다.",
      })}
      ${metric(capacityLabel, `${guide.max_regular_capacity_hours}h`, {
        tip: guide.uses_emergency_range
          ? "예외 범위 최대값까지 포함했을 때 정규 근무로 배치 가능한 최대 인정시간입니다."
          : "기본 인원 범위 기준으로 정규 근무를 최대한 배치했을 때 가능한 최대 인정시간입니다.",
      })}
      ${metric("연가 부족", `${guide.suggested_leave_days}일`, {
        danger: needsExtraLeave,
        tip: "인원 설정 조건을 만족하려면 추가로 필요한 연가 또는 기타 근무 일수입니다.",
      })}
    </div>
    <p>${escapeHtml(guide.message)}</p>
  `;
  if (panel.dataset.lastHtml === html) return;
  panel.innerHTML = html;
  panel.dataset.lastHtml = html;
  panel.dataset.hasGuide = "true";
}

function scheduleGuideRefresh() {
  clearTimeout(state.guideTimer);
  state.guideTimer = setTimeout(refreshGuide, 250);
}

async function refreshGuide() {
  collectSidebar();
  const requestId = ++state.guideRequestId;
  const panel = $("guidePanel");
  const hasExistingGuide = panel.dataset.hasGuide === "true";
  panel.classList.remove("hidden");
  panel.setAttribute("aria-busy", "true");
  if (!hasExistingGuide) {
    panel.innerHTML = `<div class="inline-progress"><span class="spinner"></span><span>연가 필요 일수 계산 중</span></div>`;
    delete panel.dataset.lastHtml;
  }
  try {
    const data = await requestJson("/api/leave-guide", {
      method: "POST",
      json: buildSchedulePayload(),
    });
    if (requestId !== state.guideRequestId) return;
    renderGuide(data.guide);
  } catch (error) {
    if (requestId !== state.guideRequestId) return;
    if (hasExistingGuide) {
      showTopToast(error.message, "error", 7000);
      return;
    }
    const html = `<p class="guide-error">${escapeHtml(error.message)}</p>`;
    if (panel.dataset.lastHtml !== html) {
      panel.innerHTML = html;
      panel.dataset.lastHtml = html;
      delete panel.dataset.hasGuide;
    }
  } finally {
    if (requestId === state.guideRequestId) {
      panel.removeAttribute("aria-busy");
    }
  }
}

function metric(label, value, options = {}) {
  const className = options.danger ? "metric metric-danger" : "metric";
  const tip = options.tip
    ? `<span class="help metric-help" tabindex="0" data-tip="${escapeHtml(options.tip)}">?</span>`
    : "";
  return `
    <div class="${className}">
      <div class="metric-label"><span>${escapeHtml(label)}</span>${tip}</div>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `;
}

function renderHistory() {
  const select = $("historySelect");
  select.innerHTML = state.histories
    .map((item, index) => `<option value="${index}" ${index === state.selectedHistory ? "selected" : ""}>${escapeHtml(item.label)}</option>`)
    .join("");
  select.disabled = state.histories.length === 0;
  $("rerollBtn").disabled = state.isSolving || state.histories.length === 0;
  $("repairEditBtn").disabled = state.isSolving || state.isDownloading || state.repairEditing || state.histories.length === 0;
}

function renderResult() {
  const wrap = $("resultTableWrap");
  const entry = state.histories[state.selectedHistory];
  if (!entry) {
    wrap.innerHTML = "";
    $("resultMeta").textContent = "생성된 결과가 없습니다.";
    $("downloadBtn").disabled = true;
    $("repairEditBtn").disabled = true;
    renderResultWarnings(null);
    return;
  }

  const result = state.repairEditing && state.repairDraftResult ? state.repairDraftResult : entry.result;
  const isRepairEditing = state.repairEditing && state.repairDraftResult;
  $("resultMeta").textContent = `${result.year}년 ${result.month}월 | ${result.status}`;
  $("downloadBtn").disabled = state.isDownloading || state.repairEditing;
  $("repairEditBtn").classList.toggle("hidden", state.repairEditing);
  $("repairApplyBtn").classList.toggle("hidden", !state.repairEditing);
  $("repairCancelBtn").classList.toggle("hidden", !state.repairEditing);
  $("repairEditBtn").disabled = state.isSolving || state.histories.length === 0;
  $("repairApplyBtn").disabled = state.isSolving || Object.keys(state.repairEdits).length === 0;
  renderResultWarnings(result);
  const repairChangedKeys = isRepairEditing ? new Set() : repairChangedCellSet(result);

  const head = [
    `<th class="row-index-head row-index-sticky"></th>`,
    `<th class="name-sticky">이름</th>`,
    ...range(1, result.days_in_month).map((day) => `<th class="date-head">${day}</th>`),
    `<th class="stat-cell stat-count-cell">주</th>`,
    `<th class="stat-cell stat-count-cell">야</th>`,
    `<th class="stat-cell stat-count-cell">비</th>`,
    `<th class="stat-cell stat-count-cell">휴</th>`,
    `<th class="stat-cell stat-count-cell">연</th>`,
    `<th class="stat-cell work-hours-cell">시간</th>`,
  ].join("");

  const rows = result.rows
    .map((row, rowIndex) => {
      const stats = rowStats(row.raw_days);
      return `<tr>
        <td class="row-index-cell row-index-sticky">${rowIndex + 1}</td>
        <td class="name-sticky">${escapeHtml(row.name)}</td>
        ${row.raw_days.map((raw, dayIndex) => renderResultDayCell(rowIndex, dayIndex + 1, raw, isRepairEditing, repairChangedKeys)).join("")}
        <td class="stat-cell stat-count-cell">${stats.day}</td>
        <td class="stat-cell stat-count-cell">${stats.night}</td>
        <td class="stat-cell stat-count-cell">${stats.offNight}</td>
        <td class="stat-cell stat-count-cell">${stats.off}</td>
        <td class="stat-cell stat-count-cell">${stats.leave}</td>
        <td class="stat-cell work-hours-cell">${stats.workHours}</td>
      </tr>`;
    })
    .join("");

  const dayCounts = countResultDays(result, "day");
  const nightCounts = countResultDays(result, "night");
  const dayCountRow = `<tr><td class="row-index-cell row-index-sticky"></td><td class="name-sticky stat-cell">주간</td>${dayCounts.map((value) => `<td class="stat-cell day-count-cell">${value}</td>`).join("")}<td colspan="6"></td></tr>`;
  const nightCountRow = `<tr><td class="row-index-cell row-index-sticky"></td><td class="name-sticky stat-cell">야간</td>${nightCounts.map((value) => `<td class="stat-cell day-count-cell">${value}</td>`).join("")}<td colspan="6"></td></tr>`;
  const colgroup = [
    `<col class="result-row-index-col" />`,
    `<col class="result-name-col" />`,
    ...range(1, result.days_in_month).map(() => `<col class="result-day-col" />`),
    ...range(1, 5).map(() => `<col class="result-stat-col" />`),
    `<col class="result-work-hours-col" />`,
  ].join("");

  wrap.innerHTML = `<table class="schedule-table result-table ${isRepairEditing ? "repair-editing-table" : ""}" style="--result-days: ${result.days_in_month};"><colgroup>${colgroup}</colgroup><thead><tr>${head}</tr></thead><tbody>${rows}${dayCountRow}${nightCountRow}</tbody></table>`;
}

function renderResultDayCell(rowIndex, day, raw, isRepairEditing, repairChangedKeys = new Set()) {
  const code = rawToShiftValue(raw);
  const display = displayShiftValue(raw);
  const visualCode = code || labelToCode(display);
  const shiftAttr = `data-result-shift="${escapeHtml(visualCode)}"`;
  const colorClass = state.flags.showShiftColors ? shiftClass(visualCode) : "";
  const repairChanged = repairChangedKeys.has(repairEditKey(rowIndex, day));
  const changedClass = repairChanged ? " repair-changed-cell" : "";
  const changedTitle = repairChanged ? ` title="부분 재생성으로 변경된 칸"` : "";
  if (!isRepairEditing) {
    return `<td ${shiftAttr}${changedTitle} class="result-day${changedClass} ${colorClass}">${escapeHtml(display)}</td>`;
  }

  const locked = isResultCellInputFixed(rowIndex, day);
  const outsidePeriod = isOutsideWorkPeriod(state.workers[rowIndex] || {}, day);
  if (locked || outsidePeriod || display === "") {
    const title = locked ? "입력표에서 고정한 근무" : outsidePeriod ? "근무 기간 밖" : "";
    return `<td ${shiftAttr} class="result-day repair-locked-cell ${colorClass}" title="${escapeHtml(title)}">${escapeHtml(display)}</td>`;
  }

  const options = repairShiftOptions(rowIndex)
    .filter((option) => option.value)
    .map((option) => `<option value="${escapeHtml(option.value)}" ${option.value === code ? "selected" : ""}>${escapeHtml(option.label)}</option>`)
    .join("");
  const edited = state.repairEdits[repairEditKey(rowIndex, day)] ? " repair-edited-cell" : "";
  return `<td ${shiftAttr} class="result-day repair-edit-cell${edited} ${colorClass}">
    <select data-repair-row="${rowIndex}" data-repair-day="${day}" class="compact-select ${colorClass}">${options}</select>
  </td>`;
}

function repairChangedCellSet(result) {
  const keys = new Set();
  for (const cell of result?.repair_changed_cells || []) {
    const rowIndex = Number(cell.worker_index);
    const day = Number(cell.day);
    if (Number.isInteger(rowIndex) && Number.isInteger(day)) {
      keys.add(repairEditKey(rowIndex, day));
    }
  }
  return keys;
}

function updateRenderedShiftColors() {
  const cells = $("resultTableWrap").querySelectorAll("[data-result-shift]");
  cells.forEach((cell) => {
    applyShiftColorClass(cell, cell.dataset.resultShift || "");
    const select = cell.querySelector("select");
    if (select) applyShiftColorClass(select, cell.dataset.resultShift || "");
  });
}

function applyShiftColorClass(element, code) {
  element.classList.remove("shift-day", "shift-night", "shift-off-night", "shift-off", "shift-leave", "shift-custom");
  if (!state.flags.showShiftColors) return;
  const colorClass = shiftClass(code);
  if (colorClass) element.classList.add(colorClass);
}

function startRepairEdit() {
  const entry = state.histories[state.selectedHistory];
  if (!entry || state.repairEditing) return;
  state.repairEditing = true;
  state.repairBaseResult = deepClone(entry.result);
  state.repairDraftResult = deepClone(entry.result);
  state.repairEdits = {};
  renderResult();
  setStatus("부분 편집 중");
}

function cancelRepairEdit(shouldRender = true) {
  state.repairEditing = false;
  state.repairBaseResult = null;
  state.repairDraftResult = null;
  state.repairEdits = {};
  if (shouldRender) {
    renderHistory();
    renderResult();
    setStatus("부분 편집 취소");
  }
}

async function applyRepairEdit() {
  if (state.isSolving) return;
  if (!state.repairEditing || !state.repairBaseResult || !state.repairDraftResult) return;
  const edits = Object.values(state.repairEdits);
  if (edits.length === 0) {
    showTopToast("부분 편집된 칸이 없습니다.", "warning", 3500);
    return;
  }

  collectSidebar();
  const payload = {
    request: buildSchedulePayload(),
    result: state.repairBaseResult,
    edits,
  };

  const controller = new AbortController();
  state.solveAbortController = controller;
  state.activeSolveTaskId = "resultSolve";
  setSolveBusy(true);
  startProgressTask("resultSolve", "부분 재생성 중", 42000);
  setStatus("부분 재생성 중");
  clearMessage();

  try {
    const data = await requestJson("/api/repair", {
      method: "POST",
      json: payload,
      signal: controller.signal,
    });
    if (controller.signal.aborted) return;
    if (!data.success) throw new Error(data.error || "부분 재생성에 실패했습니다.");
    completeProgressTask("resultSolve", "부분 재생성 완료");
    const changedCount = Number(data.result.repair_changed_count || 0);
    state.histories.unshift({
      label: `${new Date().toLocaleTimeString()} 부분 재생성`,
      result: data.result,
    });
    state.selectedHistory = 0;
    cancelRepairEdit(false);
    markDirty();
    renderHistory();
    renderResult();
    showTopToast(`부분 재생성 완료. 총 ${changedCount}칸이 변경되었습니다.`, "success", 5000);
    setStatus("부분 재생성 완료");
  } catch (error) {
    if (isAbortError(error)) {
      cancelProgressTask("resultSolve", "부분 재생성 취소");
      setStatus("부분 재생성 취소");
      return;
    }
    failProgressTask("resultSolve", "부분 재생성 실패");
    showTopToast(error.message, "error", 7000);
    setStatus("부분 재생성 실패");
  } finally {
    if (state.solveAbortController === controller) {
      state.solveAbortController = null;
      state.activeSolveTaskId = null;
      setSolveBusy(false);
    }
  }
}

function handleResultTableChange(event) {
  if (!state.repairEditing || !state.repairDraftResult || !state.repairBaseResult) return;
  const target = event.target;
  if (!(target instanceof HTMLSelectElement) || !target.dataset.repairDay) return;

  const rowIndex = Number(target.dataset.repairRow);
  const day = Number(target.dataset.repairDay);
  if (!Number.isInteger(rowIndex) || !Number.isInteger(day)) return;
  const raw = shiftValueToRaw(target.value);
  state.repairDraftResult.rows[rowIndex].raw_days[day - 1] = raw;
  state.repairDraftResult.rows[rowIndex].days[day - 1] = displayShiftValue(raw);

  const baseRaw = normalizeRawForCompare(state.repairBaseResult.rows[rowIndex].raw_days[day - 1]);
  const nextRaw = normalizeRawForCompare(raw);
  const key = repairEditKey(rowIndex, day);
  if (baseRaw === nextRaw) {
    delete state.repairEdits[key];
  } else {
    state.repairEdits[key] = {
      worker_index: rowIndex,
      day,
      shift: target.value,
    };
  }
  renderResult();
}

function isResultCellInputFixed(rowIndex, day) {
  const worker = state.workers[rowIndex];
  if (!worker) return true;
  const fixed = worker.fixed_shifts || {};
  return Boolean(fixed[day] || fixed[String(day)]);
}

function repairEditKey(rowIndex, day) {
  return `${rowIndex}:${day}`;
}

function rawToShiftValue(raw) {
  if (typeof raw === "number") return RAW_TO_CODE[raw] || "";
  const token = String(raw || "").trim();
  if (!token) return "";
  if (SHIFT_LABELS[token]) return token;
  for (const [code, label] of Object.entries(SHIFT_LABELS)) {
    if (label === token) return code;
  }
  return token;
}

function shiftValueToRaw(value) {
  if (value in SHIFT_LABELS) {
    const entry = Object.entries(RAW_TO_CODE).find(([, code]) => code === value);
    return entry ? Number(entry[0]) : value;
  }
  return value;
}

function displayShiftValue(raw) {
  const code = rawToShiftValue(raw);
  return SHIFT_LABELS[code] ?? code;
}

function normalizeRawForCompare(raw) {
  return JSON.stringify(shiftValueToRaw(rawToShiftValue(raw)));
}

function countResultDays(result, targetCode) {
  const counts = Array.from({ length: result.days_in_month }, () => 0);
  for (const row of result.rows) {
    row.raw_days.forEach((raw, index) => {
      if (rawToShiftValue(raw) === targetCode) counts[index] += 1;
    });
  }
  return counts;
}

function renderResultWarnings(result) {
  const box = $("resultWarnings");
  if (!box) return;

  const streaks = result?.long_off_streaks || [];
  if (streaks.length === 0) {
    box.classList.add("hidden");
    box.innerHTML = "";
    return;
  }

  const fallbackText = result.long_off_streak_fallback_used
    ? "5연속 휴무 금지 조건으로는 해를 찾지 못해 fallback을 사용했습니다."
    : "fallback 없이 생성된 결과입니다. 이 경우는 재검토가 필요합니다.";
  const details = streaks
    .slice(0, 4)
    .map((item) => {
      const fixedCount = Array.isArray(item.fixed_days) ? item.fixed_days.length : 0;
      const autoCount = Array.isArray(item.auto_days) ? item.auto_days.length : 0;
      return `${escapeHtml(item.worker_name || `#${Number(item.worker_index) + 1}`)}: ${item.start_day}~${item.end_day}일 ${item.length}연속 휴/연(자동 휴 ${autoCount}, 고정 ${fixedCount})`;
    })
    .join(" · ");
  const more = streaks.length > 4 ? ` 외 ${streaks.length - 4}건` : "";

  box.innerHTML = `
    <strong>연속 휴무 경고</strong>
    <span>${escapeHtml(fallbackText)} ${details}${more}</span>
  `;
  box.classList.remove("hidden");
  console.warn("[V3] long off streaks", {
    fallbackUsed: result.long_off_streak_fallback_used,
    streaks,
  });
}

function rowStats(rawDays) {
  const stats = { day: 0, night: 0, offNight: 0, off: 0, leave: 0, workHours: 0 };
  for (const raw of rawDays) {
    const code = RAW_TO_CODE[raw] || raw;
    if (code === "day") stats.day += 1;
    if (code === "night") stats.night += 1;
    if (code === "off_night") stats.offNight += 1;
    if (code === "off") stats.off += 1;
    if (code === "leave") stats.leave += 1;
  }
  stats.workHours = (stats.day + stats.night + stats.offNight) * 8;
  return stats;
}

async function downloadExcel() {
  const entry = state.histories[state.selectedHistory];
  if (!entry) return;
  setDownloadBusy(true);
  startProgressTask("download", "엑셀 생성 중", 9000);
  setStatus("엑셀 생성 중");
  try {
    const response = await fetch("/api/export-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        template_id: state.uploadedTemplate?.file_id || null,
        result: entry.result,
        apply_shift_colors: state.flags.showShiftColors,
      }),
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      failProgressTask("download", "엑셀 생성 실패");
      showTopToast(data.error || "엑셀 다운로드에 실패했습니다.", "error", 7000);
      return;
    }
    setProgressTask("download", 48, "엑셀 파일 수신 중");
    const blob = await responseToBlobWithProgress(response, "download");
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `근무표_${entry.result.year}_${entry.result.month}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
    completeProgressTask("download", "다운로드 준비 완료");
    setStatus("엑셀 다운로드 준비 완료");
  } catch (error) {
    failProgressTask("download", "엑셀 다운로드 실패");
    showTopToast(error.message, "error", 7000);
    setStatus("엑셀 다운로드 실패");
  } finally {
    setDownloadBusy(false);
  }
}

function setUploadBusy(isBusy) {
  state.isUploading = isBusy;
  $("templateFile").disabled = isBusy;
}

function setSolveBusy(isBusy) {
  state.isSolving = isBusy;
  $("solveBtn").disabled = isBusy;
  $("rerollBtn").disabled = isBusy || state.histories.length === 0;
  $("repairEditBtn").disabled = isBusy || state.repairEditing || state.histories.length === 0;
  $("repairApplyBtn").disabled = isBusy || Object.keys(state.repairEdits).length === 0;
  $("repairCancelBtn").disabled = isBusy;
  $("solveCancelBtn").disabled = !isBusy || state.activeSolveTaskId !== "solve";
  $("resultSolveCancelBtn").disabled = !isBusy || state.activeSolveTaskId !== "resultSolve";
}

function cancelActiveSolve(showToast = true) {
  const controller = state.solveAbortController;
  const taskId = state.activeSolveTaskId;
  if (!controller || !taskId) return;

  controller.abort();
  const label = taskId === "resultSolve" ? (state.repairEditing ? "부분 재생성 취소" : "재생성 취소") : "생성 취소";
  cancelProgressTask(taskId, label);
  state.solveAbortController = null;
  state.activeSolveTaskId = null;
  setSolveBusy(false);
  setStatus(label);
  if (showToast) {
    showTopToast("요청을 취소했습니다. 이미 시작된 계산은 잠시 후 백그라운드에서 종료될 수 있습니다.", "warning", 5000);
  }
}

function isAbortError(error) {
  return error?.name === "AbortError" || /abort|aborted|취소/i.test(String(error?.message || ""));
}

function setDownloadBusy(isBusy) {
  state.isDownloading = isBusy;
  $("downloadSpinner").classList.toggle("hidden", !isBusy);
  $("downloadText").textContent = isBusy ? "엑셀 생성 중" : "엑셀 다운로드";
  $("downloadBtn").disabled = isBusy || !state.histories[state.selectedHistory];
}

function startProgressTask(taskId, label, estimateMs) {
  clearProgressTask(taskId);
  state.progressTasks[taskId] = {
    startedAt: Date.now(),
    estimateMs,
    label,
    percent: 1,
    timer: null,
    hideTimer: null,
    lastEtaSeconds: null,
    failed: false,
    cancelled: false,
  };
  const element = $(`${taskId}Progress`);
  element.classList.remove("hidden", "progress-error", "progress-cancelled");
  renderProgressTask(taskId);
  updateToastStack();
  state.progressTasks[taskId].timer = setInterval(() => {
    const task = state.progressTasks[taskId];
    if (!task) return;
    const elapsed = Date.now() - task.startedAt;
    const simulated = Math.min(91, Math.round((1 - Math.exp(-elapsed / (task.estimateMs * 0.45))) * 91));
    setProgressTask(taskId, simulated, task.label);
  }, 250);
}

function setProgressTask(taskId, percent, label) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  task.percent = Math.max(task.percent, Math.min(100, Math.round(percent)));
  if (label) task.label = label;
  renderProgressTask(taskId);
}

function completeProgressTask(taskId, label) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  if (task.timer) clearInterval(task.timer);
  task.percent = 100;
  task.label = label || task.label;
  task.failed = false;
  renderProgressTask(taskId);
  task.hideTimer = setTimeout(() => hideProgressTask(taskId), 900);
}

function failProgressTask(taskId, label) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  if (task.timer) clearInterval(task.timer);
  task.label = label || task.label;
  task.failed = true;
  $(`${taskId}Progress`).classList.add("progress-error");
  renderProgressTask(taskId);
  task.hideTimer = setTimeout(() => hideProgressTask(taskId), 1600);
}

function cancelProgressTask(taskId, label) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  if (task.timer) clearInterval(task.timer);
  task.label = label || task.label;
  task.failed = false;
  task.cancelled = true;
  $(`${taskId}Progress`).classList.add("progress-cancelled");
  renderProgressTask(taskId);
  task.hideTimer = setTimeout(() => hideProgressTask(taskId), 1200);
}

function clearProgressTask(taskId) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  if (task.timer) clearInterval(task.timer);
  if (task.hideTimer) clearTimeout(task.hideTimer);
  delete state.progressTasks[taskId];
}

function hideProgressTask(taskId) {
  const element = $(`${taskId}Progress`);
  if (element) element.classList.add("hidden");
  clearProgressTask(taskId);
  updateToastStack();
}

function renderProgressTask(taskId) {
  const task = state.progressTasks[taskId];
  if (!task) return;
  const percent = Math.min(100, Math.max(0, task.percent));
  $(`${taskId}ProgressLabel`).textContent = task.label;
  $(`${taskId}ProgressBar`).style.width = `${percent}%`;
  $(`${taskId}ProgressMeta`).textContent = progressMeta(task, percent);
}

function progressMeta(task, percent) {
  if (task.cancelled) return `${percent}% · 취소됨`;
  if (task.failed) return `${percent}% · 중단됨`;
  if (percent >= 100) return "100% · 완료";
  const elapsed = Math.max(250, Date.now() - task.startedAt);
  if (elapsed > task.estimateMs && percent >= 90) return `${percent}% · 마무리 중`;
  const etaMs = percent > 3 ? (elapsed * (100 - percent)) / percent : task.estimateMs;
  const etaSeconds = Math.max(1, Math.ceil(etaMs / 1000));
  task.lastEtaSeconds = Math.min(task.lastEtaSeconds ?? etaSeconds, etaSeconds);
  return `${percent}% · 약 ${task.lastEtaSeconds}초 남음`;
}

function scheduleFailureToastMessage(errorMessage, guide) {
  const message = String(errorMessage || "");
  const shortageHours = Number(guide?.shortage_hours || 0);
  const suggestedLeaveDays = Number(guide?.suggested_leave_days || 0);
  if ((shortageHours > 0 || suggestedLeaveDays > 0) && isGenericScheduleFailureMessage(message)) {
    return guide.message || message;
  }
  return message;
}

function isGenericScheduleFailureMessage(message) {
  return (
    message.includes("현재 조건으로는 스케줄 생성이 불가능합니다") ||
    message.includes("인원 설정(TO)") ||
    message.includes("기본 TO 설정")
  );
}

function uploadJsonWithProgress(url, form, onProgress) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", url);
    xhr.responseType = "text";
    xhr.upload.onprogress = (event) => {
      if (!event.lengthComputable) return;
      const percent = Math.min(68, Math.max(8, Math.round((event.loaded / event.total) * 68)));
      onProgress(percent, "파일 전송 중");
    };
    xhr.onload = () => {
      const data = parseJson(xhr.responseText);
      if (xhr.status >= 200 && xhr.status < 300 && data?.success !== false) {
        onProgress(94, "엑셀 분석 중");
        resolve(data);
        return;
      }
      reject(new Error(errorMessageFromData(data, "업로드에 실패했습니다.")));
    };
    xhr.onerror = () => reject(new Error("업로드 요청에 실패했습니다."));
    xhr.send(form);
  });
}

async function responseToBlobWithProgress(response, taskId) {
  const total = Number(response.headers.get("content-length") || 0);
  if (!response.body || !total) {
    setProgressTask(taskId, 86, "엑셀 파일 준비 중");
    return response.blob();
  }

  const reader = response.body.getReader();
  const chunks = [];
  let loaded = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    loaded += value.length;
    const percent = Math.min(96, 48 + Math.round((loaded / total) * 48));
    setProgressTask(taskId, percent, "엑셀 파일 수신 중");
  }

  return new Blob(chunks, {
    type: response.headers.get("content-type") || "application/octet-stream",
  });
}

function parseJson(text) {
  try {
    return JSON.parse(text || "{}");
  } catch {
    return {};
  }
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function errorMessageFromData(data, fallback) {
  if (data?.error) return String(data.error);
  if (typeof data?.detail === "string") return data.detail;
  if (Array.isArray(data?.detail) && data.detail.length > 0) {
    return data.detail
      .map((item) => item?.msg || item?.message || String(item))
      .join(" ");
  }
  if (data?.detail) return String(data.detail);
  return fallback;
}

function trimFixedShiftsToMonth() {
  for (const worker of state.workers) {
    const next = {};
    for (const [day, shift] of Object.entries(worker.fixed_shifts || {})) {
      if (Number(day) <= state.daysInMonth) next[day] = shift;
    }
    worker.fixed_shifts = next;
  }
}

function shiftClass(code) {
  if (!code) return "";
  if (code === "day") return "shift-day";
  if (code === "night") return "shift-night";
  if (code === "off_night") return "shift-off-night";
  if (code === "off") return "shift-off";
  if (code === "leave") return "shift-leave";
  return "shift-custom";
}

function labelToCode(label) {
  for (const [code, value] of Object.entries(SHIFT_LABELS)) {
    if (value === label) return code;
  }
  return label ? "custom" : "";
}

function range(start, endInclusive) {
  return Array.from({ length: endInclusive - start + 1 }, (_, index) => start + index);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function toInt(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function optionalInt(value) {
  if (value === "" || value === null || value === undefined) return null;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function nullableInt(value) {
  return optionalInt(value);
}

async function requestJson(url, options = {}) {
  const init = { ...options };
  if (options.json) {
    init.body = JSON.stringify(options.json);
    init.headers = { "Content-Type": "application/json", ...(options.headers || {}) };
    delete init.json;
  }
  const response = await fetch(url, init);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.success === false) {
    throw new Error(errorMessageFromData(data, "요청에 실패했습니다."));
  }
  return data;
}

function showMessage(text, type) {
  const area = $("messageArea");
  area.textContent = text;
  area.className = `message-area ${type}`;
  area.classList.remove("hidden");
}

function clearMessage() {
  $("messageArea").className = "message-area hidden";
  $("messageArea").textContent = "";
}

function showTopToast(text, type = "error", durationMs = 7000) {
  const normalizedText = normalizeToastText(text);
  const duplicate = state.toasts.find(
    (toast) =>
      !toast.removing &&
      toast.type === type &&
      toast.text === normalizedText &&
      Date.now() - toast.createdAt < TOAST_DUPLICATE_WINDOW_MS,
  );
  if (duplicate) {
    if (duplicate.timer) clearTimeout(duplicate.timer);
    duplicate.timer = setTimeout(() => dismissTopToast(duplicate.id), durationMs);
    return;
  }

  const stack = ensureToastStack();
  const id = state.nextToastId++;
  const toast = document.createElement("div");
  toast.className = `top-toast ${type}`;
  toast.dataset.toastId = String(id);
  toast.setAttribute("role", type === "error" ? "alert" : "status");
  toast.textContent = normalizedText;
  stack.appendChild(toast);

  const item = {
    id,
    type,
    text: normalizedText,
    createdAt: Date.now(),
    element: toast,
    timer: null,
    removing: false,
  };
  state.toasts.push(item);
  updateToastStack();

  requestAnimationFrame(() => {
    toast.classList.add("visible");
  });

  item.timer = setTimeout(() => dismissTopToast(id), durationMs);
}

function normalizeToastText(text) {
  return String(text ?? "").replace(/\s+/g, " ").trim();
}

function ensureToastStack() {
  let stack = $("topToastStack");
  if (!stack) {
    stack = document.createElement("div");
    stack.id = "topToastStack";
    stack.className = "top-toast-stack";
    stack.setAttribute("aria-live", "polite");
    document.body.appendChild(stack);
  }
  return stack;
}

function dismissTopToast(id) {
  const item = state.toasts.find((toast) => toast.id === id);
  if (!item || item.removing) return;
  if (item.timer) clearTimeout(item.timer);
  item.removing = true;
  item.element.classList.remove("visible");
  item.element.classList.add("toast-exit");
  updateToastStack();

  setTimeout(() => {
    item.element.remove();
    state.toasts = state.toasts.filter((toast) => toast.id !== id);
    updateToastStack();
  }, TOAST_EXIT_MS);
}

function updateToastStack() {
  const stack = $("topToastStack");
  if (!stack) return;
  stack.classList.toggle("with-progress", hasActiveFloatingProgress());

  let offset = 0;
  state.toasts.forEach((item) => {
    if (item.removing) return;
    item.element.style.setProperty("--toast-y", `${offset}px`);
    offset += item.element.offsetHeight + TOAST_GAP_PX;
  });
}

function hasActiveFloatingProgress() {
  return ["solve", "resultSolve"].some((taskId) => {
    const task = state.progressTasks[taskId];
    const element = $(`${taskId}Progress`);
    return Boolean(task && element && !element.classList.contains("hidden"));
  });
}

function setStatus(text) {
  $("statusLine").textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

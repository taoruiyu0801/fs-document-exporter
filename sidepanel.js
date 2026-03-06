import { MESSAGE_TYPES, STATUS } from "./utils/constants.js";

const PROGRESS_HISTORY_KEY = "progressHistory";

const elements = {
  pageTitle: document.getElementById("pageTitle"),
  pageType: document.getElementById("pageType"),
  pageUrl: document.getElementById("pageUrl"),
  pathInput: document.getElementById("pathInput"),
  exportMarkdown: document.getElementById("exportMarkdown"),
  exportHtml: document.getElementById("exportHtml"),
  exportDocx: document.getElementById("exportDocx"),
  exportPdf: document.getElementById("exportPdf"),
  exportBitableJson: document.getElementById("exportBitableJson"),
  exportBitableCsv: document.getElementById("exportBitableCsv"),
  exportBitableXlsx: document.getElementById("exportBitableXlsx"),
  clearProgressBtn: document.getElementById("clearProgressBtn"),
  logList: document.getElementById("logList"),
  resultBox: document.getElementById("resultBox"),
  statusDot: document.getElementById("statusDot")
};

const state = {
  pageInfo: null,
  logs: [],
  isExporting: false,
  exportGuardTimer: null,
  activeTaskId: ""
};

function nowTime() {
  const date = new Date();
  return date.toLocaleTimeString("zh-CN", { hour12: false });
}

function setStatusDot(status) {
  const dot = elements.statusDot;
  dot.className = "dot";
  if (status) {
    dot.classList.add(status);
  }
}

function applyExportButtonState() {
  const hasPage = Boolean(state.pageInfo);
  const isBitable = state.pageInfo?.pageType === "bitable";
  const disabledByExport = state.isExporting;

  elements.exportMarkdown.disabled = disabledByExport || !hasPage;
  elements.exportHtml.disabled = disabledByExport || !hasPage;
  elements.exportDocx.disabled = disabledByExport || !hasPage;
  elements.exportPdf.disabled = disabledByExport || !hasPage;
  elements.exportBitableJson.disabled = disabledByExport || !hasPage || !isBitable;
  elements.exportBitableCsv.disabled = disabledByExport || !hasPage || !isBitable;
  elements.exportBitableXlsx.disabled = disabledByExport || !hasPage || !isBitable;
}

function setExporting(flag) {
  state.isExporting = Boolean(flag);
  if (!state.isExporting) {
    state.activeTaskId = "";
  }
  if (state.exportGuardTimer) {
    clearTimeout(state.exportGuardTimer);
    state.exportGuardTimer = null;
  }
  if (state.isExporting) {
    state.exportGuardTimer = setTimeout(() => {
      state.isExporting = false;
      applyExportButtonState();
      pushLog({
        title: "导出任务",
        message: "导出超时保护触发，已自动解除按钮锁定",
        status: STATUS.WARNING,
        source: "sidepanel"
      });
    }, 2 * 60 * 1000);
  }
  applyExportButtonState();
}

function parsePathParts(input) {
  return String(input || "")
    .split(/[\\/]+/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function pushLog({ title, message, status = STATUS.RUNNING, source = "runtime", taskId = "" }) {
  state.logs.unshift({
    time: nowTime(),
    title: title || "日志",
    message: message || "",
    status,
    source,
    taskId: String(taskId || "")
  });
  state.logs = state.logs.slice(0, 120);
  renderLogs();
  setStatusDot(status);
}

function replaceLogsFromStorage(entries) {
  const list = Array.isArray(entries) ? entries : [];
  state.logs = list.slice(0, 120).map((entry) => ({
    time: new Date(entry.timestamp || Date.now()).toLocaleTimeString("zh-CN", {
      hour12: false
    }),
    title: entry.title || "进度",
    message: entry.message || "",
    status: entry.status || STATUS.RUNNING,
    source: entry.source || "runtime",
    taskId: String(entry.taskId || "")
  }));
  renderLogs();
  if (state.logs[0]) {
    setStatusDot(state.logs[0].status);
  }
}

function renderLogs() {
  elements.logList.innerHTML = "";
  if (!state.logs.length) {
    const empty = document.createElement("li");
    empty.className = "log-item";
    empty.textContent = "暂无日志。";
    elements.logList.appendChild(empty);
    return;
  }

  for (const log of state.logs) {
    const item = document.createElement("li");
    item.className = `log-item ${log.status || ""}`;
    item.innerHTML = `
      <div class="log-head">
        <span>${log.time}</span>
        <span>${log.source}</span>
      </div>
      <div><strong>${log.title}</strong></div>
      <div>${log.message}${log.taskId ? ` [task:${log.taskId.slice(0, 8)}]` : ""}</div>
    `;
    elements.logList.appendChild(item);
  }
}

async function clearProgressLogs() {
  state.logs = [];
  renderLogs();
  setStatusDot("");
  try {
    await chrome.storage.local.set({ [PROGRESS_HISTORY_KEY]: [] });
  } catch (error) {
    pushLog({
      title: "日志清理",
      message: error instanceof Error ? error.message : String(error),
      status: STATUS.WARNING,
      source: "sidepanel"
    });
  }
}

function renderPageInfo() {
  const info = state.pageInfo;
  if (!info) {
    elements.pageTitle.textContent = "-";
    elements.pageType.textContent = "-";
    elements.pageUrl.textContent = "-";
    elements.pageUrl.href = "#";
    applyExportButtonState();
    return;
  }
  elements.pageTitle.textContent = info.title || "(无标题)";
  elements.pageType.textContent = info.pageType || "unknown";
  elements.pageUrl.textContent = info.url || "-";
  elements.pageUrl.href = info.url || "#";
  applyExportButtonState();
}

function renderResult(message) {
  const fileName = message.filename || "(未知文件)";
  const downloadId =
    typeof message.downloadId === "number" ? `，下载ID: ${message.downloadId}` : "";
  const exportTypeRaw = String(message.exportType || "").toUpperCase();
  let exportTypeLabel = exportTypeRaw;
  if (exportTypeRaw === "MARKDOWN") exportTypeLabel = "Markdown";
  if (exportTypeRaw === "HTML") exportTypeLabel = "HTML";
  if (exportTypeRaw === "WORD" || exportTypeRaw === "DOCX") exportTypeLabel = "DOCX";
  if (exportTypeRaw === "PDF") exportTypeLabel = "PDF";
  if (exportTypeRaw === "BITABLE_JSON") exportTypeLabel = "Bitable JSON";
  if (exportTypeRaw === "BITABLE_CSV") exportTypeLabel = "Bitable CSV";
  if (exportTypeRaw === "BITABLE_XLSX") exportTypeLabel = "Bitable XLSX";
  if (!exportTypeLabel) {
    exportTypeLabel =
      message.type === MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE ? "Bitable" : "导出任务";
  }
  elements.resultBox.textContent = `${exportTypeLabel} 导出完成：${fileName}${downloadId}`;
}

async function refreshPageInfo() {
  const response = await chrome.runtime.sendMessage({
    type: MESSAGE_TYPES.REQUEST_ACTIVE_TAB_INFO
  });
  if (!response?.ok) {
    pushLog({
      title: "页面识别",
      message: response?.error || "读取页面信息失败",
      status: STATUS.WARNING,
      source: "sidepanel"
    });
    return;
  }
  state.pageInfo = response.info;
  renderPageInfo();
}

async function loadProgressHistory() {
  try {
    const data = await chrome.storage.local.get(PROGRESS_HISTORY_KEY);
    replaceLogsFromStorage(data[PROGRESS_HISTORY_KEY]);
  } catch (error) {
    pushLog({
      title: "日志加载",
      message: error instanceof Error ? error.message : String(error),
      status: STATUS.WARNING,
      source: "sidepanel"
    });
  }
}

function buildExportOptions(exportType) {
  const pathParts = parsePathParts(elements.pathInput.value);
  const finalPathParts = pathParts.length ? pathParts : ["download"];
  if (exportType === "markdown") {
    return { pathParts: finalPathParts, zipImages: true };
  }
  if (exportType === "html") {
    return { pathParts: finalPathParts };
  }
  if (exportType === "docx" || exportType === "word") {
    return { pathParts: finalPathParts };
  }
  if (exportType === "pdf") {
    return { pathParts: finalPathParts };
  }
  if (
    exportType === "bitable-json" ||
    exportType === "bitable-csv" ||
    exportType === "bitable-xlsx"
  ) {
    return { pathParts: finalPathParts };
  }
  return { pathParts: finalPathParts };
}

async function startExport(exportType) {
  if (state.isExporting) {
    pushLog({
      title: "导出请求",
      message: "已有导出任务进行中，请等待完成",
      status: STATUS.WARNING,
      source: "sidepanel"
    });
    return;
  }

  const options = buildExportOptions(exportType);
  pushLog({
    title: "导出请求",
    message: `请求导出 ${exportType.toUpperCase()}`,
    status: STATUS.RUNNING,
    source: "sidepanel"
  });
  setExporting(true);
  const response = await chrome.runtime.sendMessage({
    type: MESSAGE_TYPES.START_EXPORT,
    exportType,
    options
  });
  if (!response?.ok) {
    setExporting(false);
    pushLog({
      title: "导出请求",
      message: response?.error || "导出请求失败",
      status: STATUS.ERROR,
      source: "background",
      taskId: response?.taskId || ""
    });
    return;
  }
  state.activeTaskId = String(response.taskId || "");
  if (state.activeTaskId) {
    pushLog({
      title: "导出任务",
      message: `任务已创建：${state.activeTaskId.slice(0, 8)}`,
      status: STATUS.RUNNING,
      source: "background",
      taskId: state.activeTaskId
    });
  }
}

function bindEvents() {
  elements.exportMarkdown.addEventListener("click", () => startExport("markdown"));
  elements.exportHtml.addEventListener("click", () => startExport("html"));
  elements.exportDocx.addEventListener("click", () => startExport("docx"));
  elements.exportPdf.addEventListener("click", () => startExport("pdf"));
  elements.exportBitableJson.addEventListener("click", () => startExport("bitable-json"));
  elements.exportBitableCsv.addEventListener("click", () => startExport("bitable-csv"));
  elements.exportBitableXlsx.addEventListener("click", () => startExport("bitable-xlsx"));
  elements.clearProgressBtn.addEventListener("click", () => {
    clearProgressLogs().catch(() => {});
  });

  chrome.runtime.onMessage.addListener((message) => {
    if (!message) return;

    if (message.type === MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE) {
      const data = message.data || {};
      const messageTaskId = String(data.taskId || "");
      const shouldAffectActiveTask =
        !state.activeTaskId || (messageTaskId && messageTaskId === state.activeTaskId);
      if (data.status === STATUS.ERROR && shouldAffectActiveTask) {
        setExporting(false);
      }
      if (String(data.taskState || "").toLowerCase() === "done" && shouldAffectActiveTask) {
        setExporting(false);
      }
      pushLog({
        title: data.title || "进度",
        message: data.message || "",
        status: data.status || STATUS.RUNNING,
        source: message.source || "runtime",
        taskId: messageTaskId
      });
      return;
    }

    if (
      message.type === MESSAGE_TYPES.EXPORT_COMPLETE ||
      message.type === MESSAGE_TYPES.MARKDOWN_EXPORT_COMPLETE ||
      message.type === MESSAGE_TYPES.HTML_EXPORT_COMPLETE ||
      message.type === MESSAGE_TYPES.WORD_EXPORT_COMPLETE ||
      message.type === MESSAGE_TYPES.PDF_EXPORT_COMPLETE ||
      message.type === MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE
    ) {
      const messageTaskId = String(message.taskId || "");
      if (state.activeTaskId) {
        if (!messageTaskId || state.activeTaskId !== messageTaskId) {
          return;
        }
      }
      setExporting(false);
      renderResult(message);
    }
  });

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local") return;
    if (!changes[PROGRESS_HISTORY_KEY]) return;
    replaceLogsFromStorage(changes[PROGRESS_HISTORY_KEY].newValue);
  });

  chrome.tabs.onActivated.addListener(() => {
    refreshPageInfo().catch(() => {});
  });
  chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (changeInfo.status === "complete" && tab?.active) {
      refreshPageInfo().catch(() => {});
    }
  });
}

async function init() {
  bindEvents();
  renderLogs();
  renderPageInfo();
  applyExportButtonState();
  await loadProgressHistory();
  await refreshPageInfo();
}

init().catch((error) => {
  pushLog({
    title: "初始化",
    message: error instanceof Error ? error.message : String(error),
    status: STATUS.ERROR,
    source: "sidepanel"
  });
});

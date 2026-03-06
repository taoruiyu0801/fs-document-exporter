import {
  IDB_META,
  LARGE_DOCUMENT_THRESHOLD,
  MESSAGE_TYPES,
  STATUS
} from "./utils/constants.js";
import { addOnChunkedMessageListener } from "./utils/chunked-messaging.js";
import {
  getImageDataMap,
  idbDelete,
  idbPut,
  openDatabase,
  removeImagesByTokens
} from "./utils/idb.js";
import {
  buildHtmlExport,
  composeHtmlForPreview
} from "./utils/exporters/html-exporter.js";
import { buildMarkdownExport } from "./utils/exporters/markdown-exporter.js";
import {
  buildBitableCsvExport,
  buildBitableJsonExport,
  buildBitableXlsxExport
} from "./utils/exporters/bitable-exporter.js";
import { buildWordExport } from "./utils/exporters/word-exporter.js";
import {
  blobToDataUrl,
  buildExportBaseName,
  buildDownloadPath,
  dataUrlToUint8Array,
  ensureExtension
} from "./utils/filename.js";
import { detectPageType } from "./utils/dom-extractor.js";

const pdfTaskLock = new Set();
const directDownloadDedupeMap = new Map();
const injectedExportLockMap = new Map();
const pdfPreviewSourceTabMap = new Map();
const PROGRESS_HISTORY_KEY = "progressHistory";
const PROGRESS_HISTORY_MAX = 200;
const DIRECT_DOWNLOAD_DEDUPE_TTL = 6000;
const INJECTED_EXPORT_LOCK_TTL = 3 * 60 * 1000;
const EXPORT_TASK_TTL = 20 * 60 * 1000;
const IMAGE_PROXY_FETCH_TIMEOUT_MS = 20000;
const INJECTED_EXPORT_TYPES = new Set(["markdown", "html", "pdf", "docx"]);
const KNOWN_FEISHU_SUFFIXES = [
  "feishu.cn",
  "feishu.net",
  "larksuite.com",
  "larkoffice.com",
  "larkenterprise.com",
  "feishu-pre.net"
];
const EXPORT_TASK_STATE = {
  PENDING: "pending",
  RUNNING: "running",
  COLLECTING: "collecting",
  ASSEMBLING: "assembling",
  DOWNLOADING: "downloading",
  DONE: "done",
  ERROR: "error",
  CANCELLED: "cancelled"
};
const exportTaskMap = new Map();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function emitMessage(message) {
  chrome.runtime.sendMessage(message).catch(() => {});
}

function emitMessageToTab(tabId, message) {
  if (!tabId) return;
  chrome.tabs.sendMessage(tabId, message).catch(() => {});
}

function cleanupExportTasks() {
  const now = Date.now();
  for (const [taskId, task] of exportTaskMap) {
    const endTime = task.endedAt || task.updatedAt || task.createdAt || 0;
    if (now - endTime > EXPORT_TASK_TTL) {
      exportTaskMap.delete(taskId);
    }
  }
}

function touchExportTask(taskId, patch = {}) {
  if (!taskId) return null;
  cleanupExportTasks();
  const now = Date.now();
  const current = exportTaskMap.get(taskId) || {
    taskId,
    state: EXPORT_TASK_STATE.PENDING,
    createdAt: now
  };
  const merged = {
    ...current,
    ...patch,
    taskId,
    updatedAt: now
  };
  if (patch.state === EXPORT_TASK_STATE.DONE || patch.state === EXPORT_TASK_STATE.ERROR) {
    merged.endedAt = now;
  }
  exportTaskMap.set(taskId, merged);
  return merged;
}

function createExportTask({ taskId, tabId, exportType }) {
  return touchExportTask(taskId, {
    tabId: tabId || null,
    exportType: normalizeExportTypeLower(exportType),
    state: EXPORT_TASK_STATE.PENDING
  });
}

function normalizeTaskState(input) {
  const state = String(input || "").toLowerCase();
  if (Object.values(EXPORT_TASK_STATE).includes(state)) {
    return state;
  }
  return EXPORT_TASK_STATE.RUNNING;
}

function inferTaskStateFromStatus(status, fallback = EXPORT_TASK_STATE.RUNNING) {
  if (status === STATUS.SUCCESS) return EXPORT_TASK_STATE.DONE;
  if (status === STATUS.ERROR) return EXPORT_TASK_STATE.ERROR;
  return fallback;
}

function attachTaskError(error, taskId) {
  if (error instanceof Error) {
    error.taskId = taskId;
    return error;
  }
  const wrapped = new Error(String(error));
  wrapped.taskId = taskId;
  return wrapped;
}

async function appendProgressHistory(entry) {
  try {
    const current = await chrome.storage.local.get(PROGRESS_HISTORY_KEY);
    const list = Array.isArray(current[PROGRESS_HISTORY_KEY])
      ? current[PROGRESS_HISTORY_KEY]
      : [];
    list.unshift(entry);
    await chrome.storage.local.set({
      [PROGRESS_HISTORY_KEY]: list.slice(0, PROGRESS_HISTORY_MAX)
    });
  } catch {
    // ignore storage failures
  }
}

function emitProgress(progress, source = "background") {
  const taskId = progress?.taskId || null;
  if (taskId) {
    const inferredState = normalizeTaskState(
      progress?.taskState ||
        inferTaskStateFromStatus(progress?.status, EXPORT_TASK_STATE.RUNNING)
    );
    touchExportTask(taskId, {
      state: inferredState,
      lastProgressTitle: progress?.title || "",
      lastProgressMessage: progress?.message || ""
    });
  }
  const message = {
    type: MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE,
    source,
    data: progress,
    timestamp: Date.now(),
    __rebroadcast: true
  };
  emitMessage(message);
  appendProgressHistory({
    source,
    timestamp: message.timestamp,
    title: progress?.title || "进度",
    message: progress?.message || "",
    status: progress?.status || STATUS.RUNNING,
    resultLink: progress?.resultLink || "",
    downloadId: progress?.downloadId ?? null,
    taskId: progress?.taskId || "",
    taskState: progress?.taskState || ""
  });
}

function normalizeExportTypeLower(value) {
  return String(value || "").toLowerCase();
}

function isBitableAssembleExportType(value) {
  const normalized = normalizeExportTypeLower(value);
  return (
    normalized === "bitable-json" ||
    normalized === "bitable-csv" ||
    normalized === "bitable-xlsx"
  );
}

function buildInjectedExportLockKey(tabId, exportType) {
  return `${tabId}:${normalizeExportTypeLower(exportType)}`;
}

function cleanupInjectedExportLocks() {
  const now = Date.now();
  for (const [key, ts] of injectedExportLockMap) {
    if (now - ts > INJECTED_EXPORT_LOCK_TTL) {
      injectedExportLockMap.delete(key);
    }
  }
}

function acquireInjectedExportLock(tabId, exportType) {
  if (!tabId || !INJECTED_EXPORT_TYPES.has(normalizeExportTypeLower(exportType))) {
    return true;
  }
  cleanupInjectedExportLocks();
  const key = buildInjectedExportLockKey(tabId, exportType);
  if (injectedExportLockMap.has(key)) {
    return false;
  }
  injectedExportLockMap.set(key, Date.now());
  return true;
}

function releaseInjectedExportLock(tabId, exportType) {
  if (!tabId || !INJECTED_EXPORT_TYPES.has(normalizeExportTypeLower(exportType))) {
    return;
  }
  const key = buildInjectedExportLockKey(tabId, exportType);
  injectedExportLockMap.delete(key);
}

function buildDirectDownloadDedupeKey(exportType, data = {}, filenameForPath = "") {
  const url = String(data.url || "");
  const urlSignature =
    url.length > 2048 ? `${url.slice(0, 1024)}::${url.length}` : url;
  const pathSig = String(filenameForPath || data.filename || "");
  return `${String(exportType || "FILE").toUpperCase()}|${pathSig}|${urlSignature}`;
}

function shouldSkipDirectDownloadByDedupe(dedupeKey) {
  const now = Date.now();
  for (const [key, ts] of directDownloadDedupeMap) {
    if (now - ts > DIRECT_DOWNLOAD_DEDUPE_TTL) {
      directDownloadDedupeMap.delete(key);
    }
  }
  const previous = directDownloadDedupeMap.get(dedupeKey);
  if (previous && now - previous <= DIRECT_DOWNLOAD_DEDUPE_TTL) {
    return true;
  }
  directDownloadDedupeMap.set(dedupeKey, now);
  return false;
}

function isSupportedFeishuUrl(url) {
  return /https:\/\/[^/]+\.(feishu\.cn|feishu\.net|larksuite\.com|larkoffice\.com|larkenterprise\.com|feishu-pre\.net)\//i.test(
    String(url || "")
  );
}

async function getActiveTabInfo() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) {
    throw new Error("未找到当前激活标签页");
  }
  const info = {
    tabId: tab.id,
    title: tab.title || "",
    url: tab.url || "",
    pageType: detectPageType(tab.url || "")
  };

  // Prefer runtime page detection when content script is available.
  if (info.tabId && isSupportedFeishuUrl(info.url)) {
    try {
      await ensureContentScriptReady(info.tabId);
      const response = await chrome.tabs.sendMessage(info.tabId, {
        type: MESSAGE_TYPES.REQUEST_PAGE_INFO,
        requestId: crypto.randomUUID()
      });
      if (response?.ok && response.info) {
        const runtimeInfo = response.info;
        if (runtimeInfo.title) {
          info.title = runtimeInfo.title;
        }
        if (runtimeInfo.url) {
          info.url = runtimeInfo.url;
        }
        if (runtimeInfo.pageType && runtimeInfo.pageType !== "unknown") {
          info.pageType = runtimeInfo.pageType;
        }
      }
    } catch {
      // fallback to tab URL based detection
    }
  }

  return info;
}

async function ensureContentScriptReady(tabId) {
  try {
    const response = await chrome.tabs.sendMessage(tabId, {
      type: MESSAGE_TYPES.REQUEST_PAGE_INFO,
      requestId: crypto.randomUUID()
    });
    if (response?.ok) {
      return;
    }
  } catch {
    // ignore and inject below
  }

  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content-scripts/content.js"]
  });
  await sleep(80);
}

async function ensureContentScriptReadyForTab(tabId, tabUrl) {
  if (!tabId || !isSupportedFeishuUrl(tabUrl || "")) {
    return;
  }
  try {
    await ensureContentScriptReady(tabId);
  } catch {
    // silent pre-injection path
  }
}

async function downloadByUrl(url, { filename, pathParts = [] }) {
  const finalPath = buildDownloadPath({ filename, pathParts });
  let downloadId;
  try {
    downloadId = await chrome.downloads.download({
      url,
      filename: finalPath,
      saveAs: false
    });
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    const invalidFilename = /invalid filename/i.test(msg);
    if (!invalidFilename) {
      throw new Error(`下载失败(${msg})，filename=${finalPath}`);
    }

    const extMatch = /\.([a-z0-9]+)$/i.exec(finalPath);
    const ext = extMatch ? extMatch[1].toLowerCase() : "bin";
    const stamp = new Date()
      .toISOString()
      .replace(/[-:]/g, "")
      .replace(/\..+/, "")
      .replace("T", "_");
    const fallbackFilename = `feishu-export-${stamp}.${ext}`;
    const fallbackPath = buildDownloadPath({
      filename: fallbackFilename,
      pathParts: []
    });

    emitProgress({
      title: "导出任务",
      message: `文件名不合法，自动回退为 ${fallbackFilename}`,
      status: STATUS.WARNING
    });

    try {
      downloadId = await chrome.downloads.download({
        url,
        filename: fallbackPath,
        saveAs: false
      });
      return {
        downloadId,
        filename: fallbackPath,
        resultLink: fallbackPath
      };
    } catch (retryError) {
      const retryMsg =
        retryError instanceof Error ? retryError.message : String(retryError);
      throw new Error(
        `下载失败(${msg})，重试失败(${retryMsg})，filename=${finalPath}`
      );
    }
  }
  return {
    downloadId,
    filename: finalPath,
    resultLink: finalPath
  };
}

async function downloadBlob(blob, { filename, pathParts = [] }) {
  const dataUrl = await blobToDataUrl(blob);
  return downloadByUrl(dataUrl, { filename, pathParts });
}

function callbackToPromise(executor) {
  return new Promise((resolve, reject) => {
    executor((result) => {
      const lastError = chrome.runtime.lastError;
      if (lastError) {
        reject(new Error(lastError.message));
        return;
      }
      resolve(result);
    });
  });
}

async function debuggerAttach(target) {
  await callbackToPromise((done) => chrome.debugger.attach(target, "1.3", done));
}

async function debuggerDetach(target) {
  await callbackToPromise((done) => chrome.debugger.detach(target, done));
}

async function debuggerSendCommand(target, method, params = {}) {
  return callbackToPromise((done) =>
    chrome.debugger.sendCommand(target, method, params, done)
  );
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function isUuidLikeToken(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    String(value || "")
  );
}

function isLikelyDriveToken(token) {
  const value = String(token || "").trim();
  if (!/^[a-zA-Z0-9_-]{16,}$/.test(value)) return false;
  if (isUuidLikeToken(value)) return false;
  if (!/[a-zA-Z]/.test(value)) return false;
  return true;
}

function resolveFeishuRootDomain(hostname) {
  const host = String(hostname || "").toLowerCase();
  for (const suffix of KNOWN_FEISHU_SUFFIXES) {
    if (host === suffix || host.endsWith(`.${suffix}`)) {
      return suffix;
    }
  }
  return "";
}

function normalizeFetchCandidateUrl(rawUrl, baseUrl = "") {
  const value = String(rawUrl || "").trim();
  if (!value) return "";
  if (/^data:/i.test(value)) return value;
  if (/^blob:/i.test(value)) return "";
  if (/^https?:\/\//i.test(value)) return value;
  if (!baseUrl) return "";
  try {
    const normalized = new URL(value, baseUrl).href;
    return /^https?:\/\//i.test(normalized) ? normalized : "";
  } catch {
    return "";
  }
}

function buildTokenPreviewCandidates(token, currentUrl = "") {
  if (!isLikelyDriveToken(token)) return [];
  let currentHost = "";
  let currentOrigin = "";
  try {
    const parsed = new URL(String(currentUrl || ""));
    currentHost = parsed.hostname.toLowerCase();
    currentOrigin = parsed.origin;
  } catch {
    // ignore parse failure
  }
  const rootDomain = resolveFeishuRootDomain(currentHost);
  const hosts = new Set();
  if (currentHost) hosts.add(currentHost);
  if (rootDomain) {
    hosts.add(`internal-api-drive-stream.${rootDomain}`);
    hosts.add(`drive-stream.${rootDomain}`);
  } else {
    hosts.add("internal-api-drive-stream.feishu.cn");
    hosts.add("drive-stream.feishu.cn");
  }
  const urls = [];
  const add = (value) => {
    const normalized = String(value || "").trim();
    if (normalized) urls.push(normalized);
  };

  if (currentOrigin) {
    add(
      `${currentOrigin}/space/api/box/stream/download/preview/${token}?preview_type=16&mount_point=docx_image`
    );
    add(
      `${currentOrigin}/space/api/box/stream/download/preview/${token}?preview_type=16&mount_point=docx_file`
    );
    add(`${currentOrigin}/space/api/box/stream/download/preview/${token}?preview_type=16`);
    add(`${currentOrigin}/space/api/box/stream/download/preview/${token}`);
  }

  for (const host of hosts) {
    add(
      `https://${host}/space/api/box/stream/download/preview/${token}?preview_type=16&mount_point=docx_image`
    );
    add(
      `https://${host}/space/api/box/stream/download/preview/${token}?preview_type=16&mount_point=docx_file`
    );
    add(`https://${host}/space/api/box/stream/download/preview/${token}?preview_type=16`);
    add(`https://${host}/space/api/box/stream/download/preview/${token}`);
  }
  return Array.from(new Set(urls));
}

function parseDataUrlToBlob(dataUrl, fallbackMimeType = "application/octet-stream") {
  const value = String(dataUrl || "");
  const mimeMatch = /^data:([^;,]+)[;,]/i.exec(value);
  const mimeType = mimeMatch?.[1] || fallbackMimeType;
  const bytes = dataUrlToUint8Array(value);
  return new Blob([bytes], { type: mimeType });
}

async function fetchWithTimeout(url, options = {}, timeoutMs = IMAGE_PROXY_FETCH_TIMEOUT_MS) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      ...options,
      signal: controller.signal
    });
  } finally {
    clearTimeout(timer);
  }
}

async function fetchImageBlobFromCandidates({ candidates, token, currentUrl, mimeType }) {
  const urls = Array.from(
    new Set(
      [
        ...(Array.isArray(candidates) ? candidates : []),
        ...buildTokenPreviewCandidates(token, currentUrl)
      ]
        .map((value) => normalizeFetchCandidateUrl(value, currentUrl))
        .filter(Boolean)
    )
  ).slice(0, 16);

  if (!urls.length) {
    throw new Error("无可用图片地址");
  }

  let lastError = null;
  for (const url of urls) {
    try {
      if (/^data:/i.test(url)) {
        return {
          blob: parseDataUrlToBlob(url, mimeType),
          sourceUrl: url
        };
      }

      const response = await fetchWithTimeout(
        url,
        {
          credentials: "include",
          cache: "force-cache"
        },
        IMAGE_PROXY_FETCH_TIMEOUT_MS
      );
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const blob = await response.blob();
      if (!(blob instanceof Blob) || blob.size <= 0) {
        throw new Error("空图片数据");
      }
      return {
        blob,
        sourceUrl: url
      };
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("图片抓取失败");
}

async function printTabToPdf(tabId) {
  const target = { tabId };
  await debuggerAttach(target);
  try {
    await debuggerSendCommand(target, "Page.enable");
    await debuggerSendCommand(target, "Runtime.enable");
    await debuggerSendCommand(target, "Emulation.setEmulatedMedia", {
      media: "print"
    });
    const printed = await debuggerSendCommand(target, "Page.printToPDF", {
      printBackground: true,
      preferCSSPageSize: true,
      scale: 1
    });
    if (!printed?.data) {
      throw new Error("Page.printToPDF 未返回数据");
    }
    return printed.data;
  } finally {
    await debuggerDetach(target).catch(() => {});
  }
}

async function openPdfPreview(data) {
  const previewId = `pdf-preview-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;

  await idbPut(IDB_META.stores.previews, {
    id: previewId,
    ...data,
    updatedAt: Date.now()
  });

  const tab = await chrome.tabs.create({
    url: `${chrome.runtime.getURL("doc-preview.html")}?previewId=${encodeURIComponent(
      previewId
    )}`,
    active: true
  });
  if (!tab?.id) {
    throw new Error("无法创建 PDF 预览标签页");
  }
  return { tabId: tab.id, previewId, taskId: data?.taskId || "" };
}

async function finalizePdfExport(message, sender) {
  const tabId = sender?.tab?.id || message.tabId;
  const taskId = message?.taskId || null;
  const sourceTabId =
    message?.sourceTabId ||
    (message?.previewId ? pdfPreviewSourceTabMap.get(message.previewId) : null);
  if (!tabId) {
    throw new Error("未获取到 PDF 预览页 tabId");
  }

  const lockKey = `${tabId}:${message.previewId || "default"}`;
  if (pdfTaskLock.has(lockKey)) {
    return { ok: true, skipped: true, taskId };
  }
  pdfTaskLock.add(lockKey);

  try {
    emitProgress({
      title: "PDF 导出",
      message: "正在调用浏览器打印引擎生成 PDF",
      status: STATUS.RUNNING,
      taskId,
      taskState: EXPORT_TASK_STATE.DOWNLOADING
    });

    const base64 = await printTabToPdf(tabId);
    const bytes = base64ToUint8Array(base64);
    const blob = new Blob([bytes], { type: "application/pdf" });
    const cleanTitle = buildExportBaseName(message.title || "feishu-document");
    const downloaded = await downloadBlob(blob, {
      filename: ensureExtension(cleanTitle, "pdf"),
      pathParts: message.pathParts || []
    });

    emitProgress({
      title: "PDF 导出",
      message: "PDF 导出完成",
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });
    emitMessage({
      type: MESSAGE_TYPES.PDF_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(sourceTabId, {
      type: MESSAGE_TYPES.PDF_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "PDF",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(sourceTabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "PDF",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });

    if (message.previewId) {
      await idbDelete(IDB_META.stores.previews, message.previewId).catch(() => {});
    }
    return { ok: true, taskId, ...downloaded };
  } finally {
    pdfTaskLock.delete(lockKey);
    if (message?.previewId) {
      pdfPreviewSourceTabMap.delete(message.previewId);
    }
    if (sourceTabId) {
      releaseInjectedExportLock(sourceTabId, "pdf");
    }
  }
}

async function handleAssembleExport(data, context = {}) {
  const exportType = String(data.exportType || "").toLowerCase();
  const taskId = data.taskId || context.taskId || null;
  const title = data.title || "未命名文档";
  const currentUrl = data.currentUrl || data.url || "";
  const htmlWithPlaceholders = data.htmlWithPlaceholders || data.htmlContent || "";
  const imageTokens = data.imageTokens || [];
  const htmlBlobSize = Number(data.htmlBlobSize || new Blob([htmlWithPlaceholders]).size);
  const options = data.options || {};
  const isBitableExport = isBitableAssembleExportType(exportType);

  try {
    if (!htmlWithPlaceholders && !isBitableExport) {
      throw new Error("HTML 内容为空，无法继续导出");
    }

    const imageMap = isBitableExport ? new Map() : await getImageDataMap(currentUrl, imageTokens);

  emitProgress({
    title: isBitableExport ? "Bitable 导出" : "导出任务",
    message: `开始组装导出内容，类型：${exportType || "unknown"}`,
    status: STATUS.RUNNING,
    taskId,
    taskState: EXPORT_TASK_STATE.ASSEMBLING
  });

    if (htmlBlobSize > LARGE_DOCUMENT_THRESHOLD) {
      emitProgress({
        title: "导出任务",
        message: "检测到大文档，启用分块/渐进组装策略",
        status: STATUS.WARNING,
        taskId,
        taskState: EXPORT_TASK_STATE.ASSEMBLING
      });
    }

  if (exportType === "markdown") {
    const output = await buildMarkdownExport(
      {
        title,
        currentUrl,
        htmlWithPlaceholders
      },
      imageMap,
      {
        zipImages: true,
        bigMode: options.bigMode === true
      }
    );
    const downloaded = await downloadBlob(output.blob, {
      filename: output.filename,
      pathParts: options.pathParts || []
    });
    emitProgress({
      title: "Markdown 导出",
      message: "导出完成",
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });
    emitMessage({
      type: MESSAGE_TYPES.MARKDOWN_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.MARKDOWN_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "MARKDOWN",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "MARKDOWN",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    await removeImagesByTokens(currentUrl, imageTokens);
    return { ok: true, taskId, ...downloaded };
  }

  if (exportType === "html") {
    const output = await buildHtmlExport(
      {
        title,
        currentUrl,
        htmlWithPlaceholders
      },
      imageMap
    );
    const downloaded = await downloadBlob(output.blob, {
      filename: output.filename,
      pathParts: options.pathParts || []
    });
    emitProgress({
      title: "HTML 导出",
      message: "导出完成",
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });
    emitMessage({
      type: MESSAGE_TYPES.HTML_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.HTML_EXPORT_COMPLETE,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "HTML",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: "HTML",
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    await removeImagesByTokens(currentUrl, imageTokens);
    return { ok: true, ...downloaded };
  }

  if (exportType === "pdf") {
    try {
      const htmlContent = composeHtmlForPreview(
        {
          title,
          currentUrl,
          htmlWithPlaceholders
        },
        imageMap
      );
      const preview = await openPdfPreview({
        title,
        currentUrl,
        htmlContent,
        pathParts: options.pathParts || [],
        taskId
      });
      if (preview?.previewId && context?.tabId) {
        pdfPreviewSourceTabMap.set(preview.previewId, context.tabId);
      }
      await removeImagesByTokens(currentUrl, imageTokens);
      emitProgress({
        title: "PDF 导出",
        message: "已打开预览页，正在生成 PDF",
        status: STATUS.RUNNING,
        taskId,
        taskState: EXPORT_TASK_STATE.RUNNING
      });
      return { ok: true, pending: true, taskId, ...preview };
    } catch (error) {
      if (context?.tabId) {
        releaseInjectedExportLock(context.tabId, "pdf");
      }
      throw error;
    }
  }

  if (exportType === "docx" || exportType === "word") {
    const output = await buildWordExport(
      {
        title,
        currentUrl,
        htmlWithPlaceholders
      },
      imageMap
    );
    const downloaded = await downloadBlob(output.blob, {
      filename: output.filename,
      pathParts: options.pathParts || []
    });
    emitProgress({
      title: "Word 导出",
      message: "导出完成",
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });
    emitMessage({
      type: MESSAGE_TYPES.WORD_EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.WORD_EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    await removeImagesByTokens(currentUrl, imageTokens);
    return { ok: true, taskId, ...downloaded };
  }

  if (isBitableExport) {
    const payload = {
      title,
      bitableData: data.bitableData || {}
    };
    let output;
    if (exportType === "bitable-json") {
      output = await buildBitableJsonExport(payload);
    } else if (exportType === "bitable-csv") {
      output = await buildBitableCsvExport(payload);
    } else {
      output = await buildBitableXlsxExport(payload);
    }
    const downloaded = await downloadBlob(output.blob, {
      filename: output.filename,
      pathParts: options.pathParts || []
    });
    const displayLabel =
      output.exportType === "BITABLE_JSON"
        ? "JSON"
        : output.exportType === "BITABLE_CSV"
          ? "CSV"
          : "XLSX";
    const tableMessage =
      typeof output.tableCount === "number" ? `，共 ${output.tableCount} 张表` : "";
    emitProgress({
      title: "Bitable 导出",
      message: `${displayLabel} 导出完成${tableMessage}`,
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });
    emitMessage({
      type: MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(context.tabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType: output.exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    return { ok: true, taskId, ...downloaded };
  }

    throw new Error(`不支持的导出类型: ${exportType || "unknown"}`);
  } catch (error) {
    emitProgress({
      title: isBitableExport ? "Bitable 导出" : "导出任务",
      message: error instanceof Error ? error.message : String(error),
      status: STATUS.ERROR,
      taskId,
      taskState: EXPORT_TASK_STATE.ERROR
    });
    throw error;
  }
}

async function handleStoreImage(data) {
  if (!data?.id || !data?.token || !data?.documentUrl) {
    throw new Error("store-image 参数不完整");
  }
  let blob = null;
  let mimeType = String(data?.mimeType || "").trim() || "application/octet-stream";

  if (typeof data.dataUrl === "string" && data.dataUrl.startsWith("data:")) {
    const mimeMatch = /^data:([^;,]+)[;,]/i.exec(data.dataUrl);
    if (mimeMatch?.[1]) {
      mimeType = mimeMatch[1];
    }
    const bytes = dataUrlToUint8Array(data.dataUrl);
    blob = new Blob([bytes], { type: mimeType });
  } else if (typeof data.base64Data === "string" && data.base64Data.length > 0) {
    const bytes = base64ToUint8Array(data.base64Data);
    blob = new Blob([bytes], { type: mimeType });
  }

  if (!(blob instanceof Blob)) {
    throw new Error("store-image 缺少有效图片数据");
  }

  await idbPut(IDB_META.stores.images, {
    id: data.id,
    token: data.token,
    blob,
    mimeType: blob.type || mimeType,
    size: blob.size,
    documentUrl: data.documentUrl,
    updatedAt: Date.now()
  });
  return { ok: true };
}

async function handleStoreImageByCandidates(data) {
  if (!data?.id || !data?.token || !data?.documentUrl) {
    throw new Error("store-image-by-candidates 参数不完整");
  }
  const fetched = await fetchImageBlobFromCandidates({
    candidates: data.candidates || [],
    token: String(data.token || ""),
    currentUrl: String(data.documentUrl || ""),
    mimeType: String(data.mimeType || "").trim() || "application/octet-stream"
  });
  const blob = fetched.blob;
  if (!(blob instanceof Blob) || blob.size <= 0) {
    throw new Error("store-image-by-candidates 未获取到有效图片数据");
  }
  await idbPut(IDB_META.stores.images, {
    id: data.id,
    token: data.token,
    blob,
    mimeType: blob.type || String(data?.mimeType || "application/octet-stream"),
    size: blob.size,
    sourceUrl: fetched.sourceUrl || "",
    documentUrl: data.documentUrl,
    updatedAt: Date.now()
  });
  return { ok: true, sourceUrl: fetched.sourceUrl || "" };
}

function normalizeExportType(value) {
  return String(value || "FILE").toUpperCase();
}

function inferExtensionForDirectDownload(exportType, data = {}) {
  if (exportType === "MARKDOWN") {
    return data.isZip ? "zip" : "md";
  }
  if (exportType === "HTML") return "html";
  if (exportType === "WORD" || exportType === "DOCX") return "docx";
  if (exportType === "PDF") return "pdf";
  if (exportType === "BITABLE_JSON") return "json";
  if (exportType === "BITABLE_CSV" || exportType === "BITABLE-CSV") {
    return data.isZip ? "zip" : "csv";
  }
  if (exportType === "BITABLE_XLSX" || exportType === "BITABLE-XLSX") return "xlsx";
  return "bin";
}

function titleForExportType(exportType) {
  if (exportType === "MARKDOWN") return "Markdown 导出";
  if (exportType === "HTML") return "HTML 导出";
  if (exportType === "WORD" || exportType === "DOCX") return "Word 导出";
  if (exportType === "PDF") return "PDF 导出";
  if (exportType === "BITABLE_JSON") return "Bitable 导出";
  if (exportType === "BITABLE_CSV" || exportType === "BITABLE-CSV") return "Bitable CSV 导出";
  if (exportType === "BITABLE_XLSX" || exportType === "BITABLE-XLSX") return "Bitable XLSX 导出";
  return "导出任务";
}

function completeEventTypeForExport(exportType) {
  if (exportType === "MARKDOWN") return MESSAGE_TYPES.MARKDOWN_EXPORT_COMPLETE;
  if (exportType === "HTML") return MESSAGE_TYPES.HTML_EXPORT_COMPLETE;
  if (exportType === "WORD" || exportType === "DOCX") return MESSAGE_TYPES.WORD_EXPORT_COMPLETE;
  if (exportType === "PDF") return MESSAGE_TYPES.PDF_EXPORT_COMPLETE;
  if (exportType === "BITABLE_JSON") return MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE;
  if (exportType === "BITABLE_CSV" || exportType === "BITABLE-CSV") {
    return MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE;
  }
  if (exportType === "BITABLE_XLSX" || exportType === "BITABLE-XLSX") {
    return MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE;
  }
  return null;
}

async function handleDirectDownload(payload, context = {}) {
  const exportType = normalizeExportType(payload?.exportType);
  const data = payload?.data || {};
  const taskId = payload?.taskId || data?.taskId || context?.taskId || null;
  const url = data.url;
  const sourceTabId = context?.tabId || data?.sourceTabId || null;
  try {
    if (!url || typeof url !== "string") {
      throw new Error("directDownload 缺少 url");
    }

    const pathParts = Array.isArray(data.pathParts) ? data.pathParts : [];
    let filename = data.filename;
    if ((!filename || !String(filename).trim()) && pathParts.length === 0) {
      const ext = inferExtensionForDirectDownload(exportType, data);
      const baseName = buildExportBaseName(
        data.title || `${exportType.toLowerCase()}-export`,
        "feishu-export"
      );
      filename = ensureExtension(baseName, ext);
    }

    const pathPreview = buildDownloadPath({
      filename: filename || "",
      pathParts
    });
    const dedupeKey = buildDirectDownloadDedupeKey(exportType, data, pathPreview);
    if (shouldSkipDirectDownloadByDedupe(dedupeKey)) {
      emitProgress({
        title: titleForExportType(exportType),
        message: "检测到重复下载请求，已自动忽略",
        status: STATUS.WARNING,
        taskId,
        taskState: EXPORT_TASK_STATE.RUNNING
      });
      return { ok: true, skipped: true };
    }

    const downloaded = await downloadByUrl(url, {
      filename,
      pathParts
    });

    emitProgress({
      title: titleForExportType(exportType),
      message: "导出完成",
      status: STATUS.SUCCESS,
      downloadId: downloaded.downloadId,
      resultLink: downloaded.resultLink,
      taskId,
      taskState: EXPORT_TASK_STATE.DONE
    });

    const completeType = completeEventTypeForExport(exportType);
    if (completeType) {
      emitMessage({
        type: completeType,
        exportType,
        taskId,
        ...downloaded,
        timestamp: Date.now()
      });
      emitMessageToTab(sourceTabId, {
        type: completeType,
        exportType,
        taskId,
        ...downloaded,
        timestamp: Date.now()
      });
    }
    emitMessage({
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });
    emitMessageToTab(sourceTabId, {
      type: MESSAGE_TYPES.EXPORT_COMPLETE,
      exportType,
      taskId,
      ...downloaded,
      timestamp: Date.now()
    });

    return { ok: true, exportType, taskId, ...downloaded };
  } finally {
    if (sourceTabId) {
      releaseInjectedExportLock(sourceTabId, String(exportType || "").toLowerCase());
    }
  }
}

addOnChunkedMessageListener(
  async (payload, context) => {
    const senderTabId = context?.sender?.tab?.id || null;
    const taskId = payload?.taskId || payload?.data?.taskId || null;
    if (payload.action === "store-image") {
      return handleStoreImage(payload.data);
    }
    if (payload.action === "store-image-by-candidates") {
      return handleStoreImageByCandidates(payload.data);
    }
    if (payload.action === "directDownload") {
      return handleDirectDownload(payload, { tabId: senderTabId, taskId });
    }
    if (payload.action === "assemble-export" || payload.action === "feishu-collect-assemble") {
      return handleAssembleExport(payload.data, { tabId: senderTabId, taskId });
    }
    return { ok: false, error: `未知 action: ${payload.action}` };
  },
  { context: "background" }
);

async function startExport(exportType, options) {
  const tabInfo = await getActiveTabInfo();
  if (!tabInfo.tabId) {
    throw new Error("当前标签页无效");
  }
  if (!isSupportedFeishuUrl(tabInfo.url)) {
    throw new Error("当前页面不在飞书/Lark支持域名下");
  }

  await ensureContentScriptReady(tabInfo.tabId);
  const normalizedExportType = String(exportType || "").toLowerCase();
  const taskId = crypto.randomUUID();
  createExportTask({
    taskId,
    tabId: tabInfo.tabId,
    exportType: normalizedExportType
  });
  const isInjectedExportType = INJECTED_EXPORT_TYPES.has(normalizedExportType);

  if (isInjectedExportType) {
    const acquired = acquireInjectedExportLock(tabInfo.tabId, normalizedExportType);
    if (!acquired) {
      throw attachTaskError(
        new Error("同类型导出任务正在执行，请等待当前任务完成"),
        taskId
      );
    }
  }

  emitProgress({
    title: "导出任务",
    message: `开始准备 ${String(normalizedExportType || "").toUpperCase()} 导出`,
    status: STATUS.RUNNING,
    taskId,
    taskState: EXPORT_TASK_STATE.RUNNING
  });

  if (
    normalizedExportType === "markdown" ||
    normalizedExportType === "html" ||
    normalizedExportType === "pdf" ||
    normalizedExportType === "docx"
  ) {
    try {
      const registerResponse = await chrome.tabs.sendMessage(tabInfo.tabId, {
        type: MESSAGE_TYPES.REGISTER_EXPORT_TASK,
        taskId,
        requestId: taskId,
        exportType: normalizedExportType,
        options: options || {}
      });
      if (registerResponse?.ok === false) {
        throw new Error(registerResponse.error || "内容脚本拒绝注册导出任务");
      }
      const injectedFile =
        normalizedExportType === "markdown"
          ? "injected-md.js"
          : normalizedExportType === "html" || normalizedExportType === "docx"
            ? "injected-html.js"
            : "injected-pdf.js";
      await chrome.scripting.executeScript({
        target: { tabId: tabInfo.tabId },
        files: [injectedFile],
        world: "MAIN"
      });
    } catch (error) {
      releaseInjectedExportLock(tabInfo.tabId, normalizedExportType);
      emitProgress({
        title: "导出任务",
        message: error instanceof Error ? error.message : String(error),
        status: STATUS.ERROR,
        taskId,
        taskState: EXPORT_TASK_STATE.ERROR
      });
      throw attachTaskError(error, taskId);
    }
    return { taskId };
  }

  const payload = {
    type: MESSAGE_TYPES.RUN_EXPORT,
    exportType: normalizedExportType,
    options: options || {},
    requestId: taskId,
    taskId
  };

  const sendRunExport = async () => {
    const response = await chrome.tabs.sendMessage(tabInfo.tabId, payload);
    if (response?.ok === false) {
      throw new Error(response.error || "内容脚本拒绝执行导出");
    }
  };

  try {
    await sendRunExport();
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    const shouldRetryByInjection =
      /Receiving end does not exist|Could not establish connection/i.test(msg);
    if (!shouldRetryByInjection) {
      throw attachTaskError(error, taskId);
    }

    emitProgress({
      title: "导出任务",
      message: "检测到内容脚本未就绪，正在自动注入并重试",
      status: STATUS.WARNING,
      taskId,
      taskState: EXPORT_TASK_STATE.RUNNING
    });

    await chrome.scripting.executeScript({
      target: { tabId: tabInfo.tabId },
      files: ["content-scripts/content.js"]
    });

    await sendRunExport().catch((retryError) => {
      throw attachTaskError(retryError, taskId);
    });
  }
  return { taskId };
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message) return false;

  if (message.type === MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE) {
    if (!message.__rebroadcast) {
      emitProgress(message.data, message.source || "runtime");
    }
    return false;
  }

  if (message.type === MESSAGE_TYPES.REQUEST_ACTIVE_TAB_INFO) {
    getActiveTabInfo()
      .then((info) => sendResponse({ ok: true, info }))
      .catch((error) =>
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error)
        })
      );
    return true;
  }

  if (message.type === MESSAGE_TYPES.START_EXPORT) {
    startExport(message.exportType, message.options)
      .then((result) => sendResponse({ ok: true, ...(result || {}) }))
      .catch((error) => {
        const taskId = error?.taskId || null;
        emitProgress({
          title: "导出任务",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskId,
          taskState: EXPORT_TASK_STATE.ERROR
        });
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error),
          taskId
        });
      });
    return true;
  }

  if (
    message.action === "directDownload" ||
    (message.type === "directDownload" && message.data)
  ) {
    handleDirectDownload(message, { tabId: sender?.tab?.id || null })
      .then((result) => sendResponse(result))
      .catch((error) => {
        const taskId = message?.taskId || message?.data?.taskId || null;
        emitProgress({
          title: "导出任务",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskId,
          taskState: EXPORT_TASK_STATE.ERROR
        });
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error),
          taskId
        });
      });
    return true;
  }

  if (message.type === MESSAGE_TYPES.PDF_PREVIEW_READY) {
    finalizePdfExport(message, sender)
      .then((result) => sendResponse(result))
      .catch((error) => {
        emitProgress({
          title: "PDF 导出",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskId: message?.taskId || null,
          taskState: EXPORT_TASK_STATE.ERROR
        });
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error),
          taskId: message?.taskId || null
        });
      });
    return true;
  }

  if (message.type === "OPEN_SIDE_PANEL") {
    chrome.windows
      .getCurrent()
      .then((win) => {
        if (win?.id) {
          return chrome.sidePanel.open({ windowId: win.id });
        }
        return null;
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) =>
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error)
        })
      );
    return true;
  }

  return false;
});

chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: true })
    .catch(() => {});

  chrome.tabs
    .query({})
    .then((tabs) =>
      Promise.all(
        tabs.map((tab) => ensureContentScriptReadyForTab(tab.id, tab.url || ""))
      )
    )
    .catch(() => {});
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status !== "complete") {
    return;
  }
  ensureContentScriptReadyForTab(tabId, tab?.url || "").catch(() => {});
});

openDatabase().catch(() => {});

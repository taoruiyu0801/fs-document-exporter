import { IDB_META, MESSAGE_TYPES } from "./utils/constants.js";
import { idbGet } from "./utils/idb.js";

const elements = {
  docTitle: document.getElementById("docTitle"),
  docUrl: document.getElementById("docUrl"),
  content: document.getElementById("content"),
  status: document.getElementById("status")
};

let hasNotified = false;

function setStatus(text, kind = "running") {
  elements.status.textContent = text;
  elements.status.className = `status ${kind}`;
}

function getPreviewIdFromUrl() {
  const url = new URL(location.href);
  return url.searchParams.get("previewId") || "";
}

function renderDocument(payload) {
  const title = payload.title || "PDF 预览";
  const currentUrl = payload.currentUrl || "";
  const htmlContent = payload.htmlContent || "<div>无内容</div>";

  document.title = `${title} - PDF 预览`;
  elements.docTitle.textContent = title;
  elements.docUrl.textContent = currentUrl || "-";
  elements.content.innerHTML = htmlContent;
}

function getCurrentTabId() {
  return new Promise((resolve) => {
    if (!chrome.tabs?.getCurrent) {
      resolve(null);
      return;
    }
    chrome.tabs.getCurrent((tab) => {
      resolve(tab?.id || null);
    });
  });
}

async function notifyPreviewReady(payload) {
  if (hasNotified) {
    return;
  }
  hasNotified = true;

  const tabId = await getCurrentTabId();
  setStatus("页面已渲染，正在生成 PDF...", "running");
  const response = await chrome.runtime.sendMessage({
    type: MESSAGE_TYPES.PDF_PREVIEW_READY,
    previewId: payload.previewId,
    taskId: payload.taskId || "",
    title: payload.title || "feishu-document",
    currentUrl: payload.currentUrl || "",
    pathParts: payload.pathParts || [],
    tabId
  });

  if (!response?.ok) {
    hasNotified = false;
    setStatus(response?.error || "PDF 生成失败", "error");
    return;
  }
  setStatus("PDF 已生成并触发下载。", "success");
}

async function renderAndGenerate(payload) {
  renderDocument(payload);
  await new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  });
  await notifyPreviewReady(payload);
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== MESSAGE_TYPES.GENERATE_PDF) {
    return false;
  }
  renderAndGenerate({
    previewId: message.previewId || "",
    taskId: message.taskId || "",
    title: message.title || "feishu-document",
    currentUrl: message.originalUrl || message.currentUrl || "",
    htmlContent: message.htmlContent || "",
    pathParts: message.pathParts || []
  })
    .then(() => sendResponse({ ok: true }))
    .catch((error) =>
      sendResponse({
        ok: false,
        error: error instanceof Error ? error.message : String(error)
      })
    );
  return true;
});

async function init() {
  try {
    const previewId = getPreviewIdFromUrl();
    if (!previewId) {
      throw new Error("缺少 previewId");
    }
    const preview = await idbGet(IDB_META.stores.previews, previewId);
    if (!preview) {
      throw new Error("未找到预览数据");
    }
    await renderAndGenerate({
      ...preview,
      previewId
    });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error), "error");
  }
}

init();

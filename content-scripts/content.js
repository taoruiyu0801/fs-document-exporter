(function () {
  if (window.__feishuExporterContentLoaded__) {
    return;
  }
  window.__feishuExporterContentLoaded__ = true;

  const CHUNKED_MESSAGE_FLAG = "__FEISHU_EXPORTER_CHUNKED_MESSAGE__";
  const MAX_CHUNK_SIZE = 32 * 1024 * 1024;
  const LARGE_DOCUMENT_THRESHOLD = 20 * 1024 * 1024;
  const IMAGE_BATCH_SIZE = 10;
  const IMAGE_CONCURRENCY = 4;
  const IMAGE_MAX_RETRY = 3;
  const IMAGE_FETCH_TIMEOUT_MS = 12000;
  const FALLBACK_EXPORT_TIMEOUT_MS = 8000;

  const MESSAGE_TYPES = {
    GLOBAL_PROGRESS_UPDATE: "GLOBAL_PROGRESS_UPDATE",
    RUN_EXPORT: "RUN_EXPORT",
    REGISTER_EXPORT_TASK: "REGISTER_EXPORT_TASK",
    REQUEST_PAGE_INFO: "REQUEST_PAGE_INFO",
    EXPORT_COMPLETE: "EXPORT_COMPLETE",
    MARKDOWN_EXPORT_COMPLETE: "MARKDOWN_EXPORT_COMPLETE",
    HTML_EXPORT_COMPLETE: "HTML_EXPORT_COMPLETE",
    WORD_EXPORT_COMPLETE: "WORD_EXPORT_COMPLETE",
    PDF_EXPORT_COMPLETE: "PDF_EXPORT_COMPLETE",
    BITABLE_EXPORT_COMPLETE: "BITABLE_EXPORT_COMPLETE"
  };

  const STATUS = {
    RUNNING: "running",
    SUCCESS: "success",
    WARNING: "warning",
    ERROR: "error"
  };

  const EXTENSION_BRIDGE_SOURCE = "feishu-exporter-extension";
  const INJECTED_BRIDGE_SOURCE = "feishu-exporter-injected";

  const inboundChunkCache = new Map();
  const monitorCache = new Map();
  const pendingPageInfoRequests = new Map();
  const pendingInjectedImageRequests = new Map();
  const pendingWorkerJobs = new Map();

  const scriptTagId = "__feishu_exporter_injected_script__";

  let isExporting = false;
  let worker = null;
  let exportWatchdogTimer = null;
  let activeExportRequest = null;
  const directResultGuard = new Map();
  let activeTask = null;

  function isBitableExportType(value) {
    const normalized = String(value || "").toLowerCase();
    return (
      normalized === "bitable-json" ||
      normalized === "bitable-csv" ||
      normalized === "bitable-xlsx"
    );
  }

  function hasBitableUrlHintLocal(url) {
    return /\/(?:share\/)?(?:base|bitable)\//i.test(String(url || ""));
  }

  function detectPageTypeLocal(url) {
    const value = String(url || "");
    if (hasBitableUrlHintLocal(value)) {
      return "bitable";
    }
    try {
      const embedded = document.querySelector(
        "iframe[src*='/share/base/'],iframe[src*='/base/'],iframe[src*='/bitable/'],a[href*='/share/base/'],a[href*='/base/'],a[href*='/bitable/']"
      );
      if (embedded) {
        return "bitable";
      }
    } catch {
      // ignore detect errors
    }
    if (value.includes("/wiki/") || value.includes("/docx/") || value.includes("/docs/")) {
      return "document";
    }
    return "unknown";
  }

  function resolveCurrentTaskId(fallback = "") {
    return (
      String(fallback || "") ||
      String(activeTask?.taskId || "") ||
      String(activeExportRequest?.taskId || activeExportRequest?.requestId || "")
    );
  }

  function withTaskProgress(progress, taskId = "") {
    const resolvedTaskId = resolveCurrentTaskId(taskId || progress?.taskId || "");
    if (!resolvedTaskId) {
      return progress;
    }
    return {
      ...progress,
      taskId: resolvedTaskId
    };
  }

  function beginTask({ taskId, exportType }) {
    const normalizedTaskId = String(taskId || "").trim();
    activeTask = normalizedTaskId
      ? {
        taskId: normalizedTaskId,
        exportType: String(exportType || "").toLowerCase(),
        startedAt: Date.now()
      }
      : null;
  }

  function clearTask() {
    activeTask = null;
  }

  function isTaskCompletionMessage(message) {
    const type = String(message?.type || "");
    return (
      type === MESSAGE_TYPES.EXPORT_COMPLETE ||
      type === MESSAGE_TYPES.MARKDOWN_EXPORT_COMPLETE ||
      type === MESSAGE_TYPES.HTML_EXPORT_COMPLETE ||
      type === MESSAGE_TYPES.WORD_EXPORT_COMPLETE ||
      type === MESSAGE_TYPES.PDF_EXPORT_COMPLETE ||
      type === MESSAGE_TYPES.BITABLE_EXPORT_COMPLETE
    );
  }

  function splitText(text, maxSize = MAX_CHUNK_SIZE) {
    const chunks = [];
    for (let i = 0; i < text.length; i += maxSize) {
      chunks.push(text.slice(i, i + maxSize));
    }
    return chunks;
  }

  function cleanupInbound(requestId) {
    inboundChunkCache.delete(requestId);
  }

  function cleanupMonitor(requestId) {
    monitorCache.delete(requestId);
  }

  function ensureMonitor(requestId) {
    const current = monitorCache.get(requestId);
    if (current) return current;
    const initial = {
      chunks: [],
      expectedChunks: 0,
      resolve: null,
      reject: null,
      timer: null,
      completed: false,
      result: null,
      error: null
    };
    monitorCache.set(requestId, initial);
    return initial;
  }

  function waitForMonitor(requestId, timeoutMs = 120000) {
    const monitor = ensureMonitor(requestId);
    if (monitor.completed) {
      if (monitor.error) {
        cleanupMonitor(requestId);
        return Promise.reject(monitor.error);
      }
      const value = monitor.result;
      cleanupMonitor(requestId);
      return Promise.resolve(value);
    }
    return new Promise((resolve, reject) => {
      monitor.resolve = resolve;
      monitor.reject = reject;
      monitor.timer = setTimeout(() => {
        cleanupMonitor(requestId);
        reject(new Error(`Chunked monitor timeout: ${requestId}`));
      }, timeoutMs);
    });
  }

  function resolveMonitorChunk(message) {
    const monitor = ensureMonitor(message.requestId);
    if (typeof message.totalChunks === "number" && message.totalChunks > 0) {
      monitor.expectedChunks = message.totalChunks;
    }
    if (message.type === "chunked-message-result-chunk") {
      monitor.chunks[message.chunkIndex] = message.chunk;
      return;
    }
    if (message.type === "chunked-message-result-done") {
      const text = monitor.chunks.join("");
      try {
        const parsed = JSON.parse(text);
        monitor.completed = true;
        monitor.result = parsed;
        if (monitor.timer) clearTimeout(monitor.timer);
        if (monitor.resolve) {
          monitor.resolve(parsed);
          cleanupMonitor(message.requestId);
        }
      } catch (error) {
        monitor.completed = true;
        monitor.error = error;
        if (monitor.timer) clearTimeout(monitor.timer);
        if (monitor.reject) {
          monitor.reject(error);
          cleanupMonitor(message.requestId);
        }
      }
    }
  }

  async function sendMessageToSender({ sender, message, runtime, context }) {
    if (context === "background" && sender?.tab?.id) {
      return chrome.tabs.sendMessage(sender.tab.id, message);
    }
    return runtime.sendMessage(message);
  }

  async function pushChunkedResponse({
    result,
    sender,
    runtime,
    context,
    monitorRequestId
  }) {
    const serialized = JSON.stringify(result);
    const chunks = splitText(serialized, MAX_CHUNK_SIZE);
    for (let i = 0; i < chunks.length; i += 1) {
      await sendMessageToSender({
        sender,
        runtime,
        context,
        message: {
          [CHUNKED_MESSAGE_FLAG]: true,
          type: "chunked-message-result-chunk",
          requestId: monitorRequestId,
          chunkIndex: i,
          totalChunks: chunks.length,
          chunk: chunks[i],
          timestamp: Date.now()
        }
      });
    }
    await sendMessageToSender({
      sender,
      runtime,
      context,
      message: {
        [CHUNKED_MESSAGE_FLAG]: true,
        type: "chunked-message-result-done",
        requestId: monitorRequestId,
        totalChunks: chunks.length,
        done: true,
        timestamp: Date.now()
      }
    });
  }

  async function sendChunkedMessage(payload, options = {}) {
    const sendMessageFn = options.sendMessageFn || ((message) => chrome.runtime.sendMessage(message));
    const timeoutMs = options.timeoutMs || 120000;

    if (options.requestIdToMonitor) {
      return waitForMonitor(options.requestIdToMonitor, timeoutMs);
    }

    const requestId = options.requestId || crypto.randomUUID();
    const serialized = JSON.stringify(payload);
    const chunks = splitText(serialized, MAX_CHUNK_SIZE);

    for (let i = 0; i < chunks.length; i += 1) {
      const ack = await sendMessageFn({
        [CHUNKED_MESSAGE_FLAG]: true,
        type: "chunked-message-chunk",
        requestId,
        chunkIndex: i,
        totalChunks: chunks.length,
        chunk: chunks[i],
        timestamp: Date.now()
      });
      if (!ack || ack[CHUNKED_MESSAGE_FLAG] !== true || ack.type !== "chunked-message-ack") {
        throw new Error(`Chunk ACK missing for ${requestId}:${i}`);
      }
    }

    const doneAck = await sendMessageFn({
      [CHUNKED_MESSAGE_FLAG]: true,
      type: "chunked-message-done",
      requestId,
      done: true,
      timestamp: Date.now()
    });

    if (!doneAck) {
      throw new Error(`Chunk done ACK missing for ${requestId}`);
    }
    if (doneAck.type === "chunked-message-error") {
      throw new Error(doneAck.error || `Chunk handler failed for ${requestId}`);
    }
    if (doneAck.type === "chunked-message-monitor" && doneAck.monitorRequestId) {
      return waitForMonitor(doneAck.monitorRequestId, timeoutMs);
    }
    return doneAck.result;
  }

  function addOnChunkedMessageListener(
    handler,
    { runtime = chrome.runtime, context = "background" } = {}
  ) {
    const listener = (message, sender, sendResponse) => {
      if (!message || message[CHUNKED_MESSAGE_FLAG] !== true) {
        return false;
      }

      if (
        message.type === "chunked-message-result-chunk" ||
        message.type === "chunked-message-result-done"
      ) {
        resolveMonitorChunk(message);
        sendResponse({
          [CHUNKED_MESSAGE_FLAG]: true,
          type: "chunked-message-ack",
          requestId: message.requestId,
          status: "RESULT_PENDING"
        });
        return false;
      }

      if (message.type === "chunked-message-chunk") {
        const bucket =
          inboundChunkCache.get(message.requestId) ||
          {
            chunks: [],
            totalChunks: message.totalChunks || 0
          };
        bucket.chunks[message.chunkIndex] = message.chunk;
        if (message.totalChunks) {
          bucket.totalChunks = message.totalChunks;
        }
        inboundChunkCache.set(message.requestId, bucket);
        sendResponse({
          [CHUNKED_MESSAGE_FLAG]: true,
          type: "chunked-message-ack",
          requestId: message.requestId,
          chunkIndex: message.chunkIndex,
          status: "PENDING"
        });
        return false;
      }

      if (message.type === "chunked-message-done") {
        const processAsync = async () => {
          const bucket = inboundChunkCache.get(message.requestId);
          if (!bucket) {
            sendResponse({
              [CHUNKED_MESSAGE_FLAG]: true,
              type: "chunked-message-error",
              requestId: message.requestId,
              error: "Chunk bucket missing"
            });
            return;
          }

          try {
            const jsonText = bucket.chunks.join("");
            const payload = JSON.parse(jsonText);
            cleanupInbound(message.requestId);

            const result = await handler(payload, {
              sender,
              requestId: message.requestId
            });

            if (typeof result === "undefined") {
              sendResponse({
                [CHUNKED_MESSAGE_FLAG]: true,
                type: "chunked-message-ack",
                requestId: message.requestId,
                status: "DONE"
              });
              return;
            }

            const monitorRequestId = `${message.requestId}:response`;
            await pushChunkedResponse({
              result,
              sender,
              runtime,
              context,
              monitorRequestId
            });
            sendResponse({
              [CHUNKED_MESSAGE_FLAG]: true,
              type: "chunked-message-monitor",
              requestId: message.requestId,
              monitorRequestId
            });
          } catch (error) {
            cleanupInbound(message.requestId);
            sendResponse({
              [CHUNKED_MESSAGE_FLAG]: true,
              type: "chunked-message-error",
              requestId: message.requestId,
              error: error instanceof Error ? error.message : String(error)
            });
          }
        };

        processAsync();
        return true;
      }

      return false;
    };

    runtime.onMessage.addListener(listener);
    return () => runtime.onMessage.removeListener(listener);
  }

  function buildImageRecordId(documentUrl, token) {
    return `${documentUrl}_${token}`;
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function clearExportWatchdog() {
    if (exportWatchdogTimer) {
      clearTimeout(exportWatchdogTimer);
      exportWatchdogTimer = null;
    }
  }

  function markAndCheckDuplicateDirectResult(kind, signature, ttlMs = 8000) {
    const key = `${kind}:${signature}`;
    const now = Date.now();
    for (const [cacheKey, ts] of directResultGuard) {
      if (now - ts > ttlMs) {
        directResultGuard.delete(cacheKey);
      }
    }
    const prev = directResultGuard.get(key);
    if (prev && now - prev <= ttlMs) {
      return true;
    }
    directResultGuard.set(key, now);
    return false;
  }

  function emitProgress(progress, source = "content") {
    chrome.runtime
      .sendMessage({
        type: MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE,
        source,
        data: withTaskProgress(progress),
        timestamp: Date.now()
      })
      .catch(() => { });
  }

  function sanitizeFilenameForDownload(input, fallback = "untitled") {
    const value = String(input || "")
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
      .replace(/[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g, "")
      .replace(/[<>:"/\\|?*]/g, "_")
      .replace(/\s+/g, " ")
      .trim();
    return value || fallback;
  }

  function blobToDataUrlLocal(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("读取文件数据失败"));
      reader.readAsDataURL(blob);
    });
  }

  function bytesToBase64Local(bytes) {
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function decodeTextDataUrl(dataUrl) {
    const value = String(dataUrl || "");
    const match = /^data:([^;,]+(?:;charset=[^;,]+)?)?(?:;(base64))?,([\s\S]*)$/i.exec(value);
    if (!match) return null;
    const mimeType = match[1] || "text/plain;charset=utf-8";
    const isBase64 = Boolean(match[2]);
    const payload = match[3] || "";
    try {
      if (isBase64) {
        const binary = atob(payload);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i += 1) {
          bytes[i] = binary.charCodeAt(i);
        }
        return {
          mimeType,
          text: new TextDecoder("utf-8").decode(bytes)
        };
      }
      return {
        mimeType,
        text: decodeURIComponent(payload)
      };
    } catch {
      return null;
    }
  }

  function encodeTextDataUrl(text, mimeType = "text/html;charset=utf-8") {
    const bytes = new TextEncoder().encode(String(text || ""));
    return `data:${mimeType};base64,${bytesToBase64Local(bytes)}`;
  }

  function escapeHtmlLocal(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function normalizeRecoveredUrl(rawUrl) {
    const raw = String(rawUrl || "").trim();
    if (!raw) return "";
    const unescaped = raw
      .replace(/\\u003a/gi, ":")
      .replace(/\\u002f/gi, "/")
      .replace(/\\\//g, "/")
      .replace(/&amp;/gi, "&")
      .replace(/^['"(<\[]+/, "")
      .replace(/[>'")\],.;]+$/, "");
    const abs = toAbsoluteUrl(unescaped);
    if (!/^https?:\/\//i.test(abs)) return "";
    return abs;
  }

  function isLikelyUserFacingLink(url) {
    try {
      const parsed = new URL(url);
      const host = parsed.hostname.toLowerCase();
      const path = parsed.pathname.toLowerCase();
      const mountPoint = (parsed.searchParams.get("mount_point") || "").toLowerCase();
      const trustedHost =
        /(^|\.)feishu\.cn$/i.test(host) ||
        /(^|\.)feishu\.net$/i.test(host) ||
        /(^|\.)larksuite\.com$/i.test(host) ||
        /(^|\.)larkoffice\.com$/i.test(host) ||
        /(^|\.)chromewebstore\.google\.com$/i.test(host) ||
        /(^|\.)microsoftedge\.microsoft\.com$/i.test(host) ||
        /(^|\.)crxsoso\.com$/i.test(host) ||
        /(^|\.)github\.com$/i.test(host);
      if (!trustedHost) {
        return false;
      }
      if (host === "localhost" || host === "127.0.0.1") return false;
      if (/^internal-api-(?:drive-stream|space|drive)\./i.test(host)) return false;
      if (/\/space\/api\/box\/stream\/download\//i.test(path)) return false;
      if (/\/space\/api\/medias?\/batch_get\//i.test(path)) return false;
      if (mountPoint === "docx_image" || mountPoint === "docx_file") return false;
      if (/\.(?:js|css|png|jpe?g|webp|gif|svg|ico|woff2?|ttf|map|mp4|webm|m3u8|json)$/i.test(path)) {
        return false;
      }
      if (/\/(?:static|assets?|resource|images?)\//i.test(path) && !/\/(?:wiki|docx|docs|share|base)\//i.test(path)) {
        return false;
      }
      return true;
    } catch {
      return false;
    }
  }

  function extractUrlsFromTextForRecovery(text) {
    const source = String(text || "")
      .replace(/\\u003a/gi, ":")
      .replace(/\\u002f/gi, "/")
      .replace(/\\\//g, "/");
    const matches = source.match(/https?:\/\/[^\s"'<>`\\]+/gi) || [];
    return matches.map((value) => normalizeRecoveredUrl(value)).filter(Boolean);
  }

  function linkPriority(url) {
    const value = String(url || "").toLowerCase();
    if (value.includes("chromewebstore.google.com")) return 120;
    if (value.includes("microsoftedge.microsoft.com")) return 120;
    if (value.includes("crxsoso.com")) return 120;
    if (value.includes("github.com")) return 90;
    if (value.includes("feishu.cn") || value.includes("larksuite.com") || value.includes("larkoffice.com")) return 80;
    return 10;
  }

  const INSTALL_FALLBACK_LINKS = [
    {
      id: "chrome",
      url: "https://chromewebstore.google.com/detail/cbcbonoeikcdhobbmoaakgodaiknmidk",
      label: "Chrome 应用商店安装链接",
      hostPattern: /(^|\.)chromewebstore\.google\.com$/i,
      headingPattern:
        /<h[1-6][^>]*>(?:(?!<\/h[1-6]>)[\s\S])*(?:Chrome\s*应用商店|Chrome\s*web\s*store|访问Chrome)(?:(?!<\/h[1-6]>)[\s\S])*<\/h[1-6]>/i
    },
    {
      id: "edge",
      url: "https://microsoftedge.microsoft.com/addons/detail/cbcbonoeikcdhobbmoaakgodaiknmidk",
      label: "Edge 应用商店安装链接",
      hostPattern: /(^|\.)microsoftedge\.microsoft\.com$/i,
      headingPattern:
        /<h[1-6][^>]*>(?:(?!<\/h[1-6]>)[\s\S])*(?:无法访问chrome商店|edge浏览器|微软的edge)(?:(?!<\/h[1-6]>)[\s\S])*<\/h[1-6]>/i
    },
    {
      id: "crxsoso",
      url: "https://www.crxsoso.com/webstore/detail/cfenjfhlhjpkaaobmhbobajnnhifilbl",
      label: "CRXSoso 安装链接",
      hostPattern: /(^|\.)crxsoso\.com$/i,
      headingPattern:
        /<h[1-6][^>]*>(?:(?!<\/h[1-6]>)[\s\S])*(?:点下面链接安装|无法访问\s*应用商店|crxsoso)(?:(?!<\/h[1-6]>)[\s\S])*<\/h[1-6]>/i
    }
  ];

  function dedupeRecoveredLinks(links) {
    const map = new Map();
    for (const item of Array.isArray(links) ? links : []) {
      if (!item?.url) continue;
      const url = normalizeRecoveredUrl(item.url);
      if (!url || !isLikelyUserFacingLink(url)) continue;
      if (map.has(url)) continue;
      const text = String(item.text || "").trim();
      map.set(url, {
        url,
        text: text || url
      });
    }
    return Array.from(map.values());
  }

  function findRecoveredLinkByHost(links, hostPattern) {
    for (const item of links) {
      try {
        const host = new URL(item.url).hostname || "";
        if (hostPattern.test(host)) {
          return item;
        }
      } catch {
        // ignore invalid URL
      }
    }
    return null;
  }

  function ensureInstallFallbackLinks(links, sourceText) {
    const deduped = dedupeRecoveredLinks(links);
    const combinedText = String(sourceText || "").replace(/\s+/g, " ");
    const hasInstallHint =
      /应用商店|chrome\s*web\s*store|chromewebstore|microsoftedge|edge浏览器|crxsoso/i.test(
        combinedText
      );
    if (!hasInstallHint) {
      return deduped;
    }

    for (const preset of INSTALL_FALLBACK_LINKS) {
      if (findRecoveredLinkByHost(deduped, preset.hostPattern)) {
        continue;
      }
      deduped.push({
        url: preset.url,
        text: preset.label
      });
    }
    return dedupeRecoveredLinks(deduped);
  }

  function buildInlineRecoveredLinkCard(link, blockId, fallbackLabel = "") {
    const text = String(link?.text || fallbackLabel || link?.url || "").trim();
    if (!link?.url) return "";
    return `<div data-nv-install-link="${escapeHtmlLocal(
      blockId
    )}" style="margin:10px 0 14px;padding:10px 12px;border:1px solid #d7deea;border-radius:10px;background:#f8fafc;">
  <a href="${escapeHtmlLocal(
      link.url
    )}" target="_blank" rel="noopener noreferrer" style="color:#1d4ed8;word-break:break-all;text-decoration:none;font-weight:600;">${escapeHtmlLocal(
      text || fallbackLabel || link.url
    )}</a>
  <div style="margin-top:4px;font-size:12px;color:#64748b;word-break:break-all;">${escapeHtmlLocal(
      link.url
    )}</div>
</div>`;
  }

  function injectInstallLinksByHeading(htmlText, links) {
    let html = String(htmlText || "");
    const usedUrls = new Set();
    let insertedCount = 0;

    for (const preset of INSTALL_FALLBACK_LINKS) {
      const marker = `data-nv-install-link="${preset.id}"`;
      if (html.includes(marker)) {
        continue;
      }
      const matchedLink =
        findRecoveredLinkByHost(links, preset.hostPattern) || {
          url: preset.url,
          text: preset.label
        };
      const card = buildInlineRecoveredLinkCard(matchedLink, preset.id, preset.label);
      if (!card) continue;
      const replaced = html.replace(preset.headingPattern, (matched) => `${matched}\n${card}`);
      if (replaced !== html) {
        html = replaced;
        usedUrls.add(matchedLink.url);
        insertedCount += 1;
      }
    }

    return {
      html,
      usedUrls,
      insertedCount
    };
  }

  function collectPageExternalLinksForRecovery() {
    const map = new Map();
    const append = (rawUrl, rawText = "") => {
      const url = normalizeRecoveredUrl(rawUrl);
      if (!url || !isLikelyUserFacingLink(url)) return;
      if (map.has(url)) return;
      const text = String(rawText || "")
        .replace(/\s+/g, " ")
        .trim();
      map.set(url, {
        url,
        text: text || url
      });
    };

    for (const a of Array.from(document.querySelectorAll("a[href]"))) {
      append(a.getAttribute("href") || a.href, a.textContent || a.getAttribute("title") || "");
    }

    const attrNodes = Array.from(
      document.querySelectorAll("[data-href],[data-url],[data-link],[data-link-url],[data-redirect-url]")
    );
    for (const node of attrNodes) {
      append(node.getAttribute("data-href"), node.textContent || "");
      append(node.getAttribute("data-url"), node.textContent || "");
      append(node.getAttribute("data-link"), node.textContent || "");
      append(node.getAttribute("data-link-url"), node.textContent || "");
      append(node.getAttribute("data-redirect-url"), node.textContent || "");
    }

    const rawSources = [];
    rawSources.push(document.documentElement?.outerHTML || "");
    for (const scriptNode of Array.from(document.querySelectorAll("script"))) {
      if (rawSources.length >= 40) break;
      const text = scriptNode.textContent || "";
      if (text.length < 20) continue;
      rawSources.push(text.slice(0, 200000));
    }

    for (const source of rawSources) {
      for (const url of extractUrlsFromTextForRecovery(source)) {
        append(url, url);
      }
    }

    const bodyText = String(document.body?.innerText || "");
    const hasChromeStoreHint =
      /chrome\s*应用商店|chrome\s*web\s*store|无法访问.*应用商店/i.test(bodyText);
    const hasEdgeHint = /edge浏览器|microsoftedge/i.test(bodyText);
    const hasCrxHint = /crxsoso|点下面链接安装/i.test(bodyText);

    const links = Array.from(map.values())
      .sort((a, b) => {
        const p = linkPriority(b.url) - linkPriority(a.url);
        if (p !== 0) return p;
        return String(a.url).length - String(b.url).length;
      });

    const installHintText = [
      bodyText,
      hasChromeStoreHint ? "chrome" : "",
      hasEdgeHint ? "edge" : "",
      hasCrxHint ? "crx" : ""
    ]
      .filter(Boolean)
      .join(" ");

    return ensureInstallFallbackLinks(links, installHintText).slice(0, 80);
  }

  function buildRecoveredLinksBlock(links) {
    if (!Array.isArray(links) || !links.length) return "";
    const items = links
      .map(
        (item) =>
          `<li style="margin:6px 0;"><a href="${escapeHtmlLocal(item.url)}" target="_blank" rel="noopener noreferrer" style="color:#1d4ed8;word-break:break-all;text-decoration:none;">${escapeHtmlLocal(item.text)}</a></li>`
      )
      .join("");
    return `
<section id="nv-link-recovery" style="margin:16px 0;padding:12px;border:1px solid #d7deea;border-radius:10px;background:#f8fafc;">
  <div style="font-weight:700;color:#0f172a;margin-bottom:8px;">链接补充（自动恢复）</div>
  <ul style="margin:0;padding-left:18px;">${items}</ul>
</section>`;
  }

  function hasInstallRecoveryHint(value) {
    const text = String(value || "").replace(/\s+/g, " ");
    return /chrome\s*应用商店|chrome\s*web\s*store|无法访问chrome商店|edge浏览器|应用商店|点下面链接安装|crxsoso/i.test(
      text
    );
  }

  function appendRecoveredLinksToHtml(htmlText, links) {
    const html = String(htmlText || "");
    if (!html) {
      return html;
    }
    const validLinks = ensureInstallFallbackLinks(
      (Array.isArray(links) ? links : []).filter((item) => item?.url),
      html
    ).slice(0, 40);
    if (!validLinks.length) {
      return html;
    }

    const hasInstallHint =
      hasInstallRecoveryHint(html) ||
      validLinks.some((item) =>
        /chromewebstore\.google\.com|microsoftedge\.microsoft\.com|crxsoso\.com/i.test(
          String(item?.url || "")
        )
      );
    if (!hasInstallHint) {
      return html;
    }

    const injected = injectInstallLinksByHeading(html, validLinks);
    if (injected.insertedCount > 0) {
      return injected.html;
    }

    if (html.includes("id=\"nv-link-recovery\"")) {
      return html;
    }

    const block = buildRecoveredLinksBlock(validLinks);
    if (!block) return html;
    const installHeadingPattern =
      /<h[1-6][^>]*>[^<]*(?:Chrome\s*应用商店|无法访问chrome商店|edge浏览器|点下面链接安装|zip安装包)[^<]*<\/h[1-6]>/i;
    if (installHeadingPattern.test(html)) {
      return html.replace(installHeadingPattern, (matched) => `${matched}\n${block}`);
    }
    if (/<\/body>/i.test(html)) {
      return html.replace(/<\/body>/i, `${block}\n</body>`);
    }
    return `${html}\n${block}`;
  }

  function recoverMissingLinksInReferenceDataUrl(url) {
    const decoded = decodeTextDataUrl(url);
    if (!decoded || !/text\/html/i.test(decoded.mimeType)) {
      return String(url || "");
    }
    const links = collectPageExternalLinksForRecovery();
    const enhancedHtml = appendRecoveredLinksToHtml(decoded.text, links);
    if (enhancedHtml === decoded.text) {
      return String(url || "");
    }
    return encodeTextDataUrl(enhancedHtml, decoded.mimeType || "text/html;charset=utf-8");
  }

  function recoverMissingLinksInHtmlContent(htmlText) {
    return appendRecoveredLinksToHtml(htmlText, collectPageExternalLinksForRecovery());
  }

  async function buildMarkdownDownloadTarget(meta) {
    const titleSafe = sanitizeFilenameForDownload(
      meta?.title || document.title || "飞书文档",
      "飞书文档"
    );
    const mimeType = String(meta?.mimeType || "").toLowerCase();
    const isZipByMime = mimeType.includes("zip");
    const isZip = Boolean(meta?.isZip || isZipByMime);
    const ext = isZip ? "zip" : "md";
    const recommendBase = sanitizeFilenameForDownload(
      meta?.recommendName || titleSafe,
      titleSafe
    );
    const filename = `${recommendBase}.${ext}`;
    const requestedPathParts = Array.isArray(activeExportRequest?.options?.pathParts)
      ? activeExportRequest.options.pathParts
        .map((part) => String(part || "").trim())
        .filter(Boolean)
      : [];
    return {
      filename,
      pathParts: requestedPathParts.length ? requestedPathParts : undefined
    };
  }

  async function handleReferenceMarkdownSuccess(rawMessage) {
    const data = rawMessage?.data || {};
    const arrayBuffer = data.arrayBuffer;
    if (!(arrayBuffer instanceof ArrayBuffer)) {
      throw new Error("MARKDOWN_EXPORT_SUCCESS 缺少有效的 arrayBuffer");
    }

    const byteLength = Number(arrayBuffer?.byteLength || 0);
    const duplicateSignature = `${data?.title || ""}|${data?.mimeType || ""}|${byteLength}`;
    if (markAndCheckDuplicateDirectResult("markdown-success", duplicateSignature)) {
      emitProgress({
        title: "Markdown 导出",
        message: "检测到重复成功消息，已忽略",
        status: STATUS.WARNING
      });
      return { ok: true, skipped: true };
    }

    emitProgress({
      title: "Markdown 导出",
      message: `页面已生成 Markdown（${data?.mimeType || "unknown"}），正在提交下载任务`,
      status: STATUS.RUNNING
    });

    const blob = new Blob([arrayBuffer], {
      type: data.mimeType || "application/octet-stream"
    });
    const dataUrl = await blobToDataUrlLocal(blob);
    const target = await buildMarkdownDownloadTarget(data);

    const taskId = resolveCurrentTaskId();
    return sendChunkedMessage({
      action: "directDownload",
      taskId,
      exportType: "MARKDOWN",
      data: {
        url:
          typeof dataUrl === "string" && dataUrl.startsWith("data:")
            ? dataUrl
            : `data:application/octet-stream;base64,${dataUrl}`,
        filename: target.filename,
        pathParts: target.pathParts,
        title: data.title || document.title || "飞书文档",
        isBatchDownload: Boolean(data.isBatchDownload),
        hasImages: Boolean(data.hasImages),
        isBigMd: Boolean(data.isBigMd),
        isZip: Boolean(data.isZip),
        taskId
      }
    });
  }

  async function handleReferenceHtmlDownload(rawMessage) {
    const data = rawMessage?.data || {};
    if (!data?.url) {
      throw new Error("downloadHtml 缺少 url");
    }

    const currentExportType = String(
      activeTask?.exportType || activeExportRequest?.exportType || ""
    ).toLowerCase();

    // DOCX export: decode HTML from data URL, then route through assemble-export
    // so background.js calls buildWordExport instead of direct downloading HTML
    if (currentExportType === "docx" || currentExportType === "word") {
      const duplicateSignature = `${data?.filename || ""}|${String(data?.url || "").length}|docx`;
      if (markAndCheckDuplicateDirectResult("docx-success", duplicateSignature)) {
        emitProgress({
          title: "DOCX 导出",
          message: "检测到重复成功消息，已忽略",
          status: STATUS.WARNING
        });
        return { ok: true, skipped: true };
      }

      emitProgress({
        title: "DOCX 导出",
        message: "页面已生成 HTML，正在转换为 DOCX...",
        status: STATUS.RUNNING
      });

      // Decode HTML from data URL
      const decoded = decodeTextDataUrl(data.url);
      const htmlContent = decoded?.text || "";
      if (!htmlContent) {
        throw new Error("无法解码 HTML 内容用于 DOCX 转换");
      }

      // Route through handleExportDataMessage which sends assemble-export to background
      // Images are already embedded as data URLs in the HTML src attributes,
      // so word-exporter.js resolveImageEntry will read them directly - no need to fetch
      return handleExportDataMessage({
        exportType: "docx",
        taskId: resolveCurrentTaskId(),
        requestId: resolveCurrentTaskId(),
        options: activeExportRequest?.options || {},
        payload: {
          title: data.title || document.title || "飞书文档",
          currentUrl: location.href,
          htmlContent: htmlContent,
          htmlWithPlaceholders: htmlContent,
          imageInfoList: [],
          imageTokens: [],
          htmlBlobSize: new Blob([htmlContent]).size
        }
      });
    }

    const duplicateSignature = `${data?.filename || ""}|${String(data?.url || "").length}`;
    if (markAndCheckDuplicateDirectResult("html-success", duplicateSignature)) {
      emitProgress({
        title: "HTML 导出",
        message: "检测到重复成功消息，已忽略",
        status: STATUS.WARNING
      });
      return { ok: true, skipped: true };
    }

    emitProgress({
      title: "HTML 导出",
      message: "页面已生成 HTML，正在提交下载任务",
      status: STATUS.RUNNING
    });

    const recoveredUrl = recoverMissingLinksInReferenceDataUrl(data.url);
    const taskId = resolveCurrentTaskId();
    return sendChunkedMessage({
      action: "directDownload",
      taskId,
      exportType: "HTML",
      data: {
        url: recoveredUrl || data.url,
        filename: data.filename,
        title: data.title || document.title || "飞书文档",
        taskId
      }
    });
  }

  async function handleReferencePdfImageProcessed(rawMessage) {
    const data = rawMessage?.data || {};
    const token = data.token;
    const dataUrl = data.dataUrl;
    const documentUrl = data.originalUrl || location.href;
    if (!token || !dataUrl) {
      return { ok: false, skipped: true };
    }

    const taskId = resolveCurrentTaskId();
    return sendChunkedMessage({
      action: "store-image",
      taskId,
      data: {
        id: buildImageRecordId(documentUrl, token),
        token,
        dataUrl,
        documentUrl,
        taskId
      }
    });
  }

  async function handleReferencePdfAllImagesProcessed(rawMessage) {
    const data = rawMessage?.data || {};
    const title = data.title || document.title || "飞书文档";
    const currentUrl = data.originalUrl || location.href;
    const htmlWithPlaceholders = recoverMissingLinksInHtmlContent(data.htmlContent || "");
    const imageTokens = Array.isArray(data.imageTokens) ? data.imageTokens : [];

    if (!htmlWithPlaceholders) {
      throw new Error("PDF htmlContent 为空");
    }

    emitProgress({
      title: "PDF 导出",
      message: "图片处理完成，开始组装 PDF 预览",
      status: STATUS.RUNNING
    });

    const taskId = resolveCurrentTaskId();
    return sendChunkedMessage({
      action: "assemble-export",
      taskId,
      data: {
        title,
        currentUrl,
        htmlWithPlaceholders,
        imageTokens,
        exportType: "pdf",
        options: {},
        htmlBlobSize: new Blob([htmlWithPlaceholders]).size,
        taskId
      }
    });
  }

  function ensureInjectedScript() {
    return new Promise((resolve) => {
      const existing = document.getElementById(scriptTagId);
      if (existing) {
        resolve();
        return;
      }
      const script = document.createElement("script");
      script.id = scriptTagId;
      script.src = chrome.runtime.getURL("injected.js");
      script.async = false;
      script.onload = () => resolve();
      script.onerror = () => resolve();
      (document.head || document.documentElement).appendChild(script);
    });
  }

  function getWorker() {
    if (worker) return worker;
    worker = new Worker(chrome.runtime.getURL("image-processor.worker.js"));
    worker.addEventListener("message", (event) => {
      const data = event.data || {};
      if (!data.jobId) return;
      const job = pendingWorkerJobs.get(data.jobId);
      if (!job) return;
      pendingWorkerJobs.delete(data.jobId);
      clearTimeout(job.timer);
      if (data.action === "image-ready") {
        job.resolve(data);
      } else {
        job.reject(new Error(data.error || "Image worker failed"));
      }
    });
    worker.addEventListener("error", (event) => {
      for (const [jobId, job] of pendingWorkerJobs) {
        clearTimeout(job.timer);
        job.reject(new Error(event.message || "Image worker crashed"));
        pendingWorkerJobs.delete(jobId);
      }
    });
    return worker;
  }

  function postToInjected(message) {
    window.postMessage(message, "*");
  }

  function requestPageInfoFromInjected() {
    const requestId = crypto.randomUUID();
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        pendingPageInfoRequests.delete(requestId);
        resolve({
          title: document.title || "",
          url: location.href,
          pageType: detectPageTypeLocal(location.href)
        });
      }, 1200);
      pendingPageInfoRequests.set(requestId, (info) => {
        clearTimeout(timer);
        resolve(info);
      });
      postToInjected({
        source: EXTENSION_BRIDGE_SOURCE,
        type: MESSAGE_TYPES.REQUEST_PAGE_INFO,
        requestId
      });
    });
  }

  function sanitizeToken(value) {
    const base = String(value || "")
      .replace(/[^a-zA-Z0-9._-]+/g, "-")
      .replace(/^-+|-+$/g, "");
    return base || "img";
  }

  function toAbsoluteUrl(rawUrl) {
    const value = String(rawUrl || "").trim();
    if (!value || /^data:/i.test(value)) return "";
    try {
      return new URL(value, location.href).href;
    } catch {
      return "";
    }
  }

  function guessMimeType(url) {
    const value = String(url || "").toLowerCase();
    if (value.includes(".png")) return "image/png";
    if (value.includes(".jpg") || value.includes(".jpeg")) return "image/jpeg";
    if (value.includes(".webp")) return "image/webp";
    if (value.includes(".gif")) return "image/gif";
    if (value.includes(".svg")) return "image/svg+xml";
    return "";
  }

  function tokenFromUrl(url, index) {
    const value = String(url || "");
    if (/^(?:blob|data):/i.test(value)) {
      return `img-${index + 1}`;
    }
    let candidate = "";
    const tokenFromQuery = /[?&](?:token|image|img|id)=([^&#]+)/i.exec(value);
    if (tokenFromQuery) {
      candidate = decodeURIComponent(tokenFromQuery[1]);
    }
    if (!candidate) {
      const parts = value.split(/[/?#]/).filter(Boolean);
      candidate = parts[parts.length - 1] || `img-${index + 1}`;
    }
    return sanitizeToken(candidate);
  }

  function extractBackgroundUrl(styleValue) {
    const style = String(styleValue || "");
    const matched = /background(?:-image)?\s*:\s*[^;]*url\((['"]?)(.*?)\1\)/i.exec(style);
    return matched ? matched[2] : "";
  }

  function getMainContainerSelectorsLocal() {
    return [
      ".wiki-content",
      ".docs-editor-container",
      ".lark-doc-content",
      ".docs-content",
      ".docx-content",
      ".doc-body",
      ".text-viewer",
      "[class*='wiki-content']",
      "[class*='doc-content']",
      "[class*='editor-content']",
      "article",
      "[role='main']",
      "main",
      "body"
    ];
  }

  const INVISIBLE_CHARS_RE_LOCAL = /[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g;
  const CTRL_CHARS_RE_LOCAL = /[\u0000-\u001F\u007F-\u009F]/g;
  const EXPORT_TITLE_SUFFIX_RE_LOCAL =
    /\s*[-_–—|]\s*(飞书云文档|飞书文档|Feishu\s*Docs?|Lark\s*Docs?)\s*$/i;
  const NOISY_ATTR_RE_LOCAL =
    /(?:^|[\s_-])(nav|navbar|header|footer|sidebar|sider|toolbar|tool|comment|catalog|outline|search|avatar|profile|popover|modal|dialog|drawer|floating|help|feedback)(?:$|[\s_-])/i;
  const DECORATIVE_IMAGE_URL_RE_LOCAL =
    /(?:\/(?:static|assets?|resource|icons?|emoji|avatar)\b|(?:^|[._-])(icon|logo|avatar|emoji)(?:[._-]|$))/i;
  const DOC_IMAGE_TOKEN_RE_LOCAL = /\/preview\/([a-zA-Z0-9_-]{16,})/i;
  const KNOWN_FEISHU_SUFFIXES_LOCAL = [
    "feishu.cn",
    "feishu.net",
    "larksuite.com",
    "larkoffice.com",
    "larkenterprise.com",
    "feishu-pre.net"
  ];

  function isUuidLikeTokenLocal(value) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
      String(value || "")
    );
  }

  function isLikelyDriveTokenLocal(token) {
    const value = String(token || "").trim();
    if (!/^[a-zA-Z0-9_-]{16,}$/.test(value)) return false;
    if (isUuidLikeTokenLocal(value)) return false;
    if (!/[a-zA-Z]/.test(value)) return false;
    return true;
  }

  function resolveFeishuRootDomainLocal(hostname) {
    const host = String(hostname || "").toLowerCase();
    for (const suffix of KNOWN_FEISHU_SUFFIXES_LOCAL) {
      if (host === suffix || host.endsWith(`.${suffix}`)) {
        return suffix;
      }
    }
    return "";
  }

  function buildTokenPreviewCandidatesLocal(token) {
    const safeToken = sanitizeToken(token);
    if (!isLikelyDriveTokenLocal(safeToken)) return [];
    const currentHost = String(location.hostname || "").toLowerCase();
    const currentOrigin = location.origin;
    const rootDomain = resolveFeishuRootDomainLocal(currentHost);
    const hosts = new Set([currentHost]);
    if (rootDomain) {
      hosts.add(`internal-api-drive-stream.${rootDomain}`);
      hosts.add(`drive-stream.${rootDomain}`);
    } else {
      hosts.add("internal-api-drive-stream.feishu.cn");
      hosts.add("drive-stream.feishu.cn");
    }
    const list = [];
    const add = (value) => {
      const normalized = String(value || "").trim();
      if (normalized) list.push(normalized);
    };
    add(
      `${currentOrigin}/space/api/box/stream/download/preview/${safeToken}?preview_type=16&mount_point=docx_image`
    );
    add(
      `${currentOrigin}/space/api/box/stream/download/preview/${safeToken}?preview_type=16&mount_point=docx_file`
    );
    add(`${currentOrigin}/space/api/box/stream/download/preview/${safeToken}?preview_type=16`);
    add(`${currentOrigin}/space/api/box/stream/download/preview/${safeToken}`);
    for (const host of hosts) {
      add(
        `https://${host}/space/api/box/stream/download/preview/${safeToken}?preview_type=16&mount_point=docx_image`
      );
      add(
        `https://${host}/space/api/box/stream/download/preview/${safeToken}?preview_type=16&mount_point=docx_file`
      );
      add(`https://${host}/space/api/box/stream/download/preview/${safeToken}?preview_type=16`);
      add(`https://${host}/space/api/box/stream/download/preview/${safeToken}`);
    }
    return Array.from(new Set(list));
  }

  function normalizeExportTextLocal(value) {
    return String(value || "")
      .replace(INVISIBLE_CHARS_RE_LOCAL, "")
      .replace(CTRL_CHARS_RE_LOCAL, "")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeExportTitleLocal(value) {
    const text = normalizeExportTextLocal(value).replace(EXPORT_TITLE_SUFFIX_RE_LOCAL, "").trim();
    return text || "未命名文档";
  }

  function buildPreviewUrlByTokenLocal(token) {
    return buildTokenPreviewCandidatesLocal(token)[0] || "";
  }

  function extractTokenFromUrlLocal(url) {
    const value = String(url || "");
    if (!value) return "";
    const fromPreview = DOC_IMAGE_TOKEN_RE_LOCAL.exec(value);
    if (fromPreview?.[1]) {
      const token = sanitizeToken(fromPreview[1]);
      if (isLikelyDriveTokenLocal(token)) return token;
    }
    const fromQuery = /[?&](?:token|image|img|id)=([^&#]+)/i.exec(value);
    if (fromQuery?.[1]) {
      const token = sanitizeToken(decodeURIComponent(fromQuery[1]));
      if (isLikelyDriveTokenLocal(token)) return token;
    }
    return "";
  }

  function getImageTokenFromNodeLocal(node, rawUrl = "") {
    if (!node) return extractTokenFromUrlLocal(rawUrl);
    const tokenAttrs = [
      "data-token",
      "data-resource-token",
      "data-asset-token",
      "data-file-token",
      "data-image-token",
      "data-origin-token",
      "data-attachment-token"
    ];
    for (const attr of tokenAttrs) {
      const value = node.getAttribute?.(attr);
      if (!value) continue;
      const token = sanitizeToken(value);
      if (isLikelyDriveTokenLocal(token)) {
        return token;
      }
    }
    return extractTokenFromUrlLocal(rawUrl);
  }

  function isLikelyDecorativeImageUrlLocal(url) {
    const value = String(url || "");
    if (!value) return true;
    if (/internal-api-drive-stream\.feishu\.cn\/space\/api\/box\/stream\/download\/preview/i.test(value)) {
      return false;
    }
    return DECORATIVE_IMAGE_URL_RE_LOCAL.test(value);
  }

  function getImageRawUrlFromNodeLocal(img) {
    const candidates = getImageSourceCandidatesFromNodeLocal(img);
    return candidates.find((value) => !/^data:/i.test(String(value || ""))) || "";
  }

  function normalizeImageCandidateLocal(rawValue) {
    const value = String(rawValue || "").trim();
    if (!value) return "";
    if (/^data:/i.test(value)) return value;
    const abs = toAbsoluteUrl(value);
    if (/^https?:\/\//i.test(abs) || /^blob:/i.test(abs)) {
      return abs;
    }
    if (/^blob:/i.test(value) || /^https?:\/\//i.test(value)) {
      return value;
    }
    return "";
  }

  function getImageSourceCandidatesFromNodeLocal(img) {
    const list = [];
    const push = (value) => {
      const normalized = normalizeImageCandidateLocal(value);
      if (normalized) {
        list.push(normalized);
      }
    };

    push(img?.currentSrc);
    push(img?.getAttribute?.("src"));
    push(img?.getAttribute?.("data-src"));
    push(img?.getAttribute?.("data-origin-src"));
    push(img?.getAttribute?.("data-original"));
    push(img?.getAttribute?.("data-image-url"));
    push(img?.getAttribute?.("data-lazy-src"));
    push(img?.getAttribute?.("data-download-url"));
    push(img?.getAttribute?.("data-actualsrc"));

    const tokenHint = getImageTokenFromNodeLocal(img, list[0] || "");
    for (const previewUrl of buildTokenPreviewCandidatesLocal(tokenHint)) {
      push(previewUrl);
    }

    return Array.from(new Set(list));
  }

  function findLiveImageForCloneLocal(cloneImg, liveImages = []) {
    if (!cloneImg || !Array.isArray(liveImages) || !liveImages.length) return null;
    const attrKeys = [
      "data-token",
      "data-resource-token",
      "data-asset-token",
      "data-file-token",
      "data-image-token",
      "data-origin-token",
      "data-attachment-token",
      "src",
      "data-src",
      "data-origin-src",
      "data-image-url",
      "alt"
    ];
    for (const key of attrKeys) {
      const value = String(cloneImg.getAttribute?.(key) || "").trim();
      if (!value) continue;
      for (const liveImg of liveImages) {
        const liveValue = String(liveImg.getAttribute?.(key) || "").trim();
        if (liveValue && liveValue === value) {
          return liveImg;
        }
      }
      if (key === "src") {
        const absValue = toAbsoluteUrl(value) || value;
        for (const liveImg of liveImages) {
          const liveSrc = String(
            liveImg.currentSrc ||
            liveImg.getAttribute?.("src") ||
            liveImg.getAttribute?.("data-src") ||
            ""
          ).trim();
          if (!liveSrc) continue;
          const absLive = toAbsoluteUrl(liveSrc) || liveSrc;
          if (absLive === absValue || liveSrc === value) {
            return liveImg;
          }
        }
      }
    }
    return null;
  }

  function collectNodeImageCandidatesLocal(node) {
    const candidates = [];
    if (!node) return candidates;
    const push = (value) => {
      const normalized = normalizeImageCandidateLocal(value);
      if (normalized) {
        candidates.push(normalized);
      }
    };
    if (node.attributes) {
      for (const attr of Array.from(node.attributes)) {
        const name = String(attr?.name || "").toLowerCase();
        const value = String(attr?.value || "");
        if (!value) continue;
        if (
          /(src|url|href|origin|download|image|img|file|thumb|preview|poster)/i.test(name) ||
          /^https?:\/\//i.test(value) ||
          /^blob:/i.test(value) ||
          /\/preview\//i.test(value) ||
          /space\/api\/box\/stream\/download\//i.test(value)
        ) {
          push(value);
        }
      }
    }
    const inlineBg = extractBackgroundUrl(node.getAttribute?.("style") || "");
    if (inlineBg) {
      push(inlineBg);
    }
    return Array.from(new Set(candidates));
  }

  function extractBackgroundUrlFromComputedStyleLocal(node) {
    if (!node) return "";
    let bgImage = "";
    try {
      bgImage = String(window.getComputedStyle(node)?.backgroundImage || "");
    } catch {
      bgImage = "";
    }
    if (!bgImage || bgImage === "none") return "";
    const matched = /url\((['"]?)(.*?)\1\)/i.exec(bgImage);
    return matched ? matched[2] : "";
  }

  function isPossibleImageContainerNodeLocal(node) {
    if (!node || String(node.tagName || "").toLowerCase() === "img") return false;
    const classId = [
      node.getAttribute?.("class") || "",
      node.getAttribute?.("id") || "",
      node.getAttribute?.("data-type") || "",
      node.getAttribute?.("data-block-type") || ""
    ]
      .join(" ")
      .toLowerCase();
    if (!/(image|img|media|photo|picture|attachment)/.test(classId)) {
      return false;
    }
    const textLen = normalizeExportTextLocal(node.textContent || "").length;
    if (textLen > 120) return false;
    try {
      const rect = node.getBoundingClientRect();
      if (Number(rect.width || 0) > 0 && Number(rect.height || 0) > 0) {
        if (rect.width < 24 && rect.height < 24) return false;
      }
    } catch {
      // ignore
    }
    return true;
  }

  function tryInjectBackgroundImagePlaceholdersLocal({
    cloned,
    liveRoot,
    appendImageInfo
  }) {
    if (!cloned || !liveRoot || !appendImageInfo) return 0;
    const selector =
      "[data-block-type*='image'],[data-type*='image'],figure,[class*='image'],[class*='img'],[class*='media']";
    const liveCandidates = Array.from(liveRoot.querySelectorAll(selector))
      .filter((node) => isPossibleImageContainerNodeLocal(node))
      .map((node) => {
        const sourceCandidates = [
          ...collectNodeImageCandidatesLocal(node),
          ...Array.from(node.querySelectorAll("img")).flatMap((img) =>
            collectNodeImageCandidatesLocal(img)
          )
        ];
        const computedBg = extractBackgroundUrlFromComputedStyleLocal(node);
        if (computedBg) {
          sourceCandidates.unshift(computedBg);
        }
        const deduped = Array.from(new Set(sourceCandidates));
        const rawUrl = deduped.find((value) => !/^data:/i.test(value)) || "";
        if (!rawUrl) return null;
        return {
          rawUrl,
          tokenHint: getImageTokenFromNodeLocal(node, rawUrl),
          altText:
            node.getAttribute?.("aria-label") ||
            node.getAttribute?.("title") ||
            node.getAttribute?.("alt") ||
            "embedded-image",
          sourceCandidates: deduped
        };
      })
      .filter(Boolean)
      .slice(0, 120);

    if (!liveCandidates.length) return 0;

    const cloneCandidates = Array.from(cloned.querySelectorAll(selector)).slice(0, 200);
    if (!cloneCandidates.length) return 0;

    let inserted = 0;
    const usedToken = new Set();
    for (let i = 0; i < Math.min(liveCandidates.length, cloneCandidates.length); i += 1) {
      const liveItem = liveCandidates[i];
      const cloneNode = cloneCandidates[i];
      if (!cloneNode || cloneNode.querySelector("img[data-token-placeholder='true']")) continue;
      if (cloneNode.querySelector("img")) continue;
      const imageMeta = appendImageInfo(
        liveItem.rawUrl,
        liveItem.altText,
        liveItem.tokenHint,
        false,
        liveItem.sourceCandidates
      );
      if (!imageMeta || usedToken.has(imageMeta.token)) continue;
      usedToken.add(imageMeta.token);
      const placeholder = document.createElement("img");
      placeholder.setAttribute("data-token", imageMeta.token);
      placeholder.setAttribute("data-name", imageMeta.name);
      placeholder.setAttribute("data-token-placeholder", "true");
      placeholder.setAttribute("src", "");
      placeholder.setAttribute("alt", imageMeta.name);
      cloneNode.appendChild(placeholder);
      inserted += 1;
    }
    return inserted;
  }

  function shouldDropNodeByAttrsLocal(node) {
    if (!node) return false;
    const tag = String(node.tagName || "").toLowerCase();
    if (
      tag === "script" ||
      tag === "style" ||
      tag === "noscript" ||
      tag === "template" ||
      tag === "button" ||
      tag === "input" ||
      tag === "textarea" ||
      tag === "select" ||
      tag === "form" ||
      tag === "dialog"
    ) {
      return true;
    }
    if (node.getAttribute?.("aria-hidden") === "true") {
      return true;
    }
    const textLength = normalizeExportTextLocal(node.textContent || "").length;
    if (textLength > 300) {
      return false;
    }
    try {
      const structuredCount = node.querySelectorAll(
        "p,li,h1,h2,h3,h4,h5,h6,table,pre,blockquote"
      ).length;
      if (structuredCount >= 4) {
        return false;
      }
    } catch {
      // ignore
    }
    const joined = [
      node.getAttribute?.("class") || "",
      node.getAttribute?.("id") || "",
      node.getAttribute?.("data-testid") || "",
      node.getAttribute?.("data-test-id") || "",
      node.getAttribute?.("data-qa") || ""
    ]
      .join(" ")
      .toLowerCase();
    return NOISY_ATTR_RE_LOCAL.test(joined);
  }

  function pruneNonContentNodesLocal(root) {
    if (!root?.querySelectorAll) return;
    const structuralNoise = root.querySelectorAll(
      "header,footer,nav,aside,[role='navigation'],[role='banner'],[role='search'],[role='dialog'],[aria-modal='true']"
    );
    for (const node of structuralNoise) {
      node.remove();
    }
    const allNodes = root.querySelectorAll("*");
    for (const node of allNodes) {
      if (shouldDropNodeByAttrsLocal(node)) {
        node.remove();
      }
    }
  }

  function scoreNodeForDocumentLocal(node) {
    if (!node?.querySelectorAll) return Number.NEGATIVE_INFINITY;
    const text = normalizeExportTextLocal(node.textContent || "");
    const textLength = text.length;
    if (textLength < 80) return Number.NEGATIVE_INFINITY;

    const paragraphs = node.querySelectorAll("p").length;
    const listItems = node.querySelectorAll("li").length;
    const headings = node.querySelectorAll("h1,h2,h3,h4,h5,h6").length;
    const images = node.querySelectorAll("img").length;
    const links = node.querySelectorAll("a[href]").length;
    const controls = node.querySelectorAll("button,input,select,textarea,[role='button']").length;
    const attr = `${node.getAttribute?.("class") || ""} ${node.getAttribute?.("id") || ""}`.toLowerCase();

    if (textLength < 160 && paragraphs + listItems + headings < 3) {
      return Number.NEGATIVE_INFINITY;
    }

    let score = textLength * 2 + paragraphs * 280 + listItems * 180 + headings * 360 + images * 40;
    score -= links * 10;
    score -= controls * 120;
    if (/wiki|doc|editor|content|article|block|canvas/.test(attr)) {
      score += 900;
    }
    if (/nav|toolbar|sidebar|header|footer|comment|catalog/.test(attr)) {
      score -= 1200;
    }
    return score;
  }

  function selectBestMainContentNodeLocal() {
    const candidates = [];
    const seen = new Set();
    for (const selector of getMainContainerSelectorsLocal()) {
      const nodes = Array.from(document.querySelectorAll(selector));
      for (const node of nodes) {
        if (!node || seen.has(node)) continue;
        seen.add(node);
        const score = scoreNodeForDocumentLocal(node);
        if (!Number.isFinite(score)) continue;
        candidates.push({ node, score });
      }
    }
    if (!candidates.length) {
      return document.body || null;
    }
    candidates.sort((a, b) => b.score - a.score);
    const best = candidates[0]?.node || null;
    if (!best) return document.body || null;
    if (String(best.tagName || "").toLowerCase() === "body" && candidates[1]?.node) {
      return candidates[1].node;
    }
    return best;
  }

  function cloneMainContentNodeLocal() {
    const node = selectBestMainContentNodeLocal();
    if (!node) return null;
    const cloned = node.cloneNode(true);
    pruneNonContentNodesLocal(cloned);
    const prunedTextLength = normalizeExportTextLocal(cloned.textContent || "").length;
    if (prunedTextLength >= 80) {
      return cloned;
    }
    const fallbackClone = node.cloneNode(true);
    const fallbackTextLength = normalizeExportTextLocal(fallbackClone.textContent || "").length;
    if (fallbackTextLength >= 80) {
      return fallbackClone;
    }
    if (document.body && node !== document.body) {
      const bodyClone = document.body.cloneNode(true);
      pruneNonContentNodesLocal(bodyClone);
      if (normalizeExportTextLocal(bodyClone.textContent || "").length >= 80) {
        return bodyClone;
      }
      return document.body.cloneNode(true);
    }
    return cloned;
  }

  function extractUrlFromAttributesLocal(node, keys) {
    for (const key of keys) {
      const value = node.getAttribute?.(key);
      if (!value) continue;
      const abs = toAbsoluteUrl(value);
      if (/^https?:\/\//i.test(abs)) {
        return abs;
      }
    }
    return "";
  }

  function buildLinkCardNodeLocal(url, titleText = "") {
    const card = document.createElement("div");
    card.setAttribute("data-export-link-card", "true");
    card.style.border = "1px solid #d7deea";
    card.style.borderRadius = "10px";
    card.style.padding = "10px 12px";
    card.style.margin = "10px 0";
    card.style.background = "#f8fafc";

    if (titleText) {
      const title = document.createElement("div");
      title.textContent = titleText;
      title.style.marginBottom = "6px";
      title.style.color = "#0f172a";
      title.style.fontWeight = "600";
      card.appendChild(title);
    }

    const link = document.createElement("a");
    link.href = url;
    link.textContent = url;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.style.color = "#1d4ed8";
    link.style.wordBreak = "break-all";
    card.appendChild(link);
    return card;
  }

  function normalizeEmbedLinksLocal(root) {
    if (!root?.querySelectorAll) return;

    const iframes = Array.from(root.querySelectorAll("iframe"));
    for (const iframe of iframes) {
      const url = extractUrlFromAttributesLocal(iframe, [
        "src",
        "data-src",
        "data-origin-src",
        "data-url"
      ]);
      if (!url) {
        iframe.remove();
        continue;
      }
      const title = (iframe.getAttribute("title") || "").trim();
      iframe.replaceWith(buildLinkCardNodeLocal(url, title || "嵌入链接"));
    }

    const possibleCards = Array.from(
      root.querySelectorAll("[data-href],[data-url],[data-link],[data-link-url],[data-redirect-url]")
    );
    for (const node of possibleCards) {
      if (!node || node.closest?.("a[href]")) continue;
      if (node.querySelector?.("a[href]")) continue;
      const url = extractUrlFromAttributesLocal(node, [
        "data-href",
        "data-url",
        "data-link",
        "data-link-url",
        "data-redirect-url"
      ]);
      if (!url) continue;
      const text = (node.textContent || "").trim().replace(/\s+/g, " ");
      const link = document.createElement("a");
      link.href = url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.style.color = "#1d4ed8";
      link.style.wordBreak = "break-all";
      link.textContent = text || url;
      node.textContent = "";
      node.appendChild(link);
    }
  }

  function collectDocumentPayloadLocally() {
    const title = normalizeExportTitleLocal(document.title || "未命名文档");
    const currentUrl = location.href;
    const cloned = cloneMainContentNodeLocal();
    const liveRoot = selectBestMainContentNodeLocal();
    if (!cloned) {
      throw new Error("未找到文档主体区域");
    }

    cloned.querySelectorAll("script,style,noscript,template").forEach((node) => node.remove());
    normalizeEmbedLinksLocal(cloned);

    const imageInfoList = [];
    const tokenSet = new Set();
    let imageIndex = 0;

    const appendImageInfo = (
      rawUrl,
      altText,
      tokenHint = "",
      allowDecorative = false,
      sourceCandidates = []
    ) => {
      let absUrl = toAbsoluteUrl(rawUrl);
      const hintedToken = sanitizeToken(tokenHint || extractTokenFromUrlLocal(absUrl || rawUrl));
      if (!absUrl && Array.isArray(sourceCandidates)) {
        absUrl =
          sourceCandidates
            .map((value) => normalizeImageCandidateLocal(value))
            .find((value) => value && !/^data:/i.test(value)) || "";
      }
      if (!absUrl && hintedToken && hintedToken !== "img") {
        absUrl = buildPreviewUrlByTokenLocal(hintedToken);
      }
      if (!absUrl || /^data:/i.test(absUrl)) return null;
      if (!allowDecorative && isLikelyDecorativeImageUrlLocal(absUrl)) return null;
      const token =
        hintedToken && hintedToken !== "img" ? hintedToken : tokenFromUrl(absUrl, imageIndex);
      imageIndex += 1;
      if (!tokenSet.has(token)) {
        tokenSet.add(token);
        const normalizedCandidates = Array.from(
          new Set(
            [
              ...sourceCandidates,
              absUrl,
              ...(hintedToken && hintedToken !== "img"
                ? buildTokenPreviewCandidatesLocal(hintedToken)
                : [])
            ]
              .map((value) => normalizeImageCandidateLocal(value))
              .filter((value) => value && !/^data:/i.test(value))
          )
        );
        imageInfoList.push({
          token,
          url: absUrl,
          originalUrl: String(rawUrl || ""),
          mimeType: guessMimeType(absUrl),
          sourceCandidates: normalizedCandidates
        });
      }
      return {
        token,
        name: altText || token
      };
    };

    const images = Array.from(cloned.querySelectorAll("img"));
    const liveImages = Array.from(document.querySelectorAll("img"));
    for (const img of images) {
      const liveImg = findLiveImageForCloneLocal(img, liveImages);
      const sourceImg = liveImg || img;
      const sourceCandidates = getImageSourceCandidatesFromNodeLocal(sourceImg);
      const rawUrl = sourceCandidates[0] || getImageRawUrlFromNodeLocal(sourceImg);
      const tokenHint = getImageTokenFromNodeLocal(sourceImg, rawUrl);
      const imageMeta = appendImageInfo(
        rawUrl,
        img.getAttribute("alt") || "",
        tokenHint,
        false,
        sourceCandidates
      );
      if (!imageMeta) continue;
      img.setAttribute("data-token", imageMeta.token);
      img.setAttribute("data-name", imageMeta.name);
      img.setAttribute("data-token-placeholder", "true");
      img.setAttribute("src", "");
      img.removeAttribute("srcset");
      img.removeAttribute("loading");
    }

    const imageLikeNodes = Array.from(
      cloned.querySelectorAll("[style*='background'],[data-src],[data-origin-src],[data-image-url]")
    );
    for (const node of imageLikeNodes) {
      if (!node || String(node.tagName || "").toLowerCase() === "img") continue;
      const styleValue = node.getAttribute("style") || "";
      const bgUrl = extractBackgroundUrl(styleValue);
      const rawUrl =
        bgUrl ||
        node.getAttribute("data-src") ||
        node.getAttribute("data-origin-src") ||
        node.getAttribute("data-image-url") ||
        "";
      if (!rawUrl) continue;
      const tokenHint = getImageTokenFromNodeLocal(node, rawUrl);
      const sourceCandidates = collectNodeImageCandidatesLocal(node);
      if (bgUrl) {
        sourceCandidates.unshift(bgUrl);
      }
      const imageMeta = appendImageInfo(
        rawUrl,
        "embedded-image",
        tokenHint,
        false,
        Array.from(new Set(sourceCandidates))
      );
      if (!imageMeta) continue;

      if (bgUrl) {
        node.setAttribute(
          "style",
          styleValue.replace(
            /background(?:-image)?\s*:\s*[^;]*url\((['"]?)(.*?)\1\)/gi,
            "background-image:none"
          )
        );
      }

      const placeholder = document.createElement("img");
      placeholder.setAttribute("data-token", imageMeta.token);
      placeholder.setAttribute("data-name", imageMeta.name);
      placeholder.setAttribute("data-token-placeholder", "true");
      placeholder.setAttribute("src", "");
      placeholder.setAttribute("alt", imageMeta.name);
      node.appendChild(placeholder);
    }

    const injectedInPlaceCount = tryInjectBackgroundImagePlaceholdersLocal({
      cloned,
      liveRoot,
      appendImageInfo
    });

    if (imageInfoList.length < 4) {
      const supplementContainer = document.createElement("section");
      supplementContainer.setAttribute("data-export-fallback-images", "true");
      supplementContainer.style.marginTop = "16px";
      const heading = document.createElement("h3");
      heading.textContent = "补采集图片";
      supplementContainer.appendChild(heading);
      const addedTokenInContainer = new Set();

      const addSupplementImage = (rawUrl, tokenHint = "", altText = "supplement-image") => {
        const imageMeta = appendImageInfo(rawUrl, altText, tokenHint, true);
        if (!imageMeta) return;
        if (addedTokenInContainer.has(imageMeta.token)) return;
        addedTokenInContainer.add(imageMeta.token);
        const img = document.createElement("img");
        img.setAttribute("data-token", imageMeta.token);
        img.setAttribute("data-name", imageMeta.name);
        img.setAttribute("data-token-placeholder", "true");
        img.setAttribute("src", "");
        img.setAttribute("alt", imageMeta.name);
        supplementContainer.appendChild(img);
      };

      try {
        const globalImages = document.querySelectorAll("img");
        for (const img of globalImages) {
          const width = Number(img?.naturalWidth || img?.width || 0);
          const height = Number(img?.naturalHeight || img?.height || 0);
          if (width > 0 && height > 0 && width < 48 && height < 48) continue;
          const rawUrl =
            img?.currentSrc ||
            img?.getAttribute?.("src") ||
            img?.getAttribute?.("data-src") ||
            img?.getAttribute?.("data-origin-src") ||
            img?.getAttribute?.("data-image-url") ||
            "";
          const tokenHint = getImageTokenFromNodeLocal(img, rawUrl);
          addSupplementImage(rawUrl, tokenHint, img?.getAttribute?.("alt") || "img");
          if (addedTokenInContainer.size >= 40) break;
        }
      } catch {
        // ignore
      }

      if (addedTokenInContainer.size < 40) {
        const html = String(document.documentElement?.innerHTML || "");
        const previewTokenRegex = /\/preview\/([a-zA-Z0-9_-]{16,})/g;
        let previewMatch = previewTokenRegex.exec(html);
        while (previewMatch && addedTokenInContainer.size < 45) {
          addSupplementImage("", previewMatch[1], previewMatch[1]);
          previewMatch = previewTokenRegex.exec(html);
        }
        const encodedPreviewTokenRegex = /preview%2F([a-zA-Z0-9_-]{16,})/gi;
        let encodedPreviewMatch = encodedPreviewTokenRegex.exec(html);
        while (encodedPreviewMatch && addedTokenInContainer.size < 50) {
          addSupplementImage("", encodedPreviewMatch[1], encodedPreviewMatch[1]);
          encodedPreviewMatch = encodedPreviewTokenRegex.exec(html);
        }
      }

      if (addedTokenInContainer.size > 0) {
        cloned.appendChild(supplementContainer);
      }
    }

    const htmlWithPlaceholders = cloned.outerHTML || "";
    return {
      title,
      currentUrl,
      htmlContent: htmlWithPlaceholders,
      htmlWithPlaceholders,
      imageInfoList,
      imageTokens: Array.from(tokenSet),
      htmlBlobSize: new Blob([htmlWithPlaceholders]).size,
      supplementedImageCount: injectedInPlaceCount,
      timestamp: Date.now()
    };
  }

  function safeStringLocal(value) {
    if (typeof value !== "string") return "";
    return value.trim();
  }

  function parseBitableUrlContextLocal(url) {
    const fallback = {
      appToken: "",
      tableId: "",
      viewId: "",
      shareViewToken: "",
      host: ""
    };
    try {
      const parsed = new URL(url);
      const segments = parsed.pathname.split("/").filter(Boolean);
      const reserved = new Set([
        "view",
        "table",
        "form",
        "dashboard",
        "share",
        "embed"
      ]);

      let appToken = "";
      let shareViewToken = "";

      for (let index = 0; index < segments.length; index += 1) {
        const current = (segments[index] || "").toLowerCase();
        const next = segments[index + 1] || "";
        const nextLower = next.toLowerCase();

        if ((current === "base" || current === "bitable") && next) {
          if (nextLower === "view" && segments[index + 2]) {
            shareViewToken = segments[index + 2];
          } else if (!reserved.has(nextLower)) {
            appToken = next;
          }
        }

        if (current === "view" && !shareViewToken && segments[index + 1]) {
          shareViewToken = segments[index + 1];
        }
      }

      const tableId =
        parsed.searchParams.get("table") || parsed.searchParams.get("tableId") || "";
      const queryViewId =
        parsed.searchParams.get("view") || parsed.searchParams.get("viewId") || "";
      const viewId = queryViewId || shareViewToken;

      return {
        appToken,
        tableId,
        viewId,
        shareViewToken,
        host: parsed.host
      };
    } catch {
      return fallback;
    }
  }

  function getTableRowsFromDomLocal() {
    const table = document.querySelector("table");
    if (table) {
      const headerCells = Array.from(table.querySelectorAll("thead th"));
      const headers =
        headerCells.length > 0
          ? headerCells.map(
            (th, index) => (th.textContent || "").trim() || `column_${index + 1}`
          )
          : Array.from(table.querySelectorAll("tr:first-child th, tr:first-child td")).map(
            (cell, index) => (cell.textContent || "").trim() || `column_${index + 1}`
          );
      const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) =>
        Array.from(tr.children).map((cell) => (cell.textContent || "").trim())
      );
      return { headers, rows };
    }

    const grid =
      document.querySelector("[role='grid']") || document.querySelector("[data-sheet-container]");
    if (!grid) {
      return { headers: [], rows: [] };
    }

    const headers = Array.from(grid.querySelectorAll("[role='columnheader']"))
      .map((cell, index) => (cell.textContent || "").trim() || `column_${index + 1}`)
      .filter(Boolean);

    const rows = Array.from(grid.querySelectorAll("[role='row']"))
      .map((rowNode) =>
        Array.from(rowNode.querySelectorAll("[role='gridcell']"))
          .map((cell) => (cell.textContent || "").trim())
          .filter((cellText, cellIndex) => cellText || headers[cellIndex])
      )
      .filter((row) => row.length > 0);

    if (!headers.length && rows.length > 0) {
      const generatedHeaders = rows[0].map((_, index) => `column_${index + 1}`);
      return { headers: generatedHeaders, rows };
    }
    return { headers, rows };
  }

  function collectBitablePayloadLocally() {
    const urlContext = parseBitableUrlContextLocal(location.href);
    const tableName = document.title || "Bitable";
    const { headers, rows } = getTableRowsFromDomLocal();

    if (!headers.length && !rows.length) {
      throw new Error("未从页面 DOM 读取到多维表数据，请确认表格内容已加载后重试");
    }

    const records = rows.map((row, index) => ({
      recordId: `row_${index + 1}`,
      fields: headers.reduce((acc, key, keyIndex) => {
        acc[key] = row[keyIndex] ?? "";
        return acc;
      }, {})
    }));

    return {
      title: tableName,
      currentUrl: location.href,
      bitableData: {
        meta: {
          appToken: safeStringLocal(urlContext.appToken),
          tableId: safeStringLocal(urlContext.tableId),
          viewId: safeStringLocal(urlContext.viewId),
          shareViewToken: safeStringLocal(urlContext.shareViewToken),
          host: safeStringLocal(urlContext.host),
          source: "content-dom-fallback",
          exportedAt: Date.now(),
          currentUrl: location.href
        },
        tables: [
          {
            tableId: urlContext.tableId || "visible-table",
            tableName,
            fields: headers.map((name) => ({
              id: sanitizeToken(name),
              name,
              type: "text"
            })),
            records
          }
        ]
      },
      htmlContent: "<div>Bitable Export</div>",
      htmlWithPlaceholders: "<div>Bitable Export</div>",
      imageInfoList: [],
      imageTokens: [],
      htmlBlobSize: 24,
      timestamp: Date.now()
    };
  }

  async function runFallbackExport(request) {
    if (!isExporting || !request) {
      return;
    }
    try {
      emitProgress({
        title: "页面抓取",
        message: "注入脚本未响应，启用兜底抓取模式",
        status: STATUS.WARNING
      });
      const exportType = String(request.exportType || "").toLowerCase();
      const payload =
        isBitableExportType(exportType)
          ? collectBitablePayloadLocally()
          : collectDocumentPayloadLocally();
      await handleExportDataMessage({
        requestId: request.requestId,
        exportType,
        options: request.options || {},
        payload
      });
    } catch (error) {
      emitProgress({
        title: "导出任务",
        message: error instanceof Error ? error.message : String(error),
        status: STATUS.ERROR,
        taskState: "error"
      });
    } finally {
      clearExportWatchdog();
      activeExportRequest = null;
      isExporting = false;
      clearTask();
    }
  }

  function decodeDataUrlToBlobLocal(dataUrl, fallbackMimeType = "application/octet-stream") {
    const value = String(dataUrl || "");
    const match = /^data:([^;,]+)?(?:;charset=[^;,]+)?(?:;(base64))?,([\s\S]*)$/i.exec(value);
    if (!match) {
      throw new Error("Invalid data URL");
    }
    const mimeType = match[1] || fallbackMimeType;
    const isBase64 = Boolean(match[2]);
    const payload = match[3] || "";
    if (isBase64) {
      const binary = atob(payload);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) {
        bytes[i] = binary.charCodeAt(i);
      }
      return new Blob([bytes], { type: mimeType });
    }
    return new Blob([decodeURIComponent(payload)], { type: mimeType });
  }

  function buildImageFetchCandidates(imageInfo = {}) {
    const token = String(imageInfo?.token || "").trim();
    const candidates = [
      ...(Array.isArray(imageInfo?.sourceCandidates) ? imageInfo.sourceCandidates : []),
      imageInfo?.url,
      imageInfo?.originalUrl,
      imageInfo?.sourceUrl,
      imageInfo?.rawUrl
    ]
      .map((value) => String(value || "").trim())
      .filter(Boolean);
    if (isLikelyDriveTokenLocal(token)) {
      candidates.push(...buildTokenPreviewCandidatesLocal(token));
    }
    return Array.from(new Set(candidates)).slice(0, 12);
  }

  async function fetchWithTimeoutLocal(url, options = {}, timeoutMs = IMAGE_FETCH_TIMEOUT_MS) {
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

  async function fetchImageBlobByUrl(url) {
    let response;
    try {
      response = await fetchWithTimeoutLocal(
        url,
        {
          credentials: "include",
          cache: "force-cache"
        },
        IMAGE_FETCH_TIMEOUT_MS
      );
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        throw new Error("timeout");
      }
      throw error;
    }
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return response.blob();
  }

  function requestImageDataViaInjected({ candidates, token, timeoutMs = 15000 }) {
    const requestId = `image-fetch-${crypto.randomUUID()}`;
    return new Promise(async (resolve, reject) => {
      let timer = null;
      try {
        await ensureInjectedScript();
        timer = setTimeout(() => {
          pendingInjectedImageRequests.delete(requestId);
          reject(new Error("页面上下文取图超时"));
        }, timeoutMs);
        pendingInjectedImageRequests.set(requestId, {
          resolve: (data) => {
            clearTimeout(timer);
            resolve(data);
          },
          reject: (error) => {
            clearTimeout(timer);
            reject(error);
          }
        });
        postToInjected({
          source: EXTENSION_BRIDGE_SOURCE,
          type: "FETCH_IMAGE_FROM_PAGE",
          requestId,
          token: token || "",
          candidates: Array.isArray(candidates) ? candidates : []
        });
      } catch (error) {
        if (timer) clearTimeout(timer);
        pendingInjectedImageRequests.delete(requestId);
        reject(error);
      }
    });
  }

  async function fetchImageBlobWithRetry(imageInfo, maxRetry, candidatesInput = []) {
    const candidates =
      Array.isArray(candidatesInput) && candidatesInput.length > 0
        ? candidatesInput
        : buildImageFetchCandidates(imageInfo);
    const retryCount = Math.max(1, Math.min(2, Number(maxRetry) || 1));
    let lastError = null;

    for (const candidate of candidates) {
      for (let attempt = 1; attempt <= retryCount; attempt += 1) {
        try {
          return await fetchImageBlobByUrl(candidate);
        } catch (error) {
          lastError = error;
          if (attempt < retryCount) {
            await sleep(200 * 2 ** (attempt - 1));
          }
        }
      }
    }

    try {
      const injectedData = await requestImageDataViaInjected({
        candidates,
        token: imageInfo?.token || "",
        timeoutMs: 60000
      });
      const dataUrl = String(injectedData?.dataUrl || "");
      if (dataUrl.startsWith("data:")) {
        return decodeDataUrlToBlobLocal(
          dataUrl,
          String(injectedData?.mimeType || imageInfo?.mimeType || "application/octet-stream")
        );
      }
    } catch (error) {
      lastError = error;
    }

    throw lastError || new Error("Image fetch failed");
  }

  function shouldUseBackgroundImageProxy(candidates) {
    return (Array.isArray(candidates) ? candidates : []).some((value) =>
      /^https?:\/\//i.test(String(value || ""))
    );
  }

  async function storeImageViaBackgroundProxy({
    token,
    currentUrl,
    imageInfo,
    candidates
  }) {
    if (!shouldUseBackgroundImageProxy(candidates)) {
      return false;
    }
    const response = await sendChunkedMessage({
      action: "store-image-by-candidates",
      taskId: resolveCurrentTaskId(),
      data: {
        id: buildImageRecordId(currentUrl, token),
        token,
        documentUrl: currentUrl,
        candidates,
        mimeType: String(imageInfo?.mimeType || "").trim(),
        taskId: resolveCurrentTaskId()
      }
    });
    return Boolean(response?.ok);
  }

  function convertBlobByWorker({ token, blob, batchIndex, imageIndex }) {
    return new Promise((resolve, reject) => {
      const imageWorker = getWorker();
      const jobId = `${token}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const timer = setTimeout(() => {
        pendingWorkerJobs.delete(jobId);
        reject(new Error("Image worker timeout"));
      }, 20000);
      pendingWorkerJobs.set(jobId, { resolve, reject, timer });
      imageWorker.postMessage({
        action: "convert-blob",
        token,
        blob,
        batchIndex,
        imageIndex,
        jobId
      });
    });
  }

  async function runWithConcurrency(items, limit, runner) {
    let index = 0;
    const workers = new Array(Math.min(limit, items.length)).fill(0).map(async () => {
      while (true) {
        const current = index;
        index += 1;
        if (current >= items.length) {
          return;
        }
        await runner(items[current], current);
      }
    });
    await Promise.all(workers);
  }

  async function processImagesAndStore(payload) {
    const imageList = payload.imageInfoList || [];
    const total = imageList.length;

    if (!total) {
      return [];
    }

    const successfulTokens = [];
    let processed = 0;
    let failed = 0;

    emitProgress({
      title: "图片处理",
      message: `开始处理图片 0/${total}`,
      status: STATUS.RUNNING,
      taskState: "collecting"
    });

    for (
      let batchStart = 0;
      batchStart < imageList.length;
      batchStart += IMAGE_BATCH_SIZE
    ) {
      const batch = imageList.slice(batchStart, batchStart + IMAGE_BATCH_SIZE);

      await runWithConcurrency(batch, IMAGE_CONCURRENCY, async (imageInfo, offset) => {
        const index = batchStart + offset;
        const token = imageInfo.token || `img-${index + 1}`;
        const candidates = buildImageFetchCandidates(imageInfo);
        try {
          let storedByProxy = false;
          try {
            storedByProxy = await storeImageViaBackgroundProxy({
              token,
              currentUrl: payload.currentUrl,
              imageInfo,
              candidates
            });
          } catch {
            storedByProxy = false;
          }
          if (storedByProxy) {
            successfulTokens.push(token);
            return;
          }

          const blob = await fetchImageBlobWithRetry(imageInfo, IMAGE_MAX_RETRY, candidates);
          const converted = await convertBlobByWorker({
            token,
            blob,
            batchIndex: Math.floor(index / IMAGE_BATCH_SIZE),
            imageIndex: index % IMAGE_BATCH_SIZE
          });
          const base64Data =
            converted.base64Data ||
            (typeof converted.dataUrl === "string"
              ? String(converted.dataUrl).replace(/^data:[^;]+;base64,/, "")
              : "");
          if (!base64Data) {
            throw new Error("图片转换结果为空");
          }

          await sendChunkedMessage({
            action: "store-image",
            taskId: resolveCurrentTaskId(),
            data: {
              id: buildImageRecordId(payload.currentUrl, token),
              token,
              base64Data,
              mimeType: converted.mimeType || blob.type || "application/octet-stream",
              documentUrl: payload.currentUrl,
              taskId: resolveCurrentTaskId()
            }
          });
          successfulTokens.push(token);
        } catch (error) {
          failed += 1;
          if (failed <= 3) {
            emitProgress({
              title: "图片处理",
              message: `图片拉取失败(${token})：${error instanceof Error ? error.message : String(error)
                }`,
              status: STATUS.WARNING,
              taskState: "collecting"
            });
          }
        } finally {
          processed += 1;
          emitProgress({
            title: "图片处理",
            message: `已处理 ${processed}/${total}（成功 ${successfulTokens.length}，失败 ${failed}）`,
            status: failed > 0 && processed === total ? STATUS.WARNING : STATUS.RUNNING,
            taskState: "collecting"
          });
        }
      });
    }

    return successfulTokens;
  }

  async function handleExportDataMessage(message) {
    let payload = message.payload || {};
    const exportType = String(message.exportType || "").toLowerCase();
    const taskId = resolveCurrentTaskId(message.taskId || message.requestId);
    const options = { ...(message.options || {}) };
    const isBitableExport = isBitableExportType(exportType);

    if (isBitableExport) {
      const hasTables =
        Array.isArray(payload?.bitableData?.tables) && payload.bitableData.tables.length > 0;
      if (!hasTables) {
        emitProgress({
          title: "Bitable 导出",
          message: "注入结果未返回表格数据，切换本地抓取兜底",
          status: STATUS.WARNING,
          taskId,
          taskState: "collecting"
        });
        const localPayload = collectBitablePayloadLocally();
        payload = {
          ...localPayload,
          ...payload,
          bitableData: localPayload.bitableData,
          title: payload.title || localPayload.title,
          currentUrl: payload.currentUrl || payload.url || localPayload.currentUrl
        };
      }
    }

    const htmlValue = payload.htmlWithPlaceholders || payload.htmlContent || "";
    const htmlBlobSize = Number(payload.htmlBlobSize || new Blob([htmlValue]).size);
    const currentUrl = payload.currentUrl || payload.url || location.href;

    if (!payload?.title || !currentUrl) {
      throw new Error("页面数据不完整");
    }

    if (!isBitableExport && !htmlValue) {
      throw new Error("未获取到可导出的文档 HTML");
    }

    if (htmlBlobSize > LARGE_DOCUMENT_THRESHOLD) {
      if (exportType === "markdown") {
        options.zipImages = true;
        options.bigMode = true;
      }
      emitProgress({
        title: "导出任务",
        message: "检测到大文档，启用大文档模式",
        status: STATUS.WARNING,
        taskId,
        taskState: "collecting"
      });
    }

    const imageTokens = isBitableExport
      ? []
      : await processImagesAndStore({
        ...payload,
        currentUrl
      });

    if (isBitableExport) {
      emitProgress({
        title: "Bitable 导出",
        message: "数据抓取完成，开始组装导出文件",
        status: STATUS.RUNNING,
        taskId,
        taskState: "assembling"
      });
    } else {
      emitProgress({
        title: "导出任务",
        message: "图片阶段完成，开始组装导出文件",
        status: STATUS.RUNNING,
        taskId,
        taskState: "assembling"
      });
    }

    await sendChunkedMessage({
      action: "assemble-export",
      taskId,
      data: {
        title: payload.title,
        currentUrl,
        htmlWithPlaceholders: htmlValue,
        imageTokens,
        exportType,
        options,
        htmlBlobSize,
        bitableData: payload.bitableData || null,
        taskId
      }
    });
  }

  window.addEventListener("message", async (event) => {
    if (event.source !== window) return;
    const message = event.data || {};

    if (
      message.source === INJECTED_BRIDGE_SOURCE &&
      message.type === "FETCH_IMAGE_FROM_PAGE_RESULT" &&
      message.requestId
    ) {
      const pending = pendingInjectedImageRequests.get(message.requestId);
      if (!pending) return;
      pendingInjectedImageRequests.delete(message.requestId);
      if (message.ok && message.data?.dataUrl) {
        pending.resolve(message.data);
      } else {
        pending.reject(new Error(message.error || "页面上下文取图失败"));
      }
      return;
    }

    if (message.type === "NV_MARKDOWN_EXPORT_SUCCESS") {
      try {
        await handleReferenceMarkdownSuccess(message);
      } catch (error) {
        emitProgress({
          title: "Markdown 导出",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskState: "error"
        });
        clearExportWatchdog();
        activeExportRequest = null;
        isExporting = false;
        clearTask();
      }
      return;
    }

    if (message.action === "NV_DOWNLOAD_HTML") {
      try {
        await handleReferenceHtmlDownload(message);
      } catch (error) {
        emitProgress({
          title: "HTML 导出",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskState: "error"
        });
        clearExportWatchdog();
        activeExportRequest = null;
        isExporting = false;
        clearTask();
      }
      return;
    }

    if (message.type === "NV_MARKDOWN_EXPORT_FAIL") {
      const failMsg =
        message?.data?.message ||
        message?.error ||
        message?.message ||
        "Markdown 导出失败";
      emitProgress({
        title: "Markdown 导出",
        message: failMsg,
        status: STATUS.ERROR,
        taskState: "error"
      });
      clearExportWatchdog();
      activeExportRequest = null;
      isExporting = false;
      clearTask();
      return;
    }

    if (message.type === "NV_IMAGE_PROCESSED") {
      try {
        await handleReferencePdfImageProcessed(message);
      } catch (error) {
        emitProgress({
          title: "PDF 图片处理",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.WARNING
        });
      }
      return;
    }

    if (message.type === "NV_ALL_IMAGES_PROCESSED") {
      try {
        await handleReferencePdfAllImagesProcessed(message);
      } catch (error) {
        emitProgress({
          title: "PDF 导出",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskState: "error"
        });
        clearExportWatchdog();
        activeExportRequest = null;
        isExporting = false;
        clearTask();
      }
      return;
    }

    if (message.type === MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE && message.data) {
      emitProgress(message.data, "injected");
      return;
    }

    if (message.source !== INJECTED_BRIDGE_SOURCE) return;

    if (message.type === "PAGE_INFO") {
      const resolver = pendingPageInfoRequests.get(message.requestId);
      if (resolver) {
        pendingPageInfoRequests.delete(message.requestId);
        resolver(message.pageInfo || {});
      }
      return;
    }

    if (message.type === MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE) {
      const progress = message.data || message.progress;
      if (progress) {
        emitProgress(progress, "injected");
      }
      return;
    }

    if (message.type === "EXPORT_ERROR") {
      const incomingTaskId = String(message.taskId || message.requestId || "");
      const currentTaskId = resolveCurrentTaskId();
      if (incomingTaskId && currentTaskId && incomingTaskId !== currentTaskId) {
        return;
      }
      if (
        activeExportRequest &&
        message.requestId &&
        message.requestId !== activeExportRequest.requestId
      ) {
        return;
      }
      clearExportWatchdog();
      emitProgress({
        title: "导出任务",
        message: message.error || "页面抓取失败",
        status: STATUS.ERROR,
        taskId: incomingTaskId || currentTaskId,
        taskState: "error"
      });
      activeExportRequest = null;
      isExporting = false;
      clearTask();
      return;
    }

    if (message.type === "EXPORT_DATA") {
      const incomingTaskId = String(message.taskId || message.requestId || "");
      const currentTaskId = resolveCurrentTaskId();
      if (incomingTaskId && currentTaskId && incomingTaskId !== currentTaskId) {
        return;
      }
      if (
        activeExportRequest &&
        message.requestId &&
        message.requestId !== activeExportRequest.requestId
      ) {
        return;
      }
      clearExportWatchdog();
      try {
        await handleExportDataMessage(message);
      } catch (error) {
        emitProgress({
          title: "导出任务",
          message: error instanceof Error ? error.message : String(error),
          status: STATUS.ERROR,
          taskId: incomingTaskId || currentTaskId,
          taskState: "error"
        });
      } finally {
        activeExportRequest = null;
        isExporting = false;
        clearTask();
      }
    }
  });

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message?.[CHUNKED_MESSAGE_FLAG]) {
      return false;
    }

    (async () => {
      if (isTaskCompletionMessage(message)) {
        const completedTaskId = String(message?.taskId || "");
        const currentTaskId = resolveCurrentTaskId();
        if (!currentTaskId) {
          sendResponse({ ok: true });
          return;
        }
        if (completedTaskId && completedTaskId === currentTaskId) {
          clearExportWatchdog();
          activeExportRequest = null;
          isExporting = false;
          clearTask();
        }
        sendResponse({ ok: true });
        return;
      }

      if (message?.type === MESSAGE_TYPES.REGISTER_EXPORT_TASK) {
        const incomingTaskId = String(message.taskId || message.requestId || "").trim();
        if (!incomingTaskId) {
          sendResponse({ ok: false, error: "缺少 taskId" });
          return;
        }
        if (isExporting) {
          const currentTaskId = resolveCurrentTaskId();
          if (currentTaskId && currentTaskId === incomingTaskId) {
            sendResponse({ ok: true, duplicate: true });
            return;
          }
          if (currentTaskId && currentTaskId !== incomingTaskId) {
            sendResponse({ ok: false, error: "当前已有导出任务进行中" });
            return;
          }
        }
        isExporting = true;
        beginTask({
          taskId: incomingTaskId,
          exportType: message.exportType
        });
        activeExportRequest = {
          requestId: incomingTaskId,
          taskId: incomingTaskId,
          exportType: message.exportType,
          options: message.options || {}
        };
        clearExportWatchdog();
        exportWatchdogTimer = setTimeout(() => {
          emitProgress({
            title: "导出任务",
            message: "任务超时保护触发，已自动重置当前任务状态",
            status: STATUS.WARNING,
            taskId: incomingTaskId,
            taskState: "error"
          });
          clearExportWatchdog();
          activeExportRequest = null;
          isExporting = false;
          clearTask();
        }, 2 * 60 * 1000);
        emitProgress({
          title: "导出任务",
          message: `任务已登记，等待 ${String(message.exportType || "").toUpperCase()} 导出结果`,
          status: STATUS.RUNNING,
          taskId: incomingTaskId,
          taskState: "pending"
        });
        sendResponse({ ok: true });
        return;
      }

      if (message?.type === MESSAGE_TYPES.RUN_EXPORT) {
        const incomingTaskId = String(message.taskId || message.requestId || "").trim();
        if (isExporting) {
          const currentTaskId = resolveCurrentTaskId();
          if (incomingTaskId && currentTaskId && incomingTaskId === currentTaskId) {
            sendResponse({ ok: true, duplicate: true });
            return;
          }
          sendResponse({ ok: false, error: "当前已有导出任务进行中" });
          return;
        }
        isExporting = true;
        beginTask({
          taskId: incomingTaskId,
          exportType: message.exportType
        });
        activeExportRequest = {
          requestId: message.requestId || incomingTaskId || crypto.randomUUID(),
          taskId: incomingTaskId || message.requestId || "",
          exportType: message.exportType,
          options: message.options || {}
        };
        clearExportWatchdog();
        await ensureInjectedScript();
        emitProgress({
          title: "导出任务",
          message: "已接收导出请求，正在抓取页面内容",
          status: STATUS.RUNNING,
          taskId: resolveCurrentTaskId(incomingTaskId),
          taskState: "collecting"
        });
        postToInjected({
          source: EXTENSION_BRIDGE_SOURCE,
          type: MESSAGE_TYPES.RUN_EXPORT,
          requestId: activeExportRequest.requestId,
          taskId: resolveCurrentTaskId(incomingTaskId),
          exportType: message.exportType,
          options: message.options || {}
        });
        const watchdogMs = isBitableExportType(message.exportType)
          ? 20000
          : FALLBACK_EXPORT_TIMEOUT_MS;
        exportWatchdogTimer = setTimeout(() => {
          runFallbackExport(activeExportRequest).catch(() => { });
        }, watchdogMs);
        sendResponse({ ok: true });
        return;
      }

      if (message?.type === MESSAGE_TYPES.REQUEST_PAGE_INFO) {
        await ensureInjectedScript();
        const info = await requestPageInfoFromInjected();
        sendResponse({ ok: true, info });
        return;
      }

      sendResponse({ ok: false, error: "unsupported message" });
    })().catch((error) => {
      clearExportWatchdog();
      activeExportRequest = null;
      isExporting = false;
      clearTask();
      sendResponse({
        ok: false,
        error: error instanceof Error ? error.message : String(error)
      });
    });
    return true;
  });

  addOnChunkedMessageListener(async () => undefined, {
    context: "content"
  });

  ensureInjectedScript().catch(() => { });
})();

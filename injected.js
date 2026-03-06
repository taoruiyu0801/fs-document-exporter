(function () {
  if (window.__feishuExporterInjectedLoaded__) {
    return;
  }
  window.__feishuExporterInjectedLoaded__ = true;

  const MESSAGE_TYPES = {
    GLOBAL_PROGRESS_UPDATE: "GLOBAL_PROGRESS_UPDATE",
    RUN_EXPORT: "RUN_EXPORT",
    REQUEST_PAGE_INFO: "REQUEST_PAGE_INFO"
  };
  const EXTENSION_BRIDGE_SOURCE = "feishu-exporter-extension";
  const INJECTED_BRIDGE_SOURCE = "feishu-exporter-injected";
  const LARGE_DOCUMENT_THRESHOLD = 20 * 1024 * 1024;
  const IMAGE_FETCH_TIMEOUT_MS = 12000;
  let currentTaskId = "";

  function postToExtension(message) {
    window.postMessage(
      {
        source: INJECTED_BRIDGE_SOURCE,
        ...message
      },
      "*"
    );
  }

  function emitProgress(progress) {
    const taskId = progress?.taskId || currentTaskId || "";
    postToExtension({
      type: MESSAGE_TYPES.GLOBAL_PROGRESS_UPDATE,
      data: taskId ? { ...progress, taskId } : progress,
      timestamp: Date.now()
    });
  }

  function detectPageType(url) {
    const value = String(url || "");
    if (hasBitableUrlHint(value)) {
      return "bitable";
    }
    try {
      if (hasBitableStore(window)) {
        return "bitable";
      }
      const embeddedLink = document.querySelector(
        "iframe[src*='/share/base/'],iframe[src*='/base/'],iframe[src*='/bitable/'],a[href*='/share/base/'],a[href*='/base/'],a[href*='/bitable/']"
      );
      if (embeddedLink) {
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

  function getMainContainerSelectors() {
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

  const INVISIBLE_CHARS_RE = /[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g;
  const CTRL_CHARS_RE = /[\u0000-\u001F\u007F-\u009F]/g;
  const EXPORT_TITLE_SUFFIX_RE =
    /\s*[-_–—|]\s*(飞书云文档|飞书文档|Feishu\s*Docs?|Lark\s*Docs?)\s*$/i;
  const NOISY_ATTR_RE =
    /(?:^|[\s_-])(nav|navbar|header|footer|sidebar|sider|toolbar|tool|comment|catalog|outline|search|avatar|profile|popover|modal|dialog|drawer|floating|help|feedback)(?:$|[\s_-])/i;
  const DECORATIVE_IMAGE_URL_RE =
    /(?:\/(?:static|assets?|resource|icons?|emoji|avatar)\b|(?:^|[._-])(icon|logo|avatar|emoji)(?:[._-]|$))/i;
  const DOC_IMAGE_TOKEN_RE = /\/preview\/([a-zA-Z0-9_-]{16,})/i;
  const KNOWN_FEISHU_SUFFIXES = [
    "feishu.cn",
    "feishu.net",
    "larksuite.com",
    "larkoffice.com",
    "larkenterprise.com",
    "feishu-pre.net"
  ];

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

  function buildTokenPreviewCandidates(token) {
    const safeToken = sanitizeToken(token);
    if (!isLikelyDriveToken(safeToken)) return [];
    const currentHost = String(location.hostname || "").toLowerCase();
    const currentOrigin = location.origin;
    const rootDomain = resolveFeishuRootDomain(currentHost);
    const hosts = new Set([currentHost]);
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
      if (normalized) {
        urls.push(normalized);
      }
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
    return Array.from(new Set(urls));
  }

  function normalizeExportText(value) {
    return String(value || "")
      .replace(INVISIBLE_CHARS_RE, "")
      .replace(CTRL_CHARS_RE, "")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeExportTitle(value) {
    const text = normalizeExportText(value).replace(EXPORT_TITLE_SUFFIX_RE, "").trim();
    return text || "未命名文档";
  }

  function buildPreviewUrlByToken(token) {
    return buildTokenPreviewCandidates(token)[0] || "";
  }

  function parseMimeTypeFromDataUrl(dataUrl) {
    const match = /^data:([^;,]+)[;,]/i.exec(String(dataUrl || ""));
    return match?.[1] || "application/octet-stream";
  }

  function extractTokenFromUrl(url) {
    const value = String(url || "");
    if (!value) return "";
    const fromPreview = DOC_IMAGE_TOKEN_RE.exec(value);
    if (fromPreview?.[1]) {
      const token = sanitizeToken(fromPreview[1]);
      if (isLikelyDriveToken(token)) return token;
    }
    const fromQuery = /[?&](?:token|image|img|id)=([^&#]+)/i.exec(value);
    if (fromQuery?.[1]) {
      const token = sanitizeToken(decodeURIComponent(fromQuery[1]));
      if (isLikelyDriveToken(token)) return token;
    }
    return "";
  }

  function getImageTokenFromNode(node, rawUrl = "") {
    if (!node) return extractTokenFromUrl(rawUrl);
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
      if (isLikelyDriveToken(token)) {
        return token;
      }
    }
    return extractTokenFromUrl(rawUrl);
  }

  function isLikelyDecorativeImageUrl(url) {
    const value = String(url || "");
    if (!value) return true;
    if (/internal-api-drive-stream\.feishu\.cn\/space\/api\/box\/stream\/download\/preview/i.test(value)) {
      return false;
    }
    return DECORATIVE_IMAGE_URL_RE.test(value);
  }

  function getImageRawUrlFromNode(img) {
    const candidates = getImageSourceCandidatesFromNode(img);
    return candidates.find((value) => !/^data:/i.test(String(value || ""))) || "";
  }

  function normalizeImageCandidate(value) {
    const text = String(value || "").trim();
    if (!text) return "";
    if (/^data:/i.test(text)) return text;
    const abs = toAbsoluteUrl(text);
    if (/^https?:\/\//i.test(abs) || /^blob:/i.test(abs)) {
      return abs;
    }
    if (/^https?:\/\//i.test(text) || /^blob:/i.test(text)) {
      return text;
    }
    return "";
  }

  function getImageSourceCandidatesFromNode(img) {
    const list = [];
    const push = (value) => {
      const normalized = normalizeImageCandidate(value);
      if (normalized) list.push(normalized);
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

    const tokenHint = getImageTokenFromNode(img, list[0] || "");
    for (const previewUrl of buildTokenPreviewCandidates(tokenHint)) {
      push(previewUrl);
    }

    return Array.from(new Set(list));
  }

  function findLiveImageForClone(cloneImg, liveImages = []) {
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

  function collectNodeImageCandidates(node) {
    const candidates = [];
    if (!node) return candidates;
    const push = (value) => {
      const normalized = normalizeImageCandidate(value);
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

  function extractBackgroundUrlFromComputedStyle(node) {
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

  function isPossibleImageContainerNode(node) {
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
    const textLen = normalizeExportText(node.textContent || "").length;
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

  function tryInjectBackgroundImagePlaceholders({
    cloned,
    liveRoot,
    appendImageInfo
  }) {
    if (!cloned || !liveRoot || !appendImageInfo) return 0;
    const selector =
      "[data-block-type*='image'],[data-type*='image'],figure,[class*='image'],[class*='img'],[class*='media']";
    const liveCandidates = Array.from(liveRoot.querySelectorAll(selector))
      .filter((node) => isPossibleImageContainerNode(node))
      .map((node) => {
        const sourceCandidates = [
          ...collectNodeImageCandidates(node),
          ...Array.from(node.querySelectorAll("img")).flatMap((img) =>
            collectNodeImageCandidates(img)
          )
        ];
        const computedBg = extractBackgroundUrlFromComputedStyle(node);
        if (computedBg) {
          sourceCandidates.unshift(computedBg);
        }
        const deduped = Array.from(new Set(sourceCandidates));
        const rawUrl = deduped.find((value) => !/^data:/i.test(value)) || "";
        if (!rawUrl) return null;
        return {
          rawUrl,
          tokenHint: getImageTokenFromNode(node, rawUrl),
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

  function shouldDropNodeByAttrs(node) {
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
    const textLength = normalizeExportText(node.textContent || "").length;
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
    return NOISY_ATTR_RE.test(joined);
  }

  function pruneNonContentNodes(root) {
    if (!root?.querySelectorAll) return;
    const structuralNoise = root.querySelectorAll(
      "header,footer,nav,aside,[role='navigation'],[role='banner'],[role='search'],[role='dialog'],[aria-modal='true']"
    );
    for (const node of structuralNoise) {
      node.remove();
    }
    const allNodes = root.querySelectorAll("*");
    for (const node of allNodes) {
      if (shouldDropNodeByAttrs(node)) {
        node.remove();
      }
    }
  }

  function scoreNodeForDocument(node) {
    if (!node?.querySelectorAll) return Number.NEGATIVE_INFINITY;
    const text = normalizeExportText(node.textContent || "");
    const textLength = text.length;
    if (textLength < 80) return Number.NEGATIVE_INFINITY;

    const paragraphs = node.querySelectorAll("p").length;
    const listItems = node.querySelectorAll("li").length;
    const headings = node.querySelectorAll("h1,h2,h3,h4,h5,h6").length;
    const images = node.querySelectorAll("img").length;
    const links = node.querySelectorAll("a[href]").length;
    const controls = node.querySelectorAll("button,input,select,textarea,[role='button']").length;

    if (textLength < 160 && paragraphs + listItems + headings < 3) {
      return Number.NEGATIVE_INFINITY;
    }

    const attr = `${node.getAttribute?.("class") || ""} ${node.getAttribute?.("id") || ""}`.toLowerCase();
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

  function selectBestMainContentNode() {
    const candidates = [];
    const seen = new Set();

    for (const selector of getMainContainerSelectors()) {
      const nodes = Array.from(document.querySelectorAll(selector));
      for (const node of nodes) {
        if (!node || seen.has(node)) continue;
        seen.add(node);
        const score = scoreNodeForDocument(node);
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

  function getPageInfo() {
    return {
      title: normalizeExportTitle(document.title || ""),
      url: location.href,
      pageType: detectPageType(location.href)
    };
  }

  function toAbsoluteUrl(rawUrl) {
    const value = String(rawUrl || "").trim();
    if (!value) return "";
    if (/^data:/i.test(value)) return value;
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

  function sanitizeToken(value) {
    const base = String(value || "")
      .replace(/[^a-zA-Z0-9._-]+/g, "-")
      .replace(/^-+|-+$/g, "");
    return base || "img";
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

  function cloneMainContentNode() {
    const node = selectBestMainContentNode();
    if (!node) return null;
    const cloned = node.cloneNode(true);
    pruneNonContentNodes(cloned);
    const prunedTextLength = normalizeExportText(cloned.textContent || "").length;
    if (prunedTextLength >= 80) {
      return cloned;
    }
    const fallbackClone = node.cloneNode(true);
    const fallbackTextLength = normalizeExportText(fallbackClone.textContent || "").length;
    if (fallbackTextLength >= 80) {
      return fallbackClone;
    }
    if (document.body && node !== document.body) {
      const bodyClone = document.body.cloneNode(true);
      pruneNonContentNodes(bodyClone);
      if (normalizeExportText(bodyClone.textContent || "").length >= 80) {
        return bodyClone;
      }
      return document.body.cloneNode(true);
    }
    return cloned;
  }

  function getLikelyScrollContainer() {
    const candidates = [
      document.scrollingElement,
      document.documentElement,
      document.body,
      ...Array.from(
        document.querySelectorAll(
          ".docs-editor-container,.wiki-content,[role='main'],main,[class*='scroll'],[class*='viewport']"
        )
      )
    ].filter(Boolean);
    let best = null;
    let bestScore = -1;
    for (const node of candidates) {
      const clientHeight = Number(node.clientHeight || 0);
      const scrollHeight = Number(node.scrollHeight || 0);
      const scrollable = scrollHeight - clientHeight;
      if (scrollable <= 120) continue;
      const score = scrollable + clientHeight * 0.4;
      if (score > bestScore) {
        best = node;
        bestScore = score;
      }
    }
    return best || document.scrollingElement || document.documentElement || document.body;
  }

  async function preloadDocumentByAutoScroll() {
    const scroller = getLikelyScrollContainer();
    if (!scroller || typeof scroller.scrollTop !== "number") return;
    const originalTop = Number(scroller.scrollTop || 0);
    const maxSteps = 20;
    let stableBottomHits = 0;
    let lastHeight = Number(scroller.scrollHeight || 0);
    for (let step = 0; step < maxSteps; step += 1) {
      const clientHeight = Number(scroller.clientHeight || 0);
      const scrollHeight = Number(scroller.scrollHeight || 0);
      const jump = Math.max(280, Math.floor(clientHeight * 0.9));
      const nextTop = Math.min(Number(scroller.scrollTop || 0) + jump, scrollHeight);
      scroller.scrollTop = nextTop;
      await sleep(180);
      const newHeight = Number(scroller.scrollHeight || 0);
      const nearBottom = Number(scroller.scrollTop || 0) + clientHeight >= newHeight - 6;
      if (nearBottom) {
        stableBottomHits += 1;
      } else {
        stableBottomHits = 0;
      }
      if (stableBottomHits >= 2 && Math.abs(newHeight - lastHeight) < 8) {
        break;
      }
      lastHeight = newHeight;
    }
    scroller.scrollTop = originalTop;
    await sleep(80);
  }

  function supplementImagesFromDocumentTokens({ cloned, imageInfoList, tokenSet }) {
    const discoveredItems = [];
    const discoveredKeys = new Set();
    const addImage = (rawUrl = "", rawToken = "", fallbackName = "supplement-image") => {
      let token = sanitizeToken(rawToken || "");
      let url = String(rawUrl || "").trim();

      if (!token || token === "img") {
        token = extractTokenFromUrl(url);
      }
      if ((!url || /^blob:/i.test(url)) && isLikelyDriveToken(token)) {
        url = buildPreviewUrlByToken(token);
      }

      const absUrl =
        /^https?:\/\//i.test(url) || /^blob:/i.test(url) || /^data:/i.test(url)
          ? url
          : toAbsoluteUrl(url);
      if (!absUrl) return;

      if (!token || token === "img") {
        token = tokenFromUrl(absUrl, imageInfoList.length + discoveredItems.length);
      }
      token = sanitizeToken(token);
      if (!token || token === "img") return;

      const key = `${token}|${absUrl.slice(0, 180)}`;
      if (discoveredKeys.has(key) || tokenSet.has(token)) return;
      discoveredKeys.add(key);
      tokenSet.add(token);
      discoveredItems.push({
        token,
        url: absUrl,
        name: sanitizeToken(fallbackName) || token
      });
      imageInfoList.push({
        token,
        url: absUrl,
        originalUrl: String(rawUrl || ""),
        mimeType: guessMimeType(absUrl)
      });
    };

    const addToken = (rawToken) => {
      const token = sanitizeToken(rawToken);
      if (!isLikelyDriveToken(token)) return;
      addImage("", token, token);
    };

    try {
      const tokenAttrNodes = document.querySelectorAll(
        "[data-token],[data-resource-token],[data-asset-token],[data-file-token],[data-image-token],[data-origin-token],[data-attachment-token]"
      );
      for (const node of tokenAttrNodes) {
        addToken(
          node.getAttribute("data-token") ||
            node.getAttribute("data-resource-token") ||
            node.getAttribute("data-asset-token") ||
            node.getAttribute("data-file-token") ||
            node.getAttribute("data-image-token") ||
            node.getAttribute("data-origin-token") ||
            node.getAttribute("data-attachment-token") ||
            ""
        );
        if (discoveredItems.length >= 30) break;
      }
    } catch {
      // ignore
    }

    try {
      const globalImages = document.querySelectorAll("img");
      for (const img of globalImages) {
        const width = Number(img?.naturalWidth || img?.width || 0);
        const height = Number(img?.naturalHeight || img?.height || 0);
        if (width > 0 && height > 0 && width < 48 && height < 48) {
          continue;
        }
        const rawUrl =
          img?.currentSrc ||
          img?.getAttribute?.("src") ||
          img?.getAttribute?.("data-src") ||
          img?.getAttribute?.("data-origin-src") ||
          img?.getAttribute?.("data-image-url") ||
          "";
        const token = getImageTokenFromNode(img, rawUrl);
        addImage(rawUrl, token, img?.getAttribute?.("alt") || "img");
        if (discoveredItems.length >= 40) break;
      }
    } catch {
      // ignore
    }

    if (discoveredItems.length < 40) {
      const html = String(document.documentElement?.innerHTML || "");
      const previewTokenRegex = /\/preview\/([a-zA-Z0-9_-]{16,})/g;
      let previewMatch = previewTokenRegex.exec(html);
      while (previewMatch && discoveredItems.length < 40) {
        addToken(previewMatch[1]);
        previewMatch = previewTokenRegex.exec(html);
      }
      const encodedPreviewTokenRegex = /preview%2F([a-zA-Z0-9_-]{16,})/gi;
      let encodedPreviewMatch = encodedPreviewTokenRegex.exec(html);
      while (encodedPreviewMatch && discoveredItems.length < 45) {
        addToken(encodedPreviewMatch[1]);
        encodedPreviewMatch = encodedPreviewTokenRegex.exec(html);
      }
      const tokenFieldRegex =
        /"(?:attachmentToken|imageToken|token|resource_token|resourceToken)"\s*:\s*"([a-zA-Z0-9_-]{16,})"/g;
      let tokenFieldMatch = tokenFieldRegex.exec(html);
      while (tokenFieldMatch && discoveredItems.length < 50) {
        addToken(tokenFieldMatch[1]);
        tokenFieldMatch = tokenFieldRegex.exec(html);
      }
      const imageUrlRegex =
        /https?:\/\/[^"'<>\s]+?\.(?:png|jpe?g|webp|gif)(?:\?[^"'<>\s]*)?/gi;
      let imageUrlMatch = imageUrlRegex.exec(html);
      while (imageUrlMatch && discoveredItems.length < 50) {
        addImage(imageUrlMatch[0], "", "regex-image");
        imageUrlMatch = imageUrlRegex.exec(html);
      }
    }

    if (!discoveredItems.length || !cloned?.appendChild) {
      return 0;
    }
    const container = document.createElement("section");
    container.setAttribute("data-export-fallback-images", "true");
    container.style.marginTop = "16px";
    const heading = document.createElement("h3");
    heading.textContent = "补采集图片";
    container.appendChild(heading);
    for (const item of discoveredItems) {
      const img = document.createElement("img");
      img.setAttribute("data-token", item.token);
      img.setAttribute("data-name", item.name || item.token);
      img.setAttribute("data-token-placeholder", "true");
      img.setAttribute("src", "");
      img.setAttribute("alt", item.name || item.token);
      container.appendChild(img);
    }
    cloned.appendChild(container);
    return discoveredItems.length;
  }

  function extractUrlFromAttributes(node, keys) {
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

  function buildLinkCardNode(url, titleText = "") {
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

  function normalizeEmbedLinks(root) {
    if (!root?.querySelectorAll) return;

    const iframes = Array.from(root.querySelectorAll("iframe"));
    for (const iframe of iframes) {
      const url = extractUrlFromAttributes(iframe, [
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
      iframe.replaceWith(buildLinkCardNode(url, title || "嵌入链接"));
    }

    const possibleCards = Array.from(
      root.querySelectorAll("[data-href],[data-url],[data-link],[data-link-url],[data-redirect-url]")
    );
    for (const node of possibleCards) {
      if (!node || node.closest?.("a[href]")) continue;
      if (node.querySelector?.("a[href]")) continue;
      const url = extractUrlFromAttributes(node, [
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

  function collectDocumentPayload() {
    const title = normalizeExportTitle(document.title || "未命名文档");
    const currentUrl = location.href;
    const cloned = cloneMainContentNode();
    const liveRoot = selectBestMainContentNode();
    if (!cloned) {
      throw new Error("未找到文档主体区域");
    }

    cloned.querySelectorAll("script,style,noscript,template").forEach((node) => node.remove());
    normalizeEmbedLinks(cloned);

    const imageInfoList = [];
    const tokenSet = new Set();
    let imageIndex = 0;

    const appendImageInfo = (rawUrl, altText, tokenHint = "", sourceCandidates = []) => {
      let absUrl = toAbsoluteUrl(rawUrl);
      const hintedToken = sanitizeToken(tokenHint || extractTokenFromUrl(absUrl || rawUrl));
      if (!absUrl && Array.isArray(sourceCandidates)) {
        absUrl =
          sourceCandidates
            .map((value) => normalizeImageCandidate(value))
            .find((value) => value && !/^data:/i.test(value)) || "";
      }
      if (!absUrl && hintedToken && hintedToken !== "img") {
        absUrl = buildPreviewUrlByToken(hintedToken);
      }
      if (!absUrl || /^data:/i.test(absUrl)) {
        return null;
      }
      if (isLikelyDecorativeImageUrl(absUrl)) {
        return null;
      }
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
                ? buildTokenPreviewCandidates(hintedToken)
                : [])
            ]
              .map((value) => normalizeImageCandidate(value))
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
      const liveImg = findLiveImageForClone(img, liveImages);
      const sourceImg = liveImg || img;
      const sourceCandidates = getImageSourceCandidatesFromNode(sourceImg);
      const rawUrl = sourceCandidates[0] || getImageRawUrlFromNode(sourceImg);
      const tokenHint = getImageTokenFromNode(sourceImg, rawUrl);
      const imageMeta = appendImageInfo(
        rawUrl,
        img.getAttribute("alt") || "",
        tokenHint,
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
      const tokenHint = getImageTokenFromNode(node, rawUrl);
      const sourceCandidates = collectNodeImageCandidates(node);
      if (bgUrl) {
        sourceCandidates.unshift(bgUrl);
      }
      const imageMeta = appendImageInfo(
        rawUrl,
        "embedded-image",
        tokenHint,
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

    const injectedInPlaceCount = tryInjectBackgroundImagePlaceholders({
      cloned,
      liveRoot,
      appendImageInfo
    });

    const supplementedCount =
      imageInfoList.length < 4
        ? supplementImagesFromDocumentTokens({ cloned, imageInfoList, tokenSet })
        : 0;

    const htmlWithPlaceholders = cloned.outerHTML || "";
    const htmlBlobSize = new Blob([htmlWithPlaceholders]).size;
    const imageTokens = Array.from(tokenSet);

    return {
      title,
      currentUrl,
      htmlContent: htmlWithPlaceholders,
      htmlWithPlaceholders,
      imageInfoList,
      imageTokens,
      htmlBlobSize,
      supplementedImageCount: supplementedCount + injectedInPlaceCount,
      timestamp: Date.now()
    };
  }

  function safeString(value) {
    if (typeof value !== "string") return "";
    return value.trim();
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("blob 转 dataUrl 失败"));
      reader.readAsDataURL(blob);
    });
  }

  async function fetchWithTimeout(url, options = {}, timeoutMs = IMAGE_FETCH_TIMEOUT_MS) {
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

  function normalizeImageFetchCandidates(candidates, token = "") {
    const list = [];
    const pushCandidate = (value) => {
      const item = String(value || "").trim();
      if (!item) return;
      if (!/^https?:\/\//i.test(item) && !/^blob:/i.test(item) && !/^data:/i.test(item)) return;
      list.push(item);
    };
    if (Array.isArray(candidates)) {
      for (const value of candidates) {
        pushCandidate(value);
      }
    } else {
      pushCandidate(candidates);
    }
    const safeToken = sanitizeToken(token || "");
    if (isLikelyDriveToken(safeToken)) {
      list.push(...buildTokenPreviewCandidates(safeToken));
    }
    return Array.from(new Set(list)).slice(0, 12);
  }

  async function fetchImageFromPageContext({ candidates, token }) {
    const urls = normalizeImageFetchCandidates(candidates, token);
    if (!urls.length) {
      throw new Error("无可用图片地址");
    }
    let lastError = null;
    for (const url of urls) {
      try {
        if (/^data:/i.test(url)) {
          return {
            dataUrl: url,
            mimeType: parseMimeTypeFromDataUrl(url),
            url
          };
        }
        let response;
        try {
          response = await fetchWithTimeout(
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
        const blob = await response.blob();
        if (!(blob instanceof Blob) || blob.size <= 0) {
          throw new Error("空图片数据");
        }
        const dataUrl = await blobToDataUrl(blob);
        if (!dataUrl.startsWith("data:")) {
          throw new Error("图片编码失败");
        }
        return {
          dataUrl,
          mimeType: blob.type || "application/octet-stream",
          url
        };
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("页面上下文图片拉取失败");
  }

  function hasBitableUrlHint(url) {
    return /\/(?:share\/)?(?:base|bitable)\//i.test(String(url || ""));
  }

  function getTableRowsFromDom(runtimeDocument = document) {
    if (!runtimeDocument) {
      return { headers: [], rows: [] };
    }

    const table = runtimeDocument.querySelector("table");
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

      const rowNodes = Array.from(table.querySelectorAll("tbody tr"));
      const rows = rowNodes.map((tr) =>
        Array.from(tr.children).map((cell) => (cell.textContent || "").trim())
      );
      return { headers, rows };
    }

    const grid =
      runtimeDocument.querySelector("[role='grid']") ||
      runtimeDocument.querySelector("[data-sheet-container]");
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

  function parseBitableUrlContext(url) {
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

  function hasBitableStore(runtimeWindow) {
    const store = runtimeWindow?.bitableStore;
    return Boolean(
      store &&
        typeof store.getActiveTableId === "function" &&
        store.modelOperator?.base &&
        typeof store.modelOperator.base.getTable === "function"
    );
  }

  function createRuntimeContext(runtimeWindow, origin) {
    let runtimeDocument = null;
    let runtimeUrl = "";
    try {
      runtimeDocument = runtimeWindow?.document || null;
      runtimeUrl = safeString(runtimeWindow?.location?.href || "");
    } catch {
      runtimeDocument = null;
      runtimeUrl = "";
    }
    return {
      runtimeWindow,
      runtimeDocument,
      runtimeUrl,
      origin: origin || "unknown"
    };
  }

  function appendFrameRuntimeContexts(baseContext, result, depth = 0) {
    if (!baseContext?.runtimeDocument || depth > 2) {
      return;
    }
    let frames = [];
    try {
      frames = Array.from(baseContext.runtimeDocument.querySelectorAll("iframe"));
    } catch {
      frames = [];
    }
    for (const frame of frames) {
      let frameWindow = null;
      try {
        frameWindow = frame.contentWindow;
      } catch {
        frameWindow = null;
      }
      if (!frameWindow) {
        continue;
      }
      const context = createRuntimeContext(frameWindow, `${baseContext.origin}-iframe`);
      if (!context.runtimeWindow) {
        continue;
      }
      if (!context.runtimeUrl) {
        context.runtimeUrl = safeString(frame.getAttribute("src") || "");
      }
      result.push(context);
      appendFrameRuntimeContexts(context, result, depth + 1);
    }
  }

  function scoreRuntimeContext(context) {
    let score = 0;
    if (!context) return score;
    if (hasBitableStore(context.runtimeWindow)) {
      score += 100;
    }
    if (hasBitableUrlHint(context.runtimeUrl)) {
      score += 30;
    }
    if (context.runtimeDocument) {
      const { headers, rows } = getTableRowsFromDom(context.runtimeDocument);
      if (headers.length || rows.length) {
        score += 20;
      }
    }
    if (String(context.origin || "").includes("iframe")) {
      score += 5;
    }
    return score;
  }

  function resolveBitableRuntimeContext() {
    const contexts = [];
    const rootContext = createRuntimeContext(window, "top");
    contexts.push(rootContext);
    appendFrameRuntimeContexts(rootContext, contexts);
    contexts.sort((left, right) => scoreRuntimeContext(right) - scoreRuntimeContext(left));
    return contexts[0] || rootContext;
  }

  class BitableCompressedPayloadDecoder {
    base64DecodeToUint8Array(base64Value) {
      const binary = atob(base64Value);
      const size = binary.length;
      const bytes = new Uint8Array(size);
      for (let index = 0; index < size; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
      return bytes;
    }

    async decode(base64Value) {
      if (!base64Value || typeof base64Value !== "string") {
        throw new Error("压缩数据为空");
      }
      if (typeof DecompressionStream === "undefined") {
        throw new Error("当前浏览器不支持 DecompressionStream");
      }
      const compressedBytes = this.base64DecodeToUint8Array(base64Value);
      const reader = new Blob([compressedBytes])
        .stream()
        .pipeThrough(new DecompressionStream("gzip"))
        .getReader();

      const chunks = [];
      let total = 0;
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        total += value.length;
      }

      const merged = new Uint8Array(total);
      let offset = 0;
      for (const chunk of chunks) {
        merged.set(chunk, offset);
        offset += chunk.length;
      }

      const jsonText = new TextDecoder("utf-8").decode(merged);
      return JSON.parse(jsonText);
    }
  }

  const bitableDecoder = new BitableCompressedPayloadDecoder();

  function cleanupBitableTextFieldArtifacts(tableData) {
    if (!tableData?.fieldMap || !tableData?.recordMap) return;
    for (const record of Object.values(tableData.recordMap)) {
      if (!record || typeof record !== "object") continue;
      for (const fieldId of Object.keys(record)) {
        const fieldMeta = tableData.fieldMap[fieldId];
        const value = record[fieldId];
        if (
          fieldMeta?.type === 1 &&
          value?.value &&
          Array.isArray(value.value)
        ) {
          value.value = value.value.filter((item) =>
            item && typeof item === "object" && "text" in item
              ? Boolean(String(item.text || "").trim())
              : true
          );
        }
      }
    }
  }

  function ensureFeishuSubdomain(hostValue) {
    const host = safeString(hostValue);
    const match = host.match(/^([^.]+)\.feishu\.cn$/i);
    if (match && match[1]) {
      return match[1];
    }
    return "";
  }

  function decodeParamValue(raw) {
    if (!raw) return "";
    try {
      return decodeURIComponent(raw);
    } catch {
      return String(raw);
    }
  }

  function extractBitableParamsFromResourceEntries(runtimeWindow, preferredViewId) {
    const performanceApi = runtimeWindow?.performance;
    if (!performanceApi || typeof performanceApi.getEntriesByType !== "function") {
      return null;
    }
    const entries = performanceApi.getEntriesByType("resource");
    if (!Array.isArray(entries) || entries.length === 0) {
      return null;
    }

    const candidates = [];
    for (const entry of entries) {
      const name = safeString(entry?.name || "");
      if (!name) continue;
      const matched = name.match(
        /^https:\/\/([^.]+)\.feishu\.cn\/space\/api\/v1\/bitable\/([^/]+)\/clientvars\?(.+)$/i
      );
      if (!matched) continue;

      const subdomain = safeString(matched[1]);
      const appToken = decodeParamValue(matched[2]);
      const search = new URLSearchParams(matched[3]);
      const tableId =
        search.get("tableID") || search.get("tableId") || search.get("table_id") || "";
      const viewId =
        search.get("viewID") || search.get("viewId") || search.get("view_id") || "";
      if (!subdomain || !appToken || !tableId) continue;

      candidates.push({
        subdomain,
        appToken,
        tableId,
        viewId,
        score:
          Number(entry?.responseEnd || 0) +
          (preferredViewId && preferredViewId === viewId ? 1000000 : 0)
      });
    }

    if (!candidates.length) return null;
    candidates.sort((left, right) => right.score - left.score);
    return candidates[0];
  }

  function extractBitableParamsFromLinks(runtimeDocument) {
    if (!runtimeDocument) return null;
    const anchors = Array.from(runtimeDocument.querySelectorAll("a[href]"));
    for (const anchor of anchors) {
      const href = safeString(anchor.getAttribute("href") || "");
      if (!href) continue;
      let parsed;
      try {
        parsed = new URL(href, location.href);
      } catch {
        parsed = null;
      }
      if (!parsed) continue;

      const baseMatch = parsed.pathname.match(/\/base\/([^/?#]+)/i);
      const appToken = baseMatch ? decodeParamValue(baseMatch[1]) : "";
      const tableId =
        parsed.searchParams.get("table") ||
        parsed.searchParams.get("tableId") ||
        parsed.searchParams.get("table_id") ||
        "";
      const viewId =
        parsed.searchParams.get("view") ||
        parsed.searchParams.get("viewId") ||
        parsed.searchParams.get("view_id") ||
        "";
      const subdomain = ensureFeishuSubdomain(parsed.host);
      if (appToken && tableId && subdomain) {
        return {
          subdomain,
          appToken,
          tableId,
          viewId
        };
      }
    }
    return null;
  }

  function extractBitableParamsFromScriptText(runtimeDocument, preferredViewId) {
    if (!runtimeDocument) return null;
    const scripts = Array.from(runtimeDocument.querySelectorAll("script"));
    if (!scripts.length) return null;

    let merged = "";
    for (const script of scripts) {
      const text = script.textContent || "";
      if (!text) continue;
      if (text.length > 400000) continue;
      if (
        text.includes("appToken") ||
        text.includes("tableId") ||
        text.includes("tableID") ||
        (preferredViewId && text.includes(preferredViewId))
      ) {
        merged += `${text}\n`;
      }
      if (merged.length > 1200000) {
        break;
      }
    }
    if (!merged) return null;

    const focused = (() => {
      if (!preferredViewId) return merged;
      const index = merged.indexOf(preferredViewId);
      if (index < 0) return merged;
      const start = Math.max(0, index - 10000);
      const end = Math.min(merged.length, index + 10000);
      return merged.slice(start, end);
    })();

    const appTokenMatch =
      focused.match(/"appToken"\s*:\s*"([^"]+)"/) ||
      focused.match(/"baseToken"\s*:\s*"([^"]+)"/) ||
      focused.match(/"obj_token"\s*:\s*"([^"]+)"/);
    const tableIdMatch =
      focused.match(/"tableId"\s*:\s*"([^"]+)"/) ||
      focused.match(/"tableID"\s*:\s*"([^"]+)"/) ||
      focused.match(/"table_id"\s*:\s*"([^"]+)"/);
    const viewIdMatch =
      focused.match(/"viewId"\s*:\s*"([^"]+)"/) ||
      focused.match(/"viewID"\s*:\s*"([^"]+)"/) ||
      focused.match(/"view_id"\s*:\s*"([^"]+)"/) ||
      focused.match(/"shareViewToken"\s*:\s*"([^"]+)"/);

    const appToken = decodeParamValue(appTokenMatch?.[1] || "");
    const tableId = decodeParamValue(tableIdMatch?.[1] || "");
    const viewId = decodeParamValue(viewIdMatch?.[1] || "") || preferredViewId || "";
    if (!appToken || !tableId) {
      return null;
    }
    return {
      appToken,
      tableId,
      viewId
    };
  }

  function resolveBitableApiParams(urlContext, runtimeContext) {
    const runtimeUrl = runtimeContext?.runtimeUrl || location.href;
    const runtimeParsed = parseBitableUrlContext(runtimeUrl);
    const host = safeString(runtimeParsed.host || urlContext.host);
    const fallbackSubdomain = ensureFeishuSubdomain(host);

    let appToken = safeString(urlContext.appToken || runtimeParsed.appToken);
    let tableId = safeString(urlContext.tableId || runtimeParsed.tableId);
    let viewId = safeString(urlContext.viewId || runtimeParsed.viewId);
    let subdomain = fallbackSubdomain;

    const fromResource = extractBitableParamsFromResourceEntries(
      runtimeContext?.runtimeWindow || window,
      viewId
    );
    if (fromResource) {
      appToken = appToken || safeString(fromResource.appToken);
      tableId = tableId || safeString(fromResource.tableId);
      viewId = viewId || safeString(fromResource.viewId);
      subdomain = subdomain || safeString(fromResource.subdomain);
    }

    if (!appToken || !tableId) {
      const fromLinks = extractBitableParamsFromLinks(runtimeContext?.runtimeDocument || document);
      if (fromLinks) {
        appToken = appToken || safeString(fromLinks.appToken);
        tableId = tableId || safeString(fromLinks.tableId);
        viewId = viewId || safeString(fromLinks.viewId);
        subdomain = subdomain || safeString(fromLinks.subdomain);
      }
    }

    if (!appToken || !tableId) {
      const fromScript = extractBitableParamsFromScriptText(
        runtimeContext?.runtimeDocument || document,
        viewId || safeString(urlContext.shareViewToken)
      );
      if (fromScript) {
        appToken = appToken || safeString(fromScript.appToken);
        tableId = tableId || safeString(fromScript.tableId);
        viewId = viewId || safeString(fromScript.viewId);
      }
    }

    return {
      host,
      subdomain,
      appToken,
      tableId,
      viewId
    };
  }

  async function fetchBitableClientvarsByApi({
    subdomain,
    appToken,
    tableId = "",
    viewId = "",
    needBase = true
  }) {
    const search = new URLSearchParams({
      tableID: tableId,
      viewID: viewId,
      recordLimit: "3000",
      ondemandLimit: "3000",
      needBase: String(Boolean(needBase)),
      ondemandVer: "2",
      openType: "0",
      noMissCS: "true",
      optimizationFlag: "1",
      removeFmlExtra: "true"
    });
    const url = `https://${subdomain}.feishu.cn/space/api/v1/bitable/${appToken}/clientvars?${search}`;
    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json"
      }
    });
    if (!response.ok) {
      throw new Error(`获取 clientvars 失败: ${response.status} ${response.statusText}`);
    }
    const payload = await response.json();
    if (payload.code !== 0) {
      throw new Error(`获取 clientvars 失败: ${payload.msg || payload.code}`);
    }
    const result = {
      table: await bitableDecoder.decode(payload.data.table)
    };
    cleanupBitableTextFieldArtifacts(result.table);
    if (needBase && payload.data.base) {
      result.base = await bitableDecoder.decode(payload.data.base);
    }
    return result;
  }

  async function fetchBitableRecordsByApi({
    subdomain,
    appToken,
    tableId,
    tableRev,
    offset,
    limit = 3000
  }) {
    const search = new URLSearchParams({
      tableId,
      tableRev: String(tableRev),
      depRev: "{}",
      viewLazyLoad: "true",
      offset: String(offset),
      limit: String(limit),
      tableID: tableId,
      removeFmlExtra: "true"
    });
    const url = `https://${subdomain}.feishu.cn/space/api/v1/bitable/${appToken}/records?${search}`;
    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json"
      }
    });
    if (!response.ok) {
      throw new Error(`获取 records 失败: ${response.status} ${response.statusText}`);
    }
    const payload = await response.json();
    if (payload.code !== 0) {
      throw new Error(`获取 records 失败: ${payload.msg || payload.code}`);
    }
    return bitableDecoder.decode(payload.data.records);
  }

  async function collectBitablePayloadByApi(urlContext, runtimeContext = {}) {
    const runtimeUrl = runtimeContext.runtimeUrl || location.href;
    const runtimeDocument = runtimeContext.runtimeDocument || document;
    const apiParams = resolveBitableApiParams(urlContext, runtimeContext);

    if (!apiParams.subdomain) {
      throw new Error("未解析到子域名，无法调用多维表 API");
    }
    if (!apiParams.appToken || !apiParams.tableId) {
      throw new Error("未解析到 appToken/tableId，无法调用多维表 API");
    }

    emitProgress({
      title: "Bitable 导出",
      message: "正在通过 API 拉取表格数据",
      status: "running",
      stage: "metadata",
      tableName: safeString(runtimeDocument?.title || document.title || "") || "表格",
      currentTable: 1,
      totalTables: 1,
      currentRecords: 0,
      totalRecords: 0
    });

    const clientVars = await fetchBitableClientvarsByApi({
      subdomain: apiParams.subdomain,
      appToken: apiParams.appToken,
      tableId: apiParams.tableId,
      viewId: apiParams.viewId || urlContext.viewId || "",
      needBase: false
    });

    const table = clientVars.table || {};
    const resolvedTableId =
      safeString(table?.meta?.id || table?.id || apiParams.tableId) || apiParams.tableId;
    const tableRev = Number(table.latestCSRev || table.remoteRev || 255);
    const totalRecords =
      Number(table.recordCount || 0) || Object.keys(table.recordMap || {}).length;

    if (totalRecords > 3000) {
      let offset = 2999;
      while (offset < totalRecords - 1) {
        emitProgress({
          title: "Bitable 导出",
          message: `API 分页拉取记录 ${Math.min(offset + 1, totalRecords)}/${totalRecords}`,
          status: "running",
          stage: "records",
          tableName:
            safeString(table?.meta?.name || table?.meta?.title || runtimeDocument?.title || "") ||
            "表格",
          currentTable: 1,
          totalTables: 1,
          currentRecords: Math.min(offset + 1, totalRecords),
          totalRecords
        });

        const recordsChunk = await fetchBitableRecordsByApi({
          subdomain: apiParams.subdomain,
          appToken: apiParams.appToken,
          tableId: resolvedTableId,
          tableRev,
          offset,
          limit: 3000
        });
        if (recordsChunk?.recordMap) {
          table.recordMap = {
            ...(table.recordMap || {}),
            ...recordsChunk.recordMap
          };
        }
        offset += 3000;
        await sleep(100);
      }
      cleanupBitableTextFieldArtifacts(table);
    }

    const fieldMap = table.fieldMap && typeof table.fieldMap === "object" ? table.fieldMap : {};
    const viewMap = table.viewMap && typeof table.viewMap === "object" ? table.viewMap : {};
    const currentViewId =
      safeString(
        apiParams.viewId ||
          table.currentView ||
          table.meta?.view_id ||
          urlContext.viewId
      ) || "";
    const selectedView =
      (currentViewId && viewMap[currentViewId]) || viewMap[table.currentView] || null;
    const visibleFieldIds =
      selectedView?.property?.fields && Array.isArray(selectedView.property.fields)
        ? selectedView.property.fields
        : Object.keys(fieldMap);

    const fields = visibleFieldIds
      .filter((fieldId) => fieldMap[fieldId])
      .map((fieldId) => ({
        id: String(fieldId),
        name: safeString(fieldMap[fieldId]?.name || `field_${fieldId}`),
        type: Number(fieldMap[fieldId]?.type ?? -1)
      }));

    const tableName =
      safeString(
        table?.meta?.name ||
          table?.meta?.title ||
          runtimeDocument?.title ||
          document.title ||
          "表格"
      ) || "表格";
    const recordMap = table.recordMap && typeof table.recordMap === "object" ? table.recordMap : {};
    const recordIds = Object.keys(recordMap);
    const records = [];

    for (let index = 0; index < recordIds.length; index += 1) {
      const recordId = String(recordIds[index]);
      const rawRecord = recordMap[recordId] || {};
      const fieldPayload = {};

      for (const field of fields) {
        const fieldId = field.id;
        const fieldMeta = fieldMap[fieldId] || {};
        fieldPayload[field.name] = normalizeBitableCellValue(rawRecord[fieldId], fieldMeta, {
          tableId: resolvedTableId,
          recordId,
          fieldId,
          tableRev
        });
      }

      records.push({
        rowIndex: index + 1,
        recordId,
        fields: fieldPayload
      });

      if (index === 0 || (index + 1) % 200 === 0 || index === recordIds.length - 1) {
        emitProgress({
          title: "Bitable 导出",
          message: `读取记录 ${index + 1}/${recordIds.length}`,
          status: "running",
          stage: "records",
          tableName,
          currentTable: 1,
          totalTables: 1,
          currentRecords: index + 1,
          totalRecords: recordIds.length
        });
      }
    }

    return {
      title: tableName,
      currentUrl: runtimeUrl,
      bitableData: {
        meta: {
          appToken: safeString(apiParams.appToken),
          tableId: safeString(resolvedTableId),
          viewId: safeString(currentViewId || urlContext.viewId),
          shareViewToken: safeString(urlContext.shareViewToken),
          host: safeString(apiParams.host || urlContext.host),
          source: "api-clientvars",
          exportedAt: Date.now(),
          currentUrl: runtimeUrl
        },
        tables: [
          {
            tableId: safeString(resolvedTableId),
            tableName,
            viewId: safeString(currentViewId || urlContext.viewId),
            viewName:
              safeString(selectedView?.meta?.name || selectedView?.name || "") || "当前视图",
            fields,
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

  function mapRecordIds(rawRecordIds) {
    const list = Array.isArray(rawRecordIds) ? rawRecordIds : [];
    const unique = new Set();
    for (const item of list) {
      let id = "";
      if (typeof item === "string" || typeof item === "number") {
        id = String(item);
      } else if (item && typeof item === "object") {
        id = safeString(item.recordId || item.id || item._id || "");
      }
      if (id) {
        unique.add(id);
      }
    }
    return Array.from(unique);
  }

  function buildBitableImagePreviewUrl({
    attachmentToken,
    tableId,
    recordId,
    fieldId,
    tableRev
  }) {
    const token = safeString(attachmentToken);
    if (!token) return "";
    const extra = encodeURIComponent(
      JSON.stringify({
        bitablePerm: {
          tableId,
          rev: Number.isFinite(tableRev) ? tableRev : 255,
          attachments: {
            [fieldId]: {
              [recordId]: [token]
            }
          }
        }
      })
    );
    return `https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/preview/${token}?extra=${extra}&mount_point=bitable&preview_type=16`;
  }

  function optionNameById(fieldMeta, optionId) {
    const options = fieldMeta?.property?.options;
    if (!Array.isArray(options)) {
      return optionId;
    }
    const hit = options.find((item) => String(item?.id || "") === String(optionId));
    return hit?.name || optionId;
  }

  function normalizePrimitive(value) {
    if (value == null) return "";
    if (
      typeof value === "string" ||
      typeof value === "number" ||
      typeof value === "boolean"
    ) {
      return value;
    }
    if (Array.isArray(value)) {
      return value.map((item) => normalizePrimitive(item));
    }
    if (typeof value === "object") {
      if (typeof value.text === "string") return value.text;
      if (Array.isArray(value.users)) {
        return value.users
          .map((user) => user?.name || user?.enName || user?.id || "")
          .filter(Boolean);
      }
      try {
        return JSON.parse(JSON.stringify(value));
      } catch {
        return String(value);
      }
    }
    return String(value);
  }

  function normalizeAttachmentValue(value, context) {
    if (!Array.isArray(value)) {
      return [];
    }
    return value.map((item) => {
      const attachmentToken = item?.attachmentToken || item?.id || "";
      return {
        token: attachmentToken,
        id: item?.id || attachmentToken,
        name: item?.name || attachmentToken || "attachment",
        mimeType: item?.mimeType || "",
        size: typeof item?.size === "number" ? item.size : null,
        url: buildBitableImagePreviewUrl({
          attachmentToken,
          tableId: context.tableId,
          recordId: context.recordId,
          fieldId: context.fieldId,
          tableRev: context.tableRev
        })
      };
    });
  }

  function normalizeBitableCellValue(cell, fieldMeta, context) {
    const raw =
      cell && typeof cell === "object" && Object.prototype.hasOwnProperty.call(cell, "value")
        ? cell.value
        : cell;
    const fieldType = Number(fieldMeta?.type);

    if (fieldType === 3) {
      return optionNameById(fieldMeta, raw);
    }
    if (fieldType === 4) {
      if (!Array.isArray(raw)) return [];
      return raw.map((optionId) => optionNameById(fieldMeta, optionId));
    }
    if (fieldType === 5 && typeof raw === "number" && Number.isFinite(raw)) {
      const date = new Date(raw);
      return Number.isNaN(date.getTime()) ? raw : date.toISOString();
    }
    if (fieldType === 11 && raw && Array.isArray(raw.users)) {
      return raw.users
        .map((user) => user?.name || user?.enName || user?.id || "")
        .filter(Boolean);
    }
    if (fieldType === 15 && Array.isArray(raw)) {
      return raw.map((item) => {
        if (item?.type === "url") {
          return {
            type: "url",
            text: item?.text || "",
            link: item?.link || item?.text || ""
          };
        }
        if (item?.type === "mention") {
          return {
            type: "mention",
            text: item?.text || "",
            link: item?.link || ""
          };
        }
        return item?.text || normalizePrimitive(item);
      });
    }
    if (fieldType === 17) {
      return normalizeAttachmentValue(raw, context);
    }
    return normalizePrimitive(raw);
  }

  function collectBitablePayloadFromStore(urlContext, runtimeContext = {}) {
    const runtimeWindow = runtimeContext.runtimeWindow || window;
    const runtimeDocument = runtimeContext.runtimeDocument || document;
    const runtimeUrl = runtimeContext.runtimeUrl || location.href;
    const store = runtimeWindow.bitableStore;
    if (
      !store ||
      typeof store.getActiveTableId !== "function" ||
      !store.modelOperator ||
      !store.modelOperator.base ||
      typeof store.modelOperator.base.getTable !== "function"
    ) {
      return null;
    }

    const baseOperator = store.modelOperator.base;
    const activeTableId =
      safeString(urlContext.tableId) || safeString(store.getActiveTableId() || "");
    if (!activeTableId) {
      return null;
    }

    const table = baseOperator.getTable(activeTableId);
    if (!table) {
      return null;
    }

    const allFields = table.fields && typeof table.fields === "object" ? table.fields : {};
    const views =
      table.views && typeof table.views === "object" ? Object.values(table.views) : [];
    const preferredViewId =
      safeString(urlContext.viewId) ||
      (typeof store.getActiveViewId === "function" ? safeString(store.getActiveViewId()) : "");

    let activeView = null;
    if (preferredViewId) {
      activeView =
        views.find(
          (view) =>
            safeString(view?._id || view?.id || "") === preferredViewId
        ) || null;
    }
    if (!activeView && views.length) {
      activeView = views[0];
    }

    const visibleFieldIds =
      Array.isArray(activeView?._visibleFieldIds) && activeView._visibleFieldIds.length
        ? activeView._visibleFieldIds
        : Object.keys(allFields);

    let recordIds = [];
    if (
      activeView &&
      activeView._viewSort &&
      typeof activeView._viewSort._getCurrentRecords === "function"
    ) {
      try {
        recordIds = mapRecordIds(activeView._viewSort._getCurrentRecords());
      } catch {
        recordIds = [];
      }
    }
    if (!recordIds.length) {
      recordIds = Object.keys(table.records || {});
    }

    const fields = visibleFieldIds
      .filter((fieldId) => allFields[fieldId])
      .map((fieldId) => ({
        id: String(fieldId),
        name: safeString(allFields[fieldId]?.name) || `field_${fieldId}`,
        type: Number(allFields[fieldId]?.type ?? -1)
      }));

    const tableName =
      safeString(table.name || table.title || runtimeDocument?.title || document.title || "") ||
      "表格";
    const totalRecords = recordIds.length;
    const tableRev = Number(table.remoteRev || 255);

    emitProgress({
      title: "Bitable 导出",
      message: `读取记录 0/${totalRecords}`,
      status: "running",
      stage: "records",
      tableName,
      currentTable: 1,
      totalTables: 1,
      currentRecords: 0,
      totalRecords
    });

    const records = [];
    for (let index = 0; index < recordIds.length; index += 1) {
      const recordId = String(recordIds[index]);
      const rawRecord = table.records?.[recordId] || {};
      const fieldPayload = {};

      for (const field of fields) {
        const fieldId = field.id;
        const fieldMeta = allFields[fieldId] || {};
        fieldPayload[field.name] = normalizeBitableCellValue(rawRecord[fieldId], fieldMeta, {
          tableId: activeTableId,
          recordId,
          fieldId,
          tableRev
        });
      }

      records.push({
        rowIndex: index + 1,
        recordId,
        fields: fieldPayload
      });

      if (index === 0 || (index + 1) % 200 === 0 || index === totalRecords - 1) {
        emitProgress({
          title: "Bitable 导出",
          message: `读取记录 ${index + 1}/${totalRecords}`,
          status: "running",
          stage: "records",
          tableName,
          currentTable: 1,
          totalTables: 1,
          currentRecords: index + 1,
          totalRecords
        });
      }
    }

    const viewId =
      safeString(activeView?._id || activeView?.id || "") || safeString(urlContext.viewId);
    const appToken =
      safeString(urlContext.appToken) ||
      safeString(baseOperator.appToken || "") ||
      safeString(baseOperator.baseToken || "") ||
      safeString(baseOperator.baseId || "") ||
      safeString(baseOperator.objToken || "") ||
      safeString(baseOperator.token || "") ||
      "";

    return {
      title: tableName,
      currentUrl: runtimeUrl,
      bitableData: {
        meta: {
          appToken,
          tableId: activeTableId,
          viewId,
          shareViewToken: safeString(urlContext.shareViewToken),
          host: safeString(urlContext.host) || safeString(parseBitableUrlContext(runtimeUrl).host),
          source: "bitableStore",
          exportedAt: Date.now(),
          currentUrl: runtimeUrl
        },
        tables: [
          {
            tableId: activeTableId,
            tableName,
            viewId,
            viewName:
              safeString(activeView?._name || activeView?.name || "") || "当前视图",
            fields,
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

  function collectBitablePayloadFromDom(urlContext, runtimeContext = {}, options = {}) {
    const runtimeDocument = runtimeContext.runtimeDocument || document;
    const runtimeUrl = runtimeContext.runtimeUrl || location.href;
    const allowEmpty = options.allowEmpty === true;
    const silent = options.silent === true;

    const tableName = safeString(runtimeDocument?.title || document.title || "") || "Bitable";
    const { headers, rows } = getTableRowsFromDom(runtimeDocument);

    if (!allowEmpty && headers.length === 0 && rows.length === 0) {
      throw new Error("未从页面读取到多维表数据，请确认表格已加载后重试");
    }

    const records = rows.map((row, index) => ({
      recordId: `row_${index + 1}`,
      fields: headers.reduce((acc, key, keyIndex) => {
        acc[key] = row[keyIndex] ?? "";
        return acc;
      }, {})
    }));

    if (!silent) {
      emitProgress({
        title: "Bitable 导出",
        message: `读取记录 ${records.length}/${records.length}`,
        status: "running",
        stage: "records",
        tableName,
        currentTable: 1,
        totalTables: 1,
        currentRecords: records.length,
        totalRecords: records.length
      });
    }

    return {
      title: tableName,
      currentUrl: runtimeUrl,
      bitableData: {
        meta: {
          appToken: urlContext.appToken,
          tableId: urlContext.tableId,
          viewId: urlContext.viewId,
          shareViewToken: urlContext.shareViewToken,
          host: urlContext.host,
          source: "dom-fallback",
          exportedAt: Date.now(),
          currentUrl: runtimeUrl
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

  async function collectBitablePayload() {
    const maxAttempts = 24;
    let lastContext = resolveBitableRuntimeContext();
    let tableName = safeString(lastContext?.runtimeDocument?.title || document.title || "") || "Bitable";
    let lastApiError = null;

    emitProgress({
      title: "Bitable 导出",
      message: "读取表结构信息",
      status: "running",
      stage: "metadata",
      tableName,
      currentTable: 1,
      totalTables: 1,
      currentRecords: 0,
      totalRecords: 0
    });

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      lastContext = resolveBitableRuntimeContext();
      const runtimeUrl = lastContext?.runtimeUrl || location.href;
      const urlContext = parseBitableUrlContext(runtimeUrl);
      tableName =
        safeString(lastContext?.runtimeDocument?.title || document.title || "") || "Bitable";

      const payloadByStore = collectBitablePayloadFromStore(urlContext, lastContext);
      if (payloadByStore) {
        const table = payloadByStore?.bitableData?.tables?.[0];
        const hasFieldOrRecord =
          (Array.isArray(table?.fields) && table.fields.length > 0) ||
          (Array.isArray(table?.records) && table.records.length > 0);
        if (hasFieldOrRecord || attempt >= 4) {
          return payloadByStore;
        }
      }

      const previewFromDom = collectBitablePayloadFromDom(urlContext, lastContext, {
        allowEmpty: true,
        silent: true
      });
      const domTable = previewFromDom?.bitableData?.tables?.[0];
      const domHasFieldOrRecord =
        (Array.isArray(domTable?.fields) && domTable.fields.length > 0) ||
        (Array.isArray(domTable?.records) && domTable.records.length > 0);
      if (domHasFieldOrRecord && attempt >= 2) {
        emitProgress({
          title: "Bitable 导出",
          message: "未命中 bitableStore，使用可见网格导出",
          status: "warning",
          stage: "metadata",
          tableName,
          currentTable: 1,
          totalTables: 1,
          currentRecords: 0,
          totalRecords: 0
        });
        return previewFromDom;
      }

      if (attempt === 1 || attempt % 6 === 0) {
        emitProgress({
          title: "Bitable 导出",
          message: `等待多维表数据加载（${attempt}/${maxAttempts}）`,
          status: "running",
          stage: "metadata",
          tableName,
          currentTable: 1,
          totalTables: 1,
          currentRecords: 0,
          totalRecords: 0
        });
      }
      await sleep(250);
    }

    const finalContext = lastContext || resolveBitableRuntimeContext();
    const finalRuntimeUrl = finalContext?.runtimeUrl || location.href;
    const finalUrlContext = parseBitableUrlContext(finalRuntimeUrl);
    const finalPayloadByStore = collectBitablePayloadFromStore(finalUrlContext, finalContext);
    if (finalPayloadByStore) {
      return finalPayloadByStore;
    }

    try {
      const payloadByApi = await collectBitablePayloadByApi(finalUrlContext, finalContext);
      const apiTable = payloadByApi?.bitableData?.tables?.[0];
      const apiHasData =
        (Array.isArray(apiTable?.fields) && apiTable.fields.length > 0) ||
        (Array.isArray(apiTable?.records) && apiTable.records.length > 0);
      if (apiHasData) {
        return payloadByApi;
      }
      lastApiError = new Error("API 返回空数据");
    } catch (error) {
      lastApiError = error;
      emitProgress({
        title: "Bitable 导出",
        message: `API 拉取失败，回退 DOM 抓取：${
          error instanceof Error ? error.message : String(error)
        }`,
        status: "warning",
        stage: "metadata",
        tableName,
        currentTable: 1,
        totalTables: 1,
        currentRecords: 0,
        totalRecords: 0
      });
    }

    emitProgress({
      title: "Bitable 导出",
      message: "未读取到 bitableStore，降级为 DOM 抓取",
      status: "warning",
      stage: "metadata",
      tableName,
      currentTable: 1,
      totalTables: 1,
      currentRecords: 0,
      totalRecords: 0
    });
    try {
      return collectBitablePayloadFromDom(finalUrlContext, finalContext, {
        allowEmpty: false,
        silent: false
      });
    } catch (domError) {
      if (lastApiError) {
        throw new Error(
          `未读取到可导出数据。API: ${
            lastApiError instanceof Error ? lastApiError.message : String(lastApiError)
          }; DOM: ${domError instanceof Error ? domError.message : String(domError)}`
        );
      }
      throw domError;
    }
  }

  function isBitableExportType(value) {
    const normalized = String(value || "").toLowerCase();
    return (
      normalized === "bitable-json" ||
      normalized === "bitable-csv" ||
      normalized === "bitable-xlsx"
    );
  }

  async function handleRunExport(message) {
    try {
      currentTaskId = String(message?.taskId || message?.requestId || "");
      const exportType = String(message.exportType || "").toLowerCase();
      emitProgress({
        title: "页面抓取",
        message: "开始提取页面内容",
        status: "running"
      });

      let payload;
      if (isBitableExportType(exportType)) {
        payload = await collectBitablePayload();
      } else {
        await preloadDocumentByAutoScroll().catch(() => {});
        payload = collectDocumentPayload();
      }

      if (!payload.htmlWithPlaceholders && !isBitableExportType(exportType)) {
        throw new Error("抓取结果为空，无法导出");
      }

      if (payload.htmlBlobSize > LARGE_DOCUMENT_THRESHOLD) {
        emitProgress({
          title: "页面抓取",
          message: "检测到大文档，已切换大文档策略",
          status: "warning"
        });
      }

      emitProgress({
        title: "页面抓取",
        message: `提取完成：图片 ${payload.imageInfoList.length} 张${
          payload.supplementedImageCount ? `（补采集 ${payload.supplementedImageCount}）` : ""
        }`,
        status: "success"
      });

      postToExtension({
        type: "EXPORT_DATA",
        requestId: message.requestId,
        taskId: currentTaskId,
        exportType: exportType || "markdown",
        options: message.options || {},
        payload
      });
    } catch (error) {
      postToExtension({
        type: "EXPORT_ERROR",
        requestId: message.requestId,
        taskId: currentTaskId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    const message = event.data || {};
    if (message.source !== EXTENSION_BRIDGE_SOURCE) return;

    if (message.type === MESSAGE_TYPES.REQUEST_PAGE_INFO) {
      postToExtension({
        type: "PAGE_INFO",
        requestId: message.requestId,
        pageInfo: getPageInfo()
      });
      return;
    }

    if (message.type === MESSAGE_TYPES.RUN_EXPORT) {
      handleRunExport(message);
      return;
    }

    if (message.type === "FETCH_IMAGE_FROM_PAGE") {
      const requestId = String(message.requestId || "");
      fetchImageFromPageContext({
        candidates: message.candidates || [],
        token: message.token || ""
      })
        .then((data) => {
          postToExtension({
            type: "FETCH_IMAGE_FROM_PAGE_RESULT",
            requestId,
            ok: true,
            data
          });
        })
        .catch((error) => {
          postToExtension({
            type: "FETCH_IMAGE_FROM_PAGE_RESULT",
            requestId,
            ok: false,
            error: error instanceof Error ? error.message : String(error)
          });
        });
    }
  });
})();

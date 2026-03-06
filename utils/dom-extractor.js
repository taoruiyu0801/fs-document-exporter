function isBitableLikeUrl(url) {
  const value = String(url || "");
  if (!value) return false;
  if (/\/(?:share\/)?(?:base|bitable)\//i.test(value)) return true;
  if (/[?&](?:view|viewId|table|tableId|shareViewToken)=/i.test(value) && /\/(?:share\/)?base\//i.test(value)) {
    return true;
  }
  if (/#\/?(?:share\/)?(?:base|bitable)\//i.test(value)) return true;
  return false;
}

export function detectPageType(url) {
  const value = String(url || "");
  if (isBitableLikeUrl(value)) {
    return "bitable";
  }
  if (value.includes("/wiki/") || value.includes("/docx/") || value.includes("/docs/")) {
    return "document";
  }
  return "unknown";
}

export function getMainContainerSelectors() {
  return [
    "article",
    "[role='main']",
    ".wiki-content",
    ".docs-editor-container",
    ".lark-doc-content",
    ".doc-body",
    ".text-viewer",
    "main",
    "body"
  ];
}

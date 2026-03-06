import {
  buildExportBaseName,
  ensureExtension
} from "../filename.js";

function escapeAttr(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function replaceImagePlaceholderTag(tag, srcValue) {
  const cleaned = tag
    .replace(/data-token-placeholder="true"\s*/gi, "")
    .replace(/data-token="[^"]*"\s*/gi, "")
    .replace(/data-name="[^"]*"\s*/gi, "")
    .replace(/\swidth="[^"]*"/gi, "")
    .replace(/\sheight="[^"]*"/gi, "")
    .replace(
      /\sstyle="([^"]*)"/gi,
      (_full, styleValue) => {
        const style = String(styleValue || "")
          .replace(/(?:^|;)\s*width\s*:[^;]*/gi, "")
          .replace(/(?:^|;)\s*height\s*:[^;]*/gi, "")
          .replace(/;;+/g, ";")
          .trim()
          .replace(/^;/, "")
          .replace(/;$/, "");
        return style ? ` style="${style}"` : "";
      }
    );
  if (/src\s*=\s*"[^"]*"/i.test(cleaned)) {
    return cleaned.replace(/src\s*=\s*"[^"]*"/i, `src="${escapeAttr(srcValue)}"`);
  }
  return cleaned.replace(/<img/i, `<img src="${escapeAttr(srcValue)}"`);
}

function replaceImagePlaceholdersWithDataUrl(html, imageMap) {
  return String(html || "").replace(
    /<img\b[^>]*data-token="([^"]+)"[^>]*>/gi,
    (full, token) => {
      const src = imageMap.get(token);
      if (!src) {
        return "";
      }
      return replaceImagePlaceholderTag(full, src);
    }
  );
}

function buildDocumentStyles(mode = "html") {
  const shared = `
    :root {
      --text-main: #1f2937;
      --text-sub: #475569;
      --line: #d7deea;
      --bg-soft: #f4f7fb;
      --brand: #1d4ed8;
      --brand-soft: #dbeafe;
    }
    body {
      font-family: "Noto Sans SC", "PingFang SC", "Microsoft YaHei UI", "Segoe UI", sans-serif;
      margin: 0;
      padding: 24px;
      color: var(--text-main);
      background: #fff;
      line-height: 1.72;
      font-size: 16px;
      text-rendering: optimizeLegibility;
      -webkit-font-smoothing: antialiased;
    }
    .export-meta {
      margin-bottom: 18px;
      padding: 12px 14px;
      border-left: 3px solid var(--brand);
      background: #eef4ff;
      border-radius: 8px;
      color: var(--text-sub);
      font-size: 14px;
    }
    .export-meta strong {
      color: var(--text-main);
    }
    .export-meta a {
      color: var(--brand);
      text-decoration: none;
      word-break: break-all;
    }
    h1, h2, h3, h4, h5, h6 {
      color: #0f172a;
      margin-top: 1.3em;
      margin-bottom: 0.5em;
      line-height: 1.35;
      font-weight: 700;
      letter-spacing: 0.01em;
      break-after: avoid;
      page-break-after: avoid;
    }
    h1 { font-size: 2em; border-bottom: 1px solid #e6edf8; padding-bottom: 0.24em; }
    h2 { font-size: 1.58em; border-bottom: 1px solid #edf2fb; padding-bottom: 0.18em; }
    h3 { font-size: 1.28em; }
    h4 { font-size: 1.12em; color: #1e293b; }
    h5, h6 { font-size: 1em; color: #334155; }
    p {
      margin: 0.72em 0;
      color: var(--text-main);
    }
    b, strong {
      color: #0f172a;
      font-weight: 700;
    }
    ul, ol {
      padding-left: 1.4em;
    }
    li {
      margin: 0.28em 0;
    }
    hr {
      border: 0;
      border-top: 1px solid #e5ebf5;
      margin: 1.2em 0;
    }
    img {
      max-width: 100% !important;
      width: auto !important;
      height: auto !important;
      max-height: none !important;
      object-fit: contain;
    }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; margin: 1em 0; }
    th, td {
      border: 1px solid #d1d9e6;
      padding: 7px 10px;
      vertical-align: top;
      word-break: break-word;
      overflow-wrap: anywhere;
    }
    th {
      background: #f1f5fb;
      color: #0f172a;
      font-weight: 700;
    }
    pre { background: #111827; color: #f9fafb; border-radius: 8px; padding: 12px; overflow: auto; white-space: pre-wrap; word-break: break-word; }
    blockquote { margin: 10px 0; padding: 10px 12px; border-left: 3px solid #94a3b8; background: #f8fafc; color: #334155; border-radius: 0 6px 6px 0; }
    .toc-container,
    .feishu-toc,
    [class*="-toc"] {
      border: 1px solid #cfddf7 !important;
      background: linear-gradient(180deg, #f8fbff 0%, #eff5ff 100%) !important;
      border-radius: 10px !important;
      padding: 12px 14px !important;
      margin: 12px 0 18px !important;
    }
    .toc-container h1, .toc-container h2, .toc-container h3, .toc-container h4, .toc-container h5, .toc-container h6,
    .feishu-toc-title,
    [class*="-toc-title"] {
      margin: 0 0 8px !important;
      padding: 0 !important;
      border: 0 !important;
      color: #0f3f90 !important;
      font-size: 17px !important;
      font-weight: 700 !important;
      line-height: 1.35 !important;
    }
    .toc-container a,
    .feishu-toc a,
    [class*="-toc"] a {
      color: #0f3f90 !important;
      text-decoration: none !important;
      border-radius: 6px;
      display: inline-block;
      padding: 2px 6px;
    }
    .toc-container a:hover,
    .feishu-toc a:hover,
    [class*="-toc"] a:hover {
      background: var(--brand-soft);
    }
    figure, img, table, pre, blockquote { break-inside: avoid; page-break-inside: avoid; }
  `;

  if (mode === "pdf") {
    return `
    ${shared}
    body.pdf-mode { padding: 0; }
    @page {
      size: A4;
      margin: 12mm 10mm;
    }
    @media print {
      html, body { width: auto; background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .export-meta { margin-bottom: 10mm; }
      a { color: inherit; text-decoration: none; }
      table { font-size: 12px; }
      tr, td, th { break-inside: avoid; page-break-inside: avoid; }
      img { max-width: 100% !important; max-height: none !important; width: auto !important; height: auto !important; object-fit: contain; }
      .toc-container,
      .feishu-toc,
      [class*="-toc"] {
        break-inside: avoid;
        page-break-inside: avoid;
      }
    }
  `;
  }

  return `
    ${shared}
    body.html-mode { padding: 24px; }
  `;
}

function wrapHtmlDocument({ title, originalUrl, bodyHtml, mode = "html" }) {
  const safeTitle = escapeAttr(title || "untitled");
  const safeUrl = escapeAttr(originalUrl || "");
  const safeMode = mode === "pdf" ? "pdf" : "html";
  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${safeTitle}</title>
  <style>
    ${buildDocumentStyles(safeMode)}
  </style>
</head>
<body class="${safeMode}-mode">
  <div class="export-meta">
    <div><strong>导出标题：</strong>${safeTitle}</div>
    <div><strong>原文链接：</strong><a href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeUrl}</a></div>
    <div><strong>导出时间：</strong>${new Date().toLocaleString("zh-CN")}</div>
  </div>
  ${bodyHtml}
</body>
</html>`;
}

export function composeHtmlForPreview(payload, imageMap) {
  const html = replaceImagePlaceholdersWithDataUrl(payload.htmlWithPlaceholders, imageMap);
  return wrapHtmlDocument({
    title: payload.title,
    originalUrl: payload.currentUrl,
    bodyHtml: html,
    mode: "pdf"
  });
}

export async function buildHtmlExport(payload, imageMap) {
  const baseName = buildExportBaseName(payload.title || "feishu-document");
  const singleFileHtml = wrapHtmlDocument({
    title: payload.title,
    originalUrl: payload.currentUrl,
    bodyHtml: replaceImagePlaceholdersWithDataUrl(payload.htmlWithPlaceholders, imageMap),
    mode: "html"
  });
  return {
    blob: new Blob([singleFileHtml], { type: "text/html;charset=utf-8" }),
    filename: ensureExtension(baseName, "html"),
    exportType: "HTML"
  };
}

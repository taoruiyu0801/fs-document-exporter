export const CHUNKED_MESSAGE_FLAG = "__FEISHU_EXPORTER_CHUNKED_MESSAGE__";
export const MAX_CHUNK_SIZE = 32 * 1024 * 1024;
export const LARGE_DOCUMENT_THRESHOLD = 20 * 1024 * 1024;

export const IMAGE_BATCH_SIZE = 10;
export const IMAGE_CONCURRENCY = 4;
export const IMAGE_MAX_RETRY = 3;

export const MESSAGE_TYPES = {
  GLOBAL_PROGRESS_UPDATE: "GLOBAL_PROGRESS_UPDATE",
  START_EXPORT: "START_EXPORT",
  RUN_EXPORT: "RUN_EXPORT",
  REGISTER_EXPORT_TASK: "REGISTER_EXPORT_TASK",
  REQUEST_ACTIVE_TAB_INFO: "REQUEST_ACTIVE_TAB_INFO",
  REQUEST_PAGE_INFO: "REQUEST_PAGE_INFO",
  GENERATE_PDF: "GENERATE_PDF",
  PDF_PREVIEW_READY: "PDF_PREVIEW_READY",
  EXPORT_COMPLETE: "EXPORT_COMPLETE",
  MARKDOWN_EXPORT_COMPLETE: "MARKDOWN_EXPORT_COMPLETE",
  HTML_EXPORT_COMPLETE: "HTML_EXPORT_COMPLETE",
  WORD_EXPORT_COMPLETE: "WORD_EXPORT_COMPLETE",
  PDF_EXPORT_COMPLETE: "PDF_EXPORT_COMPLETE",
  BITABLE_EXPORT_COMPLETE: "BITABLE_EXPORT_COMPLETE"
};

export const EXTENSION_BRIDGE_SOURCE = "feishu-exporter-extension";
export const INJECTED_BRIDGE_SOURCE = "feishu-exporter-injected";

export const FEISHU_MATCHES = [
  "https://*.feishu.cn/*",
  "https://*.feishu.net/*",
  "https://*.larksuite.com/*",
  "https://*.larkoffice.com/*",
  "https://*.larkenterprise.com/*",
  "https://*.feishu-pre.net/*"
];

export const IDB_META = {
  name: "feishu-doc-efficient-exporter",
  version: 1,
  stores: {
    images: "images",
    previews: "previews",
    exportTasks: "exportTasks"
  }
};

export const STATUS = {
  RUNNING: "running",
  SUCCESS: "success",
  WARNING: "warning",
  ERROR: "error"
};

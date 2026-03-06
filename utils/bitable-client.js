export function isBitableUrl(url) {
  const value = String(url || "");
  return /\/(?:share\/)?(?:base|bitable)\//i.test(value);
}

export function parseBitableUrl(url) {
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
    const reserved = new Set(["view", "table", "form", "dashboard", "share", "embed"]);
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

    const tableId = parsed.searchParams.get("table") || parsed.searchParams.get("tableId") || "";
    const queryViewId =
      parsed.searchParams.get("view") || parsed.searchParams.get("viewId") || "";

    return {
      appToken,
      tableId,
      viewId: queryViewId || shareViewToken,
      shareViewToken,
      host: parsed.host
    };
  } catch {
    return fallback;
  }
}

export function normalizeVisibleBitableRows(rows) {
  return rows.map((row, index) => ({
    rowIndex: index + 1,
    values: row
  }));
}

/** =========================================================
 * Weekly Tidbits — Server (Code.gs)
 * =========================================================
 * - Serves the web app (index.html + base.html via include)
 * - Reads Google Sheet with headers:
 *   A: Date submitted
 *   B: Weekly Tidbit Information   (title)
 *   C: Weekly Tidbit Test          (content)
 *   D: BU/OSO Sector               (categories)
 *   E: Post no later than date
 *   F: Published
 *   G: Additional Notes (ignored)
 * - Appends a computed "ContentHTML" column built from RichText
 *   so inline links (anchor text) are preserved.
 * =========================================================
 */

/** ====== CONFIG ====== */
const WEBAPP_TITLE = "Weekly Tidbits";
const SHEET_NAME = "Weekly Tidbits"; // change if your tab name differs
const HEADER_ROW = 1;
const CACHE_KEY = "tidbitsData_v6"; // bump to invalidate old cache
const CACHE_TTL_SEC = 300; // 5 minutes cache
const CACHE_BACKUP_KEY = "tidbitsData_backup_v6"; // longer backup cache
const CACHE_BACKUP_TTL_SEC = 1800; // 30 minutes backup

/** Optional: server-side allowlist for categories (not enforced here) */
const ALLOWED_CATEGORIES = [
  "Operations and Policy",
  "Trainings",
  "General Information",
  "HR",
  "Capital",
  "RTO",
];

/** =========================================================
 * Web App Entrypoint
 * ======================================================= */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile("index");
  t.allowedCategories = ALLOWED_CATEGORIES; // available to client if needed
  return t
    .evaluate()
    .setTitle(WEBAPP_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Server-side HTML include helper for <?!= include('file') ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** =========================================================
 * Sheet Menu: "Push Updates"
 * ======================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Weekly Tidbits")
    .addItem("Push Updates", "refreshWebAppData")
    .addToUi();
}

function refreshWebAppData() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY);
    cache.remove(CACHE_BACKUP_KEY);
    SpreadsheetApp.getActive().toast(
      "Weekly Tidbits: All caches cleared, updates pushed."
    );
    SpreadsheetApp.getUi().alert(
      "Updates pushed successfully. All caches cleared."
    );
  } catch (err) {
    SpreadsheetApp.getActive().toast("Weekly Tidbits: Failed to push updates.");
    SpreadsheetApp.getUi().alert("Failed to push updates: " + err);
  }
}

/** =========================================================
 * Data Endpoint for Client
 * Returns: 2D array with header row + rows, plus ContentHTML
 * ======================================================= */
function getSheetData() {
  const startTime = new Date().getTime();
  const cache = CacheService.getScriptCache();

  // Try primary cache first
  const cached = cache.get(CACHE_KEY);
  if (cached) {
    try {
      const result = JSON.parse(cached);
      console.log(
        `Cache HIT - returned ${result.length} rows in ${new Date().getTime() - startTime
        }ms`
      );
      return result;
    } catch (e) {
      console.warn("Primary cache parse error:", e);
    }
  }

  // Try backup cache if primary fails
  const backupCached = cache.get(CACHE_BACKUP_KEY);
  if (backupCached) {
    try {
      const result = JSON.parse(backupCached);
      console.log(
        `Backup cache HIT - returned ${result.length} rows in ${new Date().getTime() - startTime
        }ms`
      );
      // Restore to primary cache
      cache.put(CACHE_KEY, backupCached, CACHE_TTL_SEC);
      return result;
    } catch (e) {
      console.warn("Backup cache parse error:", e);
    }
  }

  console.log("Cache MISS - processing sheet data");
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];

  // Batch read all data at once for better performance
  const range = sh.getRange(HEADER_ROW, 1, lastRow, lastCol);
  const display = range.getDisplayValues();
  const rich = range.getRichTextValues();

  if (!display || !display.length) return [];

  const header = (display[0] || []).map((s) => (s || "").toString().trim());
  const findAnyCol = (names, fallbackIndex) => {
    const lower = header.map((h) => h.toLowerCase());
    for (const n of names) {
      const idx = lower.indexOf(n.toLowerCase());
      if (idx >= 0) return idx;
    }
    return fallbackIndex;
  };

  const idx = {
    dateSubmitted: findAnyCol(["Date submitted", "Date Submitted"], 0),
    title: findAnyCol(["Weekly Tidbit Information", "Title"], 1),
    content: findAnyCol(
      ["Weekly Tidbit Test", "Content", "Weekly Tidbit Description"],
      2
    ),
    metatags: findAnyCol(
      ["BU/OSO Sector", "Metatags", "Category", "Categories", "Tags"],
      3
    ),
    postBy: findAnyCol(
      ["Post no later than date", "Post No Later Than Date"],
      4
    ),
    published: findAnyCol(["Published"], -1),
    notes: findAnyCol(["Additional Notes"], -1),
  };

  const outHeader = header.slice();
  let htmlColIndex = outHeader.findIndex(
    (h) => h.toLowerCase() === "contenthtml"
  );
  if (htmlColIndex === -1) {
    outHeader.push("ContentHTML");
    htmlColIndex = outHeader.length - 1;
  }

  // Process rows with optimized rich text handling
  const outRows = [];
  const batchSize = 25; // Process in smaller batches for better performance

  for (
    let batchStart = 1;
    batchStart < display.length;
    batchStart += batchSize
  ) {
    const batchEnd = Math.min(batchStart + batchSize, display.length);

    for (let r = batchStart; r < batchEnd; r++) {
      const dispRow = display[r] || [];
      const richRow = rich[r] || [];
      const rt = richRow[idx.content] || null;
      const dispText = (dispRow[idx.content] || "").toString();

      // Only process rich text if there's actual content
      const html = dispText.trim() ? richTextToHtml_(rt, dispText) : "";
      const newRow = dispRow.slice();

      while (newRow.length < outHeader.length - 1) newRow.push("");
      newRow.push(html);

      outRows.push(newRow);
    }

    // Allow other processes to run between batches
    if (batchStart % 100 === 0) {
      Utilities.sleep(1);
    }
  }

  const result = [outHeader].concat(outRows);
  const resultJson = JSON.stringify(result);

  // Store in both primary and backup caches
  try {
    cache.put(CACHE_KEY, resultJson, CACHE_TTL_SEC);
    cache.put(CACHE_BACKUP_KEY, resultJson, CACHE_BACKUP_TTL_SEC);
    console.log(
      `Processed and cached ${result.length} rows in ${new Date().getTime() - startTime
      }ms`
    );
  } catch (e) {
    console.warn("Failed to cache result:", e);
  }

  return result;
}

/** =========================================================
 * RichText → HTML (preserve links + simple styles)
 * ======================================================= */

/**
 * Converts a Spreadsheet RichTextValue to safe HTML:
 * - Keeps <a href="...">anchor text</a> with target=_blank, rel=noopener
 * - Keeps bold/italic/underline
 * - Keeps foreground color via <span style="color:#RRGGBB">
 * - Preserves newlines as <br>
 * Falls back to linkifying plain text if RichText is absent.
 */
function richTextToHtml_(rtValue, fallbackText) {
  try {
    if (rtValue && typeof rtValue.getText === "function") {
      const text = rtValue.getText() || "";
      if (!text) return "";

      let html = "";
      let i = 0;
      while (i < text.length) {
        const runUrl = safeStr_(rtValue.getLinkUrl(i, i + 1));
        const runStyle = rtValue.getTextStyle(i, i + 1);
        let j = i + 1;
        while (j < text.length) {
          const nextUrl = safeStr_(rtValue.getLinkUrl(j, j + 1));
          const nextStyle = rtValue.getTextStyle(j, j + 1);
          if (!sameUrl_(runUrl, nextUrl) || !sameStyle_(runStyle, nextStyle))
            break;
          j++;
        }

        let seg = escapeHtml_(text.substring(i, j));
        // Apply inline formatting (order: link wraps inside the style wrappers or vice versa — either is fine)
        if (runUrl) {
          seg =
            '<a href="' +
            escapeHtmlAttr_(runUrl) +
            '" target="_blank" rel="noopener noreferrer">' +
            seg +
            "</a>";
        }
        if (runStyle) {
          if (runStyle.isUnderline && runStyle.isUnderline())
            seg = "<u>" + seg + "</u>";
          if (runStyle.isItalic && runStyle.isItalic())
            seg = "<em>" + seg + "</em>";
          if (runStyle.isBold && runStyle.isBold())
            seg = "<strong>" + seg + "</strong>";
          const fg =
            runStyle.getForegroundColor && runStyle.getForegroundColor();
          if (fg)
            seg =
              '<span style="color:' +
              escapeCssColor_(fg) +
              ';">' +
              seg +
              "</span>";
        }

        html += seg;
        i = j;
      }

      // Preserve line breaks
      html = html.replace(/\n/g, "<br>");
      return html;
    }
  } catch (e) {
    // Fall through to linkify fallback on any API quirk
  }
  // Fallback: plain-text linkify (http(s), www., email)
  return linkifyPlainText_(fallbackText || "");
}

/** ===== Small helpers for RichText → HTML ===== */

function sameUrl_(a, b) {
  return (a || "") === (b || "");
}

function sameStyle_(a, b) {
  const ab = a && a.isBold && a.isBold();
  const ai = a && a.isItalic && a.isItalic();
  const au = a && a.isUnderline && a.isUnderline();
  const ac = a && a.getForegroundColor && a.getForegroundColor();

  const bb = b && b.isBold && b.isBold();
  const bi = b && b.isItalic && b.isItalic();
  const bu = b && b.isUnderline && b.isUnderline();
  const bc = b && b.getForegroundColor && b.getForegroundColor();

  return ab === bb && ai === bi && au === bu && (ac || "") === (bc || "");
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeHtmlAttr_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;");
}

function escapeCssColor_(s) {
  // Very light guard; color strings from Sheets are already validated
  return String(s || "").replace(/[^#a-zA-Z0-9(),.\s]/g, "");
}

function safeStr_(s) {
  return s == null ? "" : String(s);
}

/**
 * Minimal linkifier for fallback text. (Client also linkifies, but
 * returning clickable HTML here makes server-only previews simpler.)
 */
function linkifyPlainText_(txt) {
  let s = escapeHtml_(txt);
  s = s.replace(
    /((?:https?:\/\/|ftp:\/\/)[^\s<)]+[^\s<\.,)])/gi,
    (m) =>
      '<a href="' +
      m +
      '" target="_blank" rel="noopener noreferrer">' +
      m +
      "</a>"
  );
  s = s.replace(
    /(^|[\s>])((?:www\.)[^\s<)]+[^\s<\.,)])/gi,
    (match, p1, p2) =>
      p1 +
      '<a href="https://' +
      p2 +
      '" target="_blank" rel="noopener noreferrer">' +
      p2 +
      "</a>"
  );
  s = s.replace(
    /([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi,
    (m) => '<a href="mailto:' + m + '">' + m + "</a>"
  );
  return s.replace(/\n/g, "<br>");
}

/** =========================================================
 * Performance and Cache Management Functions
 * ======================================================= */

/**
 * Preload data into cache for faster initial loads
 */
function preloadData() {
  try {
    const startTime = new Date().getTime();
    getSheetData(); // This will populate both caches
    const duration = new Date().getTime() - startTime;
    return {
      success: true,
      message: `Data preloaded successfully in ${duration}ms`,
      duration: duration,
    };
  } catch (error) {
    console.error("Preload error:", error);
    return {
      success: false,
      message: error.toString(),
      duration: 0,
    };
  }
}

/**
 * Get cache statistics for monitoring
 */
function getCacheStats() {
  try {
    const cache = CacheService.getScriptCache();
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();

    const stats = {
      primaryCache: !!cache.get(CACHE_KEY),
      backupCache: !!cache.get(CACHE_BACKUP_KEY),
      totalRows: sh.getLastRow(),
      lastColumn: sh.getLastColumn(),
      timestamp: new Date().toISOString(),
      cacheConfig: {
        primaryTTL: CACHE_TTL_SEC,
        backupTTL: CACHE_BACKUP_TTL_SEC,
      },
    };

    return stats;
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Clear all caches manually
 */
function clearAllCaches() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY);
    cache.remove(CACHE_BACKUP_KEY);

    return {
      success: true,
      message: "All caches cleared successfully",
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
    };
  }
}

/**
 * Warm up the cache by preloading data
 */
function warmUpCache() {
  try {
    const cache = CacheService.getScriptCache();

    // Clear existing caches first
    cache.remove(CACHE_KEY);
    cache.remove(CACHE_BACKUP_KEY);

    // Preload fresh data
    const result = preloadData();

    return {
      success: result.success,
      message: `Cache warmed up: ${result.message}`,
      duration: result.duration,
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
      duration: 0,
    };
  }
}

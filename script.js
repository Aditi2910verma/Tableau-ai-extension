// ---------------------------------------------
// Configuration
// ---------------------------------------------
const DEFAULT_LOADING_DELAY_MS = 800; // 0.8 seconds

// Worksheet that contains the insights data + filters
const INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend";

// Worksheets that should trigger refresh when filters change
const WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];

// Parameter that controls date logic (your Date Range Selector)
const DATE_RANGE_PARAM_NAME = "Date Range Selector";

let dashboard = null;
let loadingTimeoutId = null;
let currentDelayMs = DEFAULT_LOADING_DELAY_MS;

const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");
const insightsTableEl = () => document.getElementById("insights-table");

// ---------------------------------------------
// Utilities
// ---------------------------------------------
function log(message) {
  console.log("[AI Extension]", message);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${message}\n`;
}

// Simple typing for header text (used on init)
function typeText(element, text, speed = 35) {
  if (!element) return;
  element.textContent = "";
  let i = 0;
  const interval = setInterval(() => {
    element.textContent += text[i];
    i++;
    if (i >= text.length) clearInterval(interval);
  }, speed);
}

// Escape text for HTML attribute
function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const loadingScreen = document.getElementById("loading-screen");
  const insightsScreen = document.getElementById("insights-screen");

  // Splash screen for 0.8s
  setTimeout(() => {
    if (loadingScreen) loadingScreen.style.display = "none";
    if (insightsScreen) {
      insightsScreen.style.display = "block";
      insightsScreen.classList.add("show");
      const status = statusEl();
      if (status) status.textContent = "Initializing AI Insights…";
    }
  }, 800);

  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log(`Dashboard name: ${dashboard.name}`);
        dashboard.worksheets.forEach(ws => log(`Worksheet available: ${ws.name}`));

        subscribeToFilterChanges();
        subscribeToDateRangeParameter();

        // First-time nice message
        typeText(statusEl(), "AI Generated Insights");

        // Initial data load
        return refreshInsights();
      })
      .catch(err => {
        console.error("Failed to initialize Tableau Extension:", err);
        const status = statusEl();
        if (status) status.textContent = "Failed to initialize extension (see log).";
        log(`Tableau init failed: ${err.message || err}`);
      });
  } catch (e) {
    console.warn("Not running inside Tableau, demo mode:", e);
    log("Tableau Extensions API not found. UI shown, but logic disabled.");
    const status = statusEl();
    if (status) status.textContent = "Running outside Tableau (demo mode).";
  }
});

// ---------------------------------------------
// Subscriptions
// ---------------------------------------------
function subscribeToFilterChanges() {
  dashboard.worksheets.forEach((ws) => {
    if (WORKSHEETS_TO_SUBSCRIBE.includes(ws.name)) {
      ws.addEventListener(
        tableau.TableauEventType.FilterChanged,
        () => onSomethingChanged("filter", ws.name)
      );
      log(`Subscribed to FilterChanged on ${ws.name}`);
    } else {
      log(`Skipping worksheet: ${ws.name}`);
    }
  });
}

function subscribeToDateRangeParameter() {
  dashboard.getParametersAsync().then(params => {
    const dateParam = params.find(p => p.name === DATE_RANGE_PARAM_NAME);
    if (!dateParam) {
      log(`Date Range parameter "${DATE_RANGE_PARAM_NAME}" not found.`);
      return;
    }

    log(`Parameter available: ${dateParam.name}`);
    dateParam.addEventListener(
      tableau.TableauEventType.ParameterChanged,
      () => onSomethingChanged("parameter", dateParam.name)
    );
    log(`Subscribed to ParameterChanged on "${dateParam.name}"`);
  }).catch(err => {
    log(`Error subscribing to date parameter: ${err.message || err}`);
  });
}

// ---------------------------------------------
// Change handler
// ---------------------------------------------
function onSomethingChanged(type, name) {
  log(`${type} changed: ${name}`);

  // Debounce rapid changes
  if (loadingTimeoutId) {
    clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
    log("Existing timer cleared (debounce).");
  }

  const status = statusEl();
  const tableContainer = insightsTableEl();

  // Immediately show loading & hide cards
  if (status) status.textContent = "Loading your Insight...";
  if (tableContainer) tableContainer.style.visibility = "hidden";

  // After 0.8s, refresh and show cards again
  loadingTimeoutId = setTimeout(() => {
    refreshInsights()
      .then(() => {
        if (tableContainer) tableContainer.style.visibility = "visible";
        if (status) status.textContent = "AI Generated Insights";
      })
      .catch(err => {
        log(`Error during refreshInsights: ${err.message || err}`);
        if (status) status.textContent = "Error updating AI Insights (see log).";
        if (tableContainer) tableContainer.style.visibility = "visible";
      })
      .finally(() => {
        loadingTimeoutId = null;
      });
  }, currentDelayMs);
}

// ---------------------------------------------
// Data & rendering
// ---------------------------------------------
async function refreshInsights() {
  if (!dashboard) return;

  const container = insightsTableEl();
  if (!container) return;

  const sheet = dashboard.worksheets.find(ws => ws.name === INSIGHTS_WORKSHEET_NAME);
  if (!sheet) {
    log(`Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.`);
    container.innerHTML =
      `<em>Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.</em>`;
    return;
  }

  log(`Fetching summary data from "${INSIGHTS_WORKSHEET_NAME}"`);
  const dataTable = await sheet.getSummaryDataAsync();

  const cols = dataTable.columns;
  const rows = dataTable.data;

  // EXTRA DEBUG INFO: column names and a couple of sample rows
  log("Columns found: " + cols.map(c => c.fieldName).join(" | "));
  if (rows.length > 0) {
    const previewCount = Math.min(3, rows.length);
    for (let i = 0; i < previewCount; i++) {
      const row = rows[i];
      log(`Row ${i} preview (raw formatted values): ` +
        cols.map((c, idx) => `${c.fieldName}=${row[idx].formattedValue}`).join(" | ")
      );
    }
  }

  renderInsightsAsCards(cols, rows);
}

// Typewriter effect for all insight bodies
function animateInsightBodies(speedPerChar = 35) {
  const bodies = document.querySelectorAll(".insight-body[data-fulltext]");
  bodies.forEach((el, index) => {
    const fullText = el.getAttribute("data-fulltext") || "";
    el.textContent = "";

    let i = 0;
    const startDelay = index * 100; // stagger between cards

    setTimeout(() => {
      const interval = setInterval(() => {
        el.textContent += fullText[i];
        i++;
        if (i >= fullText.length) {
          clearInterval(interval);
        }
      }, speedPerChar);
    }, startDelay);
  });
}

// Renders cards with typing effect
function renderInsightsAsCards(columns, rows) {
  const container = insightsTableEl();
  if (!container) return;

  if (!rows || rows.length === 0) {
    container.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  // Map column names -> index
  const colIndex = {};
  columns.forEach((col, idx) => {
    colIndex[col.fieldName] = idx;
  });

  // Column indices (adjust names if your headers differ)
  const brandIdx  = colIndex["Brand"];
  const hcpIdx    = colIndex["Hcp Dtc Identifier"];
  const sourceIdx = colIndex["Source"];
  const dateIdx   = colIndex["Current Period Date Range"];
  const insightIdx =
    colIndex["Estimated Spend - Insight1"] ??
    colIndex["Estimated Spend - Insight"] ??
    colIndex["Estimated Spend - Insight 1"];

  let html = "<div class='insights-grid'>";

  rows.forEach(row => {
    const brandVal   = brandIdx   != null ? row[brandIdx].formattedValue   : "";
    const hcpVal     = hcpIdx     != null ? row[hcpIdx].formattedValue     : "";
    const sourceVal  = sourceIdx  != null ? row[sourceIdx].formattedValue  : "";
    const dateVal    = dateIdx    != null ? row[dateIdx].formattedValue    : "";
    const insightVal = insightIdx != null ? row[insightIdx].formattedValue : "";

    // Treat "Null" / "(Null)" / empty as "no AI insight yet"
    const isNullish =
      !insightVal ||
      insightVal === "Null" ||
      insightVal === "(Null)";

    const safeInsight = isNullish
      ? "No AI insight is available yet for this combination."
      : insightVal;

    html += `
<div class="insight-card">
  <div class="insight-card-header">
    <div class="insight-brand">${escapeHtml(brandVal || "—")}</div>
    ${hcpVal    ? `<span class="insight-badge">${escapeHtml(hcpVal)}</span>` : ""}
    ${sourceVal ? `<span class="insight-badge">${escapeHtml(sourceVal)}</span>` : ""}
  </div>
  ${dateVal ? `<div class="insight-date">${escapeHtml(dateVal)}</div>` : ""}
  <div class="insight-body" data-fulltext="${escapeHtml(safeInsight)}">
    <!-- text will be filled by typewriter -->
  </div>
</div>
`;
  });

  html += "</div>";

  container.innerHTML = html;

  // Typewriter effect for card bodies
  animateInsightBodies();
}

// ---------------------------------------------
// Configuration
// ---------------------------------------------
const INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend";
const WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];
const DATE_RANGE_PARAM_NAME = "Date Range Selector";

let dashboard = null;
let isRefreshing = false;

// DOM helpers
const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");
const insightsTableEl = () => document.getElementById("insights-table");

// Simple logger (kept silent in UI)
function log(msg) {
  console.log("[AI Insights]", msg);
  const el = logEl();
  if (!el) return;
  el.textContent += msg + "\n";
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const loadingScreen = document.getElementById("loading-screen");
  const insightsScreen = document.getElementById("insights-screen");

  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log(`Dashboard: ${dashboard.name}`);

        // Final title text – we force it once here
        if (statusEl()) statusEl().textContent = "AI Generated Insights";

        subscribeToFilterChanges();
        subscribeToDateRangeParameter();

        // Initial load of insights
        return refreshInsights();
      })
      .then(() => {
        // Hide loading screen, show insights UI
        if (loadingScreen) loadingScreen.style.display = "none";
        if (insightsScreen) {
          insightsScreen.style.display = "block";
          insightsScreen.classList.add("show");
        }
      })
      .catch(err => {
        console.error("Failed to initialize Tableau Extension:", err);
        if (loadingScreen) loadingScreen.textContent =
          "AI Insights extension failed to initialize.";
      });
  } catch (e) {
    // Running outside Tableau (for local preview)
    console.warn("Not running inside Tableau (demo mode):", e);
    if (loadingScreen) loadingScreen.textContent =
      "Running outside Tableau (demo mode).";
  }
});

// ---------------------------------------------
// Subscriptions
// ---------------------------------------------
function subscribeToFilterChanges() {
  if (!dashboard) return;
  dashboard.worksheets.forEach(ws => {
    if (WORKSHEETS_TO_SUBSCRIBE.includes(ws.name)) {
      ws.addEventListener(
        tableau.TableauEventType.FilterChanged,
        () => onSomethingChanged("filter", ws.name)
      );
      log(`Subscribed to FilterChanged on ${ws.name}`);
    }
  });
}

function subscribeToDateRangeParameter() {
  if (!dashboard) return;
  dashboard.getParametersAsync()
    .then(params => {
      const dateParam = params.find(p => p.name === DATE_RANGE_PARAM_NAME);
      if (!dateParam) {
        log(`Date range parameter "${DATE_RANGE_PARAM_NAME}" not found.`);
        return;
      }
      dateParam.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        () => onSomethingChanged("parameter", dateParam.name)
      );
      log(`Subscribed to ParameterChanged on "${dateParam.name}"`);
    })
    .catch(err => log(`Error subscribing to date param: ${err.message || err}`));
}

// ---------------------------------------------
// Change handler
// ---------------------------------------------
async function onSomethingChanged(type, name) {
  log(`${type} changed: ${name}`);
  if (isRefreshing) {
    log("Refresh already in progress; skipping new event.");
    return;
  }
  isRefreshing = true;

  const status = statusEl();
  if (status) status.textContent = "AI Generated Insights";

  try {
    await refreshInsights();
  } catch (err) {
    log(`Error during refresh: ${err.message || err}`);
  } finally {
    if (status) status.textContent = "AI Generated Insights";
    isRefreshing = false;
  }
}

// ---------------------------------------------
// Data + Rendering
// ---------------------------------------------
async function refreshInsights() {
  if (!dashboard) return;
  const tableContainer = insightsTableEl();
  if (!tableContainer) return;

  const insightsSheet = dashboard.worksheets.find(
    ws => ws.name === INSIGHTS_WORKSHEET_NAME
  );
  if (!insightsSheet) {
    log(`Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.`);
    tableContainer.innerHTML =
      `<em>Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.</em>`;
    return;
  }

  log(`Fetching summary data from "${INSIGHTS_WORKSHEET_NAME}"`);
  const dataTable = await insightsSheet.getSummaryDataAsync();
  const cols = dataTable.columns;
  const rows = dataTable.data;

  renderInsightsTable(cols, rows);
}

function renderInsightsTable(columns, rows) {
  const tableContainer = insightsTableEl();
  if (!tableContainer) return;

  if (!rows || rows.length === 0) {
    tableContainer.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  // Map column indexes by name (adjust names if your sheet uses slightly different labels)
  const colIndex = {};
  columns.forEach((c, idx) => {
    colIndex[c.fieldName] = idx;
  });

  const brandIdx   = colIndex["Brand"] ?? colIndex["Brand "];
  const hcpIdx     = colIndex["Hcp Dtc Identifier"] ?? colIndex["HCP Dtc Identifier"];
  const sourceIdx  = colIndex["Source"];
  const dateIdx    = colIndex["Current Period Date Range"];
  const textIdx    = colIndex["Estimated Spend - Insight1"] ?? colIndex["Insight1"];

  let html = '<div class="insights-grid">';

  rows.forEach(row => {
    const brand = brandIdx != null ? row[brandIdx].formattedValue : "";
    const hcp   = hcpIdx   != null ? row[hcpIdx].formattedValue   : "";
    const src   = sourceIdx!= null ? row[sourceIdx].formattedValue: "";
    const date  = dateIdx  != null ? row[dateIdx].formattedValue  : "";
    const raw   = textIdx  != null ? row[textIdx].formattedValue  : "";

    const bodyHtml = highlightNumbers(raw || "");

    html += `
      <div class="insight-card">
        <div class="insight-card-header">
          <span class="insight-brand">${escapeHtml(brand || "")}</span>
          ${hcp ? `<span class="insight-badge">${escapeHtml(hcp)}</span>` : ""}
          ${src ? `<span class="insight-badge">${escapeHtml(src)}</span>` : ""}
        </div>
        <div class="insight-date">${escapeHtml(date || "")}</div>
        <div class="insight-body">${bodyHtml}</div>
      </div>
    `;
  });

  html += "</div>";
  tableContainer.innerHTML = html;
}

// ---------------------------------------------
// Helpers
// ---------------------------------------------

// Bold numeric values: 19.8%, $15.6K, 27.4%, 90K, etc.
function highlightNumbers(text) {
  if (!text) return "";
  const safe = escapeHtml(text);
  const numberRegex = /\b(\$?\d[\d,]*(?:\.\d+)?%?K?)\b/g;
  return safe.replace(numberRegex, "<strong>$1</strong>");
}

// Basic HTML escape to avoid issues when inserting into innerHTML
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

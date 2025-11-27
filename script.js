// ---------------------------------------------
// Configuration
// ---------------------------------------------
const INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend"; // must match sheet name
const WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];
const DATE_RANGE_PARAM_NAME = "Date Range Selector"; // if present, we subscribe

let dashboard = null;
let isRefreshing = false;

// DOM helpers
const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");
const insightsTableEl = () => document.getElementById("insights-table");
const loadingScreenEl = () => document.getElementById("loading-screen");
const insightsScreenEl = () => document.getElementById("insights-screen");

// Simple logger (hidden in UI unless you turn #log on)
function log(msg) {
  console.log("[AI Extension]", msg);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${msg}\n`;
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log(`Dashboard: ${dashboard.name}`);

        // Show main UI, hide loading
        const loading = loadingScreenEl();
        const main = insightsScreenEl();
        if (loading) loading.style.display = "none";
        if (main) {
          main.style.display = "block";
          main.classList.add("show");
        }

        // Title stays "AI Generated Insights" from HTML
        subscribeToFilterChanges();
        subscribeToDateRangeParameter();
        refreshInsights();
      })
      .catch(err => {
        console.error("Failed to initialize Tableau Extension:", err);
        log(`Tableau init failed: ${err.message || err}`);
        const loading = loadingScreenEl();
        const main = insightsScreenEl();
        if (loading) loading.textContent = "AI Insights extension failed to initialize.";
        if (main) main.style.display = "none";
      });
  } catch (e) {
    // Running outside Tableau (browser preview)
    console.warn("Not running inside Tableau, demo mode:", e);
    log("Tableau Extensions API not found. Demo mode only.");
    const loading = loadingScreenEl();
    const main = insightsScreenEl();
    if (loading) loading.style.display = "none";
    if (main) {
      main.style.display = "block";
      main.classList.add("show");
    }
    demoRender();
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
        () => handleChange("filter", ws.name)
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
        log(`Date parameter "${DATE_RANGE_PARAM_NAME}" not found (ok if you don't use it).`);
        return;
      }

      dateParam.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        () => handleChange("parameter", dateParam.name)
      );
      log(`Subscribed to ParameterChanged on "${DATE_RANGE_PARAM_NAME}"`);
    })
    .catch(err => log(`Error subscribing to parameters: ${err.message || err}`));
}

// ---------------------------------------------
// Change handler
// ---------------------------------------------
function handleChange(type, name) {
  log(`${type} changed: ${name}`);

  if (isRefreshing) {
    log("Refresh already in progress; skipping new event.");
    return;
  }

  isRefreshing = true;
  const status = statusEl();
  if (status) status.textContent = "AI Generated Insights";

  refreshInsights()
    .catch(err => {
      log(`Error during refresh: ${err.message || err}`);
    })
    .finally(() => {
      isRefreshing = false;
    });
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
    container.innerHTML = `<em>Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.</em>`;
    return;
  }

  try {
    log(`Fetching summary data from "${INSIGHTS_WORKSHEET_NAME}"`);
    const dataTable = await sheet.getSummaryDataAsync();
    const cols = dataTable.columns;
    const rows = dataTable.data;

    renderInsightsAsCards(cols, rows);
  } catch (err) {
    log(`Error fetching insights data: ${err.message || err}`);
    insightsTableEl().innerHTML = "<em>Error loading insights data.</em>";
  }
}

// Map a column name to its index (case-insensitive)
function buildColumnIndex(columns) {
  const map = {};
  columns.forEach((c, i) => {
    if (!c || !c.fieldName) return;
    map[c.fieldName.trim().toLowerCase()] = i;
  });
  return map;
}

// Simple helper for HTML escaping
function escapeHtml(str) {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// Bold numbers in the text
function highlightNumbers(text) {
  if (!text) return "";
  return text.replace(/(\$?\d[\d,]*(\.\d+)?%?)/g, "<strong>$1</strong>");
}

function renderInsightsAsCards(columns, rows) {
  const container = insightsTableEl();
  if (!container) return;

  if (!rows || rows.length === 0) {
    container.innerHTML = `<div class="no-insights">No insights for the current selection.</div>`;
    return;
  }

  const colIndex = buildColumnIndex(columns);

  const brandIdx  = colIndex["brand"] ?? colIndex["brandname"];
  const hcpIdx    = colIndex["hcp dtc identifier"] ?? colIndex["hcp dtc"];
  const srcIdx    = colIndex["source"];
  const dateIdx   = colIndex["current period date range"] ?? colIndex["date range"];
  const textIdx   = colIndex["estimated spend - insight1"] ?? colIndex["insight"] ?? colIndex["ai insight"];

  let html = `<div class="insights-grid">`;

  rows.forEach(row => {
    const getVal = idx =>
      (idx == null || idx < 0 || idx >= row.length)
        ? ""
        : row[idx].formattedValue;

    const brand = getVal(brandIdx);
    const hcp   = getVal(hcpIdx);
    const src   = getVal(srcIdx);
    const date  = getVal(dateIdx);
    const text  = getVal(textIdx);

    const bodyHtml = text ? highlightNumbers(escapeHtml(text)) : "<em>Null</em>";

    html += `
      <div class="insight-card">
        <div class="insight-card-header">
          <span class="insight-brand">${escapeHtml(brand || "")}</span>
          ${hcp ? `<span class="insight-badge">${escapeHtml(hcp)}</span>` : ""}
          ${src ? `<span class="insight-badge">${escapeHtml(src)}</span>` : ""}
        </div>
        ${date ? `<div class="insight-date">${escapeHtml(date)}</div>` : ""}
        <div class="insight-body">${bodyHtml}</div>
      </div>
    `;
  });

  html += `</div>`;
  container.innerHTML = html;
}

// ---------------------------------------------
// Demo mode (when not in Tableau)
// ---------------------------------------------
function demoRender() {
  const container = insightsTableEl();
  if (!container) return;

  container.innerHTML = `
    <div class="insights-grid">
      <div class="insight-card">
        <div class="insight-card-header">
          <span class="insight-brand">DEMO BRAND</span>
          <span class="insight-badge">HCP</span>
          <span class="insight-badge">Google</span>
        </div>
        <div class="insight-date">Demo Date Range</div>
        <div class="insight-body">
          Spend increased to <strong>$25K</strong>, up <strong>12.5%</strong> vs prior period.
        </div>
      </div>
    </div>
  `;
}

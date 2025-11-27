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

// Logging
function log(message) {
  console.log("[AI Extension]", message);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${message}\n`;
}

// Typewriter effect for insight text
function typeText(element, text, speed = 25, onComplete = null) {
  if (!element) return;
  element.textContent = "";
  let i = 0;
  const interval = setInterval(() => {
    element.textContent += text[i];
    i++;
    if (i >= text.length) {
      clearInterval(interval);
      if (typeof onComplete === "function") {
        onComplete();
      }
    }
  }, speed);
}

// Wrap numeric tokens in <strong> tags
function boldNumbersInElement(el) {
  if (!el) return;
  const raw = el.textContent || "";
  if (!raw) return;
  const html = raw.replace(/(\d[\d.,%]*)/g, "<strong>$1</strong>");
  el.innerHTML = html;
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const status = statusEl();
  if (status) status.textContent = "";
  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log(`Dashboard name: ${dashboard.name}`);

        subscribeToFilterChanges();
        subscribeToParameterChanges();

        refreshInsights();
      })
      .catch((err) => {
        const msg = (err && (err.message || err.toString())) || "Unknown error";
        console.error("Failed to initialize Tableau Extension:", err);
        if (status) status.textContent = "AI Insights extension failed to initialize.";
        const l = logEl();
        if (l) {
          l.style.display = "block";
          l.textContent += `\n[INIT ERROR] ${msg}\n`;
        }
      });
  } catch (e) {
    console.warn("Not running inside Tableau, demo mode:", e);
    if (status) status.textContent = "Running outside Tableau (demo mode).";
    log("Tableau Extensions API not found. UI shown, but logic disabled.");
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
    }
  });
}

function subscribeToParameterChanges() {
  dashboard.getParametersAsync()
    .then(params => {
      const dateParam = params.find(p => p.name === DATE_RANGE_PARAM_NAME);
      if (!dateParam) {
        log(`Parameter "${DATE_RANGE_PARAM_NAME}" not found.`);
        return;
      }
      dateParam.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        () => onSomethingChanged("parameter", dateParam.name)
      );
      log(`Subscribed to ParameterChanged on "${dateParam.name}"`);
    })
    .catch(err => log(`Error subscribing to parameters: ${err.message || err}`));
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
  if (status) status.textContent = "Updating insights…";

  try {
    await refreshInsights();
    if (status) status.textContent = "";
  } catch (err) {
    log(`Error during refresh: ${err.message || err}`);
    if (status) status.textContent = "Error updating AI insights.";
  } finally {
    isRefreshing = false;
  }
}

// ---------------------------------------------
// Data & rendering
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

  renderInsightsCards(cols, rows);
}

function renderInsightsCards(columns, rows) {
  const tableContainer = insightsTableEl();
  if (!tableContainer) return;

  if (!rows || rows.length === 0) {
    tableContainer.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  const getIndex = (name) =>
    columns.findIndex(c => c.fieldName === name);

  const brandIdx = getIndex("Brand");
  const hcpDtcIdx = getIndex("Hcp Dtc Identifier");
  const sourceIdx = getIndex("Source");
  const dateIdx = getIndex("Current Period Date Range");
  const insightIdx = getIndex("Estimated Spend - Insight1");

  let html = "";
  rows.forEach((row, i) => {
    const brand = brandIdx >= 0 ? row[brandIdx].formattedValue : "";
    const hcpDtc = hcpDtcIdx >= 0 ? row[hcpDtcIdx].formattedValue : "";
    const source = sourceIdx >= 0 ? row[sourceIdx].formattedValue : "";
    const date = dateIdx >= 0 ? row[dateIdx].formattedValue : "";
    const insightText = insightIdx >= 0 ? row[insightIdx].formattedValue : "";

    const cardId = `insight-card-${i}`;
    const textId = `insight-text-${i}`;

    html += `
      <div class="insight-card" id="${cardId}">
        <div class="insight-header-row">
          <div class="insight-brand">${brand || ""}</div>
          <div class="pill-row">
            ${hcpDtc ? `<span class="pill">${hcpDtc}</span>` : ""}
            ${source ? `<span class="pill">${source}</span>` : ""}
          </div>
        </div>
        <div class="insight-date">${date || ""}</div>
        <div class="insight-text" id="${textId}"></div>
      </div>
    `;
  });

  tableContainer.innerHTML = html;

  // Apply typing effect + bold numbers for each card text
  rows.forEach((row, i) => {
    const textEl = document.getElementById(`insight-text-${i}`);
    if (!textEl) return;
    const insightIdx = columns.findIndex(c => c.fieldName === "Estimated Spend - Insight1");
    const raw = insightIdx >= 0 ? row[insightIdx].formattedValue || "" : "";
    if (!raw) return;

    typeText(textEl, raw, 25, () => boldNumbersInElement(textEl));
  });
}

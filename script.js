// ---------------------------------------------
// Configuration
// ---------------------------------------------
const AI_STATE_PARAM_NAME = "p_AI_Insight_State";
const AI_DELAY_PARAM_NAME = null; // optional parameter name if you want delay from Tableau
const DEFAULT_LOADING_DELAY_MS = 800; // 0.8 seconds

// Worksheet that contains the insights data + filters
const INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend";

// Worksheets that should trigger AI state + refresh when filters change
const WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];

// Parameter that should also trigger refresh (your date range selector)
const DATE_RANGE_PARAM_NAME = "Date Range Selector";

let dashboard = null;
let aiStateParam = null;
let loadingTimeoutId = null;
let currentDelayMs = DEFAULT_LOADING_DELAY_MS;

const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");
const insightsTableEl = () => document.getElementById("insights-table");

// Regex to bold numbers after typing
const NUM_REGEX = /(\$?\d[\d,\.]*\s?[MK%]?)/g;

// Utility: append a log line
function log(message) {
  console.log("[AI Loading Extension]", message);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${message}\n`;
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const loadingScreen = document.getElementById("loading-screen");
  const insightsScreen = document.getElementById("insights-screen");

  // Show loading splash for 0.8s, then fade in insights UI
  setTimeout(() => {
    if (loadingScreen) loadingScreen.style.display = "none";
    if (insightsScreen) {
      insightsScreen.style.display = "block";
      insightsScreen.classList.add("show");
      const s = statusEl();
      if (s) s.textContent = "AI Generated Insights | Estimated Spend";
    }
  }, 800);

  // Initialize Tableau extension
  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log(`Dashboard name: ${dashboard.name}`);

        return findAiStateParameter();
      })
      .then(() => maybeReadDelayParameter())
      .then(() => {
        subscribeToFilterChanges();
        subscribeToParameterChanges();   // listen to Date Range Selector
        // Initial load of insights table
        return refreshInsights();
      })
      .catch((err) => {
        console.error("Failed to initialize Tableau Extension:", err);
        const s = statusEl();
        if (s) s.textContent = "AI Insights extension failed to initialize.";
        log(`Tableau init failed: ${err.message || err}`);
      });
  } catch (e) {
    // This happens when running in a normal browser, not inside Tableau
    console.warn("Not running inside Tableau, demo mode:", e);
    log("Tableau Extensions API not found. UI shown, but parameter logic disabled.");
    const s = statusEl();
    if (s) s.textContent = "AI Generated Insights | Estimated Spend (demo)";
  }
});

// ---------------------------------------------
// Parameter helpers
// ---------------------------------------------
async function findAiStateParameter() {
  log(`Looking for parameter: ${AI_STATE_PARAM_NAME}`);
  const param = await dashboard.findParameterAsync(AI_STATE_PARAM_NAME);
  if (!param) throw new Error(`Parameter "${AI_STATE_PARAM_NAME}" not found.`);
  aiStateParam = param;
  log(`Found parameter "${param.name}". Current value: ${param.currentValue?.value}`);
}

async function maybeReadDelayParameter() {
  if (!AI_DELAY_PARAM_NAME) return;
  const delayParam = await dashboard.findParameterAsync(AI_DELAY_PARAM_NAME);
  if (delayParam && delayParam.currentValue) {
    const value = Number(delayParam.currentValue.value);
    if (Number.isFinite(value) && value > 0) {
      currentDelayMs = value;
      log(`Using delay from parameter: ${currentDelayMs} ms`);
    }
  }
}

// ---------------------------------------------
// Subscriptions
// ---------------------------------------------
function subscribeToFilterChanges() {
  dashboard.worksheets.forEach((ws) => {
    if (WORKSHEETS_TO_SUBSCRIBE.includes(ws.name)) {
      ws.addEventListener(
        tableau.TableauEventType.FilterChanged,
        () => onSomethingChanged("FilterChanged", ws.name)
      );
      log(`Subscribed to FilterChanged on ${ws.name}`);
    } else {
      log(`Skipping worksheet: ${ws.name}`);
    }
  });
}

// Subscribe to Date Range Selector parameter only
function subscribeToParameterChanges() {
  dashboard.getParametersAsync().then(params => {
    params.forEach(p => {
      if (p.name === DATE_RANGE_PARAM_NAME) {
        p.addEventListener(
          tableau.TableauEventType.ParameterChanged,
          () => onSomethingChanged("ParameterChanged", p.name)
        );
        log(`Subscribed to ParameterChanged on "${DATE_RANGE_PARAM_NAME}"`);
      }
    });
  }).catch(err => {
    log(`Error subscribing to parameters: ${err.message || err}`);
  });
}

// ---------------------------------------------
// Change handler
// ---------------------------------------------
function onSomethingChanged(type, name) {
  log(`${type} event from: ${name}`);
  handleChange();
}

// shared logic for any change (filter or param)
function handleChange() {
  if (loadingTimeoutId) {
    clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
    log("Existing timer cleared (debounce).");
  }

  const tableContainer = insightsTableEl();
  if (tableContainer) tableContainer.style.opacity = "0.4";

  setAiStateSafe("Loading").then(() => {
    loadingTimeoutId = setTimeout(() => {
      Promise.all([
        setAiStateSafe("Ready"),
        refreshInsights()
      ])
        .then(() => {
          if (tableContainer) tableContainer.style.opacity = "1";
        })
        .catch(err => {
          log(`Error during refresh: ${err.message || err}`);
          if (tableContainer) tableContainer.style.opacity = "1";
        })
        .finally(() => { loadingTimeoutId = null; });
    }, currentDelayMs);
  });
}

async function setAiStateSafe(value) {
  try {
    if (!aiStateParam) return;
    log(`Setting "${AI_STATE_PARAM_NAME}" to "${value}"`);
    await aiStateParam.changeValueAsync(value);
  } catch (e) {
    log(`Error setting AI state to ${value} (ignored): ${e.message || e}`);
  }
}

// ---------------------------------------------
// Insights data & rendering
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

  try {
    log(`Fetching summary data from "${INSIGHTS_WORKSHEET_NAME}"`);
    const dataTable = await insightsSheet.getSummaryDataAsync();
    const cols = dataTable.columns;
    const rows = dataTable.data;

    renderInsightsAsCards(cols, rows);
  } catch (err) {
    log(`Error fetching insights data: ${err.message || err}`);
    tableContainer.innerHTML = "<em>Error loading insights data.</em>";
  }
}

function renderInsightsAsCards(columns, rows) {
  const tableContainer = insightsTableEl();
  if (!tableContainer) return;

  if (!rows || rows.length === 0) {
    tableContainer.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  // Map column names → index
  const colIndex = {};
  columns.forEach((c, idx) => {
    colIndex[c.fieldName] = idx;
  });

  const brandIdx   = colIndex["Brand"] ?? 0;
  const hcpIdx     = colIndex["Hcp Dtc Identifier"] ?? colIndex["HCP DTC Identifier"] ?? null;
  const srcIdx     = colIndex["Source"] ?? null;
  const dateIdx    = colIndex["Current Period Date Range"] ?? null;
  const insightKey = Object.keys(colIndex).find(name => name.includes("Insight"));
  const insightIdx = insightKey != null ? colIndex[insightKey] : null;

  let html = `<div class="insights-grid">`;

  rows.forEach(row => {
    const brand   = row[brandIdx]?.formattedValue || "";
    const hcpDtc  = hcpIdx != null ? (row[hcpIdx]?.formattedValue || "") : "";
    const source  = srcIdx != null ? (row[srcIdx]?.formattedValue || "") : "";
    const date    = dateIdx != null ? (row[dateIdx]?.formattedValue || "") : "";
    let insight   = insightIdx != null ? (row[insightIdx]?.formattedValue || "") : "";

    if (!insight || insight === "Null" || insight === "*") {
      insight = "No narrative available for this combination.";
    }

    // Escape for safe HTML in data attribute
    const safeInsight = insight
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

    html += `
      <div class="insight-card">
        <div class="insight-card-header">
          <div class="insight-brand">${brand || "—"}</div>
          ${hcpDtc ? `<div class="insight-badge">${hcpDtc}</div>` : ""}
          ${source ? `<div class="insight-badge">${source}</div>` : ""}
        </div>
        ${date ? `<div class="insight-date">${date}</div>` : ""}
        <div class="insight-body" data-fulltext="${safeInsight}"></div>
      </div>
    `;
  });

  html += `</div>`;
  tableContainer.innerHTML = html;

  // Start the typing animation once cards are in the DOM
  startTypingAnimation();
}

// ---------------------------------------------
// Typing animation helpers
// ---------------------------------------------
function startTypingAnimation() {
  const bodies = document.querySelectorAll(".insight-body");
  const typingSpeed = 20;   // ms per character → slower = more “AI-like”
  const staggerDelay = 150; // ms extra delay between cards

  bodies.forEach((el, idx) => {
    const full = el.getAttribute("data-fulltext");
    if (!full) return;

    // Clear existing content before typing
    el.textContent = "";

    const decoded = full
      .replace(/&quot;/g, '"')
      .replace(/&gt;/g, ">")
      .replace(/&lt;/g, "<")
      .replace(/&amp;/g, "&");

    // Slight stagger so all cards don’t start at the exact same moment
    setTimeout(() => {
      typeOneCard(el, decoded, 0, typingSpeed);
    }, idx * staggerDelay);
  });
}

function typeOneCard(el, text, index, speed) {
  if (!el) return;

  if (index >= text.length) {
    // Typing finished; now bold numeric values
    applyBoldNumbers(el);
    return;
  }

  el.textContent += text.charAt(index);
  setTimeout(() => typeOneCard(el, text, index + 1, speed), speed);
}

function applyBoldNumbers(el) {
  const plain = el.textContent;
  // Replace numbers with <strong>wrapped</strong> equivalents
  const html = plain.replace(NUM_REGEX, "<strong>$1</strong>");
  el.innerHTML = html;
}

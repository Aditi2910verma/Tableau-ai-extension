// ---------------------------------------------
// Configuration
// ---------------------------------------------

// Worksheet with the insights data
const INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend";

// Parameter that should also refresh insights
const DATE_RANGE_PARAM_NAME = "Date Range Selector";

// Typing / loading behaviour
const TITLE_TEXT = "AI Insights ready. Listening for filter changes...";
const STATUS_LOADING_TEXT = "Updating insights for your current selection…";
const CARD_TYPING_SPEED_MS = 15;         // letter speed for card text
const INITIAL_FADE_DELAY_MS = 800;       // first splash screen delay

let dashboard = null;
let isRefreshing = false;

// DOM helpers
const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");
const insightsTableEl = () => document.getElementById("insights-table");

// ---------------------------------------------
// Utility: logging
// ---------------------------------------------
function log(message) {
  console.log("[AI Insights Extension]", message);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${message}\n`;
}

// Simple typewriter for a single element's textContent
function typeText(element, text, speed = 30) {
  if (!element) return;
  element.textContent = "";
  let i = 0;
  const timer = setInterval(() => {
    element.textContent += text[i];
    i++;
    if (i >= text.length) {
      clearInterval(timer);
    }
  }, speed);
}

// Highlight numeric values (applied AFTER typing finishes)
function highlightNumbers(text) {
  if (!text) return "";

  // Optional currency symbol ($/£/€), then number with optional commas/decimals,
  // then optional unit (K, M, %, bps)
  const numberRegex = /([$£€]?\d[\d.,]*\s*(?:K|M|%|bps)?)/g;

  return text.replace(numberRegex, '<span class="insight-number">$1</span>');
}
// Animate one card body: type text, then bold numbers
function animateCardBody(element, fullText) {
  if (!element) return;
  element.textContent = "";
  if (!fullText) return;

  let i = 0;
  const chars = fullText.split("");
  const timer = setInterval(() => {
    element.textContent += chars[i];
    i++;
    if (i >= chars.length) {
      clearInterval(timer);
      // After typing completes, apply number highlighting
      element.innerHTML = highlightNumbers(element.textContent);
    }
  }, CARD_TYPING_SPEED_MS);
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const loadingScreen = document.getElementById("loading-screen");
  const insightsScreen = document.getElementById("insights-screen");

  // First splash: 0.8s, then show extension UI
  setTimeout(() => {
    if (loadingScreen) loadingScreen.style.display = "none";
    if (insightsScreen) {
      insightsScreen.style.display = "block";
      insightsScreen.classList.add("show");
      typeText(statusEl(), TITLE_TEXT);
    }
  }, INITIAL_FADE_DELAY_MS);

  // Try initializing Tableau Extensions API
  try {
    tableau.extensions.initializeAsync().then(() => {
      dashboard = tableau.extensions.dashboardContent.dashboard;
      log(`Dashboard: ${dashboard.name}`);

      subscribeToFilterChanges();
      subscribeToDateRangeParameter();

      // Initial load
      refreshInsights();
    }).catch(err => {
      console.error("Failed to initialize Tableau Extension:", err);
      const s = statusEl();
      if (s) s.textContent = "AI Insights extension failed to initialize.";
      log(`Init failed: ${err.message || err}`);
    });
  } catch (e) {
    // This branch is only hit when opened directly in a browser, not in Tableau
    console.warn("Tableau Extensions API not available; demo mode.", e);
    const s = statusEl();
    if (s) s.textContent = "Running outside Tableau (demo mode).";
    log("Tableau Extensions API not found. Logic disabled.");
  }
});

// ---------------------------------------------
// Subscriptions
// ---------------------------------------------
function subscribeToFilterChanges() {
  if (!dashboard) return;
  dashboard.worksheets.forEach(ws => {
    if (ws.name === INSIGHTS_WORKSHEET_NAME) {
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
  dashboard.getParametersAsync().then(params => {
    const dateParam = params.find(p => p.name === DATE_RANGE_PARAM_NAME);
    if (!dateParam) {
      log(`Parameter "${DATE_RANGE_PARAM_NAME}" not found (optional).`);
      return;
    }
    dateParam.addEventListener(
      tableau.TableauEventType.ParameterChanged,
      () => onSomethingChanged("parameter", dateParam.name)
    );
    log(`Subscribed to ParameterChanged on "${dateParam.name}"`);
  }).catch(err => {
    log(`Error subscribing to parameters: ${err.message || err}`);
  });
}

// Unified handler for any change
function onSomethingChanged(type, name) {
  log(`${type} changed: ${name}`);
  if (isRefreshing) {
    log("Refresh already in progress; skipping.");
    return;
  }
  handleRefresh();
}

async function handleRefresh() {
  isRefreshing = true;

  const s = statusEl();
  if (s) s.textContent = STATUS_LOADING_TEXT;

  try {
    await refreshInsights();
    // After data is rendered, re-type the title
    typeText(statusEl(), TITLE_TEXT);
  } catch (err) {
    log(`Error refreshing insights: ${err.message || err}`);
    if (s) s.textContent = "Error updating insights (see log).";
  } finally {
    isRefreshing = false;
  }
}

// ---------------------------------------------
// Data & rendering
// ---------------------------------------------
async function refreshInsights() {
  if (!dashboard) return;
  const container = insightsTableEl();
  if (!container) return;

  const sheet = dashboard.worksheets.find(
    ws => ws.name === INSIGHTS_WORKSHEET_NAME
  );

  if (!sheet) {
    log(`Worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.`);
    container.innerHTML =
      `<em>Insights worksheet "${INSIGHTS_WORKSHEET_NAME}" not found.</em>`;
    return;
  }

  log(`Fetching summary data from "${INSIGHTS_WORKSHEET_NAME}"…`);
  const dataTable = await sheet.getSummaryDataAsync();
  const cols = dataTable.columns;
  const rows = dataTable.data;

  renderInsightsCards(cols, rows);
}

function renderInsightsCards(columns, rows) {
  const container = insightsTableEl();
  if (!container) return;

  container.innerHTML = "";

  if (!rows || rows.length === 0) {
    container.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  // Map the key columns we care about
  const idxBrand = columns.findIndex(c => c.fieldName === "Brand");
  const idxHcpDtc = columns.findIndex(
    c => c.fieldName === "Hcp Dtc Identifier"
  );
  const idxSource = columns.findIndex(c => c.fieldName === "Source");
  const idxDate = columns.findIndex(
    c => c.fieldName === "Current Period Date Range"
  );
  // Adjust this if your insight text field has a slightly different name
  const idxInsight = columns.findIndex(
    c => c.fieldName.indexOf("Insight") !== -1
  );

  const grid = document.createElement("div");
  grid.className = "insights-grid";

  rows.forEach(row => {
    const brand = idxBrand >= 0 ? row[idxBrand].formattedValue : "";
    const hcpDtc = idxHcpDtc >= 0 ? row[idxHcpDtc].formattedValue : "";
    const source = idxSource >= 0 ? row[idxSource].formattedValue : "";
    const dateRange = idxDate >= 0 ? row[idxDate].formattedValue : "";
    const insightText =
      idxInsight >= 0 ? row[idxInsight].formattedValue : "";

    const card = document.createElement("div");
    card.className = "insight-card";

    const header = document.createElement("div");
    header.className = "insight-card-header";

    const brandEl = document.createElement("div");
    brandEl.className = "insight-brand";
    brandEl.textContent = brand || "—";

    header.appendChild(brandEl);

    if (hcpDtc) {
      const badge = document.createElement("span");
      badge.className = "insight-badge";
      badge.textContent = hcpDtc;
      header.appendChild(badge);
    }

    if (source) {
      const badge = document.createElement("span");
      badge.className = "insight-badge";
      badge.textContent = source;
      header.appendChild(badge);
    }

    const dateEl = document.createElement("div");
    dateEl.className = "insight-date";
    dateEl.textContent = dateRange || "";

    const body = document.createElement("div");
    body.className = "insight-body";

    card.appendChild(header);
    if (dateRange) card.appendChild(dateEl);
    card.appendChild(body);

    grid.appendChild(card);

    // Animate the insight text per card
    if (insightText) {
      animateCardBody(body, insightText);
    } else {
      body.innerHTML = "<em>No data</em>";
    }
  });

  container.appendChild(grid);
}


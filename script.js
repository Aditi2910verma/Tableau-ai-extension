// ---------------------------------------------
// Configuration
// ---------------------------------------------
const AI_STATE_PARAM_NAME = "p_AI_Insight_State";
const AI_DELAY_PARAM_NAME = null; // optional
const DEFAULT_LOADING_DELAY_MS = 1000;
const WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];

let dashboard = null;
let aiStateParam = null;
let loadingTimeoutId = null;
let currentDelayMs = DEFAULT_LOADING_DELAY_MS;

const statusEl = () => document.getElementById("status");
const logEl = () => document.getElementById("log");

// Utility: append a log line
function log(message) {
  console.log("[AI Loading Extension]", message);
  const el = logEl();
  if (!el) return;
  const time = new Date().toISOString().substr(11, 8);
  el.textContent += `[${time}] ${message}\n`;
}

// Typing animation for status text
function typeText(element, text, speed = 50) {
  if (!element) return;
  element.textContent = "";
  let i = 0;
  const interval = setInterval(() => {
    element.textContent += text[i];
    i++;
    if (i >= text.length) clearInterval(interval);
  }, speed);
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const loadingScreen = document.getElementById("loading-screen");
  const insightsScreen = document.getElementById("insights-screen");

  // UI transition: show loading for 0.8s, then fade in insights
 setTimeout(() => {
  if (loadingScreen) loadingScreen.style.display = "none";
  if (insightsScreen) {
    insightsScreen.style.display = "block";
    insightsScreen.classList.add("show");
    statusEl().textContent = "Initializing AI Insights…";
  }
}, 800);

  // Always try to initialize Tableau; if it fails, fall back to demo mode
  try {
    tableau.extensions.initializeAsync()
      .then(() => {
        dashboard = tableau.extensions.dashboardContent.dashboard;
       statusEl().textContent = "AI Insights ready. Listening for filter changes…";
        log(`Dashboard name: ${dashboard.name}`);

        return findAiStateParameter();
      })
      .then(() => maybeReadDelayParameter())
      .then(() => {
        subscribeToFilterChanges();
        typeText(statusEl(), "AI Insights ready. Listening for filter changes…");
      })
      .catch((err) => {
        console.error("Failed to initialize Tableau Extension:", err);
        typeText(statusEl(), "Failed to initialize extension (see log).");
        log(`Tableau init failed: ${err.message || err}`);
      });
  } catch (e) {
    // This happens when running in a normal browser, not inside Tableau
    console.warn("Not running inside Tableau, demo mode:", e);
    log("Tableau Extensions API not found. UI shown, but parameter logic disabled.");
    typeText(statusEl(), "Running outside Tableau (demo mode).");
  }
});


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

function subscribeToFilterChanges() {
  dashboard.worksheets.forEach((ws) => {
    if (WORKSHEETS_TO_SUBSCRIBE.includes(ws.name)) {
      ws.addEventListener(
        tableau.TableauEventType.FilterChanged,
        (event) => onFilterChanged(event, ws)
      );
      log(`Subscribed to FilterChanged on ${ws.name}`);
    } else {
      log(`Skipping worksheet: ${ws.name}`);
    }
  });
}

function onFilterChanged(filterEvent, worksheet) {
  log(`FilterChanged event fired on ${worksheet.name}`);
  if (loadingTimeoutId) {
    clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
    log("Existing timer cleared (debounce).");
  }
  setAiState("Loading").then(() => {
    loadingTimeoutId = setTimeout(() => {
      setAiState("Ready")
        .catch(err => log(`Error setting state to Ready: ${err.message || err}`))
        .finally(() => { loadingTimeoutId = null; });
    }, currentDelayMs);
  }).catch(err => {
    log(`Error setting state to Loading: ${err.message || err}`);
  });
}

async function setAiState(value) {
  if (!aiStateParam) throw new Error("AI state parameter not initialized.");
  log(`Setting "${AI_STATE_PARAM_NAME}" to "${value}"`);
  await aiStateParam.changeValueAsync(value);
}

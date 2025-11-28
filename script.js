// ---------------------------------------------
// Configuration
// ---------------------------------------------
var DEFAULT_LOADING_DELAY_MS = 800; // 0.8 seconds

// Worksheet that contains the insights data + filters
var INSIGHTS_WORKSHEET_NAME = "AI Insights- Estimated Spend";

// Worksheets that should trigger refresh when filters change
var WORKSHEETS_TO_SUBSCRIBE = ["AI Insights- Estimated Spend"];

// Parameter that controls date logic (your Date Range Selector)
var DATE_RANGE_PARAM_NAME = "Date Range Selector";

var dashboard = null;
var loadingTimeoutId = null;
var currentDelayMs = DEFAULT_LOADING_DELAY_MS;

function statusEl()        { return document.getElementById("status"); }
function logEl()           { return document.getElementById("log"); }
function insightsTableEl() { return document.getElementById("insights-table"); }

// ---------------------------------------------
// Utilities
// ---------------------------------------------
function log(message) {
  console.log("[AI Extension]", message);
  var el = logEl();
  if (!el) return;
  var time = new Date().toISOString().substr(11, 8);
  el.textContent += "[" + time + "] " + message + "\n";
}

// Escape text for HTML / attributes
function escapeHtml(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Simple helper for older browsers (no Array.includes)
function arrayContains(arr, value) {
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] === value) return true;
  }
  return false;
}

// ---------------------------------------------
// Initialization
// ---------------------------------------------
document.addEventListener("DOMContentLoaded", function () {
  var loadingScreen  = document.getElementById("loading-screen");
  var insightsScreen = document.getElementById("insights-screen");

  // Splash screen for 0.8s
  setTimeout(function () {
    if (loadingScreen) loadingScreen.style.display = "none";
    if (insightsScreen) {
      insightsScreen.style.display = "block";
      insightsScreen.classList.add("show");
      var st = statusEl();
      if (st) st.textContent = "AI Insights ready. Listening for filter changes...";
    }
  }, 800);

  try {
    tableau.extensions.initializeAsync()
      .then(function () {
        dashboard = tableau.extensions.dashboardContent.dashboard;
        log("Dashboard name: " + dashboard.name);
        dashboard.worksheets.forEach(function (ws) {
          log("Worksheet available: " + ws.name);
        });

        subscribeToFilterChanges();
        subscribeToDateRangeParameter();

        // Initial data load
        return refreshInsights();
      })
      .catch(function (err) {
        console.error("Failed to initialize Tableau Extension:", err);
        var st = statusEl();
        if (st) st.textContent = "Failed to initialize extension (see log).";
        log("Tableau init failed: " + (err.message || err));
      });
  } catch (e) {
    console.warn("Not running inside Tableau, demo mode:", e);
    log("Tableau Extensions API not found. UI shown, but logic disabled.");
    var st2 = statusEl();
    if (st2) st2.textContent = "Running outside Tableau (demo mode).";
  }
});

// ---------------------------------------------
// Subscriptions
// ---------------------------------------------
function subscribeToFilterChanges() {
  if (!dashboard) return;

  dashboard.worksheets.forEach(function (ws) {
    if (arrayContains(WORKSHEETS_TO_SUBSCRIBE, ws.name)) {
      ws.addEventListener(
        tableau.TableauEventType.FilterChanged,
        function () { onSomethingChanged("filter", ws.name); }
      );
      log("Subscribed to FilterChanged on " + ws.name);
    } else {
      log("Skipping worksheet: " + ws.name);
    }
  });
}

function subscribeToDateRangeParameter() {
  if (!dashboard) return;

  dashboard.getParametersAsync()
    .then(function (params) {
      var dateParam = null;
      for (var i = 0; i < params.length; i++) {
        if (params[i].name === DATE_RANGE_PARAM_NAME) {
          dateParam = params[i];
          break;
        }
      }

      if (!dateParam) {
        log('Date Range parameter "' + DATE_RANGE_PARAM_NAME + '" not found.');
        return;
      }

      log("Parameter available: " + dateParam.name);
      dateParam.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        function () { onSomethingChanged("parameter", dateParam.name); }
      );
      log('Subscribed to ParameterChanged on "' + dateParam.name + '"');
    })
    .catch(function (err) {
      log("Error subscribing to date parameter: " + (err.message || err));
    });
}

// ---------------------------------------------
// Change handler
// ---------------------------------------------
function onSomethingChanged(type, name) {
  log(type + " changed: " + name);

  // Debounce rapid changes
  if (loadingTimeoutId) {
    clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
    log("Existing timer cleared (debounce).");
  }

  var st    = statusEl();
  var table = insightsTableEl();

  if (st)    st.textContent = "Loading your Insight...";
  if (table) table.style.visibility = "hidden";

  loadingTimeoutId = setTimeout(function () {
    refreshInsights()
      .then(function () {
        if (table) table.style.visibility = "visible";
        if (st)    st.textContent = "AI Generated Insights";
      })
      .catch(function (err) {
        log("Error during refreshInsights: " + (err.message || err));
        if (st)    st.textContent = "Error updating AI Insights (see log).";
        if (table) table.style.visibility = "visible";
      })
      .finally(function () {
        loadingTimeoutId = null;
      });
  }, currentDelayMs);
}

// ---------------------------------------------
// Data & rendering
// ---------------------------------------------
function refreshInsights() {
  if (!dashboard) return Promise.resolve();

  var container = insightsTableEl();
  if (!container) return Promise.resolve();

  var sheet = null;
  dashboard.worksheets.forEach(function (ws) {
    if (ws.name === INSIGHTS_WORKSHEET_NAME) sheet = ws;
  });

  if (!sheet) {
    log('Insights worksheet "' + INSIGHTS_WORKSHEET_NAME + '" not found.');
    container.innerHTML =
      '<em>Insights worksheet "' + INSIGHTS_WORKSHEET_NAME + '" not found.</em>';
    return Promise.resolve();
  }

  log('Fetching summary data from "' + INSIGHTS_WORKSHEET_NAME + '"');

  return sheet.getSummaryDataAsync()
    .then(function (dataTable) {
      var cols = dataTable.columns;
      var rows = dataTable.data;
      renderInsightsAsCards(cols, rows);
    })
    .catch(function (err) {
      log("Error fetching summary data: " + (err.message || err));
      var st = statusEl();
      if (st) st.textContent = "Error loading insights (see log).";
    });
}

// Typewriter effect for all insight bodies
function animateInsightBodies(speedPerChar) {
  if (speedPerChar === undefined) speedPerChar = 35;

  var bodies = document.querySelectorAll(".insight-body[data-fulltext]");
  for (var i = 0; i < bodies.length; i++) {
    (function (el, index) {
      var fullText = el.getAttribute("data-fulltext") || "";
      el.textContent = "";

      var j = 0;
      var startDelay = index * 100; // stagger cards

      setTimeout(function () {
        var interval = setInterval(function () {
          el.textContent += fullText.charAt(j);
          j++;
          if (j >= fullText.length) {
            clearInterval(interval);
          }
        }, speedPerChar);
      }, startDelay);
    })(bodies[i], i);
  }
}

// Renders cards with typing effect
function renderInsightsAsCards(columns, rows) {
  var container = insightsTableEl();
  if (!container) return;

  if (!rows || rows.length === 0) {
    container.innerHTML = "<em>No insights for the current selection.</em>";
    return;
  }

  // Map column names -> index
  var colIndex = {};
  for (var i = 0; i < columns.length; i++) {
    colIndex[columns[i].fieldName] = i;
  }

  var brandIdx  = colIndex["Brand"];
  var hcpIdx    = colIndex["Hcp Dtc Identifier"];
  var sourceIdx = colIndex["Source"];
  var dateIdx   = colIndex["Current Period Date Range"];

  // Try several possible names for the insight column
  var insightIdx = null;
  var insightNames = [
    "Estimated Spend - Insight1",
    "Estimated Spend - Insight",
    "Estimated Spend - Insight 1"
  ];
  for (i = 0; i < insightNames.length; i++) {
    if (colIndex.hasOwnProperty(insightNames[i])) {
      insightIdx = colIndex[insightNames[i]];
      break;
    }
  }

  var html = "<div class='insights-grid'>";

  rows.forEach(function (row) {
    var brand   = brandIdx  != null ? row[brandIdx].formattedValue  : "";
    var hcp     = hcpIdx    != null ? row[hcpIdx].formattedValue    : "";
    var source  = sourceIdx != null ? row[sourceIdx].formattedValue : "";
    var date    = dateIdx   != null ? row[dateIdx].formattedValue   : "";
    var insight = insightIdx != null ? row[insightIdx].formattedValue : "";

    var safeInsight = insight || "No narrative for this combination.";

    html +=
      "<div class='insight-card'>" +
        "<div class='insight-card-header'>" +
          "<div class='insight-brand'>" + escapeHtml(brand || "—") + "</div>" +
          (hcp    ? "<span class='insight-badge'>" + escapeHtml(hcp) + "</span>" : "") +
          (source ? "<span class='insight-badge'>" + escapeHtml(source) + "</span>" : "") +
        "</div>" +
        (date ? "<div class='insight-date'>" + escapeHtml(date) + "</div>" : "") +
        "<div class='insight-body' data-fulltext='" + escapeHtml(safeInsight) + "'></div>" +
      "</div>";
  });

  html += "</div>";
  container.innerHTML = html;

  // Trigger typing animation for all cards
  animateInsightBodies();
}

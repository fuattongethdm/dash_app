// Links the "Daily Repair Amount" bar chart and the "Daily Production vs.
// Repair, and Backlog Trend" chart (pages/home.py, render_dashboard) so
// hovering a bar on one lightens the matching date's bars on the other —
// makes it easy to compare "how much" vs "how many" for the same day
// without hunting across two charts.
//
// Wired directly to Plotly's own plotly_hover/plotly_unhover events on each
// graph div (not through Dash's hoverData prop): hoverData didn't reliably
// reset to null on mouse-leave, which left the highlight color stuck. The
// native events fire symmetrically and don't round-trip through Dash at all.
window.dash_clientside = window.dash_clientside || {};
window.dash_clientside.clientside = window.dash_clientside.clientside || {};

function _dailySyncPlotDiv(id) {
    var wrapper = document.getElementById(id);
    return wrapper ? wrapper.querySelector(".js-plotly-plot") : null;
}

// "Daily Repair Amount" stays on a continuous date axis (x = ISO-ish
// datetime strings); the pipe chart uses a categorical axis with
// pre-formatted "dd.mm.yy" labels (see render_dashboard — a continuous
// date axis with very few days breaks grouped-bar width/offset). Reduce
// both to the same "yyyy-mm-dd" key so hover matching still works across
// the two different x-axis representations.
function _dailySyncDateKey(x) {
    var s = String(x);
    var m = s.match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
    if (m) {
        return "20" + m[3] + "-" + m[2] + "-" + m[1];
    }
    return s.slice(0, 10);
}

// traceIndices restricts the restyle to specific traces (e.g. just the bar
// traces, leaving a secondary-axis line trace's own color alone) — mirrors
// the Pareto sync functions below. Omitting it (amountDiv's single-bar
// chart) restyles every trace, same as before.
function _dailySyncResetColors(div, baseColor, traceIndices) {
    if (!div || !div.data || !div.data.length) return;
    var n = div.data[0].x.length;
    var colors = [];
    for (var i = 0; i < n; i++) colors.push(baseColor);
    Plotly.restyle(div, { "marker.color": [colors] }, traceIndices);
}

function _dailySyncHighlightMatch(div, baseColor, lightColor, hoveredDateKey, traceIndices) {
    if (!div || !div.data || !div.data.length) return;
    var xValues = div.data[0].x;
    var colors = xValues.map(function (x) {
        return _dailySyncDateKey(x) === hoveredDateKey ? lightColor : baseColor;
    });
    Plotly.restyle(div, { "marker.color": [colors] }, traceIndices);
}

function _wireDailyChartsSync() {
    var amountDiv = _dailySyncPlotDiv("daily-amount-graph");
    var pipeDiv = _dailySyncPlotDiv("daily-pipe-count-graph");
    if (!amountDiv || !pipeDiv || amountDiv._dailySyncWired) {
        return;
    }
    amountDiv._dailySyncWired = true;

    var AMOUNT_COLOR = "#2563eb";
    var AMOUNT_LIGHT = "rgba(37, 99, 235, 0.45)";
    // The pipe chart now has two bar traces (Produced=orange, Repaired=blue)
    // plus a stock-level line (same shared axis, drawn on top) — only
    // traces 0/1 (the bars) get restyled on hover-sync; the line trace (2)
    // keeps its own fixed color untouched.
    var PIPE_BAR_TRACES = [0, 1];
    var PIPE_PRODUCED_COLOR = "#f97316";
    var PIPE_PRODUCED_LIGHT = "rgba(249, 115, 22, 0.45)";
    var PIPE_REPAIRED_COLOR = "#2563eb";
    var PIPE_REPAIRED_LIGHT = "rgba(37, 99, 235, 0.45)";

    amountDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length) return;
        var hoveredDateKey = _dailySyncDateKey(evt.points[0].x);
        if (!pipeDiv || !pipeDiv.data || pipeDiv.data.length < 2) return;
        var producedColors = pipeDiv.data[0].x.map(function (x) {
            return _dailySyncDateKey(x) === hoveredDateKey ? PIPE_PRODUCED_LIGHT : PIPE_PRODUCED_COLOR;
        });
        var repairedColors = pipeDiv.data[1].x.map(function (x) {
            return _dailySyncDateKey(x) === hoveredDateKey ? PIPE_REPAIRED_LIGHT : PIPE_REPAIRED_COLOR;
        });
        Plotly.restyle(pipeDiv, { "marker.color": [producedColors, repairedColors] }, PIPE_BAR_TRACES);
    });
    amountDiv.on("plotly_unhover", function () {
        if (!pipeDiv || !pipeDiv.data || pipeDiv.data.length < 2) return;
        var n0 = pipeDiv.data[0].x.length;
        var n1 = pipeDiv.data[1].x.length;
        var producedColors = [];
        for (var i = 0; i < n0; i++) producedColors.push(PIPE_PRODUCED_COLOR);
        var repairedColors = [];
        for (var j = 0; j < n1; j++) repairedColors.push(PIPE_REPAIRED_COLOR);
        Plotly.restyle(pipeDiv, { "marker.color": [producedColors, repairedColors] }, PIPE_BAR_TRACES);
    });

    pipeDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length) return;
        _dailySyncHighlightMatch(amountDiv, AMOUNT_COLOR, AMOUNT_LIGHT, _dailySyncDateKey(evt.points[0].x));
    });
    pipeDiv.on("plotly_unhover", function () {
        _dailySyncResetColors(amountDiv, AMOUNT_COLOR);
    });
}

// Fired by a clientside callback whenever the Dashboard tab's content
// (re)renders. Figures paint asynchronously after that, so this retries
// briefly until both graph divs exist, then wires the native listeners
// above (idempotent — _dailySyncWired guards against attaching twice on a
// re-render of the same divs).
window.dash_clientside.clientside.wireDailyChartsSync = function () {
    var attempts = 0;
    var timer = setInterval(function () {
        attempts += 1;
        var amountDiv = _dailySyncPlotDiv("daily-amount-graph");
        var pipeDiv = _dailySyncPlotDiv("daily-pipe-count-graph");
        if ((amountDiv && pipeDiv) || attempts > 25) {
            clearInterval(timer);
            if (amountDiv && pipeDiv) {
                _wireDailyChartsSync();
            }
        }
    }, 200);
    return window.dash_clientside.no_update;
};

// Same cross-chart highlight for "Repair Amount Pareto" / "Repair Ratio
// Pareto". Both bar traces already use the same blue, so both directions
// share one base/light pair. Only trace 0 (the bar) is restyled — trace 1
// (the cumulative-% line) is left untouched — and matching is by exact
// project label text rather than date truncation, since the two Paretos
// are sorted by different metrics and don't share an x-axis order.
function _paretoSyncResetColors(div, baseColor) {
    if (!div || !div.data || !div.data.length) return;
    var n = div.data[0].x.length;
    var colors = [];
    for (var i = 0; i < n; i++) colors.push(baseColor);
    Plotly.restyle(div, { "marker.color": [colors] }, [0]);
}

function _paretoSyncHighlightMatch(div, baseColor, lightColor, hoveredLabel) {
    if (!div || !div.data || !div.data.length) return;
    var xValues = div.data[0].x;
    var colors = xValues.map(function (x) {
        return String(x) === hoveredLabel ? lightColor : baseColor;
    });
    Plotly.restyle(div, { "marker.color": [colors] }, [0]);
}

function _wireParetoChartsSync() {
    var amountDiv = _dailySyncPlotDiv("pareto-amount-graph");
    var ratioDiv = _dailySyncPlotDiv("pareto-ratio-graph");
    if (!amountDiv || !ratioDiv || amountDiv._paretoSyncWired) {
        return;
    }
    amountDiv._paretoSyncWired = true;

    var BAR_COLOR = "#2563eb";
    var BAR_LIGHT = "rgba(37, 99, 235, 0.45)";

    amountDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length || evt.points[0].curveNumber !== 0) return;
        _paretoSyncHighlightMatch(ratioDiv, BAR_COLOR, BAR_LIGHT, String(evt.points[0].x));
    });
    amountDiv.on("plotly_unhover", function () {
        _paretoSyncResetColors(ratioDiv, BAR_COLOR);
    });

    ratioDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length || evt.points[0].curveNumber !== 0) return;
        _paretoSyncHighlightMatch(amountDiv, BAR_COLOR, BAR_LIGHT, String(evt.points[0].x));
    });
    ratioDiv.on("plotly_unhover", function () {
        _paretoSyncResetColors(amountDiv, BAR_COLOR);
    });
}

window.dash_clientside.clientside.wireParetoChartsSync = function () {
    var attempts = 0;
    var timer = setInterval(function () {
        attempts += 1;
        var amountDiv = _dailySyncPlotDiv("pareto-amount-graph");
        var ratioDiv = _dailySyncPlotDiv("pareto-ratio-graph");
        if ((amountDiv && ratioDiv) || attempts > 25) {
            clearInterval(timer);
            if (amountDiv && ratioDiv) {
                _wireParetoChartsSync();
            }
        }
    }, 200);
    return window.dash_clientside.no_update;
};

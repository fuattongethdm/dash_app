// Links the "Daily Repair Amount" and "Pipes Repaired per Day" bar charts
// (pages/home.py, render_dashboard) so hovering a bar on one lightens the
// matching date's bar on the other — makes it easy to compare "how much"
// vs "how many" for the same day without hunting across two charts.
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

function _dailySyncResetColors(div, baseColor) {
    if (!div || !div.data || !div.data.length) return;
    var n = div.data[0].x.length;
    var colors = [];
    for (var i = 0; i < n; i++) colors.push(baseColor);
    Plotly.restyle(div, { "marker.color": [colors] });
}

function _dailySyncHighlightMatch(div, baseColor, lightColor, hoveredDate) {
    if (!div || !div.data || !div.data.length) return;
    var xValues = div.data[0].x;
    var colors = xValues.map(function (x) {
        return String(x).slice(0, 10) === hoveredDate ? lightColor : baseColor;
    });
    Plotly.restyle(div, { "marker.color": [colors] });
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
    var PIPE_COLOR = "#f97316";
    var PIPE_LIGHT = "rgba(249, 115, 22, 0.45)";

    amountDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length) return;
        _dailySyncHighlightMatch(pipeDiv, PIPE_COLOR, PIPE_LIGHT, String(evt.points[0].x).slice(0, 10));
    });
    amountDiv.on("plotly_unhover", function () {
        _dailySyncResetColors(pipeDiv, PIPE_COLOR);
    });

    pipeDiv.on("plotly_hover", function (evt) {
        if (!evt.points || !evt.points.length) return;
        _dailySyncHighlightMatch(amountDiv, AMOUNT_COLOR, AMOUNT_LIGHT, String(evt.points[0].x).slice(0, 10));
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

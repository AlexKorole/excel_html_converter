/*
 * chart-tools-pie.js
 *
 * Отдельный статичный SVG-рендер круговой диаграммы. Геометрия
 * принципиально другая, чем у line/bar (сектора по накопленному углу,
 * а не оси/сетка) — поэтому отдельный файл, а не ветка внутри
 * chart-tools.js. Без доп. функционала (без hover/тултипов) — статичный
 * снимок того, что посчитано в chart.xml, тот же принцип, что и у остальных.
 *
 * Использование: ChartToolsPie.render(containerEl, config) — config тот
 * же формат, что и для ChartTools.init(), плюс config.point_colors
 * ({индекс_кусочка: '#RRGGBB'}) — цвет кусочка, если задан явно в файле.
 */
(function (global) {
  'use strict';

  var PALETTE = ['#1a73e8', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6',
    '#16a085', '#e91e63', '#795548', '#3498db', '#e67e22'];

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function polarToCartesian(cx, cy, r, angleDeg) {
    var rad = (angleDeg - 90) * Math.PI / 180;
    return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
  }

  function render(container, config) {
    var W = container.clientWidth || 800;
    var H = container.clientHeight || 420;

    var series = (config.series && config.series[0]) || { values: [] };
    var categories = config.categories || [];
    var values = series.values || [];
    var pointColors = config.point_colors || {};

    var total = 0;
    values.forEach(function (v) { total += (v || 0); });

    var hasLegend = !!config.has_legend;
    var cx = hasLegend ? W * 0.35 : W / 2;
    var cy = config.title ? H / 2 + 10 : H / 2;
    var r = Math.max(Math.min(cx, cy, W - cx) - 20, 10);

    var svg = [];
    svg.push('<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" ' +
      'font-family="-apple-system, Segoe UI, sans-serif" width="100%" height="100%">');

    if (config.title) {
      svg.push('<text x="' + (W / 2) + '" y="24" text-anchor="middle" font-size="15" ' +
        'font-weight="600" fill="#1a1a1a">' + escapeHtml(config.title) + '</text>');
    }

    var slices = [];
    if (total > 0) {
      var startAngle = 0;
      values.forEach(function (v, i) {
        if (!v) return;
        var angle = (v / total) * 360;
        var endAngle = startAngle + angle;
        var color = pointColors[i] || PALETTE[i % PALETTE.length];
        var largeArc = angle > 180 ? 1 : 0;
        var p1 = polarToCartesian(cx, cy, r, startAngle);
        var p2 = polarToCartesian(cx, cy, r, endAngle);
        var path = 'M ' + cx + ',' + cy +
          ' L ' + p1[0] + ',' + p1[1] +
          ' A ' + r + ',' + r + ' 0 ' + largeArc + ' 1 ' + p2[0] + ',' + p2[1] + ' Z';
        svg.push('<path d="' + path + '" fill="' + color + '" stroke="#fff" stroke-width="1.5"/>');
        slices.push({ color: color, label: categories[i] || '', value: v, pct: v / total * 100 });
        startAngle = endAngle;
      });
    }

    if (hasLegend) {
      var legendX = cx + r + 30;
      var legendY = Math.max(cy - r, 30);
      slices.forEach(function (s) {
        svg.push('<rect x="' + legendX + '" y="' + (legendY - 10) + '" width="12" height="12" fill="' + s.color + '"/>');
        svg.push('<text x="' + (legendX + 18) + '" y="' + legendY + '" font-size="12" fill="#333">' +
          escapeHtml(s.label) + ' (' + s.pct.toFixed(1) + '%)</text>');
        legendY += 22;
      });
    }

    svg.push('</svg>');
    container.innerHTML = svg.join('');
  }

  global.ChartToolsPie = { render: render };
})(window);

/*
 * chart-tools.js
 *
 * Статичный SVG-рендер графика по уже готовым данным (без интерактива —
 * снимок того, что посчитано в chart.xml). Поддерживает line и bar
 * (вертикальный/column). Горизонтальный bar (barDir="bar") пока НЕ
 * поддерживается — явно предупреждает в консоли, а не рисует неверно.
 *
 * Использование:
 *   ChartTools.init(containerEl, {
 *     chart_type: 'line' | 'bar',
 *     bar_dir: 'col' | 'bar' | null,
 *     title: string | null,
 *     has_legend: boolean,
 *     categories: [string, ...],
 *     series: [{ name: string, values: [number|null, ...] }, ...]
 *   });
 */
(function (global) {
  'use strict';

  var PALETTE = ['#1a73e8', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#16a085', '#e91e63', '#795548'];

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function formatNumber(v) {
    if (Number.isInteger(v)) return v.toLocaleString('ru-RU');
    return v.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function niceNumber(range, round) {
    if (range <= 0) return 1;
    var exponent = Math.floor(Math.log10(range));
    var fraction = range / Math.pow(10, exponent);
    var niceFraction;
    if (round) {
      niceFraction = fraction < 1.5 ? 1 : fraction < 3 ? 2 : fraction < 7 ? 5 : 10;
    } else {
      niceFraction = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
    }
    return niceFraction * Math.pow(10, exponent);
  }

  function niceScale(dataMin, dataMax, tickCount) {
    if (dataMin === dataMax) {
      dataMin -= 1;
      dataMax += 1;
    }
    var range = niceNumber(dataMax - dataMin, false);
    var step = niceNumber(range / (tickCount - 1), true);
    var min = Math.floor(dataMin / step) * step;
    var max = Math.ceil(dataMax / step) * step;
    return { min: min, max: max, step: step };
  }

  function render(container, config) {
    if (config.chart_type === 'bar' && config.bar_dir === 'bar') {
      console.warn('ChartTools: горизонтальные bar-графики (barDir="bar") пока не поддерживаются, ' +
        'рендерю как вертикальные — оси могут не соответствовать оригиналу.');
    }

    var W = container.clientWidth || 800;
    var H = container.clientHeight || 420;
    var margin = { top: config.title ? 40 : 20, right: 20, bottom: 70, left: 70 };
    var innerW = Math.max(W - margin.left - margin.right, 10);
    var innerH = Math.max(H - margin.top - margin.bottom, 10);

    var categories = config.categories || [];
    var series = config.series || [];

    var allValues = [];
    series.forEach(function (s) {
      (s.values || []).forEach(function (v) {
        if (v !== null && v !== undefined) allValues.push(v);
      });
    });
    var dataMin = allValues.length ? Math.min.apply(null, allValues) : 0;
    var dataMax = allValues.length ? Math.max.apply(null, allValues) : 1;

    // Excel сам решает, начинать ли ось с нуля — bar не обязан включать 0
    // (проверено на реальном файле: у Excel valAx scaling там же, без
    // принудительного нуля, авто-масштаб подстраивается под данные).
    // Если в файле явно задан min/max оси — используем как есть, без
    // собственных вычислений.
    var minVal, maxVal, tickStep;
    if (config.axis_min !== null && config.axis_min !== undefined &&
        config.axis_max !== null && config.axis_max !== undefined) {
      minVal = config.axis_min;
      maxVal = config.axis_max;
      tickStep = niceNumber((maxVal - minVal) / 5, true);
    } else {
      var scale = niceScale(dataMin, dataMax, 6);
      minVal = scale.min;
      maxVal = scale.max;
      tickStep = scale.step;
    }

    function yFor(v) {
      return innerH - ((v - minVal) / (maxVal - minVal)) * innerH;
    }

    var n = categories.length;
    var step = n > 0 ? innerW / n : innerW;

    var svg = [];
    svg.push('<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" ' +
      'font-family="-apple-system, Segoe UI, sans-serif" width="100%" height="100%">');

    if (config.title) {
      svg.push('<text x="' + (W / 2) + '" y="24" text-anchor="middle" font-size="15" ' +
        'font-weight="600" fill="#1a1a1a">' + escapeHtml(config.title) + '</text>');
    }

    svg.push('<g transform="translate(' + margin.left + ',' + margin.top + ')">');

    var tickCount = Math.round((maxVal - minVal) / tickStep);
    for (var t = 0; t <= tickCount; t++) {
      var val = minVal + tickStep * t;
      var y = yFor(val);
      svg.push('<line x1="0" y1="' + y + '" x2="' + innerW + '" y2="' + y + '" stroke="#eee" stroke-width="1"/>');
      svg.push('<text x="-8" y="' + (y + 4) + '" text-anchor="end" font-size="11" fill="#888">' +
        formatNumber(Math.round(val * 100) / 100) + '</text>');
    }
    svg.push('<line x1="0" y1="' + innerH + '" x2="' + innerW + '" y2="' + innerH + '" stroke="#ccc" stroke-width="1"/>');

    for (var i = 0; i < n; i++) {
      var cx = step * i + step / 2;
      var label = String(categories[i] || '');
      var rotate = n > 8;
      var attrs = rotate
        ? 'transform="rotate(-35 ' + cx + ' ' + (innerH + 14) + ')" text-anchor="end"'
        : 'text-anchor="middle"';
      svg.push('<text x="' + cx + '" y="' + (innerH + 18) + '" font-size="11" fill="#666" ' + attrs + '>' +
        escapeHtml(label) + '</text>');
    }

    if (config.chart_type === 'line') {
      series.forEach(function (s, si) {
        var color = (s.color && s.color.charAt(0) === '#') ? s.color : PALETTE[si % PALETTE.length];
        var points = (s.values || []).map(function (v, idx) {
          if (v === null || v === undefined) return null;
          return [step * idx + step / 2, yFor(v)];
        });
        var path = '';
        points.forEach(function (p) {
          if (!p) return;
          path += (path === '' ? 'M' : 'L') + p[0] + ',' + p[1] + ' ';
        });
        svg.push('<path d="' + path + '" fill="none" stroke="' + color + '" stroke-width="2"/>');
        points.forEach(function (p) {
          if (!p) return;
          svg.push('<circle cx="' + p[0] + '" cy="' + p[1] + '" r="3" fill="' + color + '"/>');
        });
      });
    } else if (config.chart_type === 'bar') {
      var groupW = step * 0.7;
      var barW = groupW / Math.max(series.length, 1);
      series.forEach(function (s, si) {
        var color = (s.color && s.color.charAt(0) === '#') ? s.color : PALETTE[si % PALETTE.length];
        (s.values || []).forEach(function (v, idx) {
          if (v === null || v === undefined) return;
          var groupX = step * idx + (step - groupW) / 2;
          var x = groupX + si * barW;
          // Столбик рисуется от низа видимой области графика (= minVal)
          // до значения — НЕ от нуля: если ось не включает 0 (см. выше),
          // "высота от нуля" увела бы столбик далеко за пределы графика.
          var y1 = yFor(v);
          var barY = y1;
          var barH = innerH - y1;
          svg.push('<rect x="' + x + '" y="' + barY + '" width="' + (barW * 0.85) +
            '" height="' + barH + '" fill="' + color + '"/>');
        });
      });
    }

    svg.push('</g>');

    if (config.has_legend && series.length) {
      var legendY = H - 20;
      var lx = margin.left;
      series.forEach(function (s, si) {
        var color = (s.color && s.color.charAt(0) === '#') ? s.color : PALETTE[si % PALETTE.length];
        var name = s.name || '';
        svg.push('<rect x="' + lx + '" y="' + (legendY - 10) + '" width="10" height="10" fill="' + color + '"/>');
        svg.push('<text x="' + (lx + 14) + '" y="' + legendY + '" font-size="12" fill="#333">' +
          escapeHtml(name) + '</text>');
        lx += 14 + name.length * 7 + 24;
      });
    }

    svg.push('</svg>');
    container.innerHTML = svg.join('');
  }

  global.ChartTools = {
    init: function (container, config) {
      if (config.chart_type === 'pie' && global.ChartToolsPie) {
        global.ChartToolsPie.render(container, config);
        window.addEventListener('resize', function () { global.ChartToolsPie.render(container, config); });
        return { rerender: function () { global.ChartToolsPie.render(container, config); } };
      }
      render(container, config);
      window.addEventListener('resize', function () { render(container, config); });
      return { rerender: function () { render(container, config); } };
    }
  };
})(window);

/*
 * table-tools.js
 *
 * Внешние sort/filter для плоских таблиц (не PivotGrid — тот для сводных).
 * Работает поверх готового массива строк, переданного генератором
 * (xlsx_table_to_page.py) в TABLE_COLUMNS/TABLE_ROWS. Рассчитан на объёмы
 * до ~5000 строк (см. README/комментарии в xlsx_table_to_page.py) — выше
 * этого порога прямые DOM-операции сортировки/фильтра начинают заметно
 * тормозить, нужна виртуализация, которой тут нет.
 *
 * Формат данных на входе:
 *   columns: [{key, html, type}], type: 'number' | 'date' | 'bool' | 'text'
 *   rows:    [{ [key]: { v: <сравниваемое значение>, h: <готовый HTML для показа> } }, ...]
 *
 * Использование:
 *   TableTools.init(containerEl, TABLE_COLUMNS, TABLE_ROWS);
 */
(function (global) {
  'use strict';

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function cellKey(v) {
    return v === null || v === undefined ? '\u0000__blank__' : String(v);
  }

  function compareValues(a, b, type) {
    if (a === null && b === null) return 0;
    if (a === null) return -1;
    if (b === null) return 1;
    if (type === 'number') return a - b;
    if (type === 'bool') return (a === b) ? 0 : (a ? 1 : -1);
    return String(a).localeCompare(String(b), 'ru');
  }

  function uniqueValues(rows, key) {
    var seen = new Map();
    rows.forEach(function (row) {
      var cell = row[key];
      var v = cell ? cell.v : null;
      var h = cell ? cell.h : '';
      var k = cellKey(v);
      if (!seen.has(k)) {
        seen.set(k, { v: v, h: h === '' ? '(пусто)' : h });
      }
    });
    return Array.from(seen.values());
  }

  function TableTools(container, columns, rows) {
    this.container = container;
    this.columns = columns;
    this.rows = rows;
    this.sortKey = null;
    this.sortDir = 1;
    this.filters = {}; // key -> Set(строковых ключей разрешённых значений) | отсутствует = все
    this.currentPanel = null;
    this._outsideClickHandler = null;
    this.render();
  }

  TableTools.prototype.columnType = function (key) {
    var col = this.columns.filter(function (c) { return c.key === key; })[0];
    return col ? col.type : 'text';
  };

  TableTools.prototype.visibleRows = function () {
    var self = this;
    var filtered = this.rows.filter(function (row) {
      for (var key in self.filters) {
        var allowed = self.filters[key];
        if (!allowed) continue;
        var cell = row[key];
        var k = cellKey(cell ? cell.v : null);
        if (!allowed.has(k)) return false;
      }
      return true;
    });
    if (this.sortKey) {
      var key = this.sortKey, type = this.columnType(key), dir = this.sortDir;
      filtered = filtered.slice().sort(function (a, b) {
        var av = a[key] ? a[key].v : null;
        var bv = b[key] ? b[key].v : null;
        return compareValues(av, bv, type) * dir;
      });
    }
    return filtered;
  };

  TableTools.prototype.render = function () {
    var self = this;
    var visible = this.visibleRows();

    var theadCells = this.columns.map(function (col) {
      var isSorted = self.sortKey === col.key;
      var isFiltered = !!self.filters[col.key];
      var isActive = isSorted || isFiltered;
      var sortMark = isSorted ? (self.sortDir === 1 ? ' &#9650;' : ' &#9660;') : '';
      var activeCls = isFiltered ? ' tt-filter-active' : '';
      var clearVisibleCls = isActive ? ' tt-clear-visible' : '';
      return '<th>' +
        '<span class="tt-th-label" data-key="' + col.key + '">' + col.html + sortMark + '</span>' +
        '<button class="tt-filter-btn' + activeCls + '" data-key="' + col.key + '" aria-label="Фильтр">&#9662;</button>' +
        '<button class="tt-clear-btn' + clearVisibleCls + '" data-key="' + col.key + '" aria-label="Сбросить" title="Сбросить сортировку/фильтр">&times;</button>' +
        '</th>';
    }).join('');

    var tbodyRows = visible.map(function (row) {
      var tds = self.columns.map(function (col) {
        var cell = row[col.key];
        var cls = (col.type === 'number' || col.type === 'date') ? ' class="num"' : '';
        return '<td' + cls + '>' + (cell ? cell.h : '') + '</td>';
      }).join('');
      return '<tr>' + tds + '</tr>';
    }).join('');

    this.container.innerHTML =
      '<table class="tt-table"><thead><tr>' + theadCells + '</tr></thead>' +
      '<tbody>' + tbodyRows + '</tbody></table>' +
      '<div class="tt-status">Показано строк: ' + visible.length + ' из ' + this.rows.length + '</div>';

    Array.prototype.forEach.call(this.container.querySelectorAll('.tt-th-label'), function (el) {
      el.addEventListener('click', function () {
        var key = el.getAttribute('data-key');
        if (self.sortKey === key) {
          self.sortDir = -self.sortDir;
        } else {
          self.sortKey = key;
          self.sortDir = 1;
        }
        self.render();
      });
    });

    Array.prototype.forEach.call(this.container.querySelectorAll('.tt-filter-btn'), function (el) {
      el.addEventListener('click', function (e) {
        e.stopPropagation();
        self.openFilterDropdown(el.getAttribute('data-key'), el);
      });
    });

    Array.prototype.forEach.call(this.container.querySelectorAll('.tt-clear-btn'), function (el) {
      el.addEventListener('click', function (e) {
        e.stopPropagation();
        var key = el.getAttribute('data-key');
        if (self.sortKey === key) {
          self.sortKey = null;
          self.sortDir = 1;
        }
        delete self.filters[key];
        self.render();
      });
    });
  };

  TableTools.prototype.closeDropdown = function () {
    if (this.currentPanel) {
      this.currentPanel.remove();
      this.currentPanel = null;
    }
    if (this._outsideClickHandler) {
      document.removeEventListener('click', this._outsideClickHandler);
      this._outsideClickHandler = null;
    }
  };

  TableTools.prototype.openFilterDropdown = function (key, anchorEl) {
    var self = this;
    this.closeDropdown();

    var allValues = uniqueValues(this.rows, key);
    var type = this.columnType(key);
    allValues.sort(function (a, b) { return compareValues(a.v, b.v, type); });
    var currentAllowed = this.filters[key];

    var panel = document.createElement('div');
    panel.className = 'tt-filter-panel';

    var searchWrap = document.createElement('div');
    searchWrap.className = 'tt-filter-search';
    var modeId = 'tt-mode-' + key.replace(/[^a-zA-Z0-9]/g, '_');
    searchWrap.innerHTML =
      '<input type="text" class="tt-filter-search-input" placeholder="Поиск...">' +
      '<label><input type="radio" name="' + modeId + '" value="contains" checked> Содержит</label>' +
      '<label><input type="radio" name="' + modeId + '" value="starts"> Начинается с</label>';
    panel.appendChild(searchWrap);

    var listWrap = document.createElement('div');
    listWrap.className = 'tt-filter-list';
    panel.appendChild(listWrap);

    var footerWrap = document.createElement('div');
    footerWrap.className = 'tt-filter-footer';
    footerWrap.innerHTML =
      '<button class="tt-filter-ok">OK</button>' +
      '<button class="tt-filter-cancel">Отмена</button>';
    panel.appendChild(footerWrap);

    function renderList(query, mode) {
      var filtered = allValues.filter(function (item) {
        if (!query) return true;
        var h = String(item.h).toLowerCase();
        var q = query.toLowerCase();
        return mode === 'starts' ? h.indexOf(q) === 0 : h.indexOf(q) !== -1;
      });
      var shown = filtered.slice(0, 50);
      var truncated = filtered.length > 50;

      listWrap.innerHTML =
        '<label class="tt-filter-item tt-select-all"><input type="checkbox" class="tt-select-all-cb"> (Выбрать все)</label>' +
        shown.map(function (item) {
          var k = cellKey(item.v);
          var checked = !currentAllowed || currentAllowed.has(k);
          return '<label class="tt-filter-item"><input type="checkbox" data-k="' + escapeHtml(k) + '"' +
            (checked ? ' checked' : '') + '> ' + escapeHtml(item.h) + '</label>';
        }).join('') +
        (truncated ? '<div class="tt-filter-truncated">Показаны первые 50 из ' + filtered.length + ' — уточните поиск</div>' : '');

      var selectAllCb = listWrap.querySelector('.tt-select-all-cb');
      var itemCbs = Array.prototype.slice.call(listWrap.querySelectorAll('input[data-k]'));
      selectAllCb.checked = itemCbs.length > 0 && itemCbs.every(function (cb) { return cb.checked; });
      selectAllCb.addEventListener('change', function () {
        itemCbs.forEach(function (cb) { cb.checked = selectAllCb.checked; });
      });
    }

    var searchInput = searchWrap.querySelector('.tt-filter-search-input');
    var currentMode = 'contains';
    renderList('', currentMode);

    searchInput.addEventListener('input', function () { renderList(searchInput.value, currentMode); });
    Array.prototype.forEach.call(searchWrap.querySelectorAll('input[type=radio]'), function (r) {
      r.addEventListener('change', function () {
        currentMode = r.value;
        renderList(searchInput.value, currentMode);
      });
    });

    footerWrap.querySelector('.tt-filter-ok').addEventListener('click', function () {
      var checked = Array.prototype.slice.call(listWrap.querySelectorAll('input[data-k]:checked'))
        .map(function (cb) { return cb.getAttribute('data-k'); });
      if (checked.length === allValues.length) {
        delete self.filters[key];
      } else {
        self.filters[key] = new Set(checked);
      }
      self.closeDropdown();
      self.render();
    });
    footerWrap.querySelector('.tt-filter-cancel').addEventListener('click', function () {
      self.closeDropdown();
    });

    document.body.appendChild(panel);
    var rect = anchorEl.getBoundingClientRect();
    var panelWidth = panel.offsetWidth;
    var viewportWidth = window.innerWidth || document.documentElement.clientWidth;

    var left;
    if (rect.left + panelWidth > viewportWidth) {
      // не помещается по правому краю — прижимаем панель правым краем к кнопке
      left = window.scrollX + rect.right - panelWidth;
      if (left < 0) left = window.scrollX + 4; // совсем узкий экран — хотя бы не за пределами слева
    } else {
      left = window.scrollX + rect.left;
    }

    panel.style.position = 'absolute';
    panel.style.top = (window.scrollY + rect.bottom + 4) + 'px';
    panel.style.left = left + 'px';
    this.currentPanel = panel;

    setTimeout(function () {
      self._outsideClickHandler = function (e) {
        if (!panel.contains(e.target)) self.closeDropdown();
      };
      document.addEventListener('click', self._outsideClickHandler);
    }, 0);
  };

  global.TableTools = {
    init: function (container, columns, rows) {
      return new TableTools(container, columns, rows);
    }
  };
})(window);

"""
report_nav.py

Общий кусок навигации ("бургер-меню") для многостраничного отчёта.
Один и тот же паттерн вставляется на каждую сгенерированную страницу
(таблица / сводная / график) — список страниц у каждого конкретного
отчёта свой, поэтому это не статичный файл, а генерируемый фрагмент.
"""

from report_messages import rt

NAV_CSS = """
.report-nav-bar {
  display: flex;
  align-items: center;
  height: 48px;
  flex-shrink: 0;
  padding: 0 4px;
  position: relative; /* точка отсчёта для выпадающего меню */
}

.report-nav-toggle {
  width: 36px;
  height: 36px;
  border: none;
  border-radius: 8px;
  background: #1a1a1a;
  color: #fff;
  font-size: 18px;
  cursor: pointer;
  flex-shrink: 0;
}
.report-nav-toggle:hover { background: #333; }

.report-nav-menu {
  position: absolute;
  top: 100%;
  left: 4px;
  z-index: 999;
  background: #fff;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.15);
  min-width: 180px;
  overflow: hidden;
  display: none;
  font-family: -apple-system, 'Segoe UI', sans-serif;
}
.report-nav-menu.open { display: block; }

.report-nav-menu a {
  display: block;
  padding: 10px 16px;
  text-decoration: none;
  color: #1a1a1a;
  font-size: 14px;
  border-bottom: 1px solid #f0f0f0;
}
.report-nav-menu a:last-child { border-bottom: none; }
.report-nav-menu a:hover { background: #f4f5f7; }
.report-nav-menu a.current {
  font-weight: 600;
  background: #eef3ff;
  color: #1a73e8;
}

.report-nav-menu a.report-nav-back {
  background: #f9fafb;
  font-weight: 600;
  color: #555;
  border-bottom: 2px solid #ddd;
}
.report-nav-menu a.report-nav-back:hover { background: #f0f0f0; }
"""

NAV_JS = """
(function () {
  var btn = document.getElementById('report-nav-toggle');
  var menu = document.getElementById('report-nav-menu');
  if (!btn || !menu) return;
  btn.addEventListener('click', function (e) {
    e.stopPropagation();
    menu.classList.toggle('open');
  });
  document.addEventListener('click', function (e) {
    if (!menu.contains(e.target) && e.target !== btn) {
      menu.classList.remove('open');
    }
  });
})();
"""


def render_nav(pages, current_filename):
    """
    pages: список (label, filename), например
        [('Таблица', 'table.html'), ('Сводная', 'pivot.html'), ('График', 'chart.html')]
    current_filename: имя файла текущей страницы — подсвечивается в меню.

    Возвращает (css, html, js) — вставлять соответственно в <style>, <body> и <script>.
    """
    links = []
    for label, filename in pages:
        cls = ' class="current"' if filename == current_filename else ''
        links.append(f'    <a href="{filename}"{cls}>{label}</a>')

    html = (
        '<div class="report-nav-bar">\n'
        '  <button id="report-nav-toggle" class="report-nav-toggle" aria-label="Меню">&#9776;</button>\n'
        '  <nav id="report-nav-menu" class="report-nav-menu">\n'
        f'    <a href="/" class="report-nav-back">{rt("back_to_reports")}</a>\n'
        + '\n'.join(links) + '\n'
        '  </nav>\n'
        '</div>'
    )
    return NAV_CSS, html, NAV_JS

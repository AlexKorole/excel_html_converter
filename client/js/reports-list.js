(function () {
  const listEl = document.getElementById('report-list');
  let pollTimer = null;

  function escapeHtml(s) {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  function formatElapsed(createdAt) {
    const seconds = Math.max(0, Math.floor((Date.now() - new Date(createdAt).getTime()) / 1000));
    const s = I18N.t('seconds_short');
    if (seconds < 60) return `${seconds} ${s}`;
    const minutes = Math.floor(seconds / 60);
    return `${minutes} ${I18N.t('minutes_short')} ${seconds % 60} ${s}`;
  }

  function renderItem(r) {
    let nameHtml, statusHtml;
    if (r.status === 'ready') {
      nameHtml = `<a href="/server/reports/${r.id}/index.html">${escapeHtml(r.name)}</a>`;
      statusHtml = '';
    } else if (r.status === 'processing') {
      nameHtml = escapeHtml(r.name);
      statusHtml = `<span class="report-status status-processing">${I18N.t('processing')} (${formatElapsed(r.created_at)})</span>`;
    } else {
      nameHtml = escapeHtml(r.name);
      statusHtml = `<span class="report-status status-error">${I18N.t('error_prefix')} ${escapeHtml(r.error || '')}</span>`;
    }

    const li = document.createElement('li');
    li.className = 'report-item';
    const disabledAttr = r.status === 'processing' ? 'disabled' : '';
    li.innerHTML = `
      <div><span class="report-name">${nameHtml}</span>${statusHtml}</div>
      <button class="btn btn-danger" data-id="${r.id}" data-name="${escapeHtml(r.name)}" ${disabledAttr}>${I18N.t('delete_btn')}</button>
    `;
    li.querySelector('button').addEventListener('click', async (e) => {
      const id = e.target.getAttribute('data-id');
      const name = e.target.getAttribute('data-name');
      if (!confirm(I18N.t('confirm_delete', { name }))) return;
      await Api.deleteReport(id);
      await refresh();
    });
    return li;
  }

  async function refresh() {
    let reports;
    try {
      reports = await Api.listReports();
    } catch (e) {
      listEl.innerHTML = `<li>${escapeHtml(e.message)}</li>`;
      return;
    }

    listEl.innerHTML = '';
    if (reports.length === 0) {
      listEl.innerHTML = `<li>${I18N.t('no_reports_yet')}</li>`;
    } else {
      reports.forEach((r) => listEl.appendChild(renderItem(r)));
    }

    // Пока хоть один отчёт в процессе сборки - опрашиваем сервер, чтобы
    // статус обновился сам, без перезагрузки страницы пользователем.
    const stillProcessing = reports.some((r) => r.status === 'processing');
    if (stillProcessing && !pollTimer) {
      pollTimer = setInterval(refresh, CONFIG.POLL_INTERVAL_MS);
    } else if (!stillProcessing && pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }
  }

  refresh();
})();

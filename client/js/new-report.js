(function () {
  const form = document.getElementById('new-report-form');
  const errorEl = document.getElementById('form-error');
  const titleEl = document.getElementById('form-title');
  const pageTitleEl = document.getElementById('page-title');
  const submitBtn = document.getElementById('submit-btn');
  const nameInput = document.getElementById('name-input');

  const params = new URLSearchParams(window.location.search);
  const reportId = params.get('report_id');
  const reportName = params.get('name');

  if (reportId) {
    // Режим обновления существующего отчёта: имя зафиксировано (id и имя
    // в списке не меняются), редактировать его через эту форму нельзя -
    // поле disabled, поэтому браузер вообще не отправит его на сервер.
    if (reportName !== null) nameInput.value = reportName;
    nameInput.disabled = true;

    // data-i18n убираем, иначе I18N.apply() на DOMContentLoaded перезатрёт
    // текст обратно на дефолтный ("Новый отчёт"/"Создать").
    [titleEl, pageTitleEl, submitBtn].forEach((el) => el && el.removeAttribute('data-i18n'));
    if (titleEl) titleEl.textContent = I18N.t('update_report_title');
    if (pageTitleEl) pageTitleEl.textContent = I18N.t('update_report_title');
    if (submitBtn) submitBtn.textContent = I18N.t('update_btn');
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorEl.style.display = 'none';

    const formData = new FormData(form);
    try {
      if (reportId) {
        await Api.updateReport(reportId, formData);
      } else {
        await Api.createReport(formData);
      }
      window.location.href = '/';
    } catch (err) {
      errorEl.textContent = err.message;
      errorEl.style.display = 'block';
    }
  });
})();

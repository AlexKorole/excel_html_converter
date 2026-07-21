(function () {
  const form = document.getElementById('new-report-form');
  const errorEl = document.getElementById('form-error');

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorEl.style.display = 'none';

    const formData = new FormData(form);
    try {
      await Api.createReport(formData);
      window.location.href = '/';
    } catch (err) {
      errorEl.textContent = err.message;
      errorEl.style.display = 'block';
    }
  });
})();

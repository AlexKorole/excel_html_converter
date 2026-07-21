const Api = {
  async listReports() {
    const res = await fetch('/api/reports');
    if (!res.ok) throw new Error('Не удалось получить список отчётов');
    return res.json();
  },

  async createReport(formData) {
    const res = await fetch('/api/reports', { method: 'POST', body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Не удалось создать отчёт');
    return data;
  },

  async deleteReport(id) {
    const res = await fetch(`/api/reports/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (!res.ok) throw new Error('Не удалось удалить отчёт');
    return res.json().catch(() => ({}));
  },
};

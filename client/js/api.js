const Api = {
  async listReports() {
    const res = await fetch('/api/reports');
    if (!res.ok) throw new Error(I18N.t('fetch_list_failed'));
    return res.json();
  },

  async createReport(formData) {
    const res = await fetch('/api/reports', { method: 'POST', body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || I18N.t('create_failed'));
    return data;
  },

  async deleteReport(id) {
    const res = await fetch(`/api/reports/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (!res.ok) throw new Error(I18N.t('delete_failed'));
    return res.json().catch(() => ({}));
  },
};

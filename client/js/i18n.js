const I18N = {
  lang: (typeof CONFIG !== 'undefined' && CONFIG.LANGUAGE) || 'ru',

  dict: {
    reports_title: { ru: 'Отчёты', en: 'Reports' },
    create_report_btn: { ru: '+ Создать отчёт', en: '+ Create report' },
    loading: { ru: 'Загрузка...', en: 'Loading...' },

    new_report_title: { ru: 'Новый отчёт', en: 'New report' },
    report_name_label: { ru: 'Название отчёта', en: 'Report name' },
    excel_file_label: { ru: 'Excel-файл (.xlsx)', en: 'Excel file (.xlsx)' },
    include_connected_label: {
      ru: 'Показывать таблицы, использующиеся как источник сводной',
      en: 'Show tables used as a pivot table\u2019s data source',
    },
    create_btn: { ru: 'Создать', en: 'Create' },
    back_to_list: { ru: '\u2190 К списку отчётов', en: '\u2190 Back to reports' },

    no_reports_yet: { ru: 'Пока нет ни одного отчёта.', en: 'No reports yet.' },
    delete_btn: { ru: 'Удалить', en: 'Delete' },
    processing: { ru: 'обрабатывается...', en: 'processing...' },
    error_prefix: { ru: 'ошибка:', en: 'error:' },
    confirm_delete: { ru: 'Удалить отчёт «{name}»?', en: 'Delete report "{name}"?' },
    seconds_short: { ru: 'сек', en: 'sec' },
    minutes_short: { ru: 'мин', en: 'min' },

    fetch_list_failed: { ru: 'Не удалось получить список отчётов', en: 'Failed to load the report list' },
    create_failed: { ru: 'Не удалось создать отчёт', en: 'Failed to create the report' },
    delete_failed: { ru: 'Не удалось удалить отчёт', en: 'Failed to delete the report' },
  },

  t(key, vars) {
    const entry = this.dict[key];
    let text = (entry && (entry[this.lang] || entry.ru)) || key;
    if (vars) {
      Object.keys(vars).forEach((k) => {
        text = text.replace(`{${k}}`, vars[k]);
      });
    }
    return text;
  },

  apply(root) {
    document.documentElement.lang = this.lang;
    (root || document).querySelectorAll('[data-i18n]').forEach((el) => {
      el.textContent = this.t(el.getAttribute('data-i18n'));
    });
  },
};

document.addEventListener('DOMContentLoaded', () => I18N.apply());

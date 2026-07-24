const CONFIG = {
  // Site UI language: 'ru' or 'en'. Independent of LANGUAGE in
  // server/.env — that one controls server messages and generated
  // reports, this one controls only the site itself (list/form).
  LANGUAGE: 'en',

  // How often to poll /api/reports while at least one report is "processing"
  POLL_INTERVAL_MS: 3000,
};

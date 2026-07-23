# excel-html-converter

Turn Excel pivot tables, plain tables, and charts into browsable HTML reports — no Excel and no plugins needed to view them, just a browser. A small self-hosted Python web app does the work: upload an `.xlsx`, get a shareable report page.

## Quick start

```
npm install excel-html-converter
python node_modules/excel-html-converter/server/server.py
```

Open http://localhost:8000 in your browser.

## What it does

- Reads classic (worksheet-based) pivot tables, plain worksheet tables, and charts (line / bar / pie) directly from `.xlsx` files.
- Respects what Excel already shows: hidden filter values, autofilter-hidden rows, and date groupings (Year / Quarter / Month) are read from the file, not re-derived.
- Renders each as its own standalone page — a PivotGrid-based grid for pivot tables (with expand/collapse), a sortable/filterable table for plain data, and an SVG chart for line/bar/pie — all read-only snapshots of what was already built in Excel.
- A tiny web UI lists all generated reports: create, open, delete. Report generation runs in a separate OS process, so a large file never blocks the server or takes the whole app down if it runs out of memory.

## Configuration

Server settings live in `server/.env` (copy `server/.env.example` to get started):

- `PORT`, `HOST` — where the server listens.
- `TABLE_DEFAULT_ROW_LIMIT`, `PIVOT_DEFAULT_ROW_LIMIT` — row caps applied when the upload form doesn't specify one, to keep very large files from taking excessive time or memory.

## Project layout

```
server/
  server.py        - REST API + static file serving
  converters/       - the actual xlsx -> html parsers/generators
  reports/          - generated reports (created automatically)
client/
  index.html, new.html, js/, css/  - the small report-list web UI
  table-tools/, chart-tools/       - vanilla-JS rendering libraries used by generated reports
  node_modules/pivotgrid-js/       - pivot grid rendering engine (npm dependency)
```

## License

Free for personal, educational, or non-commercial use — see [LICENSE](LICENSE). For commercial use, see [LICENSE.commercial](LICENSE.commercial) or contact korolevalexa@gmail.com.

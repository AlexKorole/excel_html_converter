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
- `REPORTS_DIR` — where uploaded files and generated reports are stored. Defaults to a `reports/` folder next to wherever you run the server from — deliberately *not* inside the package itself, since `node_modules` gets wiped and recreated by `npm install` at any time, and that would take your reports with it.

Row limits (`TABLE_DEFAULT_ROW_LIMIT`, `PIVOT_DEFAULT_ROW_LIMIT`) are covered in their own section below.

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

## Excel file requirements

- **One kind of content per sheet.** A sheet is either a plain table, or hosts pivot table(s), or hosts chart(s) — not a mix. Multiple pivot tables or charts *can* share one sheet (each becomes its own page in the report); a plain table cannot share a sheet with anything else.
- **Only classic (worksheet-based) pivot tables** are supported. Pivot tables built on the Data Model / Power Pivot are skipped (with a message in the server log) — rebuild them as a regular pivot on a cell range if you need them in the report.
- **Supported chart types:** line, bar (column), pie. Other chart types are skipped.
- Excel's own hard ceiling is 1,048,576 rows per sheet — not something this tool adds, just a fact about `.xlsx` itself.

## Limits & configuration

Two *different* kinds of limit, both configurable in `server/.env` (copy from `server/.env.example`), server restart required after changing:

- **`TABLE_DEFAULT_ROW_LIMIT`** (default 5000) — a hard architectural ceiling, not just a nicety. Sorting/filtering a plain table happens directly on the DOM in the browser, with no virtualization; raising this a lot will make the browser itself noticeably slow on sort/filter, regardless of server power.
- **`PIVOT_DEFAULT_ROW_LIMIT`** (default 500000) — a soft safety default, not a hard limit. It exists because decoding raw pivot cache rows takes real server time and memory. The pivot grid itself renders fine at any size — feel free to raise this if your server has the time/memory to spare.

If a file has more rows than the limit in effect, the generated report footer says so explicitly ("shown: first N, source may have more"); if everything fit, no such note appears.

## License

Free for personal, educational, or non-commercial use — see [LICENSE](LICENSE). For commercial use, see [LICENSE.commercial](LICENSE.commercial) or contact korolevalexa@gmail.com.

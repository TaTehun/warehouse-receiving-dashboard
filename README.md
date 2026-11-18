# Warehouse Receiving Dashboard

A Python pipeline that collects five warehouses' daily receiving logs from email, merges them into an Excel template,
lets Excel recalculate the dashboard, and mails it as an HTML body with the workbook attached.

> **Visual overview:** [project-overview.html](project-overview.html)

**Status: testing.** The full run works end to end; the result mail goes to the developer until the recipient list and the daily schedule are signed off.

---

## What it does

1. **Picks the report date**: the previous weekday in Los Angeles time (Monday reports Friday).
2. **Scans the mailbox newest first** with POP3 `TOP` (headers only) and stops at the first mail received before the report day.
3. **Saves one receiving log per site** when the sender, subject, and report date match; a site already saved is skipped.
4. **Reads and normalizes the logs**: expands one provider's merged metadata cells, maps a site alias and a `SCHEDULED` status
   to the template's values, and drops placeholder rows.
5. **Writes the template**: rows into `RAW!A13`, Updated / Not Received per site into `B5:B9`, the report date into `B2`.
6. **Recalculates in Excel** (xlwings), exports the `DASHBOARD` sheet as HTML, and scales it to 80% for the mail body.
7. **Mails** the HTML dashboard with the recalculated workbook attached.

---

## Project Structure (private repo)

```
warehouse-receiving-dashboard/
├── main.py                 # report date, mail scan, template write, Excel export, send
├── processors/
│   ├── report_reader.py    # receiving-log reader, merged cells, normalization
│   ├── report_builder.py   # Updated / Not Received per site
│   └── dashboard_html.py   # 80% resize of the exported HTML
├── email_pipeline/         # sources, attachment save, SMTP
├── utils/                  # paths, logging
└── tests/                  # 21 isolated checks
```

This public repo holds `main_example.py` and `project-overview.html`.

```
python main_example.py   # report-date and resize demo
```

---

## Changelog

- **v0.9**: Report date is the previous LA weekday; the recalculated workbook is attached; the mail body is scaled to 80%.
- **v0.7**: Merged-cell reader and normalization; the dashboard matched the completed sample with zero cell differences.
- **v0.5**: Header-first POP3 scan with an early stop at the report day.

"""Warehouse receiving dashboard: five receiving logs from email -> Excel template -> HTML mail.

Portfolio example. Mail server, senders, sites, and paths are placeholders.

Flow
  1. report_date(): the previous weekday in LA time
  2. collect_logs(): newest-first POP3 scan with TOP, stop at the first mail older than the report day
  3. read_report(): every sheet of a log, merged metadata expanded, values normalized
  4. The template gets RAW!A13, B5:B9 and B2; Excel recalculates and exports the DASHBOARD sheet
  5. resize_dashboard(): scale the exported HTML for the mail body
"""
import poplib
import re
from datetime import date, datetime, timedelta
from email import message_from_bytes, policy
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

LA = ZoneInfo("America/Los_Angeles")
DOWNLOAD_DIR = Path("download")

EMAIL_TARGETS = [
    {"name": "WH1", "sender": "provider-a.example.com", "subject": "[DAILY EOD REPORT] WH1", "attachment": "WH1 INB Rec Log"},
    {"name": "WH2", "sender": "provider-a.example.com", "subject": "[DAILY EOD REPORT] WH2", "attachment": "WH2 INB Rec Log"},
    {"name": "WH3", "sender": "provider-a.example.com", "subject": "[DAILY EOD REPORT] WH3", "attachment": "WH3 INB Rec Log"},
    {"name": "WH4", "sender": "provider-a.example.com", "subject": "WH4 Inbound Rec Log", "attachment": "Inbound Log"},
    {"name": "WH5", "sender": "provider-b.example.com", "subject": "WH5 inbound rec log", "attachment": "WH5 INB Rec Log"},
]

# Metadata columns a provider merges down a shipment. Quantity and serial columns are left out on purpose:
# a merged quantity belongs to the shipment once, not to every row under it.
META_COLS = (1, 2, 4, 5, 6, 8, 9, 10, 11, 12)
TRANSPORT_MODES = ("AIR", "TRUCK", "OTR", "RAIL", "VESSEL", "SEA", "OCEAN")


# ============================================================
# Dates and mail
# ============================================================

def report_date(now: datetime | None = None) -> date:
    """The previous weekday in LA time. No holiday calendar."""
    day = (now or datetime.now(LA)).date() - timedelta(days=1)
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def match_target(msg):
    sender = parseaddr(str(msg.get("From", "")))[1].lower()
    subject = str(msg.get("Subject", "")).lower()
    for target in EMAIL_TARGETS:
        if target["sender"] in sender and target["subject"].lower() in subject:
            return target
    return None


def save_attachment(msg, target) -> Path | None:
    for part in msg.walk():
        name = part.get_filename()
        if name and name.lower().endswith((".xlsx", ".xls")) and target["attachment"].lower() in name.lower():
            DOWNLOAD_DIR.mkdir(exist_ok=True)
            path = DOWNLOAD_DIR / Path(name).name
            path.write_bytes(part.get_payload(decode=True))
            return path
    return None


def collect_logs(mailbox: poplib.POP3_SSL, day: date) -> list[Path]:
    """Newest first, headers only. The age check runs before any filter, so unrelated mail
    cannot keep the loop going past the report day."""
    scan_start = datetime(day.year, day.month, day.day, tzinfo=LA)
    count, _ = mailbox.stat()
    saved, files = set(), []
    for index in range(count, 0, -1):
        _, header, _ = mailbox.top(index, 0)
        msg = message_from_bytes(b"\r\n".join(header), policy=policy.default)
        if msg["Received"] is None:
            continue
        received = parsedate_to_datetime(str(msg["Received"]).rsplit(";", 1)[-1])
        if received < scan_start:
            break
        target = match_target(msg)
        if target is None or target["name"] in saved:
            continue
        subject = str(msg.get("Subject", ""))
        if day.strftime("%m%d%Y") not in subject and day.strftime("%m/%d/%Y") not in subject:
            continue
        _, lines, _ = mailbox.retr(index)
        path = save_attachment(message_from_bytes(b"\r\n".join(lines), policy=policy.default), target)
        if path:
            saved.add(target["name"])
            files.append(path)
        if len(saved) == len(EMAIL_TARGETS):
            break
    return files


# ============================================================
# Receiving logs
# ============================================================

def normalize_report_row(row) -> tuple:
    values = list(row)
    if str(values[0]).strip().upper() == "WH5-OLD":
        values[0] = "WH5"                      # site alias -> template value
    if str(values[11]).strip().upper() == "SCHEDULED":
        values[11] = "Planned"
    # One legacy log keeps the transport mode in Remark; the HBL column identifies the shipment
    if values[8] in (None, "") and str(values[7]).strip().upper() in TRANSPORT_MODES:
        values[8] = str(values[7]).strip().upper()
        values[7] = values[5]
    return tuple(values)


def read_report(path: Path) -> list[tuple]:
    wb = load_workbook(path, data_only=True)
    rows = []
    try:
        for sheet in wb:
            merged = {}
            for span in sheet.merged_cells.ranges:
                if span.min_col == span.max_col and span.min_col in META_COLS:
                    value = sheet.cell(span.min_row, span.min_col).value
                    for r in range(span.min_row, span.max_row + 1):
                        merged[r, span.min_col] = value
            started = False
            for r, row in enumerate(sheet.iter_rows(max_col=12, values_only=True), start=1):
                if not started:
                    started = row[0] == "WH"           # data starts below the header row
                    continue
                if not any(v not in (None, "") for v in row[1:]):
                    continue                           # WH-only placeholder, not a shipment
                row = [merged.get((r, c), v) for c, v in enumerate(row, start=1)]
                rows.append(normalize_report_row(row))
    finally:
        wb.close()
    return rows


# ============================================================
# Mail body
# ============================================================

DIM_CSS = (r"(?<![\w-])((?:width|height|font-size|line-height|padding(?:-(?:top|right|bottom|left))?)\s*:\s*)"
           r"(\d*\.?\d+)(px|pt)\b")


def resize_dashboard(html_body: str, scale: float = 0.8) -> str:
    """Excel's HTML export is too wide for a mail body. Scale every dimension by one factor
    so the table keeps its proportions; the attached workbook is not touched."""
    def resize_css(css):
        return re.sub(DIM_CSS, lambda m: f"{m[1]}{float(m[2]) * scale:g}{m[3]}", css, flags=re.I)

    def resize_tag(match):
        tag = re.sub(r"(\s(?:width|height)\s*=\s*)([\"']?)(\d+(?:\.\d+)?)\2(?=[\s/>])",
                     lambda s: f"{s[1]}{s[2]}{round(float(s[3]) * scale)}{s[2]}", match[0], flags=re.I)
        return re.sub(r"(\bstyle\s*=\s*)([\"'])(.*?)\2",
                      lambda s: f"{s[1]}{s[2]}{resize_css(s[3])}{s[2]}", tag, flags=re.I | re.S)

    html_body = re.sub(r"(<style\b[^>]*>)(.*?)(</style\s*>)",
                       lambda s: f"{s[1]}{resize_css(s[2])}{s[3]}", html_body, flags=re.I | re.S)
    return re.sub(r"<[a-z][^>]*>", resize_tag, html_body, flags=re.I)


if __name__ == "__main__":
    print(report_date(datetime(2026, 11, 16, 9, tzinfo=LA)))   # Monday -> 2026-11-13 (Friday)
    print(resize_dashboard('<table width=1120 style="font-size:10pt;padding:5px">'))
    # <table width=896 style="font-size:8pt;padding:4px">

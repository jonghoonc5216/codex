# -*- coding: utf-8 -*-
import argparse
import hashlib
import json
import os
import re
import sqlite3
import sys
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
import tkinter as tk


APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = Path(os.environ.get("USERPROFILE", "")) / "Documents" / "★최종훈노트"
DATA_DIR = Path(os.environ.get("STICKY_ALARM_DATA_DIR", str(DEFAULT_DATA_DIR)))
STATE_PATH = DATA_DIR / "스티커메모_알람상태.json"
ALARM_TEXT_PATH = DATA_DIR / "스티커메모_알람문구.json"
WORD_ARCHIVE_LOG_PATH = DATA_DIR / "스티커메모_워드저장로그.txt"
PID_PATH = APP_DIR / "sticky_alarm.pid"
LOG_PATH = APP_DIR / "sticky_alarm.log"
LOCK_PATH = APP_DIR / "sticky_alarm.lock"
DEFAULT_DB_PATH = (
    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Packages"
    / "Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe"
    / "LocalState"
    / "plum.sqlite"
)

CHECK_INTERVAL_SECONDS = 60
DEFAULT_GRACE_MINUTES = 12 * 60
WORD_ARCHIVE_STABLE_SECONDS = 60
KST = timezone(timedelta(hours=9))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", W_NS)

ALARM_LINE_RE = re.compile(
    r"^\s*(?:[@#]\s*)?(?:알람|alarm|reminder)\s*[:=：]?\s*(?P<expr>.+?)\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Alarm:
    note_id: str
    due_at: datetime
    message: str
    source_line: str

    @property
    def key(self) -> str:
        digest = hashlib.sha1(self.message.encode("utf-8", errors="ignore")).hexdigest()
        return f"{self.note_id}|{self.due_at.isoformat()}|{digest}"


@dataclass(frozen=True)
class StickyNote:
    note_id: str
    text: str
    created_at: int
    updated_at: int

    @property
    def clean_text(self) -> str:
        return clean_sticky_text(self.text)

    @property
    def signature(self) -> str:
        digest = hashlib.sha1(self.clean_text.encode("utf-8", errors="ignore")).hexdigest()
        return f"{self.updated_at}|{digest}"

    @property
    def updated_datetime(self) -> datetime:
        return ticks_to_local_datetime(self.updated_at) or datetime.now()


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def load_state() -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not STATE_PATH.exists():
        return {"fired": {}}
    try:
        with STATE_PATH.open("r", encoding="utf-8") as f:
            state = json.load(f)
        if not isinstance(state.get("fired"), dict):
            state["fired"] = {}
        return state
    except Exception as exc:
        log(f"state load failed: {exc}")
        return {"fired": {}}


def save_state(state: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = STATE_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    temp_path.replace(STATE_PATH)


def save_alarm_texts(alarms, now: datetime) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "updated_at": now.isoformat(timespec="seconds"),
        "source": str(DEFAULT_DB_PATH),
        "alarms": [
            {
                "due_at": alarm.due_at.isoformat(timespec="minutes"),
                "message": alarm.message,
                "alarm_line": alarm.source_line,
                "note_id": alarm.note_id,
            }
            for alarm in sorted(alarms, key=lambda item: item.due_at)
        ],
    }
    temp_path = ALARM_TEXT_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    temp_path.replace(ALARM_TEXT_PATH)


def ticks_to_local_datetime(value):
    try:
        ticks = int(value or 0)
        if ticks <= 0:
            return None
        utc = datetime(1, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ticks / 10)
        return utc.astimezone(KST).replace(tzinfo=None)
    except Exception:
        return None


def word_log(message: str) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with WORD_ARCHIVE_LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def w_tag(name):
    return f"{{{W_NS}}}{name}"


def set_w_attr(element, name, value):
    element.set(w_tag(name), str(value))


def make_docx_run(text, bold=False, size=22):
    run = ET.Element(w_tag("r"))
    rpr = ET.SubElement(run, w_tag("rPr"))
    fonts = ET.SubElement(rpr, w_tag("rFonts"))
    set_w_attr(fonts, "ascii", "Malgun Gothic")
    set_w_attr(fonts, "hAnsi", "Malgun Gothic")
    set_w_attr(fonts, "eastAsia", "Malgun Gothic")
    if bold:
        ET.SubElement(rpr, w_tag("b"))
    sz = ET.SubElement(rpr, w_tag("sz"))
    set_w_attr(sz, "val", size)
    sz_cs = ET.SubElement(rpr, w_tag("szCs"))
    set_w_attr(sz_cs, "val", size)
    text_el = ET.SubElement(run, w_tag("t"))
    if text != text.strip():
        text_el.set(f"{{{XML_NS}}}space", "preserve")
    text_el.text = text
    return run


def make_docx_paragraph(text="", bold=False, size=22, align=None):
    paragraph = ET.Element(w_tag("p"))
    if align:
        ppr = ET.SubElement(paragraph, w_tag("pPr"))
        jc = ET.SubElement(ppr, w_tag("jc"))
        set_w_attr(jc, "val", align)
    paragraph.append(make_docx_run(text, bold=bold, size=size))
    return paragraph


def make_docx_section():
    sect = ET.Element(w_tag("sectPr"))
    page_size = ET.SubElement(sect, w_tag("pgSz"))
    set_w_attr(page_size, "w", 11906)
    set_w_attr(page_size, "h", 16838)
    page_margin = ET.SubElement(sect, w_tag("pgMar"))
    for key, value in {
        "top": 1440,
        "right": 1440,
        "bottom": 1440,
        "left": 1440,
        "header": 708,
        "footer": 708,
        "gutter": 0,
    }.items():
        set_w_attr(page_margin, key, value)
    return sect


def docx_xml_bytes(element):
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def month_docx_path(entry_time: datetime) -> Path:
    return DATA_DIR / f"{entry_time.year}년{entry_time.month}월.docx"


def docx_content_types():
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def docx_package_rels():
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def docx_word_rels():
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""


def docx_core_xml(now: datetime, title: str):
    stamp = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    safe_title = (
        title.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{safe_title}</dc:title>
  <dc:creator>sticky_alarm</dc:creator>
  <cp:lastModifiedBy>sticky_alarm</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}</dcterms:modified>
</cp:coreProperties>
""".encode("utf-8")


def docx_app_xml():
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>sticky_alarm</Application>
</Properties>
"""


def create_month_docx(path: Path, entry_time: datetime):
    title = f"{entry_time.year}년 {entry_time.month}월 스티커 메모"
    document = ET.Element(w_tag("document"))
    body = ET.SubElement(document, w_tag("body"))
    body.append(make_docx_paragraph(title, bold=True, size=32, align="center"))
    body.append(make_docx_paragraph(""))
    body.append(make_docx_section())
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", docx_content_types())
        z.writestr("_rels/.rels", docx_package_rels())
        z.writestr("word/_rels/document.xml.rels", docx_word_rels())
        z.writestr("word/document.xml", docx_xml_bytes(document))
        z.writestr("docProps/core.xml", docx_core_xml(entry_time, title))
        z.writestr("docProps/app.xml", docx_app_xml())


def append_note_to_month_docx(note: StickyNote, archive_time: datetime):
    text = note.clean_text
    if not text:
        return None

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    entry_time = note.updated_datetime
    path = month_docx_path(entry_time)
    if not path.exists():
        create_month_docx(path, entry_time)

    temp_path = path.with_suffix(".tmp")
    with zipfile.ZipFile(path, "r") as zin:
        document = ET.fromstring(zin.read("word/document.xml"))
        body = document.find(w_tag("body"))
        if body is None:
            raise RuntimeError("Word document body not found")

        section = None
        children = list(body)
        if children and children[-1].tag == w_tag("sectPr"):
            section = children[-1]
            body.remove(section)

        body.append(make_docx_paragraph(""))
        body.append(make_docx_paragraph(entry_time.strftime("%Y-%m-%d %H:%M"), bold=True, size=24))
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            body.append(make_docx_paragraph(line, size=22))
        body.append(make_docx_paragraph(""))
        body.append(section if section is not None else make_docx_section())

        updated_document = docx_xml_bytes(document)
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = updated_document if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    temp_path.replace(path)
    word_log(f"saved note {note.note_id} to {path} at {archive_time.isoformat(timespec='seconds')}")
    return path


def prune_state(state: dict, now: datetime) -> None:
    cutoff = now - timedelta(days=45)
    fired = state.setdefault("fired", {})
    for key, value in list(fired.items()):
        try:
            if datetime.fromisoformat(value) < cutoff:
                fired.pop(key, None)
        except Exception:
            fired.pop(key, None)


def clean_sticky_text(text: str) -> str:
    if not text:
        return ""
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n")
    cleaned = re.sub(r"\\id=[0-9a-fA-F-]+", "", cleaned)
    cleaned = cleaned.replace("\uFFFC", "")
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def parse_time(expr: str):
    time_matches = list(
        re.finditer(
            r"(?P<ampm>오전|오후|am|pm)?\s*(?P<hour>\d{1,2})\s*(?::|시)\s*(?P<minute>\d{1,2})?\s*(?:분)?",
            expr,
            re.IGNORECASE,
        )
    )
    if not time_matches:
        return None

    match = time_matches[-1]
    ampm = (match.group("ampm") or "").lower()
    hour = int(match.group("hour"))
    minute_text = match.group("minute")
    minute = int(minute_text) if minute_text is not None else 0

    if hour > 23 or minute > 59:
        return None

    if ampm in ("오후", "pm") and hour < 12:
        hour += 12
    if ampm in ("오전", "am") and hour == 12:
        hour = 0

    return hour, minute


def parse_date(expr: str, now: datetime):
    if "모레" in expr:
        return (now + timedelta(days=2)).date(), True
    if "내일" in expr:
        return (now + timedelta(days=1)).date(), True
    if "오늘" in expr:
        return now.date(), True

    full = re.search(
        r"(?P<year>20\d{2})\s*(?:년|[./-])\s*(?P<month>\d{1,2})\s*(?:월|[./-])\s*(?P<day>\d{1,2})\s*(?:일)?",
        expr,
    )
    if full:
        try:
            return (
                datetime(
                    int(full.group("year")),
                    int(full.group("month")),
                    int(full.group("day")),
                ).date(),
                True,
            )
        except ValueError:
            return None, True

    short = re.search(
        r"(?<!\d)(?P<month>\d{1,2})\s*(?:월|[./-])\s*(?P<day>\d{1,2})\s*(?:일)?(?!\d)",
        expr,
    )
    if short:
        month = int(short.group("month"))
        day = int(short.group("day"))
        try:
            alarm_date = datetime(now.year, month, day).date()
            return alarm_date, True
        except ValueError:
            return None, True

    return now.date(), False


def parse_alarm_datetime(expr: str, now: datetime):
    expr = expr.strip()
    time_part = parse_time(expr)
    if time_part is None:
        return None

    alarm_date, has_explicit_date = parse_date(expr, now)
    if alarm_date is None:
        return None

    hour, minute = time_part
    due_at = datetime.combine(alarm_date, datetime.min.time()).replace(hour=hour, minute=minute)

    if not has_explicit_date and due_at <= now:
        due_at += timedelta(days=1)
    elif has_explicit_date and due_at < now - timedelta(days=1) and re.search(r"(?<!\d)\d{1,2}\s*(?:월|[./-])\s*\d{1,2}", expr) and not re.search(r"20\d{2}", expr):
        try:
            due_at = due_at.replace(year=due_at.year + 1)
        except ValueError:
            return None

    return due_at


def extract_alarms(note_id: str, raw_text: str, now: datetime):
    visible = clean_sticky_text(raw_text)
    if not visible:
        return []

    alarm_lines = []
    message_lines = []
    for line in visible.splitlines():
        match = ALARM_LINE_RE.match(line)
        if match:
            alarm_lines.append((line.strip(), match.group("expr").strip()))
        else:
            message_lines.append(line)

    if not alarm_lines:
        return []

    message = "\n".join(message_lines).strip() or visible
    alarms = []
    for source_line, expr in alarm_lines:
        due_at = parse_alarm_datetime(expr, now)
        if due_at is None:
            log(f"ignored alarm line: {source_line}")
            continue
        alarms.append(Alarm(note_id=note_id, due_at=due_at, message=message, source_line=source_line))
    return alarms


def read_sticky_notes(db_path: Path):
    uri = f"file:{db_path}?mode=ro"
    con = sqlite3.connect(uri, uri=True, timeout=5)
    try:
        cur = con.cursor()
        rows = cur.execute(
            """
            select Id, Text, CreatedAt, UpdatedAt
            from Note
            where DeletedAt is null or DeletedAt = 0
            """
        ).fetchall()
        return [
            StickyNote(
                note_id=str(note_id),
                text=text or "",
                created_at=int(created_at or 0),
                updated_at=int(updated_at or 0),
            )
            for note_id, text, created_at, updated_at in rows
        ]
    finally:
        con.close()


def collect_alarms(db_path: Path, now: datetime):
    rows = read_sticky_notes(db_path)
    alarms = []
    for note in rows:
        alarms.extend(extract_alarms(note.note_id, note.text, now))
    return alarms


def collect_alarms_from_notes(notes, now: datetime):
    alarms = []
    for note in notes:
        alarms.extend(extract_alarms(note.note_id, note.text, now))
    return alarms


def parse_iso_datetime(value):
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def initialize_word_archive_state(state: dict, notes, now: datetime) -> bool:
    if "word_archive_seen" in state:
        return False

    seen = {}
    for note in notes:
        if note.clean_text:
            seen[note.note_id] = note.signature

    state["word_archive_seen"] = seen
    state["word_archive_pending"] = {}
    state["word_archive_initialized_at"] = now.isoformat(timespec="seconds")
    word_log(f"baseline initialized with {len(seen)} existing sticky notes")
    return True


def handle_word_archive(state: dict, notes, now: datetime, stable_seconds: int):
    if initialize_word_archive_state(state, notes, now):
        return

    seen = state.setdefault("word_archive_seen", {})
    pending = state.setdefault("word_archive_pending", {})
    current_ids = set()
    stable_for = timedelta(seconds=max(0, int(stable_seconds)))

    for note in sorted(notes, key=lambda item: (item.updated_at or item.created_at or 0, item.note_id)):
        text = note.clean_text
        if not text:
            continue

        current_ids.add(note.note_id)
        signature = note.signature
        if seen.get(note.note_id) == signature:
            pending.pop(note.note_id, None)
            continue

        pending_entry = pending.get(note.note_id)
        if not pending_entry or pending_entry.get("signature") != signature:
            pending[note.note_id] = {
                "signature": signature,
                "first_seen": now.isoformat(timespec="seconds"),
                "updated_at": note.updated_at,
            }
            continue

        first_seen = parse_iso_datetime(pending_entry.get("first_seen"))
        if first_seen is None or now - first_seen < stable_for:
            continue

        try:
            append_note_to_month_docx(note, now)
            seen[note.note_id] = signature
            pending.pop(note.note_id, None)
        except Exception as exc:
            word_log(f"save failed for note {note.note_id}: {exc}")
            log(f"word archive failed for note {note.note_id}: {exc}")

    for note_id in list(pending.keys()):
        if note_id not in current_ids:
            pending.pop(note_id, None)


class FullscreenAlarmApp:
    def __init__(self, args):
        self.args = args
        self.db_path = Path(args.db).expanduser()
        self.state = load_state()
        self.queue = []
        self.current_window = None
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("Sticky Memo Alarm")
        self.root.protocol("WM_DELETE_WINDOW", self.root.withdraw)

    def run(self):
        PID_PATH.write_text(str(os.getpid()), encoding="utf-8")
        log("monitor started")
        self.scan()
        self.root.mainloop()

    def scan(self):
        now = datetime.now()
        try:
            prune_state(self.state, now)
            if not self.db_path.exists():
                log(f"database not found: {self.db_path}")
            else:
                notes = read_sticky_notes(self.db_path)
                alarms = collect_alarms_from_notes(notes, now)
                save_alarm_texts(alarms, now)
                handle_word_archive(self.state, notes, now, WORD_ARCHIVE_STABLE_SECONDS)
                self.handle_alarms(alarms, now)
                save_state(self.state)
        except Exception as exc:
            log(f"scan failed: {exc}")

        interval_ms = max(5, int(self.args.interval)) * 1000
        self.root.after(interval_ms, self.scan)

    def handle_alarms(self, alarms, now: datetime):
        fired = self.state.setdefault("fired", {})
        grace = timedelta(minutes=max(1, int(self.args.grace_minutes)))
        for alarm in alarms:
            if alarm.key in fired:
                continue
            if alarm.due_at <= now:
                fired[alarm.key] = now.isoformat()
                if now - alarm.due_at <= grace:
                    self.queue.append(alarm)
                    log(f"queued alarm: {alarm.due_at.isoformat()} {alarm.source_line}")
                else:
                    log(f"expired alarm skipped: {alarm.due_at.isoformat()} {alarm.source_line}")

        if self.current_window is None and self.queue:
            self.show_next_alarm()

    def show_next_alarm(self):
        if not self.queue:
            return
        alarm = self.queue.pop(0)
        self.current_window = AlarmWindow(self.root, alarm, self.close_current_alarm, self.args.auto_close)

    def close_current_alarm(self):
        if self.current_window is not None:
            self.current_window.destroy()
            self.current_window = None
        if self.queue:
            self.root.after(150, self.show_next_alarm)


class AlarmWindow:
    def __init__(self, root, alarm: Alarm, on_close, auto_close_seconds=0):
        self.root = root
        self.alarm = alarm
        self.on_close = on_close
        self.window = tk.Toplevel(root)
        self.window.title("스티커 메모 알람")
        self.window.configure(bg="#240006")
        self.window.attributes("-fullscreen", True)
        self.window.attributes("-topmost", True)
        self.window.bind("<Escape>", lambda event: self.on_close())
        self.window.bind("<Return>", lambda event: self.on_close())
        self.window.bind("<space>", lambda event: self.on_close())

        width = self.window.winfo_screenwidth()
        height = self.window.winfo_screenheight()
        wrap = max(420, int(width * 0.78))

        container = tk.Frame(self.window, bg="#240006")
        container.place(relx=0.5, rely=0.5, anchor="center", width=int(width * 0.88), height=int(height * 0.78))

        title = tk.Label(
            container,
            text="스티커 메모 알람",
            bg="#240006",
            fg="#ffffff",
            font=("Malgun Gothic", 30, "bold"),
        )
        title.pack(pady=(10, 18))

        time_text = alarm.due_at.strftime("%Y-%m-%d %H:%M")
        time_label = tk.Label(
            container,
            text=time_text,
            bg="#240006",
            fg="#ffd9de",
            font=("Malgun Gothic", 20, "bold"),
        )
        time_label.pack(pady=(0, 26))

        message = tk.Label(
            container,
            text=alarm.message,
            bg="#240006",
            fg="#ffffff",
            justify="center",
            wraplength=wrap,
            font=("Malgun Gothic", 34, "bold"),
        )
        message.pack(expand=True, fill="both")

        button = tk.Button(
            container,
            text="확인",
            command=self.on_close,
            bg="#ffffff",
            fg="#240006",
            activebackground="#ffd9de",
            activeforeground="#240006",
            font=("Malgun Gothic", 18, "bold"),
            relief="flat",
            padx=42,
            pady=12,
        )
        button.pack(pady=(28, 10))

        footer = tk.Label(
            container,
            text="Enter / Space / Esc 로 닫기",
            bg="#240006",
            fg="#ffc7cf",
            font=("Malgun Gothic", 12),
        )
        footer.pack()

        self.window.lift()
        self.window.focus_force()

        if auto_close_seconds:
            self.window.after(int(auto_close_seconds * 1000), self.on_close)

    def destroy(self):
        if self.window.winfo_exists():
            self.window.destroy()


def list_alarms(db_path: Path) -> int:
    now = datetime.now()
    alarms = collect_alarms(db_path, now)
    save_alarm_texts(alarms, now)
    if not alarms:
        print("등록된 알람 문구가 없습니다.")
        return 0
    for alarm in sorted(alarms, key=lambda item: item.due_at):
        preview = " ".join(alarm.message.split())
        if len(preview) > 80:
            preview = preview[:77] + "..."
        print(f"{alarm.due_at:%Y-%m-%d %H:%M} | {preview}")
    return 0


def run_test_alert(auto_close_seconds: int) -> int:
    root = tk.Tk()
    root.withdraw()
    alarm = Alarm(
        note_id="test",
        due_at=datetime.now(),
        message="전체화면 알림 테스트입니다.\n검붉은 배경과 흰색 글씨가 잘 보이면 성공입니다.",
        source_line="test",
    )

    def close():
        root.quit()

    window = AlarmWindow(root, alarm, close, auto_close_seconds)
    root.mainloop()
    window.destroy()
    root.destroy()
    return 0


def parse_args():
    parser = argparse.ArgumentParser(description="Sticky Notes alarm monitor")
    parser.add_argument("--db", default=str(DEFAULT_DB_PATH), help="Sticky Notes plum.sqlite path")
    parser.add_argument("--interval", type=int, default=CHECK_INTERVAL_SECONDS, help="scan interval seconds")
    parser.add_argument("--grace-minutes", type=int, default=DEFAULT_GRACE_MINUTES, help="show alarms missed within this many minutes")
    parser.add_argument("--list", action="store_true", help="list alarms parsed from Sticky Notes")
    parser.add_argument("--test-alert", type=int, metavar="SECONDS", help="show a test full-screen alert and auto-close")
    parser.add_argument("--auto-close", type=int, default=0, help="auto-close real alarm windows after N seconds; 0 disables")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    db_path = Path(args.db).expanduser()

    if args.test_alert is not None:
        return run_test_alert(max(1, args.test_alert))

    if args.list:
        if not db_path.exists():
            print(f"스티커 메모 데이터베이스를 찾지 못했습니다: {db_path}")
            return 1
        return list_alarms(db_path)

    if not db_path.exists():
        print(f"스티커 메모 데이터베이스를 찾지 못했습니다: {db_path}")
        return 1

    lock_file = acquire_single_instance_lock()
    if lock_file is None:
        log("monitor already running")
        print("스티커 메모 알람 감시 프로그램이 이미 실행 중입니다.")
        return 0

    app = FullscreenAlarmApp(args)
    try:
        app.run()
    finally:
        try:
            if PID_PATH.exists() and PID_PATH.read_text(encoding="utf-8").strip() == str(os.getpid()):
                PID_PATH.unlink()
        except Exception:
            pass
        log("monitor stopped")
    return 0


def acquire_single_instance_lock():
    try:
        import msvcrt

        lock_file = LOCK_PATH.open("a+", encoding="utf-8")
        lock_file.seek(0)
        try:
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            return lock_file
        except OSError:
            lock_file.close()
            return None
    except Exception as exc:
        log(f"single-instance lock unavailable: {exc}")
        return LOCK_PATH.open("a+", encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())

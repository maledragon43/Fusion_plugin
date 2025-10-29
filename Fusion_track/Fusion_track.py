import adsk.core, adsk.fusion, adsk.cam
import csv, tempfile, traceback, sys, os
from pathlib import Path
from datetime import datetime, timedelta
from io import StringIO

# ======================= Paths (save in add-in folder) =======================
def _addon_dir() -> str:
    return os.path.dirname(__file__)

def _sessions_path() -> str:
    return os.path.join(_addon_dir(), "sessions.csv")

SESSIONS_CSV    = Path(_sessions_path())
SESSIONS_HEADER = ["project_name", "document_name", "status", "start_time", "end_time", "duration_seconds", "duration_hms"]

# ======================= Tunables =======================
IDLE_THRESHOLD_SEC = 120      # go idle after 120s of no input
POLL_INTERVAL_SEC  = 2        # heartbeat cadence from palette JS
TIME_FMT = "%m/%d/%Y %H:%M"   # 10/29/2025 19:07
MIN_SESSION_SECONDS = 1        # drop ultra-short segments

# ======================= Globals =======================
_handlers = []
_palette = None
_html_ready = False
_buffer = []

# Current doc + segment state
_active_doc_key = None                  # stable key for current doc
_active_names   = ("Unsaved", "Unknown")# (project, document)
_curr_status    = "active"              # "active" | "idle"
_seg_start_dt   = None                  # start time of current segment

# ======================= UI helpers =======================
def _append_ui(msg: str):
    global _html_ready, _buffer
    s = (msg or "").strip()
    if not s: return
    try:
        if _palette and _html_ready:
            _palette.sendInfoToHTML("appendLog", s)
        else:
            _buffer.append(s)
    except:
        _buffer.append(s)

def _flush_ui():
    global _buffer
    if not _palette: return
    for m in _buffer:
        try: _palette.sendInfoToHTML("appendLog", m)
        except: break
    _buffer = []

class UIPaletteWriter(StringIO):
    def write(self, s):
        if s and s.strip():
            _append_ui(s.strip())
        return len(s or "")

# ======================= CSV / time helpers =======================
def _ensure_csv(path: Path, header):
    try:
        Path(_addon_dir()).mkdir(parents=True, exist_ok=True)
        if not path.exists():
            with open(path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(header)
            _append_ui(f"[TimeSessions] Created CSV: {path}")
    except Exception:
        _append_ui("[TimeSessions] Failed to create CSV:\n" + traceback.format_exc())

def _hms(seconds: int) -> str:
    s = max(0, int(seconds))
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    return f"{h}:{m:02d}:{sec:02d}"

def _doc_key_and_names(doc: adsk.core.Document):
    proj, name, key = "Unsaved", "Unknown", "None"
    try:
        if doc:
            name = getattr(doc, "name", "Unknown") or "Unknown"
            df = getattr(doc, "dataFile", None)
            if df and getattr(df, "project", None) and getattr(df.project, "name", None):
                proj = df.project.name
            urn = getattr(df, "id", None) or getattr(df, "versionId", None) or ""
            key = urn if urn else f"{proj}|{name}"
    except:
        pass
    return key, (proj, name)

def _write_segment_row(names, status, start_dt: datetime, end_dt: datetime):
    dur = int((end_dt - start_dt).total_seconds())
    if dur < MIN_SESSION_SECONDS:
        _append_ui(f"[TimeSessions] (dropped short {status}) {names[0]} / {names[1]} | {_hms(dur)}")
        return
    try:
        _ensure_csv(SESSIONS_CSV, SESSIONS_HEADER)
        row = [names[0], names[1], status,
               start_dt.strftime(TIME_FMT), end_dt.strftime(TIME_FMT),
               dur, _hms(dur)]
        with open(SESSIONS_CSV, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(row)
            f.flush()
        _append_ui(f"[TimeSessions] Saved {status} | {names[0]} / {names[1]} | {_hms(dur)}")
    except Exception:
        _append_ui("[TimeSessions] Error writing CSV:\n" + traceback.format_exc())

# ======================= OS idle time (Windows) =======================
def _get_idle_seconds_windows() -> int:
    try:
        import ctypes
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]
        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if not user32.GetLastInputInfo(ctypes.byref(lii)):
            return 0
        millis = kernel32.GetTickCount() - lii.dwTime
        return int(millis / 1000)
    except Exception:
        return 0

# ======================= Session control =======================
def _switch_to(doc: adsk.core.Document, reason: str):
    """End current segment, start a new ACTIVE segment for the new doc."""
    global _active_doc_key, _active_names, _curr_status, _seg_start_dt
    now = datetime.now()

    # end current (if any)
    if _active_doc_key and _seg_start_dt:
        _write_segment_row(_active_names, _curr_status, _seg_start_dt, now)

    # begin new active segment
    new_key, new_names = _doc_key_and_names(doc)
    _active_doc_key = new_key
    _active_names   = new_names
    _curr_status    = "active"
    _seg_start_dt   = now
    _append_ui(f"[TimeSessions] ▶ ACTIVE start ({reason}) | {new_names[0]} / {new_names[1]} @ {now.strftime('%H:%M:%S')}")

def _stop_current(reason: str):
    """End current segment (active/idle) and clear."""
    global _active_doc_key, _seg_start_dt
    if not _active_doc_key or not _seg_start_dt:
        return
    now = datetime.now()
    _write_segment_row(_active_names, _curr_status, _seg_start_dt, now)
    _append_ui(f"[TimeSessions] ■ Stop ({reason}) @ {now.strftime('%H:%M:%S')}")
    _active_doc_key = None
    _seg_start_dt   = None

def _to_idle(idle_start_dt: datetime):
    """ACTIVE → IDLE at precise idle start."""
    global _curr_status, _seg_start_dt
    if _curr_status == "idle":
        return
    if _active_doc_key and _seg_start_dt:
        _write_segment_row(_active_names, "active", _seg_start_dt, idle_start_dt)
    _curr_status  = "idle"
    _seg_start_dt = idle_start_dt
    _append_ui(f"[TimeSessions] ⏸ Idle start @ {idle_start_dt.strftime('%H:%M:%S')}")

def _from_idle_to_active(resume_dt: datetime):
    """IDLE → ACTIVE at resume time."""
    global _curr_status, _seg_start_dt
    if _curr_status == "active":
        return
    if _active_doc_key and _seg_start_dt:
        _write_segment_row(_active_names, "idle", _seg_start_dt, resume_dt)
    _curr_status  = "active"
    _seg_start_dt = resume_dt
    _append_ui(f"[TimeSessions] ▶ Resume ACTIVE @ {resume_dt.strftime('%H:%M:%S')}")

# ======================= Event handlers =======================
class HtmlIncomingHandler(adsk.core.HTMLEventHandler):
    def notify(self, args):
        try:
            global _html_ready
            a = adsk.core.HTMLEventArgs.cast(args)
            if not a: return

            if a.action == "ready":
                _html_ready = True
                _flush_ui()

            elif a.action == "tick":
                # Palette heartbeat (every POLL_INTERVAL_SEC). Compute OS idle here.
                idle_sec = _get_idle_seconds_windows()
                now = datetime.now()
                if not _active_doc_key or not _seg_start_dt:
                    return
                if idle_sec >= IDLE_THRESHOLD_SEC:
                    # exact idle start time
                    idle_start = now - timedelta(seconds=idle_sec)
                    # don't allow idle start before the current segment began
                    idle_start = max(idle_start, _seg_start_dt)
                    _to_idle(idle_start)
                else:
                    _from_idle_to_active(now)

        except Exception:
            _append_ui("[TimeSessions] HTML incoming error:\n" + traceback.format_exc())

class DocEventHandler(adsk.core.DocumentEventHandler):
    def __init__(self, name):
        super().__init__()
        self.name = name
    def notify(self, args):
        try:
            if self.name == "DocumentActivated":
                doc = getattr(args, "document", None)
                _switch_to(doc, "activated")
            elif self.name == "DocumentClosed":
                closed_doc = getattr(args, "document", None)
                key_closed, _ = _doc_key_and_names(closed_doc)
                if _active_doc_key and key_closed == _active_doc_key:
                    _stop_current("close")
        except Exception:
            _append_ui(f"[TimeSessions] Error in {self.name}:\n" + traceback.format_exc())

# ======================= Palette UI =======================
def _write_html():
    try:
        temp_dir = Path(tempfile.gettempdir()) / "FusionTimeSessions"
        temp_dir.mkdir(parents=True, exist_ok=True)
        html_path = temp_dir / "ui.html"
        html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Fusion Time Sessions</title>
<style>
body {{ font-family: Consolas, monospace; background:#0d0f14; color:#d7f5d4; margin:0; }}
h3 {{ margin:10px 12px; }}
.wrap {{ padding: 0 12px 12px; }}
#log {{ white-space: pre-line; font-size:12px; background:#0b0d12; border:1px solid #2a3340; padding:8px; height: calc(100vh - 80px); overflow:auto; }}
small {{ color:#94a3b8; }}
button {{ background:#1f2937; color:#d7f5d4; border:1px solid #334155; padding:6px 10px; cursor:pointer; margin-left: 12px; }}
button:hover {{ background:#263241; }}
</style>
<script>
function appendLog(msg) {{
  var el = document.getElementById('log');
  el.textContent += msg + "\\n";
  el.scrollTop = el.scrollHeight;
}}
function clearLog() {{ document.getElementById('log').textContent = ''; }}
// Heartbeat to Python every {POLL_INTERVAL_SEC}s
window.addEventListener('load', () => {{
  try {{ adsk.fusionSendData('ready',''); }} catch(e){{}}
  setInterval(() => {{ try {{ adsk.fusionSendData('tick',''); }} catch(e){{}} }}, {POLL_INTERVAL_SEC*1000});
}});
</script>
</head>
<body>
  <h3>Fusion Time Sessions <small>active/idle</small>
    <button onclick="clearLog()">Clear</button>
  </h3>
  <div class="wrap">
    <div id="log">--- UI Ready ---\\n</div>
  </div>
</body>
</html>"""
        html_path.write_text(html, encoding="utf-8")
        return html_path.as_uri()
    except Exception:
        print("[TimeSessions] Failed to write HTML:\n" + traceback.format_exc())
        return ""

def _create_palette_and_redirect_console():
    global _palette
    app = adsk.core.Application.get()
    ui = app.userInterface
    old = ui.palettes.itemById("FusionTimeSessionsPalette")
    if old:
        old.deleteMe()
    _palette = ui.palettes.add("FusionTimeSessionsPalette", "Fusion Time Sessions", "", True, True, True, 420, 520)
    _palette.isResizable = True
    _palette.isVisible   = True
    _palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateRight

    # hook HTML events
    html_in = HtmlIncomingHandler()
    _palette.incomingFromHTML.add(html_in)
    _handlers.append(html_in)

    url = _write_html()
    if url:
        _palette.htmlFileURL = url
        print(f"[TimeSessions] UI: {url}")
    else:
        print("[TimeSessions] UI NOT SET")

    # route print() to palette
    sys.stdout = UIPaletteWriter()
    sys.stderr = sys.stdout

# ======================= Entry points =======================
def run(context):
    try:
        app = adsk.core.Application.get()

        _create_palette_and_redirect_console()
        _ensure_csv(SESSIONS_CSV, SESSIONS_HEADER)

        # Only Activated + Closed to avoid duplicate opens
        events = {
            "DocumentActivated": app.documentActivated,
            "DocumentClosed":    app.documentClosed
        }
        for name, ev in events.items():
            h = DocEventHandler(name)
            ev.add(h)
            _handlers.append(h)

        # Start tracking current doc
        _switch_to(app.activeDocument, "startup")
        _append_ui(f"[TimeSessions] Tracking → {SESSIONS_CSV} (idle ≥ {IDLE_THRESHOLD_SEC}s)")

    except Exception:
        try:
            adsk.core.Application.get().userInterface.messageBox("Error in run():\n" + traceback.format_exc())
        except:
            pass

def stop(context):
    try:
        global _palette, _html_ready, _buffer
        _stop_current("stop")
        _handlers.clear()
        _html_ready = False
        _buffer = []
        if _palette:
            try: _palette.deleteMe()
            except: pass
            _palette = None
        try:
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
        except:
            pass
        print("[TimeSessions] Stopped.")
    except Exception:
        print("[TimeSessions] Error in stop():\n" + traceback.format_exc())

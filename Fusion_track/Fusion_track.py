import adsk.core, adsk.fusion, adsk.cam
import csv, tempfile, traceback, sys
from pathlib import Path
from datetime import datetime
from io import StringIO

# ======================= Config =======================
# Where to write session CSV (change if you want):
SESSIONS_CSV = Path("D:/New folder (3)/fusion_time_sessions.csv")  # <- per your D:\ path
SESSIONS_HEADER = ["project_name", "document_name", "start_time", "end_time", "duration_seconds", "duration_hms"]

# ======================= Globals =======================
_handlers = []
_palette = None
_html_ready = False
_buffer = []

# Active session state
_active_doc_key = None
_active_start   = None
_active_names   = ("Unsaved", "Unknown")  # (project_name, document_name)

# Temp HTML path
_temp_dir = Path(tempfile.gettempdir()) / "FusionTimeSessions"
_html_file = _temp_dir / "ui.html"

# ======================= Utilities =======================
def _safe_print(msg: str):
    try:
        print(msg)
    except:
        pass

def _append_ui(msg: str):
    global _html_ready, _buffer
    msg = msg.strip()
    if not msg:
        return
    try:
        if _palette and _html_ready:
            _palette.sendInfoToHTML("appendLog", msg)
        else:
            _buffer.append(msg)
    except:
        _buffer.append(msg)

def _flush_ui():
    global _buffer
    if not _palette:
        return
    for m in _buffer:
        try: _palette.sendInfoToHTML("appendLog", m)
        except: break
    _buffer = []

class UIPaletteWriter(StringIO):
    """Redirect print() to palette (with buffering until HTML ready)."""
    def write(self, s):
        if s and s.strip():
            _append_ui(s.strip())
        return len(s or "")

def _ensure_csv(path: Path, header):
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
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
    return f"{h:02d}:{m:02d}:{sec:02d}"

def _doc_key_and_names(doc: adsk.core.Document):
    """Return (unique_key, (project_name, document_name)) for a Document (doc can be None)."""
    proj, name = "Unsaved", "Unknown"
    key = "None"
    try:
        if doc:
            name = getattr(doc, "name", "Unknown") or "Unknown"
            df = getattr(doc, "dataFile", None)
            if df and getattr(df, "project", None) and getattr(df.project, "name", None):
                proj = df.project.name
            # Build a stable-ish key: prefer cloud URN if available, else project+name
            urn = getattr(df, "id", None) or getattr(df, "versionId", None) or ""
            key = urn if urn else f"{proj}|{name}"
    except:
        pass
    return key, (proj, name)

def _write_session_row(start_dt: datetime, end_dt: datetime, names):
    """Write one session line to CSV."""
    try:
        _ensure_csv(SESSIONS_CSV, SESSIONS_HEADER)
        dur = int((end_dt - start_dt).total_seconds())
        row = [names[0], names[1], start_dt.strftime("%Y-%m-%d %H:%M:%S"),
               end_dt.strftime("%Y-%m-%d %H:%M:%S"), dur, _hms(dur)]
        with open(SESSIONS_CSV, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(row)
            f.flush()
        _append_ui(f"[TimeSessions] Saved session | {names[0]} / {names[1]} | {_hms(dur)}")
    except Exception:
        _append_ui("[TimeSessions] Error writing CSV:\n" + traceback.format_exc())

def _start_session_for(doc: adsk.core.Document, reason: str):
    """Stop current session (if any) and start a new one for 'doc'."""
    global _active_doc_key, _active_start, _active_names
    # Stop existing session if running
    if _active_doc_key and _active_start:
        _stop_active_session("switch")
    # Start new session (if doc exists)
    key, names = _doc_key_and_names(doc)
    _active_doc_key = key
    _active_names = names
    _active_start = datetime.now()
    _append_ui(f"[TimeSessions] ▶ Start ({reason}) | {names[0]} / {names[1]} @ {_active_start.strftime('%H:%M:%S')}")

def _stop_active_session(reason: str):
    """Stop the current session and write CSV."""
    global _active_doc_key, _active_start, _active_names
    if not (_active_doc_key and _active_start):
        return
    end = datetime.now()
    _write_session_row(_active_start, end, _active_names)
    _append_ui(f"[TimeSessions] ■ Stop ({reason})  @ {end.strftime('%H:%M:%S')}")
    # Clear
    _active_doc_key = None
    _active_start = None

# ======================= Events =======================
class HtmlIncomingHandler(adsk.core.HTMLEventHandler):
    def notify(self, args):
        try:
            global _html_ready
            a = adsk.core.HTMLEventArgs.cast(args)
            if a and a.action == "ready":
                _html_ready = True
                _flush_ui()
        except:
            _append_ui("[TimeSessions] HTML incoming error:\n" + traceback.format_exc())

class DocEventHandler(adsk.core.DocumentEventHandler):
    def __init__(self, name):
        super().__init__()
        self.name = name  # event name for UI

    def notify(self, args):
        try:
            # Decide what to do based on event type
            if self.name in ("DocumentOpened", "DocumentActivated"):
                # Start/Restart on the newly active/opened doc
                doc = getattr(args, "document", None)
                _start_session_for(doc, self.name)
            elif self.name == "DocumentClosed":
                # If the closed doc is our active doc, stop that session
                closed_doc = getattr(args, "document", None)
                key_closed, _ = _doc_key_and_names(closed_doc)
                if _active_doc_key and key_closed == _active_doc_key:
                    _stop_active_session("close")
        except Exception:
            _append_ui(f"[TimeSessions] Error in {self.name}:\n" + traceback.format_exc())

# ======================= UI (palette) =======================
def _write_html():
    try:
        _temp_dir.mkdir(parents=True, exist_ok=True)
        html = """<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Fusion Time Sessions</title>
<style>
body { font-family: Consolas, monospace; background:#0d0f14; color:#d7f5d4; margin:0; }
h3 { margin:10px 12px; }
.wrap { padding: 0 12px 12px; }
#log { white-space: pre-line; font-size:12px; background:#0b0d12; border:1px solid #2a3340; padding:8px; height: calc(100vh - 80px); overflow:auto; }
small { color:#94a3b8; }
button { background:#1f2937; color:#d7f5d4; border:1px solid #334155; padding:6px 10px; cursor:pointer; margin-left: 12px; }
button:hover { background:#263241; }
</style>
<script>
function appendLog(msg){
  var el = document.getElementById('log');
  el.textContent += msg + "\\n";
  el.scrollTop = el.scrollHeight;
}
function clearLog(){ document.getElementById('log').textContent = ''; }
window.addEventListener('load', ()=> {
  try { adsk.fusionSendData('ready',''); } catch(e){}
});
</script>
</head>
<body>
  <h3>Fusion Time Sessions <small>(per-document)</small><button onclick="clearLog()">Clear</button></h3>
  <div class="wrap">
    <div id="log">--- UI Ready ---\\n</div>
  </div>
</body>
</html>"""
        _html_file.write_text(html, encoding="utf-8")
        return _html_file.as_uri()
    except Exception:
        _safe_print("[TimeSessions] Failed to write HTML:\n" + traceback.format_exc())
        return ""

def _create_palette():
    global _palette
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        old = ui.palettes.itemById("FusionTimeSessionsPalette")
        if old:
            old.deleteMe()

        _palette = ui.palettes.add("FusionTimeSessionsPalette", "Fusion Time Sessions", "", True, True, True, 420, 520)
        _palette.isResizable = True
        _palette.isVisible = True
        _palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateRight

        # incoming events handler (for 'ready')
        html_in = HtmlIncomingHandler()
        _palette.incomingFromHTML.add(html_in)
        _handlers.append(html_in)

        url = _write_html()
        if url:
            _palette.htmlFileURL = url
            _safe_print(f"[TimeSessions] UI: {url}")
        else:
            _safe_print("[TimeSessions] UI NOT SET")

    except Exception:
        _safe_print("[TimeSessions] Failed to create palette:\n" + traceback.format_exc())

# ======================= Entry Points =======================
def run(context):
    try:
        app = adsk.core.Application.get()

        # UI first, then redirect stdout to palette (buffer until 'ready')
        _create_palette()
        sys.stdout = UIPaletteWriter()
        sys.stderr = sys.stdout

        # Ensure CSV exists
        _ensure_csv(SESSIONS_CSV, SESSIONS_HEADER)

        # Subscribe to doc events
        events = {
            "DocumentOpened": app.documentOpened,
            "DocumentActivated": app.documentActivated,
            "DocumentClosed": app.documentClosed
        }
        for name, ev in events.items():
            h = DocEventHandler(name)
            ev.add(h)
            _handlers.append(h)

        # If there is an active document at startup, begin a session right away
        _start_session_for(app.activeDocument, "startup")

        _append_ui(f"[TimeSessions] Tracking sessions → {SESSIONS_CSV}")

    except Exception:
        try:
            adsk.core.Application.get().userInterface.messageBox("Error in run():\n" + traceback.format_exc())
        except:
            pass

def stop(context):
    try:
        global _palette, _html_ready, _buffer, _active_doc_key, _active_start

        # Stop whatever is active and write last session
        _stop_active_session("stop")

        _html_ready = False
        _buffer = []
        # Clear handlers
        _handlers.clear()

        # Remove palette
        if _palette:
            try: _palette.deleteMe()
            except: pass
            _palette = None

        # Restore normal stdout
        try:
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
        except:
            pass

        print("[TimeSessions] Stopped.")
    except Exception:
        print("[TimeSessions] Error in stop():\n" + traceback.format_exc())

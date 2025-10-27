import adsk.core, adsk.fusion, adsk.cam, traceback, csv
from datetime import datetime
from pathlib import Path

# ---------------- Globals ----------------
_handlers = []

# Preferred log path order
# _PRIMARY      = Path("D:/New folder (3)/fusion_event_log.csv")
_PRIMARY      = Path.home() / "Desktop" / "fusion_event_log.csv"
_FALLBACK_DIR = Path("D:/New folder (3)/FusionLogs")
_DESKTOP      = Path.home() / "Desktop" / "fusion_event_log.csv"
_log_file: Path = None  # resolved in run()

# ---------------- Path & CSV helpers ----------------
def _try_open_for_append(p: Path) -> bool:
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, 'a', encoding='utf-8'):
            pass
        return True
    except Exception:
        print(f"[FusionEventLogger] Cannot write to: {p}\n{traceback.format_exc()}")
        return False

def _resolve_log_path() -> Path:
    if _try_open_for_append(_PRIMARY):
        print(f"[FusionEventLogger] Using log file: {_PRIMARY}")
        return _PRIMARY
    fb = _FALLBACK_DIR / "fusion_event_log.csv"
    if _try_open_for_append(fb):
        print(f"[FusionEventLogger] Using fallback log file: {fb}")
        return fb
    if _try_open_for_append(_DESKTOP):
        print(f"[FusionEventLogger] Using Desktop log file: {_DESKTOP}")
        return _DESKTOP
    # last resort (still return Desktop so later attempts are consistent)
    print("[FusionEventLogger] ERROR: No writable location found. Logging only to console.")
    return _DESKTOP

def _ensure_csv_exists():
    global _log_file
    if _log_file is None:
        _log_file = _resolve_log_path()
    try:
        if not _log_file.exists():
            with open(_log_file, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(["timestamp", "event_name", "project_name", "document_name"])
            print(f"[FusionEventLogger] Created CSV at: {_log_file}")
    except Exception:
        print("[FusionEventLogger] Failed to create CSV:\n" + traceback.format_exc())

def _append_row(event_name: str, doc: adsk.core.Document):
    """Append one record to CSV; safe if doc is None."""
    global _log_file
    if _log_file is None:
        _log_file = _resolve_log_path()

    try:
        _ensure_csv_exists()

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        project_name = "Unsaved"
        document_name = "Unknown"

        if doc:
            try:
                if getattr(doc, 'name', None):
                    document_name = doc.name
                df = getattr(doc, 'dataFile', None)
                if df and getattr(df, 'project', None) and getattr(df.project, 'name', None):
                    project_name = df.project.name
            except Exception:
                pass

        wrote = False
        try:
            with open(_log_file, 'a', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow([ts, event_name, project_name, document_name])
                f.flush()
            wrote = True
        except Exception:
            print(f"[FusionEventLogger] Write failed: {_log_file}\n" + traceback.format_exc())

        # Always echo to console
        print(f"{ts} | {event_name} | {project_name} | {document_name}")

        # If initial write failed, try re-resolving path once and retry
        if not wrote:
            newp = _resolve_log_path()
            if newp != _log_file:
                _log_file = newp
                _ensure_csv_exists()
                try:
                    with open(_log_file, 'a', newline='', encoding='utf-8') as f:
                        csv.writer(f).writerow([ts, event_name, project_name, document_name])
                        f.flush()
                except Exception:
                    print(f"[FusionEventLogger] Retry write failed: {_log_file}\n" + traceback.format_exc())

    except Exception:
        print("[FusionEventLogger] Error writing to CSV:\n" + traceback.format_exc())

# ---------------- Event Handler ----------------
class FusionDocEventHandler(adsk.core.DocumentEventHandler):
    def __init__(self, event_name: str):
        super().__init__()
        self.event_name = event_name

    def notify(self, args):
        try:
            doc = None
            try:
                doc = args.document  # may be None (e.g., DocumentClosed)
            except Exception:
                pass
            _append_row(self.event_name, doc)
        except Exception:
            print(f"[FusionEventLogger] Error in {self.event_name}:\n" + traceback.format_exc())

# ---------------- Entry Points ----------------
def run(context):
    try:
        app = adsk.core.Application.get()

        # Resolve path & ensure CSV up front
        global _log_file
        _log_file = _resolve_log_path()
        _ensure_csv_exists()

        # Subscribe to supported document events
        events = {
            "DocumentOpened": app.documentOpened,
            "DocumentActivated": app.documentActivated,
            "DocumentClosed": app.documentClosed
        }
        for name, ev in events.items():
            h = FusionDocEventHandler(name)
            ev.add(h)
            _handlers.append(h)

        # Immediate test row so you can confirm it logs even before events
        _append_row("LoggerStarted", app.activeDocument)

        print("[FusionEventLogger] Event logger started.")

    except Exception:
        # No message box: stay silent except console
        print("[FusionEventLogger] Error in run():\n" + traceback.format_exc())

def stop(context=None):
    try:
        global _handlers
        _handlers.clear()
        print("[FusionEventLogger] Event logger stopped.")
    except Exception:
        print("[FusionEventLogger] Error in stop():\n" + traceback.format_exc())

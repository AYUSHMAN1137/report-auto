import os
import sys
import time
import threading
from pathlib import Path
from typing import Callable, Dict, Optional, Any

from flask import Flask, render_template, jsonify, request, send_from_directory, abort
from flask_socketio import SocketIO


class ReportDashboard:
    def __init__(self, port: int = 5001, host: str = "127.0.0.1"):
        self.app = Flask(__name__, template_folder="templates", static_folder="static")
        self.app.config['SECRET_KEY'] = 'thyrocare-report-dashboard'
        # threading mode avoids extra deps (gevent/eventlet)
        self.socketio = SocketIO(self.app, cors_allowed_origins="*", async_mode="threading")
        self.port = port
        self.host = host

        self.lock = threading.Lock()
        self.stats: Dict[str, Any] = {
            'total_barcodes': 0,
            'done_barcodes': 0,
            'total_pdfs': 0,
            'done_pdfs': 0,
            'current_step': '-',
            'current_step_index': 0,
            'current_step_total': 5,
            'current_barcode': '-',
            'status': 'idle',
            'logs': [],
            'download_folder': "",
            'whatsapp_contact': "",
            'whatsapp_status': 'pending',
            'whatsapp_preset_contacts': [
                'Ayushman1137',
                'Mummy',
                'Gaur Report',
                'Thyrocare Report'
            ]
        }

        self.on_control: Optional[Callable[[str, Dict[str, Any]], Dict[str, Any]]] = None
        self.download_folder: Optional[Path] = None
        self.logs_history: list[str] = []
        self._stop_threads = threading.Event()
        self.start_time = time.time()

        self._setup_routes()

        # Background threads placeholders (reserved for future use)
        self._sys_thread = None
        self._server_thread: Optional[threading.Thread] = None

    # ----- Public API -----
    def set_download_folder(self, folder: Path) -> None:
        try:
            self.download_folder = Path(folder)
            self.stats['download_folder'] = str(self.download_folder)
        except Exception:
            self.download_folder = None
            self.stats['download_folder'] = ""

    def set_whatsapp_contact(self, contact: str) -> None:
        with self.lock:
            self.stats['whatsapp_contact'] = contact
    
    def get_whatsapp_contact(self) -> str:
        with self.lock:
            return self.stats.get('whatsapp_contact', 'Ayushman1137')

    def set_handlers(self, on_control: Optional[Callable[[str, Dict[str, Any]], Dict[str, Any]]] = None) -> None:
        self.on_control = on_control

    def update_stats(self, data: Dict[str, Any]) -> None:
        with self.lock:
            self.stats.update(data or {})
            self.socketio.emit('update', self.stats)

    def add_log(self, message: str) -> None:
        with self.lock:
            timestamp = time.strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {message}"
            self.logs_history.append(log_entry)
            # Keep only last 200 logs in history
            if len(self.logs_history) > 200:
                self.logs_history = self.logs_history[-200:]
            # Include last 50 logs in full-state updates
            self.stats['logs'] = self.logs_history[-50:]
            self.socketio.emit('log', {'message': log_entry})

    def clear_logs(self) -> None:
        with self.lock:
            self.logs_history.clear()
            self.stats['logs'] = []
        self.socketio.emit('update', self.stats)

    def start_background(self) -> str:
        # Start Flask-SocketIO server in a background thread
        if self._server_thread and self._server_thread.is_alive():
            return f"http://{self.host}:{self.port}"

        def run_server():
            # Avoid passing unsupported args when async_mode="threading"
            self.socketio.run(
                self.app,
                host=self.host,
                port=self.port,
                debug=False,
                use_reloader=False
            )

        self._server_thread = threading.Thread(target=run_server, daemon=True)
        self._server_thread.start()
        print(f"Report Dashboard running at http://{self.host}:{self.port}")
        return f"http://{self.host}:{self.port}"

    # ----- Internal: routes -----
    def _setup_routes(self) -> None:
        @self.app.route("/")
        def index():
            return render_template("dashboard.html")

        @self.app.route("/stats")
        def get_stats():
            with self.lock:
                return jsonify(self.stats)

        @self.app.route("/api/control", methods=["POST"])
        def api_control():
            data = request.get_json(silent=True) or {}
            action = (data.get("action") or "").lower()

            # Server-side actions
            if action == "open_download_folder":
                try:
                    if self.download_folder and self.download_folder.exists():
                        if os.name == "nt":
                            os.startfile(str(self.download_folder))  # type: ignore[attr-defined]
                        else:
                            opener = "open" if sys.platform == "darwin" else "xdg-open"
                            os.system(f'{opener} "{self.download_folder}"')
                        return jsonify({"ok": True})
                    return jsonify({"ok": False, "error": "Folder not set or missing"}), 400
                except Exception as e:
                    return jsonify({"ok": False, "error": str(e)}), 500

            if action == "clear_logs":
                self.clear_logs()
                return jsonify({"ok": True})
            
            if action == "clean_profile":
                # Forward to main handler
                if self.on_control:
                    try:
                        res = self.on_control("clean_profile", {})
                        return jsonify(res or {"ok": True})
                    except Exception as e:
                        return jsonify({"ok": False, "error": str(e)}), 500
                return jsonify({"ok": False, "error": "No handler"}), 400
            
            if action == "set_whatsapp_contact":
                contact = data.get("contact", "").strip()
                if contact:
                    self.set_whatsapp_contact(contact)
                    return jsonify({"ok": True, "contact": contact})
                return jsonify({"ok": False, "error": "No contact provided"}), 400

            if self.on_control:
                try:
                    res = self.on_control(action, data)
                    return jsonify(res or {"ok": True})
                except Exception as e:
                    return jsonify({"ok": False, "error": str(e)}), 500
            return jsonify({"ok": False, "error": "No control handler"}), 400

        @self.app.route("/api/start", methods=["POST"])
        def api_start():
            data = request.get_json(silent=True) or {}
            if self.on_control:
                try:
                    res = self.on_control("start", data)
                    return jsonify(res or {"ok": True})
                except Exception as e:
                    return jsonify({"ok": False, "error": str(e)}), 500
            return jsonify({"ok": False, "error": "No start handler"}), 400

        @self.app.route("/api/cancel", methods=["POST"])
        def api_cancel():
            if self.on_control:
                try:
                    res = self.on_control("cancel", {})
                    return jsonify(res or {"ok": True})
                except Exception as e:
                    return jsonify({"ok": False, "error": str(e)}), 500
            return jsonify({"ok": False, "error": "No cancel handler"}), 400

        @self.app.route("/download/<path:filename>")
        def serve_download(filename: str):
            try:
                if not self.download_folder or not self.download_folder.exists():
                    abort(404)

                file_path = (self.download_folder / filename).resolve()

                # Security: ensure requested file is inside download folder tree
                try:
                    file_path.relative_to(self.download_folder.resolve())
                except ValueError:
                    abort(403)

                if not file_path.exists() or not file_path.is_file():
                    abort(404)

                return send_from_directory(str(file_path.parent), file_path.name)
            except Exception:
                abort(404)

    # ----- Public API: Error Handling -----
    def send_error(self, title: str, message: str, screenshot_path: Optional[str] = None) -> None:
        """Send error with optional screenshot to dashboard (via 'ui_error' event)."""
        error_data: Dict[str, Any] = {
            'title': title,
            'message': message
        }

        if screenshot_path:
            shot_path = Path(screenshot_path)
            if shot_path.exists() and self.download_folder:
                try:
                    rel_path = shot_path.relative_to(self.download_folder)
                    error_data['screenshot'] = f"/download/{rel_path.as_posix()}"
                except ValueError:
                    # If not relative to download folder, attempt a screenshots/ mapping
                    error_data['screenshot'] = f"/download/screenshots/{shot_path.name}"

        self.socketio.emit('ui_error', error_data)
        self.add_log(f"❌ {title}: {message}")
"""
DOP Automation Dashboard
========================
Live web dashboard + control panel for the DOP Agent Portal automation.
Runs as a Flask server on a daemon thread alongside the Selenium automation.

Usage (standalone test):
    python3 dop_dashboard.py

Usage (from dop_automate.py):
    from dop_dashboard import (
        DashboardState, ControlFlags, SkipLotException,
        StopAfterCurrentException, checkpoint, start_dashboard
    )
"""

import json
import time
import threading
from dataclasses import dataclass, field
from collections import deque
from datetime import datetime


# ── Exceptions ──

class SkipLotException(Exception):
    """Raised when user requests to skip the current LOT via dashboard."""
    pass


class StopAfterCurrentException(Exception):
    """Raised when user requests to stop after the current LOT finishes."""
    pass


# ── Shared State ──

@dataclass
class DashboardState:
    lock: threading.Lock = field(default_factory=threading.Lock)

    # Progress
    current_phase: str = "Startup"
    current_lot: str = ""
    current_step: str = ""
    lots_done: int = 0
    lots_total: int = 0
    lots_skipped: int = 0
    lots_failed: int = 0

    # Per-LOT status list
    lot_statuses: list = field(default_factory=list)

    # System
    memory_mb: float = 0.0
    start_time: float = 0.0
    is_paused: bool = False
    is_finished: bool = False

    # Log buffer (last 80 messages)
    log_messages: deque = field(default_factory=lambda: deque(maxlen=80))

    # Live-editable config
    delay_short: float = 1.5
    delay_medium: float = 3.0
    delay_long: float = 5.0
    delay_checkbox: float = 0.4

    def to_dict(self):
        with self.lock:
            return {
                "current_phase": self.current_phase,
                "current_lot": self.current_lot,
                "current_step": self.current_step,
                "lots_done": self.lots_done,
                "lots_total": self.lots_total,
                "lots_skipped": self.lots_skipped,
                "lots_failed": self.lots_failed,
                "lot_statuses": [dict(s) for s in self.lot_statuses],
                "memory_mb": round(self.memory_mb, 1),
                "elapsed_seconds": int(time.time() - self.start_time) if self.start_time else 0,
                "is_paused": self.is_paused,
                "is_finished": self.is_finished,
                "log_messages": list(self.log_messages),
                "config": {
                    "delay_short": self.delay_short,
                    "delay_medium": self.delay_medium,
                    "delay_long": self.delay_long,
                    "delay_checkbox": self.delay_checkbox,
                },
            }


@dataclass
class ControlFlags:
    pause_event: threading.Event = field(default_factory=threading.Event)
    skip_lot: threading.Event = field(default_factory=threading.Event)
    stop_after_current: threading.Event = field(default_factory=threading.Event)
    skip_lots_set: set = field(default_factory=set)
    lock: threading.Lock = field(default_factory=threading.Lock)

    def __post_init__(self):
        self.pause_event.set()  # Start in running state


# ── Checkpoint ──

def checkpoint(state: DashboardState, control: ControlFlags, step_name: str = ""):
    """Called between automation steps. Blocks if paused, raises on skip/stop."""
    if step_name:
        with state.lock:
            state.current_step = step_name
    if not control.pause_event.wait(timeout=300):
        # Still paused after 5 minutes — log a warning but keep waiting
        with state.lock:
            state.log_messages.append(
                f"{time.strftime('%H:%M:%S')}  WARNING: Paused for 5+ minutes at '{step_name}'"
            )
        control.pause_event.wait()
    if control.skip_lot.is_set():
        control.skip_lot.clear()
        raise SkipLotException()
    if control.stop_after_current.is_set():
        raise StopAfterCurrentException()


# ── Flask App ──

def _create_app(state: DashboardState, control: ControlFlags):
    from flask import Flask, Response, request, jsonify

    app = Flask(__name__)

    @app.route("/")
    def index():
        return DASHBOARD_HTML

    @app.route("/api/state")
    def sse_stream():
        def generate():
            try:
                while True:
                    data = json.dumps(state.to_dict())
                    yield f"data: {data}\n\n"
                    if state.is_finished:
                        break
                    time.sleep(0.5)
            except GeneratorExit:
                return
        return Response(generate(), mimetype="text/event-stream",
                        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

    @app.route("/api/control", methods=["POST"])
    def handle_control():
        body = request.json or {}
        action = body.get("action")

        if action == "pause":
            control.pause_event.clear()
            with state.lock:
                state.is_paused = True
            return jsonify({"ok": True, "status": "paused"})

        elif action == "resume":
            control.pause_event.set()
            with state.lock:
                state.is_paused = False
            return jsonify({"ok": True, "status": "resumed"})

        elif action == "skip":
            control.skip_lot.set()
            control.pause_event.set()
            with state.lock:
                state.is_paused = False
            return jsonify({"ok": True, "status": "skipping"})

        elif action == "stop_after_current":
            control.stop_after_current.set()
            return jsonify({"ok": True, "status": "stopping"})

        elif action == "update_config":
            config = body.get("config", {})
            try:
                with state.lock:
                    if "delay_short" in config:
                        state.delay_short = max(0.1, float(config["delay_short"]))
                    if "delay_medium" in config:
                        state.delay_medium = max(0.1, float(config["delay_medium"]))
                    if "delay_long" in config:
                        state.delay_long = max(0.1, float(config["delay_long"]))
                    if "delay_checkbox" in config:
                        state.delay_checkbox = max(0.05, float(config["delay_checkbox"]))
            except (ValueError, TypeError) as e:
                return jsonify({"ok": False, "error": f"Invalid config value: {e}"}), 400
            return jsonify({"ok": True})

        elif action == "toggle_lot":
            lot_num = str(body.get("lot", ""))
            with control.lock:
                if lot_num in control.skip_lots_set:
                    control.skip_lots_set.discard(lot_num)
                else:
                    control.skip_lots_set.add(lot_num)
            return jsonify({"ok": True, "skip_lots": list(control.skip_lots_set)})

        return jsonify({"ok": False, "error": "unknown action"}), 400

    return app


def start_dashboard(state: DashboardState, control: ControlFlags, port=5555):
    """Start Flask dashboard on a daemon thread. Tries ports 5555-5560."""
    app = _create_app(state, control)

    import logging
    log = logging.getLogger("werkzeug")
    log.setLevel(logging.ERROR)

    import socket

    for p in range(port, port + 6):
        # Pre-check if port is available
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", p))
        except OSError:
            continue

        def _run(port_num):
            try:
                app.run(host="127.0.0.1", port=port_num, debug=False, use_reloader=False)
            except OSError:
                pass

        thread = threading.Thread(target=_run, args=(p,), daemon=True)
        thread.start()

        # Poll until Flask is actually accepting TCP connections (up to 3 s).
        # This replaces the old before_request approach which required a browser
        # to connect within 3 s — that never happened, so all 6 ports got threads
        # and "Could not start dashboard" was always printed.
        deadline = time.time() + 3.0
        while time.time() < deadline:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(0.1)
                    s.connect(("127.0.0.1", p))
                print(f"  Dashboard running at http://127.0.0.1:{p}")
                return thread
            except OSError:
                time.sleep(0.05)

    print("  Could not start dashboard (ports 5555-5560 in use)")
    return None


# ── Dashboard HTML ──

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DOP Automation</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
    --bg-primary: #0f1117;
    --bg-card: #1a1d27;
    --bg-card-hover: #1f2233;
    --bg-input: #252836;
    --border: #2a2d3a;
    --border-light: #353849;
    --text-primary: #e4e6ef;
    --text-secondary: #8b8fa3;
    --text-muted: #5d6177;
    --accent: #6c5ce7;
    --accent-light: #a29bfe;
    --green: #00cec9;
    --green-bg: rgba(0,206,201,0.12);
    --red: #ff6b6b;
    --red-bg: rgba(255,107,107,0.12);
    --yellow: #feca57;
    --yellow-bg: rgba(254,202,87,0.12);
    --blue: #54a0ff;
    --blue-bg: rgba(84,160,255,0.12);
    --radius: 12px;
    --radius-sm: 8px;
    --shadow: 0 4px 24px rgba(0,0,0,0.3);
}

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    line-height: 1.6;
    min-height: 100vh;
}

.container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 24px 20px;
}

/* Header */
.header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
    padding-bottom: 20px;
    border-bottom: 1px solid var(--border);
}

.header-left { display: flex; align-items: center; gap: 14px; }

.logo {
    width: 40px; height: 40px;
    background: linear-gradient(135deg, var(--accent), var(--accent-light));
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 18px; color: #fff;
}

.header h1 { font-size: 20px; font-weight: 600; letter-spacing: -0.3px; }
.header-sub { font-size: 12px; color: var(--text-secondary); margin-top: 2px; }

.header-right { display: flex; align-items: center; gap: 16px; }

.timer {
    font-size: 28px; font-weight: 600;
    font-variant-numeric: tabular-nums;
    color: var(--text-primary);
    letter-spacing: 1px;
}

.connection-dot {
    width: 8px; height: 8px; border-radius: 50%;
    background: var(--green);
    box-shadow: 0 0 8px var(--green);
    animation: pulse 2s infinite;
}
.connection-dot.disconnected { background: var(--red); box-shadow: 0 0 8px var(--red); }

@keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }

/* Status Cards */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 14px;
    margin-bottom: 24px;
}

.stat-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 18px 20px;
    transition: border-color 0.2s;
}
.stat-card:hover { border-color: var(--border-light); }

.stat-label {
    font-size: 11px; font-weight: 500;
    text-transform: uppercase; letter-spacing: 0.8px;
    color: var(--text-secondary);
    margin-bottom: 8px;
}

.stat-value {
    font-size: 24px; font-weight: 700;
    font-variant-numeric: tabular-nums;
}

.stat-value.green { color: var(--green); }
.stat-value.red { color: var(--red); }
.stat-value.yellow { color: var(--yellow); }
.stat-value.blue { color: var(--blue); }

/* Progress Section */
.progress-section {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px;
    margin-bottom: 24px;
}

.progress-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 6px;
}

.progress-phase {
    font-size: 14px; font-weight: 600; color: var(--accent-light);
}

.progress-step {
    font-size: 13px; color: var(--text-secondary);
    margin-bottom: 16px;
}

.progress-bar-container {
    width: 100%; height: 8px;
    background: var(--bg-input);
    border-radius: 4px;
    overflow: hidden;
    margin-bottom: 8px;
}

.progress-bar {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--accent-light));
    border-radius: 4px;
    transition: width 0.5s ease;
    min-width: 0;
}

.progress-text {
    font-size: 12px; color: var(--text-muted);
    text-align: right;
    font-variant-numeric: tabular-nums;
}

/* Paused overlay */
.paused-badge {
    display: none;
    align-items: center; gap: 8px;
    background: var(--yellow-bg);
    border: 1px solid rgba(254,202,87,0.3);
    color: var(--yellow);
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 12px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.5px;
}
.paused-badge.visible { display: inline-flex; }

/* Two-column layout */
.two-col {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 24px;
}
@media (max-width: 768px) { .two-col { grid-template-columns: 1fr; } }

/* Controls card */
.card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px;
}

.card-title {
    font-size: 13px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.8px;
    color: var(--text-secondary);
    margin-bottom: 18px;
}

.controls-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
}

.btn {
    padding: 10px 16px;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--bg-input);
    color: var(--text-primary);
    font-size: 13px; font-weight: 500;
    cursor: pointer;
    transition: all 0.2s;
    display: flex; align-items: center; justify-content: center; gap: 6px;
}
.btn:hover { background: var(--bg-card-hover); border-color: var(--border-light); }
.btn:active { transform: scale(0.97); }

.btn-pause {
    background: var(--accent);
    border-color: var(--accent);
    color: #fff;
    grid-column: span 2;
}
.btn-pause:hover { background: var(--accent-light); border-color: var(--accent-light); }
.btn-pause.paused {
    background: var(--green);
    border-color: var(--green);
}

.btn-skip { border-color: var(--yellow); color: var(--yellow); }
.btn-skip:hover { background: var(--yellow-bg); }

.btn-stop { border-color: var(--red); color: var(--red); grid-column: span 2; }
.btn-stop:hover { background: var(--red-bg); }

/* Delay sliders */
.delay-group { margin-bottom: 16px; }
.delay-group:last-child { margin-bottom: 0; }

.delay-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 6px;
}

.delay-label { font-size: 13px; color: var(--text-secondary); }

.delay-value {
    font-size: 13px; font-weight: 600;
    color: var(--text-primary);
    font-variant-numeric: tabular-nums;
    min-width: 40px; text-align: right;
}

input[type="range"] {
    -webkit-appearance: none; appearance: none;
    width: 100%; height: 4px;
    background: var(--bg-input);
    border-radius: 2px;
    outline: none;
}
input[type="range"]::-webkit-slider-thumb {
    -webkit-appearance: none; appearance: none;
    width: 16px; height: 16px;
    background: var(--accent-light);
    border-radius: 50%;
    cursor: pointer;
    border: 2px solid var(--bg-card);
    box-shadow: 0 0 6px rgba(108,92,231,0.4);
}
input[type="range"]::-moz-range-thumb {
    width: 16px; height: 16px;
    background: var(--accent-light);
    border-radius: 50%;
    cursor: pointer;
    border: 2px solid var(--bg-card);
}

/* LOT Table */
.table-container {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    margin-bottom: 24px;
}

.table-header {
    padding: 18px 24px 14px;
    border-bottom: 1px solid var(--border);
}

.table-scroll {
    max-height: min(320px, 40vh);
    overflow-y: auto;
}
.table-scroll::-webkit-scrollbar { width: 6px; }
.table-scroll::-webkit-scrollbar-track { background: transparent; }
.table-scroll::-webkit-scrollbar-thumb { background: var(--border-light); border-radius: 3px; }

table { width: 100%; border-collapse: collapse; }

th {
    padding: 10px 16px;
    font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.6px;
    color: var(--text-muted);
    text-align: left;
    background: var(--bg-input);
    position: sticky; top: 0; z-index: 1;
}

td {
    padding: 10px 16px;
    font-size: 13px;
    border-bottom: 1px solid var(--border);
    font-variant-numeric: tabular-nums;
}

tr:last-child td { border-bottom: none; }
tr.running { background: rgba(108,92,231,0.06); }

/* Status pills */
.pill {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.4px;
}
.pill-done { background: var(--green-bg); color: var(--green); }
.pill-running { background: var(--blue-bg); color: var(--blue); animation: pulse 1.5s infinite; }
.pill-failed { background: var(--red-bg); color: var(--red); }
.pill-skipped { background: var(--yellow-bg); color: var(--yellow); }
.pill-pending { background: rgba(93,97,119,0.15); color: var(--text-muted); }

.skip-btn {
    padding: 3px 8px;
    border: 1px solid var(--border);
    border-radius: 6px;
    background: transparent;
    color: var(--text-muted);
    font-size: 11px;
    cursor: pointer;
    transition: all 0.2s;
}
.skip-btn:hover { border-color: var(--yellow); color: var(--yellow); }
.skip-btn.active { border-color: var(--yellow); color: var(--yellow); background: var(--yellow-bg); }

/* Log panel */
.log-container {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
}

.log-header {
    padding: 18px 24px 14px;
    border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
}

.log-scroll {
    max-height: min(260px, 30vh);
    overflow-y: auto;
    padding: 12px 0;
}
.log-scroll::-webkit-scrollbar { width: 6px; }
.log-scroll::-webkit-scrollbar-track { background: transparent; }
.log-scroll::-webkit-scrollbar-thumb { background: var(--border-light); border-radius: 3px; }

.log-line {
    padding: 3px 24px;
    font-family: 'SF Mono', 'Fira Code', 'Cascadia Code', 'Consolas', monospace;
    font-size: 12px;
    color: var(--text-secondary);
    line-height: 1.7;
    white-space: pre-wrap;
    word-break: break-all;
}
.log-line:hover { background: rgba(255,255,255,0.02); }

/* Finished state */
.finished-banner {
    display: none;
    align-items: center;
    justify-content: center;
    gap: 10px;
    padding: 14px 24px;
    background: var(--green-bg);
    border: 1px solid rgba(0,206,201,0.3);
    border-radius: var(--radius);
    margin-bottom: 24px;
    color: var(--green);
    font-weight: 600;
    font-size: 14px;
}
.finished-banner.visible { display: flex; }
</style>
</head>
<body>
<div class="container">

    <!-- Header -->
    <div class="header">
        <div class="header-left">
            <div class="logo">D</div>
            <div>
                <h1>DOP Automation</h1>
                <div class="header-sub">Agent Portal | RD Installments</div>
            </div>
        </div>
        <div class="header-right">
            <span class="paused-badge" id="pausedBadge">PAUSED</span>
            <div class="timer" id="timer">00:00</div>
            <div class="connection-dot" id="connectionDot"></div>
        </div>
    </div>

    <!-- Finished banner -->
    <div class="finished-banner" id="finishedBanner">Automation Complete</div>

    <!-- Stats -->
    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-label">Phase</div>
            <div class="stat-value blue" id="statPhase">--</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Current LOT</div>
            <div class="stat-value" id="statLot">--</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Completed</div>
            <div class="stat-value green" id="statDone">0</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Failed</div>
            <div class="stat-value red" id="statFailed">0</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Skipped</div>
            <div class="stat-value yellow" id="statSkipped">0</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Memory</div>
            <div class="stat-value" id="statMemory">--</div>
        </div>
    </div>

    <!-- Progress -->
    <div class="progress-section">
        <div class="progress-header">
            <span class="progress-phase" id="progressPhase">Waiting...</span>
        </div>
        <div class="progress-step" id="progressStep">--</div>
        <div class="progress-bar-container">
            <div class="progress-bar" id="progressBar" style="width:0%"></div>
        </div>
        <div class="progress-text" id="progressText">0 / 0</div>
    </div>

    <!-- Controls + Delays -->
    <div class="two-col">
        <div class="card">
            <div class="card-title">Controls</div>
            <div class="controls-grid">
                <button class="btn btn-pause" id="btnPause" onclick="togglePause()">Pause</button>
                <button class="btn btn-skip" onclick="sendControl('skip')">Skip LOT</button>
                <button class="btn btn-stop" onclick="sendControl('stop_after_current')">Stop After Current</button>
            </div>
        </div>
        <div class="card">
            <div class="card-title">Delays (Live)</div>
            <div class="delay-group">
                <div class="delay-header">
                    <span class="delay-label">Short</span>
                    <span class="delay-value" id="valShort">1.5s</span>
                </div>
                <input type="range" min="0.1" max="5" step="0.1" value="1.5"
                       id="sliderShort" oninput="updateDelay('delay_short', this.value, 'valShort')">
            </div>
            <div class="delay-group">
                <div class="delay-header">
                    <span class="delay-label">Medium</span>
                    <span class="delay-value" id="valMedium">3.0s</span>
                </div>
                <input type="range" min="0.5" max="10" step="0.1" value="3.0"
                       id="sliderMedium" oninput="updateDelay('delay_medium', this.value, 'valMedium')">
            </div>
            <div class="delay-group">
                <div class="delay-header">
                    <span class="delay-label">Long</span>
                    <span class="delay-value" id="valLong">5.0s</span>
                </div>
                <input type="range" min="1" max="15" step="0.1" value="5.0"
                       id="sliderLong" oninput="updateDelay('delay_long', this.value, 'valLong')">
            </div>
            <div class="delay-group">
                <div class="delay-header">
                    <span class="delay-label">Checkbox</span>
                    <span class="delay-value" id="valCheckbox">0.4s</span>
                </div>
                <input type="range" min="0.05" max="2" step="0.05" value="0.4"
                       id="sliderCheckbox" oninput="updateDelay('delay_checkbox', this.value, 'valCheckbox')">
            </div>
        </div>
    </div>

    <!-- LOT Table -->
    <div class="table-container">
        <div class="table-header">
            <div class="card-title" style="margin-bottom:0">LOT Status</div>
        </div>
        <div class="table-scroll">
            <table>
                <thead>
                    <tr>
                        <th>LOT</th>
                        <th>Count</th>
                        <th>Status</th>
                        <th>Reference ID</th>
                        <th>Step</th>
                        <th></th>
                    </tr>
                </thead>
                <tbody id="lotTableBody"></tbody>
            </table>
        </div>
    </div>

    <!-- Log -->
    <div class="log-container">
        <div class="log-header">
            <div class="card-title" style="margin-bottom:0">Live Log</div>
            <span style="font-size:11px; color:var(--text-muted)" id="logCount">0 entries</span>
        </div>
        <div class="log-scroll" id="logScroll">
            <div id="logBody"></div>
        </div>
    </div>

</div>

<script>
let isPaused = false;
let skipLots = new Set();
let debounceTimers = {};

// SSE connection with exponential backoff
let sseBackoff = 2000;
function connectSSE() {
    const dot = document.getElementById('connectionDot');
    const evtSource = new EventSource('/api/state');

    evtSource.onmessage = function(e) {
        dot.classList.remove('disconnected');
        sseBackoff = 2000;  // Reset on successful message
        const data = JSON.parse(e.data);
        updateDashboard(data);
    };

    evtSource.onerror = function() {
        dot.classList.add('disconnected');
        evtSource.close();
        setTimeout(connectSSE, sseBackoff);
        sseBackoff = Math.min(sseBackoff * 2, 16000);
    };
}

function updateDashboard(d) {
    // Timer
    const mins = Math.floor(d.elapsed_seconds / 60);
    const secs = d.elapsed_seconds % 60;
    document.getElementById('timer').textContent =
        String(mins).padStart(2, '0') + ':' + String(secs).padStart(2, '0');

    // Finished
    const fb = document.getElementById('finishedBanner');
    fb.classList.toggle('visible', d.is_finished);

    // Stats
    document.getElementById('statPhase').textContent = d.current_phase || '--';
    document.getElementById('statLot').textContent = d.current_lot || '--';
    document.getElementById('statDone').textContent = d.lots_done;
    document.getElementById('statFailed').textContent = d.lots_failed;
    document.getElementById('statSkipped').textContent = d.lots_skipped;
    document.getElementById('statMemory').textContent =
        d.memory_mb > 0 ? d.memory_mb.toFixed(0) + ' MB' : '--';

    // Paused badge
    isPaused = d.is_paused;
    document.getElementById('pausedBadge').classList.toggle('visible', d.is_paused);
    const btn = document.getElementById('btnPause');
    btn.textContent = d.is_paused ? 'Resume' : 'Pause';
    btn.classList.toggle('paused', d.is_paused);

    // Progress
    document.getElementById('progressPhase').textContent = d.current_phase || 'Waiting...';
    document.getElementById('progressStep').textContent = d.current_step || '--';
    const pct = d.lots_total > 0 ? Math.round((d.lots_done / d.lots_total) * 100) : 0;
    document.getElementById('progressBar').style.width = pct + '%';
    document.getElementById('progressText').textContent = d.lots_done + ' / ' + d.lots_total;

    // Sliders (skip the one actively being dragged)
    var activeId = document.activeElement ? document.activeElement.id : '';
    if (activeId !== 'sliderShort') setSlider('sliderShort', 'valShort', d.config.delay_short);
    if (activeId !== 'sliderMedium') setSlider('sliderMedium', 'valMedium', d.config.delay_medium);
    if (activeId !== 'sliderLong') setSlider('sliderLong', 'valLong', d.config.delay_long);
    if (activeId !== 'sliderCheckbox') setSlider('sliderCheckbox', 'valCheckbox', d.config.delay_checkbox);

    // LOT table
    const tbody = document.getElementById('lotTableBody');
    tbody.innerHTML = '';
    (d.lot_statuses || []).forEach(lot => {
        const tr = document.createElement('tr');
        if (lot.status === 'running') tr.classList.add('running');

        const pillClass = {
            done: 'pill-done', running: 'pill-running',
            failed: 'pill-failed', skipped: 'pill-skipped', pending: 'pill-pending'
        }[lot.status] || 'pill-pending';

        const isSkipped = skipLots.has(String(lot.lot));

        function td(text, style) {
            const cell = document.createElement('td');
            cell.textContent = text;
            if (style) cell.setAttribute('style', style);
            return cell;
        }

        tr.appendChild(td(lot.lot));
        tr.appendChild(td(lot.count));

        const statusTd = document.createElement('td');
        const pill = document.createElement('span');
        pill.className = 'pill ' + pillClass;
        pill.textContent = lot.status;
        statusTd.appendChild(pill);
        tr.appendChild(statusTd);

        tr.appendChild(td(lot.ref_id || '--', 'font-family:monospace;font-size:12px'));
        tr.appendChild(td(lot.step || '--', 'color:var(--text-secondary);font-size:12px'));

        const actionTd = document.createElement('td');
        if (lot.status === 'pending') {
            const btn = document.createElement('button');
            btn.className = 'skip-btn' + (isSkipped ? ' active' : '');
            btn.textContent = isSkipped ? 'Unskip' : 'Skip';
            btn.addEventListener('click', function() { toggleLot(String(lot.lot)); });
            actionTd.appendChild(btn);
        }
        tr.appendChild(actionTd);

        tbody.appendChild(tr);
    });

    // Log
    const logBody = document.getElementById('logBody');
    const logScroll = document.getElementById('logScroll');
    const wasAtBottom = logScroll.scrollTop + logScroll.clientHeight >= logScroll.scrollHeight - 30;

    logBody.innerHTML = '';
    (d.log_messages || []).forEach(msg => {
        const div = document.createElement('div');
        div.className = 'log-line';
        div.textContent = msg;
        logBody.appendChild(div);
    });

    document.getElementById('logCount').textContent = d.log_messages.length + ' entries';
    if (wasAtBottom) logScroll.scrollTop = logScroll.scrollHeight;
}

function setSlider(sliderId, valId, value) {
    document.getElementById(sliderId).value = value;
    document.getElementById(valId).textContent = parseFloat(value).toFixed(1) + 's';
}

function togglePause() {
    sendControl(isPaused ? 'resume' : 'pause');
}

function sendControl(action, extra) {
    fetch('/api/control', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(Object.assign({action: action}, extra || {}))
    }).catch(function(err) { console.error('Control request failed:', err); });
}

function updateDelay(key, value, valId) {
    document.getElementById(valId).textContent = parseFloat(value).toFixed(1) + 's';
    clearTimeout(debounceTimers[key]);
    debounceTimers[key] = setTimeout(function() {
        const config = {};
        config[key] = parseFloat(value);
        sendControl('update_config', {config: config});
    }, 200);
}

function toggleLot(lot) {
    if (skipLots.has(lot)) skipLots.delete(lot); else skipLots.add(lot);
    sendControl('toggle_lot', {lot: lot});
}

connectSSE();
</script>
</body>
</html>
"""


# ── Standalone test ──

if __name__ == "__main__":
    state = DashboardState()
    control = ControlFlags()
    state.start_time = time.time()
    state.current_phase = "Phase 1"
    state.current_lot = "3"
    state.current_step = "Step 4: Verifying count"
    state.lots_total = 10
    state.lots_done = 2
    state.lot_statuses = [
        {"lot": str(i), "count": 7, "status": "done" if i <= 2 else ("running" if i == 3 else "pending"),
         "ref_id": f"C32046{i}082" if i <= 2 else "", "step": "Verifying count" if i == 3 else ""}
        for i in range(1, 11)
    ]
    state.log_messages.append(f"{datetime.now().strftime('%H:%M:%S')}  Dashboard started in standalone test mode")
    state.log_messages.append(f"{datetime.now().strftime('%H:%M:%S')}  Open http://127.0.0.1:5555 to preview")

    start_dashboard(state, control)
    print("Dashboard running at http://127.0.0.1:5555 (test mode)")
    print("Press Ctrl+C to stop")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopped.")

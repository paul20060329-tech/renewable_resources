"""Renewable Resource Lab - simulated coding demo.

This is a visual demo only. It does not perform real analysis; the final
step opens the configured result URL in the default browser.
"""

from __future__ import annotations

import random
import math
import tkinter as tk
import webbrowser
from tkinter import ttk


TARGET_URL = "https://renewable-resources.onrender.com"


class RenewableDemo:
    BG = "#0a0f14"
    PANEL = "#101820"
    PANEL_2 = "#151f28"
    BORDER = "#263744"
    TEXT = "#d7e2e8"
    MUTED = "#71838e"
    GREEN = "#63e6a4"
    CYAN = "#5ed7e8"
    AMBER = "#ffc857"
    RED = "#ff7181"

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Renewable Resource Lab  |  Simulated Run")
        self.root.geometry("1120x700")
        self.root.minsize(900, 580)
        self.root.configure(bg=self.BG)
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

        self.running = False
        self.sequence_index = 0
        self.activity_frame = 0
        self.log_widget: tk.Text
        self.status_var = tk.StringVar(value="IDLE / READY")
        self.phase_var = tk.StringVar(value="Awaiting simulated pipeline")
        self.progress_var = tk.DoubleVar(value=0)
        self.metric_vars: dict[str, tk.StringVar] = {}

        self._build_style()
        self._build_ui()

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Demo.Horizontal.TProgressbar",
            troughcolor="#202b34",
            background=self.GREEN,
            bordercolor="#202b34",
            lightcolor=self.GREEN,
            darkcolor=self.GREEN,
        )
        style.configure(
            "Demo.Vertical.TScrollbar",
            background=self.PANEL_2,
            troughcolor=self.BG,
            bordercolor=self.BG,
            arrowcolor=self.MUTED,
        )

    def _build_ui(self) -> None:
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        header = tk.Frame(self.root, bg=self.BG, padx=24, pady=18)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        logo = tk.Canvas(header, width=40, height=40, bg=self.BG, highlightthickness=0)
        logo.grid(row=0, column=0, rowspan=2, padx=(0, 12))
        logo.create_oval(3, 3, 37, 37, outline=self.GREEN, width=2)
        logo.create_arc(9, 9, 31, 31, start=35, extent=220, outline=self.CYAN, width=2)
        logo.create_line(20, 8, 20, 32, fill=self.GREEN, width=2)

        tk.Label(
            header,
            text="RENEWABLE RESOURCE LAB",
            fg=self.TEXT,
            bg=self.BG,
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=1, sticky="w")
        tk.Label(
            header,
            text="Virtual Code Demo Console  /  SIMULATION MODE",
            fg=self.MUTED,
            bg=self.BG,
            font=("Consolas", 9),
        ).grid(row=1, column=1, sticky="w", pady=(2, 0))

        self.status_badge = tk.Label(
            header,
            textvariable=self.status_var,
            fg=self.GREEN,
            bg="#12251d",
            font=("Consolas", 9, "bold"),
            padx=12,
            pady=7,
        )
        self.status_badge.grid(row=0, column=2, rowspan=2, sticky="e")

        body = tk.Frame(self.root, bg=self.BG, padx=24, pady=0)
        body.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        body.grid_rowconfigure(1, weight=1)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=0, minsize=285)

        toolbar = tk.Frame(body, bg=self.BG)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        self.start_button = self._button(toolbar, "▶  开始运行", self.start_demo, self.GREEN, "#102a20")
        self.start_button.pack(side="left")
        self.open_button = self._button(toolbar, "Open the target page", self.open_target, self.AMBER, "#2a2411")
        self.open_button.pack(side="right")

        terminal_panel = tk.Frame(body, bg=self.PANEL, highlightbackground=self.BORDER, highlightthickness=1)
        terminal_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        terminal_panel.grid_rowconfigure(1, weight=1)
        terminal_panel.grid_columnconfigure(0, weight=1)

        terminal_bar = tk.Frame(terminal_panel, bg=self.PANEL_2, padx=12, pady=9)
        terminal_bar.grid(row=0, column=0, sticky="ew")
        for color in (self.RED, self.AMBER, self.GREEN):
            tk.Label(terminal_bar, text="●", fg=color, bg=self.PANEL_2, font=("Segoe UI", 10)).pack(side="left", padx=2)
        tk.Label(
            terminal_bar,
            text="  python  renewable_pipeline.py",
            fg=self.MUTED,
            bg=self.PANEL_2,
            font=("Consolas", 9),
        ).pack(side="left", padx=8)

        text_frame = tk.Frame(terminal_panel, bg=self.PANEL)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", style="Demo.Vertical.TScrollbar")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_widget = tk.Text(
            text_frame,
            bg=self.PANEL,
            fg=self.TEXT,
            insertbackground=self.GREEN,
            selectbackground="#29414a",
            relief="flat",
            wrap="word",
            padx=16,
            pady=14,
            font=("Consolas", 10),
            yscrollcommand=scrollbar.set,
            state="disabled",
        )
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.configure(command=self.log_widget.yview)
        self.log_widget.tag_configure("prompt", foreground=self.GREEN)
        self.log_widget.tag_configure("info", foreground=self.CYAN)
        self.log_widget.tag_configure("muted", foreground=self.MUTED)
        self.log_widget.tag_configure("warn", foreground=self.AMBER)
        self.log_widget.tag_configure("success", foreground=self.GREEN)
        self.log_widget.tag_configure("error", foreground=self.RED)

        side = tk.Frame(body, bg=self.PANEL, highlightbackground=self.BORDER, highlightthickness=1, padx=16, pady=16)
        side.grid(row=1, column=1, sticky="nsew")
        tk.Label(side, text="PIPELINE TELEMETRY", fg=self.TEXT, bg=self.PANEL, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(side, text="实时模拟状态", fg=self.MUTED, bg=self.PANEL, font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(3, 18))

        self._metric(side, "CURRENT PHASE", self.phase_var)
        self._metric(side, "DATA SOURCES", self.metric_vars.setdefault("sources", tk.StringVar(value="--")))
        self._metric(side, "RECORDS SCANNED", self.metric_vars.setdefault("records", tk.StringVar(value="--")))
        self._metric(side, "CONFIDENCE INDEX", self.metric_vars.setdefault("confidence", tk.StringVar(value="--")))

        tk.Frame(side, height=1, bg=self.BORDER).pack(fill="x", pady=18)
        tk.Label(side, text="EXECUTION PROGRESS", fg=self.MUTED, bg=self.PANEL, font=("Consolas", 8)).pack(anchor="w")
        ttk.Progressbar(side, style="Demo.Horizontal.TProgressbar", variable=self.progress_var, maximum=100).pack(fill="x", pady=(8, 6))
        self.percent_label = tk.Label(side, text="0%", fg=self.GREEN, bg=self.PANEL, font=("Consolas", 9, "bold"))
        self.percent_label.pack(anchor="e")

        tk.Frame(side, height=1, bg=self.BORDER).pack(fill="x", pady=18)
        tk.Label(side, text="OUTPUT ROUTE", fg=self.MUTED, bg=self.PANEL, font=("Consolas", 8)).pack(anchor="w")
        tk.Label(
            side,
            text=TARGET_URL,
            fg=self.CYAN,
            bg=self.PANEL,
            font=("Consolas", 9),
            wraplength=245,
            justify="left",
        ).pack(anchor="w", pady=(7, 0))

        self._write_line("Renewable Resource Lab :: simulated coding session", "info")
        self._write_line("This is a visual demonstration, and all computation states are simulated outputs.", "muted")
        self._write_line("Click to start the simulation run\n", "muted")

    def _button(self, parent: tk.Widget, text: str, command, fg: str, bg: str) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            fg=fg,
            bg=bg,
            activeforeground="#ffffff",
            activebackground="#1e3c43",
            relief="flat",
            borderwidth=0,
            padx=13,
            pady=8,
            cursor="hand2",
            font=("Microsoft YaHei UI", 9, "bold"),
        )

    def _metric(self, parent: tk.Widget, label: str, variable: tk.StringVar) -> None:
        frame = tk.Frame(parent, bg=self.PANEL)
        frame.pack(fill="x", pady=6)
        tk.Label(frame, text=label, fg=self.MUTED, bg=self.PANEL, font=("Consolas", 8)).pack(anchor="w")
        tk.Label(frame, textvariable=variable, fg=self.TEXT, bg=self.PANEL, font=("Consolas", 11, "bold"), wraplength=245, justify="left").pack(anchor="w", pady=(3, 0))

    def _write_line(self, message: str, tag: str = "") -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", message + "\n", tag)
        self.log_widget.configure(state="disabled")
        self.log_widget.see("end")

    def _set_status(self, text: str, color: str) -> None:
        self.status_var.set(text)
        self.status_badge.configure(fg=color)

    def _animate_activity(self) -> None:
        """Keep a small heartbeat animation visible while the simulation runs."""
        if not self.running:
            return
        frames = ("|", "/", "-", "\\")
        self.status_var.set(f"RUNNING / SIMULATION {frames[self.activity_frame]}")
        self.status_badge.configure(fg=self.GREEN)
        self.activity_frame = (self.activity_frame + 1) % len(frames)
        self.root.after(180, self._animate_activity)

    def start_demo(self) -> None:
        if self.running:
            return
        self.running = True
        self.sequence_index = 0
        self.start_button.configure(state="disabled")
        self._set_status("RUNNING / SIMULATION", self.GREEN)
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")
        self.progress_var.set(0)
        self._update_metrics("Initializing virtual execution environment", "--", "--", "--")
        self._animate_activity()
        self._run_next()

    def _run_next(self) -> None:
        if not self.running:
            return
        steps = [
            ("$ python renewable_pipeline.py", "prompt", 700),
            ("[BOOT] Loading isolated runtime ...", "info", 700),
            ("[BOOT] Resolving 42 virtual analysis modules ...", "info", 900),
            ("[BOOT] Mounting renewable-resource knowledge graph ...", "info", 850),
            ("[OK]   Virtual environment ready  |  kernel=RRL-4.2", "success", 850),
            ("", "muted", 700),
            ("[SCAN] Detecting solar / wind / hydro data connectors ...", "info", 900),
            ("[SCAN] Establishing encrypted telemetry channel ...", "info", 850),
            ("[SCAN] Normalizing 18,492 time-series records ...", "info", 750),
            ("[OK]   7 data channels linked  |  latency=42ms", "success", 850),
            ("", "muted", 850),
            ("[QUERY] Indexing regional generation signatures ...", "warn", 900),
            ("[QUERY] Deep-searching historical scenario database ...", "warn", 700),
            ("[QUERY] Comparing 2030 / 2040 / 2050 reference curves ...", "warn", 800),
            ("[OK]   2,406 candidate patterns ranked", "success", 900),
            ("", "muted", 800),
            ("[THINK] Building multi-source scenario graph ...", "warn", 750),
            ("[THINK] Cross-checking regional generation signals ...", "warn", 850),
            ("[THINK] Running constraint and sensitivity passes ...", "warn", 900),
            ("[OK]   Confidence index stabilized at 96.8%", "success", 700),
            ("[RENDER] Preparing interactive result surface ...", "info", 800),
            ("[RENDER] Sealing visual report and routing metadata ...", "info", 850),
            ("[DONE]  Analysis package sealed and routed to browser", "success", 1100),
        ]
        if self.sequence_index >= len(steps):
            self.finish_demo()
            return
        message, tag, delay = steps[self.sequence_index]
        self.sequence_index += 1
        self._write_line(message, tag)
        progress = min(96, round((self.sequence_index / len(steps)) * 100))
        self.progress_var.set(progress)
        self.percent_label.configure(text=f"{progress}%")
        self._update_live_metrics()

        popup_at = {
            3: ("软件构建中", "正在装配虚拟分析组件", self.CYAN),
            8: ("数据接入中", "正在连接多源可再生能源数据", self.GREEN),
            13: ("深度检索中", "正在深度思考并检索数据库", self.AMBER),
            18: ("模型编排中", "正在合并多源推演结果", self.GREEN),
            22: ("输出准备中", "正在生成交互式成果页面", self.CYAN),
        }
        if self.sequence_index in popup_at:
            title, body, color = popup_at[self.sequence_index]
            self.root.after(120, lambda: self._show_popup(title, body, color))
        self.root.after(delay, self._run_next)

    def _update_live_metrics(self) -> None:
        phase = self.sequence_index
        if phase < 6:
            self._update_metrics("Booting virtual environment", "--", "--", "--")
        elif phase < 11:
            self._update_metrics("Connector scan + normalization", "7 channels", f"{random.randint(12_000, 18_492):,}", "--")
        elif phase < 16:
            self._update_metrics("Deep database retrieval", "7 channels", f"{random.randint(18_492, 22_640):,}", "--")
        elif phase < 21:
            self._update_metrics("Scenario graph inference", "7 channels", "18,492", f"{random.uniform(94.2, 97.1):.1f}%")
        else:
            self._update_metrics("Result surface ready", "7 channels", "18,492", "96.8%")

    def _update_metrics(self, phase: str, sources: str, records: str, confidence: str) -> None:
        self.phase_var.set(phase)
        self.metric_vars["sources"].set(sources)
        self.metric_vars["records"].set(records)
        self.metric_vars["confidence"].set(confidence)

    def _show_popup(self, title: str, body: str, color: str) -> None:
        if not self.running:
            return
        popup = tk.Toplevel(self.root)
        popup.title(f"系统消息 · {title}")
        popup.configure(bg=self.BG)
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("540x300")
        self._center_popup(popup)

        shell = tk.Frame(popup, bg=self.PANEL, highlightbackground=color, highlightthickness=1)
        shell.pack(fill="both", expand=True, padx=1, pady=1)

        header = tk.Frame(shell, bg=self.PANEL_2, padx=18, pady=12)
        header.pack(fill="x")
        tk.Label(header, text="●  SYSTEM EVENT  /  LIVE ORCHESTRATOR", fg=color, bg=self.PANEL_2, font=("Consolas", 9, "bold")).pack(side="left")
        tk.Label(header, text="RRL-4.2", fg=self.MUTED, bg=self.PANEL_2, font=("Consolas", 8)).pack(side="right")

        content = tk.Frame(shell, bg=self.PANEL, padx=18, pady=14)
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        radar = tk.Canvas(content, width=116, height=116, bg=self.PANEL, highlightthickness=0)
        radar.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 18), pady=(3, 0))
        radar.create_oval(7, 7, 109, 109, outline="#29434c", width=1)
        radar.create_oval(22, 22, 94, 94, outline="#29434c", width=1)
        radar.create_line(58, 7, 58, 109, fill="#203842")
        radar.create_line(7, 58, 109, 58, fill="#203842")
        radar.create_oval(53, 53, 63, 63, fill=color, outline="")
        pulse_ring = radar.create_oval(42, 42, 74, 74, outline=color, width=2)
        sweep = radar.create_line(58, 58, 102, 58, fill=color, width=2)
        tk.Label(content, text=title, fg=self.TEXT, bg=self.PANEL, font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=1, sticky="nw")
        tk.Label(content, text=body, fg=self.MUTED, bg=self.PANEL, font=("Microsoft YaHei UI", 9)).grid(row=0, column=1, sticky="sw", pady=(0, 3))

        details = {
            "软件构建中": ("BUILD MATRIX", "COMPONENT ASSEMBLY", ("allocating worker nodes", "resolving module graph", "warming execution cache")),
            "数据接入中": ("INGESTION BUS", "MULTI-SOURCE LINK", ("negotiating data channels", "synchronizing evidence cache", "validating stream signatures")),
            "深度检索中": ("SEARCH ENGINE", "DEEP INDEX QUERY", ("ranking historical patterns", "cross-checking reference curves", "reconciling regional signals")),
            "模型编排中": ("REASONING FABRIC", "SCENARIO INFERENCE", ("assembling scenario graph", "running constraint passes", "stabilizing confidence index")),
            "输出准备中": ("DELIVERY LAYER", "RESULT SURFACE", ("serializing result package", "sealing route metadata", "preparing browser handoff")),
        }
        module, phase, messages = details.get(title, ("SYSTEM MODULE", "ACTIVE PROCESS", ("synchronizing state", "processing signals", "finalizing output")))

        meta = tk.Frame(content, bg=self.PANEL)
        meta.grid(row=1, column=1, sticky="ew", pady=(6, 0))
        meta.grid_columnconfigure(1, weight=1)
        tk.Label(meta, text=f"{module}  /  {phase}", fg=color, bg=self.PANEL, font=("Consolas", 8, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        latency_var = tk.StringVar(value=f"{random.randint(28, 68)} ms")
        signal_var = tk.StringVar(value=f"{random.uniform(96.1, 99.9):.1f}%")
        state_var = tk.StringVar(value=messages[0])
        self._popup_metric(meta, "LATENCY", latency_var, 1, 0)
        self._popup_metric(meta, "SIGNAL", signal_var, 1, 1)
        tk.Label(meta, textvariable=state_var, fg=self.TEXT, bg=self.PANEL, font=("Consolas", 9)).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 6))

        progress_var = tk.DoubleVar(value=0)
        percent_var = tk.StringVar(value="00%")
        progress_row = tk.Frame(shell, bg=self.PANEL, padx=18, pady=0)
        progress_row.pack(fill="x", pady=(0, 14))
        progress_row.grid_columnconfigure(0, weight=1)
        tk.Label(progress_row, text="PROCESSING STREAM", fg=self.MUTED, bg=self.PANEL, font=("Consolas", 8)).grid(row=0, column=0, sticky="w")
        tk.Label(progress_row, textvariable=percent_var, fg=color, bg=self.PANEL, font=("Consolas", 8, "bold")).grid(row=0, column=1, sticky="e")
        bar = ttk.Progressbar(progress_row, style="Demo.Horizontal.TProgressbar", variable=progress_var, maximum=100)
        bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 8))
        signal_strip = tk.Canvas(progress_row, height=22, bg=self.PANEL, highlightthickness=0)
        signal_strip.grid(row=2, column=0, columnspan=2, sticky="ew")
        signal_strip.update_idletasks()
        strip_bars = [signal_strip.create_rectangle(i * 18, 22, i * 18 + 10, 22, fill=color, outline="") for i in range(27)]

        tick = 0

        def animate_popup() -> None:
            nonlocal tick
            try:
                if not popup.winfo_exists():
                    return
            except tk.TclError:
                return
            progress = min(100, round((tick / 23) * 100))
            progress_var.set(progress)
            percent_var.set(f"{progress:02d}%")
            state_var.set(messages[(tick // 4) % len(messages)])
            angle = math.radians((tick * 18) % 360)
            radar.coords(sweep, 58, 58, 58 + 45 * math.cos(angle), 58 + 45 * math.sin(angle))
            radius = 18 + (tick % 8) * 2
            radar.coords(pulse_ring, 58 - radius, 58 - radius, 58 + radius, 58 + radius)
            latency_var.set(f"{random.randint(28, 68)} ms")
            signal_var.set(f"{random.uniform(96.1, 99.9):.1f}%")
            for index, bar_id in enumerate(strip_bars):
                height = random.randint(4, 20) if index < (tick % 27) + 2 else random.randint(2, 7)
                signal_strip.coords(bar_id, index * 18, 22 - height, index * 18 + 10, 22)
            tick += 1
            if tick < 24:
                popup.after(80, animate_popup)

        animate_popup()
        popup.after(1950, self._close_popup, popup)

    def _popup_metric(self, parent: tk.Widget, label: str, variable: tk.StringVar, row: int, column: int) -> None:
        cell = tk.Frame(parent, bg=self.PANEL_2, padx=10, pady=6)
        cell.grid(row=row, column=column, sticky="ew", padx=(0, 6 if column == 0 else 0))
        tk.Label(cell, text=label, fg=self.MUTED, bg=self.PANEL_2, font=("Consolas", 7)).pack(anchor="w")
        tk.Label(cell, textvariable=variable, fg=self.TEXT, bg=self.PANEL_2, font=("Consolas", 10, "bold")).pack(anchor="w", pady=(2, 0))

    @staticmethod
    def _close_popup(popup: tk.Toplevel) -> None:
        try:
            popup.grab_release()
            popup.destroy()
        except tk.TclError:
            pass

    def _center_popup(self, popup: tk.Toplevel) -> None:
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 540) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        popup.geometry(f"540x300+{x}+{y}")

    def finish_demo(self) -> None:
        self.running = False
        self.progress_var.set(100)
        self.percent_label.configure(text="100%")
        self._update_metrics("Result surface ready", "7 channels", "18,492", "96.8%")
        self._set_status("COMPLETE / ROUTED", self.CYAN)
        self._write_line("", "muted")
        self._write_line("[ROUTE] Opening renewable-resources.onrender.com ...", "info")
        self._write_line(f"[URL]   {TARGET_URL}", "success")
        self._write_line("[EXIT]  Demo pipeline complete.\n", "success")
        self.start_button.configure(state="normal")
        self.open_target()

    @staticmethod
    def open_target() -> None:
        webbrowser.open_new_tab(TARGET_URL)

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    RenewableDemo().run()

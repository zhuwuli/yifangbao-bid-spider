from __future__ import annotations

import os
import subprocess
import sys
import threading
import contextlib
from pathlib import Path
from tkinter import BooleanVar, IntVar, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = APP_DIR / "2026-乙方宝招标信息统计.xlsx"
SPIDER = APP_DIR / "yfb_bid_spider.py"


class SpiderApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("乙方宝招标信息爬取工具")
        self.root.geometry("900x620")
        self.root.minsize(820, 560)

        self.cookie = StringVar()
        self.xlsx = StringVar(value=str(DEFAULT_XLSX))
        self.days = IntVar(value=31)
        self.dry_run = BooleanVar(value=False)
        self.process: subprocess.Popen[str] | None = None
        self.running_in_process = False

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(6, weight=1)

        ttk.Label(outer, text="Cookie").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        cookie_entry = ttk.Entry(outer, textvariable=self.cookie, show="*", width=80)
        cookie_entry.grid(row=0, column=1, columnspan=3, sticky="ew", pady=6)

        ttk.Label(outer, text="Excel 文件").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(outer, textvariable=self.xlsx).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(outer, text="选择", command=self.choose_xlsx).grid(row=1, column=2, padx=8, pady=6)
        ttk.Button(outer, text="打开", command=self.open_xlsx).grid(row=1, column=3, pady=6)

        ttk.Label(outer, text="时间范围").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        presets = ttk.Frame(outer)
        presets.grid(row=2, column=1, columnspan=3, sticky="w", pady=6)
        ttk.Radiobutton(presets, text="近 7 天", variable=self.days, value=7).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(presets, text="近 1 个月", variable=self.days, value=31).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(presets, text="近 3 个月", variable=self.days, value=90).pack(side="left", padx=(0, 16))
        ttk.Label(presets, text="自定义天数").pack(side="left")
        ttk.Spinbox(presets, from_=1, to=365, textvariable=self.days, width=7).pack(side="left", padx=(6, 0))

        ttk.Label(outer, text="当前规则").grid(row=3, column=0, sticky="nw", padx=(0, 10), pady=6)
        rule_text = (
            "地区：山东济南、莱芜；关键词：监测、水土保持、测绘、测量；"
            "组合：全文/标题 × 智能/精准；输出：主表摘要 + 公告详情跳转。"
        )
        ttk.Label(outer, text=rule_text, wraplength=720).grid(row=3, column=1, columnspan=3, sticky="w", pady=6)

        options = ttk.Frame(outer)
        options.grid(row=4, column=1, columnspan=3, sticky="w", pady=6)
        ttk.Checkbutton(options, text="演练模式，不写入 Excel", variable=self.dry_run).pack(side="left")

        actions = ttk.Frame(outer)
        actions.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(8, 10))
        self.run_button = ttk.Button(actions, text="开始爬取", command=self.start)
        self.run_button.pack(side="left")
        self.stop_button = ttk.Button(actions, text="停止", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        ttk.Button(actions, text="清空日志", command=self.clear_log).pack(side="left", padx=8)
        ttk.Button(actions, text="打开说明书", command=self.open_manual).pack(side="left", padx=8)

        ttk.Label(outer, text="运行日志").grid(row=6, column=0, sticky="nw", padx=(0, 10))
        self.log = ScrolledText(outer, height=18, wrap="word")
        self.log.grid(row=6, column=1, columnspan=3, sticky="nsew")

        self.status = StringVar(value="就绪")
        ttk.Label(outer, textvariable=self.status).grid(row=7, column=0, columnspan=4, sticky="w", pady=(8, 0))

    def choose_xlsx(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            initialdir=str(APP_DIR),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.xlsx.set(path)

    def open_xlsx(self) -> None:
        path = Path(self.xlsx.get())
        if not path.exists():
            messagebox.showwarning("文件不存在", "请选择有效的 Excel 文件。")
            return
        os.startfile(path)

    def open_manual(self) -> None:
        manual = APP_DIR / "乙方宝爬虫使用说明.md"
        txt_manual = APP_DIR / "乙方宝爬虫使用说明.txt"
        target = manual if manual.exists() else txt_manual
        if not target.exists():
            messagebox.showwarning("未找到说明书", str(manual))
            return
        try:
            os.startfile(target)
        except OSError:
            subprocess.Popen(["notepad.exe", str(target)], cwd=str(APP_DIR))

    def clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def write_log(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def start(self) -> None:
        cookie = self.cookie.get().strip()
        xlsx = Path(self.xlsx.get().strip())
        if not getattr(sys, "frozen", False) and not SPIDER.exists():
            messagebox.showerror("脚本不存在", f"未找到 {SPIDER}")
            return
        if not cookie:
            messagebox.showwarning("缺少 Cookie", "请先粘贴乙方宝 Cookie。")
            return
        if not xlsx.exists():
            messagebox.showwarning("Excel 不存在", "请选择有效的 Excel 文件。")
            return
        if self.process and self.process.poll() is None:
            messagebox.showinfo("正在运行", "爬虫正在运行，请等待结束或点击停止。")
            return

        cmd = [
            sys.executable,
            str(SPIDER),
            "--days",
            str(self.days.get()),
            "--xlsx",
            str(xlsx),
        ]
        if self.dry_run.get():
            cmd.append("--dry-run")

        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status.set("正在运行")
        self.write_log("\n=== 开始爬取 ===\n")
        self.write_log("命令：" + " ".join(cmd) + "\n")

        if getattr(sys, "frozen", False):
            thread = threading.Thread(target=self._run_embedded, args=(cmd, cookie), daemon=True)
        else:
            thread = threading.Thread(target=self._run_process, args=(cmd, cookie), daemon=True)
        thread.start()

    def _run_embedded(self, cmd: list[str], cookie: str) -> None:
        class LogWriter:
            def __init__(self, app: SpiderApp) -> None:
                self.app = app

            def write(self, text: str) -> int:
                if text:
                    self.app.root.after(0, self.app.write_log, text)
                return len(text)

            def flush(self) -> None:
                return None

        old_cookie = os.environ.get("YFB_COOKIE")
        old_argv = sys.argv[:]
        os.environ["YFB_COOKIE"] = cookie
        sys.argv = [str(SPIDER)] + cmd[2:]
        self.running_in_process = True
        try:
            import yfb_bid_spider

            with contextlib.redirect_stdout(LogWriter(self)), contextlib.redirect_stderr(LogWriter(self)):
                code = yfb_bid_spider.main()
            self.root.after(0, self._finish, int(code or 0))
        except Exception as exc:
            self.root.after(0, self.write_log, f"\n运行失败：{exc}\n")
            self.root.after(0, self._finish, 1)
        finally:
            sys.argv = old_argv
            if old_cookie is None:
                os.environ.pop("YFB_COOKIE", None)
            else:
                os.environ["YFB_COOKIE"] = old_cookie
            self.running_in_process = False

    def _run_process(self, cmd: list[str], cookie: str) -> None:
        env = os.environ.copy()
        env["YFB_COOKIE"] = cookie
        try:
            self.process = subprocess.Popen(
                cmd,
                cwd=str(APP_DIR),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.root.after(0, self.write_log, line)
            code = self.process.wait()
            self.root.after(0, self._finish, code)
        except Exception as exc:
            self.root.after(0, self.write_log, f"\n运行失败：{exc}\n")
            self.root.after(0, self._finish, 1)

    def stop(self) -> None:
        if self.running_in_process:
            messagebox.showinfo("正在运行", "打包版正在内部执行任务，无法强制停止；请等待本轮完成。")
            return
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.write_log("\n已发送停止请求。\n")
            self.status.set("正在停止")

    def _finish(self, code: int) -> None:
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.status.set("完成" if code == 0 else f"结束，退出码 {code}")
        self.write_log(f"=== 结束，退出码 {code} ===\n")


def main() -> None:
    root = Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    SpiderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

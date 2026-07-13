from __future__ import annotations

import contextlib
import os
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import BooleanVar, IntVar, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from yfb_browser_auth import BrowserInfo, browser_from_path, detect_browsers, interactive_login, read_saved_cookie


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = APP_DIR / "2026-乙方宝招标信息统计.xlsx"
SPIDER = APP_DIR / "yfb_bid_spider.py"
BROWSER_DATA_DIR = APP_DIR / "browser_data"
AUTO_BROWSER = "自动选择（推荐）"


class SpiderApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("乙方宝招标信息爬取工具")
        self.root.geometry("940x700")
        self.root.minsize(860, 640)

        self.cookie = StringVar()
        self.browser_choice = StringVar(value=AUTO_BROWSER)
        self.auth_status = StringVar(value="尚未读取登录状态")
        self.xlsx = StringVar(value=str(DEFAULT_XLSX))
        self.days = IntVar(value=31)
        self.dry_run = BooleanVar(value=False)
        self.process: subprocess.Popen[str] | None = None
        self.running_in_process = False
        self.browser_options: dict[str, BrowserInfo] = {}

        self._build_ui()
        self.refresh_browsers()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(7, weight=1)

        ttk.Label(outer, text="登录浏览器").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        browser_row = ttk.Frame(outer)
        browser_row.grid(row=0, column=1, columnspan=3, sticky="ew", pady=6)
        browser_row.columnconfigure(0, weight=1)
        self.browser_combo = ttk.Combobox(browser_row, textvariable=self.browser_choice, state="readonly")
        self.browser_combo.grid(row=0, column=0, sticky="ew")
        ttk.Button(browser_row, text="刷新", command=self.refresh_browsers).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(browser_row, text="手动选择", command=self.choose_browser).grid(row=0, column=2, padx=(8, 0))
        self.login_button = ttk.Button(browser_row, text="登录乙方宝", command=self.login_browser)
        self.login_button.grid(row=0, column=3, padx=(8, 0))

        ttk.Label(outer, text="Cookie（备用）").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(outer, textvariable=self.cookie, show="*", width=70).grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=6
        )
        ttk.Label(outer, textvariable=self.auth_status).grid(row=1, column=3, sticky="e", pady=6)

        ttk.Label(outer, text="Excel 文件").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(outer, textvariable=self.xlsx).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(outer, text="选择", command=self.choose_xlsx).grid(row=2, column=2, padx=8, pady=6)
        ttk.Button(outer, text="打开", command=self.open_xlsx).grid(row=2, column=3, pady=6)

        ttk.Label(outer, text="时间范围").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=6)
        presets = ttk.Frame(outer)
        presets.grid(row=3, column=1, columnspan=3, sticky="w", pady=6)
        ttk.Radiobutton(presets, text="近 7 天", variable=self.days, value=7).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(presets, text="近 1 个月", variable=self.days, value=31).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(presets, text="近 3 个月", variable=self.days, value=90).pack(side="left", padx=(0, 16))
        ttk.Label(presets, text="自定义天数").pack(side="left")
        ttk.Spinbox(presets, from_=1, to=365, textvariable=self.days, width=7).pack(side="left", padx=(6, 0))

        ttk.Label(outer, text="当前规则").grid(row=4, column=0, sticky="nw", padx=(0, 10), pady=6)
        rule_text = (
            "地区：山东济南、莱芜；关键词：监测、水土保持、测绘、测量、绿色建筑评价、绿色建筑验收；"
            "输出：原始表 + 筛选后表；标题命中保留，标题未命中再按公告资质规则筛选。"
        )
        ttk.Label(outer, text=rule_text, wraplength=740).grid(row=4, column=1, columnspan=3, sticky="w", pady=6)

        options = ttk.Frame(outer)
        options.grid(row=5, column=1, columnspan=3, sticky="w", pady=6)
        ttk.Checkbutton(options, text="演练模式，不写入 Excel", variable=self.dry_run).pack(side="left")

        actions = ttk.Frame(outer)
        actions.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(8, 10))
        self.run_button = ttk.Button(actions, text="开始爬取", command=self.start)
        self.run_button.pack(side="left")
        self.stop_button = ttk.Button(actions, text="停止", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        ttk.Button(actions, text="清空日志", command=self.clear_log).pack(side="left", padx=8)
        ttk.Button(actions, text="打开说明书", command=self.open_manual).pack(side="left", padx=8)

        ttk.Label(outer, text="运行日志").grid(row=7, column=0, sticky="nw", padx=(0, 10))
        self.log = ScrolledText(outer, height=18, wrap="word")
        self.log.grid(row=7, column=1, columnspan=3, sticky="nsew")

        self.status = StringVar(value="就绪")
        ttk.Label(outer, textvariable=self.status).grid(row=8, column=0, columnspan=4, sticky="w", pady=(8, 0))
        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.grid(row=9, column=0, columnspan=4, sticky="ew", pady=(6, 0))

    def refresh_browsers(self) -> None:
        current = self.browser_choice.get()
        self.browser_options = {}
        values = [AUTO_BROWSER]
        counts: dict[str, int] = {}
        for browser in detect_browsers():
            counts[browser.name] = counts.get(browser.name, 0) + 1
            suffix = f" ({counts[browser.name]})" if counts[browser.name] > 1 else ""
            label = f"{browser.name}{suffix}"
            self.browser_options[label] = browser
            values.append(label)
        self.browser_combo.configure(values=values)
        self.browser_choice.set(current if current in values else AUTO_BROWSER)
        if len(values) == 1:
            self.auth_status.set("未检测到支持的浏览器")

    def choose_browser(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Chromium 内核浏览器程序",
            filetypes=[("浏览器程序", "*.exe"), ("所有文件", "*.*")],
        )
        if not path:
            return
        browser = browser_from_path(path)
        label = f"手动：{browser.name}"
        self.browser_options[label] = browser
        values = list(self.browser_combo.cget("values"))
        if label not in values:
            values.append(label)
            self.browser_combo.configure(values=values)
        self.browser_choice.set(label)
        self.auth_status.set("已选择浏览器，尚未登录")

    def selected_browser(self) -> BrowserInfo | None:
        choice = self.browser_choice.get()
        if choice == AUTO_BROWSER:
            return next(iter(self.browser_options.values()), None)
        return self.browser_options.get(choice)

    def login_browser(self) -> None:
        if self.running_in_process or (self.process and self.process.poll() is None):
            messagebox.showinfo("正在运行", "请等待本轮爬取结束后再登录。")
            return
        browser = self.selected_browser()
        if browser is None:
            messagebox.showwarning("未找到浏览器", "请选择 Edge、Chrome、Brave，或手动选择 Chromium 浏览器程序。")
            return
        self.login_button.config(state="disabled")
        self.status.set("等待浏览器登录")
        self.auth_status.set("等待登录")
        self.progress.start(12)
        self.write_log(f"\n[登录] 正在打开 {browser.name}，请在浏览器中完成登录。\n")
        threading.Thread(target=self._login_worker, args=(browser,), daemon=True).start()

    def _login_worker(self, browser: BrowserInfo) -> None:
        try:
            cookie = interactive_login(browser, BROWSER_DATA_DIR, status_callback=self._login_status)
            self.root.after(0, self._login_finished, cookie, "")
        except Exception as exc:
            self.root.after(0, self._login_finished, "", str(exc))

    def _login_status(self, text: str) -> None:
        self.root.after(0, self.auth_status.set, text)
        self.root.after(0, self.write_log, f"[登录] {text}\n")

    def _login_finished(self, cookie: str, error: str) -> None:
        self.login_button.config(state="normal")
        self.progress.stop()
        if error:
            self.status.set("登录未完成")
            self.auth_status.set("未取得登录状态")
            self.write_log(f"[登录] 失败：{error}\n")
            messagebox.showwarning("登录未完成", error)
            return
        self.cookie.set(cookie)
        self.status.set("登录状态已保存")
        self.auth_status.set("登录状态可用")
        self.write_log("[登录] 登录状态已保存，以后可自动复用。\n")
        messagebox.showinfo("登录成功", "已保存乙方宝登录状态，以后可直接点击“开始爬取”。")

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
        line = text.strip().splitlines()[-1] if text.strip() else ""
        if line.startswith(("[列表]", "[详情]", "[写入]", "已生成", "已追加", "演练模式")):
            self.status.set(line[:160])

    def start(self) -> None:
        xlsx = Path(self.xlsx.get().strip())
        if not getattr(sys, "frozen", False) and not SPIDER.exists():
            messagebox.showerror("脚本不存在", f"未找到 {SPIDER}")
            return
        if not xlsx.exists():
            messagebox.showwarning("Excel 不存在", "请选择有效的 Excel 文件。")
            return
        if self.running_in_process or (self.process and self.process.poll() is None):
            messagebox.showinfo("正在运行", "爬虫正在运行，请等待结束或点击停止。")
            return

        if getattr(sys, "frozen", False):
            cmd = [sys.executable, str(SPIDER), "--days", str(self.days.get()), "--xlsx", str(xlsx)]
        else:
            cmd = [sys.executable, "-u", str(SPIDER), "--days", str(self.days.get()), "--xlsx", str(xlsx)]
        if self.dry_run.get():
            cmd.append("--dry-run")

        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.login_button.config(state="disabled")
        self.progress.start(12)
        cookie = self.cookie.get().strip()
        if cookie:
            self._start_with_cookie(cmd, cookie)
            return

        browser = self.selected_browser()
        if browser is None:
            self._auth_prepare_failed("未检测到支持的浏览器，请手动选择浏览器或粘贴 Cookie。")
            return
        self.status.set("正在读取已保存的登录状态")
        self.write_log(f"\n[登录] 正在从 {browser.name} 的专用登录环境读取状态。\n")
        threading.Thread(target=self._prepare_cookie, args=(browser, cmd), daemon=True).start()

    def _prepare_cookie(self, browser: BrowserInfo, cmd: list[str]) -> None:
        try:
            cookie = read_saved_cookie(browser, BROWSER_DATA_DIR)
            if not cookie:
                raise RuntimeError("没有找到有效登录状态，请先点击“登录乙方宝”。")
            self.root.after(0, self._cookie_ready, cmd, cookie)
        except Exception as exc:
            self.root.after(0, self._auth_prepare_failed, str(exc))

    def _cookie_ready(self, cmd: list[str], cookie: str) -> None:
        self.cookie.set(cookie)
        self.auth_status.set("已自动读取登录状态")
        self.write_log("[登录] 已自动读取保存的登录状态。\n")
        self._start_with_cookie(cmd, cookie)

    def _auth_prepare_failed(self, error: str) -> None:
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.login_button.config(state="normal")
        self.progress.stop()
        self.status.set("需要登录")
        self.auth_status.set("登录状态不可用")
        self.write_log(f"[登录] {error}\n")
        messagebox.showwarning("需要登录", error)

    def _start_with_cookie(self, cmd: list[str], cookie: str) -> None:
        self.status.set("正在运行")
        self.write_log("\n=== 开始爬取 ===\n")
        self.write_log("命令：" + " ".join(cmd) + "\n")
        if getattr(sys, "frozen", False):
            target = self._run_embedded
        else:
            target = self._run_process
        threading.Thread(target=target, args=(cmd, cookie), daemon=True).start()

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
        self.login_button.config(state="normal")
        self.progress.stop()
        if code == 2:
            self.cookie.set("")
            self.auth_status.set("登录可能已失效，请重新登录")
            messagebox.showwarning("登录已失效", "乙方宝登录状态可能已失效，请点击登录乙方宝重新登录。")
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

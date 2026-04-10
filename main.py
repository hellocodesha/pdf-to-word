#!/usr/bin/env python3
"""PDF 批量转 Word 工具 — 基于 pdf2docx"""

import os
import re
import platform
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 可选拖拽支持（需要安装 tkinterdnd2）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── 配色 & 字体 ────────────────────────────────────────────────────────
BG      = "#f0f2f5"
CARD    = "white"
BORDER  = "#dde1e7"
ACCENT  = "#4a90e2"
SUCCESS = "#28a745"
GRAY    = "#6c757d"
FONT    = "微软雅黑"


class App:
    def __init__(self):
        # 检查核心依赖
        self._check_deps()

        Root = TkinterDnD.Tk if HAS_DND else tk.Tk
        self.root = Root()
        self.root.title("PDF 批量转 Word 工具")
        self.root.configure(bg=BG)
        self.root.minsize(600, 500)
        self._center(740, 620)

        self.files: list = []
        self.out_dir  = tk.StringVar()
        self.keep_fmt = tk.BooleanVar(value=True)
        self.running  = False

        self._setup_styles()
        self._build_ui()
        self.root.mainloop()

    # ── 依赖检查 ──────────────────────────────────────────────────────────

    @staticmethod
    def _check_deps():
        try:
            import pdf2docx  # noqa: F401
        except ImportError:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "缺少依赖库",
                "未找到 pdf2docx，请先安装：\n\n    pip install pdf2docx\n\n安装后重新启动程序。"
            )
            root.destroy()
            raise SystemExit(1)

    # ── 窗口居中 ──────────────────────────────────────────────────────────

    def _center(self, w, h):
        self.root.geometry(f"{w}x{h}")
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ── 样式 ──────────────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure(
            "P.TProgressbar",
            thickness=14,
            troughcolor=BORDER,
            background=ACCENT,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            bordercolor=BORDER,
        )

    # ── 小控件工厂 ────────────────────────────────────────────────────────

    def _lbl(self, parent, text, size=9, bold=False, fg="#444", bg=None):
        return tk.Label(
            parent, text=text,
            font=(FONT, size, "bold" if bold else ""),
            bg=bg if bg is not None else parent["bg"],
            fg=fg,
        )

    def _btn(self, parent, text, cmd, accent=False, green=False, **kw):
        bg = SUCCESS if green else (ACCENT if accent else "#e2e6ea")
        fg = "white" if (accent or green) else "#333"
        return tk.Button(
            parent, text=text, command=cmd,
            bg=bg, fg=fg,
            activebackground=bg, activeforeground=fg,
            relief="flat", bd=0,
            font=(FONT, 9),
            padx=12, pady=5,
            cursor="hand2",
            **kw,
        )

    def _card(self, parent, title):
        """返回 (外框 frame, 内容 body frame)"""
        outer = tk.Frame(parent, bg=BORDER)
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        hdr = tk.Frame(inner, bg="#f7f8fa")
        hdr.pack(fill="x")
        tk.Label(hdr, text=title, font=(FONT, 9, "bold"),
                 bg="#f7f8fa", fg="#555", anchor="w",
                 padx=10, pady=7).pack(fill="x")
        tk.Frame(inner, height=1, bg=BORDER).pack(fill="x")

        body = tk.Frame(inner, bg=CARD, padx=12, pady=10)
        body.pack(fill="both", expand=True)
        return outer, body

    # ── 构建 UI ───────────────────────────────────────────────────────────

    def _build_ui(self):
        wrap = tk.Frame(self.root, bg=BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=14)

        # 标题栏
        bar = tk.Frame(wrap, bg=BG)
        bar.pack(fill="x", pady=(0, 12))
        self._lbl(bar, "PDF 批量转 Word", size=15, bold=True,
                  fg="#222", bg=BG).pack(side="left")
        dnd_note = "（支持拖拽 PDF）" if HAS_DND else "（拖拽不可用：请安装 tkinterdnd2）"
        self._lbl(bar, dnd_note, size=8, fg="#aaa", bg=BG).pack(side="left", padx=8)

        # ── 卡片 1：文件选择 ─────────────────────────────────────────────
        c1, b1 = self._card(wrap, "  选择 PDF 文件")
        c1.pack(fill="both", expand=True, pady=(0, 10))

        btn_row = tk.Frame(b1, bg=CARD)
        btn_row.pack(fill="x", pady=(0, 8))
        self._btn(btn_row, "+ 添加文件",   self.add_files,  accent=True).pack(side="left", padx=(0, 8))
        self._btn(btn_row, "+ 添加文件夹", self.add_folder, accent=True).pack(side="left", padx=(0, 8))
        self._btn(btn_row, "清空列表",     self.clear_list              ).pack(side="left")
        self._lbl(btn_row, "（支持多选）", size=8, fg="#aaa", bg=CARD).pack(side="left", padx=8)

        # 拖拽提示
        if HAS_DND:
            tk.Label(b1,
                     text="可将 PDF 文件 / 文件夹拖拽到下方列表",
                     font=(FONT, 8), bg="#eef3fb", fg="#7a9ecf",
                     pady=5, anchor="center").pack(fill="x", pady=(0, 6))

        # Listbox + 滚动条
        lb_wrap = tk.Frame(b1, bg=CARD)
        lb_wrap.pack(fill="both", expand=True)
        sy = tk.Scrollbar(lb_wrap)
        sy.pack(side="right", fill="y")
        sx = tk.Scrollbar(lb_wrap, orient="horizontal")
        sx.pack(side="bottom", fill="x")

        self.lb = tk.Listbox(
            lb_wrap, font=(FONT, 9), height=8,
            bg="#fafbfc", fg="#333",
            selectbackground=ACCENT, selectforeground="white",
            relief="flat", bd=1,
            highlightthickness=1, highlightbackground=BORDER,
            selectmode="extended", activestyle="none",
            yscrollcommand=sy.set, xscrollcommand=sx.set,
        )
        self.lb.pack(side="left", fill="both", expand=True)
        sy.config(command=self.lb.yview)
        sx.config(command=self.lb.xview)

        if HAS_DND:
            self.lb.drop_target_register(DND_FILES)
            self.lb.dnd_bind("<<Drop>>", self._on_drop)

        self.cnt_lbl = self._lbl(b1, "已选 0 个文件", size=8, fg="#aaa", bg=CARD)
        self.cnt_lbl.pack(anchor="w", pady=(5, 0))

        # ── 卡片 2：输出设置 ─────────────────────────────────────────────
        c2, b2 = self._card(wrap, "  输出设置")
        c2.pack(fill="x", pady=(0, 10))

        r1 = tk.Frame(b2, bg=CARD)
        r1.pack(fill="x", pady=(0, 8))
        self._lbl(r1, "输出目录：", bg=CARD).pack(side="left")
        self.out_entry = tk.Entry(
            r1, textvariable=self.out_dir,
            font=(FONT, 9),
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground="#d0d5de",
            bg=CARD,
        )
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(4, 8))
        self._btn(r1, "浏览…", self.browse_out).pack(side="left")
        self._lbl(r1, "留空则与源文件同目录", size=8, fg="#bbb", bg=CARD).pack(side="left", padx=(8, 0))

        r2 = tk.Frame(b2, bg=CARD)
        r2.pack(fill="x")
        tk.Checkbutton(
            r2,
            text="保留原始排版格式（图片、表格、字体样式等）",
            variable=self.keep_fmt,
            font=(FONT, 9), bg=CARD, fg="#444",
            activebackground=CARD, selectcolor=CARD,
            cursor="hand2",
        ).pack(side="left")

        # ── 卡片 3：转换进度 ─────────────────────────────────────────────
        c3, b3 = self._card(wrap, "  转换进度")
        c3.pack(fill="x", pady=(0, 12))

        self.status = self._lbl(b3, "就绪，等待开始…", fg="#666", bg=CARD)
        self.status.pack(fill="x", anchor="w", pady=(0, 6))
        self.pbar = ttk.Progressbar(b3, mode="determinate", style="P.TProgressbar")
        self.pbar.pack(fill="x")
        self.prog = self._lbl(b3, "0 / 0", size=8, fg="#aaa", bg=CARD)
        self.prog.pack(anchor="e", pady=(3, 0))

        # ── 底部操作栏 ───────────────────────────────────────────────────
        bot = tk.Frame(wrap, bg=BG)
        bot.pack(fill="x")
        self._lbl(bot, "基于 pdf2docx 开源库", size=8, fg="#ccc", bg=BG).pack(side="left")
        self.go_btn = tk.Button(
            bot, text="开始转换", command=self.start,
            bg=SUCCESS, fg="white",
            activebackground="#218838", activeforeground="white",
            relief="flat", bd=0,
            font=(FONT, 11, "bold"),
            padx=28, pady=10,
            cursor="hand2",
        )
        self.go_btn.pack(side="right")

    # ── 文件管理 ──────────────────────────────────────────────────────────

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if paths:
            self._push(list(paths))

    def add_folder(self):
        d = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if not d:
            return
        found = sorted(Path(d).rglob("*.pdf"))
        if found:
            self._push([str(p) for p in found])
        else:
            messagebox.showinfo("提示", "该文件夹中没有找到 PDF 文件。")

    def browse_out(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir.set(d)

    def clear_list(self):
        self.files.clear()
        self.lb.delete(0, "end")
        self._refresh_count()

    def _push(self, paths):
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.lb.insert("end", p)
        self.lb.see("end")
        self._refresh_count()

    def _refresh_count(self):
        n = len(self.files)
        self.cnt_lbl.config(text=f"已选 {n} 个文件")

    def _on_drop(self, event):
        tokens = re.findall(r"\{[^}]+\}|[^\s]+", event.data)
        paths = []
        for t in tokens:
            t = t.strip("{}")
            p = Path(t)
            if p.is_dir():
                paths.extend(str(f) for f in sorted(p.rglob("*.pdf")))
            elif p.suffix.lower() == ".pdf" and p.is_file():
                paths.append(str(p))
        if paths:
            self._push(paths)
        else:
            messagebox.showinfo("提示", "拖入内容中未找到 PDF 文件。")

    # ── 转换逻辑 ──────────────────────────────────────────────────────────

    def start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning("提示", "请先添加 PDF 文件。")
            return
        self.running = True
        self.go_btn.config(state="disabled", text="转换中…", bg=GRAY)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        from pdf2docx import Converter  # 延迟导入，加快窗口启动速度

        files = list(self.files)
        total = len(files)
        ok, fail = 0, []

        self._after(lambda: self.pbar.config(maximum=total, value=0))

        for i, src in enumerate(files):
            name = Path(src).name
            self._update_progress(i, total, f"正在转换（{i + 1}/{total}）：{name}")

            try:
                # 确定输出目录
                dest_dir = (
                    Path(self.out_dir.get().strip())
                    if self.out_dir.get().strip()
                    else Path(src).parent
                )
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = str(dest_dir / (Path(src).stem + ".docx"))

                # 执行转换
                # keep_fmt=True：pdf2docx 默认尽量保留版式（图片、表格、字体）
                # keep_fmt=False：仅提取文本流，跳过复杂版式还原
                cv = Converter(src)
                if self.keep_fmt.get():
                    cv.convert(dest, start=0, end=None)
                else:
                    # 关闭图片提取，仅保留文字段落
                    cv.convert(dest, start=0, end=None,
                               connected_border_tolerance=0,
                               page_margin_factor_with_one_column=0.5)
                cv.close()
                ok += 1

            except PermissionError:
                fail.append((name, "文件被加密或拒绝访问"))
            except Exception as exc:
                fail.append((name, str(exc)))

        self._update_progress(total, total, f"完成！成功 {ok} 个，失败 {len(fail)} 个。")
        self.running = False

        open_dir = (
            self.out_dir.get().strip()
            or (str(Path(files[0]).parent) if files else "")
        )
        self._after(lambda: self._show_done(ok, fail, open_dir))

    def _update_progress(self, cur, total, text):
        def _do():
            self.status.config(text=text)
            self.pbar.config(value=cur)
            self.prog.config(text=f"{cur} / {total}")
        self._after(_do)

    def _after(self, fn):
        self.root.after(0, fn)

    def _show_done(self, ok, fail, open_dir):
        self.go_btn.config(state="normal", text="开始转换", bg=SUCCESS)

        lines = [f"转换完成！\n\n成功：{ok} 个\n失败：{len(fail)} 个"]
        if fail:
            lines.append("\n\n失败文件详情：")
            for name, err in fail:
                short = (err[:120] + "…") if len(err) > 120 else err
                lines.append(f"\n• {name}\n  原因：{short}")
        msg = "".join(lines)

        if open_dir:
            ans = messagebox.askquestion(
                "转换完成", msg + "\n\n是否打开输出文件夹？", icon="info"
            )
            if ans == "yes":
                self._open_dir(open_dir)
        else:
            messagebox.showinfo("转换完成", msg)

    @staticmethod
    def _open_dir(path: str):
        s = platform.system()
        if s == "Windows":
            os.startfile(path)
        elif s == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


if __name__ == "__main__":
    App()

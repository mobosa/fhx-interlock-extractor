#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
FHX Interlock Extractor GUI
DeltaV FHX文件联锁信息提取工具 - Apple风格图形界面版
"""

import os
import threading
import customtkinter as ctk

from core import parse_fhx, generate_excel

# Apple-inspired theme setup
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Color palette
COLORS = {
    "bg": "#FFFFFF",
    "bg_secondary": "#F5F5F7",
    "text_primary": "#1D1D1F",
    "text_secondary": "#86868B",
    "text_tertiary": "#AEAEB2",
    "accent": "#007AFF",
    "accent_hover": "#0063D1",
    "success": "#34C759",
    "success_bg": "#E8F9ED",
    "error": "#FF3B30",
    "border": "#E5E5EA",
    "card_bg": "#F5F5F7",
    "input_border": "#D2D2D7",
}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("FHX Interlock Extractor")
        self.geometry("640x520")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg"])

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exp_logo.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        self.fhx_path = ctk.StringVar()
        self.output_path = ctk.StringVar()

        self._build_ui()

    def _build_ui(self):
        # ── Main container ──
        container = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        container.pack(fill="both", expand=True, padx=0, pady=0)

        # ── Header ──
        header = ctk.CTkFrame(container, fg_color=COLORS["bg"], corner_radius=0, height=80)
        header.pack(fill="x", padx=0, pady=(16, 0))
        header.pack_propagate(False)

        ctk.CTkLabel(header, text="FHX Interlock Extractor",
                     font=("Segoe UI", 24, "bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", padx=28, pady=(0, 2))
        ctk.CTkLabel(header, text="DeltaV FHX 文件联锁信息提取工具",
                     font=("Segoe UI", 13),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", padx=28)

        # ── Separator line ──
        sep = ctk.CTkFrame(container, fg_color=COLORS["border"], height=1)
        sep.pack(fill="x", padx=28, pady=(8, 0))

        # ── Content area ──
        content = ctk.CTkFrame(container, fg_color=COLORS["bg"], corner_radius=0)
        content.pack(fill="both", expand=True, padx=28, pady=(12, 10))

        # ── File selection card ──
        file_card = ctk.CTkFrame(content, fg_color=COLORS["card_bg"], corner_radius=12)
        file_card.pack(fill="x", pady=(0, 16))

        file_inner = ctk.CTkFrame(file_card, fg_color="transparent")
        file_inner.pack(fill="x", padx=14, pady=12)

        ctk.CTkLabel(file_inner, text="FHX 文件",
                     font=("Segoe UI", 13, "bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", pady=(0, 8))

        input_row = ctk.CTkFrame(file_inner, fg_color="transparent")
        input_row.pack(fill="x")

        self.file_entry = ctk.CTkEntry(input_row,
                                       textvariable=self.fhx_path,
                                       placeholder_text="选择要解析的 FHX 文件...",
                                       font=("Segoe UI", 12),
                                       height=36,
                                       corner_radius=8,
                                       border_width=1,
                                       border_color=COLORS["input_border"],
                                       fg_color="#FFFFFF")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_btn = ctk.CTkButton(input_row, text="浏览",
                                        command=self._browse_fhx,
                                        font=("Segoe UI", 12),
                                        width=72, height=36,
                                        corner_radius=8,
                                        fg_color=COLORS["accent"],
                                        hover_color=COLORS["accent_hover"])
        self.browse_btn.pack(side="right")

        # ── Output info ──
        self.output_var = ctk.StringVar(value="输出位置: 与 FHX 文件同目录")
        ctk.CTkLabel(content, textvariable=self.output_var,
                     font=("Segoe UI", 11),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 16))

        # ── Progress card ──
        self.progress_frame = ctk.CTkFrame(content, fg_color=COLORS["card_bg"], corner_radius=12)
        self.progress_frame.pack(fill="x", pady=(0, 16))

        progress_inner = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        progress_inner.pack(fill="x", padx=14, pady=10)

        self.progress_label = ctk.CTkLabel(progress_inner, text="准备就绪",
                                           font=("Segoe UI", 12),
                                           text_color=COLORS["text_secondary"])
        self.progress_label.pack(anchor="w", pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(progress_inner,
                                               height=6, corner_radius=3,
                                               progress_color=COLORS["accent"],
                                               fg_color=COLORS["border"])
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        # ── Status ──
        self.status_var = ctk.StringVar(value="就绪 - 请选择 FHX 文件")
        ctk.CTkLabel(content, textvariable=self.status_var,
                     font=("Segoe UI", 11),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 16))

        # ── Run button ──
        self.run_btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.run_btn_frame.pack(fill="x", pady=(0, 16))

        self.run_btn = ctk.CTkButton(self.run_btn_frame, text="开始提取",
                                     command=self._run,
                                     font=("Segoe UI", 15, "bold"),
                                     height=48, corner_radius=12,
                                     fg_color=COLORS["accent"],
                                     hover_color=COLORS["accent_hover"])
        self.run_btn.pack(fill="x")

        # ── Author footer ──
        ctk.CTkLabel(container, text="Author: Jared.Ji@emerson.com",
                     font=("Segoe UI", 10),
                     text_color=COLORS["text_tertiary"]).pack(pady=(0, 12))

    def _browse_fhx(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择 FHX 文件",
            filetypes=[("FHX files", "*.fhx"), ("All files", "*.*")]
        )
        if path:
            self.fhx_path.set(path)
            base = os.path.dirname(path)
            out = os.path.join(base, "联锁信息表.xlsx")
            self.output_path.set(out)
            self.output_var.set(f"输出: {out}")

    def _on_progress(self, current, total):
        pct = int(current / total * 100) if total > 0 else 0
        self.after(0, self._update_progress, pct)

    def _update_progress(self, pct):
        self.progress_bar.set(pct / 100)
        self.progress_label.configure(text=f"正在解析 FHX 文件... {pct}%")

    def _run(self):
        from tkinter import messagebox

        fhx = self.fhx_path.get().strip()
        if not fhx:
            messagebox.showwarning("提示", "请选择 FHX 文件")
            return
        if not os.path.exists(fhx):
            messagebox.showerror("错误", f"文件不存在:\n{fhx}")
            return
        out = os.path.join(os.path.dirname(fhx), "联锁信息表.xlsx")
        self.output_path.set(out)

        self.run_btn.configure(state="disabled", text="处理中...")
        self.progress_label.configure(text="正在解析 FHX 文件...")
        self.progress_bar.set(0)
        self.status_var.set("正在解析 FHX 文件...")

        def worker():
            try:
                instances = parse_fhx(fhx, progress_callback=self._on_progress)
                total_ilock = sum(inst.interlock.count_real() for inst in instances)
                total_perm = sum(inst.permissive.count_real() for inst in instances)
                total_force = sum(inst.force.count_real() for inst in instances)
                total_at = sum(inst.tracking.count_real() for inst in instances)
                self.after(0, lambda: self.progress_label.configure(text="正在生成 Excel 文件..."))
                self.after(0, lambda: self.progress_bar.set(1.0))
                generate_excel(instances, out)
                self.after(0, lambda: self.progress_label.configure(text="完成"))
                self.after(0, lambda: self.status_var.set(f"完成 - {len(instances)}个模块, {total_ilock + total_perm + total_force + total_at}个条件"))
                self.after(0, lambda: messagebox.showinfo("完成", f"联锁信息表已生成!\n\n{out}"))
            except Exception as e:
                self.after(0, lambda: self.status_var.set("出错"))
                self.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                self.after(0, self._on_complete)

        def _on_complete():
            self.run_btn.configure(state="normal", text="开始提取")
            self.progress_bar.set(0)
            self.progress_label.configure(text="准备就绪")

        self._on_complete_ref = _on_complete
        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()

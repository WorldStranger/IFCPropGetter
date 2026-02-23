# -*- coding: utf-8 -*-

"""GUI 界面模块"""

import os
import sys
import queue
import threading
from pathlib import Path
from tkinter import messagebox, filedialog, BooleanVar, StringVar, ttk

import customtkinter as ctk

from ifc_prop_getter import extractor, utils
from ifc_prop_getter.constants import DEFAULT_PROPERTIES


def get_resource_path(relative_path):
    """获取资源文件绝对路径"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


class IFCPropertyExtractorApp:
    def __init__(self):
        # 全局字体配置
        self.font_main = ("微软雅黑", 12)
        self.font_bold = ("微软雅黑", 12, "bold")
        self.font_tree_content = ("微软雅黑", 11)
        self.font_log = ("Consolas", 11)

        self.colors = {
            "bg": "#F5F5F5",
            "fg": "#2E3B4E",
            "primary": "#6C8B9A",
            "secondary": "#A7C7D9",
            "frame_bg": "#FFFFFF",
            "entry_bg": "#FFFFFF",
            "log_bg": "#F0F0F0",
            "button_hover": "#5A7A8A",
            "tree_bg": "#FFFFFF",
            "tree_header_bg": "#E1E1E1",
            "tree_selected": "#D9EAF0"
        }

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("IFCPropGetter (Pro)")
        self.root.geometry("820x720")
        self.root.minsize(600, 450)
        self.root.configure(fg_color=self.colors["bg"])

        icon_path = get_resource_path(os.path.join("resources", "IFCPropGetter.ico"))
        if os.path.exists(icon_path):
            self.root.iconbitmap(default=icon_path)

        self.ifc_path = StringVar()
        self.file_size = StringVar(value="未选择文件")
        self.properties = DEFAULT_PROPERTIES.copy()

        # 属性选项
        self.include_globalid = BooleanVar(value=True)
        self.include_name = BooleanVar(value=False)

        self.output_filename = StringVar(value="output_data")
        self.output_dir = StringVar(value=utils.get_default_output_dir())
        self.output_format = StringVar(value="Excel")

        self.queue = queue.Queue()
        self.worker_thread = None
        self.progress_window = None
        self.running = False
        self.stop_event = threading.Event()

        # 进度文本
        self.status_text = StringVar(value="准备就绪")

        # 动画控制
        self.marquee_val = 0.0
        self.marquee_after_id = None

        self._create_widgets()
        self.root.after(100, self._check_queue)

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self.root, fg_color=self.colors["bg"], corner_radius=0)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # --- 文件选择区 ---
        file_frame = ctk.CTkFrame(main_frame, fg_color=self.colors["frame_bg"], corner_radius=12)
        file_frame.pack(fill="x", pady=(0, 8))
        inner_file = ctk.CTkFrame(file_frame, fg_color="transparent")
        inner_file.pack(fill="x", padx=12, pady=10)

        row1 = ctk.CTkFrame(inner_file, fg_color="transparent")
        row1.pack(fill="x", pady=3)
        ctk.CTkLabel(row1, text="IFC文件:", text_color=self.colors["fg"], font=self.font_main).pack(side="left")
        self.ifc_entry = ctk.CTkEntry(row1, textvariable=self.ifc_path, state="readonly", height=30,
                                      font=self.font_main)
        self.ifc_entry.pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(row1, text="浏览", command=self._browse_ifc, width=70, height=30,
                      fg_color=self.colors["primary"], font=self.font_main).pack(side="right")

        row2 = ctk.CTkFrame(inner_file, fg_color="transparent")
        row2.pack(fill="x", pady=3)
        ctk.CTkLabel(row2, textvariable=self.file_size, font=("微软雅黑", 11), text_color=self.colors["fg"]).pack(
            side="left")

        # --- 属性管理区 ---
        prop_frame = ctk.CTkFrame(main_frame, fg_color=self.colors["frame_bg"], corner_radius=12)
        prop_frame.pack(fill="x", pady=(0, 8))
        inner_prop = ctk.CTkFrame(prop_frame, fg_color="transparent")
        inner_prop.pack(fill="both", expand=True, padx=12, pady=10)

        add_row = ctk.CTkFrame(inner_prop, fg_color="transparent")
        add_row.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(add_row, text="属性名称:", text_color=self.colors["fg"], font=self.font_main).pack(side="left")
        self.prop_entry = ctk.CTkEntry(add_row, width=280, height=30, placeholder_text="例如: Assembly/Cast unit Mark",
                                       font=self.font_main)
        self.prop_entry.pack(side="left", padx=8)
        ctk.CTkButton(add_row, text="添加", command=self._add_property, width=60, height=30,
                      fg_color=self.colors["primary"], font=self.font_main).pack(side="left")

        # Treeview 列表
        tree_container = ctk.CTkFrame(inner_prop, fg_color="transparent", border_width=1, border_color="#ccc")
        tree_container.pack(fill="both", expand=True, pady=6)

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Treeview.Heading",
                        background=self.colors["tree_header_bg"],
                        foreground=self.colors["fg"],
                        font=self.font_bold,
                        relief="raised")

        style.configure("Treeview",
                        background=self.colors["tree_bg"],
                        foreground=self.colors["fg"],
                        rowheight=28,
                        fieldbackground=self.colors["tree_bg"],
                        font=self.font_tree_content)

        style.map("Treeview",
                  background=[("selected", self.colors["tree_selected"])],
                  foreground=[("selected", self.colors["fg"])])

        self.tree = ttk.Treeview(tree_container, columns=('order', 'property'), show='headings', height=5,
                                 selectmode='browse')
        self.tree.heading('order', text='序号')
        self.tree.column('order', width=50, anchor='center')
        self.tree.heading('property', text='属性名称')
        self.tree.column('property', width=350, anchor='center')

        vsb = ttk.Scrollbar(tree_container, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 按钮组
        btn_row = ctk.CTkFrame(inner_prop, fg_color="transparent")
        btn_row.pack(fill="x", pady=3)
        btn_conf = {"width": 70, "height": 28, "fg_color": self.colors["primary"], "font": self.font_main}
        ctk.CTkButton(btn_row, text="上移", command=self._move_up, **btn_conf).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="下移", command=self._move_down, **btn_conf).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="删除选中", command=self._delete_selected, **btn_conf).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="清空列表", command=self._confirm_clear, **btn_conf).pack(side="left", padx=2)

        check_row = ctk.CTkFrame(inner_prop, fg_color="transparent")
        check_row.pack(fill="x", pady=3)
        ctk.CTkCheckBox(check_row, text="包含 GlobalId", variable=self.include_globalid,
                        fg_color=self.colors["primary"], font=self.font_main).pack(side="left", padx=5)
        ctk.CTkCheckBox(check_row, text="包含 Name", variable=self.include_name,
                        fg_color=self.colors["primary"], font=self.font_main).pack(side="left", padx=5)

        self._refresh_tree()

        # --- 导出选项区 ---
        opt_frame = ctk.CTkFrame(main_frame, fg_color=self.colors["frame_bg"], corner_radius=12)
        opt_frame.pack(fill="x", pady=(0, 8))
        inner_opt = ctk.CTkFrame(opt_frame, fg_color="transparent")
        inner_opt.pack(fill="x", padx=12, pady=10)

        name_row = ctk.CTkFrame(inner_opt, fg_color="transparent")
        name_row.pack(fill="x", pady=3)
        ctk.CTkLabel(name_row, text="输出文件名:", text_color=self.colors["fg"], font=self.font_main).pack(side="left")
        self.name_entry = ctk.CTkEntry(name_row, textvariable=self.output_filename, width=240, height=30,
                                       font=self.font_main)
        self.name_entry.pack(side="left", padx=8)

        dir_row = ctk.CTkFrame(inner_opt, fg_color="transparent")
        dir_row.pack(fill="x", pady=3)
        ctk.CTkLabel(dir_row, text="输出文件夹:", text_color=self.colors["fg"], font=self.font_main).pack(side="left")
        self.dir_entry = ctk.CTkEntry(dir_row, textvariable=self.output_dir, state="readonly", height=30,
                                      font=self.font_main)
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(dir_row, text="浏览", command=self._browse_output_dir, width=60, height=30,
                      fg_color=self.colors["primary"], font=self.font_main).pack(side="right")

        format_row = ctk.CTkFrame(inner_opt, fg_color="transparent")
        format_row.pack(fill="x", pady=3)
        ctk.CTkLabel(format_row, text="输出格式:", text_color=self.colors["fg"], font=self.font_main).pack(side="left")
        ctk.CTkRadioButton(format_row, text="Excel (.xlsx)", variable=self.output_format, value="Excel",
                           fg_color=self.colors["primary"], font=self.font_main).pack(side="left", padx=10)
        ctk.CTkRadioButton(format_row, text="CSV (.csv)", variable=self.output_format, value="CSV",
                           fg_color=self.colors["primary"], font=self.font_main).pack(side="left")

        # --- 日志区 ---
        log_frame = ctk.CTkFrame(main_frame, fg_color=self.colors["frame_bg"], corner_radius=12)
        log_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.log_text = ctk.CTkTextbox(log_frame, height=60, font=self.font_log,
                                       fg_color=self.colors["log_bg"],
                                       border_color="#D0D0D0", border_width=1)
        self.log_text.pack(fill="both", expand=True, padx=12, pady=10)
        self.log_text.configure(state="disabled")

        # --- 控制按钮 ---
        ctrl_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        ctrl_frame.pack(fill="x")
        self.start_btn = ctk.CTkButton(ctrl_frame, text="开始提取", command=self._start_extraction,
                                       width=120, height=38, fg_color=self.colors["primary"], font=("微软雅黑", 14, "bold"))
        self.start_btn.pack(side="left", padx=5)
        self.exit_btn = ctk.CTkButton(ctrl_frame, text="退出", command=self._quit,
                                      width=80, height=38, fg_color=self.colors["secondary"],
                                      text_color=self.colors["fg"], font=("微软雅黑", 14))
        self.exit_btn.pack(side="left", padx=5)

    # ---------------- 功能函数 ----------------

    def _refresh_tree(self):
        """刷新属性列表视图"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, prop in enumerate(self.properties, start=1):
            self.tree.insert('', 'end', values=(idx, prop))

    def _add_property(self):
        """添加新属性到列表"""
        prop = self.prop_entry.get().strip()
        if prop:
            if prop not in self.properties:
                self.properties.append(prop)
                self._refresh_tree()
                self._log(f"添加属性: {prop}")
            else:
                messagebox.showwarning("提示", "属性已存在")
            self.prop_entry.delete(0, 'end')

    def _move_up(self):
        """上移选中的属性"""
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            if idx > 0:
                self.properties[idx], self.properties[idx - 1] = self.properties[idx - 1], self.properties[idx]
                self._refresh_tree()
                self.tree.selection_set(self.tree.get_children()[idx - 1])

    def _move_down(self):
        """下移选中的属性"""
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            if idx < len(self.properties) - 1:
                self.properties[idx], self.properties[idx + 1] = self.properties[idx + 1], self.properties[idx]
                self._refresh_tree()
                self.tree.selection_set(self.tree.get_children()[idx + 1])

    def _delete_selected(self):
        """删除选中的属性"""
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            self.properties.pop(idx)
            self._refresh_tree()

    def _confirm_clear(self):
        """清空属性列表"""
        if messagebox.askyesno("确认", "清空所有属性？"):
            self.properties.clear()
            self._refresh_tree()

    def _browse_ifc(self):
        """浏览并选择 IFC 文件"""
        path = filedialog.askopenfilename(filetypes=[("IFC files", "*.ifc"), ("All", "*.*")])
        if path:
            self.ifc_path.set(path)
            stem = Path(path).stem
            if len(stem) > 8:
                short_stem = stem[-8:]
            else:
                short_stem = stem
            self.output_filename.set(f"{short_stem}_data")
            size = Path(path).stat().st_size
            self.file_size.set(f"文件大小: {size / 1024 / 1024:.2f} MB")

    def _browse_output_dir(self):
        """浏览并选择输出目录"""
        path = filedialog.askdirectory()
        if path: self.output_dir.set(path)

    def _log(self, msg):
        """在日志区追加信息"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{utils.format_timestamp()}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _quit(self):
        """退出程序"""
        if self.running and not messagebox.askyesno("警告", "任务运行中，确定退出？"):
            return
        self.root.quit()

    # ---------------- 核心任务控制 ----------------

    def _start_extraction(self):
        """启动属性提取后台任务"""
        if not self.ifc_path.get() or not self.properties:
            messagebox.showerror("错误", "请检查文件路径和属性列表")
            return

        self.running = True
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")

        self._show_progress_window()
        self._log("开始后台提取任务...")

        self.worker_thread = threading.Thread(
            target=extractor.extract_properties,
            args=(
                self.ifc_path.get(),
                self.properties.copy(),
                self.include_globalid.get(),
                self.include_name.get(),
                self.output_dir.get(),
                self.output_filename.get(),
                self.output_format.get(),
                self.queue,
                self.stop_event
            ),
            daemon=True
        )
        self.worker_thread.start()

    def _show_progress_window(self):
        """显示提取进度弹窗"""
        self.progress_window = ctk.CTkToplevel(self.root)
        self.progress_window.title("处理中")
        self.progress_window.geometry("400x120")
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()

        self.progress_window.protocol("WM_DELETE_WINDOW", self._cancel_task)

        x = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 120) // 2
        self.progress_window.geometry(f"+{x}+{y}")

        self.status_label = ctk.CTkLabel(self.progress_window, textvariable=self.status_text, font=("微软雅黑", 12))
        self.status_label.pack(pady=(20, 10))

        self.progress_bar = ctk.CTkProgressBar(self.progress_window, width=320, mode="determinate")
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)

        self.marquee_val = 0.0
        self._animate_marquee()

        ctk.CTkButton(self.progress_window, text="取消", command=self._cancel_task,
                      fg_color="#d9534f", hover_color="#c9302c", height=28, width=80).pack(pady=10)

    def _animate_marquee(self):
        """进度条单向循环滚动动画"""
        if self.progress_window and self.progress_window.winfo_exists():
            self.marquee_val += 0.015
            if self.marquee_val > 1.0:
                self.marquee_val = 0.0

            self.progress_bar.set(self.marquee_val)
            self.marquee_after_id = self.root.after(30, self._animate_marquee)

    def _cancel_task(self):
        """取消当前后台任务"""
        self.stop_event.set()
        self.status_text.set("正在停止...")

    def _check_queue(self):
        """检查线程通信队列并更新 UI"""
        try:
            if not self.root.winfo_exists():
                return

            while True:
                msg = self.queue.get_nowait()
                mtype = msg.get('type')

                if mtype == 'log':
                    self._log(msg['message'])

                elif mtype == 'status':
                    self.status_text.set(msg['message'])

                elif mtype == 'complete':
                    self._log(f"任务完成: {msg['filepath']}")
                    self._cleanup_task()
                    messagebox.showinfo("成功", f"导出完成！\n路径: {msg['filepath']}")

                elif mtype == 'error':
                    self._log(f"错误: {msg['message']}")
                    self._cleanup_task()
                    messagebox.showerror("错误", msg['message'])

                elif mtype == 'finished':
                    self._cleanup_task()

        except queue.Empty:
            pass
        finally:
            if self.root.winfo_exists():
                self.root.after(50, self._check_queue)

    def _cleanup_task(self):
        """清理并重置任务状态"""
        self.running = False
        self.start_btn.configure(state="normal")

        if self.marquee_after_id:
            try:
                self.root.after_cancel(self.marquee_after_id)
            except Exception:
                pass
            self.marquee_after_id = None

        if self.progress_window:
            try:
                self.progress_window.destroy()
            except Exception:
                pass
            self.progress_window = None
        self.status_text.set("准备就绪")

    def run(self):
        """启动应用主循环"""
        self.root.mainloop()
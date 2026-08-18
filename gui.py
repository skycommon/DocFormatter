# -*- coding: utf-8 -*-
"""
文档快速排版 —— Windows 桌面 GUI（tkinter，纯离线）
功能：导入 TXT/Markdown 杂乱文本 -> 一键智能整理（标题/段落/列表/重点加粗）-> 导出规范 Word(.docx)
支持：简体中文 / English 实时切换；导出文档全部内容纯黑色。
"""

import os
import json
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import formatter
import i18n
from formatter import STYLE_PRESETS, CIRCLED
from i18n import t

# 单一版本常量（避免各处版本号四分五裂）
APP_VERSION = "1.0.3"

HERE = os.path.dirname(os.path.abspath(__file__))

# 打包成单文件 exe 后，__file__ 会指向临时解压目录；用户设置必须落在 exe 真实所在
# 目录才能跨会话持久化，因此用 sys.executable 所在目录作为可写应用目录。
import sys
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = HERE

ICON_PATH = os.path.join(HERE, "app_icon.ico")
# 常用导出位置的本地持久化文件（与本机设置一起自动保存/读取），放在 exe 同目录以持久化
LOC_PATH = os.path.join(APP_DIR, "export_locations.json")
# 通用设置（西文字体等）的本地持久化文件，放在 exe 同目录以跨会话持久化
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")

# 主题配色（key 与语言无关）
THEMES = {
    "light": {
        "bg": "#f5f6f8", "fg": "#222222", "panel": "#ffffff",
        "accent": "#2d7ff9", "muted": "#7a7f87", "entry_bg": "#ffffff",
        "list_bg": "#ffffff", "sel": "#cfe2ff",
    },
    "dark": {
        "bg": "#1e1f22", "fg": "#e6e6e6", "panel": "#2a2c30",
        "accent": "#4a9bff", "muted": "#9aa0a8", "entry_bg": "#2a2c30",
        "list_bg": "#2a2c30", "sel": "#3a4a66",
    },
}

# 样式预设：内部 key -> i18n key
PRESET_KEYS = ["默认", "毕业论文", "通用论文", "课程论文", "简洁风", "阅读风"]
PRESET_I18N = {
    "默认": "preset_default", "毕业论文": "preset_graduation",
    "通用论文": "preset_general", "课程论文": "preset_course",
    "简洁风": "preset_simple", "阅读风": "preset_read",
}

# 字体族选项（覆盖常见中文（宋/黑/楷/仿宋等）与西文（Times/Arial 等）场景）
FONT_OPTIONS = [
    "宋体", "仿宋", "楷体", "黑体", "微软雅黑", "等线",
    "华文宋体", "华文楷体", "华文黑体", "华文中宋",
    "思源宋体", "思源黑体", "隶书", "幼圆",
    "Times New Roman", "Arial", "Calibri", "Cambria",
    "Georgia", "Consolas", "Courier New", "Verdana",
    "Tahoma", "Helvetica",
]

# 西文（英文/数字/罗马字符）字体选项：国标默认 Times New Roman，给常见学术西文字体备选
WESTERN_FONT_OPTIONS = [
    "Times New Roman", "Arial", "Cambria", "Calibri", "Georgia",
    "Consolas", "Courier New", "Verdana", "Tahoma", "Helvetica",
]

# 字号选项：（显示标签，磅值）；含中文「字号」名与常用磅值，覆盖全面
SIZE_OPTIONS = [
    ("初号 42pt", 42.0), ("小初 36pt", 36.0), ("一号 26pt", 26.0),
    ("小一 24pt", 24.0), ("二号 22pt", 22.0), ("小二 18pt", 18.0),
    ("三号 16pt", 16.0), ("小三 15pt", 15.0), ("四号 14pt", 14.0),
    ("小四 12pt", 12.0), ("11pt", 11.0), ("10.5pt", 10.5),
    ("10pt", 10.0), ("小五 9pt", 9.0), ("8pt", 8.0),
    ("七号 5.5pt", 5.5), ("六号 7.5pt", 7.5),
]
SIZE_LABEL_TO_PT = {lab: pt for lab, pt in SIZE_OPTIONS}


def size_label(pt: float) -> str:
    """把磅值换算成下拉标签；找不到时回退为 'Xpt'。"""
    for lab, p in SIZE_OPTIONS:
        if abs(p - pt) < 0.01:
            return lab
    return f"{pt}pt"


def size_pt(label: str) -> float:
    """把下拉标签换算成磅值；解析失败回退 11。"""
    return SIZE_LABEL_TO_PT.get(label, 11.0)


# 页脚模式（内部值 -> i18n 前缀 "footer_"）
FOOTER_MODES = [
    "none", "page", "page_of_total", "date", "date_page", "text", "text_page_split",
]

# 页码对齐方式（与 footer_mode 解耦，单独成一项）
FOOTER_ALIGNS = ["left", "center", "right"]

# 参考文献编号样式（内部值 -> i18n key）
REF_STYLES = ["gb7714", "numbered", "paren", "circle", "superscript", "none"]
REF_STYLE_I18N = {
    "gb7714": "ref_style_gb",
    "numbered": "ref_style_numbered",
    "paren": "ref_style_paren",
    "circle": "ref_style_circle",
    "superscript": "ref_style_sup",
    "none": "ref_style_none",
}


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.lang = "zh"
        self.theme_name = "light"
        self.preset_key = "默认"
        self.sample_lang = "zh"
        self.blocks = []
        self._tr = []  # (widget, i18n_key) 需要随语言刷新的文本控件
        self.export_locations = self._load_locations()  # 常用导出位置列表
        # 西文字体（英文/数字），国标默认 Times New Roman；随设置文件持久化、可随配置导出
        self.latin_font_var = tk.StringVar(value=self._load_western_font())
        self.latin_font_var.trace_add("write", lambda *a: self._save_western_font())

        try:
            if os.path.exists(ICON_PATH):
                self.root.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self._build_ui()
        self.apply_theme()
        self.retranslate()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # 顶部工具栏
        top = ttk.Frame(self.root)
        top.pack(side="top", fill="x", padx=8, pady=6)
        self.w_import = ttk.Button(top, text="", command=self.import_file)
        self.w_import.pack(side="left", padx=4)
        self.w_clear = ttk.Button(top, text="", command=self.clear_text)
        self.w_clear.pack(side="left", padx=4)
        self.w_refresh = ttk.Button(top, text="", command=self.refresh_outline)
        self.w_refresh.pack(side="left", padx=4)
        self.w_export = ttk.Button(top, text="", command=self.export_docx)
        self.w_export.pack(side="left", padx=4)
        self.w_export_pdf = ttk.Button(top, text="", command=self.export_pdf)
        self.w_export_pdf.pack(side="left", padx=4)
        self.w_batch = ttk.Button(top, text="", command=self.open_batch)
        self.w_batch.pack(side="left", padx=4)
        self.w_about = ttk.Button(top, text="", command=self.show_about)
        self.w_about.pack(side="right", padx=4)
        for w,  key in ((self.w_import, "btn_import"), (self.w_clear, "btn_clear"),
                       (self.w_refresh, "btn_refresh"), (self.w_export, "btn_export"),
                       (self.w_export_pdf, "btn_export_pdf"), (self.w_batch, "btn_batch"),
                       (self.w_about, "btn_about")):
            self._tr.append((w, key))

        # 主体：左输入 / 右选项+大纲
        body = ttk.PanedWindow(self.root, orient="horizontal")
        body.pack(side="top", fill="both", expand=True, padx=8, pady=4)

        # 左：输入
        left = ttk.Frame(body)
        self.w_lbl_input = ttk.Label(left, text="")
        self.w_lbl_input.pack(anchor="w", padx=4, pady=2)
        self._tr.append((self.w_lbl_input, "lbl_input"))
        self.text = tk.Text(left, wrap="word", undo=True, font=("Microsoft YaHei", 11))
        self.text.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        sb = ttk.Scrollbar(left, command=self.text.yview)
        sb.pack(side="right", fill="y")
        self.text.config(yscrollcommand=sb.set)
        self.text.bind("<KeyRelease>", lambda e: self._schedule_update())
        body.add(left, weight=3)

        # 右：选项 + 大纲
        right = ttk.Frame(body)

        opt = ttk.LabelFrame(right, text="")
        opt.pack(fill="x", padx=6, pady=6)
        self._tr.append((opt, "frame_options"))
        self.smart_var = tk.BooleanVar(value=True)
        self.toc_var = tk.BooleanVar(value=False)
        self.math_var = tk.BooleanVar(value=True)
        self.cite_var = tk.BooleanVar(value=True)
        # 自定义开关：选中显示对号✓，未选显示○，不用原生勾选框（避免出现叉号）
        self.toggles = []
        self.toggles.append(self._make_toggle(opt, self.smart_var, "chk_smart"))
        self.toggles.append(self._make_toggle(opt, self.toc_var, "chk_toc"))
        self.toggles.append(self._make_toggle(opt, self.math_var, "chk_math"))
        self.toggles.append(self._make_toggle(opt, self.cite_var, "chk_cite"))

        # 样式预设行
        row_preset = ttk.Frame(opt)
        row_preset.pack(anchor="w", padx=6, pady=3, fill="x")
        self.w_lbl_preset = ttk.Label(row_preset, text="")
        self.w_lbl_preset.pack(side="left")
        self._tr.append((self.w_lbl_preset, "lbl_preset"))
        self.preset_disp = tk.StringVar()
        self.cb_preset = ttk.Combobox(row_preset, textvariable=self.preset_disp,
                                      values=[], state="readonly", width=12)
        self.cb_preset.pack(side="left", padx=4)
        self.cb_preset.bind("<<ComboboxSelected>>", self.on_preset)
        self.w_btn_export2 = ttk.Button(row_preset, text="", command=self.export_docx)
        self.w_btn_export2.pack(side="right", padx=2)
        self._tr.append((self.w_btn_export2, "btn_export"))

        # 语言行
        row_lang = ttk.Frame(opt)
        row_lang.pack(anchor="w", padx=6, pady=3, fill="x")
        self.w_lbl_lang = ttk.Label(row_lang, text="")
        self.w_lbl_lang.pack(side="left")
        self._tr.append((self.w_lbl_lang, "lbl_lang"))
        self.lang_disp = tk.StringVar()
        self.cb_lang = ttk.Combobox(row_lang, textvariable=self.lang_disp,
                                    values=[], state="readonly", width=10)
        self.cb_lang.pack(side="left", padx=4)
        self.cb_lang.bind("<<ComboboxSelected>>", self.on_lang)

        # 字体与字号（独立选择，覆盖预设的字体与正文字号）
        fontf = ttk.LabelFrame(opt, text="")
        fontf.pack(fill="x", padx=6, pady=(6, 2))
        self._tr.append((fontf, "frame_font"))
        _preset0 = STYLE_PRESETS.get(self.preset_key, STYLE_PRESETS["默认"])
        self.body_font_var = tk.StringVar(value=_preset0["body_font"])
        self.head_font_var = tk.StringVar(value=_preset0["head_font"])
        self.title_font_var = tk.StringVar(value=_preset0["title_font"])
        self.body_size_var = tk.StringVar(value=size_label(_preset0["body_size"]))

        # 正文字体
        row_bf = ttk.Frame(fontf)
        row_bf.pack(anchor="w", padx=6, pady=2, fill="x")
        self.w_lbl_body_font = ttk.Label(row_bf, text="")
        self.w_lbl_body_font.pack(side="left")
        self._tr.append((self.w_lbl_body_font, "lbl_body_font"))
        self.cb_body_font = ttk.Combobox(row_bf, textvariable=self.body_font_var,
                                         values=FONT_OPTIONS, state="readonly", width=14)
        self.cb_body_font.pack(side="left", padx=4)
        self.cb_body_font.bind("<<ComboboxSelected>>", self.on_font_change)

        # 标题字体
        row_hf = ttk.Frame(fontf)
        row_hf.pack(anchor="w", padx=6, pady=2, fill="x")
        self.w_lbl_head_font = ttk.Label(row_hf, text="")
        self.w_lbl_head_font.pack(side="left")
        self._tr.append((self.w_lbl_head_font, "lbl_head_font"))
        self.cb_head_font = ttk.Combobox(row_hf, textvariable=self.head_font_var,
                                         values=FONT_OPTIONS, state="readonly", width=14)
        self.cb_head_font.pack(side="left", padx=4)
        self.cb_head_font.bind("<<ComboboxSelected>>", self.on_font_change)

        # 主标题字体
        row_tf = ttk.Frame(fontf)
        row_tf.pack(anchor="w", padx=6, pady=2, fill="x")
        self.w_lbl_title_font = ttk.Label(row_tf, text="")
        self.w_lbl_title_font.pack(side="left")
        self._tr.append((self.w_lbl_title_font, "lbl_title_font"))
        self.cb_title_font = ttk.Combobox(row_tf, textvariable=self.title_font_var,
                                          values=FONT_OPTIONS, state="readonly", width=14)
        self.cb_title_font.pack(side="left", padx=4)
        self.cb_title_font.bind("<<ComboboxSelected>>", self.on_font_change)

        # 正文字号
        row_bs = ttk.Frame(fontf)
        row_bs.pack(anchor="w", padx=6, pady=(2, 4), fill="x")
        self.w_lbl_body_size = ttk.Label(row_bs, text="")
        self.w_lbl_body_size.pack(side="left")
        self._tr.append((self.w_lbl_body_size, "lbl_body_size"))
        self.cb_body_size = ttk.Combobox(row_bs, textvariable=self.body_size_var,
                                         values=[lab for lab, _ in SIZE_OPTIONS],
                                         state="readonly", width=14)
        self.cb_body_size.pack(side="left", padx=4)
        self.cb_body_size.bind("<<ComboboxSelected>>", self.on_font_change)

        # 西文字体（英文/数字/罗马字符；国标默认 Times New Roman）
        row_wf = ttk.Frame(fontf)
        row_wf.pack(anchor="w", padx=6, pady=(2, 4), fill="x")
        self.w_lbl_western_font = ttk.Label(row_wf, text="")
        self.w_lbl_western_font.pack(side="left")
        self._tr.append((self.w_lbl_western_font, "lbl_western_font"))
        self.cb_western_font = ttk.Combobox(row_wf, textvariable=self.latin_font_var,
                                            values=WESTERN_FONT_OPTIONS,
                                            state="readonly", width=14)
        self.cb_western_font.pack(side="left", padx=4)
        self.cb_western_font.bind("<<ComboboxSelected>>", self.on_font_change)

        # 配置导入/导出行
        row_cfg = ttk.Frame(opt)
        row_cfg.pack(anchor="w", padx=6, pady=(6, 2), fill="x")
        self.w_export_cfg = ttk.Button(row_cfg, text="", command=self.export_config)
        self.w_export_cfg.pack(side="left", padx=4)
        self.w_import_cfg = ttk.Button(row_cfg, text="", command=self.import_config)
        self.w_import_cfg.pack(side="left", padx=4)
        self._tr.append((self.w_export_cfg, "btn_export_cfg"))
        self._tr.append((self.w_import_cfg, "btn_import_cfg"))

        # 页面设置（页眉 / 页脚）
        page = ttk.LabelFrame(right, text="")
        page.pack(fill="x", padx=6, pady=(6, 2))
        self._tr.append((page, "frame_page"))
        # 页眉
        row_header = ttk.Frame(page)
        row_header.pack(anchor="w", padx=6, pady=3, fill="x")
        self.w_lbl_header = ttk.Label(row_header, text="")
        self.w_lbl_header.pack(side="left")
        self._tr.append((self.w_lbl_header, "lbl_header"))
        self.header_var = tk.StringVar()
        self.header_entry = ttk.Entry(row_header, textvariable=self.header_var, width=22)
        self.header_entry.pack(side="left", padx=4, fill="x", expand=True)
        # 页脚
        row_footer = ttk.Frame(page)
        row_footer.pack(anchor="w", padx=6, pady=3, fill="x")
        self.w_lbl_footer = ttk.Label(row_footer, text="")
        self.w_lbl_footer.pack(side="left")
        self._tr.append((self.w_lbl_footer, "lbl_footer"))
        self.footer_mode_var = tk.StringVar(value="none")
        self.footer_modes = FOOTER_MODES
        self.footer_text_var = tk.StringVar(value="")
        self.footer_disp = tk.StringVar()
        self.cb_footer = ttk.Combobox(row_footer, textvariable=self.footer_disp,
                                      values=[], state="readonly", width=16)
        self.cb_footer.pack(side="left", padx=4)
        self.cb_footer.bind("<<ComboboxSelected>>", self.on_footer)
        # 页码对齐（单独一项：居左 / 居中 / 居右）
        self.footer_align_var = tk.StringVar(value="center")
        self.footer_align_disp = tk.StringVar()
        self.w_lbl_footer_align = ttk.Label(row_footer, text="")
        self.w_lbl_footer_align.pack(side="left", padx=(10, 2))
        self._tr.append((self.w_lbl_footer_align, "lbl_footer_align"))
        self.cb_footer_align = ttk.Combobox(row_footer, textvariable=self.footer_align_disp,
                                            values=[], state="readonly", width=8)
        self.cb_footer_align.pack(side="left", padx=2)
        self.cb_footer_align.bind("<<ComboboxSelected>>", self.on_footer_align)
        # 页脚自定义文字（供 text / text_page_split 使用）
        row_footer_text = ttk.Frame(page)
        row_footer_text.pack(anchor="w", padx=6, pady=3, fill="x")
        self.w_lbl_footer_text = ttk.Label(row_footer_text, text="")
        self.w_lbl_footer_text.pack(side="left")
        self._tr.append((self.w_lbl_footer_text, "lbl_footer_text"))
        self.footer_text_entry = ttk.Entry(row_footer_text, textvariable=self.footer_text_var, width=22)
        self.footer_text_entry.pack(side="left", padx=4, fill="x", expand=True)

        # 参考文献自动整理
        ref = ttk.LabelFrame(right, text="")
        ref.pack(fill="x", padx=6, pady=(6, 2))
        self._tr.append((ref, "frame_ref"))
        self.ref_var = tk.BooleanVar(value=True)
        self.toggles.append(self._make_toggle(ref, self.ref_var, "chk_ref"))

        # 编号样式
        row_ref_style = ttk.Frame(ref)
        row_ref_style.pack(anchor="w", padx=6, pady=2, fill="x")
        self.w_lbl_ref_style = ttk.Label(row_ref_style, text="")
        self.w_lbl_ref_style.pack(side="left")
        self._tr.append((self.w_lbl_ref_style, "lbl_ref_style"))
        self.ref_style_var = tk.StringVar(value="gb7714")
        self.ref_styles = REF_STYLES
        self.ref_style_i18n = REF_STYLE_I18N
        self.ref_style_disp = tk.StringVar()
        self.cb_ref_style = ttk.Combobox(row_ref_style, textvariable=self.ref_style_disp,
                                         values=[], state="readonly", width=18)
        self.cb_ref_style.pack(side="left", padx=4)
        self.cb_ref_style.bind("<<ComboboxSelected>>", self.on_ref_style)

        # 悬挂缩进 / 行距
        row_ref_params = ttk.Frame(ref)
        row_ref_params.pack(anchor="w", padx=6, pady=(1, 4), fill="x")
        self.w_lbl_ref_hang = ttk.Label(row_ref_params, text="")
        self.w_lbl_ref_hang.pack(side="left")
        self._tr.append((self.w_lbl_ref_hang, "lbl_ref_hang"))
        self.ref_hang_var = tk.DoubleVar(value=0.74)
        self.sp_ref_hang = ttk.Spinbox(row_ref_params, textvariable=self.ref_hang_var,
                                       from_=0.0, to=2.0, increment=0.1, width=6)
        self.sp_ref_hang.pack(side="left", padx=4)
        self.w_lbl_ref_line = ttk.Label(row_ref_params, text="")
        self.w_lbl_ref_line.pack(side="left", padx=(10, 0))
        self._tr.append((self.w_lbl_ref_line, "lbl_ref_line"))
        self.ref_line_var = tk.DoubleVar(value=1.5)
        self.sp_ref_line = ttk.Spinbox(row_ref_params, textvariable=self.ref_line_var,
                                       from_=1.0, to=2.5, increment=0.05, width=6)
        self.sp_ref_line.pack(side="left", padx=4)

        # 右侧：大纲 / 预览（选项卡）
        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb

        ol = ttk.Frame(nb)
        self.outline = tk.Listbox(ol, font=("Microsoft YaHei", 10))
        self.outline.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        osb = ttk.Scrollbar(ol, command=self.outline.yview)
        osb.pack(side="right", fill="y")
        self.outline.config(yscrollcommand=osb.set)
        nb.add(ol, text="")

        pv = ttk.Frame(nb)
        self.preview = tk.Text(pv, wrap="word", font=("Microsoft YaHei", 10),
                               state="disabled", undo=False)
        self.preview.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        psb = ttk.Scrollbar(pv, command=self.preview.yview)
        psb.pack(side="right", fill="y")
        self.preview.config(yscrollcommand=psb.set)
        nb.add(pv, text="")

        # 导出位置（常用导出位置，可勾选）
        loc = ttk.Frame(nb)
        self._build_export_loc(loc)
        nb.add(loc, text="")

        body.add(right, weight=2)

        # 底部状态栏
        self.status = ttk.Label(self.root, text="", relief="sunken", anchor="w")
        self.status.pack(side="bottom", fill="x")

        # 初始示例文本 + 首次解析（填充大纲与预览）
        self.text.insert("1.0", t("sample", self.lang))
        self.refresh_outline()

    # ----------------------------------------------------- 自定义开关
    def _make_toggle(self, parent, var, i18n_key):
        """构建一个对号开关：选中=✓，未选=○，点击切换。"""
        btn = tk.Button(parent, relief="flat", anchor="w", padx=6, pady=3,
                        font=("Microsoft YaHei", 10), cursor="hand2",
                        borderwidth=0, highlightthickness=0,
                        command=lambda: self._toggle_flip(var))
        btn.pack(anchor="w", fill="x", padx=2, pady=1)
        entry = {"btn": btn, "var": var, "key": i18n_key}
        self._refresh_one_toggle(entry)
        return entry

    def _toggle_flip(self, var):
        var.set(not var.get())
        for e in self.toggles:
            if e["var"] is var:
                self._refresh_one_toggle(e)
                break
        self.refresh_outline()
        if getattr(self, "_batch_win", None) and self._batch_win.winfo_exists():
            self._batch_update_summary()

    def _refresh_one_toggle(self, entry):
        th = THEMES[self.theme_name]
        label = t(entry["key"], self.lang)
        if entry["var"].get():
            entry["btn"].config(text=f"✓  {label}", fg=th["accent"],
                                bg=th["bg"], activebackground=th["bg"])
        else:
            entry["btn"].config(text=f"○  {label}", fg=th["muted"],
                                bg=th["bg"], activebackground=th["bg"])

    def _refresh_toggles(self):
        for e in self.toggles:
            self._refresh_one_toggle(e)

    # -------------------------------------------------------- 语言/主题
    def retranslate(self):
        """按当前语言刷新所有静态文案与下拉选项。"""
        self.root.title(t("app_title", self.lang))
        for w, key in self._tr:
            w.config(text=t(key, self.lang))

        # 语言下拉
        self.cb_lang.config(values=[t("lang_zh", self.lang), t("lang_en", self.lang)])
        self.lang_disp.set(t("lang_" + self.lang, self.lang))

        # 样式预设下拉
        self.cb_preset.config(values=[t(PRESET_I18N[k], self.lang) for k in PRESET_KEYS])
        self.preset_disp.set(t(PRESET_I18N[self.preset_key], self.lang))

        # 页脚下拉
        footer_vals = [t("footer_" + m, self.lang) for m in self.footer_modes]
        self.cb_footer.config(values=footer_vals)
        idx = self.footer_modes.index(self.footer_mode_var.get())
        self.footer_disp.set(footer_vals[idx])
        # 页码对齐下拉
        align_vals = [t("align_" + a, self.lang) for a in FOOTER_ALIGNS]
        self.cb_footer_align.config(values=align_vals)
        aidx = FOOTER_ALIGNS.index(self.footer_align_var.get())
        self.footer_align_disp.set(align_vals[aidx])

        # 参考文献编号样式下拉
        ref_vals = [t(self.ref_style_i18n[m], self.lang) for m in self.ref_styles]
        self.cb_ref_style.config(values=ref_vals)
        idx = self.ref_styles.index(self.ref_style_var.get())
        self.ref_style_disp.set(ref_vals[idx])

        # 选项卡标签
        if hasattr(self, "nb"):
            self.nb.tab(0, text=t("tab_outline", self.lang))
            self.nb.tab(1, text=t("tab_preview", self.lang))
            self.nb.tab(2, text=t("tab_export_loc", self.lang))

        self._refresh_toggles()
        if hasattr(self, "loc_inner"):
            self._refresh_loc_list()
        self.set_status(t("status_ready", self.lang))

    def on_lang(self, event=None):
        new_lang = "zh" if self.lang_disp.get() == t("lang_zh", self.lang) else "en"
        if new_lang == self.lang:
            return
        old = self.lang
        self.lang = new_lang
        # 示例文本：若用户未改动（仍是旧语言示例或为空），则换为新语言示例
        cur = self.text.get("1.0", "end-1c")
        if cur.strip() == "" or cur == t("sample", old):
            self.text.delete("1.0", "end")
            self.text.insert("1.0", t("sample", self.lang))
            self.sample_lang = self.lang
        self.retranslate()
        self.refresh_outline()
        self.set_status(t("status_lang", self.lang).format(lang=t("lang_" + self.lang, self.lang)))

    def on_preset(self, event=None):
        val = self.preset_disp.get()
        for k in PRESET_KEYS:
            if t(PRESET_I18N[k], self.lang) == val:
                self.preset_key = k
                break
        # 预设作为「快速套用」：同步刷新字体/字号选择器（用户仍可单独覆盖）
        p = STYLE_PRESETS.get(self.preset_key, STYLE_PRESETS["默认"])
        self.body_font_var.set(p["body_font"])
        self.head_font_var.set(p["head_font"])
        self.title_font_var.set(p["title_font"])
        self.body_size_var.set(size_label(p["body_size"]))
        # 预设改变了字体/字号，立即刷新预览（所见即所得）
        self.render_preview()
        if getattr(self, "_batch_win", None) and self._batch_win.winfo_exists():
            self._batch_update_summary()

    def on_font_change(self, event=None):
        """字体 / 字号下拉改变时，立即把新字体套用到右侧预览。"""
        self.render_preview()

    def on_footer(self, event=None):
        vals = [t("footer_" + m, self.lang) for m in self.footer_modes]
        sel = self.footer_disp.get()
        if sel in vals:
            self.footer_mode_var.set(self.footer_modes[vals.index(sel)])

    def on_footer_align(self, event=None):
        vals = [t("align_" + a, self.lang) for a in FOOTER_ALIGNS]
        sel = self.footer_align_disp.get()
        if sel in vals:
            self.footer_align_var.set(FOOTER_ALIGNS[vals.index(sel)])

    def on_ref_style(self, event=None):
        vals = [t(self.ref_style_i18n[m], self.lang) for m in self.ref_styles]
        sel = self.ref_style_disp.get()
        if sel in vals:
            self.ref_style_var.set(self.ref_styles[vals.index(sel)])

    # ------------------------------------------------------------ 主题
    def apply_theme(self):
        th = THEMES[self.theme_name]
        self.root.configure(bg=th["bg"])
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=th["bg"])
        style.configure("TLabel", background=th["bg"], foreground=th["fg"])
        style.configure("TLabelFrame", background=th["bg"], foreground=th["fg"])
        style.configure("TButton", background=th["panel"], foreground=th["fg"])
        style.configure("TCheckbutton", background=th["bg"], foreground=th["fg"])
        style.configure("TCombobox", fieldbackground=th["entry_bg"], foreground=th["fg"])
        style.configure("TPanedwindow", background=th["bg"])

        self.text.configure(bg=th["entry_bg"], fg=th["fg"],
                            insertbackground=th["fg"], selectbackground=th["sel"])
        self.outline.configure(bg=th["list_bg"], fg=th["fg"], selectbackground=th["sel"])
        self.status.configure(background=th["panel"], foreground=th["muted"])
        if hasattr(self, "preview"):
            self.preview.configure(bg=th["entry_bg"], fg=th["fg"],
                                   selectbackground=th["sel"], insertbackground=th["fg"])
            self._configure_preview_tags()
        self._refresh_toggles()

    def _configure_preview_tags(self):
        """配置预览 Text 的样式标签（与导出一致：全部内容纯黑色）。"""
        pv = self.preview
        # 套用用户在界面选择的字体与正文字号（所见即所得）
        bf = self.body_font_var.get() or "Microsoft YaHei"
        hf = self.head_font_var.get() or "Microsoft YaHei"
        tf = self.title_font_var.get() or "Microsoft YaHei"
        try:
            bs = float(size_pt(self.body_size_var.get()))
        except Exception:
            bs = 11.0
        b = int(round(bs))  # tkinter 字号必须为整数
        BLACK = "#000000"  # 导出全部黑色，预览同步全黑
        pv.tag_configure("title", font=(tf, b + 6, "bold"),
                         foreground=BLACK, spacing1=6, spacing3=8)
        pv.tag_configure("h1", font=(hf, b + 4, "bold"),
                         foreground=BLACK, spacing1=6, spacing3=4)
        pv.tag_configure("h2", font=(hf, b + 2, "bold"),
                         foreground=BLACK, spacing1=4, spacing3=3)
        pv.tag_configure("h3", font=(hf, b + 1, "bold"),
                         foreground=BLACK, spacing1=3, spacing3=2)
        pv.tag_configure("body", font=(bf, b), foreground=BLACK,
                         spacing1=2, spacing3=4, lmargin1=4, lmargin2=4)
        pv.tag_configure("bold", font=(bf, b, "bold"),
                         foreground=BLACK)
        pv.tag_configure("ital", font=(bf, b, "italic"),
                         foreground=BLACK)
        pv.tag_configure("sup", font=(bf, max(6, b - 3)), foreground=BLACK)
        pv.tag_configure("list", font=(bf, b), foreground=BLACK,
                         lmargin1=12, lmargin2=24, spacing1=1, spacing3=2)
        pv.tag_configure("ref", font=(bf, b), foreground=BLACK,
                         lmargin1=14, lmargin2=26, spacing1=1, spacing3=2)
        pv.tag_configure("hr", font=(bf, b), foreground=BLACK)
        pv.tag_configure("mono", font=("Consolas", 10), foreground=BLACK)
        pv.tag_configure("monohead", font=("Consolas", 10,  "bold"),
                         foreground=BLACK)
        pv.tag_configure("code", font=("Consolas", max(9, b - 1)), foreground=BLACK)
        pv.tag_configure("link", font=(bf, b), foreground="#1a0dab", underline=True)
        pv.tag_configure("center", justify="center")

    # ------------------------------------------------------------ 功能
    def refresh_outline(self):
        text = self.text.get("1.0", "end-1c")
        try:
            self.blocks = formatter.parse_document(text, smart_heading=self.smart_var.get())
        except Exception as e:
            self.set_status(f"{t('err_export', self.lang)}: {e}")
            return
        self.outline.delete(0, "end")
        for lvl, txt in formatter.outline(self.blocks):
            indent = "    " * lvl
            prefix = {0: "■ ", 1: "● ", 2: "○ ", 3: "· "}.get(lvl, "· ")
            self.outline.insert("end", f"{indent}{prefix}{txt}")
        self.set_status(t("status_parsed", self.lang).format(
            n=len(self.blocks), h=len(formatter.outline(self.blocks))))
        self.render_preview()

    # --------------------------------------------------- 实时预览（防抖）
    def _schedule_update(self):
        """输入防抖：停止输入 300ms 后才重新解析并刷新预览。"""
        if getattr(self, "_after_id", None):
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.root.after(300, self._live_update)

    def _live_update(self):
        self.refresh_outline()

    def render_preview(self):
        """把当前 blocks 以近似样式渲染到预览面板（所见即所得的轻量近似）。"""
        pv = getattr(self, "preview", None)
        if pv is None:
            return
        # 字体/字号下拉改变后，先把最新字体套用到各标签（所见即所得）
        self._configure_preview_tags()
        try:
            pv.configure(state="normal")
        except Exception:
            return
        pv.delete("1.0", "end")
        blocks = self.blocks or []
        preset = formatter.STYLE_PRESETS.get(self.preset_key, formatter.STYLE_PRESETS["默认"])
        pv.tag_configure("title", justify="center" if preset.get("title_center") else "left")
        for b in blocks:
            t = b["type"]
            if t == "title":
                pv.insert("end", b["text"] + "\n\n", "title")
            elif t == "heading":
                tag = {1: "h1", 2: "h2", 3: "h3"}.get(b["level"], "h3")
                pv.insert("end", b["text"] + "\n\n", tag)
            elif t == "hr":
                pv.insert("end", "─" * 28 + "\n\n", "hr")
            elif t == "list":
                counters = {}
                for level, ordered, text in b["items"]:
                    counters[level] = counters.get(level, 0) + 1
                    for deeper in [k for k in counters if k > level]:
                        counters[deeper] = 0
                    indent = "    " * level
                    prefix = (f"{counters[level]}. " if ordered else "• ")
                    self._preview_inline(pv, indent + prefix + text + "\n", "list")
                pv.insert("end", "\n")
            elif t == "code":
                for line in (b.get("text") or "").split("\n"):
                    pv.insert("end", line + "\n", "mono")
                pv.insert("end", "\n")
            elif t == "image":
                alt = b.get("alt") or ""
                src = b.get("src") or ""
                pv.insert("end", f"[图片] {alt}  ({src})\n\n", "body")
            elif t == "math":
                for line in (b.get("text") or "").split("\n"):
                    pv.insert("end", line + "\n", "ital")
                pv.insert("end", "\n")
            elif t == "table":
                self._preview_table(pv, b)
                pv.insert("end", "\n")
            elif t == "paragraph":
                self._preview_inline(pv, b["text"] + "\n\n", "body")
            elif t == "reference":
                for k, en in enumerate(b["entries"]):
                    if self.ref_var.get():
                        rs = self.ref_style_var.get()
                        n = k + 1
                        if rs == "gb7714":
                            prefix = f"[{n}] "
                        elif rs == "paren":
                            prefix = f"({n}) "
                        elif rs == "circle":
                            prefix = (CIRCLED[n - 1] + " ") if 1 <= n <= len(CIRCLED) else f"[{n}] "
                        elif rs == "superscript":
                            prefix = f"^{n} "
                        elif rs == "none":
                            prefix = ""
                        else:
                            prefix = f"{n}. "
                        self._preview_inline(pv, prefix + en + "\n", "ref")
                    else:
                        self._preview_inline(pv, en + "\n", "body")
                pv.insert("end", "\n")
        pv.configure(state="disabled")

    def _preview_inline(self, pv, text, base_tag):
        """段落内联渲染：加粗 / 公式斜体 / 引用上标 / 行内代码 / 超链接。"""
        for seg, bold, is_math, is_cite, is_code, link in formatter.iter_runs(
                text, math_mode=True, cite_mode=self.cite_var.get()):
            if not seg:
                continue
            if is_math:
                disp = formatter._strip_math_delim(seg) if self.math_var.get() else seg
                pv.insert("end", disp, "ital")
            elif is_code:
                pv.insert("end", seg, "code")
            elif link:
                pv.insert("end", seg, "link")
            elif is_cite:
                pv.insert("end", seg, "sup")
            elif bold:
                pv.insert("end", seg, "bold")
            else:
                pv.insert("end", seg, base_tag)

    def _preview_table(self, pv, b):
        """用等宽 + 制表线近似渲染表格（列宽按字体实际像素宽度测量，兼容全角字符）。"""
        import tkinter.font as tkfont
        header = list(b.get("header", []))
        rows = [list(r) for r in b.get("rows", [])]
        n_cols = max(len(header), max((len(r) for r in rows), default=0), 1)

        def norm(cells):
            cells = list(cells)
            while len(cells) < n_cols:
                cells.append("")
            return cells[:n_cols]

        header = norm(header)
        rows = [norm(r) for r in rows]

        # 用表格字体（Consolas）真实测量像素宽度，避免全角标点/汉字估算偏窄
        f = tkfont.Font(family="Consolas", size=10)
        space_w = max(1, f.measure(" "))
        bar_w = max(1, f.measure("─"))
        def measure(s):
            return max(1, f.measure(s))

        col_w = []
        for c in range(n_cols):
            wmax = measure(header[c])
            for r in rows:
                wmax = max(wmax, measure(r[c]))
            col_w.append(wmax + 6)  # 预留左右内边距

        def pad(s, c):
            need = col_w[c] - measure(s)
            return s + " " * max(0, round(need / space_w))

        def fmt(cells):
            return "│ " + " │ ".join(pad(cells[c], c) for c in range(n_cols)) + " │"

        def fmt_sep():
            bars = []
            for c in range(n_cols):
                nbar = max(1, round(col_w[c] / bar_w))
                bars.append("─" * nbar)
            return "├─" + "─┼─".join(bars) + "─┤"

        pv.insert("end", fmt(header) + "\n", "monohead")
        pv.insert("end", fmt_sep() + "\n", "mono")
        for r in rows:
            pv.insert("end", fmt(r) + "\n", "mono")

    def import_file(self):
        path = filedialog.askopenfilename(
            title="选择文本文件" if self.lang == "zh" else "Select text file",
            filetypes=[("文本文件", "*.txt *.md *.markdown"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            data = self._read_file(path)
        except Exception as e:
            messagebox.showerror(t("err_read", self.lang),
                                 t("err_read_msg", self.lang).format(e=e))
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", data)
        self.sample_lang = None  # 已导入，不再是示例
        self.source_dir = os.path.dirname(path)  # 用于解析图片相对路径
        self.refresh_outline()
        self.set_status(t("status_imported", self.lang).format(name=os.path.basename(path)))

    @staticmethod
    def _read_file(path: str) -> str:
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                with open(path, "r", encoding=enc) as f:
                    return f.read()
            except (UnicodeDecodeError, UnicodeError):
                continue
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    def clear_text(self):
        self.text.delete("1.0", "end")
        self.outline.delete(0, "end")
        self.blocks = []
        self.sample_lang = None
        self.set_status(t("status_cleared", self.lang))

    def export_docx(self):
        text = self.text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning(t("warn_empty", self.lang), t("warn_empty_msg", self.lang))
            return
        path = filedialog.asksaveasfilename(
            title="保存 Word 文档" if self.lang == "zh" else "Save Word document",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialdir=self._default_export_dir(),
            initialfile="排版结果.docx" if self.lang == "zh" else "formatted.docx")
        if not path:
            return
        try:
            self.set_status(t("status_exporting", self.lang))
            self.root.update_idletasks()
            blocks = formatter.format_text_to_docx(
                text, path,
                preset_name=self.preset_key,
                smart_heading=self.smart_var.get(),
                auto_toc=self.toc_var.get(),
                math_pretty=self.math_var.get(),
                cite_sup=self.cite_var.get(),
                header_text=self.header_var.get(),
                footer_mode=self.footer_mode_var.get(),
                footer_text=self.footer_text_var.get(),
                footer_align=self.footer_align_var.get(),
                ref_auto=self.ref_var.get(),
                ref_style=self.ref_style_var.get(),
                ref_hang=self.ref_hang_var.get(),
                ref_line=self.ref_line_var.get(),
                body_font=self.body_font_var.get(),
                head_font=self.head_font_var.get(),
                title_font=self.title_font_var.get(),
                body_size=size_pt(self.body_size_var.get()),
                latin_font=self.latin_font_var.get(),
                base_dir=getattr(self, "source_dir", None))
            self.blocks = blocks
            self.populate_outline_from(blocks)
            messagebox.showinfo(t("info_done", self.lang),
                                t("info_done_msg", self.lang).format(path=path, n=len(blocks)))
            self.set_status(t("status_exported", self.lang).format(name=os.path.basename(path)))
        except Exception as e:
            messagebox.showerror(t("err_export", self.lang), f"{e}")
            self.set_status(t("status_export_fail", self.lang))

    def export_pdf(self):
        """先生成临时 docx，再调用 Word / LibreOffice 转换为 PDF。"""
        from pdf_export import docx_to_pdf
        import tempfile

        text = self.text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning(t("warn_empty", self.lang), t("warn_empty_msg", self.lang))
            return
        path = filedialog.asksaveasfilename(
            title="保存 PDF" if self.lang == "zh" else "Save PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文档", "*.pdf"), ("所有文件", "*.*")],
            initialdir=self._default_export_dir(),
            initialfile="排版结果.pdf" if self.lang == "zh" else "formatted.pdf")
        if not path:
            return
        tmp = tempfile.mktemp(suffix=".docx")
        try:
            self.set_status(t("pdf_status_exporting", self.lang))
            self.root.update_idletasks()
            formatter.format_text_to_docx(
                text, tmp,
                preset_name=self.preset_key,
                smart_heading=self.smart_var.get(),
                auto_toc=self.toc_var.get(),
                math_pretty=self.math_var.get(),
                cite_sup=self.cite_var.get(),
                header_text=self.header_var.get(),
                footer_mode=self.footer_mode_var.get(),
                footer_text=self.footer_text_var.get(),
                footer_align=self.footer_align_var.get(),
                ref_auto=self.ref_var.get(),
                ref_style=self.ref_style_var.get(),
                ref_hang=self.ref_hang_var.get(),
                ref_line=self.ref_line_var.get(),
                body_font=self.body_font_var.get(),
                head_font=self.head_font_var.get(),
                title_font=self.title_font_var.get(),
                body_size=size_pt(self.body_size_var.get()),
                latin_font=self.latin_font_var.get(),
                base_dir=getattr(self, "source_dir", None))
            if docx_to_pdf(tmp, path):
                messagebox.showinfo(t("info_done", self.lang),
                                    t("info_done_msg", self.lang).format(path=path, n=len(self.blocks)))
                self.set_status(t("pdf_status_exported", self.lang).format(name=os.path.basename(path)))
            else:
                messagebox.showerror(t("pdf_fail_title", self.lang), t("pdf_fail_msg", self.lang))
                self.set_status(t("status_export_fail", self.lang))
        except Exception as e:
            messagebox.showerror(t("err_export", self.lang), f"{e}")
            self.set_status(t("status_export_fail", self.lang))
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass

    def populate_outline_from(self, blocks):
        self.outline.delete(0, "end")
        for lvl, txt in formatter.outline(blocks):
            indent = "    " * lvl
            prefix = {0: "■ ", 1: "● ", 2: "○ ", 3: "· "}.get(lvl, "· ")
            self.outline.insert("end", f"{indent}{prefix}{txt}")

    # --------------------------------------------------- 配置导入/导出
    def _current_config(self) -> dict:
        return {
            "app": "DocFormatter",
            "version": APP_VERSION,
            "config": {
                "smart_heading": self.smart_var.get(),
                "auto_toc": self.toc_var.get(),
                "math_pretty": self.math_var.get(),
                "cite_sup": self.cite_var.get(),
                "header_text": self.header_var.get(),
                "footer_mode": self.footer_mode_var.get(),
                "footer_align": self.footer_align_var.get(),
                "ref_auto": self.ref_var.get(),
                "ref_style": self.ref_style_var.get(),
                "ref_hang": self.ref_hang_var.get(),
                "ref_line": self.ref_line_var.get(),
                "body_font": self.body_font_var.get(),
                "head_font": self.head_font_var.get(),
                "title_font": self.title_font_var.get(),
                "body_size": size_pt(self.body_size_var.get()),
                "footer_text": self.footer_text_var.get(),
                "western_font": self.latin_font_var.get(),
                "preset": self.preset_key,
                "language": self.lang,
                "export_locations": [dict(x) for x in self.export_locations],
            },
        }

    def export_config(self):
        cfg = self._current_config()
        path = filedialog.asksaveasfilename(
            title=t("btn_export_cfg", self.lang),
            defaultextension=".json",
            filetypes=[(t("cfg_filetype", self.lang), "*.json"), ("所有文件", "*.*")],
            initialfile="DocFormatter配置.json")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self.set_status(t("status_cfg_exported", self.lang).format(name=os.path.basename(path)))
        except Exception as e:
            messagebox.showerror(t("err_export", self.lang), f"{e}")

    def import_config(self):
        path = filedialog.askopenfilename(
            title=t("btn_import_cfg", self.lang),
            filetypes=[(t("cfg_filetype", self.lang), "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror(t("err_read", self.lang),
                                 t("err_read_msg", self.lang).format(e=e))
            return
        cfg = data.get("config") if isinstance(data, dict) else None
        if not isinstance(cfg, dict) or data.get("app") != "DocFormatter":
            messagebox.showerror(t("status_cfg_invalid", self.lang),
                                 t("warn_cfg_invalid_msg", self.lang))
            return
        self.apply_config(cfg)
        self.set_status(t("status_cfg_imported", self.lang).format(name=os.path.basename(path)))

    def apply_config(self, cfg: dict):
        old_lang = self.lang
        self.smart_var.set(bool(cfg.get("smart_heading", self.smart_var.get())))
        self.toc_var.set(bool(cfg.get("auto_toc", self.toc_var.get())))
        self.math_var.set(bool(cfg.get("math_pretty", self.math_var.get())))
        self.cite_var.set(bool(cfg.get("cite_sup", self.cite_var.get())))
        self.ref_var.set(bool(cfg.get("ref_auto", self.ref_var.get())))
        if "ref_style" in cfg and cfg["ref_style"] in self.ref_styles:
            self.ref_style_var.set(cfg["ref_style"])
        if "ref_hang" in cfg:
            try:
                self.ref_hang_var.set(float(cfg["ref_hang"]))
            except Exception:
                pass
        if "ref_line" in cfg:
            try:
                self.ref_line_var.set(float(cfg["ref_line"]))
            except Exception:
                pass
        self.header_var.set(str(cfg.get("header_text", self.header_var.get())))
        if "footer_mode" in cfg and cfg["footer_mode"] in self.footer_modes:
            self.footer_mode_var.set(cfg["footer_mode"])
        if "footer_align" in cfg and cfg["footer_align"] in FOOTER_ALIGNS:
            self.footer_align_var.set(cfg["footer_align"])
        preset = cfg.get("preset", self.preset_key)
        if preset in PRESET_KEYS:
            self.preset_key = preset
        # 先把预设作为基线套用字体/字号；随后若配置含显式覆盖值则再覆盖
        _p = STYLE_PRESETS.get(self.preset_key, STYLE_PRESETS["默认"])
        self.body_font_var.set(_p["body_font"])
        self.head_font_var.set(_p["head_font"])
        self.title_font_var.set(_p["title_font"])
        self.body_size_var.set(size_label(_p["body_size"]))
        if "body_font" in cfg:
            self.body_font_var.set(str(cfg["body_font"]))
        if "head_font" in cfg:
            self.head_font_var.set(str(cfg["head_font"]))
        if "title_font" in cfg:
            self.title_font_var.set(str(cfg["title_font"]))
        if "body_size" in cfg:
            try:
                self.body_size_var.set(size_label(float(cfg["body_size"])))
            except Exception:
                pass
        if "footer_text" in cfg:
            self.footer_text_var.set(str(cfg["footer_text"]))
        if "western_font" in cfg and cfg["western_font"] in WESTERN_FONT_OPTIONS:
            self.latin_font_var.set(str(cfg["western_font"]))
            self._save_western_font()
        lang = cfg.get("language", self.lang)
        if lang in ("zh", "en"):
            self.lang = lang
        # 常用导出位置（含勾选状态）随配置恢复，并写入本机本地文件
        if "export_locations" in cfg and isinstance(cfg["export_locations"], list):
            loaded = []
            for it in cfg["export_locations"]:
                if isinstance(it, dict) and it.get("path"):
                    loaded.append({"path": str(it["path"]),
                                   "enabled": bool(it.get("enabled", True))})
            self.export_locations = loaded
            self._save_locations()
        self.retranslate()   # 同步 主题/语言/预设 下拉 + 对号开关
        self.apply_theme()
        # 若文本框仍是旧语言示例，则换成新语言示例
        cur = self.text.get("1.0", "end-1c")
        if cur.strip() != "" and cur == t("sample", old_lang):
            self.text.delete("1.0", "end")
            self.text.insert("1.0", t("sample", self.lang))
            self.sample_lang = self.lang
        self.refresh_outline()
        if hasattr(self, "loc_inner"):
            self._refresh_loc_list()

    # --------------------------------------------------- 常用导出位置
    def _build_export_loc(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        hint = ttk.Label(parent, text="", wraplength=320, justify="left")
        hint.grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))
        self._tr.append((hint, "lbl_export_loc_hint"))

        cv = tk.Canvas(parent, bg=THEMES[self.theme_name]["list_bg"], highlightthickness=0)
        cv.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        sb = ttk.Scrollbar(parent, command=cv.yview)
        sb.grid(row=1, column=1, sticky="ns")
        cv.config(yscrollcommand=sb.set)
        self.loc_canvas = cv
        self.loc_inner = ttk.Frame(cv)
        self._loc_win = cv.create_window((0, 0), window=self.loc_inner, anchor="nw")
        # 内层框架随画布宽度变化，使勾选项横向铺满
        cv.bind("<Configure>", lambda e: cv.itemconfig(self._loc_win, width=e.width))
        self.loc_inner.bind("<Configure>",
                            lambda e: cv.configure(scrollregion=cv.bbox("all")))

        bf = ttk.Frame(parent)
        bf.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=6)
        self.w_add_loc = ttk.Button(bf, text="", command=self._add_loc)
        self.w_add_loc.pack(side="left", padx=4)
        self._tr.append((self.w_add_loc, "btn_add_loc"))

    def _refresh_loc_list(self):
        if not hasattr(self, "loc_inner"):
            return
        for w in list(self.loc_inner.winfo_children()):
            w.destroy()
        self._loc_rows = []
        if not self.export_locations:
            empty = ttk.Label(self.loc_inner, text=t("lbl_no_loc", self.lang),
                              wraplength=300, justify="left")
            empty.pack(anchor="w", padx=6, pady=8)
            return
        for i, loc in enumerate(self.export_locations):
            var = tk.BooleanVar(value=bool(loc.get("enabled", False)))
            row = ttk.Frame(self.loc_inner)
            row.pack(anchor="w", fill="x", padx=2, pady=1)
            cb = ttk.Checkbutton(
                row, text=loc["path"], variable=var,
                command=lambda idx=i, v=var: self._toggle_loc(idx, v))
            cb.pack(side="left", anchor="w", fill="x", expand=True)
            del_btn = ttk.Button(row, text="×", width=3,
                                command=lambda idx=i: self._del_loc(idx))
            del_btn.pack(side="right", padx=2)
            self._loc_rows.append((var, cb))

    def _add_loc(self):
        d = filedialog.askdirectory(
            title="选择常用导出位置" if self.lang == "zh" else "Select folder")
        if not d:
            return
        norm = os.path.normpath(d)
        for loc in self.export_locations:
            if os.path.normpath(loc["path"]) == norm:
                loc["enabled"] = True
                self._refresh_loc_list()
                self._save_locations()
                self.set_status(t("status_loc_added", self.lang).format(path=d))
                return
        self.export_locations.append({"path": d, "enabled": True})
        self._refresh_loc_list()
        self._save_locations()
        self.set_status(t("status_loc_added", self.lang).format(path=d))

    def _del_loc(self, idx):
        if 0 <= idx < len(self.export_locations):
            self.export_locations.pop(idx)
            self._refresh_loc_list()
            self._save_locations()
            self.set_status(t("status_loc_removed", self.lang))

    def _toggle_loc(self, idx, var):
        if 0 <= idx < len(self.export_locations):
            self.export_locations[idx]["enabled"] = bool(var.get())
            self._save_locations()

    def _default_export_dir(self):
        for loc in self.export_locations:
            if loc.get("enabled") and os.path.isdir(loc["path"]):
                return loc["path"]
        for loc in self.export_locations:
            if os.path.isdir(loc["path"]):
                return loc["path"]
        return None

    def _load_locations(self):
        try:
            if os.path.exists(LOC_PATH):
                with open(LOC_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    out = []
                    for it in data:
                        if isinstance(it, dict) and it.get("path"):
                            out.append({"path": str(it["path"]),
                                        "enabled": bool(it.get("enabled", True))})
                    return out
        except Exception:
            pass
        return []

    def _save_locations(self):
        try:
            with open(LOC_PATH, "w", encoding="utf-8") as f:
                json.dump(self.export_locations, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # 西文字体（英文/数字）是本机偏好设置：写入 exe 同目录的 settings.json，关掉重开不丢
    def _load_western_font(self):
        try:
            if os.path.exists(SETTINGS_PATH):
                with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                    d = json.load(f)
                if isinstance(d, dict) and d.get("western_font"):
                    return str(d["western_font"])
        except Exception:
            pass
        return "Times New Roman"

    def _save_western_font(self):
        try:
            d = {}
            if os.path.exists(SETTINGS_PATH):
                try:
                    with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                        d = json.load(f)
                except Exception:
                    d = {}
            d["western_font"] = self.latin_font_var.get()
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def show_about(self):
        messagebox.showinfo(t("about_title", self.lang), t("about_text", self.lang))

    # ------------------------------------------------------- 批量处理
    def open_batch(self):
        if getattr(self, "_batch_win", None) and self._batch_win.winfo_exists():
            self._batch_win.lift()
            return
        win = tk.Toplevel(self.root)
        self._batch_win = win
        win.title(t("batch_title", self.lang))
        win.geometry("580x480")
        win.transient(self.root)
        try:
            win.grab_set()
        except Exception:
            pass

        self.batch_mode = tk.StringVar(value="files")
        self.batch_input = []
        self.batch_folder = ""
        self.batch_output = ""
        self._batch_running = False
        self._batch_log_q = queue.Queue()
        self._batch_files = []
        self._batch_ok = self._batch_fail = self._batch_done = self._batch_total = 0

        mode_f = ttk.LabelFrame(win, text=t("batch_title", self.lang))
        mode_f.pack(fill="x", padx=10, pady=8)
        ttk.Radiobutton(mode_f, text=t("batch_mode_files", self.lang),
                        variable=self.batch_mode, value="files",
                        command=lambda: self._batch_mode_changed()).pack(side="left", padx=12, pady=6)
        ttk.Radiobutton(mode_f, text=t("batch_mode_folder", self.lang),
                        variable=self.batch_mode, value="folder",
                        command=lambda: self._batch_mode_changed()).pack(side="left", padx=12, pady=6)

        in_f = ttk.Frame(win)
        in_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(in_f, text=t("batch_lbl_input", self.lang)).pack(side="left")
        self.batch_in_lbl = ttk.Label(in_f, text="—", foreground="#888")
        self.batch_in_lbl.pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(in_f, text=t("batch_btn_pick", self.lang),
                   command=self._batch_pick).pack(side="right", padx=4)

        out_f = ttk.Frame(win)
        out_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(out_f, text=t("batch_lbl_output", self.lang)).pack(side="left")
        self.batch_out_lbl = ttk.Label(out_f, text="—", foreground="#888")
        self.batch_out_lbl.pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(out_f, text=t("batch_btn_pick", self.lang),
                   command=self._batch_pick_out).pack(side="right", padx=4)

        sum_f = ttk.LabelFrame(win, text=t("frame_options", self.lang))
        sum_f.pack(fill="x", padx=10, pady=6)
        self.batch_sum_lbl = ttk.Label(sum_f, text="", wraplength=520, justify="left")
        self.batch_sum_lbl.pack(anchor="w", padx=8, pady=6)
        self._batch_update_summary()

        self.batch_pb = ttk.Progressbar(win, mode="determinate", maximum=100)
        self.batch_pb.pack(fill="x", padx=10, pady=6)
        self.batch_log = tk.Text(win, height=11, font=("Consolas", 9), state="disabled")
        self.batch_log.pack(fill="both", expand=True, padx=10, pady=4)
        ttk.Button(win, text=t("batch_btn_start", self.lang),
                   command=self._batch_start).pack(pady=8)

    def _batch_mode_changed(self):
        self.batch_input = []
        self.batch_folder = ""
        if hasattr(self, "batch_in_lbl"):
            self.batch_in_lbl.config(text="—")

    def _batch_pick(self):
        if self.batch_mode.get() == "files":
            paths = filedialog.askopenfilenames(
                title=t("batch_pick_files", self.lang),
                filetypes=[("文本文件", "*.txt *.md *.markdown"), ("所有文件", "*.*")])
            if paths:
                self.batch_input = list(paths)
                self.batch_in_lbl.config(text=f"{len(paths)} 个文件")
                d = os.path.dirname(paths[0])
                if not self.batch_output:
                    self.batch_output = d
                    self.batch_out_lbl.config(text=d)
        else:
            d = filedialog.askdirectory(title=t("batch_pick_folder", self.lang))
            if d:
                self.batch_folder = d
                self.batch_in_lbl.config(text=d)
                if not self.batch_output:
                    self.batch_output = d
                    self.batch_out_lbl.config(text=d)

    def _batch_pick_out(self):
        d = filedialog.askdirectory(title=t("batch_lbl_output", self.lang))
        if d:
            self.batch_output = d
            self.batch_out_lbl.config(text=d)

    def _batch_update_summary(self):
        if not hasattr(self, "batch_sum_lbl"):
            return
        mark = lambda v: "✓" if v else "○"
        preset = t(PRESET_I18N[self.preset_key], self.lang)
        s = (f"{t('lbl_preset', self.lang)} {preset}    "
             f"{mark(self.smart_var.get())}{t('chk_smart', self.lang)}  "
             f"{mark(self.toc_var.get())}{t('chk_toc', self.lang)}  "
             f"{mark(self.math_var.get())}{t('chk_math', self.lang)}  "
             f"{mark(self.cite_var.get())}{t('chk_cite', self.lang)}  "
             f"{mark(self.ref_var.get())}{t('chk_ref', self.lang)}")
        self.batch_sum_lbl.config(text=s)

    def _batch_start(self):
        if self._batch_running:
            return
        if self.batch_mode.get() == "files":
            files = list(self.batch_input)
        else:
            files = []
            if self.batch_folder:
                for root, _, fnames in os.walk(self.batch_folder):
                    for fn in fnames:
                        if fn.lower().endswith((".txt", ".md", ".markdown")):
                            files.append(os.path.join(root, fn))
        if not files:
            self._batch_log_write(t("batch_no_files", self.lang))
            return
        if not self.batch_output:
            self.batch_output = os.path.dirname(files[0]) or "."
        try:
            os.makedirs(self.batch_output, exist_ok=True)
        except Exception:
            pass
        self._batch_running = True
        self._batch_files = files
        self._batch_ok = self._batch_fail = self._batch_done = 0
        self._batch_total = len(files)
        self.batch_pb["value"] = 0
        self._batch_log_clear()
        threading.Thread(target=self._batch_worker, daemon=True).start()
        self._batch_poll()

    def _batch_worker(self):
        files = self._batch_files
        for i, fpath in enumerate(files):
            try:
                data = self._read_file(fpath)
                base = os.path.splitext(os.path.basename(fpath))[0]
                out = os.path.join(self.batch_output, base + ".docx")
                formatter.format_text_to_docx(
                    data, out, preset_name=self.preset_key,
                    smart_heading=self.smart_var.get(), auto_toc=self.toc_var.get(),
                    math_pretty=self.math_var.get(), cite_sup=self.cite_var.get(),
                    header_text=self.header_var.get(),
                    footer_mode=self.footer_mode_var.get(),
                    footer_text=self.footer_text_var.get(),
                    footer_align=self.footer_align_var.get(),
                    ref_auto=self.ref_var.get(), ref_style=self.ref_style_var.get(),
                    ref_hang=self.ref_hang_var.get(), ref_line=self.ref_line_var.get(),
                    body_font=self.body_font_var.get(),
                    head_font=self.head_font_var.get(),
                    title_font=self.title_font_var.get(),
                body_size=size_pt(self.body_size_var.get()),
                latin_font=self.latin_font_var.get(),
                base_dir=os.path.dirname(fpath))
                self._batch_log_q.put(("ok", fpath, ""))
                self._batch_ok += 1
            except Exception as e:
                self._batch_log_q.put(("fail", fpath, str(e)))
                self._batch_fail += 1
            self._batch_done = i + 1
        self._batch_running = False

    def _batch_poll(self):
        if hasattr(self, "_batch_log_q"):
            while True:
                try:
                    kind, fpath, err = self._batch_log_q.get_nowait()
                except queue.Empty:
                    break
                name = os.path.basename(fpath)
                if kind == "ok":
                    self._batch_log_write(t("batch_log_ok", self.lang).format(name=name))
                else:
                    self._batch_log_write(t("batch_log_fail", self.lang).format(name=name, e=err))
        if self._batch_total:
            self.batch_pb["value"] = int(100 * self._batch_done / self._batch_total)
        if self._batch_running:
            self.root.after(100, self._batch_poll)
        else:
            self.batch_pb["value"] = 100
            self._batch_log_write(t("batch_status_done", self.lang).format(
                ok=self._batch_ok, fail=self._batch_fail))
            self.set_status(t("status_exported", self.lang).format(
                name=os.path.basename(self.batch_output or "")))

    def _batch_log_write(self, msg):
        log = getattr(self, "batch_log", None)
        if log is None:
            return
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    def _batch_log_clear(self):
        log = getattr(self, "batch_log", None)
        if log is None:
            return
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")

    def set_status(self, msg: str):
        self.status.configure(text=msg)


def main():
    # 支持 `DocFormatter.exe --version` 查看版本（README 已记录该用法）
    if "--version" in sys.argv[1:]:
        print(f"DocFormatter {APP_VERSION}")
        return
    root = tk.Tk()
    App(root)
    # 启动时最大化（全屏窗口，带标题栏，可拖动边缘缩小到普通窗口大小）。
    # 先给一个较大默认尺寸作为兜底（非 Windows 或 zoom 失败时），再尝试最大化。
    root.geometry("1280x820")
    root.minsize(880, 560)  # 允许缩小，但不至于塌成不可用
    try:
        root.state("zoomed")  # Windows：最大化（保留标题栏/边框，可缩放）
    except Exception:
        try:
            root.attributes("-zoomed", True)
        except Exception:
            pass
    root.mainloop()


if __name__ == "__main__":
    main()

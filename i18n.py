# -*- coding: utf-8 -*-
"""
文档快速排版 —— 多语言文本 (简体中文 / English)
所有界面文案集中此处，切换语言时由 gui.py 统一刷新。
"""

STRINGS = {
    "zh": {
        "app_title": "文档快速排版",
        "btn_import": "导入文件",
        "btn_clear": "清空",
        "btn_refresh": "刷新大纲",
        "btn_export": "导出 Word",
        "btn_export_all": "一键排版导出",
        "btn_about": "关于",
        "btn_export_cfg": "导出配置",
        "btn_import_cfg": "导入配置",
        "cfg_filetype": "DocFormatter 配置",
        "status_cfg_exported": "配置已导出：{name}",
        "status_cfg_imported": "配置已导入：{name}",
        "status_cfg_invalid": "配置文件无效",
        "warn_cfg_invalid_msg": "该文件不是有效的 DocFormatter 配置。",
        "lbl_input": "原始文本（可粘贴或导入 .txt / .md）",
        "frame_options": "排版选项",
        "chk_smart": "智能识别标题（普通文本）",
        "chk_toc": "自动目录（Word 域）",
        "chk_math": "保留并美化数学公式（LaTeX）",
        "chk_cite": "引用标注 [1] 设为上标",
        "chk_ref": "自动整理参考文献（悬挂缩进+自动编号）",
        "frame_ref": "参考文献",
        "lbl_ref_style": "编号样式：",
        "ref_style_numbered": "顺序编号 1. 2. 3.",
        "ref_style_gb": "国标 [1] [2] [3]",
        "ref_style_paren": "圆括号 (1) (2) (3)",
        "ref_style_circle": "带圈数字 ① ② ③",
        "ref_style_sup": "上标编号 ¹ ² ³",
        "ref_style_none": "仅悬挂缩进（著者-出版年制）",
        "lbl_ref_hang": "悬挂缩进(cm)：",
        "lbl_ref_line": "行距：",
        "lbl_preset": "样式预设：",
        "lbl_lang": "语言：",
        "frame_outline": "结构大纲（标题层级预览）",
        "frame_font": "字体与字号",
        "frame_page": "页面设置",
        "lbl_header": "页眉内容：",
        "lbl_footer": "页脚：",
        "lbl_footer_text": "页脚文字：",
        "footer_none": "无",
        "footer_page": "页码",
        "footer_page_of_total": "第 X 页 / 共 Y 页",
        "footer_date": "日期",
        "footer_date_page": "日期 + 页码（左日期 / 右页码）",
        "footer_text": "自定义文字",
        "footer_text_page": "自定义文字（左）+ 页码（右）",
        "lbl_footer_align": "页码对齐：",
        "align_left": "居左",
        "align_center": "居中",
        "align_right": "居右",
        "lbl_body_font": "正文字体：",
        "lbl_head_font": "标题字体：",
        "lbl_title_font": "主标题字体：",
        "lbl_body_size": "正文字号：",
        "lbl_western_font": "西文字体：",
        "tab_outline": "大纲",
        "tab_preview": "预览",
        "btn_export_pdf": "导出 PDF",
        "tab_export_loc": "导出位置",
        "frame_export_loc": "常用导出位置",
        "lbl_export_loc_hint": "勾选常用导出位置，导出时自动以此为默认目录；这些位置会随“导出配置”一起保存到其它电脑。",
        "btn_add_loc": "＋ 添加文件夹",
        "btn_remove_loc": "移除选中",
        "lbl_no_loc": "（暂无常用位置，点“添加文件夹”添加）",
        "status_loc_added": "已添加导出位置：{path}",
        "status_loc_removed": "已移除导出位置",
        "status_loc_default": "默认导出目录：{path}",
        "warn_no_loc_sel": "请先勾选要移除的位置。",
        "btn_batch": "批量处理",
        "batch_title": "批量处理",
        "batch_mode_files": "多个文件",
        "batch_mode_folder": "整个文件夹",
        "batch_btn_pick": "选择…",
        "batch_lbl_input": "输入：",
        "batch_lbl_output": "输出文件夹：",
        "batch_btn_start": "开始批量",
        "batch_status_ready": "就绪，请选择文件或文件夹",
        "batch_status_done": "完成：成功 {ok} 个，失败 {fail} 个",
        "batch_status_running": "正在处理 {i}/{n}：{name}",
        "batch_log_ok": "✓ {name}",
        "batch_log_fail": "✗ {name}：{e}",
        "batch_no_files": "尚未选择任何文件",
        "batch_pick_files": "选择文本文件",
        "batch_pick_folder": "选择文件夹",
        "pdf_status_exporting": "正在导出 PDF…",
        "pdf_status_exported": "已导出 PDF：{name}",
        "pdf_fail_title": "无法导出 PDF",
        "pdf_fail_msg": "本机未检测到可用于转换的 Word 或 LibreOffice。请安装 Microsoft Word 或 LibreOffice 后重试，或先导出 Word 再用其“另存为 PDF”。",

        "lang_zh": "简体中文",
        "lang_en": "English",
        "preset_default": "默认",
        "preset_graduation": "毕业论文",
        "preset_general": "通用论文",
        "preset_course": "课程论文",
        "preset_simple": "简洁风",
        "preset_read": "阅读风",
        "status_ready": "就绪",
        "status_parsed": "已识别 {n} 个结构块，标题 {h} 个",
        "status_imported": "已导入：{name}",
        "status_cleared": "已清空",
        "status_exporting": "正在生成文档…",
        "status_exported": "已导出：{name}",
        "status_export_fail": "导出失败",
        "status_lang": "语言：{lang}",
        "warn_empty": "空内容",
        "warn_empty_msg": "请先输入或导入文本。",
        "err_read": "读取失败",
        "err_read_msg": "无法读取文件：{e}",
        "err_export": "导出失败",
        "info_done": "完成",
        "info_done_msg": "已导出：\n{path}\n共 {n} 个结构块。",
        "about_title": "关于",
        "about_text": (
            "文档快速排版 v1.0.2\n\n"
            "纯离线启发式排版工具：\n"
            "杂乱文本 / TXT / Markdown -> 一键整理标题、段落、列表、表格、公式、引用标注、参考文献、重点加粗 -> 规范 Word(.docx)\n\n"
            "支持：页眉页脚与自动页码、实时预览、批量处理、导出 PDF。\n"
            "不依赖大模型，保护隐私，打开即用。"
        ),
        "sample": (
            "我的学习笔记\n\n"
            "一、绪论\n"
            "这是绪论部分的内容。它描述了研究的背景与意义，内容比较长比较长比较长比较长比较长。\n\n"
            "1. 研究背景\n"
            "背景包括很多方面，例如技术发展与社会需求。\n\n"
            "2. 研究意义\n"
            "意义在于提升效率。\n\n"
            "重点：这一部分很重要。\n"
            "定义：模型是指一组参数。\n\n"
            "- 苹果\n"
            "- 香蕉\n"
            "- 橘子\n\n"
            "第二章 方法\n"
            "方法是本文的核心。继续写一些内容来测试段落合并是否正确，"
            "this is English and should keep space.\n"
            "\n"
            "三、方法对比\n"
            "| 方法 | 准确率 | 速度 |\n"
            "|:---|:---:|---:|\n"
            "| 方法甲 | 92% | 快 |\n"
            "| 方法乙 | 88% | 中 |\n"
            "\n"
            "四、重要公式\n"
            "欧拉公式 $e^{i\\pi}+1=0$ 被誉为数学中最优美的等式。\n"
            "质能方程如下：\n"
            "$$\n"
            "E = mc^2\n"
            "$$\n"
            "求和公式可写作 $S = \\sum_{k=1}^{n} a_k$。\n"
            "\n"
            "五、文献综述\n"
            "已有研究指出该方法行之有效[1]，并在后续工作中得到扩展[2,3]。"
            "另有综述对其局限性进行了讨论[4-6]。\n"
            "\n"
            "参考文献\n"
            "[1] 作者一, 作者二. 论文题目[J]. 期刊名, 2020, 40(2): 123-130.\n"
            "[2] Author A, Author B. Title of the paper[C]. Conference, 2021.\n"
            "[3] 作者三. 专著书名[M]. 出版社, 2019.\n"
        ),
    },
    "en": {
        "app_title": "Doc Formatter",
        "btn_import": "Import",
        "btn_clear": "Clear",
        "btn_refresh": "Refresh Outline",
        "btn_export": "Export Word",
        "btn_export_all": "Format & Export",
        "btn_about": "About",
        "btn_export_cfg": "Export Config",
        "btn_import_cfg": "Import Config",
        "cfg_filetype": "DocFormatter Config",
        "status_cfg_exported": "Config exported: {name}",
        "status_cfg_imported": "Config imported: {name}",
        "status_cfg_invalid": "Invalid config file",
        "warn_cfg_invalid_msg": "This is not a valid DocFormatter config file.",
        "lbl_input": "Source text (paste, or import .txt / .md)",
        "frame_options": "Format Options",
        "chk_smart": "Auto-detect headings (plain text)",
        "chk_toc": "Auto table of contents (Word field)",
        "chk_math": "Keep & prettify math formulas (LaTeX)",
        "chk_cite": "Citation [1] as superscript",
        "chk_ref": "Auto-format references (hanging indent + auto number)",
        "frame_ref": "References",
        "lbl_ref_style": "Number style:",
        "ref_style_numbered": "Sequential 1. 2. 3.",
        "ref_style_gb": "GB/T 7714 [1] [2] [3]",
        "ref_style_paren": "Parentheses (1) (2) (3)",
        "ref_style_circle": "Circled ① ② ③",
        "ref_style_sup": "Superscript ¹ ² ³",
        "ref_style_none": "Hanging indent only (author-year)",
        "lbl_ref_hang": "Hanging indent (cm):",
        "lbl_ref_line": "Line spacing:",
        "lbl_preset": "Style preset:",
        "lbl_lang": "Language:",
        "frame_outline": "Outline (heading hierarchy)",
        "frame_font": "Font & size",
        "frame_page": "Page Setup",
        "lbl_header": "Header:",
        "lbl_footer": "Footer:",
        "lbl_footer_text": "Footer text:",
        "footer_none": "None",
        "footer_page": "Page number",
        "footer_page_of_total": "Page X of Y",
        "footer_date": "Date",
        "footer_date_page": "Date + page (left date / right page)",
        "footer_text": "Custom text",
        "footer_text_page": "Custom text (left) + page (right)",
        "lbl_footer_align": "Alignment:",
        "align_left": "Left",
        "align_center": "Center",
        "align_right": "Right",
        "lbl_body_font": "Body font:",
        "lbl_head_font": "Heading font:",
        "lbl_title_font": "Title font:",
        "lbl_body_size": "Body size:",
        "lbl_western_font": "Western font:",
        "tab_outline": "Outline",
        "tab_preview": "Preview",
        "btn_export_pdf": "Export PDF",
        "tab_export_loc": "Export Paths",
        "frame_export_loc": "Common Export Locations",
        "lbl_export_loc_hint": "Check your frequent export folders; exports open there by default. These locations travel with the config file to other PCs.",
        "btn_add_loc": "+ Add Folder",
        "btn_remove_loc": "Remove Selected",
        "lbl_no_loc": "(No common locations yet; click 'Add Folder')",
        "status_loc_added": "Added export location: {path}",
        "status_loc_removed": "Removed export location",
        "status_loc_default": "Default export dir: {path}",
        "warn_no_loc_sel": "Select a location to remove.",
        "btn_batch": "Batch",
        "batch_title": "Batch Processing",
        "batch_mode_files": "Multiple files",
        "batch_mode_folder": "Whole folder",
        "batch_btn_pick": "Choose…",
        "batch_lbl_input": "Input:",
        "batch_lbl_output": "Output folder:",
        "batch_btn_start": "Start",
        "batch_status_ready": "Ready, please select files or a folder",
        "batch_status_done": "Done: {ok} succeeded, {fail} failed",
        "batch_status_running": "Processing {i}/{n}: {name}",
        "batch_log_ok": "✓ {name}",
        "batch_log_fail": "✗ {name}: {e}",
        "batch_no_files": "No files selected yet",
        "batch_pick_files": "Select text files",
        "batch_pick_folder": "Select folder",
        "pdf_status_exporting": "Exporting PDF…",
        "pdf_status_exported": "PDF exported: {name}",
        "pdf_fail_title": "PDF export unavailable",
        "pdf_fail_msg": "No Word or LibreOffice found for conversion. Install Microsoft Word or LibreOffice, or export Word first and use 'Save as PDF'.",

        "lang_zh": "简体中文",
        "lang_en": "English",
        "preset_default": "Default",
        "preset_graduation": "Graduation Thesis",
        "preset_general": "General Paper",
        "preset_course": "Course Paper",
        "preset_simple": "Compact",
        "preset_read": "Reading",
        "status_ready": "Ready",
        "status_parsed": "Parsed {n} blocks, {h} headings",
        "status_imported": "Imported: {name}",
        "status_cleared": "Cleared",
        "status_exporting": "Generating document…",
        "status_exported": "Exported: {name}",
        "status_export_fail": "Export failed",
        "status_lang": "Language: {lang}",
        "warn_empty": "Empty",
        "warn_empty_msg": "Please enter or import some text first.",
        "err_read": "Read failed",
        "err_read_msg": "Cannot read file: {e}",
        "err_export": "Export failed",
        "info_done": "Done",
        "info_done_msg": "Exported:\n{path}\nTotal {n} blocks.",
        "about_title": "About",
        "about_text": (
            "Doc Formatter v1.0.2\n\n"
            "A fully offline, heuristic document formatter:\n"
            "Messy text / TXT / Markdown -> one-click headings, paragraphs,\n"
            "lists, tables, formulas, citations, references, key-term bolding -> clean Word(.docx)\n\n"
            "Features: header/footer & auto page numbers, live preview,\n"
            "batch processing, export to PDF.\n"
            "No LLM required. Private and ready to use."
        ),
        "sample": (
            "My Study Notes\n\n"
            "1. Introduction\n"
            "This is the introduction. It describes the background and significance "
            "of the research, and the content is fairly long and detailed.\n\n"
            "1.1 Background\n"
            "The background covers many aspects, such as technology and demand.\n\n"
            "1.2 Significance\n"
            "The significance lies in improving efficiency.\n\n"
            "Key: this part is very important.\n"
            "Definition: a model is a set of parameters.\n\n"
            "- Apple\n"
            "- Banana\n"
            "- Orange\n\n"
            "2. Methods\n"
            "The method is the core. Continue writing to test paragraph merging, "
            "this is English and should keep space.\n"
            "\n"
            "3. Comparison\n"
            "| Method | Accuracy | Speed |\n"
            "|:---|:---:|---:|\n"
            "| Method A | 92% | Fast |\n"
            "| Method B | 88% | Medium |\n"
            "\n"
            "4. Key Formulas\n"
            "Euler's identity $e^{i\\pi}+1=0$ is often called the most beautiful equation.\n"
            "The mass-energy equivalence is:\n"
            "$$\n"
            "E = mc^2\n"
            "$$\n"
            "A summation can be written as $S = \\sum_{k=1}^{n} a_k$.\n"
            "\n"
            "5. Literature Review\n"
            "Prior work showed the method works[1], later extended[2,3]. "
            "A survey discussed its limitations[4-6].\n"
            "\n"
            "References\n"
            "[1] Author. Paper title[J]. Journal, 2020, 40(2): 123-130.\n"
            "[2] Author A, Author B. Title[C]. Conference, 2021.\n"
        ),
    },
}


def t(key: str, lang: str = "zh") -> str:
    """按语言取文案，缺失时回退到 key 本身。"""
    return STRINGS.get(lang, STRINGS["zh"]).get(key, key)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量替换 Word、Excel 与纯文本中的文本内容
支持 .docx、.xlsx 与 .txt 文件；文档/表格保持原有格式，.txt 保持原有编码不变
作者: YehOlt
"""

# 版本信息（统一在此维护）
VERSION = "2.0.1"

# 打包完整性自检清单（`--selfcheck` 用）：exe 启动前逐个 import，漏带即报。
# 清单必须覆盖源码里出现过的全部顶层 import 名——自检脚本
# （_docs/BatchReplace_selftest.py，经 exe --selftest 执行）会用 AST 比对，
# 新增 import 而忘了更新这里，源码版自检直接失败，不会等到冻结版才炸。
# encodings 两项不在源码 import 里，但 .txt 分支按编码读写文件时要用，
# PyInstaller 不会主动打包非 ASCII 编码，一并列入。
SELFCHECK_IMPORTS = (
    "PIL.Image", "PIL.ImageDraw", "PIL.ImageTk",
    "base64", "ctypes", "datetime", "docx", "hashlib", "json", "openpyxl", "os", "pathlib",
    "queue", "re", "shutil", "subprocess", "sys", "tempfile", "threading",
    "tkinter", "urllib.error", "urllib.request", "webbrowser",
    "win32com.client", "win32com",
    "zipfile", "winreg",
    "encodings.gbk", "encodings.utf_8_sig", "encodings.idna",
)

# 勾选图标的终身保管箱：TTK 元素与 Treeview 的 item 只记住 Tcl 图像名，一旦
# 持有 PhotoImage 的对象先被回收，图像就没了，重绘时报 "image xxx doesn't
# exist"。图标必须与解释器同寿，故一律登记在此，不随实例走。
_CHECK_ICON_KEEPALIVE = []

# 发布页地址（「关于」里的入口）。公开仓库的发行页，换仓库只改这一处。
RELEASE_PAGE_URL = "https://github.com/YehOlt/BatchReplace/releases"

# 更新检查入口：只维护上面那一个发布页地址，API / raw 全部从它推导。
# 清单放在仓库根目录 update.json：主路走 API contents（国内可达），raw 做备胎；
# 再降级 releases/latest、tags。不挂独立分支/标签，仓库的分支与标签页只留版本本身。
_REPO_ROOT = RELEASE_PAGE_URL.rsplit("/releases", 1)[0]
_API_ROOT = _REPO_ROOT.replace("github.com/", "api.github.com/repos/")
UPDATE_API = _API_ROOT + "/releases/latest"
TAGS_API = _API_ROOT + "/tags"
UPDATE_MANIFEST = _REPO_ROOT.replace(
    "github.com/", "raw.githubusercontent.com/") + "/refs/heads/main/update.json"
MANIFEST_API = _API_ROOT + "/contents/update.json?ref=main"

# 项目起始年份（「关于」页的年份区间左端，右端取运行当年）。
START_YEAR = 2026

# 作者与联系方式（「关于」里显示）。集中在此，理由同 RELEASE_PAGE_URL。
AUTHOR_NAME = "YehOlt"
CONTACT_QQ = "16140220"
CONTACT_EMAIL = "yelintaoooo@outlook.com"

# 内容替换模式受理的扩展名，与 _engine_for 的分支一一对应（新增格式两处同改）。
SUPPORTED_EXTS = (".docx", ".xlsx", ".txt")

# 界面文案（中英双语）。控件文案用 Title Case，弹窗正文用 Sentence case。
# 切换语言时刷新所有控件；messagebox 下次弹出用新语言。带 {var} 的是动态模板。
TEXTS = {
    # 工具栏按钮
    "add_files": {"zh": "＋ 添加文件", "en": "+ Add Files"},
    "add_folder": {"zh": "＋ 添加文件夹", "en": "+ Add Folder"},
    "remove": {"zh": "移除", "en": "Remove"},
    "clear": {"zh": "清空", "en": "Clear"},
    "about": {"zh": "关于", "en": "About"},
    "start_search": {"zh": "开始搜索", "en": "Start Search"},
    "enable_backup": {"zh": "自动备份", "en": "Auto Backup"},
    "backup_path_label": {"zh": "备份路径", "en": "Backup Path"},
    "browse_backup": {"zh": "选择备份文件夹", "en": "Select Backup Folder"},
    "open_backup_folder_btn": {"zh": "打开备份文件夹", "en": "Open Backup Folder"},
    "rollback_btn": {"zh": "回溯", "en": "Rollback"},
    # 参数区
    "find_content": {"zh": "查找内容", "en": "Find What"},
    "replace_with": {"zh": "替换为", "en": "Replace With"},
    "find_filename": {"zh": "查找文件名", "en": "Find Filename"},
    "new_filename": {"zh": "新文件名", "en": "New Filename"},
    "find_unique": {"zh": "查找内容", "en": "Find What"},
    # 条件表右键菜单：重复项会在确认时弹窗拦截，故表头不再标「不可重复」
    "mi_insert": {"zh": "插入", "en": "Insert"},
    "mi_move_top": {"zh": "移至顶部", "en": "Move to Top"},
    "mi_move_up": {"zh": "上移一格", "en": "Move Up"},
    "mi_move_down": {"zh": "下移一格", "en": "Move Down"},
    "mi_move_bottom": {"zh": "移至底部", "en": "Move to Bottom"},
    "mi_delete": {"zh": "删除", "en": "Delete"},
    "add_cond": {"zh": "添加", "en": "Add"},
    "edit_cond": {"zh": "修改", "en": "Edit"},
    "delete_cond": {"zh": "删除", "en": "Delete"},
    "clear_cond": {"zh": "清空", "en": "Clear"},
    "mode_replace_content": {"zh": "当前：替换文件内容", "en": "Mode: Replace Content"},
    "mode_replace_filename": {"zh": "当前：重命名文件（不改内容）",
                              "en": "Mode: Rename Files (Content Unchanged)"},
    "match_case": {"zh": "匹配大小写", "en": "Match Case"},
    "replace_filename_opt": {"zh": "仅文件名", "en": "Filenames Only"},
    "multi_condition": {"zh": "多条件替换", "en": "Multi-Condition"},
    # 多条件下拉框：两种执行方式。措辞的关键是"在哪段文本上找"——
    # 「操作 / 记录」一类词都会让人误会执行顺序。
    "multi_mode_parallel": {"zh": "并行替换", "en": "Parallel"},
    "multi_mode_chain": {"zh": "链式替换", "en": "Chained"},
    # 悬停下拉框时按当前执行方式弹出对应说明。每种拆三段渲染：标题（执行
    # 方式名，加粗）、说明（定宽换行）、示例（等宽字体整块缩进），末段补
    # 结论或适用场景。分段只为版面工整，措辞与旧版逐句一致。
    "multi_mode_tip_parallel": {
        "zh": "所有规则在同一轮扫描里一起生效，每条都只认原始文本，互不干扰。",
        "en": "All rules take effect together in a single scan — each matches "
              "only the original text and they do not interfere."},
    "multi_mode_tip_parallel_example": {
        "zh": "规则：a→b，b→c\n原文：ab\n一轮扫完：a 变 b，原位置的 b 变 c\n结果：ab → bc",
        "en": "rules: a→b, b→c\ninput: ab\none pass: a→b, original b→c\nresult: ab → bc"},
    "multi_mode_tip_parallel_note": {
        "zh": "第 1 条刚改出的内容，不会被后面的规则再命中。",
        "en": "Text produced by an earlier rule is never re-matched by a later rule."},
    "multi_mode_tip_chain": {
        "zh": "规则按你排的顺序一条接一条执行，上一条替换完的文本，直接作为下一条的查找对象。",
        "en": "Rules run one by one in your chosen order; the text left by each "
              "rule becomes the input for the next."},
    "multi_mode_tip_chain_example": {
        "zh": "规则：a→b，b→c\n原文：ab\n第1轮：a→b，ab 变成 bb\n第2轮：b→c，bb 变成 cc\n结果：ab → bb → cc",
        "en": "rules: a→b, b→c\ninput: ab\nround 1: a→b, ab becomes bb\n"
              "round 2: b→c, bb becomes cc\nresult: ab → bb → cc"},
    "multi_mode_tip_chain_note": {
        "zh": "越往后执行的规则，越能“看到”前面改出来的新内容。规则有交叉、需要接力时选它；拿不准就用并行。",
        "en": "Later rules can see the new text produced by earlier ones. "
              "Use this when rules overlap or need to chain; if unsure, use parallel."},
    "chain_round": {"zh": "第 {n} 轮 · 「{f}」→「{r}」", "en": 'Round {n} · "{f}" → "{r}"'},
    # 位置列后缀：这条命中落在上一轮写出来的文本里（链式才有，并行不会出现）
    "loc_from_prev": {"zh": "上一轮产出", "en": "from previous round"},
    "hint_no_jump_row": {"zh": "这一处要等第 {n} 轮才出现，文件里现在还没有——双击不跳转",
                         "en": "This one only appears at round {n}; it is not in the file yet — double-click will not jump"},
    "chain_hint": {"zh": "链式模式按整个文件勾选（轮次连在一起，不能只换其中一轮）"
                         " · 带底色的是上一轮改出来的文本",
                   "en": "Chained mode selects whole files (rounds are linked; single rounds "
                         "cannot be replaced) · highlighted rows are text produced by the previous round"},
    "confirm_line_chain": {"zh": "链式替换（{n} 条，从上到下依次执行，上一条的结果是下一条的输入）：",
                           "en": "Chained ({n} conditions, run top to bottom — each searches the previous result):"},
    "lg_batch_start_chain": {"zh": "开始链式替换（{n} 条 · 从上到下依次执行）: {summary}",
                             "en": "Chained replace started ({n} condition(s), top to bottom): {summary}"},
    "history": {"zh": "查找历史", "en": "Find History"},
    "load_rules": {"zh": "加载规则", "en": "Load Rules"},
    "save_rules": {"zh": "保存规则", "en": "Save Rules"},
    # 文件列表表头
    "col_name": {"zh": "名称", "en": "Name"},
    "col_type": {"zh": "类型", "en": "Type"},
    "col_size": {"zh": "大小", "en": "Size"},
    "col_match": {"zh": "匹配", "en": "Matches"},
    "col_status": {"zh": "状态", "en": "Status"},
    # 日志/状态栏
    "process_log": {"zh": "日志", "en": "Log"},
    "ready": {"zh": "就绪", "en": "Ready"},
    # 历史对话框
    "col_cond_group": {"zh": "条件组", "en": "Condition Set"},
    "col_case": {"zh": "区分大小写", "en": "Case Sensitive"},
    "col_target": {"zh": "作用对象", "en": "Target"},
    "col_time": {"zh": "时间", "en": "Time"},
    "select": {"zh": "选择", "en": "Select"},
    "delete": {"zh": "删除", "en": "Delete"},
    "clear_all": {"zh": "清空所有", "en": "Clear All"},
    "close": {"zh": "关闭", "en": "Close"},
    # 预览窗
    "find": {"zh": "查找", "en": "Find"},
    "select_all": {"zh": "全选", "en": "Select All"},
    "deselect_all": {"zh": "全不选", "en": "Deselect All"},
    "expand_all": {"zh": "全部展开", "en": "Expand All"},
    "collapse_all": {"zh": "全部折叠", "en": "Collapse All"},
    "cancel": {"zh": "取消", "en": "Cancel"},
    "col_file": {"zh": "文件", "en": "File"},
    "col_check": {"zh": "勾选", "en": "Check"},
    "col_filename_preview": {"zh": "文件名替换预览", "en": "Filename Preview"},
    "col_filename": {"zh": "文件名", "en": "Filename"},
    "col_matches": {"zh": "匹配数", "en": "Matches"},
    "col_location": {"zh": "位置", "en": "Location"},
    "col_context": {"zh": "上下文", "en": "Context"},
    # 通用按钮
    "ok": {"zh": "确定", "en": "OK"},
    # 弹窗标题
    "t_missing_dep": {"zh": "依赖缺失提示", "en": "Missing Dependencies"},
    "t_warning": {"zh": "警告", "en": "Warning"},
    "t_notice": {"zh": "提示", "en": "Notice"},
    "t_error": {"zh": "错误", "en": "Error"},
    "t_confirm_clear": {"zh": "确认清空", "en": "Confirm Clear"},
    "t_file_exists": {"zh": "文件已存在", "en": "File Exists"},
    "t_success": {"zh": "成功", "en": "Success"},
    "t_confirm_delete": {"zh": "确认删除", "en": "Confirm Delete"},
    "t_confirm": {"zh": "确认", "en": "Confirm"},
    "t_done": {"zh": "完成", "en": "Done"},
    "t_jump_failed": {"zh": "跳转失败", "en": "Jump Failed"},
    "t_overlap_blocked": {"zh": "重叠匹配，替换已拦下", "en": "Overlapping Matches — Replace Blocked"},
    "t_overlap_stopped": {"zh": "重叠匹配，已整批停下", "en": "Overlapping Matches — Batch Stopped"},
    # 弹窗正文（Sentence case，完整句子）
    "m_find_empty": {"zh": "查找内容不能为空！", "en": "Find content cannot be empty."},
    "m_select_to_edit": {"zh": "请先选中要修改的条件！", "en": "Please select a condition to edit."},
    "m_select_to_delete": {"zh": "请先选中要删除的条件！", "en": "Please select a condition to delete."},
    "cond_table_caption": {"zh": "拖动行可调整顺序（链式模式下，从上到下即执行次序）",
                           "en": "Drag rows to reorder (in chain mode, order is execution order)"},
    "m_confirm_clear_cond": {"zh": "确定清空全部 {n} 条条件？", "en": "Clear all {n} conditions?"},
    "m_backup_invalid": {"zh": "备份路径不存在，请先选择有效路径！", "en": "Backup path does not exist. Please select a valid path."},
    "m_cannot_open_folder": {"zh": "无法打开文件夹:\n{e}", "en": "Cannot open folder:\n{e}"},
    "m_cond_issue_save": {"zh": "条件表有问题，未保存：\n\n{errors}", "en": "Condition table has issues, not saved:\n\n{errors}"},
    "m_add_cond_first": {"zh": "请先添加替换条件！", "en": "Please add a replace condition first."},
    "m_enter_find": {"zh": "请输入查找内容！", "en": "Please enter find content."},
    "m_rule_exists": {"zh": "规则文件已存在：\n{filename}\n\n是否覆盖？", "en": "Rule file exists:\n{filename}\n\nOverwrite?"},
    "m_rule_saved": {"zh": "规则已保存为：\n{filename}", "en": "Rule saved as:\n{filename}"},
    "m_save_rule_failed": {"zh": "保存规则失败:\n{e}", "en": "Failed to save rule:\n{e}"},
    "m_load_rule_failed": {"zh": "加载规则失败:\n{e}", "en": "Failed to load rule:\n{e}"},
    "m_select_cond_group": {"zh": "请选择一组条件！", "en": "Please select a condition set."},
    "m_record_gone": {"zh": "该记录已不存在，请重新选择！", "en": "Record no longer exists. Please re-select."},
    "m_select_to_del_record": {"zh": "请选择要删除的记录！", "en": "Please select a record to delete."},
    "m_confirm_delete": {"zh": "确定删除{what}？", "en": "Delete {what}?"},
    "m_history_empty": {"zh": "查找历史已为空！", "en": "Find history is empty."},
    "m_clear_all_history": {"zh": "确定清空所有查找历史？", "en": "Clear all find history?"},
    "m_searching": {"zh": "正在搜索中，请稍候...", "en": "Searching, please wait..."},
    "m_replacing": {"zh": "正在替换中，请稍候...", "en": "Replacing, please wait..."},
    "m_select_files": {"zh": "请先选择文件！", "en": "Please select files first."},
    "m_cond_issue_search": {"zh": "条件表有问题，未开始搜索：\n\n{errors}", "en": "Condition table has issues, search not started:\n\n{errors}"},
    "m_add_cond_click": {"zh": "请先添加替换条件（点「添加」）！", "en": "Please add a condition first (click \"Add\")."},
    "m_no_match": {"zh": "未找到匹配内容！", "en": "No matches found."},
    "m_select_match": {"zh": "请至少选择一条匹配记录！", "en": "Please select at least one match."},
    "m_unsupported_type": {"zh": "不支持的文件类型: {ext}", "en": "Unsupported file type: {ext}"},
    "m_txt_jump": {"zh": "已打开文件（记事本无法精确定位）\n\n匹配位置: {loc}", "en": "File opened (Notepad can't jump precisely).\n\nLocation: {loc}"},
    "loc_text": {"zh": "第 {line} 行, 第 {col} 列", "en": "Line {line}, Col {col}"},
    "e_text_read": {"zh": "读取失败: {e}", "en": "Read failed: {e}"},
    "m_opened_no_loc": {"zh": "已打开文件（精确定位不可用）\n\n原因: {e}", "en": "File opened (precise location unavailable).\n\nReason: {e}"},
    "m_jump_failed": {"zh": "无法打开文件:\n{e}\n\n请确保文件存在且未被占用。", "en": "Cannot open file:\n{e}\n\nPlease ensure the file exists and is not in use."},
    # 关于窗口
    "about_title": {"zh": "关于", "en": "About"},
    "about_app": {"zh": "批量替换工具 v{ver}", "en": "BatchReplace v{ver}"},
    "about_years": {"zh": "{start} - {end} {name}", "en": "{start} - {end} {name}"},
    "about_free": {"zh": "本程序永久免费。", "en": "This program is free forever."},
    "about_note": {"zh": "如有任何建议或发现 Bug，欢迎随时联系我。",
                   "en": "Suggestions and bug reports are always welcome."},
    "about_no_warranty": {"zh": "本程序不作任何担保。详情见 MIT 许可协议。",
                          "en": "This program comes with absolutely no warranty. "
                                "See the MIT License (MIT) for details."},
    "about_release": {"zh": "发布页", "en": "Releases"},
    "check_update": {"zh": "检查更新", "en": "Check for Updates"},
    "t_update_found": {"zh": "发现新版本", "en": "New Version Available"},
    "lg_update_found": {"zh": "发现新版本 {new}，当前版本 {cur}",
                        "en": "New version {new} found (current: {cur})"},
    "lg_update_none": {"zh": "未发现新版本，当前版本 {cur}",
                       "en": "No new version found (current: {cur})"},
    "lg_update_fail": {"zh": "检查更新失败：{e}", "en": "Update check failed: {e}"},
    "m_update_found_short": {"zh": "发现新版本 {new}，当前版本 {cur}。",
                             "en": "New version {new} available (current: {cur})."},
    "l_update_choose": {"zh": "\n请选择处理方式：", "en": "\nChoose an action:"},
    "b_update_download": {"zh": "直接下载", "en": "Download now"},
    "b_update_page": {"zh": "发布页", "en": "Release page"},
    "b_update_skip": {"zh": "暂不更新", "en": "Not now"},
    "lg_update_download": {"zh": "选择直接下载：{new}", "en": "Chose download: {new}"},
    "lg_update_page": {"zh": "选择前往发布页：{url}", "en": "Chose release page: {url}"},
    "lg_update_skip": {"zh": "选择暂不更新：{new}", "en": "Chose skip: {new}"},
    "lg_update_progress": {"zh": "下载中：{done} / {total} 字节",
                           "en": "Downloading: {done} / {total} bytes"},
    "lg_update_dl_fail": {"zh": "下载或校验失败：{e}", "en": "Download or verification failed: {e}"},
    "lg_update_verified": {"zh": "下载完成并通过 SHA-256 校验：{path}",
                           "en": "Downloaded and SHA-256 verified: {path}"},
    "t_update_dl_fail": {"zh": "更新失败", "en": "Update failed"},
    "m_update_dl_fail": {"zh": "下载或校验失败：{e}", "en": "Download or verification failed: {e}"},
    "m_update_launch_fail": {"zh": "新版已下载并校验，但启动失败：{e}\n文件：{path}",
                             "en": "Update verified but failed to launch: {e}\nFile: {path}"},
    "t_update_done": {"zh": "更新完成", "en": "Update complete"},
    "m_update_done": {"zh": "新版 {new} 已下载并启动。\n\n新版文件已保存到同一目录，旧版文件保留不变。\n本窗口随后关闭，由新版接管。",
                      "en": "Version {new} has been downloaded and launched.\n\nThe new exe is in the same folder; the old file is kept.\nThis window will close and hand over to the new version."},
    "m_theme_close_preview": {"zh": "请先关闭替换预览窗口，再切换主题。",
                              "en": "Close the preview window before switching themes."},
    "contact": {"zh": "QQ: {qq}  |  Email: {email}", "en": "QQ: {qq}  |  Email: {email}"},
    # 语言切换按钮：中文态显示 English（提示切到英文），英文态显示 中文
    "lang_to_en": {"zh": "English", "en": "English"},
    "lang_to_zh": {"zh": "中文", "en": "中文"},

    # ================= 以下为双语补充（界面其余可见文案） =================
    # 通用
    "yes": {"zh": "是", "en": "Yes"},
    "no": {"zh": "否", "en": "No"},
    # 状态栏 / 文件列表状态列
    "status_summary": {"zh": "共 {n} 个文件 · 选中 {sel} · 匹配 {m} 处 · {backup}",
                       "en": "{n} files · {sel} selected · {m} matches · {backup}"},
    "backup_on": {"zh": "备份开启", "en": "Backup On"},
    "backup_off": {"zh": "备份关闭", "en": "Backup Off"},
    "backup_no_path": {"zh": "备份开启 · 未设路径", "en": "Backup On · No Path"},
    "st_not_searched": {"zh": "未搜索", "en": "Not Searched"},
    "st_pending": {"zh": "待替换", "en": "To Replace"},
    "st_no_match": {"zh": "无匹配", "en": "No Match"},
    "st_replaced": {"zh": "已替换", "en": "Replaced"},
    # 进度（状态栏）
    "progress_searching": {"zh": "正在搜索", "en": "Searching"},
    "progress_replacing": {"zh": "正在替换", "en": "Replacing"},
    # 模式提示的四种组合
    "mode_replace_content_multi": {"zh": "当前：多条件替换文件内容",
                                   "en": "Mode: Multi-Condition Replace Content"},
    "mode_replace_filename_multi": {"zh": "当前：多条件重命名文件",
                                    "en": "Mode: Multi-Condition Rename Files"},
    # 文件选择对话框
    "fd_select_backup": {"zh": "选择备份文件夹", "en": "Select Backup Folder"},
    "fd_select_files": {"zh": "选择文件", "en": "Select Files"},
    "fd_select_docs": {"zh": "选择 Word、Excel 或文本文件", "en": "Select Word, Excel or Text Files"},
    "fd_select_folder": {"zh": "选择文件夹", "en": "Select Folder"},
    "fd_load_rule": {"zh": "加载替换规则", "en": "Load Replace Rules"},
    "ft_supported": {"zh": "支持的文件（Word / Excel / 文本）", "en": "Supported Files (Word / Excel / Text)"},
    "ft_word": {"zh": "Word 文档", "en": "Word Documents"},
    "ft_excel": {"zh": "Excel 文档", "en": "Excel Documents"},
    "ft_text": {"zh": "纯文本文档", "en": "Text Documents"},
    "ft_all": {"zh": "所有文件", "en": "All Files"},
    "ft_json": {"zh": "JSON 文件", "en": "JSON Files"},
    # 历史对话框（动态文案）
    "history_title": {"zh": "查找历史 · 最近 {n} 组", "en": "Find History · Latest {n}"},
    "history_header": {"zh": "最近用过的查找 / 替换条件，选中即可载入复用（最多 {n} 组）",
                       "en": "Find / replace conditions you used recently — select one to load it back (up to {n})"},
    "history_note": {"zh": "只存查找条件，不记录替换结果；选中即可载入复用",
                     "en": "Only conditions are stored — not what was replaced or how. Select to reuse."},
    "history_group": {"zh": "{n} 条条件", "en": "{n} condition(s)"},
    "target_filename": {"zh": "文件名", "en": "Filename"},
    "target_content": {"zh": "内容", "en": "Content"},
    "del_what_group": {"zh": "整组条件（{n} 条）", "en": "the whole group ({n} conditions)"},
    "del_what_cond": {"zh": "这一条条件：\n\n{f} → {r}", "en": "this condition:\n\n{f} → {r}"},
    # 条件编辑对话框标题
    "t_add_cond": {"zh": "添加条件", "en": "Add Condition"},
    "t_edit_cond": {"zh": "修改条件", "en": "Edit Condition"},
    # 预览窗（动态 / 未在对照表里的）
    "preview_exec": {"zh": "▶  替换勾选项", "en": "▶  Replace Selected"},
    "preview_title": {"zh": "预览替换结果", "en": "Preview Results"},
    "empty_placeholder": {"zh": "（空）", "en": "(empty)"},
    "multi_cond_items": {"zh": "多条件替换 · {n} 项", "en": "Multi-Condition · {n} items"},
    "preview_stats": {"zh": "匹配 {n} 处 · 文件 {m} 个", "en": "{n} matches · {m} files"},
    "checked_count": {"zh": "已选: {n} 项", "en": "Selected: {n}"},
    "hint_bar_filename": {"zh": "勾选「勾选」列可单独排除某个文件，被排除的文件不会重命名",
                          "en": "Uncheck a file in the \"Check\" column to exclude it from renaming"},
    "hint_bar_content": {"zh": "点行首 ▸ 展开文件，查看每一处匹配 · 双击匹配行可跳转到文件对应位置",
                         "en": "Click ▸ to expand a file and see every match · double-click a match "
                               "to jump to its location in the file"},
    "hint_default_filename": {"zh": "勾选「勾选」列可单独排除某个文件",
                              "en": "Uncheck a file to exclude it from renaming"},
    "hint_default_content": {"zh": "双击匹配行可跳转到文件对应位置",
                             "en": "Double-click a match row to jump to its location"},
    "chip_text": {"zh": "↗ 双击跳转", "en": "↗ Double-Click to Jump"},
    "hint_expand_row": {"zh": "▶ 双击此行：展开 / 收起该文件的匹配明细",
                        "en": "▶ Double-click: expand / collapse this file's matches"},
    "hint_jump_row": {"zh": "▶ 双击此行：跳转到 {file}「{loc}」",
                      "en": "▶ Double-click: jump to {loc} in {file}"},
    "parent_hint": {"zh": "▸ 展开可见 {n} 处匹配", "en": "▸ {n} matches hidden — expand to view"},
    "cond_prefix": {"zh": "条件{n}", "en": "Cond {n}"},
    "conflict_gate": {"zh": "⚠ 有 {n} 处重叠匹配，无法替换",
                      "en": "⚠ {n} overlapping match(es) — replace blocked"},
    # 弹窗标题（补充）
    "t_backup_off": {"zh": "警告：未启用备份", "en": "Backup Disabled"},
    "t_backup_failed": {"zh": "备份失败，已中止", "en": "Backup Failed — Aborted"},
    "t_backup_no_path": {"zh": "备份路径为空，已中止", "en": "No Backup Path — Aborted"},
    "t_close": {"zh": "确认关闭", "en": "Confirm Close"},
    # 弹窗正文（补充）
    "m_backup_off": {"zh": "您已关闭「自动备份」。\n\n"
                           "关闭后，此后的替换不再生成备份与回溯记录：原文件无法恢复，"
                           "也无法用「回溯」撤销修改。\n\n"
                           "建议保持「自动备份」开启以保护文件。",
                     "en": "Auto Backup is off.\n\n"
                           "Future replacements will create no backup and no rollback "
                           "record: the originals cannot be restored, and those changes "
                           "cannot be undone with Rollback.\n\n"
                           "Keep Auto Backup on to protect your files."},
    "m_backup_failed": {"zh": "无法创建备份文件夹，为避免文件被无备份修改，替换操作已取消。\n\n"
                              "{e}\n\n请检查备份路径是否可写后重试。",
                        "en": "Cannot create the backup folder. To avoid modifying files without "
                              "a backup, the replace was cancelled.\n\n{e}\n\n"
                              "Please check that the backup path is writable and try again."},
    "m_backup_no_path": {"zh": "已启用备份，但备份路径是空的，无法备份原文件。\n\n"
                               "为保护原文件，替换已取消，未改动任何文件。\n\n"
                               "请先选择备份目录，或关闭「自动备份」后重试。",
                         "en": "Backup is on but the backup path is empty, so the originals "
                               "cannot be backed up.\n\nTo protect your files the replace was "
                               "cancelled and nothing was modified.\n\nPick a backup folder, or "
                               "turn off backup and try again."},
    "m_close_confirm": {"zh": "替换正在进行中。关闭窗口将中断当前批次：\n\n"
                              "· 已处理的文件保持替换后状态\n· 未处理的文件保持原样\n\n确定关闭？",
                        "en": "A replace batch is running. Closing the window ends it:\n\n"
                              "· Files already processed stay replaced\n"
                              "· Files not processed stay unchanged\n\nClose anyway?"},
    "m_rule_invalid": {"zh": "规则文件里的条件不合法，未加载：\n\n{errors}",
                       "en": "The rule file contains invalid conditions and was not loaded:\n\n{errors}"},
    "m_no_valid_condition": {"zh": "没有任何有效条件", "en": "no valid conditions"},
    "m_embedded": {"zh": "以下 Excel 文件包含内嵌对象，替换时需要重写文件，\n"
                         "这些对象可能丢失或变化：\n\n{files}\n\n建议先备份原文件。仍要继续吗？",
                   "en": "The following Excel files contain embedded objects. Replacing rewrites "
                         "the file, and these objects may be lost or changed:\n\n{files}\n\n"
                         "Back up the originals first. Continue anyway?"},
    "m_embedded_item": {"zh": "· {name}（{kinds}）", "en": "· {name} ({kinds})"},
    "m_confirm_replace": {"zh": "确认执行替换操作？", "en": "Run the replace operation?"},
    "confirm_line_find": {"zh": "查找: {v}", "en": "Find: {v}"},
    "confirm_line_replace": {"zh": "替换为: {v}", "en": "Replace with: {v}"},
    "confirm_line_multi": {"zh": "多条件替换（{n} 项，均在原始内容上一次替换）：",
                           "en": "Multi-condition ({n} items, all applied to the original "
                                 "content in one pass):"},
    "confirm_line_backup": {"zh": "备份: {v}", "en": "Backup: {v}"},
    "backup_enabled_word": {"zh": "启用", "en": "enabled"},
    "backup_disabled_word": {"zh": "禁用", "en": "disabled"},
    "confirm_line_rename": {"zh": "重命名文件数: {n}", "en": "Files to rename: {n}"},
    "confirm_line_matches": {"zh": "替换匹配数: {n}（涉及文件 {m} 个）",
                             "en": "Matches to replace: {n} (in {m} files)"},
    "m_done_replace": {"zh": "处理完成！\n成功: {ok} 个\n失败: {fail} 个",
                       "en": "Done!\nSucceeded: {ok}\nFailed: {fail}"},
    "m_dep_missing": {"zh": "检测到以下依赖缺失，对应格式将无法搜索 / 替换：\n\n{items}\n\n"
                            "两个出路：\n1. 使用 exe 版（依赖已打包，无需安装任何环境）\n"
                            "2. 在当前 Python 环境执行：{cmd}",
                      "en": "Missing dependencies detected; the matching formats cannot be "
                            "searched / replaced:\n\n{items}\n\nTwo ways out:\n"
                            "1. Use the exe build (dependencies are bundled; nothing to install)\n"
                            "2. Run this in the current Python environment: {cmd}"},
    "m_dep_item": {"zh": "· {lib}（{ext} 文件不可用）", "en": "· {lib} ({ext} files unavailable)"},
    "m_no_searchable": {"zh": "没有可搜索的文件。\n\n{dep}\n\n"
                              "请使用 exe 版，或在当前 Python 环境安装对应依赖。",
                        "en": "No searchable files.\n\n{dep}\n\nUse the exe build, or install the "
                              "missing dependency in the current Python environment."},
    # 引擎层（纯函数）——校验 / 冲突 / 文件名预检
    "multi_cond_tag": {"zh": "多条件{n}项", "en": "MultiCond{n}"},
    "name_empty": {"zh": "空", "en": "empty"},
    "vf_empty_name": {"zh": "文件名为空", "en": "Filename is empty"},
    "vf_invalid_char": {"zh": "含非法字符 {c}，已替换为下划线",
                        "en": "contains invalid character {c}; replaced with an underscore"},
    "vf_control": {"zh": "含控制字符，已移除", "en": "contains control characters; removed"},
    "vf_reserved": {"zh": "使用了 Windows 保留设备名，已追加下划线",
                    "en": "uses a Windows reserved device name; an underscore was appended"},
    "vf_trailing": {"zh": "已去掉文件名末尾的点/空格（Windows 不允许）",
                    "en": "trailing dots/spaces removed (Windows does not allow them)"},
    "cond_empty": {"zh": "第 {i} 条：查找内容不能为空", "en": "Item {i}: find content cannot be empty"},
    "cond_dup": {"zh": "第 {i} 条：查找内容「{find}」与第 {j} 条重复——查找内容不允许重复",
                 "en": "Item {i}: find content \"{find}\" duplicates item {j} — "
                       "find content must be unique"},
    "conflict_reason_tpl": {"zh": "查找「{hit}」的匹配与「{rival}」重叠",
                            "en": "the match for \"{hit}\" overlaps the match for \"{rival}\""},
    "conflict_note_tpl": {"zh": "{where}：{reason}", "en": "{where}: {reason}"},
    "conflict_block_head": {"zh": "以下位置被多个条件同时匹配：",
                            "en": "These locations are hit by more than one condition:"},
    "conflict_block_more": {"zh": "· 另有 {n} 处，详见日志",
                            "en": "· {n} more — see the log"},
    "conflict_block_tail": {"zh": "同一位置只能由一个条件改写，而没有人知道您想保留哪一条，\n"
                                  "所以本工具不替您挑一个执行——本次替换整批停下，一个文件都不会改动。\n\n"
                                  "处理办法：回到条件表把重叠的那条改掉或删掉；\n"
                                  "也可以在预览里取消勾选冲突的匹配行，冲突消失后按钮自动恢复可用。",
                            "en": "A single location can only be rewritten by one condition, and "
                                  "nothing tells us which one you meant,\nso this tool will not pick "
                                  "one for you — the whole batch stops and no file is changed.\n\n"
                                  "How to fix: go back to the condition table and edit or delete the "
                                  "overlapping condition;\nyou can also uncheck the conflicting match "
                                  "rows in the preview — the button re-enables once no conflict remains."},
    "cond_summary_item": {"zh": "「{f}」→「{r}」", "en": "\"{f}\" → \"{r}\""},
    "cond_summary_sep": {"zh": "、", "en": ", "},
    "cond_summary_more": {"zh": " 等 {n} 项", "en": " (+{n} more)"},
    "fname_empty_skipped": {"zh": "新文件名为空，已跳过", "en": "new filename is empty; skipped"},
    "fname_same_skipped": {"zh": "新文件名与原文件名相同，已跳过",
                           "en": "new filename equals the original; skipped"},
    "fname_dup_skipped": {"zh": "新文件名与 {other} 重复，已跳过（避免互相覆盖）",
                          "en": "new filename duplicates {other}; skipped (would overwrite each other)"},
    "fname_exists_skipped": {"zh": "目标文件 {name} 已存在，已跳过（避免覆盖）",
                             "en": "target file {name} already exists; skipped (would overwrite)"},
    # 引擎层——匹配项「位置」显示串（中文侧与旧版逐字一致，测试按字面断言）
    "loc_paragraph": {"zh": "第{n}段", "en": "para {n}"},
    "loc_table": {"zh": "表格{t} 第{r}行第{c}列", "en": "table {t}, row {r}, col {c}"},
    "loc_header": {"zh": "页眉", "en": "header"},
    "loc_footer": {"zh": "页脚", "en": "footer"},
    "loc_section": {"zh": "第{s}节 {name} 第{p}段", "en": "section {s} {name}, para {p}"},
    # 引擎层——列表拼接与备份命名（都会进文件名 / 显示串，英文界面不该混中文）
    "list_sep": {"zh": "、", "en": ", "},
    "list_more": {"zh": " 等 {n}", "en": " ... ({n} in total)"},
    "bk_content": {"zh": "内容替换", "en": "ContentReplace"},
    "bk_filename": {"zh": "文件名替换", "en": "FilenameRename"},
    # 引擎层——xlsx 内嵌对象种类（备份前的风险提示里会列出）
    "emb_image": {"zh": "内嵌图片", "en": "embedded image"},
    "emb_chart": {"zh": "图表", "en": "chart"},
    "emb_drawing": {"zh": "绘图对象", "en": "drawing"},
    "emb_object": {"zh": "嵌入对象", "en": "embedded object"},
    "e_backup_create": {"zh": "创建备份失败: {e}", "en": "failed to create backup: {e}"},
    "e_backup_missing": {"zh": "备份文件缺失: {name}", "en": "Backup file missing: {name}"},
    "e_no_orig": {"zh": "记录缺少原始文件路径", "en": "Original file path missing from record"},
    "m_mixed_batch": {"zh": "同一批里混了「文件名替换」和「内容替换」，已整批拦下。请分两次搜索分别执行。",
                      "en": "This batch mixes filename and content replacements; stopped. Run them as two separate searches."},
    "e_cloud_placeholder": {"zh": "文件处于 OneDrive 在线占位状态: {name}\n"
                                  "请先在资源管理器里打开它，等下载完成后再回溯",
                            "en": "The file is a OneDrive online placeholder: {name}\n"
                                  "Open it in File Explorer and wait for the download "
                                  "to finish before rolling back"},
    # 引擎层——提取 / 替换时被跳过的说明
    "ex_run_protected": {"zh": "{loc}：所在 run 含分页符/图片/域字符，已跳过（避免破坏文档结构）",
                         "en": "{loc}: the run contains a page break / image / field code; skipped "
                               "(would break the document structure)"},
    "ex_formula_cells": {"zh": "公式单元格 {n} 处已跳过（避免破坏公式）: {shown}",
                         "en": "{n} formula cell(s) skipped (would break the formulas): {shown}"},
    "sk_no_span": {"zh": "匹配项缺少区间信息，已跳过", "en": "match has no span info; skipped"},
    "sk_no_find": {"zh": "匹配项缺少查找词，已跳过", "en": "match has no find text; skipped"},
    "sk_no_paragraph": {"zh": "替换时定位不到对应段落（文件可能已被修改）",
                        "en": "cannot locate the paragraph at replace time (file may have changed)"},
    "sk_no_cell_info": {"zh": "匹配项缺少单元格信息，已跳过",
                        "en": "match has no cell info; skipped"},
    "sk_no_cell": {"zh": "替换时定位不到对应单元格（文件可能已被修改）",
                   "en": "cannot locate the cell at replace time (file may have changed)"},
    "sk_cell_empty": {"zh": "单元格为空，已跳过", "en": "cell is empty; skipped"},
    "sk_formula_cell": {"zh": "公式单元格，已跳过（避免破坏公式）",
                        "en": "formula cell; skipped (would break the formula)"},
    "sk_invalid_span": {"zh": "匹配区间无效，已跳过", "en": "invalid match span; skipped"},
    "sk_run_protected": {"zh": "所在 run 含分页符/图片/域字符，已跳过",
                         "en": "the run contains a page break / image / field code; skipped"},
    "sk_text_changed": {"zh": "替换时文本已变化（文件可能已被修改），已跳过",
                        "en": "text changed before replace (file may have changed); skipped"},
    # 引擎层——跳转时的异常说明（会显示在弹窗正文里）
    "e_win32com": {"zh": "win32com 不可用", "en": "win32com is unavailable"},
    "e_docx_missing": {"zh": "未安装 python-docx 库", "en": "python-docx is not installed"},
    "e_openpyxl_missing": {"zh": "未安装 openpyxl 库", "en": "openpyxl is not installed"},
    "e_unsupported_type": {"zh": "不支持的文件类型: {ext}", "en": "Unsupported file type: {ext}"},
    "e_search_failed": {"zh": "搜索失败 {name}: {e}", "en": "Search failed for {name}: {e}"},
    "e_para_out_of_range": {"zh": "段落索引 {i} 超出范围（共{n}段）",
                            "en": "paragraph index {i} is out of range ({n} paragraphs)"},
    "e_table_coord": {"zh": "表格坐标({r}, {c}) 超出范围",
                      "en": "table coordinate ({r}, {c}) is out of range"},
    "e_table_out_of_range": {"zh": "表格索引 {i} 超出范围（共{n}个表）",
                             "en": "table index {i} is out of range ({n} tables)"},
    "e_hf_para_out_of_range": {"zh": "{kind} 段落索引 {i} 超出范围",
                               "en": "{kind} paragraph index {i} is out of range"},
    "e_section_out_of_range": {"zh": "节索引 {i} 超出范围", "en": "section index {i} is out of range"},
    "e_match_type": {"zh": "不支持的匹配类型: {t}", "en": "Unsupported match type: {t}"},
    "e_sheet_missing": {"zh": "工作表 '{name}' 不存在（可用: {avail}）",
                        "en": "worksheet '{name}' does not exist (available: {avail})"},
    # 工作线程里依赖缺失的错误行——_finish_search 靠依赖缺失标记区分「搜不了」与「没匹配」
    "e_missing_docx": {"zh": "{name}: 未安装 python-docx，.docx 无法搜索"
                             "（请使用 exe 版，或安装：pip install python-docx）",
                       "en": "{name}: python-docx is not installed, so .docx cannot be searched "
                             "(use the exe build, or run: pip install python-docx)"},
    "e_missing_openpyxl": {"zh": "{name}: 未安装 openpyxl，.xlsx 无法搜索"
                                "（请使用 exe 版，或安装：pip install openpyxl）",
                           "en": "{name}: openpyxl is not installed, so .xlsx cannot be searched "
                                 "(use the exe build, or run: pip install openpyxl)"},
    # 运行时提示（启动补报 / 线程内异常）
    "note_config_failed": {"zh": "读取配置失败（已改用默认配置）: {e}",
                           "en": "Failed to read the config (defaults used): {e}"},
    "note_history_failed": {"zh": "读取查找历史失败（已按空历史启动）: {e}",
                            "en": "Failed to read the find history (started empty): {e}"},
    "note_config_save_failed": {"zh": "保存配置失败: {e}", "en": "Failed to save the config: {e}"},
    "note_history_save_failed": {"zh": "保存查找历史失败: {e}",
                                 "en": "Failed to save the find history: {e}"},
    "lg_ui_update_failed": {"zh": "界面更新失败: {e}", "en": "UI update failed: {e}"},
    # 日志正文
    "lg_rule_saved": {"zh": "规则已保存: {path}", "en": "Rule saved: {path}"},
    "lg_cond_moved": {"zh": "条件顺序已调整（链式模式下顺序即执行次序）——若预览窗开着，请重新搜索后再替换",
                      "en": "Condition order changed (in chain mode, order is execution order) — "
                            "if the preview window is open, search again before replacing"},
    "lg_rule_loaded": {"zh": "规则已加载: {path}", "en": "Rule loaded: {path}"},
    "lg_files_added": {"zh": "已添加 {n} 个文件", "en": "Added {n} file(s)"},
    "lg_files_removed": {"zh": "已从清单移除 {n} 个文件", "en": "Removed {n} item(s) from list"},
    "lg_files_cleared": {"zh": "已清空文件列表", "en": "File list cleared"},
    "lg_history_loaded": {"zh": "已载入查找历史：{n} 条条件",
                          "en": "Loaded from find history: {n} condition(s)"},
    "lg_history_deleted": {"zh": "已删除查找历史", "en": "Find history entry deleted"},
    "lg_history_cleared": {"zh": "已清空所有查找历史", "en": "All find history cleared"},
    "lg_searching": {"zh": "正在搜索匹配项，请稍候...", "en": "Searching for matches, please wait..."},
    "lg_search_stale": {"zh": "已丢弃过期的搜索结果（已有更新的搜索在进行）",
                        "en": "Stale search result discarded (a newer search is running)"},
    "lg_no_searchable_dep": {"zh": "无可搜索文件（依赖缺失）: {dep}",
                             "en": "No searchable files (missing dependency): {dep}"},
    "lg_no_match": {"zh": "未找到匹配内容", "en": "No matches found"},
    "lg_found_total": {"zh": "共找到 {n} 处匹配", "en": "Found {n} match(es)"},
    "lg_overlap": {"zh": "重叠匹配: {note}", "en": "Overlapping match: {note}"},
    "lg_cancelled_embedded": {"zh": "已取消：目标文件含内嵌对象，用户选择中止（原文件未做任何改动）",
                              "en": "Cancelled: target files contain embedded objects "
                                    "(no file was modified)"},
    "lg_overlap_abort": {"zh": "重叠匹配 {n} 处，替换整批中止（原文件未做任何改动）",
                         "en": "{n} overlapping match(es) — batch aborted (no file was modified)"},
    "lg_overlap_blocked": {"zh": "重叠阻断: {note}", "en": "Overlap blocked: {note}"},
    "lg_backup_folder": {"zh": "备份文件夹: {path}", "en": "Backup folder: {path}"},
    "lg_backup_failed": {"zh": "备份文件夹创建失败，替换已中止（原文件未做任何改动）: {e}",
                         "en": "Failed to create the backup folder; replace aborted "
                               "(no file was modified): {e}"},
    "lg_fname_precheck_item": {"zh": "文件名预检: {w}", "en": "Filename precheck: {w}"},
    "lg_fname_precheck_sum": {"zh": "文件名预检：{n} 项被修正或跳过（详见上方日志）",
                              "en": "Filename precheck: {n} item(s) corrected or skipped "
                                    "(see log above)"},
    "lg_fname_done": {"zh": "文件名替换完成: {ok} 成功, {fail} 失败",
                      "en": "Filename replace finished: {ok} succeeded, {fail} failed"},
    "lg_batch_start_single": {"zh": "开始批量替换（行级精确）: '{f}' -> '{r}'",
                              "en": "Batch replace started (line-level, precise): '{f}' -> '{r}'"},
    "lg_batch_start_multi": {"zh": "开始批量替换（多条件 {n} 项 · 在原始内容上一次替换）: {summary}",
                             "en": "Batch replace started (multi-condition, {n} item(s) · applied "
                                   "once to the original content): {summary}"},
    "lg_batch_start_plain": {"zh": "开始批量替换（行级精确）",
                             "en": "Batch replace started (line-level, precise)"},
    "lg_file_count": {"zh": "共 {n} 个文件", "en": "{n} file(s)"},
    "lg_backup_state": {"zh": "备份: {v}", "en": "Backup: {v}"},
    "lg_case_state": {"zh": "大小写匹配: {v}", "en": "Match case: {v}"},
    "lg_backup_disabled_risk": {"zh": "禁用（风险操作！）", "en": "disabled (risky!)"},
    "lg_backup_no_path": {"zh": "替换中止：备份已启用但备份路径为空（未改动任何文件）",
                          "en": "Replace aborted: backup is on but the backup path is empty "
                                "(nothing was modified)"},
    "lg_processing": {"zh": "处理: {name}", "en": "Processing: {name}"},
    "lg_skip_warn": {"zh": "  警告: 跳过 {loc} - {err}", "en": "  Warning: skipped {loc} - {err}"},
    "lg_success_one": {"zh": "  成功: {name} ({n} 处替换)",
                       "en": "  Succeeded: {name} ({n} replacements)"},
    "lg_fail_one": {"zh": "  失败: {e}", "en": "  Failed: {e}"},
    "lg_done_summary": {"zh": "处理完成! 成功: {ok}, 失败: {fail}",
                        "en": "Done! Succeeded: {ok}, failed: {fail}"},
    "lg_manifest_written": {"zh": "已记录本次替换（{n} 个文件），可在「回溯」中还原",
                            "en": "Recorded this run ({n} file(s)); restorable via \"Rollback\""},
    "lg_manifest_failed": {"zh": "回溯记录写入失败: {e}（本次替换结果不受影响，但无法回溯）",
                           "en": "Failed to write the rollback record: {e} "
                                "(this run's result is unaffected, but it cannot be rolled back)"},
    # ---- 回溯对话框 ----
    "t_rollback": {"zh": "回溯 - 恢复到替换前", "en": "Rollback — restore to before replacement"},
    "rb_header": {"zh": "选择一次内容替换操作，把它涉及的文件恢复成替换前的样子\n"
                        "（回溯只对记录时的本机路径有效；跨机器或路径已变更无法恢复）",
                  "en": "Pick a content-replacement run to restore its files to "
                        "their pre-replacement state\n"
                        "(Rollback only works for the original paths on this machine; "
                        "it cannot restore across machines or moved paths)"},
    "rb_empty": {"zh": "没有可回溯的记录。回溯只记录内容替换，且需要当时开启了备份",
                 "en": "No restorable records. Rollback only covers content "
                       "replacement, and only when backup was enabled"},
    "rb_col_time": {"zh": "时间", "en": "Time"},
    "rb_col_cond": {"zh": "条件", "en": "Conditions"},
    "rb_col_mode": {"zh": "执行方式", "en": "Mode"},
    "rb_col_files": {"zh": "文件数", "en": "Files"},
    "rb_mode_chain": {"zh": "链式", "en": "Chain"},
    "rb_mode_parallel": {"zh": "并行", "en": "Parallel"},
    "m_rb_pick": {"zh": "请先选中一条记录！", "en": "Please select a record first."},
    "m_rb_confirm": {"zh": "用替换前的文件覆盖当前文件？\n\n"
                           "· 覆盖后，这些文件在该次替换之后的全部改动都会丢失\n"
                           "· 覆盖前会先把当前文件快照到备份夹 rollback_ 子夹，仍可挽回\n"
                           "· 事后手工改过的内容同样会丢失",
                     "en": "Overwrite the current files with the pre-replacement copies?\n\n"
                           "· Every change made to these files after that run will be lost\n"
                           "· The current files are snapshotted to the backup folder's "
                           "rollback_ subfolder first, so this is still recoverable\n"
                           "· Manual edits made afterwards will be lost too"},
    "rb_later_warning": {"zh": "以下文件在更晚的替换操作中也改过，回滚将一并丢失那些改动：\n{files}",
                         "en": "These files were also changed by later replacement runs; "
                               "rolling back will lose those changes too:\n{files}"},
    "t_rb_pick_files": {"zh": "选择要回溯的文件", "en": "Select files to roll back"},
    "rb_pick_header": {"zh": "勾选要恢复的文件（默认全选）", "en": "Check the files to restore (all selected by default)"},
    "rb_col_file": {"zh": "文件", "en": "File"},
    "rb_col_path": {"zh": "原路径", "en": "Original path"},
    "m_rb_pick_none": {"zh": "请至少勾选一个文件！", "en": "Please check at least one file."},
    "lg_rb_start": {"zh": "开始回溯 {n} 个文件…", "en": "Rolling back {n} file(s)…"},
    "lg_rb_one_ok": {"zh": "  已恢复: {path}", "en": "  Restored: {path}"},
    "lg_rb_one_fail": {"zh": "  失败: {path} - {e}", "en": "  Failed: {path} - {e}"},
    "lg_rb_done": {"zh": "回溯完成：成功 {ok}，失败 {fail}",
                   "en": "Rollback complete: {ok} succeeded, {fail} failed"},
    "m_rb_result": {"zh": "回溯完成：成功 {ok} 个，失败 {fail} 个\n\n当前文件已快照在备份夹 rollback_ 子夹",
                    "en": "Rollback complete: {ok} succeeded, {fail} failed\n\n"
                          "The replaced files were snapshotted in the backup "
                          "folder's rollback_ subfolder"},
    "rb_missing_now": {"zh": "原文件已不存在，跳过", "en": "Original file no longer exists; skipped"},
    "lg_manifest_problem": {"zh": "有 {n} 个备份夹的记录文件损坏，已跳过",
                            "en": "{n} backup folder(s) had unreadable records and were skipped"},
    "lg_jump_fallback": {"zh": "跳转失败，使用降级方案: {e}",
                         "en": "Jump failed; falling back to opening the file: {e}"},
    "lg_caret_skip": {"zh": "  查找词含 ^，跳过词级高亮（避免与 Word 特殊代码混用）",
                      "en": "  Find text contains ^; skipping word-level highlighting "
                            "(would conflict with Word special codes)"},
    "lg_highlight_failed": {"zh": "  高亮失败（不影响跳转）: {e}",
                            "en": "  Highlighting failed (jump unaffected): {e}"},
}
def _system_ui_lcid():
    """Windows 用户界面语言的 LCID。抽成独立函数是为了让测试能替换它。

    import 必须留在函数内：本文件在导入期就要算 DEFAULT_LANG（见下方），
    那时模块级的 import 块还没执行到，拿不到模块级 ctypes。
    """
    import ctypes
    return ctypes.windll.kernel32.GetUserDefaultUILanguage()


def lang_from_lcid(lcid):
    """Windows LCID → 界面语言：中文系主语言 ID 为 0x04。

    抽成纯函数也是为了让判据可测——真取系统语言的 API 在测试里不好造。
    """
    return "zh" if lcid & 0x3FF == 0x04 else "en"


def detect_default_lang():
    """首次启动的默认语言（中文系统出中文，其余出英文）；仅在没有已保存配置时使用。

    取用户界面语言而非 locale：locale 受环境变量影响，中文 Windows 上也可能报出 en_US。
    """
    try:
        return lang_from_lcid(_system_ui_lcid())
    except Exception:
        pass
    # 非 Windows 或 API 不可用时退回区域环境变量（本工具只发 Windows 版，仅兜底）。
    # 直接读环境变量而非 locale.getdefaultlocale()：后者已废弃，读的也是这几个变量。
    try:
        import os
        tag = (os.environ.get("LC_ALL") or os.environ.get("LC_CTYPE")
               or os.environ.get("LANG") or "")
        return "zh" if tag.lower().startswith("zh") else "en"
    except Exception:
        return "zh"


DEFAULT_LANG = detect_default_lang()


def detect_default_theme():
    """系统当前的应用主题（浅色 / 深色），读不到按浅色。

    取的是「应用」主题而非「Windows 主题」：任务栏 / 开始菜单用的是后者，
    而本工具的窗口属于应用范畴。每次启动重新探测、不落盘——切换就在工具栏
    上，跟系统走比记住上次的手动选择更省心，也不会出现"系统改了、工具还停在
    旧主题"的半截状态。
    """
    import os
    if os.name != "nt":
        return "light"
    try:
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            light, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "light" if light else "dark"
    except Exception:
        # 键不存在（老系统 / 精简镜像）或值读不出来，一律按浅色，不影响启动
        return "light"

# 当前界面语言。放在模块级而不是只挂在实例上：校验、冲突说明这些纯函数在产品
# 逻辑区（界面区之前），它们也要出中文/英文文案，只能读模块级状态。
CURRENT_LANG = DEFAULT_LANG


def T(key, **kw):
    """取当前语言的文案，支持 {var} 格式化。找不到时回退中文再回退 key 本身。"""
    entry = TEXTS.get(key, {})
    s = entry.get(CURRENT_LANG) or entry.get("zh") or key
    return s.format(**kw) if kw else s

# 时间戳写出格式（查找历史、替换规则共用）。把单元格日期读回来另有容错格式
# 列表（见 restore_cell_value），两者不是一回事，不要合并。
TS_FMT = "%Y-%m-%d %H:%M:%S"

import tkinter as tk
import tkinter.font as tkfont
from PIL import Image, ImageDraw, ImageTk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import ctypes
import os
import queue
import sys
import threading
import webbrowser
from pathlib import Path
from datetime import datetime, date, time as time_type
import base64
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.request
import encodings.idna  # HTTPS 的 set_tunnel 会 encode('idna')，不静态 import 冻结版会缺
import zipfile

# 冻结版自检脚本会反复建窗又立刻销毁，启动静默检查的 pending 回调会在窗销毁后触发。
SKIP_STARTUP_UPDATE = "--selftest" in sys.argv


# ---------- 应用图标 ----------
# 标题栏左上角与任务栏按钮都从窗口图标取图，而 LoadImage 只认文件，
# 所以图标要以 .ico 落盘一次（_app_icon_path 负责缓存路径）。
_ICON_PATH = {}


def _render_app_icon(px):
    """画应用图标：圆角蓝底 + 白色双箭头（替换）。1024 画布超采样，小尺寸不毛边。"""
    size = 1024
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((24, 24, size - 24, size - 24), radius=224, fill="#2563EB")
    half, head = 56, 120          # 箭头杆半高 / 箭尖半高
    d.rectangle((184, 384 - half, 704, 384 + half), fill="#FFFFFF")
    d.polygon([(704, 384 - head), (840, 384), (704, 384 + head)], fill="#FFFFFF")
    d.rectangle((320, 640 - half, 840, 640 + half), fill="#FFFFFF")
    d.polygon([(320, 640 - head), (184, 640), (320, 640 + head)], fill="#FFFFFF")
    return img.resize((px, px), Image.LANCZOS)


def _app_icon_path():
    """把图标写成临时 .ico 并返回路径（同进程只写一次）。

    LoadImage 只认文件或资源、不认内存字节，所以必须先落盘；文件留在临时目录
    直到进程退出，重试加载时不必重画。
    """
    if "ico" in _ICON_PATH:
        return _ICON_PATH["ico"]
    path = None
    try:
        fd, path = tempfile.mkstemp(suffix=".ico", prefix="BatchReplace_")
        os.close(fd)
        _render_app_icon(256).save(
            path, format="ICO",
            sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    except Exception:
        path = None
    _ICON_PATH["ico"] = path
    return path


def _set_window_icon(win):
    """给窗口挂应用图标（标题栏左上角与任务栏按钮都从这里取图）。

    iconbitmap 的 default 形式一次设给解释器，之后新建的 Toplevel 直接继承；
    以前无边框窗口每次重建 HWND 都要手工 SendMessage(WM_SETICON) 补一遍，
    窗口交还系统后这条链子整根不需要了。
    """
    path = _app_icon_path()
    if not path:
        return
    try:
        win.iconbitmap(default=path)
    except Exception:
        # 图标是锦上添花：任何一步失败都不该拦住主界面
        pass


# 工作区矩形（最大化用，SPI_GETWORKAREA 的输出缓冲区）。
# 本地定义而不用 ctypes.wintypes.RECT：冻结构建只认 import 语句，
# 代码里仅属性访问 ctypes.wintypes 时它不会被打进 exe。
class _WorkArea(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


# ---------- 任务栏身份：无边框窗口在任务栏上的位置与图标 ----------
# 无边框窗口在 Windows 上是 WS_POPUP + WS_EX_TOOLWINDOW，系统据此把它从任务栏
# 摘掉：窗口开着的时候任务栏没有它的位置，只有最小化那一下因为临时交出无边框
# 标志才露一次脸。改成 APPWINDOW、并显式给窗口设图标，任务栏才会从启动那一刻
# 就带着图标出现——WM_GETICON 全为零时任务栏只能画一块空白。
_ICON_KEEPALIVE = {}


def _apply_taskbar_identity(root):
    """把根窗口登记成任务栏应用窗口，并设上大 / 小两号图标。

    overrideredirect 一变（最小化、还原）Tk 就重建 HWND，扩展样式与图标随之
    丢失，所以每次重建后都要再调一次。
    """
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        user32.GetAncestor.restype = ctypes.c_void_p
        user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        hwnd = user32.GetAncestor(ctypes.c_void_p(root.winfo_id()), 2)  # GA_ROOT
        if not hwnd:
            return
        user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
        user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
        user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
        user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                             ctypes.c_ssize_t]
        user32.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                        ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                        ctypes.c_int, ctypes.c_uint]
        GWL_EXSTYLE, WS_EX_TOOLWINDOW, WS_EX_APPWINDOW = -20, 0x80, 0x40000
        ex = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        want = (ex & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
        if want != ex:
            user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, want)
            # 扩展样式改完要过一遍 SetWindowPos：任务栏只在 FRAMECHANGED 之后
            # 重建按钮，否则样式位改了、任务栏照旧当它是工具窗。
            user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, 0x0027)
        path = _app_icon_path()
        if not path:
            return
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p,
                                      ctypes.c_uint, ctypes.c_int, ctypes.c_int,
                                      ctypes.c_uint]
        user32.SendMessageW.restype = ctypes.c_void_p
        user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                        ctypes.c_void_p, ctypes.c_void_p]
        for px, which in ((32, 1), (16, 0)):        # ICON_BIG / ICON_SMALL
            key = ("h", px)
            if key not in _ICON_KEEPALIVE:
                _ICON_KEEPALIVE[key] = user32.LoadImageW(
                    None, path, 1, px, px, 0x0010)   # IMAGE_ICON | LR_LOADFROMFILE
            hicon = _ICON_KEEPALIVE[key]
            if hicon:
                user32.SendMessageW(hwnd, 0x0080, ctypes.c_void_p(which),  # WM_SETICON
                                    ctypes.c_void_p(hicon))
    except Exception:
        # 图标是锦上添花：任何一步失败都不该拦住主界面
        pass


def _apply_dialog_titlebar_theme(win, dark):
    """把对话框的原生标题栏翻到与应用主题一致的深浅（纯 DWM 视觉开关）。

    对话框刻意保持系统原生装饰（自绘边框对话框出过前台锁死缺陷），
    但原生栏只认系统深色模式：系统深 + 应用浅时，白界面上顶着一根黑栏。
    这个属性只改栏的深浅，不动窗口结构，与 grab / 前台逻辑无交集。
    属性号 20 是正式值，个别旧版 Win10 只认预发布的 19，两个都不认就放弃。
    """
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        user32.GetAncestor.restype = ctypes.c_void_p
        user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        hwnd = user32.GetAncestor(ctypes.c_void_p(win.winfo_id()), 2)  # GA_ROOT
        if not hwnd:
            return
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    ctypes.c_void_p(hwnd), attr,
                    ctypes.byref(value), ctypes.sizeof(value)) == 0:
                return
    except Exception:
        pass


def _set_app_user_model_id():
    """给进程一个固定身份，任务栏才会把窗口归到它自己名下、用它的图标。"""
    if os.name != "nt":
        return
    try:
        shell32 = ctypes.windll.shell32
        shell32.SetCurrentProcessExplicitAppUserModelID.argtypes = [ctypes.c_wchar_p]
        shell32.SetCurrentProcessExplicitAppUserModelID("YehOlt.BatchReplace")
    except Exception:
        pass


# 匹配 / 替换 / 备份的唯一实现在本文件下方的「核心逻辑区」；工具只有一个
# .py，单元测试直接导入本文件，测的就是打包出去的同一份代码。
# 注释只解释「为什么这样写」和「踩过什么坑」，版本沿革看 git 提交记录。


try:
    from docx import Document
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


try:
    from openpyxl import load_workbook
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# win32com 处理（双击跳转）；显式导入子模块而非 import win32com，确保 PyInstaller 正确打包
WIN32COM_AVAILABLE = False
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    pass


CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".batch_replace_config.json")
HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".batch_replace_history.json")
# 规则目录不随界面语言变：磁盘上的固定数据位置，跟着语言切换等于让老用户
# 切一次就"丢"掉已保存的规则；目录名用英文，避免英文界面下留中文目录。
RULES_DIR = os.path.join(os.path.expanduser("~"), "Documents", "BatchReplace_Rules")
DEFAULT_BACKUP_DIR = os.path.join(os.path.expanduser("~"), "Documents", "BatchReplace_Backups")
# v2.0.0 之前的默认备份目录（中文名）。老配置里存着它，不迁移会被写回并锁死。
LEGACY_DEFAULT_BACKUP_DIR = os.path.join(os.path.expanduser("~"), "Documents", "批量替换备份")
MAX_HISTORY = 10  # 最大历史记录数


def migrate_backup_path(saved):
    """把配置里遗留的旧默认备份目录换成新的默认目录。

    只替换恰好等于旧默认值的路径；比较前 normpath + normcase——tkinter
    目录框返回正斜杠路径，与源码的反斜杠路径对不上就会漏掉一半情况。
    """
    if os.path.normcase(os.path.normpath(saved)) == \
            os.path.normcase(os.path.normpath(LEGACY_DEFAULT_BACKUP_DIR)):
        return DEFAULT_BACKUP_DIR
    return saved


# =====================================================================
# 核心逻辑区（匹配 / 替换 / 备份的唯一实现）
#
# 约定：不依赖 tkinter、可被测试直接导入；界面状态显式传参；跳过项以列表
# 返回；所有写回原文件都走 atomic_save（先写临时文件再原子替换）。
# =====================================================================

INVALID_CHARS = ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]

# Windows 保留设备名（大小写不敏感；带扩展名同样保留，如 NUL.txt）
WINDOWS_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"}
WINDOWS_RESERVED_NAMES.update(f"COM{i}" for i in range(1, 10))
WINDOWS_RESERVED_NAMES.update(f"LPT{i}" for i in range(1, 10))


def _clean_invalid_chars(name):
    """把非法字符替换为下划线，返回 (清洗结果, 实际出现的非法字符)。"""
    found = [c for c in INVALID_CHARS if c in name]
    for char in found:
        name = name.replace(char, "_")
    return name, found


def sanitize_filename(name):
    """取可作文件名用的安全名（复用 validate_filename，避免两套清洗规则）。"""
    return validate_filename(name)[0].strip()


def validate_filename(new_name):
    """校验并修正新文件名，返回 (安全名, [问题说明])。

    不能直接 os.rename：非法字符、保留设备名、末尾点/空格要等 rename 才报
    错，批内撞名也没人拦——动手前先收拾干净，每项改动都记录下来告知用户。
    """
    problems = []
    name = new_name or ""
    if not name.strip():
        return "", [T("vf_empty_name")]
    cleaned, invalid = _clean_invalid_chars(name)
    for char in invalid:
        problems.append(T("vf_invalid_char", c=repr(char)))
    control = [c for c in cleaned if ord(c) < 32]
    if control:
        cleaned = "".join(c for c in cleaned if ord(c) >= 32)
        problems.append(T("vf_control"))
    stem = os.path.splitext(cleaned)[0]
    if stem.upper() in WINDOWS_RESERVED_NAMES:
        cleaned = cleaned + "_"
        problems.append(T("vf_reserved"))
    stripped = cleaned.rstrip(" .")
    if stripped != cleaned:
        cleaned = stripped
        problems.append(T("vf_trailing"))
    return cleaned, problems


def plan_filename_renames(rows):
    """预检整批文件名替换，返回 (plan, warnings)。

    rows: [{"file_path": ..., "new_name": ...}]；plan: 可安全执行的
    [(file_path, new_path)]；warnings: 被修正 / 被跳过的说明。
    批内撞名、目标文件已存在都会跳过——否则「等 os.rename 才炸」或
    「POSIX 上静默覆盖别人的文件」。
    """
    plan, warnings = [], []
    seen_targets = {}
    for row in rows:
        file_path = row["file_path"]
        dir_name = os.path.dirname(file_path)
        old_name = os.path.basename(file_path)
        new_name = row.get("new_name") or ""
        safe_name, problems = validate_filename(new_name)
        for p in problems:
            warnings.append(f"{old_name}: {p}")
        if not safe_name:
            warnings.append(f"{old_name}: {T('fname_empty_skipped')}")
            continue
        if safe_name == old_name:
            warnings.append(f"{old_name}: {T('fname_same_skipped')}")
            continue
        new_path = os.path.join(dir_name, safe_name)
        # Windows 文件名大小写不敏感：A.txt 与 a.txt 是同一个目标，不归一
        # 会双双放行、第二条 rename 才抛 FileExistsError。
        target_key = os.path.normcase(os.path.normpath(new_path))
        if target_key in seen_targets:
            warnings.append(f"{old_name}: " + T(
                "fname_dup_skipped", other=os.path.basename(seen_targets[target_key])))
            continue
        target_exists = os.path.exists(new_path)
        same_file = False
        if target_exists:
            try:
                same_file = os.path.samefile(new_path, file_path)
            except OSError:
                same_file = False
        if target_exists and not same_file:
            warnings.append(f"{old_name}: " + T("fname_exists_skipped", name=safe_name))
            continue
        seen_targets[target_key] = file_path
        plan.append((file_path, new_path))
    return plan, warnings


# =====================================================================
# 多条件替换·并行语义（与单条件共用同一套匹配 / 替换实现，只是条件变成 N 条）
#
# 三条语义，改这里之前先读：
#   1. 所有条件都在「原始内容」上匹配，一次找出全部待改项；a→b 新生成的 b
#      绝不被 b→c 命中。链式模式反过来——上一条的输出是下一条的输入，
#      那是用户显式选择的另一种执行方式，见 replace_chain_files。
#   2. 查找内容不可重复，替换内容可重复（见 validate_conditions）。
#   3. 同一位置被多条命中时整批停下，冲突经搜索后弹窗 / 执行前拦截 /
#      预览按钮禁用三处拦（提示正文见 conflict_block_message）。
# =====================================================================

def _find_occurrences(text, find_text, case_sensitive):
    """返回 find_text 在 text 上的全部出现区间 [(start, end), ...]（不重叠）。"""
    if not find_text:
        return []
    if case_sensitive:
        hay, needle = text, find_text
    else:
        hay, needle = text.lower(), find_text.lower()
    # 大小写折叠在极个别字符上会改变长度（如 'İ'.lower() 是两个字符），
    # 那时下标会错位，退回正则逐段定位。
    if len(hay) == len(text) and len(needle) == len(find_text):
        out, pos = [], hay.find(needle)
        while pos != -1:
            out.append((pos, pos + len(needle)))
            pos = hay.find(needle, pos + len(needle))
        return out
    flags = 0 if case_sensitive else re.IGNORECASE
    return [(m.start(), m.end())
            for m in re.finditer(re.escape(find_text), text, flags)]


def match_hit_count(match):
    """一条命中实际会替换的处数（搜索汇总 / 预览 / 确认框 / 日志共用同一口径）。

    Excel 一个单元格只产一条命中，但 spans 覆盖格内全部出现；Word / txt /
    文件名一条命中就是一处。计数不统一就会出现「预览 1 条、落盘改 3 处」。
    """
    spans = match.get("spans")
    if match.get("type") == "cell" and spans:
        return len(spans)
    return 1


def plan_span_replacements(hits):
    """排定同一容器内的替换，返回 (accepted, conflicts)。

    hits: [{"start", "end", "cond_idx", ...}, ...]；accepted 按 start 降序
    （从后往前应用，替换不改动更靠前区间的偏移）。排序只为**确定性**——
    同一批命中每次要算出同一组冲突对，否则"哪几处冲突"会漂移；谁排前谁排后
    不代表执行优先级。
    """
    ordered = sorted(hits, key=lambda x: (-(x["end"] - x["start"]),
                                          x.get("cond_idx", 0), x["start"]))
    accepted, conflicts = [], []
    for hit in ordered:
        rival = next((a for a in accepted
                      if not (hit["end"] <= a["start"] or hit["start"] >= a["end"])),
                     None)
        if rival is not None:
            conflicts.append((hit, rival))
            continue
        accepted.append(hit)
    accepted.sort(key=lambda x: x["start"], reverse=True)
    return accepted, conflicts


def _span_text_matches(current, find_text, case_sensitive):
    """替换前校验：区间文本是否仍是查找词（文件在搜索后被改过就跳过并报告）。"""
    if case_sensitive:
        return current == find_text
    if current.lower() == find_text.lower():
        return True
    # 折叠会改变长度的字符（如 'İ'.lower() 为两字符）下 lower 口径与
    # _find_occurrences 的 re.IGNORECASE 回退不一致，用同一套正则兜底，
    # 否则预览命中的区间落盘时被误报「文件已修改」而跳过。
    return bool(re.fullmatch(re.escape(find_text), current, re.IGNORECASE))


def conflict_reason(hit, rival):
    """重叠命中的原因文案（不含位置）。

    只说明"谁和谁撞上了"：一律拦下时，谁是先判出的那条对用户没有意义。
    """
    return T("conflict_reason_tpl", hit=hit.get("find_text", ""),
             rival=rival.get("find_text", ""))


def format_conflict_note(where, hit, rival):
    """带位置的重叠提示（搜索完成后汇总、执行前拦截都用它）。"""
    return T("conflict_note_tpl", where=where, reason=conflict_reason(hit, rival))


def conflict_block_message(conflicts, max_shown=8):
    """重叠命中导致整批停下时的提示正文（搜索后弹窗与执行前拦截共用）。

    两处各写各的早晚会自相矛盾，而用户只会相信先看到的那个。
    """
    shown = conflicts[:max_shown]
    lines = [T("conflict_block_head"), ""]
    lines += [f"\u00b7 {c}" for c in shown]
    if len(conflicts) > len(shown):
        lines.append(T("conflict_block_more", n=len(conflicts) - len(shown)))
    lines += ["", T("conflict_block_tail")]
    return "\n".join(lines)


def _hit_spans(match):
    """把一条匹配展开成 [(容器键, start, end), ...]，供跨条件重叠判断用。

    Excel 的一条命中覆盖同一单元格内的多处出现（spans 在提取时算好），
    Word / txt / 文件名的一条命中就是一个区间——展开后才能与别条件做
    重叠判断。区间与容器键一律取自 _chain_spans / _chain_container_key：
    同一条命中在「搜索时提示冲突」与「链式追踪来源」两处必须看到同一套
    区间，否则两处的结论会对不上。
    """
    key = _chain_container_key(match)
    return [(key, s, e) for s, e in _chain_spans(match)]


def detect_condition_conflicts(matches):
    """检测跨条件重叠，返回 [(一条命中, 与它重叠的另一条), ...]。

    与执行前那次判断共用 plan_span_replacements，保证「搜索时提示的」与
    「按钮按死 / 点了也拦」是同一套结论。
    """
    # 同一条件的命中按构造互不重叠（_find_occurrences 只产不重叠区间），
    # 单条件批次不可能有冲突——大 txt 上万条命中时省掉每次勾选都 O(n²)。
    if len({m.get("cond_idx", 0) for m in matches}) < 2:
        return []
    by_container = {}
    for m in matches:
        for key, start, end in _hit_spans(m):
            by_container.setdefault(key, []).append({
                "start": start, "end": end,
                "cond_idx": m.get("cond_idx", 0),
                "find_text": m.get("find_text", ""),
                "replace_text": m.get("replace_text", ""),
                "location": m.get("location", ""),
            })
    conflicts = []
    for group in by_container.values():
        _, overlaps = plan_span_replacements(group)
        conflicts.extend(overlaps)
    return conflicts


def apply_filename_conditions(name, conditions, case_sensitive):
    """把多个条件一次性作用于文件名，返回 (新文件名, 冲突说明列表)。

    并行语义：全部条件在**原始文件名**上匹配。链式替换见 run_filename_chain。
    """
    hits = []
    for cond_idx, (find_text, replace_text) in enumerate(conditions):
        for start, end in _find_occurrences(name, find_text, case_sensitive):
            hits.append({"start": start, "end": end, "cond_idx": cond_idx,
                         "find_text": find_text, "replace_text": replace_text})
    accepted, conflicts = plan_span_replacements(hits)
    new_name = name
    for hit in accepted:
        new_name = (new_name[:hit["start"]] + hit["replace_text"]
                    + new_name[hit["end"]:])
    notes = [format_conflict_note(T("target_filename"), hit, rival)
             for hit, rival in conflicts]
    return new_name, notes


def _span_after_replace(ps, pe, s, e, delta):
    """替换 [s,e)（长度变化 delta）之后，把已产出区间 (ps,pe) 挪到新坐标。

    被替换区间整个吞掉的丢弃；跨在边界上的左右切开——否则贴着替换边写入
    的后续命中会被漏标。
    """
    if pe <= s:
        return [(ps, pe)]
    if ps >= e:
        return [(ps + delta, pe + delta)]
    out = []
    if ps < s:
        out.append((ps, s))
    if pe > e:
        out.append((e + delta, pe + delta))
    return out


def _spans_overlap(span, spans):
    s, e = span
    return any(not (e <= ps or s >= pe) for ps, pe in spans)


def _chain_spans(match):
    """命中在所属容器里的区间列表。

    Excel 一条命中覆盖同一单元格内的多处出现（spans 在提取时算好），Word /
    txt 一条就是一个区间——与 _hit_spans 同一套口径，来源追踪才能与替换
    引擎切分的粒度对上。
    """
    spans = match.get("spans")
    if spans:
        return list(spans)
    if "start" in match and "end" in match:
        return [(match["start"], match["end"])]
    return []


def _chain_container_key(match):
    """来源追踪用的容器键：命中必须落回同一段文本，追踪才有意义。

    Word 以段落为单位、Excel 以单元格为单位、txt 整篇一个容器——与替换
    引擎切分的粒度一致，坐标才能对上。
    """
    if match.get("type") == "cell":
        return ("cell", match.get("sheet_name"), match.get("cell_coordinate"))
    if match.get("type") == "text":
        return ("text",)
    return _container_key(match)


def _track_produced(produced, s, e, new_len):
    """在 produced（区间列表）上登记一次替换 [s, e) → 长 new_len 的文本。

    已产出的区间按新坐标平移：整个吞掉的丢弃、跨在边界上的左右切开——
    否则贴着替换边写入的后续命中会被漏标。返回新列表，原列表不改。
    """
    delta = new_len - (e - s)
    kept = []
    for ps, pe in produced:
        kept.extend(_span_after_replace(ps, pe, s, e, delta))
    kept.append((s, s + new_len))
    return kept


def run_filename_chain(name, conditions, case_sensitive):
    """文件名的链式替换：第 1 条在原文件名上找，下一条在上一条的输出上找。

    返回 (rounds, 新文件名)。命中里落在上一轮写出区间内的标 from_prev，
    预览据此把"这一处是上一轮改出来的"标出来。
    """
    rounds = []
    produced = []
    current = name
    for cond_idx, (find_text, replace_text) in enumerate(conditions, 1):
        spans = _find_occurrences(current, find_text, case_sensitive)
        matches = [{"type": "filename", "location": T("target_filename"),
                    "start": s, "end": e,
                    "context": _context_of(current, s, e - s),
                    "find_text": find_text, "replace_text": replace_text,
                    "cond_idx": cond_idx, "round": cond_idx,
                    "from_prev": _spans_overlap((s, e), produced)}
                   for s, e in spans]
        # 从后往前替换：先写右边的，左边区间的坐标不受影响，产出记录才能
        # 与下一次替换时的文本坐标保持一致
        for s, e in sorted(spans, reverse=True):
            current = current[:s] + replace_text + current[e:]
            produced = _track_produced(produced, s, e, len(replace_text))
        rounds.append({"cond_idx": cond_idx, "find_text": find_text,
                       "replace_text": replace_text, "matches": matches})
    return rounds, current


def run_replace_chain(file_path, conditions, case_sensitive,
                      dep_errors=None, commit=False):
    """把条件按顺序跑成一条链：第 1 条在原文上替换，下一条在上一条的结果上找。

    返回 (rounds, skipped)。每轮的命中区间 / 位置 / 上下文都记录在**该轮的
    输入**上——预览按轮分组展示后，看到的每一轮就是执行时真正发生的事，
    不存在"预览一套、写盘另一套"。

    全程只在临时副本上动手：collect 用途（搜索取样）的副本用完即删；
    commit=True 全部轮次跑完才把副本搬回原位，临时文件必须与原文件同盘
    （os.replace 不能跨卷），故落在原文件旁边。任何一轮抛异常，原文件
    一个字节都不会变。

    skipped 是 (位置, 原因) 列表：缺依赖、受保护 run、公式格等"搜到也
    不替换 / 搜不了"的说明；dep_errors 给了就把缺依赖单独记一份（界面
    要把"搜不了"和"真的没匹配"分开说）。
    """
    src = Path(file_path)
    ext = src.suffix.lower()
    name = os.path.basename(file_path)
    fd, tmp_name = tempfile.mkstemp(suffix=ext,
                                    dir=str(src.parent) if commit else None)
    os.close(fd)
    tmp = Path(tmp_name)
    rounds, skipped = [], []
    # 每轮写进去的区间（按容器），用来判断下一轮的命中是不是"上一轮改出来的"
    produced = {}
    try:
        shutil.copyfile(file_path, tmp)
        # 与 atomic_save 同一口径：临时文件落回原位前先继承原文件权限，
        # 否则链式替换完权限位被重置（并行路径有 copymode，这里不能没有）。
        if commit:
            shutil.copymode(file_path, tmp)
        tmp_path = str(tmp)
        engine = _engine_for(ext)
        for cond_idx, (find_text, replace_text) in enumerate(conditions, 1):
            found, notes = [], []
            # 不认识的格式与旧行为一致：当轮零命中，不报错也不替换
            if engine is not None and not engine["available"]:
                msg = T(engine["missing_search"], name=name)
                if dep_errors is not None:
                    dep_errors.append(msg)
                skipped.append((name, msg))
            elif engine is not None:
                found, notes = engine["extract"](tmp_path, find_text, case_sensitive)
                for note in notes:
                    skipped.append((name, note))
            for m in found:
                m["replace_text"] = replace_text
                m["cond_idx"] = cond_idx
                m["round"] = cond_idx
                spans = _chain_spans(m)
                m["from_prev"] = any(
                    _spans_overlap(s, produced.get(_chain_container_key(m), []))
                    for s in spans)
            # 没有命中的轮次不再过一遍"读入-写回"：xlsx 一轮就是整车
            # load/save，白跑只会放大内嵌对象丢失的风险
            if found:
                skipped.extend(engine["replace"](tmp_path, found, case_sensitive))
                applied = []
                for m in found:
                    key = _chain_container_key(m)
                    for s, e in _chain_spans(m):
                        applied.append((key, s, e, len(m["replace_text"])))
                # 与替换引擎一致：先处理靠后的区间，前面的坐标才不受影响
                applied.sort(key=lambda x: x[1], reverse=True)
                for key, s, e, new_len in applied:
                    produced[key] = _track_produced(produced.get(key, []),
                                                    s, e, new_len)
            rounds.append({"cond_idx": cond_idx, "find_text": find_text,
                           "replace_text": replace_text, "matches": found})
        if commit:
            os.replace(tmp, file_path)
        return rounds, skipped
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def validate_conditions(conditions, case_sensitive):
    """校验条件列表，返回 (规范化后的条件列表, 错误说明列表)。

    规则：查找内容不能为空、不得重复（大小写不敏感时 "A" 与 "a" 算重复）；
    替换内容可以为空（等于删除）。「添加 / 修改条件」与「开始搜索」两处都会
    调用——后者兜底防用户切「匹配大小写」造成隐性重复。
    """
    cleaned, errors, seen = [], [], {}
    for idx, pair in enumerate(conditions, 1):
        find_text = pair[0] if len(pair) > 0 else ""
        replace_text = pair[1] if len(pair) > 1 else ""
        if not find_text:
            errors.append(T("cond_empty", i=idx))
            continue
        key = find_text if case_sensitive else find_text.lower()
        if key in seen:
            errors.append(T("cond_dup", i=idx, find=find_text, j=seen[key]))
            continue
        seen[key] = idx
        cleaned.append([find_text, replace_text])
    return cleaned, errors


def _join_condition_parts(parts, limit):
    """条件简述的公共收尾：前 limit 项用分隔符连接，超出补「等N项」。"""
    text = T("cond_summary_sep").join(parts[:limit])
    if len(parts) > limit:
        text += T("cond_summary_more", n=len(parts))
    return text


def conditions_summary(conditions, limit=2):
    """把条件列表拼成给日志 / 备份文件夹名 / 预览标题用的简述。"""
    parts = [T("cond_summary_item", f=find, r=replace) for find, replace in conditions]
    return _join_condition_parts(parts, limit)


def conditions_rule_base(conditions, limit=2):
    """规则文件名用的条件摘要：条目不带引号，清洗后不会变成一片下划线。"""
    parts = [f"{find}→{replace}" for find, replace in conditions]
    return _join_condition_parts(parts, limit)


def normalize_history_item(item):
    """把一条历史记录统一成「一组条件」的结构（旧版存的是单条 find/replace）。

    不转换则对话框要写两套取值逻辑，旧记录点"选择"也填不出内容；结构认不出
    来时返回 None 丢弃，不让一条坏记录毁掉整份历史。
    """
    if not isinstance(item, dict):
        return None
    conds = item.get("conditions")
    if isinstance(conds, list) and conds:
        pairs = [[str(c[0]), str(c[1])] for c in conds
                 if isinstance(c, (list, tuple)) and len(c) >= 2]
        if not pairs:
            return None
    elif "find" in item or "replace" in item:
        pairs = [[str(item.get("find", "")), str(item.get("replace", ""))]]
    else:
        return None
    return {
        "conditions": pairs,
        "case_sensitive": bool(item.get("case_sensitive", False)),
        "replace_filename": bool(item.get("replace_filename", False)),
        "timestamp": str(item.get("timestamp", "")),
    }


def history_row_keys(history):
    """为查找历史生成稳定行键，返回与 history 等长的键列表（按下标对应）。

    行键必须与下标解耦：用下标当 Treeview iid，删一条后剩余行 iid 不重排，
    再点最后一行就按旧下标取记录而 IndexError。故用时间戳做键（同秒加序号）。
    """
    keys = []
    seen = {}
    for item in history:
        base = str(item.get("timestamp", ""))
        n = seen.get(base, 0)
        seen[base] = n + 1
        keys.append(base if n == 0 else f"{base}_{n}")
    return keys


def atomic_save(save_callable, file_path):
    """
    原子写：先在同目录写临时文件，成功后再 os.replace 替换目标文件；
    中途任何失败都不会破坏原文件，也不会留下半成品文件。
    """
    directory = os.path.dirname(os.path.abspath(file_path)) or "."
    suffix = os.path.splitext(file_path)[1]
    fd, tmp_path = tempfile.mkstemp(prefix=".batchreplace_tmp_", suffix=suffix, dir=directory)
    os.close(fd)
    try:
        if os.path.exists(file_path):
            shutil.copymode(file_path, tmp_path)
        save_callable(tmp_path)
        os.replace(tmp_path, file_path)
    except BaseException:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise
    return file_path


def _context_of(original, pos, target_len):
    start = max(0, pos - 5)
    end = min(len(original), pos + target_len + 5)
    return original[start:end]


def _ctx_parts(match):
    """把一条匹配的 context 拆成（前文, 命中, 后文），预览着色用。

    _context_of 的前缀恒为 min(pos, 5)、命中长度 = end - start，据此切出
    命中段。缺 start / end 的匹配（旧数据、测试造的精简 dict）退化为整段
    前文——只是不上色，不出错。
    """
    ctx = match.get("context", "")
    start = match.get("start")
    if start is None:
        return ctx, "", ""
    hit_len = max(0, match.get("end", start) - start)
    hit_start = min(start, 5)
    return ctx[:hit_start], ctx[hit_start:hit_start + hit_len], ctx[hit_start + hit_len:]


# ---------------------------------------------------------------------
# Word 匹配模型：段落级字符区间
#
# 不能按单个 run 查找：文本被拆进多个 run（中英文边界、加粗、超链接、
# 修订插入）会整段漏掉；用 run.text 整体赋值又会抹掉分页符 / 图片 / 域字符。
# 现行模型：容器内全部 w:t 按文档顺序拼成整段文本找命中区间；替换时按区间
# 映射回各个 w:t，只改文本，其余子元素原样保留。
# ---------------------------------------------------------------------

# 命中这些元素的 run 一律跳过并警告：br=分页/换行、drawing/pict/object=绘图对象、fldChar=域字符
PROTECTED_RUN_TAGS = ("w:br", "w:drawing", "w:pict", "w:object", "w:fldChar")


def _set_t_text(t_el, text):
    """写 w:t 文本，并按需维护 xml:space="preserve"（首尾空白不能丢）。"""
    t_el.text = text
    space = qn("xml:space")
    if text and len(text.strip()) < len(text):
        t_el.set(space, "preserve")
    elif t_el.get(space) is not None:
        del t_el.attrib[space]


def _collect_flow_runs(el):
    """收集 el 之下属于文档文本流的 w:r（按文档顺序）。

    必须覆盖超链接 / 修订插入内的 run——para.runs 只取直接子 run，另两类
    会整段漏掉；不进入 drawing/pict/object 内部，避免浮动内容制造虚假匹配。
    """
    runs = []

    def walk(node):
        for child in node:
            tag = child.tag
            if tag in (qn("w:drawing"), qn("w:pict"), qn("w:object")):
                continue  # 浮动内容整体跳过（run 自身含图片时该 run 已被收集并标记受保护）
            if tag == qn("w:r"):
                runs.append(child)
                continue  # run 内不会再嵌套 run
            walk(child)

    walk(el)
    return runs


def _run_is_protected(run_el):
    """run 内含分页符 / 图片 / 域字符等结构性元素时为 True（命中即跳过）。"""
    return any(run_el.find(qn(tag)) is not None for tag in PROTECTED_RUN_TAGS)


def _paragraph_run_records(para):
    """把段落拆成 run 记录：每个 run 的 w:t 元素及其在整段文本中的偏移区间。

    返回 (full_text, records)。records[i] = {
        "el": run 元素, "protected": 是否含结构性元素,
        "t_spans": [(w:t 元素, 起始, 结束), ...],
        "start": run 文本起始, "end": run 文本结束,
    }
    提取与替换必须走同一函数，两边看到的文本与偏移才一致。
    """
    records = []
    pos = 0
    for run_el in _collect_flow_runs(para._p):
        t_spans = []
        run_start = pos
        for t in run_el.findall(qn("w:t")):
            text = t.text or ""
            t_spans.append((t, pos, pos + len(text)))
            pos += len(text)
        records.append({
            "el": run_el,
            "protected": _run_is_protected(run_el),
            "t_spans": t_spans,
            "start": run_start,
            "end": pos,
        })
    full = "".join(t.text or "" for rec in records for (t, _, _) in rec["t_spans"])
    return full, records


def _word_paragraph_numbers(body):
    """按 Word `doc.Paragraphs` 的编号给正文每个段落编号，返回 {段落元素: 编号}。

    Word 的集合与 python-docx 的 doc.paragraphs 不是一套口径：它含表格
    单元格内的段落，且每个表格行尾还有一个行标记幻影段（实测：XML 全部
    w:p 加每行一个）。doc.paragraphs 只数顶层段落——文档一有表格，拿它的
    序号去 doc.Paragraphs(n) 落位就整体偏小，表格之后所有顶层段落全跳错
    地方。提取 / 替换侧只把这编号当容器键用，数值变化不影响两侧回查。
    """
    numbers = {}
    counter = 0

    def walk(el):
        nonlocal counter
        for child in el.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                counter += 1
                numbers[child] = counter
            elif tag == qn("w:tbl"):
                # 行尾幻影段：Word 在每个表格行结束后多数一段（行标记）
                for row in child.findall(qn("w:tr")):
                    for cell in row.findall(qn("w:tc")):
                        walk(cell)
                    counter += 1
            elif tag != qn("w:txbxContent"):
                # 内容控件（w:sdt）等透明容器要穿过去；文本框是独立 story，
                # Word 不把它算进 doc.Paragraphs，穿过去会多数
                walk(child)

    walk(body)
    return numbers


def _iter_word_containers(doc):
    """按固定顺序枚举全部文本容器；提取与替换共用（替换靠下标回查）。

    顶层段落的 para_idx 与 Word 的 doc.Paragraphs 编号对齐——跳转侧直接
    用 doc.Paragraphs(n) 落位，编号规则见 _word_paragraph_numbers。

    跳过链接到上一节的页眉 / 页脚（与所链接节是同一部件，逐节枚举会重复
    扫）；且绝不访问链接态的 .paragraphs——python-docx 会就地创建部件。
    """
    numbers = _word_paragraph_numbers(doc.element.body)
    for para in doc.paragraphs:
        yield "paragraph", {"para_idx": numbers[para._p]}, para
    for table_idx, table in enumerate(doc.tables, 1):
        for row_idx, row in enumerate(table.rows, 1):
            for col_idx, cell in enumerate(row.cells, 1):
                for para_idx, para in enumerate(cell.paragraphs, 1):
                    yield "table", {
                        "table_idx": table_idx, "row_idx": row_idx,
                        "col_idx": col_idx, "para_idx": para_idx,
                    }, para
    for section_idx, section in enumerate(doc.sections, 1):
        if not section.header.is_linked_to_previous:
            for para_idx, para in enumerate(section.header.paragraphs, 1):
                yield "header", {"section_idx": section_idx, "para_idx": para_idx}, para
        if not section.footer.is_linked_to_previous:
            for para_idx, para in enumerate(section.footer.paragraphs, 1):
                yield "footer", {"section_idx": section_idx, "para_idx": para_idx}, para


def _word_location(kind, idx):
    """匹配项的「位置」显示串（与旧版格式保持一致）。"""
    if kind == "paragraph":
        return T("loc_paragraph", n=idx["para_idx"])
    if kind == "table":
        return T("loc_table", t=idx["table_idx"], r=idx["row_idx"], c=idx["col_idx"])
    name = T("loc_header") if kind == "header" else T("loc_footer")
    return T("loc_section", s=idx["section_idx"], name=name, p=idx["para_idx"])


def _file_iid(idx):
    """文件行的 Treeview iid。

    构造与解析（_iid_index）分居两处，统一从这里出，避免两边前缀写散后
    出现不报错的错位。
    """
    return f"file_{idx}"


def _match_iid(file_iid, midx):
    """匹配子行的 Treeview iid（父行 iid + 序号）。"""
    return f"{file_iid}_match_{midx}"


def _round_iid(file_iid, ridx):
    """链式预览里轮次行的 iid（文件行 iid + 轮次号）。

    轮次行是分组标题，不是可勾选的匹配行，故单独一套前缀：iid 反查时
    轮次行走 _round_index，不会与文件行（match_idx=None）混淆。
    """
    return f"{file_iid}_round_{ridx}"


def _word_find_nth(container, text, occurrence, match_case):
    """在 container（Word 段落 / 单元格 Range）内找第 occurrence 次出现的 text。

    逐次前进而非「找到就停」：同一段里同一个词出现多次时，后者永远停在
    第一处，而用户双击的可能是第二处。Wrap 设 wdFindStop：越出容器即找错
    地方，不如放弃。
    """
    search = container.Duplicate
    pos = container.Start
    for _ in range(max(occurrence, 1)):
        search.SetRange(pos, container.End)
        find = search.Find
        find.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.Wrap = 0
        find.MatchCase = match_case
        find.MatchWildcards = False
        find.MatchWholeWord = False
        if not find.Execute():
            return None
        pos = search.End
    return search


def _container_key(match):
    """匹配项 → 容器键（替换时用它把匹配分组回同一段落）。"""
    idx = {k: v for k, v in match.items()
           if k in ("para_idx", "table_idx", "row_idx", "col_idx", "section_idx")}
    return (match["type"], tuple(sorted(idx.items())))


def _apply_word_span(records, start, end, replace_text):
    """把 [start, end) 区间替换为 replace_text：只碰 w:t，其他子元素原样保留。

    跨多个 w:t 的区间：第一个被覆盖的 w:t 放入替换词，其余被覆盖部分删掉；
    分页符 / 图片 / 域字符 / 制表符不在 w:t 里，天然不受影响。
    """
    placed = False
    for rec in records:
        for t, ts, te in rec["t_spans"]:
            if te <= start or ts >= end:
                continue
            local_s = max(start, ts) - ts
            local_e = min(end, te) - ts
            text = t.text or ""
            if not placed:
                _set_t_text(t, text[:local_s] + replace_text + text[local_e:])
                placed = True
            else:
                _set_t_text(t, text[:local_s] + text[local_e:])


def _engine_for(ext):
    """扩展名 → 引擎记录；不认识的格式返回 None。

    搜索 / 替换 / 链式 / 跳转四个调用点共用同一张表，新增一种格式只改
    这里——漏掉任一处的表现是该格式在某个环节静默失灵。
    只认 .docx / .xlsx / .txt：python-docx / openpyxl 读不了老的 .doc /
    .xls，放进来会在提取阶段就炸。.txt 一律受理、按启发式编码读取——
    二进制误带 .txt 会被 latin-1 兜底当文本读，这是已知取舍而非拦截。
    """
    if ext == ".docx":
        return {"extract": extract_word_matches, "replace": replace_word_multi,
                "available": DOCX_AVAILABLE,
                "missing_search": "e_missing_docx",
                "missing_replace": "e_docx_missing",
                "jump": "jump_to_word_location"}
    if ext == ".xlsx":
        return {"extract": extract_excel_matches, "replace": replace_excel_multi,
                "available": EXCEL_AVAILABLE,
                "missing_search": "e_missing_openpyxl",
                "missing_replace": "e_openpyxl_missing",
                "jump": "jump_to_excel_location"}
    if ext == ".txt":
        return {"extract": extract_text_matches, "replace": replace_text_multi,
                "available": True,
                "missing_search": None, "missing_replace": None,
                "jump": "jump_to_text_location"}
    return None


def _read_text_auto(file_path):
    """读 .txt：启发式探测编码，返回 (text, encoding)。

    BOM 按 utf-8-sig 处理；无 BOM 依次试 utf-8、gbk，最后 latin-1 兜底
    （字节级保真）。encoding 为 utf-8-sig 当且仅当原文件带 BOM，写回据此
    决定加不加 BOM。
    """
    with open(file_path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        try:
            return raw.decode("utf-8-sig"), "utf-8-sig"
        except UnicodeDecodeError:
            pass
    for enc in ("utf-8", "gbk"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1"), "latin-1"


def _write_text_auto(file_path, text, encoding):
    """写回：BOM 文件用 utf-8-sig 还原 BOM，其余按探测到的编码写回。
    newline="" 保留原文件换行风格（\\n / \\r\\n），不做跨平台归一。"""
    with open(file_path, "w", encoding=encoding, newline="") as f:
        f.write(text)


def _line_col_of(text, pos):
    """由绝对偏移求 1-based 行号与列号（按 \\n 切分，\\r 不计入列偏移）。"""
    line = text.count("\n", 0, pos) + 1
    col = pos - (text.rfind("\n", 0, pos) + 1) + 1
    return line, col


def extract_text_matches(file_path, find_text, case_sensitive):
    """扫描 .txt 正文，返回 (matches, skipped)。

    location 用「第 N 行, 第 M 列」（行从 1，列为命中起点在该行的字符偏移 +1）。
    skipped 恒为空：纯文本分支不依赖外部库，读取失败走异常而非静默空匹配。
    """
    if not find_text:
        return [], []
    try:
        text, _ = _read_text_auto(file_path)
    except OSError as e:
        return [], [T("e_text_read", e=str(e))]
    matches = []
    for start, end in _find_occurrences(text, find_text, case_sensitive):
        line, col = _line_col_of(text, start)
        matches.append({
            "type": "text",
            "location": T("loc_text", line=line, col=col),
            "start": start, "end": end,
            "context": _context_of(text, start, end - start),
            "find_text": find_text,
        })
    return matches, []


def replace_text_multi(file_path, matches, case_sensitive):
    """按各匹配自带的区间一次性替换 .txt 正文并原子写回。

    命中全部来自原始文本、从后往前应用（语义见「多条件替换」区头）。
    返回 [(位置, 原因)]。
    """
    skipped = []
    text, encoding = _read_text_auto(file_path)
    hits = []
    for m in matches:
        find_text = m.get("find_text")
        if "start" not in m or "end" not in m:
            skipped.append((m.get("location", "?"), T("sk_no_span")))
            continue
        if not find_text:
            skipped.append((m.get("location", "?"), T("sk_no_find")))
            continue
        hits.append({"start": m["start"], "end": m["end"],
                     "cond_idx": m.get("cond_idx", 0),
                     "find_text": find_text,
                     "replace_text": m.get("replace_text", ""),
                     "location": m.get("location", "?")})
    # 整个文件就是一个容器，命中数可能上万：不套 plan_span_replacements
    # （两两比对 O(n²)），按 start 降序应用，只需与已应用的最靠前区间比一次。
    # 重叠在搜索阶段就该整批拦下，这里仅兜底「绕过界面直接调引擎」。
    hits.sort(key=lambda h: h["start"], reverse=True)
    new_text = text
    frontier = None   # 已应用区间中最靠前的 start
    last_hit = None   # 上一条已应用的命中：重叠时它就是 rival
    for hit in hits:
        start, end = hit["start"], hit["end"]
        if frontier is not None and end > frontier:
            skipped.append((hit["location"], conflict_reason(hit, last_hit)))
            continue
        # 替换前校验区间文本仍是查找词——文件若在搜索后被改过就跳过并报告，绝不盲写
        if not _span_text_matches(new_text[start:end], hit["find_text"],
                                  case_sensitive):
            skipped.append((hit["location"], T("sk_text_changed")))
            continue
        new_text = new_text[:start] + hit["replace_text"] + new_text[end:]
        frontier = start
        last_hit = hit
    if new_text != text:
        atomic_save(lambda tmp: _write_text_auto(tmp, new_text, encoding), file_path)
    return skipped


def extract_word_matches(file_path, find_text, case_sensitive):
    """扫描 .docx 的段落/表格/页眉页脚，返回 (matches, skipped)。

    匹配按「段落级字符区间」记录，跨 run / 超链接 / 修订都能命中；run 含
    分页符 / 图片 / 域字符的命中进 skipped。另带 occurrence 等字段供双击
    跳转定位——Word 自算的字符位置与此处偏移对不上，不能复用 start/end。
    """
    doc = Document(file_path)
    matches, skipped = [], []
    if not find_text:
        return matches, skipped
    target = find_text if case_sensitive else find_text.lower()

    for kind, idx, para in _iter_word_containers(doc):
        full, records = _paragraph_run_records(para)
        if not full:
            continue
        check = full if case_sensitive else full.lower()
        location = _word_location(kind, idx)
        pos = check.find(target)
        occurrence = 0
        while pos != -1:
            # 序号统计「本容器内第几次出现」，含被跳过的命中：跳转侧是按 Word
            # 自己的文本搜的，它数得到被跳过的那些，两边序号必须一致。
            occurrence += 1
            end = pos + len(target)
            hit_runs = [r for r in records if not (r["end"] <= pos or r["start"] >= end)]
            if any(r["protected"] for r in hit_runs):
                skipped.append(T("ex_run_protected", loc=location))
            else:
                first_run_idx = next(
                    (i for i, r in enumerate(records, 1)
                     if not (r["end"] <= pos or r["start"] >= end)), 0)
                match = dict(idx)
                match.update({
                    "type": kind,
                    "start": pos, "end": end,
                    "run_idx": first_run_idx,
                    "location": location,
                    "context": _context_of(full, pos, len(target)),
                    "text": full[pos:end],
                    "occurrence": occurrence,
                    "find_text": find_text,
                    "case_sensitive": case_sensitive,
                })
                matches.append(match)
            pos = check.find(target, pos + len(target))
    return matches, skipped


def extract_excel_matches(file_path, find_text, case_sensitive):
    """扫描 .xlsx 的单元格，返回 (matches, skipped)。

    公式格（data_type == "f"）一律跳过——把公式串里的查找词当普通文本改写
    会静默损坏公式；跳过项在 skipped 里汇总。
    """
    wb = load_workbook(file_path, read_only=True)
    matches = []
    formula_cells = []
    target = find_text if case_sensitive else find_text.lower()
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    if getattr(cell, "data_type", None) == "f":
                        formula_cells.append(f"{sheet_name}!{cell.coordinate}")
                        continue
                    original = str(cell.value)
                    check_text = original if case_sensitive else original.lower()
                    if target in check_text:
                        pos = check_text.find(target)
                        matches.append({
                            "type": "cell", "sheet_name": sheet_name,
                            "cell_coordinate": cell.coordinate,
                            "location": f"{sheet_name}!{cell.coordinate}",
                            "context": _context_of(original, pos, len(target)), "value": cell.value,
                            "find_text": find_text,
                            # 命中区间照记：预览上下文列按 start / end 给命中文本上色；
                            # spans 是全部出现（多条件重叠判断用），start / end 只是首个
                            "start": pos, "end": pos + len(target),
                            # 全部出现区间：多条件要在原始值上一次算完（供重叠判断），
                            # 不在替换结果上逐条再找（见「多条件替换」）
                            "spans": _find_occurrences(original, find_text, case_sensitive),
                        })
    wb.close()
    skipped = []
    if formula_cells:
        shown = T("list_sep").join(formula_cells[:5])
        if len(formula_cells) > 5:
            shown += T("list_more", n=len(formula_cells))
        skipped.append(T("ex_formula_cells", n=len(formula_cells), shown=shown))
    return matches, skipped


def replace_word_multi(file_path, matches, case_sensitive):
    """按每条匹配自带的区间与查找/替换词，一次性替换 .docx 内容。

    命中全部来自原始文本、同一次遍历里从后往前应用（语义见「多条件替换」区头）。
    重叠命中落进 skipped——界面发现冲突即整批停下，此处仅兜底「绕过界面直接调引擎」。
    返回 [(位置, 原因)]。
    """
    doc = Document(file_path)
    skipped = []
    groups = {}
    for m in matches:
        if "type" not in m or "start" not in m:
            skipped.append((m.get("location", "?"), T("sk_no_span")))
            continue
        if not m.get("find_text"):
            skipped.append((m.get("location", "?"), T("sk_no_find")))
            continue
        groups.setdefault(_container_key(m), []).append(m)

    containers = {}
    for kind, idx, para in _iter_word_containers(doc):
        containers[(kind, tuple(sorted(idx.items())))] = para

    for key, group in groups.items():
        para = containers.get(key)
        if para is None:
            for m in group:
                skipped.append((m.get("location", "?"), T("sk_no_paragraph")))
            continue
        accepted, conflicts = plan_span_replacements(group)
        for hit, rival in conflicts:
            skipped.append((hit.get("location", "?"),
                            conflict_reason(hit, rival)))
        full, records = _paragraph_run_records(para)
        # 从后往前应用：后面的替换不会改变前面区间的偏移
        for m in accepted:
            start, end = m["start"], m["end"]
            find_text = m["find_text"]
            replace_text = m.get("replace_text", "")
            location = m.get("location", "?")
            if start >= end or end > len(full):
                skipped.append((location, T("sk_invalid_span")))
                continue
            hit_runs = [r for r in records if not (r["end"] <= start or r["start"] >= end)]
            if any(r["protected"] for r in hit_runs):
                skipped.append((location, T("sk_run_protected")))
                continue
            current = full[start:end]
            if not _span_text_matches(current, find_text, case_sensitive):
                skipped.append((location, T("sk_text_changed")))
                continue
            _apply_word_span(records, start, end, replace_text)
            # 区间文本已变，重新解析以便下一条（更靠前的）匹配做文本校验
            full, records = _paragraph_run_records(para)
    atomic_save(doc.save, file_path)
    return skipped


def restore_cell_value(original_value, new_text):
    """把替换后的文本还原成原单元格的类型；无法还原时回落为字符串。

    只还原 int / float 不够：bool 与日期时间会永久退化成字符串；bool 必须
    排在 int 前判断——isinstance(True, int) 为真。
    """
    if isinstance(original_value, bool):
        low = new_text.strip().lower()
        if low in ("true", "1"):
            return True
        if low in ("false", "0"):
            return False
        return new_text
    if isinstance(original_value, int):
        try:
            return int(new_text)
        except (TypeError, ValueError):
            return new_text
    if isinstance(original_value, float):
        try:
            return float(new_text)
        except (TypeError, ValueError):
            return new_text
    # datetime 是 date 的子类，必须排在前面
    if isinstance(original_value, datetime):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f",
                    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return datetime.strptime(new_text, fmt)
            except ValueError:
                continue
        return new_text
    if isinstance(original_value, date):
        try:
            return datetime.strptime(new_text, "%Y-%m-%d").date()
        except ValueError:
            return new_text
    if isinstance(original_value, time_type):
        for fmt in ("%H:%M:%S", "%H:%M:%S.%f", "%H:%M"):
            try:
                return datetime.strptime(new_text, fmt).time()
            except ValueError:
                continue
        return new_text
    return new_text


def replace_excel_multi(file_path, matches, case_sensitive):
    """按各条件在单元格原始值上一次算出结果，替换 .xlsx 内容。

    命中全部来自原始文本、从后往前一次应用（语义见「多条件替换」区头）；
    重叠命中落进 skipped（界面会先拦下，此处仅兜底）。
    公式格双保险：替换前再查一次 data_type，绝不改写公式。返回 [(位置, 原因)]。
    """
    wb = load_workbook(file_path)
    skipped = []
    cells = {}
    for m in matches:
        sheet_name = m.get("sheet_name")
        coord = m.get("cell_coordinate")
        if not sheet_name or not coord:
            skipped.append((m.get("location", "?"), T("sk_no_cell_info")))
            continue
        cells.setdefault((sheet_name, coord), []).append(m)

    for (sheet_name, coord), mlist in cells.items():
        location = f"{sheet_name}!{coord}"
        try:
            cell = wb[sheet_name][coord]
        except Exception:
            skipped.append((location, T("sk_no_cell")))
            continue
        if cell.value is None:
            skipped.append((location, T("sk_cell_empty")))
            continue
        if getattr(cell, "data_type", None) == "f":
            skipped.append((location, T("sk_formula_cell")))
            continue
        original = str(cell.value)
        hits = []
        for m in mlist:
            find_text = m.get("find_text")
            if not find_text:
                skipped.append((location, T("sk_no_find")))
                continue
            # spans 是搜索时按原始值算好的；外部拼出的匹配项（无 spans）现算一次，保证区间与当前值一致
            spans = m.get("spans")
            if not spans:
                spans = _find_occurrences(original, find_text, case_sensitive)
            hits.extend({"start": s, "end": e, "cond_idx": m.get("cond_idx", 0),
                         "find_text": find_text,
                         "replace_text": m.get("replace_text", "")}
                        for s, e in spans)
        accepted, conflicts = plan_span_replacements(hits)
        for hit, rival in conflicts:
            skipped.append((location, conflict_reason(hit, rival)))
        new_text = original
        for hit in accepted:
            # 替换前校验区间文本仍是查找词——文件若在搜索后被改过就跳过并报告，绝不盲写
            current = new_text[hit["start"]:hit["end"]]
            if not _span_text_matches(current, hit["find_text"], case_sensitive):
                skipped.append((location, T("sk_text_changed")))
                continue
            new_text = new_text[:hit["start"]] + hit["replace_text"] + new_text[hit["end"]:]
        if new_text != original:
            cell.value = restore_cell_value(cell.value, new_text)
    atomic_save(wb.save, file_path)
    return skipped


def xlsx_embedded_objects(file_path):
    """返回 xlsx 内可能被重写丢失 / 变化的对象类别（替换前预警用）。

    只读 zip 目录。openpyxl 整车 load/save 会重写整个包，内嵌图片 / 图表
    等能否保留取决于文件来源——替换前先告知、给取消机会，不静默损坏。
    """
    risky = []
    try:
        with zipfile.ZipFile(file_path) as z:
            names = z.namelist()
    except Exception:
        return risky
    checks = (
        ("xl/media/", T("emb_image")),
        ("xl/charts/", T("emb_chart")),
        ("xl/drawings/", T("emb_drawing")),
        ("xl/embeddings/", T("emb_object")),
    )
    for prefix, label in checks:
        if any(n.startswith(prefix) for n in names):
            risky.append(label)
    return risky


def build_backup_folder_name(find_text, replace_text, is_filename, now=None):
    """构造备份文件夹名（不含路径），与原 create_backup_folder 规则一致。"""
    now = now or datetime.now()
    find_short = sanitize_filename(find_text[:20] if find_text else "")
    replace_short = sanitize_filename(replace_text[:20] if replace_text else "")
    replace_type = T("bk_filename") if is_filename else T("bk_content")
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")
    if find_short and replace_short:
        # 查找词与替换词之间用箭头 "→"：直观表达替换方向，且明确存的是
        # "箭头左边"（替换前）的文件。占位符个数须与参数严格相等——str.format
        # 对多余参数静默忽略，少了 time_str 会让同一天重复替换的文件夹名撞车。
        return f"backup_{date_str}_{find_short}→{replace_short}_{replace_type}_{time_str}"
    if find_short:
        return f"backup_{date_str}_{find_short}_{replace_type}_{time_str}"
    return f"backup_{date_str}_{replace_type}_{time_str}"


def unique_path(path):
    """若 path 已存在，返回追加 _1 / _2 ... 的不冲突路径。"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(f"{base}_{counter}{ext}"):
        counter += 1
    return f"{base}_{counter}{ext}"


def create_backup_folder(base_dir, find_text, replace_text, is_filename, now=None):
    """在 base_dir 下按命名规则创建备份文件夹（重名自动加序号），返回完整路径。"""
    name = build_backup_folder_name(find_text, replace_text, is_filename, now)
    folder = unique_path(os.path.join(base_dir, name))
    os.makedirs(folder, exist_ok=True)
    return folder


def create_backup(file_path, backup_folder):
    """把文件复制进备份文件夹，重名自动加序号。返回备份文件路径。"""
    os.makedirs(backup_folder, exist_ok=True)
    name = os.path.basename(file_path)
    target = unique_path(os.path.join(backup_folder, name))
    try:
        shutil.copy2(file_path, target)
    except Exception as e:
        raise Exception(T("e_backup_create", e=str(e)))
    return target


# 回溯 manifest：内容模式替换成功后写进当次备份夹，回滚时按它把原文件搬回去
MANIFEST_VERSION = 1
MANIFEST_NAME = "manifest.json"


def canonical_path(path):
    """把路径规范化成回溯落盘用的规范形式。

    realpath 解掉符号链接与 ..；Documents 常被重定向到 OneDrive，两处路径
    看着一样其实不是同一个东西，按规范路径写回才落得回原处。
    """
    try:
        return os.path.realpath(os.path.abspath(path))
    except Exception:
        return os.path.abspath(path)


def write_manifest(backup_folder, data):
    """manifest 原子写：先写同目录临时文件再 rename。

    直接覆盖写，崩溃会留半截 JSON——下次扫描整本备份夹都读不出来。
    """
    os.makedirs(backup_folder, exist_ok=True)
    target = os.path.join(backup_folder, MANIFEST_NAME)
    tmp = target + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, target)
    return target


def read_manifest(backup_folder):
    """读 manifest；不存在返回 None，损坏抛异常（调用方决定提示还是跳过）。"""
    path = os.path.join(backup_folder, MANIFEST_NAME)
    if not os.path.isfile(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def scan_backups(base_dir):
    """扫 base_dir 下所有带 manifest 的备份夹，按时间倒序返回回溯条目。

    文件名替换模式不写 manifest，天然被排除——回溯只覆盖内容替换。
    manifest 损坏的那一夹跳过并在 problems 里带上原因，不让整本备份夹读不出来。
    """
    entries, problems = [], []
    if not base_dir or not os.path.isdir(base_dir):
        return entries, problems
    for name in sorted(os.listdir(base_dir)):
        folder = os.path.join(base_dir, name)
        if not os.path.isdir(folder):
            continue
        try:
            data = read_manifest(folder)
        except Exception as e:
            problems.append((folder, str(e)))
            continue
        if not data:
            continue
        if data.get("mode") != "content":
            continue
        entries.append({
            "folder": folder,
            "ts": data.get("ts", ""),
            "conditions": data.get("conditions", []),
            "chain": bool(data.get("chain", False)),
            "multi": bool(data.get("multi", False)),
            "case_sensitive": bool(data.get("case_sensitive", False)),
            "files": data.get("files", []),
        })
    entries.sort(key=lambda e: e["ts"], reverse=True)
    return entries, problems


def detect_later_operations(entries, current, orig_canon):
    """找出同一文件在**更晚**操作里也出现过的那些条目。

    manifest 是平铺列表不是栈：直接回滚较早的一次，会把其后操作的改动一并
    盖掉。返回比 current 更新的条目列表（调用方拿去拼警告）。
    """
    later = []
    for entry in entries:
        if entry is current or entry["ts"] <= current["ts"]:
            continue
        if any(canonical_path(f.get("orig", "")) == orig_canon
               for f in entry["files"]):
            later.append(entry)
    return later


# OneDrive 占位文件的属性位（文件在云端、本地只有一个壳）
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x00400000


def _long_path(path):
    r"""超长路径加 `\\?\` 前缀，绕开 Windows MAX_PATH（260）限制。

    只作用于真正做文件操作的路径：manifest 里存的一律是可读的普通路径，
    前缀加进去用户就没法看了。
    """
    if len(path) < 260 or path.startswith("\\\\?\\"):
        return path
    return "\\\\?\\" + os.path.abspath(path)


def is_cloud_placeholder(path):
    """文件是否处于 OneDrive「在线占位」状态。

    备份夹常在 Documents 下、很可能被 OneDrive 同步；占位文件读会抛
    ERROR_CLOUD_FILE，拿它当已下载用会得到空壳。读写两端都要先问这一句。
    """
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if attrs == -1:
            return False
        return bool(attrs & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS)
    except Exception:
        return False


def rollback_files(entry, selected_files):
    """把 entry 里选中的文件恢复成替换前的样子。

    每个文件：先把**当前**文件快照进备份夹 rollback_ 子夹，再把备份原子搬
    回原位——误回滚也能从 rollback_ 里捞回来。返回 (成功列表, 失败列表)。
    """
    folder = entry["folder"]
    snap_dir = os.path.join(folder, "rollback")
    ok, bad = [], []
    for item in selected_files:
        orig = _long_path(item.get("orig_canon") or item.get("orig", ""))
        backup_rel = item.get("backup_rel", "")
        backup = _long_path(os.path.join(folder, backup_rel)) if backup_rel else ""
        try:
            # orig 为空时 tmp 会落进当前工作目录且 replace 必败——先拦下，
            # 不让垃圾文件写去莫名其妙的地方。
            if not orig:
                raise Exception(T("e_no_orig"))
            if not backup or not os.path.isfile(backup):
                raise Exception(T("e_backup_missing", name=backup_rel))
            if os.path.isfile(orig) and is_cloud_placeholder(orig):
                raise Exception(T("e_cloud_placeholder", name=os.path.basename(orig)))
            if is_cloud_placeholder(backup):
                raise Exception(T("e_cloud_placeholder", name=backup_rel))
            if os.path.isfile(orig):
                os.makedirs(_long_path(snap_dir), exist_ok=True)
                shutil.copy2(orig, _long_path(unique_path(os.path.join(
                    snap_dir, os.path.basename(orig)))))
            # 原子搬回：同目录临时文件 + replace，避免写一半崩留下截断的文件
            tmp = _long_path(orig + ".rollback_tmp")
            try:
                shutil.copy2(backup, tmp)
                os.replace(tmp, orig)
            finally:
                # 失败也要清掉半成品，不能留下 .rollback_tmp 残骸
                try:
                    if os.path.exists(tmp):
                        os.remove(tmp)
                except OSError:
                    pass
            ok.append(orig)
        except Exception as e:
            bad.append((orig, str(e)))
    return ok, bad


# =====================================================================
# 工具函数 / 界面
# =====================================================================

def human_size(n):
    """把字节数格式化为人类可读的大小字符串"""
    n = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{int(n)} B" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0


def file_meta(path):
    """返回 (类型, 大小) 元组，供文件列表显示"""
    ext = os.path.splitext(path)[1].lstrip(".").upper() or "\u2014"
    try:
        size = os.path.getsize(path)
    except OSError:
        size = 0
    return ext, human_size(size)


class _ZScrollbar(tk.Canvas):
    """自绘滚动条：水槽 + 滑块，无箭头按钮。

    ttk 滚动条的水槽在空列表 / 禁用态下不按主题上色（透出父控件底色或
    退回默认浅色），深浅两种主题都因此露出白边；本类颜色取自构造参数，
    列表有没有内容、控件处于什么状态，渲染结果都只认主题色板。
    接口与 ttk.Scrollbar 兼容：set(first, last) 供 yscrollcommand 回调，
    command 只接收 "moveto" 动作。
    """

    def __init__(self, master, command=None, width=10, trough=None, thumb=None,
                 **kw):
        super().__init__(master, width=width, highlightthickness=0, bd=0,
                         takefocus=0, bg=trough, **kw)
        self._command = command
        self._thumb_color = thumb
        self._first, self._last = 0.0, 1.0
        self._drag_first = None
        self._drag_y = 0
        self.bind("<Button-1>", self._press)
        self.bind("<B1-Motion>", self._drag)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Configure>", lambda _e: self._redraw())

    # ---- ttk.Scrollbar 兼容接口：树视图通过 yscrollcommand 驱动 ----
    def set(self, first, last):
        self._first, self._last = float(first), float(last)
        self._redraw()

    def _thumb_geom(self):
        """滑块 (top, bottom)；不可滚动时返回 None（不画滑块）。"""
        h = self.winfo_height()
        if h <= 0 or self._last - self._first >= 1.0:
            return None
        top = int(self._first * h)
        return top, max(int(self._last * h), top + 24)

    def _redraw(self):
        self.delete("all")
        geom = self._thumb_geom()
        if geom is None:
            return
        top, bottom = geom
        w = self.winfo_width()
        self.create_oval(1, top, w - 1, bottom, fill=self._thumb_color,
                         outline="")

    def _moveto(self, fraction):
        if self._command:
            self._command("moveto", max(0.0, min(1.0, fraction)))

    def _press(self, e):
        geom = self._thumb_geom()
        if geom is not None and geom[0] <= e.y <= geom[1]:
            self._drag_first, self._drag_y = self._first, e.y
        else:
            # 点水槽：让滑块中心跳到点击处
            h = self.winfo_height()
            if h > 0:
                span = self._last - self._first
                self._moveto(e.y / h - span / 2)

    def _drag(self, e):
        if self._drag_first is None:
            return
        h = self.winfo_height()
        if h > 0:
            self._moveto(self._drag_first + (e.y - self._drag_y) / h)

    def _release(self, _e):
        self._drag_first = None


# ---- 更新检查与一键升级（公开仓库预留路径）----
def version_tuple(text):
    """从 vX.Y.Z / 文件名里取三段数字，取不到按 (0, 0, 0) 处理。"""
    m = re.search(r"(\d+)\.(\d+)\.(\d+)", str(text))
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def exe_version(path):
    return version_tuple(os.path.basename(str(path)))


def compare_versions(a, b):
    """返回 1 表示 a 更新，-1 更旧，0 相同。"""
    ta, tb = version_tuple(a), version_tuple(b)
    return (ta > tb) - (ta < tb)


# ---- 一键更新：清单校验 + 下载校验 + 校验即启动 ----
# SHA-256 只防下载损坏/截断，不防恶意替换（哈希与文件同源）。单一 GitHub 来源 + HTTPS 下足够；
# 若将来接入第三方镜像，必须补清单签名。
UPDATE_MAX_MANIFEST = 64 * 1024
UPDATE_MAX_EXE = 200 * 1024 * 1024
_LOOPBACK_HOSTS = {"127.0.0.1", "localhost"}


def validate_update_manifest(m):
    """字段校验；非法抛 ValueError。"""
    if not isinstance(m, dict):
        raise ValueError("清单必须是 JSON 对象")
    if version_tuple(m.get("version")) == (0, 0, 0):
        raise ValueError("清单 version 非法")
    exe = m.get("exe")
    if not isinstance(exe, str) or not exe.lower().endswith(".exe"):
        raise ValueError("清单 exe 必须以 .exe 结尾")
    if re.search(r"[\\/]", exe) or ".." in exe:
        raise ValueError("清单 exe 不得含路径分隔或 ..")
    if not re.search(r"[_-]v?(\d+\.\d+\.\d+)", exe, re.I):
        raise ValueError("清单 exe 文件名必须含可解析版本号")
    sha = m.get("sha256")
    if not isinstance(sha, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", sha):
        raise ValueError("清单 sha256 必须是 64 位十六进制")
    url = m.get("url")
    if not isinstance(url, str) or not url:
        raise ValueError("清单缺少 url")
    size = m.get("size")
    if size is not None and (isinstance(size, bool) or not isinstance(size, int) or size < 0):
        raise ValueError("清单 size 必须是非负整数")
    return m


def _github_get(url, timeout=8, limit=None):
    req = urllib.request.Request(
        url, headers={"User-Agent": "BatchReplace/%s" % VERSION})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read(limit) if limit else resp.read()


def fetch_update_manifest(url=None, timeout=8):
    """取清单 → 字段校验。返回 (清单, 错误说明)。

    默认先走 API contents（国内可达），失败再退 raw；显式传入 url 时只打该地址。
    API 的 base64 envelope 与裸清单都能认——按有没有 version 字段区分。
    两条都失败时错误信息去重合并——通常两边报的是同一个网络故障。
    """
    targets = [url] if url else [MANIFEST_API, UPDATE_MANIFEST]
    errors = []
    for target in targets:
        cap = UPDATE_MAX_MANIFEST * (2 if target == MANIFEST_API else 1)
        try:
            raw = _github_get(target, timeout, cap + 1)
            if len(raw) > cap:
                raise ValueError("清单异常大")
            m = json.loads(raw.decode("utf-8"))
            if isinstance(m, dict) and "version" not in m and "content" in m:
                m = json.loads(base64.b64decode(m["content"]))
            validate_update_manifest(m)
            return m, None
        except Exception as exc:
            errors.append(str(exc))
    return None, "；".join(dict.fromkeys(errors))


def download_verified_exe(manifest, dest_dir, on_progress=None, timeout=30):
    """下载到 dest_dir 的 .part 文件，边下边验魔数/大小，收尾验 SHA-256；
    通过后 os.replace 到清单指定的文件名。返回 (最终路径, 错误说明)。"""
    name = manifest["exe"]
    expected_sha = manifest["sha256"].lower()
    expected_size = manifest.get("size")
    part = os.path.join(dest_dir, name + ".part")
    final = os.path.join(dest_dir, name)
    url = manifest.get("url") or ""
    if not url:
        return None, "清单缺少下载地址"
    host = (urllib.parse.urlsplit(url).hostname or "").lower()
    if host not in _LOOPBACK_HOSTS and not url.startswith("https://"):
        return None, "下载地址必须是 https"
    hasher = hashlib.sha256()
    state = {"n": 0, "head": b"", "err": None}

    def fail(reason):
        if state["err"] is None:
            state["err"] = reason

    try:
        os.makedirs(dest_dir, exist_ok=True)
        req = urllib.request.Request(url, headers={"User-Agent": "BatchReplace/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp, open(part, "wb") as out:
            while True:
                block = resp.read(256 * 1024)
                if not block:
                    break
                if state["err"]:
                    break
                if len(state["head"]) < 2:
                    state["head"] = (state["head"] + block[:2])[:2]
                    if state["head"] != b"MZ":
                        fail("下载内容不是 exe（缺少 MZ 魔数）")
                        break
                state["n"] += len(block)
                if state["n"] > UPDATE_MAX_EXE:
                    fail("文件超过大小上限")
                    break
                out.write(block)
                hasher.update(block)
                if on_progress:
                    on_progress(state["n"], expected_size or -1)
        if state["err"]:
            return None, state["err"]
        if expected_size is not None and state["n"] != expected_size:
            return None, "大小与清单不符（收到 %d，清单 %s）" % (state["n"], expected_size)
        if hasher.hexdigest() != expected_sha:
            return None, "SHA-256 校验不符"
        os.replace(part, final)
        return final, None
    except Exception as exc:
        return None, str(exc)
    finally:
        if os.path.exists(part):
            try:
                os.remove(part)
            except OSError:
                pass


def launch_verified_exe(path):
    """启动已校验的新版 exe（不走 shell）。返回 (是否成功, 说明)。"""
    try:
        subprocess.Popen([path], close_fds=True)
        return True, ""
    except Exception as exc:
        return False, str(exc)


def run_one_click_update(manifest_url=None, log=print, dry_run=False, current=VERSION):
    """一键更新：清单校验 → 版本比较 → 下载校验 → （非 dry_run）启动新版。"""
    log("检查更新：正在获取版本清单…")
    m, err = fetch_update_manifest(manifest_url)
    if err:
        return "失败：%s" % err
    latest = m["version"]
    if compare_versions(latest, current) <= 0:
        return "当前已是最新版 v%s。" % current
    log("发现新版本 %s（当前 v%s），开始下载…" % (latest, current))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    path, err = download_verified_exe(m, exe_dir)
    if err:
        return "失败：%s" % err
    log("下载完成并通过 SHA-256 校验：%s" % path)
    if dry_run:
        return "试运行完成：已下载并校验 %s（未启动）。" % path
    ok, err = launch_verified_exe(path)
    if not ok:
        return "失败：新版已下载并校验，但启动失败：%s（文件：%s）" % (err, path)
    return "已启动新版本 %s；新版文件已下载到同一目录。" % latest


def silent_update_target(tag, url, err, current=VERSION):
    """静默落点：有新版返回发布页，其余返回空串（不弹窗、不打开）。"""
    if err or not tag or compare_versions(tag, current) <= 0:
        return ""
    return url or RELEASE_PAGE_URL


def fetch_latest_release(timeout=8):
    """取最新版本号与链接，返回 (tag, url, err)。

    优先读更新清单（API contents 主路、raw 备胎）；失败则降级
    releases/latest，再读 tags，兼容只有 tag 的仓库。
    """
    for manifest in (MANIFEST_API, UPDATE_MANIFEST):
        try:
            payload = _github_get(manifest, timeout).decode("utf-8")
            if manifest == MANIFEST_API:
                envelope = json.loads(payload)
                info = json.loads(base64.b64decode(envelope["content"]).decode("utf-8"))
            else:
                info = json.loads(payload)
            tag = str((info or {}).get("version") or "").strip()
            page = str((info or {}).get("page") or RELEASE_PAGE_URL).strip()
            if tag and version_tuple(tag) != (0, 0, 0):
                return tag, page, None
        except (urllib.error.URLError, OSError, ValueError, KeyError,
                AttributeError, TypeError):
            continue
    err = ""
    try:
        data = json.loads(_github_get(UPDATE_API, timeout).decode("utf-8"))
        tag = str((data or {}).get("tag_name") or "").strip()
        if tag and version_tuple(tag) != (0, 0, 0):
            url = str((data or {}).get("html_url") or RELEASE_PAGE_URL).strip()
            return tag, url, None
        err = "release page returned no version"
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            err = "no published release"
        else:
            return "", "", str(exc)
    except (urllib.error.URLError, OSError, ValueError, KeyError, AttributeError) as exc:
        return "", "", str(exc)

    try:
        data = json.loads(_github_get(TAGS_API, timeout).decode("utf-8"))
        tags = [str((item or {}).get("name") or "").strip()
                for item in data if isinstance(data, list)]
        tags = [t for t in tags if version_tuple(t) != (0, 0, 0)]
        if not tags:
            return "", "", err or "no version tag found"
        return max(tags, key=version_tuple), RELEASE_PAGE_URL, None
    except (urllib.error.URLError, OSError, ValueError, KeyError, AttributeError) as exc:
        return "", "", "%s; tags failed: %s" % (err, exc)


class BatchReplaceApp:
    # 本次搜索是否链式：预览窗与勾选行为都读它。放类级而不是只在 __init__
    # 里赋值——测试与自检有十几处按 __new__ 造最小实例，走 __init__ 的赋值
    # 到不了它们，各处就只能写 getattr 兜底，等于把"忘了初始化"藏起来。
    is_chain_mode = False

    def __init__(self, root, defer_show=False):
        self.root = root
        # 主题色板必须先铺：root 底色在标题之前就要用，而配置里的主题选项
        # 要到 load_config 才读出来——先铺默认（浅色），读后再按配置铺一次。
        self.theme = "light"
        self._apply_theme(self.theme)
        # 窗口装饰自绘（标题栏 + 窗口按钮 + 边缘热区）：标题要居中、栏底要随
        # 主题配色，系统标题栏两样都做不到。
        self.root.overrideredirect(True)
        self.root.title(f"BatchReplace v{VERSION}")
        self.root.geometry("980x700")
        # 显式拉回普通态：Windows 会记住上次退出时的窗口状态，用户试用过
        # 最大化后关闭，下次启动会直接以最大化登场——geometry 只设还原尺寸，
        # 拉不回已最大化的窗口。
        self.root.state("normal")
        # 只有打包入口要求延后显示：把刚掀开的窗口重新藏好，等 main() 里
        # 任务栏身份就位再 deiconify。测试 / 截图走默认值，行为与原来一致。
        if defer_show:
            self.root.withdraw()
        # 最小宽度取默认宽度：英文界面第二行实测需 ~940 逻辑像素，允许缩小
        # 会把最右按钮裁出窗口——看得见一半、点不着，比不能缩更糟糕。
        self.root.minsize(980, 560)
        _set_window_icon(self.root)
        # overrideredirect 重建 HWND 后 winfo_id() 暂留旧值；先跑 idle 让新句柄
        # 落地再登记任务栏身份，窗口首次可见时按钮就不用事后补。
        self.root.update_idletasks()
        _apply_taskbar_identity(self.root)
        # 无边框窗口的状态：_restore_frameless 为真表示"正在最小化，映射回
        # 来时要把无边框标志补上"；_maximized 记录是否处于最大化。
        self._restore_frameless = False
        self._maximized = False
        self._pre_max_geo = None
        self._pre_min_geo = None
        self._resize_zone = None
        self._activated = False
        self.root.bind("<Map>", self._on_root_map)
        # 边缘缩放绑在根窗口上：子控件的事件会冒泡上来，_edge_hit 负责
        # 只放行"窗口自身与背景框"上的边缘点击。
        self.root.bind("<Motion>", self._edge_motion)
        self.root.bind("<Button-1>", self._edge_press)
        self.root.bind("<B1-Motion>", self._edge_drag)
        self.root.bind("<ButtonRelease-1>", self._edge_release)

        self.selected_files = []
        self.preview_data = []
        self.find_history = []  # 查找历史（一组条件 = 一条记录）
        # 启动早期（日志框还没建）读配置 / 历史失败的信息先攒这里，见 setup_ui 尾部
        self.startup_notes = []

        # 工作线程 → 主线程的唯一 UI 通道。
        # 不能在工作线程操作控件或调 root.after——都在调用 Tcl 解释器，会触发
        # "Tcl_AsyncDelete: async handler deleted by the wrong thread"；安全
        # 做法只有纯 Python 队列 + 主线程定时排空。
        self._ui_queue = queue.Queue()
        self._ui_pump_id = None

        # 搜索重入守卫。
        # 置灰主按钮只是表面约束，挡不住别的触发路径在 worker 还在跑时再起
        # 一轮（两个 worker 竞争，预览可能显示旧结果）。用「线程还活着就
        # 拒绝 + 代号对不上即丢弃」双保险，不会把新搜索的界面改回"就绪"。
        self._search_thread = None
        self._search_gen = 0

        # 替换批次互斥，与搜索守卫同构：替换跑着时还允许再搜索，两个批次
        # 会同时开同一批文件（os.replace 撞权限错），关窗也会连锅端。这里
        # 记当前替换线程供各处查询。
        self._replace_thread = None

        self.load_config()
        self._apply_theme(self.theme)
        self.load_history()
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._ui_pump_id = self.root.after(40, self._pump_ui_queue)
        # 静默检查只给冻结版：源码版跑测试 / 调试时不会无故联网弹窗。
        if getattr(sys, "frozen", False) and not SKIP_STARTUP_UPDATE:
            self.root.after(600, self.silent_check_update)

        # 缺库时 .docx 会落入"无匹配"分支，与"真的没有匹配"无法区分（exe 版
        # 不受影响）——启动时说清楚，见 _check_dependencies。
        if not (DOCX_AVAILABLE and EXCEL_AVAILABLE):
            self.root.after(300, self._check_dependencies)

    # 语言唯一来源是模块级 CURRENT_LANG；用属性接到实例上，让 self.lang
    # 照旧可读写，避免两处各存一份而漂移（纯函数也读不到实例上那份）。
    @property
    def lang(self):
        return CURRENT_LANG

    @lang.setter
    def lang(self, value):
        global CURRENT_LANG
        CURRENT_LANG = value if value in ("zh", "en") else DEFAULT_LANG

    def _t(self, key, **kw):
        """取当前语言的文案，支持 {var} 格式化。找不到时回退中文再回退 key 本身。"""
        return T(key, **kw)

    def _reg_i18n(self, widget, key, prefix=""):
        """注册控件到 i18n 刷新表——切换语言时遍历它批量 config(text=...)。

        prefix 给「符号 + 文案」的按钮用（主按钮的 ▶）：切语言只换文案，
        符号留在原地；不带 prefix 则符号会悄悄没了（建 / 刷两处不一致）。
        """
        if not hasattr(self, "_i18n_widgets"):
            self._i18n_widgets = {}
        self._i18n_widgets[widget] = (key, prefix)
        if getattr(widget, "_br_repaint", None) is not None:
            widget._br_key = key

    def _reg_heading(self, tree, column, key):
        """注册 Treeview 表头——表头要走 heading(column, text=...)。

        不注册成控件：config(text=...) 对 Treeview 会抛 TclError，被兜底
        except 吞掉，表头静默停在旧语言。
        """
        if not hasattr(self, "_i18n_headings"):
            self._i18n_headings = []
        self._i18n_headings.append((tree, column, key))

    def switch_lang(self):
        """点击语言切换按钮：中↔英，存配置，刷新界面。"""
        self.lang = "en" if self.lang == "zh" else "zh"
        self.save_config()
        self.apply_lang()

    def apply_lang(self):
        """刷新所有静态文案控件到当前语言；动态文案（计数等）调对应更新方法。"""
        for widget, (key, prefix) in getattr(self, "_i18n_widgets", {}).items():
            try:
                if getattr(widget, "_br_key", None) is not None:
                    widget._br_text = f"{prefix}{self._t(key)}"
                    widget._br_repaint()
                else:
                    widget.config(text=f"{prefix}{self._t(key)}")
            except Exception:
                pass  # 控件可能已销毁（预览窗/对话框关闭后）
        for tree, column, key in getattr(self, "_i18n_headings", []):
            try:
                tree.heading(column, text=self._t(key))
            except Exception:
                pass
        # 动态文案单独刷新：这些位置按状态拼装，config 一次不够；一律复用
        # 原有更新方法，抄一份拼装逻辑就等于两个会漂移的实现。
        if hasattr(self, "status_label"):
            self.update_status()            # 状态栏「共 N 个文件 …」
        if hasattr(self, "checked_count_label"):
            self.update_checked_count()     # 预览窗「已选: N 项」
        if hasattr(self, "mode_label"):
            self._update_mode_hint()        # 「当前：替换文件内容」等四种组合
        if hasattr(self, "log_arrow"):
            self._refresh_log_arrow()
        if hasattr(self, "replace_label"):
            self._refresh_find_replace_labels()
        self._sync_multi_mode_texts()
        # 悬停说明窗是一次性写入文案的 Toplevel，不在静态控件注册表里；
        # 正开着时按新语言重建，否则会停在切换前的语言。
        if getattr(self, "_chain_tip", None) is not None:
            self._hide_chain_tip()
            self._chain_tip_show()
        if hasattr(self, "progress_label"):
            self.progress_label.config(text="")   # 进度文案按当前任务重写，先清空

        self._refresh_file_status_texts()

        if hasattr(self, "lang_btn"):
            # 键随当前语言二选一，不走固定键注册表；圆角按钮要改文字后重绘
            self.lang_btn._br_text = self._t(
                "lang_to_en" if self.lang == "zh" else "lang_to_zh")
            self.lang_btn._br_repaint()

    def _refresh_find_replace_labels(self):
        """参数区两个标签的文案随「替换文件名 / 替换内容」切换而变化。

        宽度一起设：文案长度两种模式不同，不设宽度整行会左右跳；宽度按语言
        实测（见 _fit_label_width），换个语言也不会被截断。
        """
        if getattr(self, "find_label", None) is not None:
            self.find_label.config(
                text=self._t("find_filename" if self.replace_filename_var.get() else "find_content"),
                width=self._fit_label_width(("find_content", "find_filename")))
        if getattr(self, "replace_label", None) is not None:
            self.replace_label.config(
                text=self._t("new_filename" if self.replace_filename_var.get() else "replace_with"),
                width=self._fit_label_width(("replace_with", "new_filename")))

    def _refresh_log_arrow(self):
        """日志折叠条：箭头方向与文案都反映当前展开状态。"""
        if getattr(self, "log_arrow", None) is None:
            return
        arrow = "\u25be" if self.log_visible else "\u25b8"
        self.log_arrow.config(text=f"{arrow}  {self._t('process_log')}")

    # 文件列表「状态」列的文案键。表格里存的是显示串（Treeview 没有别的
    # 存放处），切语言时只能按已知文案反查回键，故两个语言的写法都列上。
    _STATUS_KEYS = ("st_not_searched", "st_pending", "st_no_match", "st_replaced")

    def _status_text(self, status):
        """文件列表「状态」列的显示串 → 当前语言。已写入的行也要跟着切语言。"""
        if not status:
            return status
        for key in self._STATUS_KEYS:
            entry = TEXTS.get(key, {})
            if status in (entry.get("zh"), entry.get("en")):
                return self._t(key)
        return status

    def _refresh_file_status_texts(self):
        """把文件列表里已写入的状态文字按当前语言重写一遍。"""
        tree = getattr(self, "files_tree", None)
        if tree is None:
            return
        for iid in tree.get_children():
            try:
                self.files_tree.set(iid, "status", self._status_text(tree.set(iid, "status")))
            except tk.TclError:
                pass

    def _var(self, kind, **kw):
        """建 Tk 变量并显式绑到本实例的根窗口。

        不给 master 时 tkinter 绑的是 `_default_root`：同一进程里存在多个 Tk
        实例（测试并发建窗、将来开第二个窗口）时，变量会落到别的解释器上，
        trace 不回调、值各读各的。校验型最小实例（root 是替身对象）没有真
        窗口，仍然退回默认行为。
        """
        if isinstance(getattr(self, "root", None), tk.Misc):
            kw["master"] = self.root
        return kind(**kw)

    def load_config(self):
        # 语言与主题每次启动都取系统当前值，不读配置文件：两者都能在界面上
        # 一键切换，跟随系统比记住上一次的手动选择更符合预期（换机器、白天
        # 黑夜切换都不用回头改设置）。
        self.lang = DEFAULT_LANG
        self.theme = detect_default_theme()
        self.backup_var = self._var(tk.BooleanVar, value=True)
        # 默认备份目录是磁盘位置，不随语言切换（切换语言不该换备份去处）
        self.backup_path = self._var(tk.StringVar, value=DEFAULT_BACKUP_DIR)
        self.case_sensitive_var = self._var(tk.BooleanVar, value=False)
        self.replace_filename_var = self._var(tk.BooleanVar, value=False)
        # 多条件替换：条件表内容（[[查找, 替换], ...]）与开关。开关不进配置
        # 文件——每次启动一律从单条件开始，避免与这次要干的活对不上。
        self.multi_condition_var = self._var(tk.BooleanVar, value=False)
        # 多条件的执行方式，内部值不随界面语言变（下拉框显示的是翻译串）：
        # parallel＝全部条件在原文上一次替换；chain＝上一条的输出是下一条的输入。
        self._multi_mode = "parallel"
        self.multi_mode_var = self._var(tk.StringVar,
                                        value=self._t("multi_mode_parallel"))
        # 悬停说明的延时句柄与窗体：句柄先建好，事件来了才不会 AttributeError
        self._chain_tip_job = None
        self._chain_tip = None
        self.conditions = []
        # 预览窗操作提示条是否显示。只在本次会话内有效：菜单栏已删除，该
        # 选项没有第二个入口，持久化会变成关掉就找不回来的隐藏状态。
        self.preview_hint_var = self._var(tk.BooleanVar, value=True)

        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    # 备份是安全底线：不进配置文件，每次启动一律强制开启，避免上次关闭后本次忘记开启
                    self.backup_var.set(True)
                    self.backup_path.set(migrate_backup_path(
                        config.get("backup_path", DEFAULT_BACKUP_DIR)))
                    self.case_sensitive_var.set(config.get("case_sensitive", False))
            except Exception as e:
                # 读失败不能让启动挂掉，也不能无声无息——只攒「键 + 参数」，
                # 日志框建好再按当时的语言渲染（此刻语言可能还没读出来）。
                self.startup_notes.append(("note_config_failed", {"e": e}))

        if not os.path.exists(RULES_DIR):
            try:
                os.makedirs(RULES_DIR)
            except Exception:
                pass
        # 默认备份目录要在首次使用前建出来：已有机器上没这个目录，「打开备份
        # 文件夹」会被拦下。只建默认值——自选路径不存在仍按原样提示。
        if self.backup_path.get() == DEFAULT_BACKUP_DIR and \
                not os.path.exists(DEFAULT_BACKUP_DIR):
            try:
                os.makedirs(DEFAULT_BACKUP_DIR)
            except Exception:
                pass

    def save_config(self):
        try:
            config = {
                "backup_path": self.backup_path.get(),
                "case_sensitive": self.case_sensitive_var.get(),
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(T("note_config_save_failed", e=e))

    def on_closing(self):
        # 替换批次进行中关窗会把批次连锅端——先问一句，把后果说清楚再让用户决定。
        if self._replace_in_progress():
            if not messagebox.askyesno(
                    self._t("t_close"),
                    self._t("m_close_confirm"),
                    parent=self.root):
                return
        self.save_config()
        if getattr(self, "_ui_pump_id", None) is not None:
            try:
                self.root.after_cancel(self._ui_pump_id)
            except Exception:
                pass
            self._ui_pump_id = None
        self.root.destroy()

    def _check_dependencies(self):
        """依赖缺失时明确告知，不再静默走"无匹配"。

        没装 python-docx 时 .docx 会落入"无匹配"分支，与"真的没有匹配"完全
        无法区分（exe 版不受影响）。这里缺哪个库就说清楚，并给出两条出路。
        """
        missing = []
        if not DOCX_AVAILABLE:
            missing.append(("python-docx", ".docx"))
        if not EXCEL_AVAILABLE:
            missing.append(("openpyxl", ".xlsx"))
        if not missing:
            return
        items = "\n".join(self._t("m_dep_item", lib=lib, ext=ext) for lib, ext in missing)
        cmd = "pip install " + " ".join(lib for lib, _ in missing)
        try:
            messagebox.showinfo(self._t("t_missing_dep"),
                                self._t("m_dep_missing", items=items, cmd=cmd),
                                parent=getattr(self, "root", None))
        except tk.TclError:
            pass  # 窗口已销毁，忽略

    # ==================== 工作线程 → 主线程的 UI 通道 ====================
    def _ui(self, fn, *args, **kwargs):
        """把一个 UI 操作从工作线程投递回主线程执行。

        工作线程里凡碰界面（写日志、改按钮、弹框等）都必须走 self._ui(...)，
        绝不允许直接调用。
        """
        self._ui_queue.put((fn, args, kwargs))

    def _pump_ui_queue(self):
        """主线程心跳：把工作线程投递过来的操作逐个执行掉（40ms 一次，空转开销极小）。"""
        try:
            while True:
                try:
                    fn, args, kwargs = self._ui_queue.get_nowait()
                except queue.Empty:
                    break
                try:
                    fn(*args, **kwargs)
                except tk.TclError:
                    pass  # 控件已随窗口销毁，丢掉即可
                except Exception as e:
                    try:
                        self.log(self._t("lg_ui_update_failed", e=e))
                    except Exception:
                        pass
        finally:
            try:
                self._ui_pump_id = self.root.after(40, self._pump_ui_queue)
            except tk.TclError:
                self._ui_pump_id = None  # 窗口已销毁，停止心跳

    # 主题色表（深浅双主题）。色值唯一来源——控件创建与重绘都读实例属性，
    # 由 _apply_theme 从本表铺到实例上。列表底 / 选中 / 表头 / 悬停 / 面板
    # 分档明确，正文 13px、小字 12px。
    THEMES = {
        "light": {
            "BG": "#CECECE", "SURFACE": "#F8FAFC", "PANEL": "#DDDDDD",
            "ZEBRA": "#F1F4F8", "LINE": "#DDDDDD", "BORDER": "#808080",
            "BORDER_STRONG": "#5E5E5E", "HEAD_BG": "#D5DFE9", "SEL_BG": "#C5D5E6",
            "TXT": "#111111", "TXT_DIM": "#888888",
            "INFO_FG": "#224488", "HINT_BG": "#DDDDDD", "HINT_CLOSE_FG": "#808080",
            # 链式预览里"上一轮改出来的"命中：浅底强调蓝，与悬停蓝、斑马灰都分得开
            "HOVER_BG": "#DCE5EF", "DANGER_FG": "#CC2222",
            "PRODUCED_BG": "#D6E4FF",
            # 预览上下文列里命中文本的着色：琥珀色与蓝色系主色拉开差距，
            # 浅底（含斑马纹 / 悬停蓝 / 链式蓝底）上都够清楚
            "HIT_FG": "#B45309",
            "ACCENT": "#2563EB", "ACCENT_HOVER": "#1D4ED8", "ACCENT_FG": "#FFFFFF",
            "ACCENT_OFF_BG": "#C3CCD9", "ACCENT_OFF_FG": "#F4F6F9",
            "BAR_TROUGH": "#F8FAFC", "BAR_FILL": "#22AA22", "BAR_EDGE": "#808080",
            "BTN_BG": "#DDDDDD", "BTN_HOVER_BG": "#CFCFCF",
            "BTN_BD": "#808080", "BTN_BD_HOVER": "#5E5E5E",
            "BTN_OFF_BG": "#CECECE", "BTN_OFF_FG": "#888888", "BTN_OFF_BD": "#AAAAAA",
            "GRID": "#D2D2D2",
            # 自绘标题栏：窗口钮悬停底色（略深于 BG）/ 关闭钮悬停红（Windows 标准）
            "TITLE_HOVER": "#BEBEBE", "CLOSE_HOVER": "#C42B1C",
        },
        "dark": {
            "BG": "#353535", "SURFACE": "#222222", "PANEL": "#2A2A2A",
            "ZEBRA": "#282828", "LINE": "#222222", "BORDER": "#555555",
            "BORDER_STRONG": "#666666", "HEAD_BG": "#111111", "SEL_BG": "#444444",
            "TXT": "#F0F0F0", "TXT_DIM": "#888888",
            "INFO_FG": "#6699CC", "HINT_BG": "#222222", "HINT_CLOSE_FG": "#888888",
            "HOVER_BG": "#333333", "DANGER_FG": "#E74C3C",
            "PRODUCED_BG": "#2C3A52",
            # 深色底上的琥珀亮色：与 #F0F0F0 正文、#6699CC 信息蓝都分得开
            "HIT_FG": "#FBBF24",
            "ACCENT": "#2563EB", "ACCENT_HOVER": "#3B82F6", "ACCENT_FG": "#FFFFFF",
            "ACCENT_OFF_BG": "#3A4150", "ACCENT_OFF_FG": "#8A93A0",
            "BAR_TROUGH": "#111111", "BAR_FILL": "#22AA22", "BAR_EDGE": "#555555",
            "BTN_BG": "#3E3E3E", "BTN_HOVER_BG": "#555555",
            "BTN_BD": "#222222", "BTN_BD_HOVER": "#666666",
            "BTN_OFF_BG": "#2A2A2A", "BTN_OFF_FG": "#666666", "BTN_OFF_BD": "#3A3A3A",
            "GRID": "#414141",
            "TITLE_HOVER": "#4A4A4A", "CLOSE_HOVER": "#C42B1C",
        },
    }

    def _apply_theme(self, theme):
        """把主题色表铺到实例属性：控件创建与自绘重绘读的都是实例值。

        root 底色一并铺：框架间隙（列表四周的 padx）会直接透出它，
        只铺属性不铺窗口，切主题后窗口边缘仍留着上一个主题的底色。
        """
        for key, value in self.THEMES.get(theme, self.THEMES["light"]).items():
            setattr(self, key, value)
        self.root.configure(bg=self.BG)
        # 勾选图标按主题色板预渲染并缓存（见 _check_image）。切主题不清缓存，
        # 旧主题的图标会挂在重绘后的控件上——深色主题下未选框填 #222222，
        # 切到浅色窗口后就是"白底黑框"。这里在铺色后立即把缓存作废，下一步
        # setup_ui 重建控件时 _check_image 会用新色板重画。
        for _cache in ("_check_icons", "_check_icons_off"):
            if hasattr(self, _cache):
                delattr(self, _cache)

    # 列表行高（唯一来源：样式与"行尾小标签要塞得进一行"都看它）
    ROW_HEIGHT = 24
    # 自绘标题栏高度与窗口按钮规格（按钮宽占满栏高，悬停整块浮色）
    TITLEBAR_H = 32
    WIN_BTN_W = 44
    # 无边框窗口的边缘热区宽度：原生 resize 边框随装饰一起没了，自绘
    _EDGE = 6
    _EDGE_CURSORS = {"n": "size_ns", "s": "size_ns", "e": "size_we",
                     "w": "size_we", "ne": "size_nesw", "sw": "size_nesw",
                     "nw": "size_nwse", "se": "size_nwse"}
    # 状态栏右下角进度条自绘规格（Canvas 绘制，非 ttk.Progressbar）
    BAR_W = 170
    BAR_H = 15
    # 悬浮说明弹窗（tooltip）默认存活时长：弹出后即使无人理会也自行收掉，
    # 免得长期挂在界面上挡视线。执行方式说明、历史表头说明两处共用。
    TIP_AUTO_CLOSE_MS = 10_000

    # ttk 样式名：在 _setup_style 里注册、在控件处引用。写散了改主题名就得
    # 全文件替换，漏一处的表现是"某个控件变回系统灰"。按钮不在其列——全部自绘。
    TREE_STYLE = "Z.Treeview"

    # 预览窗口"双击"相关文案已全部搬进 TEXTS（chip_text / hint_bar_content 等）；
    # 术语必须一致——表头、行尾标签与底部说明指向同一个动作，都叫"跳转"。

    def setup_ui(self):
        """构建主界面：工具栏 → 参数栏 → 文件列表 → 日志 → 状态栏

        配色走「中性浅灰底 + 单一强调蓝」；按钮全部自绘（tk.Button）——ttk
        按钮在 Windows 原生主题下由系统绘制，配色与描边改不动。

        菜单栏已删除：菜单文案要跟着语言切换，是唯一没法靠 apply_lang 重刷
        的区域，除去它少一套需维护的文案与状态同步代码。
        """
        self.log_visible = False
        self.log_user_collapsed = False
        self._total_matches = 0
        self._i18n_widgets = {}
        self._i18n_headings = []
        tkfont.nametofont("TkDefaultFont").configure(family="Segoe UI", size=10)
        self._setup_style()
        self._build_titlebar()
        self._build_toolbar()
        self._build_path_bar()
        # 参数栏与文件列表共用一块画布，按行分高度：pack 做不到比例分配——
        # 谁先 pack 谁先占满，后 pack 的会被压到零。多条件模式下条件表需要
        # 随窗口长高，只能靠 grid 的行权重让两者各分一半富余高度。
        self.body = tk.Frame(self.root, bg=self.BG)
        self.body.pack(fill="both", expand=True, side="top")
        # 关掉尺寸传播：body 的请求高度 = 子控件之和（多条件时条件表 + 文件列表
        # 会超过窗口），pack 会因此把底部日志条与状态栏挤出窗口。让它只取分配到的
        # 空间，内部分配交给 grid 的行权重。
        self.body.grid_propagate(False)
        self.body.columnconfigure(0, weight=1)
        self.body.rowconfigure(2, weight=1)
        self._build_param_bar()
        self._build_status_bar()   # 先占底部：状态栏在最下
        self._build_log_panel()    # 再占底部：日志折叠条位于状态栏之上
        self._build_file_table()   # 中间列表填满剩余空间
        self.update_status()
        self._set_log_visible(True)  # 日志默认展开
        # 配置 / 历史在控件之前读，没地方记录失败也不知语言（键按「文案键 +
        # 参数」存着）。有内容就补报——静默只会让用户某天发现设置没生效。
        for key, kw in self.startup_notes:
            self.log(self._t(key, **kw))

    # ---------- 对话框摆位 ----------
    def _center_dialog(self, dlg, parent=None):
        """把对话框摆到父窗口正中，并确保它显示在父窗口之上。

        位置仍要显式给：Tk 的 Toplevel 默认由窗口管理器摆放，落在父窗口
        左上角附近而不是正中。必须在控件建完后调用——尺寸要按最终布局算，
        早调会把还没 pack 的控件漏掉。

        geometry 之后补一次 update 与 lift：update 让窗口完成映射（后面紧跟
        wait_window 时才不会出现「框还没显示就先挂起」），lift 兜住
        transient 尚未生效的一瞬。update 期间外部挂的定时回调可能先跑掉，
        所以最后一步要先确认窗口还在。

        全部对话框都从这条路走，原生标题栏跟随应用主题也在这里统一打点。
        """
        dlg.update_idletasks()
        parent = parent or self.root
        w, h = dlg.winfo_width(), dlg.winfo_height()
        if parent.winfo_viewable():
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        else:
            # 主窗口最小化 / 收进托盘时它的坐标是陈的，按屏幕居中
            x = (parent.winfo_screenwidth() - w) // 2
            y = (parent.winfo_screenheight() - h) // 2
        # 窗比父窗大时别推出屏幕外
        x = max(0, min(x, parent.winfo_screenwidth() - w))
        y = max(0, min(y, parent.winfo_screenheight() - h))
        dlg.geometry("+%d+%d" % (x, y))
        dlg.update()
        if dlg.winfo_exists():
            # DWM 深浅标志必须等映射完成后再打：映射前设置的值会被首次映射
            # 重置（实测深色界面顶白栏）。走到这里 update 已完成映射。
            _apply_dialog_titlebar_theme(dlg, self.theme == "dark")
            dlg.lift()

    @staticmethod
    def _wait_dialog(dlg):
        """等对话框关掉再往下走。

        tkwait 要求窗口存在：对话框若已在 _center_dialog 的 update（事件
        刷新）期间被外部定时回调销毁，直接 wait_window 会抛 bad window
        path name。先问一句还在不在，已经不在了就没什么可等的。
        """
        if dlg.winfo_exists():
            dlg.wait_window()

    # ---------- 自绘标题栏（窗口装饰） ----------
    def _build_titlebar(self):
        """标题栏：应用名居中，右侧三个窗口按钮（最小化·最大化·关闭）。

        栏底 / 按钮底 / 文字色全部取当前主题色板，切主题重建界面时自然跟随。
        拖拽移动 / 双击最大化绑在栏本体与标题上；三个按钮是独立控件，不挂
        拖拽——在按钮上按下不该顺带拖动窗口。
        """
        bar = tk.Frame(self.root, bg=self.BG, height=self.TITLEBAR_H)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)
        self._titlebar_frame = bar
        self._title_label = tk.Label(
            bar, text=f"BatchReplace v{VERSION}",
            bg=self.BG, fg=self.TXT_DIM, anchor="center")
        self._title_label.place(relx=0.5, rely=0.5, anchor="center")
        for w in (bar, self._title_label):
            w.bind("<Button-1>", self._titlebar_press)
            w.bind("<B1-Motion>", self._titlebar_drag)
            w.bind("<Double-Button-1>", lambda _e: self._toggle_maximize())
        for i, (kind, cmd) in enumerate((("close", self.on_closing),
                                         ("max", self._toggle_maximize),
                                         ("min", self._minimize_window))):
            self._make_window_button(bar, kind, cmd).place(
                relx=1.0, x=-(i + 1) * self.WIN_BTN_W, y=0,
                width=self.WIN_BTN_W, height=self.TITLEBAR_H)
        tk.Frame(self.root, bg=self.LINE, height=1).pack(fill="x", side="top")

    def _make_window_button(self, parent, kind, command):
        """标题栏窗口按钮：扁平无边框，悬停浮出底色，关闭悬停红底白字。

        符号画成位图而不用字符：字形由 GDI 渲染，小字号带 ClearType 彩边，
        而窗口本身不感知 DPI，整块界面被系统放大后彩边与字重不匀会跟着
        放大，看着发脏；位图放大只是线条变粗。
        """
        hover_bg = self.CLOSE_HOVER if kind == "close" else self.TITLE_HOVER
        btn = tk.Button(parent, command=command, relief="flat", bd=0,
                        highlightthickness=0, takefocus=0, bg=self.BG,
                        activebackground=hover_bg)
        btn._kind = kind
        btn._img = self._symbol_image(kind, self.TXT, self.BG)
        btn._img_hover = self._symbol_image(
            kind, "#FFFFFF" if kind == "close" else self.TXT, hover_bg)
        btn.configure(image=btn._img)
        btn.bind("<Enter>",
                 lambda _e: btn.configure(bg=hover_bg, image=btn._img_hover))
        btn.bind("<Leave>", lambda _e: btn.configure(bg=self.BG, image=btn._img))
        return btn

    def _symbol_image(self, kind, fg, bg):
        """画 1px 线宽的窗口符号（− □ ✕）：4 倍超采样再缩回，斜线才不锯齿。"""
        size, ss = 10, 4
        big = size * ss
        img = Image.new("RGB", (big, big), bg)
        d = ImageDraw.Draw(img)
        lo, hi = ss, big - ss - 1
        if kind == "min":
            d.line([(lo, big // 2), (hi, big // 2)], fill=fg, width=ss)
        elif kind == "max":
            d.rectangle([lo, lo, hi, hi], outline=fg, width=ss)
        else:
            d.line([(lo, lo), (hi, hi)], fill=fg, width=ss)
            d.line([(hi, lo), (lo, hi)], fill=fg, width=ss)
        return ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))

    def _titlebar_press(self, e):
        self._drag_x, self._drag_y = e.x_root, e.y_root

    def _titlebar_drag(self, e):
        dx, dy = e.x_root - self._drag_x, e.y_root - self._drag_y
        if not (dx or dy):
            return
        if self._maximized:
            # 从最大化拖出：先还原，再把标题栏拉到光标下
            self._unmaximize()
            self.root.geometry("+%d+%d" % (
                e.x_root - self.root.winfo_width() // 2,
                e.y_root - self.TITLEBAR_H // 2))
        else:
            self.root.geometry("+%d+%d" % (self.root.winfo_x() + dx,
                                           self.root.winfo_y() + dy))
        self._drag_x, self._drag_y = e.x_root, e.y_root

    def _toggle_maximize(self):
        if self._maximized:
            self._unmaximize()
        else:
            self._pre_max_geo = self.root.geometry()
            # 工作区矩形（排除任务栏）：铺满它而不是整个屏幕，任务栏仍可见
            rect = _WorkArea()
            ctypes.windll.user32.SystemParametersInfoW(
                0x30, 0, ctypes.byref(rect), 0)
            self.root.geometry("%dx%d+%d+%d" % (
                rect.right - rect.left, rect.bottom - rect.top,
                rect.left, rect.top))
            self._maximized = True

    def _unmaximize(self):
        if self._pre_max_geo:
            self.root.geometry(self._pre_max_geo)
        self._maximized = False

    def _minimize_window(self):
        """最小化：Tcl 的 iconify 拒绝无边框窗口，先临时交出无边框标志。

        最小化期间窗口是普通样式——任务栏上有按钮点得回来；映射回来时
        _on_root_map 再把无边框标志补上，并写回最小化前的几何：那圈原生
        边框的非客户区会被 Windows 并进窗口大小，不写回客户区就逐次变胖。
        """
        self._pre_min_geo = self.root.geometry()
        self.root.overrideredirect(False)
        # 交还无边框标志会重建 HWND：任务栏身份与图标要立刻补到新窗口上
        _apply_taskbar_identity(self.root)
        self.root.iconify()
        self._restore_frameless = True

    def _on_root_map(self, e):
        """窗口每次映射后：补任务栏身份；从最小化还原时再把无边框标志补上。

        必须判 e.widget：root 上绑的 <Map> 会收到所有子控件的映射事件，不判
        就会在建界面时被回调上百次，每次重设扩展样式与图标——任务栏按钮因此
        迟迟不落地，图标也要等这一串刷完才出现。
        最小化态下不补无边框：那期间必须是普通窗口，任务栏按钮才点得到。
        判据用 Tk 自己的 state 而非 Win32 IsIconic——后者有处理延迟，
        「正在最小化」的瞬间它会报 False，被误判成已还原而抢补无边框标志，
        最小化就被撤掉了。
        """
        if e.widget is not self.root:
            return
        if self.root.state() == "iconic":
            return
        _apply_taskbar_identity(self.root)
        if not getattr(self, "_activated", False):
            self._activated = True
            # 无边框窗口不会被系统自动激活，任务栏按钮要等用户点一下才落地
            self.root.after(30, self._activate_window)
        if not self._restore_frameless:
            return
        self._restore_frameless = False
        if not self.root.overrideredirect():
            self.root.overrideredirect(True)
            _apply_taskbar_identity(self.root)
        if self._pre_min_geo:
            self.root.geometry(self._pre_min_geo)

    def _activate_window(self):
        """把主窗置为前台。任务栏按钮在窗口被激活时才落地，不主动激活就要等
        用户点一下窗口，看起来就是「图标迟迟不出现」。

        前台锁会把 SetForegroundWindow 直接拒掉（返回 0），先与当前前台线程
        挂接输入再设，才稳。
        """
        if os.name != "nt":
            return
        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            user32.GetAncestor.restype = ctypes.c_void_p
            user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
            hwnd = user32.GetAncestor(ctypes.c_void_p(self.root.winfo_id()), 2)
            if not hwnd:
                return
            self.root.focus_force()
            user32.GetWindowThreadProcessId.restype = ctypes.c_uint32
            fg = user32.GetForegroundWindow()
            if fg == hwnd:
                return
            tid_fg = user32.GetWindowThreadProcessId(ctypes.c_void_p(fg), None)
            tid_me = kernel32.GetCurrentThreadId()
            attached = bool(tid_fg) and tid_fg != tid_me
            if attached:
                user32.AttachThreadInput(tid_me, tid_fg, True)
            try:
                user32.SetForegroundWindow(ctypes.c_void_p(hwnd))
            finally:
                if attached:
                    user32.AttachThreadInput(tid_me, tid_fg, False)
        except Exception:
            pass

    # ---------- 边缘缩放（无边框窗口没有系统 resize 边框） ----------
    def _edge_hit(self, e):
        """命中边缘 / 角落返回方向串，否则 None。

        只认窗口自身与背景框：树视图 / 滚动条 / 按钮 / 输入框上的点击
        不抢——在列表里拖选不该变成拖窗口边框。
        """
        if e.widget is not self.root and not isinstance(e.widget, tk.Frame):
            return None
        ox, oy = self.root.winfo_rootx(), self.root.winfo_rooty()
        x, y = e.x_root - ox, e.y_root - oy
        w, h = self.root.winfo_width(), self.root.winfo_height()
        # 标题栏区域不参与上边缘：那 32px 归拖拽移动管
        if y <= self.TITLEBAR_H:
            return None
        parts = []
        if x <= self._EDGE:
            parts.append("w")
        elif x >= w - self._EDGE:
            parts.append("e")
        if y >= h - self._EDGE:
            parts.append("s")
        elif y <= self.TITLEBAR_H + self._EDGE:
            parts.append("n")
        return "".join(parts) or None

    def _edge_motion(self, e):
        self.root.configure(cursor=self._EDGE_CURSORS.get(
            self._edge_hit(e) or "", ""))

    def _edge_press(self, e):
        zone = self._edge_hit(e)
        self._resize_zone = zone
        if zone:
            self._resize_x, self._resize_y = e.x_root, e.y_root
            self._resize_pos = (self.root.winfo_x(), self.root.winfo_y())
            self._resize_wh = (self.root.winfo_width(), self.root.winfo_height())

    def _edge_drag(self, e):
        zone = getattr(self, "_resize_zone", None)
        if not zone:
            return
        dx, dy = e.x_root - self._resize_x, e.y_root - self._resize_y
        x, y = self._resize_pos
        w, h = self._resize_wh
        if "e" in zone:
            w += dx
        if "w" in zone:
            w -= dx
            x += dx
        if "s" in zone:
            h += dy
        if "n" in zone:
            h -= dy
            y += dy
        min_w, min_h = self.root.winfo_reqwidth(), 560
        if w < min_w:
            if "w" in zone:
                x -= min_w - w
            w = min_w
        if h < min_h:
            if "n" in zone:
                y -= min_h - h
            h = min_h
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def _edge_release(self, _e):
        self._resize_zone = None

    def _setup_style(self):
        # Style 必须绑当前根窗口：不传 master 时它绑"默认根窗口"，多窗口场景下可能
        # 是别的解释器，而我们传给 element_create 的图标属于这个根窗口——两边对不上
        # 就报 image doesn't exist。
        style = ttk.Style(master=self.root)
        # clam 而非 vista：vista 的 ttk 控件由系统原生绘制，Treeview / Entry /
        # Checkbutton 的配色配置大面积被忽略（深色主题下列表白底黑字、输入框
        # 贴白块）。clam 完全吃样式配置，是深浅双主题的唯一可行主题。
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        # Treeview 选中色必须走 map：configure 的 background 会被选中态默认值盖掉。
        style.configure(self.TREE_STYLE, background=self.SURFACE,
                        foreground=self.TXT, fieldbackground=self.SURFACE,
                        bordercolor=self.BORDER, lightcolor=self.SURFACE,
                        darkcolor=self.BORDER, rowheight=self.ROW_HEIGHT,
                        font=("Segoe UI", 10))
        style.map(self.TREE_STYLE,
                  background=[("selected", self.SEL_BG)],
                  foreground=[("selected", self.TXT)])
        style.configure("Treeview.Heading", background=self.HEAD_BG,
                        foreground=self.TXT_DIM, relief="flat",
                        font=("Segoe UI", 10, "bold"))
        # 勾选框不在其列——已全部改为自绘（_make_check），ttk 那套
        # TCheckbutton 配置与图标元素没有消费者了。
        style.configure("TEntry", fieldbackground=self.SURFACE,
                        bordercolor=self.BORDER, foreground=self.TXT,
                        lightcolor=self.BORDER, darkcolor=self.BORDER,
                        font=("Segoe UI", 10), relief="flat", borderwidth=1,
                        padding=(5, 4))
        # 滚动条不在其列——全部自绘（_ZScrollbar），这里若再配一份
        # "Vertical.TScrollbar" 就是无人消费的死配置。

    def _title_font(self):
        """关于页标题字体。缓存：临时 Font 被回收会连带删掉 Tk 字体名。"""
        f = getattr(self, "_title_font_cache", None)
        if f is None:
            f = tkfont.nametofont("TkDefaultFont").copy()
            f.configure(weight="bold", size=17)
            self._title_font_cache = f
        return f

    def _btn_font(self, bold=False):
        """按钮字体。bold 给「添加文件 / 添加文件夹」这类次强按钮，其余按钮
        靠字重表达层级，不靠抢主色。缓存理由同 _title_font。
        """
        # 缓存属性名必须避开方法名自身：同名时 getattr 拿回来的是绑定方法，
        # 不是 Font 对象，调用处直接 AttributeError。
        attr = "_btn_font_bold_cache" if bold else "_btn_font_cache"
        f = getattr(self, attr, None)
        if f is None:
            f = tkfont.nametofont("TkDefaultFont").copy()
            if bold:
                f.configure(weight="bold")
            setattr(self, attr, f)
        return f

    def _note_font(self):
        """比默认小一号的说明文字字体，缓存理由同 _title_font。"""
        f = getattr(self, "_note_font_cache", None)
        if f is None:
            f = tkfont.nametofont("TkDefaultFont").copy()
            f.configure(size=max(7, f.cget("size") - 1))
            self._note_font_cache = f
        return f

    @staticmethod
    def _safe_grab(window):
        """grab_set 偶发失败时清场后重试一次。

        窗口刚建时系统 grab 没从上一次释放回过神，grab_set 会抛 "another
        application has grab"。模态丢了用户能点穿到主窗口，比不 grab 更危险
        ——先 release 清僵死记录再重试，仍失败就抛出。
        """
        try:
            window.grab_set()
        except tk.TclError:
            window.update()
            try:
                window.grab_release()
            except tk.TclError:
                pass
            window.grab_set()

    # 圆角按钮规格。tk/ttk 原生按钮都是直角矩形，圆角只能 Canvas 自绘。
    BTN_H = 28
    BTN_R = 6
    BTN_PADX = 14

    @staticmethod
    def _aa_round_image(w, h, r, fill, border, master=None):
        """抗锯齿圆角位图：4x 超采样绘制后缩小。

        Canvas 原生 arc 是硬边，高 DPI 下圆角一圈锯齿。PIL 画 4 倍再
        LANCZOS 缩小即得平滑圆角。
        边框与底色同色时视为无描边。

        master 必须显式传当前控件：不给 master 时 tkinter 会绑到"默认根窗口"，
        而默认根不一定是本控件所属的那个解释器（多窗口场景下图像会落到别的
        解释器里，画上去就报 image doesn't exist）。
        """
        s = 4
        img = Image.new("RGBA", (w * s, h * s), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if border and border != fill:
            d.rounded_rectangle([0, 0, w * s - 1, h * s - 1], radius=r * s,
                                fill=border)
            d.rounded_rectangle([s, s, w * s - 1 - s, h * s - 1 - s],
                                radius=max(0, (r - 1) * s), fill=fill)
        else:
            d.rounded_rectangle([0, 0, w * s - 1, h * s - 1], radius=r * s,
                                fill=fill)
        return ImageTk.PhotoImage(img.resize((w, h), Image.LANCZOS), master=master)

    def _check_icon(self, checked, disabled=False):
        """勾选框指示图标：未选 = 浅底描边圆角方块，选中 = 强调色底白对勾，
        禁用 = 灰底灰边空框（链式预览里行级勾选用不了，一眼要看出点不动）。

        clam 主题自带的选中态画的是「方块内打 ×」，看着像选错了。4 倍超采样
        后画白色对勾再缩小，与圆角按钮同一套绘制工艺。
        """
        s, size = 4, 16
        img = Image.new("RGBA", (size * s, size * s), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if disabled:
            d.rounded_rectangle([0, 0, size * s - 1, size * s - 1],
                                radius=4 * s, fill=self.BTN_OFF_BG,
                                outline=self.BTN_OFF_BD, width=s)
        elif checked:
            d.rounded_rectangle([0, 0, size * s - 1, size * s - 1],
                                radius=4 * s, fill=self.ACCENT)
            d.line([(int(size * s * 0.24), int(size * s * 0.52)),
                    (int(size * s * 0.43), int(size * s * 0.72)),
                    (int(size * s * 0.76), int(size * s * 0.30))],
                   fill=self.ACCENT_FG, width=2 * s, joint="curve")
        else:
            d.rounded_rectangle([0, 0, size * s - 1, size * s - 1],
                                radius=4 * s, fill=self.SURFACE,
                                outline=self.BORDER, width=s)
        # master 显式指定（不给会绑到"默认根窗口"，理由见 _aa_round_image）
        master = self.root if isinstance(getattr(self, "root", None), tk.Misc) else None
        return ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS),
                                  master=master)

    def _check_image(self, checked, disabled=False):
        """勾选图标（选中 / 未选 / 禁用），每套配色只画一次，并登记进保管箱防回收。"""
        cache = "_check_icons" if not disabled else "_check_icons_off"
        icons = getattr(self, cache, None)
        if icons is None:
            if disabled:
                # 禁用态只画一枚灰色空心框：链式预览里行级勾选用不了，
                # 画成"未勾选"会被当成可以点
                icons = (self._check_icon(False, disabled=True),)
            else:
                icons = (self._check_icon(True), self._check_icon(False))
            setattr(self, cache, icons)
            _CHECK_ICON_KEEPALIVE.extend(icons)
        return icons[0] if disabled else (icons[0] if checked else icons[1])

    def _make_check(self, parent, key, var, command=None):
        """自绘勾选框：Canvas 方框 + 独立文字，方框与文字都能点。

        ttk.Checkbutton 的文字虽然画在控件里，命中区却只在指示框附近稳定，
        用户习惯点文字，点完以为没反应。这里把方框与文字绑到同一个 toggle，
        整块（含中间 6px 间距）都可点。

        变量还可能被外部直接 set（加载规则、恢复界面状态、切主题），不跟
        变量同步重绘的话方框会停在旧状态。
        """
        holder = tk.Frame(parent, bg=parent.cget("bg"))
        cv = tk.Canvas(holder, width=16, height=16, bg=parent.cget("bg"),
                       highlightthickness=0, bd=0, cursor="hand2")
        cv.pack(side="left")
        lbl = tk.Label(holder, text=self._t(key), bg=parent.cget("bg"),
                       fg=self.TXT, cursor="hand2")
        lbl.pack(side="left", padx=(6, 0))
        self._reg_i18n(lbl, key)

        def repaint(*_a):
            try:
                cv.delete("all")
                cv._img = self._check_image(bool(var.get()))
                cv.create_image(0, 0, anchor="nw", image=cv._img)
            except tk.TclError:
                # 切主题重建界面时旧控件可能已被销毁，trace 回调仍会到
                pass

        def toggle(_e=None):
            var.set(not var.get())
            if command is not None:
                command()
            repaint()

        for w in (cv, lbl):
            w.bind("<Button-1>", toggle)
        var.trace_add("write", repaint)
        holder._br_repaint = repaint
        repaint()
        return holder

    def _round_button(self, parent, text, command, kind="outline", bold=False):
        """圆角按钮（Canvas 自绘）。

        kind=outline：底色 + 1px 圆角描边，hover 时描边加深、底色微亮；
        kind=accent：实心强调蓝，hover 变深。
        Canvas 上挂 _br_* 属性：状态、重绘闭包、i18n 键——禁用切换、语言
        刷新、主题切换都通过它们驱动，不依赖控件类型判断。
        """
        f = self._btn_font(bold)
        cv = tk.Canvas(parent, height=self.BTN_H, bg=parent.cget("bg"),
                       highlightthickness=0, bd=0, cursor="hand2")
        cv._br_kind = kind
        cv._br_text = text
        cv._br_cmd = command
        cv._br_enabled = True
        cv._br_hover = False
        cv._br_key = None

        def repaint():
            w = f.measure(cv._br_text) + self.BTN_PADX * 2
            cv.configure(width=w)
            cv.delete("all")
            if not cv._br_enabled:
                bg, fg, bd = self.BTN_OFF_BG, self.BTN_OFF_FG, self.BTN_OFF_BD
            elif kind == "accent":
                bg = self.ACCENT_HOVER if cv._br_hover else self.ACCENT
                fg, bd = self.ACCENT_FG, bg
            else:
                bg = self.BTN_HOVER_BG if cv._br_hover else self.BTN_BG
                fg = self.TXT
                bd = self.BTN_BD_HOVER if cv._br_hover else self.BTN_BD
            # PhotoImage 必须挂实例防 GC：被回收时画布上的图立刻消失
            cv._br_img = self._aa_round_image(w, self.BTN_H, self.BTN_R, bg, bd,
                                              master=cv)
            cv.create_image(w // 2, self.BTN_H // 2, image=cv._br_img)
            cv.create_text(w // 2, self.BTN_H // 2, text=cv._br_text,
                           fill=fg, font=f, tags="brtext")

        cv._br_repaint = repaint
        cv.bind("<Enter>", lambda _e: (setattr(cv, "_br_hover", True), repaint()))
        cv.bind("<Leave>", lambda _e: (setattr(cv, "_br_hover", False), repaint()))
        cv.bind("<Button-1>", lambda _e: cv._br_enabled and command())
        repaint()
        return cv

    def _make_outline_button(self, parent, text, command, bold=False):
        return self._round_button(parent, text, command, kind="outline", bold=bold)

    def _make_scrollbar(self, parent, command):
        """自绘滚动条（替代 ttk.Scrollbar）：水槽 / 滑块只认当前主题色板。

        ttk 滚动条的水槽在空列表 / 禁用态下不上主题色，会透出浅色底；
        自绘后颜色与列表有没有内容无关。
        """
        return _ZScrollbar(parent, command=command, width=10,
                           trough=self.SURFACE, thumb=self.BORDER)

    def _set_button_enabled(self, btn, enabled):
        """切换自绘按钮启用态：禁用时底色、描边、文字一起淡掉。"""
        btn._br_enabled = enabled
        btn._br_repaint()

    def _chip_font(self):
        """行尾小标签的字体：比默认字体小，且实际高度必须塞得进一行。

        高 DPI 下字体按缩放变大而行高是像素值不变，故按行高反推字号，避免
        标签比行还高、压到下一行。缓存理由同 _title_font。
        """
        f = getattr(self, "_chip_font_cache", None)
        if f is not None:
            return f
        base = tkfont.nametofont("TkDefaultFont")
        f = base.copy()
        limit = self.ROW_HEIGHT - 2
        for size in range(base.cget("size") - 1, 5, -1):
            f.configure(size=size)
            try:
                if f.metrics("linespace") <= limit:
                    break
            except Exception:
                break
        self._chip_font_cache = f
        return f

    def _make_accent_button(self, parent, text, command):
        return self._round_button(parent, text, command, kind="accent", bold=True)

    def _set_primary_state(self, state):
        """切换主操作按钮状态；禁用时整颗变灰，避免"看起来仍可按"。"""
        btn = getattr(self, "replace_btn", None)
        if btn is None:
            return
        self._set_button_enabled(btn, state != "disabled")

    def _set_file_ops_state(self, enabled):
        """搜索 / 替换运行中禁用一切"改动文件清单"的入口。

        运行中改清单虽不会崩（搜索取的是快照），但新加的文件不在本轮结果里，
        收尾后主列表会比预览窗多出一行，用户无从解释这个差额。菜单已删除，
        入口只剩工具栏三颗按钮，不再需要按钮与菜单两处一起切的同步逻辑。
        """
        for name in ("add_files_btn", "add_folder_btn", "remove_files_btn", "clear_files_btn"):
            btn = getattr(self, name, None)
            if btn is not None:
                self._set_button_enabled(btn, enabled)

    # ---------- 长任务进度 ----------

    def _show_progress(self, label, done, total, name):
        """状态栏进度文案。done 为 None 表示任务结束、清空。

        日志面板默认收起，进度必须写在常驻状态栏，否则大批量替换时用户以为卡死。
        任务期间收起「当前：…」模式方框给进度让位——它是静态提示，运行中
        没有信息量；收尾时按原 pack 信息放回。
        """
        bar = getattr(self, "progress_label", None)
        if bar is None:
            return
        if done is None:
            bar.config(text="")
            self._draw_progress_bar(0.0)
            self._set_mode_box_visible(True)
            return
        self._set_mode_box_visible(False)
        try:
            bar.config(text=self._fit_progress_text(label, done, total, name))
        except tk.TclError:
            pass
        # 文字说"到哪一步"，色条说"还剩多少"，两者同步刷新
        self._draw_progress_bar((done / total) if total else 0.0)

    def _set_mode_box_visible(self, visible):
        """收起 / 放回模式方框。放回时先撤进度标签再按原顺序装回，
        否则 pack 会把方框追加到进度标签后面、挤占它的扩展空间。"""
        box = getattr(self, "mode_box", None)
        if box is None:
            return
        shown = box.winfo_ismapped()
        if visible:
            if shown:
                return
            bar = getattr(self, "progress_label", None)
            if bar is not None:
                bar.pack_forget()
            box.pack(**self._mode_box_pack)
            if bar is not None:
                bar.pack(**self._progress_pack)
        elif shown:
            box.pack_forget()
            # 收起后立刻让布局重排一次：紧接着要按标签宽度裁剪进度文字，
            # 不刷新就还按旧宽度量，会把文字裁得比可用空间短一截
            try:
                box.update_idletasks()
            except tk.TclError:
                pass

    def _fit_progress_text(self, label, done, total, name):
        """进度文字按标签实际宽度裁剪：动作与序号永远保留，放不下的文件名
        做头尾保留省略——被砍掉的中间段通常是无信息量的路径与公司全名。
        """
        prefix = f"{label} {done}/{total}"
        if not name:
            return prefix
        full = f"{prefix} · {name}"
        try:
            f = tkfont.Font(font=self.progress_label.cget("font"))
            avail = self.progress_label.winfo_width() - 20  # padx 两侧各 10
            if avail <= 0 or f.measure(full) <= avail:
                return full
            room = avail - f.measure(prefix) - f.measure(" · ")
            if room <= f.measure("…"):
                return prefix
            n = len(name)
            while n > 1:
                head = (n + 1) // 2
                cand = name[:head] + "…" + (name[len(name) - (n - head):]
                                            if n > head else "")
                if f.measure(cand) <= room:
                    return f"{prefix} · {cand}"
                n -= 1
            return prefix
        except tk.TclError:
            return full

    def _set_search_progress(self, done, total, name):
        """搜索进度（工作线程经 UI 通道投递回主线程执行）。"""
        self._show_progress(self._t("progress_searching"), done, total, name)

    def _set_replace_progress(self, done, total, name):
        """替换进度（工作线程经 UI 通道投递回主线程执行）。"""
        self._show_progress(self._t("progress_replacing"), done, total, name)

    def _worker_alive(self, worker):
        """worker 线程是否还活着（线程退出时收尾投递必已进 UI 队列，
        不会"线程没了界面还锁着"）。"""
        is_alive = getattr(worker, "is_alive", None)
        return bool(is_alive is not None and is_alive())

    def _search_in_progress(self):
        """当前是否有搜索在跑，以 worker 线程是否存活为准。"""
        return self._worker_alive(getattr(self, "_search_thread", None))

    def _replace_in_progress(self):
        """替换批次是否正在进行；活着时禁止再发起搜索（两批次同时开同一批
        文件，os.replace 会撞权限错）。"""
        return self._worker_alive(getattr(self, "_replace_thread", None))

    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=self.BG)
        bar.pack(fill="x", side="top")

        def sep():
            tk.Frame(bar, bg=self.LINE, width=1).pack(side="left", fill="y", padx=6, pady=6)

        # 这四个按钮要存引用：搜索 / 替换运行中要整组禁用（见 _set_file_ops_state）
        self.add_files_btn = self._make_outline_button(
            bar, self._t("add_files"), self.add_files, bold=True)
        self.add_files_btn.pack(side="left", padx=(8, 2), pady=5)
        self._reg_i18n(self.add_files_btn, "add_files")
        self.add_folder_btn = self._make_outline_button(
            bar, self._t("add_folder"), self.add_folder, bold=True)
        self.add_folder_btn.pack(side="left", padx=2, pady=5)
        self._reg_i18n(self.add_folder_btn, "add_folder")
        self.remove_files_btn = self._make_outline_button(
            bar, self._t("remove"), self.remove_files)
        self.remove_files_btn.pack(side="left", padx=2, pady=5)
        self._reg_i18n(self.remove_files_btn, "remove")
        self.clear_files_btn = self._make_outline_button(
            bar, self._t("clear"), self.clear_files)
        self.clear_files_btn.pack(side="left", padx=2, pady=5)
        self._reg_i18n(self.clear_files_btn, "clear")
        sep()
        # 属性名沿用 replace_btn（_set_primary_state 按名查找）；它是「开始搜索」，替换在预览窗里点
        self.replace_btn = self._make_accent_button(
            bar, f"\u25b6  {self._t('start_search')}", self.start_search)
        self.replace_btn.pack(side="left", padx=2, pady=4)
        self._reg_i18n(self.replace_btn, "start_search", prefix="\u25b6  ")
        # 语言按钮与关于入口都排右上角：工具栏左侧只留文件操作与主操作，
        # 右侧放全局入口。关于是低频入口，用圆形「!」钮而非文字按钮。
        self.lang_btn = self._make_outline_button(
            bar, self._t("lang_to_en" if self.lang == "zh" else "lang_to_zh"),
            self.switch_lang)
        self.lang_btn.pack(side="right", padx=2, pady=5)
        self._make_theme_badge(bar)
        self._make_about_badge(bar)
        tk.Frame(self.root, bg=self.LINE, height=1).pack(fill="x", side="top")

    def _build_path_bar(self):
        bar = tk.Frame(self.root, bg=self.BG)
        bar.pack(fill="x", side="top")
        backup_check = self._make_check(bar, "enable_backup", self.backup_var,
                                        self.toggle_backup)
        backup_check.pack(side="left", padx=(10, 0), pady=5)
        backup_label = tk.Label(bar, text=self._t("backup_path_label"), bg=self.BG,
                                fg=self.TXT_DIM)
        backup_label.pack(side="left", padx=(14, 6), pady=5)
        self._reg_i18n(backup_label, "backup_path_label")
        self.backup_entry = ttk.Entry(bar, textvariable=self.backup_path)
        self.backup_entry.pack(side="left", fill="x", expand=True, pady=5)
        # 路径被清空 / 填回时状态栏要立刻改口（安全信号不能等到下一次操作才更新）
        self.backup_path.trace_add("write", lambda *_: self.update_status())
        browse_btn = self._make_outline_button(bar, self._t("browse_backup"),
                                               self.select_backup_path)
        browse_btn.pack(side="left", padx=4, pady=5)
        self._reg_i18n(browse_btn, "browse_backup")
        open_btn = self._make_outline_button(bar, self._t("open_backup_folder_btn"),
                                             self.open_backup_folder)
        open_btn.pack(side="left", padx=(0, 4), pady=5)
        self._reg_i18n(open_btn, "open_backup_folder_btn")
        rollback_btn = self._make_outline_button(bar, self._t("rollback_btn"),
                                                 self.show_rollback_dialog)
        rollback_btn.pack(side="left", padx=(0, 10), pady=5)
        self._reg_i18n(rollback_btn, "rollback_btn")
        tk.Frame(self.root, bg=self.LINE, height=1).pack(fill="x", side="top")

    def _build_param_bar(self):
        bar = tk.Frame(self.body, bg=self.BG)
        bar.grid(row=0, column=0, sticky="ew")
        # 多条件模式下这一行要跟着窗口长高（条件表占满富余高度），故存引用给
        # toggle_multi_condition 改 grid 参数与行权重；单条件模式仍是固定一行高。
        self.param_bar = bar

        # 单条件 / 多条件两套输入区占同一行，由「多条件替换」勾选框切换：常驻多行表格会撑高界面，也让人以为必须填多行
        self.single_frame = tk.Frame(bar, bg=self.BG)
        self.single_frame.pack(fill="x", padx=10, pady=(6, 3))
        row = self.single_frame
        # 这两处文案随「替换文件名 / 替换内容」切换，故存为属性；宽度与文案在 _refresh_find_replace_labels 里统一设置
        self.find_label = tk.Label(row, bg=self.BG, fg=self.TXT_DIM, anchor="w")
        self.find_label.pack(side="left", padx=(0, 6))
        self.find_entry = ttk.Entry(row)
        self.find_entry.pack(side="left", fill="x", expand=True, padx=(0, 14))
        self.replace_label = tk.Label(row, bg=self.BG, fg=self.TXT_DIM, anchor="w")
        self.replace_label.pack(side="left", padx=(0, 6))
        self.replace_entry = ttk.Entry(row)
        self.replace_entry.pack(side="left", fill="x", expand=True)

        # 多条件条件表：查找内容不可重复、替换为可重复。逐条用弹窗编辑——内联编辑要自己实现单元格编辑，出错面大得多
        # 顺序可拖拽调整：链式模式下从上到下就是执行次序
        self.multi_frame = tk.Frame(bar, bg=self.BG)
        caption = tk.Label(self.multi_frame, text=self._t("cond_table_caption"),
                           bg=self.BG, fg=self.TXT_DIM,
                           font=self._note_font())
        caption.pack(side="top", anchor="w", padx=10, pady=(2, 0))
        self._reg_i18n(caption, "cond_table_caption")
        list_wrap = tk.Frame(self.multi_frame, bg=self.BG)
        # fill="both"：list_wrap 只是表格的容器，不填满的话 pack 会把表格在
        # 富余高度里垂直居中，表格上下各留一片空白（表格越小越明显）
        list_wrap.pack(side="left", fill="both", expand=True, padx=(10, 6))
        self.cond_tree = ttk.Treeview(list_wrap, columns=("find", "replace"),
                                      show="headings", height=6, style=self.TREE_STYLE)
        # 手型光标 + 拖拽调序：抓住一行拖到目标行即实时交换位置。
        # 链式模式下顺序就是执行次序，必须能改；不设专用按钮，保持界面简洁。
        self.cond_tree.configure(cursor="hand2")
        self.cond_tree.bind("<ButtonPress-1>", self._cond_drag_press, add=True)
        self.cond_tree.bind("<B1-Motion>", self._cond_drag_motion, add=True)
        self.cond_tree.bind("<ButtonRelease-1>", self._cond_drag_release, add=True)
        self.cond_tree.heading("find", text=self._t("find_unique"))
        self.cond_tree.heading("replace", text=self._t("replace_with"))
        self._reg_heading(self.cond_tree, "find", "find_unique")
        self._reg_heading(self.cond_tree, "replace", "replace_with")
        self.cond_tree.column("find", width=320, anchor="w")
        self.cond_tree.column("replace", width=320, anchor="w")
        self.cond_tree.pack(side="left", fill="both", expand=True)
        cond_vsb = self._make_scrollbar(list_wrap, self.cond_tree.yview)
        self.cond_tree.configure(yscrollcommand=cond_vsb.set)
        cond_vsb.pack(side="right", fill="y")
        self.cond_tree.bind("<Double-Button-1>", lambda e: self.edit_condition())
        self.cond_tree.bind("<Button-3>", self._show_cond_menu)
        # 列分界线：ttk.Treeview 不带列边框，两列文字挨在一起时分不清归属。
        # 用 1px 的 Frame 盖在列边界上——列宽固定，线不必跟着滚动走。
        self.cond_divider = tk.Frame(self.cond_tree, bg=self.GRID)
        # after_idle：列宽是窗口布局时按 stretch 重算的，Configure 回调当场读到的
        # 还是旧值，线会停在上一处交界上
        self.cond_tree.bind("<Configure>", lambda _e: self.cond_tree.after_idle(
            self._place_cond_divider), add="+")
        self.cond_tree.bind("<Map>", self._place_cond_divider, add="+")

        cond_btns = tk.Frame(self.multi_frame, bg=self.BG)
        # anchor="n"：按钮组贴在表格顶端，否则会在表格高度里垂直居中
        cond_btns.pack(side="left", anchor="n", padx=(0, 10))
        for key, cmd in (("add_cond", self.add_condition), ("edit_cond", self.edit_condition),
                         ("delete_cond", self.delete_condition),
                         ("clear_cond", self.clear_conditions)):
            btn = self._make_outline_button(cond_btns, self._t(key), cmd)
            btn.pack(fill="x", pady=1)
            self._reg_i18n(btn, key)

        row2 = tk.Frame(bar, bg=self.BG)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        case_check = self._make_check(row2, "match_case", self.case_sensitive_var)
        case_check.pack(side="left", padx=(0, 16))
        fname_check = self._make_check(row2, "replace_filename_opt",
                                       self.replace_filename_var,
                                       self.toggle_replace_target)
        fname_check.pack(side="left", padx=(0, 16))
        multi_check = self._make_check(row2, "multi_condition",
                                       self.multi_condition_var,
                                       self.toggle_multi_condition)
        multi_check.pack(side="left", padx=(0, 4))
        # 并行 / 链式下拉框：勾选「多条件替换」才显现（toggle_multi_condition
        # 控制显隐），默认并行。说明不靠悬停——切换执行方式、或刚勾上多条件
        # 时直接摆到下拉框下方，用户点 × 才收（见 _on_multi_mode_selected）。
        self.multi_mode_combo = ttk.Combobox(
            row2, textvariable=self.multi_mode_var, state="readonly",
            width=self._multi_mode_width(),
            values=(self._t("multi_mode_parallel"), self._t("multi_mode_chain")))
        self.multi_mode_combo.pack(side="left", padx=(0, 16))
        # 监听变量而不是 <<ComboboxSelected>> 虚拟事件：语言切换时是程序改的
        # 值，虚拟事件不会触发；变量改动则两种来源都覆盖得到。
        self.multi_mode_var.trace_add("write", lambda *_: self._on_multi_mode_selected())
        self.multi_mode_combo.pack_forget()
        for key, cmd in (("history", self.show_history_dialog), ("load_rules", self.load_rule),
                         ("save_rules", self.save_rule)):
            btn = self._make_outline_button(row2, self._t(key), cmd)
            # 间距按文本实测：英文文案比中文宽，固定小间距会把三个按钮挤成
            # 一团（宽度是语言切换时重算的，行总宽超限时右侧按钮会被压扁）
            btn.pack(side="right", padx=(0, 6))
            self._reg_i18n(btn, key)
        # 参数栏与文件列表之间的分界线，占 body 的第 1 行（第 2 行是文件列表）
        tk.Frame(self.body, bg=self.LINE, height=1).grid(
            row=1, column=0, sticky="ew")

        # 两个标签的文案与宽度都靠它初始化（宽度要按当前语言的字体实测才不截断）
        self._refresh_find_replace_labels()
        self._update_mode_hint()

    def _text_width_units(self, keys):
        """按当前语言字体实测，把这几种文案里最宽的折算成 Tk 字符单位宽度。"""
        f = tkfont.nametofont("TkDefaultFont")
        avg = f.measure("0") or 1
        need = max(f.measure(self._t(k)) for k in keys)
        return -(-need // avg)

    def _fit_label_width(self, keys):
        """参数区标签宽度：取这几种文案里最宽的，换语言也不会被截断。"""
        return max(1, self._text_width_units(keys))

    # ========== 多条件替换：条件表编辑 ==========
    def toggle_multi_condition(self, announce=True):
        """「多条件替换」勾选 / 取消：切换单条件输入与条件表。

        多条件模式下参数栏改为随窗口伸展：条件表要能一次看到尽量多的行，
        否则窗口调高后富余高度全落到文件列表，条件表固定几行、上下两片空白。

        announce=False 用于主题重建后的状态恢复：那时不该凭空弹说明。
        """
        if self.multi_condition_var.get():
            self.single_frame.pack_forget()
            self.param_bar.grid_configure(sticky="nsew")
            self.body.rowconfigure(0, weight=1)
            self.multi_frame.pack(fill="both", expand=True, padx=10, pady=(6, 3))
            self.multi_mode_combo.pack(side="left", padx=(0, 16))
            if announce:
                # 刚勾上多条件：默认是并行，先把并行的说明摆在眼前
                self._show_mode_tip()
        else:
            self.multi_frame.pack_forget()
            self.param_bar.grid_configure(sticky="ew")
            self.body.rowconfigure(0, weight=0)
            self.single_frame.pack(fill="x", padx=10, pady=(6, 3))
            self.multi_mode_combo.pack_forget()
            self._hide_chain_tip()
        self._update_mode_hint()

    def _multi_mode_width(self):
        """下拉框宽度按当前语言最长的选项实测定，换语言不截断。"""
        return max(4, self._text_width_units(
            ("multi_mode_parallel", "multi_mode_chain")) + 3)

    def _multi_mode_from_combo(self):
        """下拉框显示串 → 内部值（parallel / chain）。串对不上就回并行：
        显示串只可能是两种之一，回错比抛错好——并行是默认语义。"""
        return "chain" if self.multi_mode_var.get() == self._t("multi_mode_chain") \
            else "parallel"

    def _on_multi_mode_selected(self):
        # 不能把方法命名为 _multi_mode：与 self._multi_mode 这个实例属性同名，
        # 属性一旦先落上去，方法就被遮蔽成字符串，调用即 TypeError（且 trace
        # 回调的异常会被 Tk 吞掉，界面看着正常、模式其实没切过去）
        previous = self._multi_mode
        self._multi_mode = self._multi_mode_from_combo()
        # 只有真换了执行方式才摆说明：换语言 / 主题恢复也会写这个变量，那时
        # 模式没变，不该凭空弹一个窗出来
        if self._multi_mode != previous:
            self._show_mode_tip()

    def _sync_multi_mode_texts(self):
        """语言切换后同步下拉框的选项与当前显示串（内部值不变）。"""
        if not hasattr(self, "multi_mode_combo"):
            return
        try:
            self.multi_mode_combo.config(values=(
                self._t("multi_mode_parallel"), self._t("multi_mode_chain")),
                width=self._multi_mode_width())
            self.multi_mode_var.set(
                self._t("multi_mode_chain" if self._multi_mode == "chain"
                        else "multi_mode_parallel"))
        except tk.TclError:
            pass

    def _show_mode_tip(self):
        """把当前执行方式的说明立即摊到下拉框下方，等用户点 × 收。

        不靠悬停、也不用展开列表：切换执行方式、或刚勾上多条件时自动弹出，
        那两个时刻正是最需要知道两者差别的时候。多条件没勾时下拉框不在场，
        没什么可说明的——程序改模式（如状态恢复）也不该凭空弹窗。
        """
        if not self.multi_condition_var.get():
            return
        self._hide_chain_tip()
        self._chain_tip_show()

    def _chain_tip_show(self):
        try:
            combo = self.multi_mode_combo
            if not combo.winfo_exists():
                return
            chain = self._multi_mode == "chain"
        except tk.TclError:
            # 界面重建期间旧下拉框可能已被销毁，此时不该再弹说明
            return
        mode = "chain" if chain else "parallel"
        tip = tk.Toplevel(self.root)
        tip.configure(bg=self.HINT_BG)
        tip.wm_overrideredirect(True)
        # 紧挨刚 pack 显形的下拉框定位：pack 只登记布局，几何要等 Tk 跑一轮
        # 事件才落地，此刻读 winfo_rootx/rooty/height 拿到的是旧值（甚至 0），
        # 悬浮窗会飞错位置（首次勾多条件必现，显示过一次后控件已布局好就正常）。
        # 先 update_idletasks 把待处理的布局算完，再读几何。
        self.root.update_idletasks()
        tip.wm_geometry("+%d+%d" % (
            combo.winfo_rootx(),
            combo.winfo_rooty() + combo.winfo_height() + 4))
        box = tk.Frame(tip, bg=self.HINT_BG, highlightthickness=1,
                       highlightbackground=self.LINE)
        box.pack()
        inner = tk.Frame(box, bg=self.HINT_BG)
        inner.pack(padx=10, pady=8)
        # 字体对象必须挂在窗体上保活：临时 Font 被回收会连带删掉 Tk 字体名，
        # 加粗标题会静默变回默认字重（自绘按钮踩过这个坑）。
        title_font = tkfont.nametofont("TkDefaultFont").copy()
        title_font.configure(weight="bold")
        mono_font = tkfont.nametofont("TkDefaultFont").copy()
        mono_font.configure(family="Consolas")

        # 标题行：左侧方式名（加粗），右侥手动关闭钮
        head = tk.Frame(inner, bg=self.HINT_BG)
        head.pack(fill="x")
        tk.Label(head, text=self._t("multi_mode_chain" if chain
                                    else "multi_mode_parallel"),
                 bg=self.HINT_BG, fg=self.INFO_FG, font=title_font,
                 anchor="w").pack(side="left")
        closer = tk.Label(head, text="\u00d7", bg=self.HINT_BG,
                          fg=self.HINT_CLOSE_FG, cursor="hand2", padx=6)
        closer._tip_closer = True      # 标记：按段收集正文时跳过它
        closer.pack(side="right")
        closer.bind("<Button-1>", lambda _e: self._hide_chain_tip())
        closer.bind("<Enter>", lambda _e: closer.config(fg=self.TXT))
        closer.bind("<Leave>", lambda _e: closer.config(fg=self.HINT_CLOSE_FG))

        def _lbl(text, font=None, padx=0, wraplength=0):
            return tk.Label(inner, text=text, bg=self.HINT_BG, fg=self.INFO_FG,
                            font=font, justify="left", anchor="w",
                            padx=padx, wraplength=wraplength)

        # 四段：标题（方式名，加粗）→ 说明（定宽换行）→ 示例（等宽整块缩进）
        # → 补充（结论 / 适用场景）。分段只为版面工整，措辞不变。
        _lbl(self._t("multi_mode_tip_%s" % mode),
             wraplength=320).pack(fill="x", pady=(4, 0))
        _lbl(self._t("multi_mode_tip_%s_example" % mode),
             font=mono_font, padx=14).pack(fill="x", pady=(6, 0))
        _lbl(self._t("multi_mode_tip_%s_note" % mode),
             wraplength=320).pack(fill="x", pady=(6, 0))
        tip._tip_fonts = (title_font, mono_font)
        self._chain_tip = tip
        # 默认存活 10s 自行收起：用户点 ×、切执行方式、关窗都会先走
        # _hide_chain_tip，它负责取消这里的定时器，不会留下野定时器。
        self._chain_tip_job = self.root.after(
            self.TIP_AUTO_CLOSE_MS, self._hide_chain_tip)

    def _hide_chain_tip(self, _e=None):
        if self._chain_tip_job is not None:
            try:
                self.root.after_cancel(self._chain_tip_job)
            except Exception:
                pass
            self._chain_tip_job = None
        if self._chain_tip is not None:
            try:
                self._chain_tip.destroy()
            except tk.TclError:
                pass
            self._chain_tip = None

    def _refresh_condition_tree(self):
        """按 self.conditions 重建条件表；行 iid 就是条件下标。"""
        if not hasattr(self, "cond_tree"):
            return
        for iid in self.cond_tree.get_children():
            self.cond_tree.delete(iid)
        for idx, (find_text, replace_text) in enumerate(self.conditions):
            self.cond_tree.insert("", "end", iid=str(idx),
                                  values=(f"{idx + 1}. {find_text}", replace_text))

    def _place_cond_divider(self, _event=None):
        """把分界线摆到「查找内容」与「替换为」的交界上（列宽固定，线不随滚动）。

        Treeview 没有列边框；画进 Canvas 会被表格重绘盖掉，盖一个 1px 的
        Frame 才是唯一能贴着列宽走的做法。
        """
        div = getattr(self, "cond_divider", None)
        tree = getattr(self, "cond_tree", None)
        if div is None or tree is None:
            return
        try:
            div.place(x=int(tree.column("find", "width")), y=0,
                      width=1, relheight=1)
            div.lift()
        except tk.TclError:
            pass

    # ---- 条件表拖拽调序 ----
    def _cond_drag_press(self, event):
        """按下：记住抓住了哪一行（点在行外空白处则不启动拖拽）。"""
        iid = self.cond_tree.identify_row(event.y)
        self._cond_drag = {"iid": iid, "moved": False} if iid else None

    def _cond_drag_motion(self, event):
        """拖动：实时与当前悬停行交换位置，用户看的就是最终顺序。

        不走 _apply_conditions_change：那会连带做去重校验，而拖动只改顺序、
        不改内容，重复项该保持拖拽前的原样，不能因为拖一下就被静默剔掉。
        """
        state = getattr(self, "_cond_drag", None)
        if not state:
            return
        src = state["iid"]
        if not src or not self.cond_tree.exists(src):
            return
        dst = self.cond_tree.identify_row(event.y)
        if not dst or dst == src:
            return
        src_idx, dst_idx = int(src), int(dst)
        conds = list(self.conditions)
        conds.insert(dst_idx, conds.pop(src_idx))
        self.conditions = conds
        self._refresh_condition_tree()
        state["iid"] = str(dst_idx)
        state["moved"] = True
        self.cond_tree.selection_set(str(dst_idx))

    def _cond_drag_release(self, _event):
        state = getattr(self, "_cond_drag", None)
        self._cond_drag = None
        if state and state["moved"]:
            # 链式模式下预览的轮次分组是按旧顺序生成的，必须重搜才一致
            self.log(self._t("lg_cond_moved"))

    def _condition_dialog(self, title_key, find_text, replace_text):
        """单条条件的编辑弹窗：确定返回 (查找, 替换为)，取消返回 None。"""
        dlg = tk.Toplevel(self.root)
        dlg.title(self._t(title_key))
        dlg.transient(self.root)
        dlg.resizable(False, False)
        self._safe_grab(dlg)
        dlg.configure(bg=self.BG)
        frm = tk.Frame(dlg, bg=self.BG)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(frm, text=self._t("find_content"), bg=self.BG,
                 fg=self.TXT_DIM).grid(row=0, column=0, sticky="w", pady=4)
        find_entry = ttk.Entry(frm, width=46)
        find_entry.grid(row=0, column=1, pady=4)
        if find_text:
            find_entry.insert(0, find_text)
        tk.Label(frm, text=self._t("replace_with"), bg=self.BG,
                 fg=self.TXT_DIM).grid(row=1, column=0, sticky="w", pady=4)
        repl_entry = ttk.Entry(frm, width=46)
        repl_entry.grid(row=1, column=1, pady=4)
        if replace_text:
            repl_entry.insert(0, replace_text)

        result = []

        def on_ok():
            result.append((find_entry.get(), repl_entry.get()))
            dlg.destroy()

        btns = tk.Frame(frm, bg=self.BG)
        btns.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        self._make_outline_button(btns, self._t("ok"), on_ok).pack(side="left", padx=4)
        self._make_outline_button(btns, self._t("cancel"),
                                  dlg.destroy).pack(side="left", padx=4)
        dlg.bind("<Return>", lambda e: on_ok())
        find_entry.focus_set()
        self._center_dialog(dlg)
        self._wait_dialog(dlg)
        return result[0] if result else None

    def _apply_conditions_change(self, trial):
        """校验候选条件列表，通过则落到 self.conditions，否则提示并保持原样。"""
        cleaned, errors = validate_conditions(trial, self.case_sensitive_var.get())
        if errors:
            messagebox.showwarning(self._t("t_warning"), "\n".join(errors))
            return False
        self.conditions = cleaned
        self._refresh_condition_tree()
        return True

    # ---- 条件表右键菜单：调序 / 插入 / 删除 ----
    def _show_cond_menu(self, event):
        """右键弹菜单：插入 / 上移 / 下移 / 移至顶底 / 删除。

        菜单每次按当前语言重建（临时弹出，无需挂 apply_lang 刷新）。点到某行就
        选中它，点空白处则保持原选中——空白右键也能插到末尾。
        """
        row = self.cond_tree.identify_row(event.y)
        if row:
            self.cond_tree.selection_set(row)
        sel = self.cond_tree.selection()
        has = bool(sel)
        idx = int(sel[0]) if has else -1
        n = len(self.conditions)
        m = tk.Menu(self.cond_tree, tearoff=0)
        m.add_command(label=self._t("mi_insert"), command=self._cond_insert)
        if has:
            top = idx == 0
            bottom = idx == n - 1
            m.add_separator()
            m.add_command(label=self._t("mi_move_top"),
                          command=lambda: self._cond_move("top"),
                          state="disabled" if top else "normal")
            m.add_command(label=self._t("mi_move_up"),
                          command=lambda: self._cond_move("up"),
                          state="disabled" if top else "normal")
            m.add_command(label=self._t("mi_move_down"),
                          command=lambda: self._cond_move("down"),
                          state="disabled" if bottom else "normal")
            m.add_command(label=self._t("mi_move_bottom"),
                          command=lambda: self._cond_move("bottom"),
                          state="disabled" if bottom else "normal")
            m.add_separator()
            m.add_command(label=self._t("mi_delete"), command=self.delete_condition)
        m.tk_popup(event.x_root, event.y_root)

    def _cond_move(self, mode):
        """把选中条件移到上/下/顶/底；越界即不动（菜单项此时已被禁用）。"""
        sel = self.cond_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        n = len(self.conditions)
        j = {"up": max(0, idx - 1), "down": min(n - 1, idx + 1),
             "top": 0, "bottom": n - 1}[mode]
        if j == idx:
            return
        item = self.conditions.pop(idx)
        self.conditions.insert(j, item)
        self._refresh_condition_tree()
        self.cond_tree.selection_set(str(j))
        self.log(self._t("lg_cond_moved"))

    def _cond_insert(self):
        """在选中行上方插入一条新条件；无选中则追加到末尾。"""
        sel = self.cond_tree.selection()
        at = int(sel[0]) if sel else len(self.conditions)
        pair = self._condition_dialog("t_add_cond", "", "")
        if pair is None:
            return
        find_text, replace_text = pair
        if not find_text:
            messagebox.showwarning(self._t("t_warning"), self._t("m_find_empty"))
            return
        trial = (self.conditions[:at] + [[find_text, replace_text]]
                 + self.conditions[at:])
        if self._apply_conditions_change(trial):
            self.cond_tree.selection_set(str(at))

    def add_condition(self):
        pair = self._condition_dialog("t_add_cond", "", "")
        if pair is None:
            return
        find_text, replace_text = pair
        if not find_text:
            messagebox.showwarning(self._t("t_warning"), self._t("m_find_empty"))
            return
        self._apply_conditions_change(self.conditions + [[find_text, replace_text]])

    def edit_condition(self):
        sel = self.cond_tree.selection()
        if not sel:
            messagebox.showinfo(self._t("t_notice"), self._t("m_select_to_edit"))
            return
        idx = int(sel[0])
        old_find, old_replace = self.conditions[idx]
        pair = self._condition_dialog("t_edit_cond", old_find, old_replace)
        if pair is None:
            return
        find_text, replace_text = pair
        if not find_text:
            messagebox.showwarning(self._t("t_warning"), self._t("m_find_empty"))
            return
        trial = (self.conditions[:idx] + [[find_text, replace_text]]
                 + self.conditions[idx + 1:])
        self._apply_conditions_change(trial)

    def delete_condition(self):
        sel = self.cond_tree.selection()
        if not sel:
            messagebox.showinfo(self._t("t_notice"), self._t("m_select_to_delete"))
            return
        del self.conditions[int(sel[0])]
        self._refresh_condition_tree()

    def clear_conditions(self):
        if not self.conditions:
            return
        if messagebox.askyesno(self._t("t_confirm_clear"),
                               self._t("m_confirm_clear_cond", n=len(self.conditions))):
            self.conditions = []
            self._refresh_condition_tree()

    def _build_file_table(self):
        wrap = tk.Frame(self.body, bg=self.BG)
        wrap.grid(row=2, column=0, sticky="nsew", padx=8, pady=(6, 0))

        cols = ("type", "size", "match", "status")
        self.files_tree = ttk.Treeview(wrap, columns=cols, show="tree headings",
                                       selectmode="extended", style=self.TREE_STYLE)
        # 表头文案与它的对齐方式、列宽放一起，新增一列时不会漏掉 i18n 注册
        for column, key, anchor in (("#0", "col_name", "w"), ("type", "col_type", "w"),
                                    ("size", "col_size", "e"), ("match", "col_match", "center"),
                                    ("status", "col_status", "w")):
            self.files_tree.heading(column, text=self._t(key), anchor=anchor)
            self._reg_heading(self.files_tree, column, key)
        self.files_tree.column("#0", width=330, minwidth=160, anchor="w", stretch=True)
        self.files_tree.column("type", width=80, minwidth=60, anchor="w", stretch=False)
        self.files_tree.column("size", width=90, minwidth=70, anchor="e", stretch=False)
        self.files_tree.column("match", width=70, minwidth=60, anchor="center", stretch=False)
        self.files_tree.column("status", width=120, minwidth=90, anchor="w", stretch=False)

        vsb = self._make_scrollbar(wrap, self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=vsb.set)
        self.files_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.files_tree.bind("<<TreeviewSelect>>", lambda e: self.update_status())
        # 斑马纹：Treeview 没有原生 zebra，用 tag 换底色实现；增删行后由
        # _restripe_file_tree 按当前行序重铺，与插入顺序无关。
        self.files_tree.tag_configure("odd", background=self.SURFACE)
        self.files_tree.tag_configure("even", background=self.ZEBRA)

    def _build_log_panel(self):
        self.log_panel = tk.Frame(self.root, bg=self.BG)
        self.log_panel.pack(fill="x", side="bottom")

        self.log_header = tk.Frame(self.log_panel, bg=self.BG, cursor="hand2")
        self.log_header.pack(fill="x", side="top")
        tk.Frame(self.log_header, bg=self.LINE, height=1).pack(fill="x", side="top")
        self.log_arrow = tk.Label(self.log_header, text=f"\u25b8  {self._t('process_log')}",
                                  bg=self.BG, fg=self.TXT_DIM, anchor="w", cursor="hand2")
        self.log_arrow.pack(side="left", padx=10, pady=3)
        for w in (self.log_header, self.log_arrow):
            w.bind("<Button-1>", lambda e: self.toggle_log())

        self.log_body = tk.Frame(self.log_panel, bg=self.BG)
        self.log_text = scrolledtext.ScrolledText(self.log_body, height=8, state="disabled",
                                                  relief="flat", borderwidth=0,
                                                  bg=self.SURFACE, fg=self.TXT)
        self.log_text.pack(fill="both", expand=True, padx=8, pady=(4, 6))

    def toggle_log(self):
        """展开 / 收起日志"""
        self._set_log_visible(not self.log_visible)
        # 记住用户的收起意图：收起后新日志不再自动顶开面板
        self.log_user_collapsed = not self.log_visible

    def _set_log_visible(self, visible):
        if visible == self.log_visible:
            return
        if visible:
            self.log_body.pack(fill="x", side="top")
        else:
            self.log_body.pack_forget()
        self.log_visible = visible
        self._refresh_log_arrow()

    def _status_field(self, parent, text="", fg=None):
        """状态栏方框：1px 描边 + 内边距，文字标签装在里面。

        描边走 highlight 而不是 relief=sunken：sunken 的明暗边由系统配色定死，
        深色主题下会亮出一圈白边；highlight 用的是本主题的 BORDER 色。
        """
        box = tk.Frame(parent, bg=self.SURFACE, bd=0,
                       highlightthickness=1, highlightbackground=self.BORDER)
        label = tk.Label(box, text=text, bg=self.SURFACE, fg=fg or self.TXT_DIM,
                         anchor="w", padx=6, pady=2)
        label.pack(side="left")
        return box, label

    def _build_status_bar(self):
        bar = tk.Frame(self.root, bg=self.BG)
        bar.pack(fill="x", side="bottom")
        tk.Frame(bar, bg=self.LINE, height=1).pack(fill="x", side="top")
        inner = tk.Frame(bar, bg=self.BG)
        inner.pack(fill="x", padx=4, pady=4)
        # 汇总句与模式提示各占一个方框：框线把「共 N 个文件 · …」和「当前：…」
        # 分成两块，段内的「·」只负责分项，不再兼任分组。
        # 初值只是占位：setup_ui 结尾会调 update_status 写成汇总句
        status_box, self.status_label = self._status_field(inner, self._t("ready"))
        status_box.pack(side="left")
        # 当前模式的常驻提示：接在汇总句右边（「备份开启」之后）。参数区第二行
        # 被「区分大小写」等选项占满，模式说明留在那里只会被挤掉。
        mode_box, self.mode_label = self._status_field(inner, fg=self.ACCENT)
        self.mode_box = mode_box
        mode_box.pack(side="left", padx=(6, 0))
        self._mode_box_pack = mode_box.pack_info()
        # 进度条先占右侧，再让文本进度标签吃剩余空间——反过来文本标签会把
        # 右侧空间全部吃掉，进度条被挤成 0 宽而看不见。
        self.progress_canvas = tk.Canvas(inner, width=self.BAR_W, height=self.BAR_H,
                                         bg=self.BG, highlightthickness=0, bd=0)
        self.progress_canvas.pack(side="right", padx=(0, 6))
        self._draw_progress_bar(0.0)
        # 进度标签常驻状态栏，搜索 / 替换期间显示「正在替换 3/10 · 文件名」；
        # update_status 只改写左侧汇总，不碰它。左对齐：文字超宽时从右端裁，
        # 居中会把首尾各切一刀，「正在替换」的动作和文件扩展名一起丢。
        self._progress_pack = dict(side="left", padx=(10, 10), fill="x", expand=True)
        self.progress_label = tk.Label(inner, text="", bg=self.BG, fg=self.ACCENT,
                                       anchor="w")
        self.progress_label.pack(**self._progress_pack)
        # 状态栏比参数栏晚建，模式提示的首次文案在这里补上
        self._update_mode_hint()

    def _draw_progress_bar(self, ratio):
        """重绘右下角进度条。ratio 为 0.0~1.0，空槽即未在跑任务。"""
        cv = getattr(self, "progress_canvas", None)
        if cv is None:
            return
        pad = 1
        w = self.BAR_W - pad * 2
        h = self.BAR_H - pad * 2
        try:
            cv.delete("all")
            cv.create_rectangle(pad, pad, pad + w, pad + h,
                                outline=self.BAR_EDGE, fill=self.BAR_TROUGH)
            ratio = 0.0 if ratio is None else max(0.0, min(1.0, ratio))
            if ratio > 0:
                cv.create_rectangle(pad + 1, pad + 1,
                                    pad + 1 + int((w - 2) * ratio), pad + h - 1,
                                    outline="", fill=self.BAR_FILL)
        except tk.TclError:
            pass

    def update_status(self):
        """刷新底部状态栏"""
        if not hasattr(self, "status_label"):
            return
        n = len(self.selected_files)
        try:
            sel = len(self.files_tree.selection())
        except (tk.TclError, AttributeError):
            sel = 0
        # 备份状态是安全信号，必须说实情：开了但没路径时替换会被拦下，
        # 状态栏就不能还挂着「备份开启」——那正是"界面说谎"的来源。
        if self.backup_var.get() and not self.backup_path.get().strip():
            backup, fg = self._t("backup_no_path"), self.DANGER_FG
        else:
            backup = self._t("backup_on" if self.backup_var.get() else "backup_off")
            fg = self.TXT
        self.status_label.config(
            text=self._t("status_summary", n=n, sel=sel, m=self._total_matches, backup=backup),
            fg=fg)

    def _insert_file_row(self, path, index="end"):
        ext, size = file_meta(path)
        stripe = "even" if len(self.files_tree.get_children()) % 2 else "odd"
        self.files_tree.insert("", index, iid=path, text=os.path.basename(path),
                               values=(ext, size, "-", self._t("st_not_searched")),
                               tags=(stripe,))

    def _restripe_file_tree(self):
        """增删行后按当前行序重铺斑马纹。"""
        tree = getattr(self, "files_tree", None)
        if tree is None:
            return
        for i, iid in enumerate(tree.get_children()):
            try:
                tree.item(iid, tags=("even" if i % 2 else "odd",))
            except tk.TclError:
                pass

    def _recalc_total_matches(self):
        """按文件列表当前的「匹配」列重算总匹配数，让状态栏与列表一致。"""
        tree = getattr(self, "files_tree", None)
        if tree is None:
            return
        total = 0
        for iid in tree.get_children():
            try:
                total += int(tree.set(iid, "match"))
            except (ValueError, tk.TclError):
                pass
        self._total_matches = total
        self.update_status()

    def _rename_file_row(self, old_path, new_path):
        """文件名替换后，主列表整行换成新文件名（列表 iid 是完整路径，重命名后旧 iid 已失效）。"""
        tree = getattr(self, "files_tree", None)
        if tree is not None:
            try:
                if tree.exists(old_path):
                    index = tree.index(old_path)
                    was_selected = old_path in tree.selection()
                    tree.delete(old_path)
                    self._insert_file_row(new_path, index=index)
                    if was_selected:
                        tree.selection_add(new_path)
                else:
                    self._insert_file_row(new_path)
                tree.set(new_path, "match", "0")
                tree.set(new_path, "status", self._t("st_replaced"))
            except tk.TclError:
                pass
        # selected_files 里存的也是完整路径，不跟着换，后续操作就全指向旧文件
        selected = getattr(self, "selected_files", None)
        if selected and old_path in selected:
            selected[selected.index(old_path)] = new_path
        self._recalc_total_matches()

    def _apply_search_status(self, counts):
        """把每个文件的匹配数写回列表"""
        for iid in self.files_tree.get_children():
            c = counts.get(iid)
            if c is None:
                self.files_tree.set(iid, "match", "-")
                self.files_tree.set(iid, "status", self._t("st_not_searched"))
            elif c > 0:
                self.files_tree.set(iid, "match", str(c))
                self.files_tree.set(iid, "status", self._t("st_pending"))
            else:
                self.files_tree.set(iid, "match", "0")
                self.files_tree.set(iid, "status", self._t("st_no_match"))

    def log(self, message, quiet=False):
        """添加日志（限制2000行）。

        quiet=True 只记账不展开面板：静默检查失败这类"知道就好"的信息，
        不该每次启动都把日志顶开。用户手动收起过面板后，本轮也不再自动展开。
        """
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > 2000:
            self.log_text.delete(1.0, f"{lines - 2000}.0")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        if not quiet and not self.log_visible and not self.log_user_collapsed:
            self._set_log_visible(True)
        self.root.update_idletasks()

    def toggle_backup(self):
        if not self.backup_var.get():
            result = messagebox.showwarning(
                self._t("t_backup_off"),
                self._t("m_backup_off"),
                icon="warning",
                type="yesno"
            )
            if result == "no":
                self.backup_var.set(True)
        # 状态栏那行字（含"备份开启/关闭"）只在别处操作时才重写，这里不主动刷新
        # 就会出现"备份已关、状态栏还说开启"——备份状态是安全信号，显示错误比
        # 不显示更糟。用户点"否"回滚后刷新也是对的：那时确实仍是开启。
        self.update_status()

    def select_backup_path(self):
        path = filedialog.askdirectory(title=self._t("fd_select_backup"),
                                       initialdir=self.backup_path.get())
        if path:
            self.backup_path.set(path)

    def open_backup_folder(self):
        path = self.backup_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning(self._t("t_warning"), self._t("m_backup_invalid"))
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror(self._t("t_error"), self._t("m_cannot_open_folder", e=str(e)))

    def generate_rule_filename(self, find_text, replace_text):
        # 「空」占位与「多条件N项」都走 T：它们会出现在用户看得见的文件名里，
        # 英文环境下不该混中文（与备份文件夹名同源同键）。
        empty = self._t("name_empty")
        # 截断放宽到 40：多条件摘要两条就奔着 40 字去，20 会把信息切断
        find_part = find_text[:40] if find_text else empty
        replace_part = replace_text[:40] if replace_text else empty
        # 替换侧留空 = 调用方只给了「查找」侧（多条件的 conditions_summary），
        # 不能再拼一个「-> 空」的箭头
        base = f"{find_part} -> {replace_part}" if replace_text else find_part
        base = sanitize_filename(base)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base} {timestamp}.json"

    def save_rule(self):
        if self.multi_condition_var.get():
            conditions, errors = validate_conditions(
                self.conditions, self.case_sensitive_var.get())
            if errors:
                messagebox.showwarning(self._t("t_warning"),
                                       self._t("m_cond_issue_save", errors="\n".join(errors)))
                return
            if not conditions:
                messagebox.showwarning(self._t("t_warning"), self._t("m_add_cond_first"))
                return
            # 规则文件名用 conditions_rule_base：与界面摘要同口径（首条 + …
            # + 等N项），但去掉引号格式——引号会被文件名清洗换成下划线
            rule_name_find = conditions_rule_base(conditions)
            rule_name_replace = ""
            rule = {
                "multi": True,
                "conditions": conditions,
                "case_sensitive": self.case_sensitive_var.get(),
                "replace_filename": self.replace_filename_var.get(),
                "backup_enabled": self.backup_var.get(),
                "create_time": datetime.now().strftime(TS_FMT)
            }
        else:
            find_text = self.find_entry.get()
            replace_text = self.replace_entry.get()
            if not find_text:
                messagebox.showwarning(self._t("t_warning"), self._t("m_enter_find"))
                return
            rule_name_find, rule_name_replace = find_text, replace_text
            rule = {
                "multi": False,
                "find": find_text,
                "replace": replace_text,
                "case_sensitive": self.case_sensitive_var.get(),
                "replace_filename": self.replace_filename_var.get(),
                "backup_enabled": self.backup_var.get(),
                "create_time": datetime.now().strftime(TS_FMT)
            }
        filename = self.generate_rule_filename(rule_name_find, rule_name_replace)
        file_path = os.path.join(RULES_DIR, filename)
        if os.path.exists(file_path):
            if not messagebox.askyesno(self._t("t_file_exists"),
                                       self._t("m_rule_exists", filename=filename)):
                base = os.path.splitext(filename)[0]
                filename = f'{base} {datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
                file_path = os.path.join(RULES_DIR, filename)
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(rule, f, ensure_ascii=False, indent=2)
            messagebox.showinfo(self._t("t_success"), self._t("m_rule_saved", filename=filename))
            self.log(self._t("lg_rule_saved", path=file_path))
        except Exception as e:
            messagebox.showerror(self._t("t_error"), self._t("m_save_rule_failed", e=str(e)))

    def load_rule(self):
        file_path = filedialog.askopenfilename(
            title=self._t("fd_load_rule"),
            initialdir=RULES_DIR,
            filetypes=[(self._t("ft_json"), "*.json"), (self._t("ft_all"), "*.*")]
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                rule = json.load(f)
            case_sensitive = rule.get("case_sensitive", False)
            # 多条件规则（带 conditions）与旧版单条件规则并存：旧规则照样能读，
            # 读进来就是单条件模式，不做格式转换、不悄悄丢字段。
            conditions = rule.get("conditions") or []
            if rule.get("multi") and conditions:
                cleaned, errors = validate_conditions(conditions, case_sensitive)
                if errors or not cleaned:
                    messagebox.showerror(
                        self._t("t_error"),
                        self._t("m_rule_invalid",
                                errors="\n".join(errors) if errors
                                else self._t("m_no_valid_condition")))
                    return
                self.conditions = cleaned
                self.multi_condition_var.set(True)
                self._refresh_condition_tree()
            else:
                self.multi_condition_var.set(False)
                self.find_entry.delete(0, tk.END)
                self.find_entry.insert(0, rule.get("find", ""))
                self.replace_entry.delete(0, tk.END)
                self.replace_entry.insert(0, rule.get("replace", ""))
            if hasattr(self, "single_frame"):
                # 程序按规则恢复输入区，不是用户刚勾上多条件：不弹执行方式说明，
                # 否则每换一次规则就盖住条件表 10 秒。
                self.toggle_multi_condition(announce=False)  # 输入区与勾选状态保持一致
            self.case_sensitive_var.set(case_sensitive)
            # save_rule 存了 replace_filename 就必须读回来，否则「文件名替换」
            # 规则加载后复选框状态丢失、规则不完整。
            self.replace_filename_var.set(rule.get("replace_filename", False))
            self.toggle_replace_target()  # 让参数栏文案与新状态一致
            self.backup_var.set(rule.get("backup_enabled", True))
            # 规则里的备份开关会改掉状态栏显示的内容，同样要主动刷新一次
            self.update_status()
            self.log(self._t("lg_rule_loaded", path=file_path))
        except Exception as e:
            messagebox.showerror(self._t("t_error"), self._t("m_load_rule_failed", e=str(e)))

    def add_files(self):
        if self.replace_filename_var.get():
            files = filedialog.askopenfilenames(
                title=self._t("fd_select_files"), filetypes=[(self._t("ft_all"), "*.*")])
        else:
            files = filedialog.askopenfilenames(
                title=self._t("fd_select_docs"),
                filetypes=[(self._t("ft_supported"), "*.docx;*.xlsx;*.txt"),
                           (self._t("ft_word"), "*.docx"),
                           (self._t("ft_excel"), "*.xlsx"),
                           (self._t("ft_text"), "*.txt"),
                           (self._t("ft_all"), "*.*")]
            )
        added = 0
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                self._insert_file_row(f)
                added += 1
        self.log(self._t("lg_files_added", n=added))
        self.update_status()

    def add_folder(self):
        """整文件夹导入：递归收下所有「合适的文件」，与 add_files 同一套去重与入列逻辑。

        合适 = 重命名模式下任何文件；内容替换模式下仅 SUPPORTED_EXTS。
        ~$ 开头的是 Office 打开文档时的锁文件，引擎读不了，直接跳过。
        """
        folder = filedialog.askdirectory(title=self._t("fd_select_folder"))
        if not folder:
            return
        rename_mode = self.replace_filename_var.get()
        added = 0
        for root, _dirs, names in os.walk(folder):
            for name in names:
                if name.startswith("~$"):
                    continue
                if not rename_mode and os.path.splitext(name)[1].lower() not in SUPPORTED_EXTS:
                    continue
                path = os.path.join(root, name)
                if path not in self.selected_files:
                    self.selected_files.append(path)
                    self._insert_file_row(path)
                    added += 1
        self.log(self._t("lg_files_added", n=added))
        self.update_status()

    def remove_files(self):
        sel = list(self.files_tree.selection())
        for iid in sel:
            if iid in self.selected_files:
                self.selected_files.remove(iid)
            self.files_tree.delete(iid)
        self.log(self._t("lg_files_removed", n=len(sel)))
        # 必须重算而不是只刷新：状态栏的「匹配 N 处」读的是 _total_matches，
        # 那是搜索当时存下的数字，删了行它不会自己变小。
        self._recalc_total_matches()
        self._restripe_file_tree()

    def clear_files(self):
        for iid in self.files_tree.get_children():
            self.files_tree.delete(iid)
        self.selected_files.clear()
        self.log(self._t("lg_files_cleared"))
        self._recalc_total_matches()

    # ========== 查找历史管理 ==========
    def load_history(self):
        """加载查找历史（读进来即统一成"一组条件"的结构，见 normalize_history_item）。"""
        self.find_history = []
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                self.find_history = [r for r in (normalize_history_item(it) for it in raw)
                                   if r is not None]
            except Exception as e:
                # 同 load_config：历史坏了不拦启动，但要让用户知道记录为什么没了。
                self.startup_notes.append(("note_history_failed", {"e": e}))

    def save_history(self):
        """保存查找历史"""
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.find_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(T("note_history_save_failed", e=e))

    def add_to_history(self, conditions, case_sensitive, replace_filename):
        """把一组条件记进查找历史：按「整组条件 + 大小写 + 作用对象」去重 + 限制组数。"""
        conditions = [[str(f), str(r)] for f, r in conditions]
        if not conditions:
            return
        for item in self.find_history:
            if (item["conditions"] == conditions
                    and item.get("case_sensitive") == case_sensitive
                    and item.get("replace_filename") == replace_filename):
                item["timestamp"] = datetime.now().strftime(TS_FMT)
                self.find_history.remove(item)
                self.find_history.insert(0, item)
                self.save_history()
                return
        self.find_history.insert(0, {
            "conditions": conditions,
            "case_sensitive": case_sensitive,
            "replace_filename": replace_filename,
            "timestamp": datetime.now().strftime(TS_FMT),
        })
        if len(self.find_history) > MAX_HISTORY:
            self.find_history = self.find_history[:MAX_HISTORY]
        self.save_history()

    def show_history_dialog(self):
        """查找历史对话框：一组条件一行，展开看具体条件。"""
        history_window = tk.Toplevel(self.root)
        history_window.title(self._t("history_title", n=MAX_HISTORY))
        history_window.geometry("840x553")
        history_window.transient(self.root)
        self._safe_grab(history_window)
        history_window.configure(bg=self.BG)

        main_frame = tk.Frame(history_window, bg=self.BG)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        header = tk.Label(main_frame, text=self._t("history_header", n=MAX_HISTORY),
                 bg=self.BG, fg=self.TXT,
                 font=self._btn_font(bold=True))
        header.pack(anchor="w")
        # 悬停标题弹出简短说明——常驻两行太啰嗦；鼠标移开即消失。
        _tip = [None]
        _tip_job = [None]     # 自动关闭定时器
        _show_job = [None]    # 悬停延迟弹出定时器

        def _cancel(job):
            if job[0] is not None:
                try: self.root.after_cancel(job[0])
                except Exception: pass
                job[0] = None

        def _hide_tip(_e=None):
            # 延迟弹出任务必须一起撤：不撤的话鼠标已离开 0.4 秒后仍会冒一个窗，
            # 扫过表头 N 次就冒 N 个（看起来像不停弹窗）。
            _cancel(_show_job)
            _cancel(_tip_job)
            if _tip[0] is not None:
                try: _tip[0].destroy()
                except tk.TclError: pass
                _tip[0] = None
        def _show_tip():
            _show_job[0] = None
            if not history_window.winfo_exists():
                # 延迟期间窗口已关：再建 Toplevel 会挂在已销毁的父窗上
                return
            _hide_tip()
            tip = tk.Toplevel(history_window)
            tip.wm_overrideredirect(True)
            tip.wm_geometry("+%d+%d" % (header.winfo_rootx(),
                                        header.winfo_rooty() + header.winfo_height() + 4))
            tk.Label(tip, text=self._t("history_note"), bg=self.HINT_BG,
                     fg=self.INFO_FG, justify="left", anchor="w", padx=8, pady=5,
                     highlightthickness=1, highlightbackground=self.LINE).pack()
            tip.bind("<Leave>", lambda _e: _hide_tip())
            _tip[0] = tip
            # 默认存活 10s 自行收起：鼠标不再停留在标题上也不会一直挂着。
            _tip_job[0] = self.root.after(self.TIP_AUTO_CLOSE_MS, _hide_tip)
        def _schedule_show(_e=None):
            # 重复进入只保留一个待弹任务：旧句柄不撤，扫过表头 N 次就会在鼠标
            # 离开之后先后弹出 N 个窗。
            _cancel(_show_job)
            _show_job[0] = self.root.after(400, _show_tip)
        header.bind("<Enter>", _schedule_show)
        header.bind("<Leave>", _hide_tip)

        list_frame = tk.Frame(main_frame, bg=self.BG)
        list_frame.pack(fill="both", expand=True, pady=5)

        # 共用属性（大小写 / 作用对象 / 时间）只写组头：每条条件都重复一遍，
        # 列表会变成一片重复文字，看不出"这几条属于同一组"。
        tree = ttk.Treeview(list_frame, columns=("case", "mode", "time"),
                            show="tree headings", height=15, style=self.TREE_STYLE)
        tree.heading("#0", text=self._t("col_cond_group"))
        tree.heading("case", text=self._t("col_case"))
        tree.heading("mode", text=self._t("col_target"))
        tree.heading("time", text=self._t("col_time"))
        tree.column("#0", width=420, anchor="w")
        tree.column("case", width=90, anchor="center")
        tree.column("mode", width=90, anchor="center")
        tree.column("time", width=160, anchor="center")

        # 字体对象必须挂在 self 上：临时 Font 被回收会连带删掉 Tk 字体名，而
        # tag 里存的只是名字，字体一没就静默退回默认字重。
        self._history_group_font = tkfont.nametofont("TkDefaultFont").copy()
        self._history_group_font.configure(weight="bold")
        tree.tag_configure("group", font=self._history_group_font)

        scrollbar = self._make_scrollbar(list_frame, tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # iid 绝不带列表下标：删一条后 find_history 变短而剩余行 iid 不重排，再点
        # 最后一行就按旧下标取记录，直接 IndexError（删除后整棵树重建）。
        node_map = {}   # iid -> ("group", item) 或 ("cond", item, 条件下标)

        def rebuild_tree():
            for iid in tree.get_children():
                tree.delete(iid)
            node_map.clear()
            for key, item in zip(history_row_keys(self.find_history), self.find_history):
                n = len(item["conditions"])
                tree.insert("", "end", iid=key, tags=("group",), open=False,
                            text=self._t("history_group", n=n),
                            values=(self._t("yes") if item.get("case_sensitive")
                                    else self._t("no"),
                                    self._t("target_filename")
                                    if item.get("replace_filename") else self._t("target_content"),
                                    item["timestamp"]))
                node_map[key] = ("group", item)
                for i, (find_text, replace_text) in enumerate(item["conditions"]):
                    cid = f"{key}::{i}"
                    tree.insert(key, "end", iid=cid, values=("", "", ""),
                                text=f"    {i + 1}. {find_text} \u2192 {replace_text}")
                    node_map[cid] = ("cond", item, i)

        rebuild_tree()

        btn_frame = tk.Frame(main_frame, bg=self.BG)
        btn_frame.pack(fill="x", pady=10)

        def on_select():
            """载入选中项：组头载入整组，条件行只载入那一条。"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(self._t("t_warning"), self._t("m_select_cond_group"))
                return
            node = node_map.get(selection[0])
            if node is None:
                messagebox.showwarning(self._t("t_warning"), self._t("m_record_gone"))
                return
            conds = node[1]["conditions"]
            if node[0] == "cond":
                conds = [conds[node[2]]]
            # 组里几条就摆成几条：单条回输入框、多条进条件表，否则载入后
            # 用户看不到自己刚选的东西。
            if len(conds) == 1:
                if self.multi_condition_var.get():
                    self.multi_condition_var.set(False)
                self.find_entry.delete(0, tk.END)
                self.find_entry.insert(0, conds[0][0])
                self.replace_entry.delete(0, tk.END)
                self.replace_entry.insert(0, conds[0][1])
            else:
                self.multi_condition_var.set(True)
                self.conditions = [list(c) for c in conds]
                self._refresh_condition_tree()
            if hasattr(self, "single_frame"):
                # 同上：从历史里挑一条回来不算「刚勾上多条件」，不弹说明
                self.toggle_multi_condition(announce=False)   # 输入区与勾选状态保持一致
            self.case_sensitive_var.set(bool(node[1].get("case_sensitive")))
            # 作用对象一并还原：只还原条件不还原它，加载出来的规则是错的。
            self.replace_filename_var.set(bool(node[1].get("replace_filename")))
            self.toggle_replace_target()
            history_window.destroy()
            self.log(self._t("lg_history_loaded", n=len(conds)))

        def on_delete():
            """删除选中项：组头删整组，条件行只删那一条（组空了连组一起删）。"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(self._t("t_warning"), self._t("m_select_to_del_record"))
                return
            node = node_map.get(selection[0])
            if node is None:
                return
            if node[0] == "group":
                what = self._t("del_what_group", n=len(node[1]["conditions"]))
            else:
                f, r = node[1]["conditions"][node[2]]
                what = self._t("del_what_cond", f=f, r=r)
            if not messagebox.askyesno(self._t("t_confirm_delete"),
                                       self._t("m_confirm_delete", what=what)):
                return
            item = node[1]
            if node[0] == "cond":
                item["conditions"].pop(node[2])
            # 按身份移除（同秒两条内容不同，按下标删会错位）；条件被删空的组不留。
            if node[0] == "group" or not item["conditions"]:
                for i, it in enumerate(self.find_history):
                    if it is item:
                        del self.find_history[i]
                        break
            self.save_history()
            rebuild_tree()
            self.log(self._t("lg_history_deleted"))

        def on_clear_all():
            """清空所有历史"""
            if not self.find_history:
                messagebox.showinfo(self._t("t_notice"), self._t("m_history_empty"))
                return
            if messagebox.askyesno(self._t("t_confirm_clear"), self._t("m_clear_all_history")):
                self.find_history = []
                self.save_history()
                rebuild_tree()
                self.log(self._t("lg_history_cleared"))

        for key, cmd in (("select", on_select), ("delete", on_delete),
                         ("clear_all", on_clear_all)):
            self._make_outline_button(btn_frame, self._t(key), cmd).pack(side="left", padx=5)
        self._make_outline_button(btn_frame, self._t("close"),
                                  history_window.destroy).pack(side="right", padx=5)

        tree.bind("<Double-Button-1>", lambda e: on_select())
        self._center_dialog(history_window)

    # ========== 回溯：恢复到替换前 ==========
    def _confirm_dialog(self, title, message, ok_key="ok", cancel_key="cancel"):
        """自建确认框：默认焦点停在「取消」上。

        messagebox.askyesno 没法指定默认焦点，而回滚是不可逆操作——误触
        回车直接覆盖文件是最坏的结果，所以这一步必须自己摆按钮。
        返回 True 表示用户确认。
        """
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.configure(bg=self.BG)
        body = tk.Frame(dlg, bg=self.BG)
        body.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(body, text=message, bg=self.BG, fg=self.TXT, justify="left",
                 anchor="w", wraplength=520).pack(anchor="w")
        btns = tk.Frame(dlg, bg=self.BG)
        btns.pack(fill="x", padx=16, pady=(0, 12))
        result = {"ok": False}

        def finish(value):
            result["ok"] = value
            dlg.destroy()

        ok_btn = self._make_outline_button(btns, self._t(ok_key),
                                           lambda: finish(True))
        ok_btn.pack(side="right", padx=(6, 0))
        cancel_btn = self._make_outline_button(btns, self._t(cancel_key),
                                               lambda: finish(False))
        cancel_btn.pack(side="right")
        cancel_btn.focus_set()
        dlg.bind("<Escape>", lambda _e: finish(False))
        dlg.bind("<Return>", lambda _e: finish(False))
        dlg.protocol("WM_DELETE_WINDOW", lambda: finish(False))
        self._safe_grab(dlg)
        self._center_dialog(dlg)
        self._wait_dialog(dlg)
        return result["ok"]

    def show_rollback_dialog(self):
        """回溯对话框：列出全部内容替换记录，选中一条回滚它涉及的文件。"""
        base_dir = (self.backup_path.get().strip()
                    or str(DEFAULT_BACKUP_DIR))
        entries, problems = scan_backups(base_dir)
        for note in problems:
            self.log(self._t("lg_manifest_problem", n=1) + ": " + os.path.basename(note[0]))

        dlg = tk.Toplevel(self.root)
        dlg.title(self._t("t_rollback"))
        dlg.geometry("860x473")
        dlg.transient(self.root)
        self._safe_grab(dlg)
        dlg.configure(bg=self.BG)

        tk.Label(dlg, text=self._t("rb_header"), bg=self.BG, fg=self.TXT,
                 font=self._btn_font(bold=True)).pack(anchor="w", padx=12, pady=(10, 6))
        wrap = tk.Frame(dlg, bg=self.BG)
        wrap.pack(fill="both", expand=True, padx=12)
        tree = ttk.Treeview(wrap, columns=("time", "cond", "mode", "files"),
                            show="headings", style=self.TREE_STYLE)
        for col, key in (("time", "rb_col_time"), ("cond", "rb_col_cond"),
                         ("mode", "rb_col_mode"), ("files", "rb_col_files")):
            tree.heading(col, text=self._t(key))
            self._reg_heading(tree, col, key)
        tree.column("time", width=150, anchor="w")
        tree.column("cond", width=380, anchor="w")
        tree.column("mode", width=80, anchor="center")
        tree.column("files", width=70, anchor="center")
        tree.pack(side="left", fill="both", expand=True)
        vsb = self._make_scrollbar(wrap, tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        for idx, entry in enumerate(entries):
            summary = conditions_summary(entry["conditions"])
            mode = self._t("rb_mode_chain" if entry["chain"] else "rb_mode_parallel")
            tree.insert("", "end", iid=str(idx),
                        values=(entry["ts"], summary, mode, len(entry["files"])))

        def on_rollback():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo(self._t("t_notice"), self._t("m_rb_pick"),
                                    parent=dlg)
                return
            self._rollback_entry(entries[int(sel[0])], entries, dlg)

        btns = tk.Frame(dlg, bg=self.BG)
        btns.pack(fill="x", padx=12, pady=10)
        rb_btn = self._make_outline_button(btns, self._t("rollback_btn"), on_rollback)
        rb_btn.pack(side="left", padx=(0, 6))
        self._make_outline_button(btns, self._t("close"), dlg.destroy).pack(side="left")
        tree.bind("<Double-Button-1>", lambda _e: on_rollback())
        if not entries:
            self._set_button_enabled(rb_btn, False)
            tk.Label(dlg, text=self._t("rb_empty"), bg=self.BG, fg=self.TXT_DIM,
                     wraplength=820, justify="center").pack(fill="x", padx=12,
                                                           pady=(0, 10))
        self._center_dialog(dlg)

    def _rollback_entry(self, entry, entries, parent):
        """回滚一条记录：确认 → 勾选文件 → 执行。"""
        # 更晚操作警告：同一文件在其后也被改过，回滚会一并盖掉
        later_hits = []
        for item in entry["files"]:
            canon = canonical_path(item.get("orig_canon") or item.get("orig", ""))
            later_hits.extend(detect_later_operations(entries, entry, canon))
        msg = self._t("m_rb_confirm")
        if later_hits:
            names = sorted({os.path.basename(f.get("orig", ""))
                            for e in later_hits for f in e["files"]})
            msg += "\n\n" + self._t("rb_later_warning", files="\n".join(names))
        if not self._confirm_dialog(self._t("t_rollback"), msg):
            return

        selected = entry["files"]
        if len(selected) > 1:
            selected = self._pick_rollback_files(entry, parent)
            if not selected:
                return

        self.log(self._t("lg_rb_start", n=len(selected)))
        ok, bad = rollback_files(entry, selected)
        for path in ok:
            self.log(self._t("lg_rb_one_ok", path=path))
        for path, err in bad:
            self.log(self._t("lg_rb_one_fail", path=path, e=err))
        self.log(self._t("lg_rb_done", ok=len(ok), fail=len(bad)))
        messagebox.showinfo(self._t("t_done"),
                            self._t("m_rb_result", ok=len(ok), fail=len(bad)),
                            parent=parent)

    def _pick_rollback_files(self, entry, parent):
        """多文件时让用户勾选要回滚哪几个；全不选返回 None，选了返回子集。"""
        dlg = tk.Toplevel(parent)
        dlg.title(self._t("t_rb_pick_files"))
        dlg.geometry("640x393")
        dlg.transient(parent)
        self._safe_grab(dlg)
        dlg.configure(bg=self.BG)
        tk.Label(dlg, text=self._t("rb_pick_header"), bg=self.BG, fg=self.TXT,
                 font=self._btn_font(bold=True)).pack(anchor="w", padx=12, pady=(10, 6))
        wrap = tk.Frame(dlg, bg=self.BG)
        wrap.pack(fill="both", expand=True, padx=12)
        tree = ttk.Treeview(wrap, columns=("file", "path"), show="tree headings",
                            style=self.TREE_STYLE)
        tree.heading("#0", text="")
        tree.heading("file", text=self._t("rb_col_file"))
        self._reg_heading(tree, "file", "rb_col_file")
        tree.heading("path", text=self._t("rb_col_path"))
        self._reg_heading(tree, "path", "rb_col_path")
        tree.column("#0", width=40, stretch=False)
        tree.column("file", width=200, anchor="w")
        tree.column("path", width=360, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        vsb = self._make_scrollbar(wrap, tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        checked = {}
        for idx, item in enumerate(entry["files"]):
            iid = str(idx)
            checked[iid] = True
            tree.insert("", "end", iid=iid,
                        image=self._check_image(True),
                        values=(os.path.basename(item.get("orig", "")),
                                item.get("orig", "")))

        def on_click(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            checked[iid] = not checked[iid]
            tree.item(iid, image=self._check_image(checked[iid]))

        tree.bind("<Button-1>", on_click)
        result = {"files": None}

        def on_ok():
            # checked 的键是 Treeview iid（字符串），files 是列表——必须先转
            # int，否则 TypeError 会被 Tk 吞掉，表现是"点了确定什么都没发生"
            picked = [entry["files"][int(i)] for i, c in checked.items() if c]
            if not picked:
                messagebox.showinfo(self._t("t_notice"), self._t("m_rb_pick_none"),
                                    parent=dlg)
                return
            result["files"] = picked
            dlg.destroy()

        btns = tk.Frame(dlg, bg=self.BG)
        btns.pack(fill="x", padx=12, pady=10)
        self._make_outline_button(btns, self._t("ok"), on_ok).pack(side="left")
        self._make_outline_button(btns, self._t("cancel"),
                                  dlg.destroy).pack(side="left", padx=6)
        self._center_dialog(dlg, parent)
        self._wait_dialog(dlg)
        return result["files"]

    def toggle_replace_target(self):
        """「替换文件名 / 替换内容」切换时同步界面提示。"""
        self._refresh_find_replace_labels()
        self._update_mode_hint()

    def _update_mode_hint(self):
        """常驻提示：把「单条件 / 多条件」与「文件名 / 内容」两个开关合成一句。

        参数栏比状态栏先建，换主题重建界面时这里会被拿旧的已销毁标签调一次——
        TclError 直接放过，状态栏建好时会再写一遍。
        """
        if getattr(self, "mode_label", None) is None:
            return
        is_filename = self.replace_filename_var.get()
        multi = self.multi_condition_var.get()
        if is_filename:
            text = self._t("mode_replace_filename_multi" if multi else "mode_replace_filename")
        else:
            text = self._t("mode_replace_content_multi" if multi else "mode_replace_content")
        try:
            self.mode_label.config(text=text)
        except tk.TclError:
            pass

    # ========== 搜索匹配提取 ==========
    def start_search(self):
        # 重入守卫放在最前面：搜索跑在后台线程里、主按钮会禁用，但「开始搜索」
        # 还有别的触发路径（替换 worker 的收尾、将来的快捷键），重复触发会
        # 让两个 worker 竞争，后完成的那个把旧结果刷进预览窗、界面状态也被
        # 改回"就绪"。
        if self._search_in_progress():
            messagebox.showinfo(self._t("t_notice"), self._t("m_searching"))
            return

        # 替换进行中也禁止再发起搜索（理由见 _replace_in_progress）。
        if self._replace_in_progress():
            messagebox.showinfo(self._t("t_notice"), self._t("m_replacing"))
            return

        if not self.selected_files:
            messagebox.showwarning(self._t("t_warning"), self._t("m_select_files"))
            return

        # 多条件在这里再校验一次（兜底）：用户可能先加好条件、再把「匹配大小写」
        # 切到不敏感，从而出现 "A" 与 "a" 这种隐性重复查找内容。
        if self.multi_condition_var.get():
            conditions, errors = validate_conditions(
                self.conditions, self.case_sensitive_var.get())
            if errors:
                messagebox.showwarning(self._t("t_warning"),
                                       self._t("m_cond_issue_search",
                                               errors="\n".join(errors)))
                return
            if not conditions:
                messagebox.showwarning(self._t("t_warning"), self._t("m_add_cond_click"))
                return
        else:
            find_text = self.find_entry.get()
            if not find_text:
                messagebox.showwarning(self._t("t_warning"), self._t("m_enter_find"))
                return
            conditions = [[find_text, self.replace_entry.get()]]

        # 搜索跑在工作线程里：全部扫描若放在 UI 线程，选一批大文档就会「界面
        # 冻结」——日志已经打完"正在搜索…"，窗口却拖不动，用户以为程序死了。
        # 界面状态一律先在主线程读成普通值再传进去，工作线程不碰任何 tk 对象。
        is_filename = self.replace_filename_var.get()
        case_sensitive = self.case_sensitive_var.get()
        files = list(self.selected_files)
        # 链式只在多条件下有意义；单条件没有"上一条"可接
        self.is_chain_mode = (self.multi_condition_var.get()
                              and self._multi_mode == "chain")

        # 升搜索代号：_finish_search 只认当前代号的结果，上一轮姗姗来迟的
        # 收尾会被丢弃，不会覆盖新搜索的界面状态。
        self._search_gen += 1
        gen = self._search_gen

        self.log(self._t("lg_searching"))
        self.preview_data = []
        self._search_counts = {}
        self._total_matches = 0
        self._set_primary_state("disabled")
        self._set_file_ops_state(False)
        self.update_status()

        self._search_thread = threading.Thread(
            target=self._search_worker,
            args=(files, conditions, case_sensitive, is_filename, gen,
                  self.is_chain_mode))
        self._search_thread.daemon = True
        self._search_thread.start()

    def _search_worker(self, files, conditions, case_sensitive, is_filename, gen=None,
                       chain=False):
        """工作线程：只做纯计算与文件读取，绝不触碰 tkinter 对象。"""
        raw = []
        counts = {}
        total = 0
        errors = []
        # 依赖缺失单独记一份：收尾时要靠它把「搜不了」和「真的没匹配」分开，
        # 不能靠翻文案猜——文案是多语言的，猜法会随语言失效。
        dep_errors = []
        conflicts = []

        total_files = len(files)
        for done, file_path in enumerate(files, 1):
            # 搜索进度进状态栏：日志面板默认收起，大批量搜索时界面像卡死，
            # 进度写在常驻状态栏才能被看见。
            self._ui(self._set_search_progress, done, total_files, os.path.basename(file_path))
            # 先把文件名取出来：若在 try 内赋值却在 except 里使用，
            # os.path.basename 先抛异常时就会演成 NameError 掩盖真实错误。
            name = os.path.basename(file_path)
            try:
                if is_filename:
                    if chain:
                        rounds, new_name = run_filename_chain(
                            name, conditions, case_sensitive)
                        notes = []
                    else:
                        rounds = None
                        new_name, notes = apply_filename_conditions(
                            name, conditions, case_sensitive)
                    if name != new_name:
                        matches = [{"type": "filename",
                                    "location": self._t("target_filename"),
                                    # 新文件名独立成字段，不靠解析显示串
                                    "new_name": new_name,
                                    "context": f"{name} -> {new_name}"}]
                        for note in notes:
                            conflicts.append(f"{name}：{note}")
                        total += 1
                        counts[file_path] = 1
                        raw.append({"file_path": file_path, "file_name": name,
                                    "matches": matches, "conflicts": notes})
                elif chain:
                    # 链式：整条链在临时副本上跑一遍，把每轮的命中按轮取回。
                    # 每轮记录的位置/上下文都基于该轮的输入，预览所见即执行结果。
                    rounds, notes = run_replace_chain(
                        file_path, conditions, case_sensitive, dep_errors=dep_errors)
                    for _, err in notes:
                        errors.append(f"{name}: {err}")
                    matches = [m for r in rounds for m in r["matches"]]
                    hits = sum(match_hit_count(m) for m in matches)
                    total += hits
                    counts[file_path] = hits
                    # 跨轮的命中区间来自不同状态的文本，拿来做重叠判断会凭空
                    # 报冲突把整批锁死（链式本就不存在"同一位置被两条条件抢"）
                    if matches:
                        raw.append({"file_path": file_path, "file_name": name,
                                    "matches": matches, "rounds": rounds,
                                    "conflicts": []})
                else:
                    ext = Path(file_path).suffix.lower()
                    engine = _engine_for(ext)
                    matches = []
                    if engine is not None and not engine["available"]:
                        # 缺库要说出来，不能静默返回空匹配——那会和
                        # "真的没有匹配"混成同一个提示
                        msg = self._t(engine["missing_search"], name=name)
                        errors.append(msg)
                        dep_errors.append(msg)
                    elif engine is not None:
                        for cond_idx, (find_text, replace_text) in enumerate(conditions, 1):
                            # 提取函数返回 (matches, skipped)；skipped 是公式格 /
                            # 受保护 run 等「搜到也不替换」的说明
                            found, notes = engine["extract"](
                                file_path, find_text, case_sensitive)
                            for note in notes:
                                errors.append(f"{name}: {note}")
                            for m in found:
                                m["replace_text"] = replace_text
                                m["cond_idx"] = cond_idx
                            matches.extend(found)
                    # 冲突清单挂在文件记录上：预览窗靠它判断"这一批还能不能执行"，
                    # 以及取消勾选冲突行后能否放行。
                    file_conflicts = [
                        format_conflict_note(hit.get("location", "?"), hit, rival)
                        for hit, rival in detect_condition_conflicts(matches)]
                    for note in file_conflicts:
                        conflicts.append(f"{name} {note}")
                    hits = sum(match_hit_count(m) for m in matches)
                    total += hits
                    counts[file_path] = hits
                    if matches:
                        raw.append({"file_path": file_path, "file_name": name,
                                    "matches": matches, "conflicts": file_conflicts})
            except Exception as e:
                errors.append(self._t("e_search_failed", name=name, e=str(e)))

        self._ui(self._set_search_progress, None, total_files, "")
        self._ui(self._finish_search, raw, counts, total, errors,
                 conditions, is_filename, gen, conflicts, dep_errors)

    def _finish_search(self, raw, counts, total, errors, conditions, is_filename,
                       gen=None, conflicts=None, dep_errors=None):
        """回到主线程：建 tk 变量、刷列表、开预览窗。"""
        # 过期结果直接丢弃：一轮搜索收尾前若已发起新一轮（代号已变），旧结果
        # 此时刷进去会把新搜索的"搜索中"界面改回就绪、预览窗也显示旧数据。
        if gen is not None and gen != getattr(self, "_search_gen", 0):
            self.log(self._t("lg_search_stale"))
            return
        # try/finally 包全程：刷日志 / 刷列表途中一旦抛异常（如 files_tree 尚未
        # 就绪），按钮会永久卡在禁用态，之后再也搜不了。
        try:
            for msg in errors:
                self.log(msg)

            self._search_counts = counts
            self._total_matches = total
            self.preview_data = []
            for item in raw:
                item["checked"] = self._var(tk.BooleanVar, value=True)
                self.preview_data.append(item)

            # 回写每个文件的匹配结果到文件列表（7z 风格列表的「匹配 / 状态」列）
            self._apply_search_status(counts)
            self.update_status()
            self._set_primary_state("normal")
            self._set_file_ops_state(True)

            if total == 0:
                # 零匹配时若存在依赖缺失，说清楚是"搜不了"而不是"没找到"——
                # 这两件事的提示一模一样，用户无法分辨。判据用 worker 单独
                # 报上来的 dep_errors，不靠翻文案（文案会随语言变）。
                if dep_errors:
                    messagebox.showinfo(
                        self._t("t_notice"),
                        self._t("m_no_searchable", dep=dep_errors[0]))
                    self.log(self._t("lg_no_searchable_dep", dep=dep_errors[0]))
                else:
                    messagebox.showinfo(self._t("t_notice"), self._t("m_no_match"))
                    self.log(self._t("lg_no_match"))
                return

            # 多条件也进历史：历史已按组展示，一套条件就是一组。
            self.add_to_history(conditions, self.case_sensitive_var.get(),
                                self.replace_filename_var.get())

            self.log(self._t("lg_found_total", n=total))
            if conflicts:
                for note in conflicts:
                    self.log(self._t("lg_overlap", note=note))
            self.create_preview_window(conditions, total, is_filename)
            # 先出预览再弹提示：别在结果还没出来时就先用弹窗拦住用户。
            if conflicts:
                self._show_conflict_notice(conflicts)
        finally:
            # 安全网——上面任何一步抛异常，界面也必须恢复可操作。
            # （正常路径上面已经恢复过一次，这里是幂等重放。）
            self._set_primary_state("normal")
            self._set_file_ops_state(True)

    def _show_conflict_notice(self, conflicts):
        """重叠命中提示：这不是"已经帮您选好了"，而是"整批停下了"。"""
        messagebox.showwarning(self._t("t_overlap_blocked"),
                               conflict_block_message(conflicts),
                               parent=self.root)

    # ========== 预览窗口（原生展开/折叠） ==========
    def _on_preview_destroyed(self, event):
        """预览窗关闭：清掉实例上指向窗内控件的引用。

        切语言重刷文案会踩到已销毁控件的死引用抛 TclError（现象：开始搜索过，
        再切换英文就会报错）；绑 <Destroy> 覆盖全部关闭路径。
        """
        if event.widget is not getattr(self, "preview_window", None):
            return          # <Destroy> 会从子控件一路冒泡上来，只认最外层那一次
        for name in ("preview_window", "preview_hint_bar", "preview_exec_btn",
                     "checked_count_label", "conflict_label", "preview_hint_label",
                     "preview_tree", "_ctx_overlays", "_ctx_text", "_ctx_pool"):
            setattr(self, name, None)

    def create_preview_window(self, conditions, total_match_count, is_filename):
        preview_window = tk.Toplevel(self.root)
        preview_window.title(self._t("preview_title"))
        preview_window.geometry("1040x733")
        preview_window.minsize(860, 553)
        preview_window.transient(self.root)
        self._safe_grab(preview_window)
        self._center_dialog(preview_window)
        preview_window.configure(bg=self.BG)
        preview_window.bind("<Escape>",
                            lambda _e: preview_window.destroy())
        # 关闭时把它留在实例上的控件引用清掉（见 _on_preview_destroyed）
        self.preview_window = preview_window
        preview_window.bind("<Destroy>", self._on_preview_destroyed, add="+")

        # 模式与悬停状态必须在建控件前就位（提示文案要按模式取值）
        self.preview_tree = None
        self.is_filename_mode = is_filename
        # 链式与否是本次搜索的属性，搜索开始时已定，预览窗重建不改它
        # 多条件时匹配行要标出"这条属于哪个条件"，故把条件表存下来
        self.active_conditions = conditions
        self._hover_iid = None
        self.preview_hint_bar = None
        # iid → (文件序号, 匹配序号或 None) 的唯一索引。
        # 不要用 "file_3_match_7" 这种字符串协议让各处自行 split 解析：任何
        # 一处格式变动都会全线出错，而且没法静态检查。
        self._iid_index = {}

        info = tk.Frame(preview_window, bg=self.BG)
        info.pack(fill="x", side="top")
        empty = self._t("empty_placeholder")
        if len(conditions) == 1:
            find_text, replace_text = conditions[0]
            tk.Label(info, text=self._t("find"), bg=self.BG,
                     fg=self.TXT_DIM).pack(side="left", padx=(10, 4), pady=5)
            tk.Label(info, text=find_text if find_text else empty, bg=self.BG,
                     fg=self.TXT).pack(side="left", pady=5)
            tk.Label(info, text="\u2192", bg=self.BG, fg=self.TXT_DIM).pack(side="left", padx=8, pady=5)
            tk.Label(info, text=replace_text if replace_text else empty, bg=self.BG,
                     fg=self.TXT).pack(side="left", pady=5)
        else:
            tk.Label(info, text=self._t("multi_cond_items", n=len(conditions)),
                     bg=self.BG, fg=self.TXT_DIM).pack(side="left", padx=(10, 8), pady=5)
            tk.Label(info, text=conditions_summary(conditions), bg=self.BG, fg=self.TXT,
                     anchor="w").pack(side="left", pady=5)
        tk.Label(info, text=self._t("preview_stats", n=total_match_count,
                                    m=len(self.preview_data)),
                 bg=self.BG, fg=self.TXT_DIM).pack(side="right", padx=10, pady=5)
        tk.Frame(preview_window, bg=self.LINE, height=1).pack(fill="x", side="top")

        bar = tk.Frame(preview_window, bg=self.BG)
        bar.pack(fill="x", side="top")
        for key, cmd in (("select_all", self.select_all_rows),
                         ("deselect_all", self.deselect_all_rows),
                         ("expand_all", self.expand_all_files),
                         ("collapse_all", self.collapse_all_files)):
            btn = self._make_outline_button(bar, self._t(key), cmd)
            btn.pack(side="left", padx=(8, 2) if key == "select_all" else 2, pady=5)
            self._reg_i18n(btn, key)
        cancel_btn = self._make_outline_button(bar, self._t("cancel"),
                                               preview_window.destroy)
        cancel_btn.pack(side="right", padx=(2, 8), pady=5)
        self._reg_i18n(cancel_btn, "cancel")
        # 留引用：有重叠命中时这颗按钮要按死（见 _refresh_conflict_gate）
        self.preview_exec_btn = self._make_accent_button(
            bar, f"\u25b6  {self._t('preview_exec')}",
            lambda: self.execute_replace_from_preview(preview_window, conditions)
        )
        self.preview_exec_btn.pack(side="right", padx=2, pady=4)
        tk.Frame(preview_window, bg=self.LINE, height=1).pack(fill="x", side="top")

        # 常驻操作提示条：放列表正上方、蓝底高对比、可点 × 关闭；只影响当前
        # 窗口、不持久化（理由见 load_config 的 preview_hint_var）。
        if self.preview_hint_var.get():
            self._build_preview_hint_bar(preview_window, is_filename)

        # 底部状态栏（先占底部，列表再填满中间）
        status = tk.Frame(preview_window, bg=self.BG)
        status.pack(fill="x", side="bottom")
        tk.Frame(status, bg=self.LINE, height=1).pack(fill="x", side="top")
        self.checked_count_label = tk.Label(status, text=self._t("checked_count", n=0),
                                            bg=self.BG, fg=self.TXT, anchor="w")
        self.checked_count_label.pack(side="left", padx=10, pady=3)
        # 阻断原因紧跟计数行，按钮灰了不用到处找为什么（文案见 _refresh_conflict_gate）。
        self.conflict_label = tk.Label(status, text="", bg=self.BG, fg=self.DANGER_FG, anchor="w")
        self.conflict_label.pack(side="left", padx=10, pady=3)
        # 这行文案随鼠标所在的行变化，直接把"双击会发生什么"说出来
        self.preview_hint_label = tk.Label(status, text=self._preview_hint_default(),
                                           bg=self.BG, fg=self.INFO_FG, anchor="e")
        self.preview_hint_label.pack(side="right", padx=10, pady=3)

        table_frame = tk.Frame(preview_window, bg=self.BG)
        table_frame.pack(fill="both", expand=True, padx=8, pady=6)

        if is_filename:
            columns = ("context",)
            tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", height=24,
                                selectmode="none", style=self.TREE_STYLE)
            tree.heading("#0", text=self._t("col_file"))
            tree.heading("context", text=self._t("col_filename_preview"))
            tree.column("#0", width=340, anchor="w")
            tree.column("context", width=600, anchor="w")
        else:
            columns = ("file", "count", "location", "context")
            tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", height=24,
                                selectmode="none", style=self.TREE_STYLE)
            # 勾选态是 #0 列的图标——ttk 的 Treeview 只支持「整行着色」，勾选列
            # 用文字符号就没法与工具栏那枚蓝底白勾的勾选框统一。#0 列原本空着，
            # 正好放图标，表头写「勾选」。
            tree.heading("#0", text=self._t("col_check"))
            tree.heading("file", text=self._t("col_filename"))
            tree.heading("count", text=self._t("col_matches"))
            tree.heading("location", text=self._t("col_location"))
            # 把"能跳转"写进列名，扫一眼表头就知道
            tree.heading("context", text=self._t("col_context"))
            tree.column("#0", width=40, anchor="center")
            tree.column("file", width=320, anchor="w")
            tree.column("count", width=70, anchor="center")
            tree.column("location", width=200, anchor="w")
            tree.column("context", width=380, anchor="w")

        scrollbar = self._make_scrollbar(table_frame, tree.yview)
        tree.configure(yscrollcommand=lambda f, l: (scrollbar.set(f, l),
                                                    self._ctx_schedule_sync()))
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.preview_tree = tree
        # 上下文列命中着色贴片（仅内容模式）：贴片是画在列表上层的画布，
        # 滚动 / 展开折叠 / 尺寸列宽变化都会让行几何失效，全部走重摆
        self._ctx_overlays = {}
        self._ctx_pool = []
        self._ctx_sync_job = None
        self._ctx_font = tkfont.nametofont("TkDefaultFont")

        # 悬停行的浅蓝底 + 行尾浮出的「↗ 双击跳转」标签：Treeview 无原生悬停，
        # 用 tag 换底色实现（只改 tag 不重绘，滚动不闪）；选中模式关掉，避免
        # 蓝色选中底与悬停高亮打架（本窗口行选择无功能在用，勾选走勾选列）。
        tree.tag_configure("hover", background=self.HOVER_BG)
        # 父行（文件行）浅灰底，与子行拉开层次
        tree.tag_configure("file", background=self.ZEBRA)
        # 链式：落在上一轮产出文本里的命中整行标色。Treeview 只支持整行 tag，
        # 单元格没法单独着色——所以位置列另加文字后缀，不靠颜色单独传达。
        tree.tag_configure("produced", background=self.PRODUCED_BG)
        tree.tag_configure("round", foreground=self.TXT_DIM)
        self._chip_iid = None
        self._chip_hide_job = None
        self.jump_chip = tk.Label(tree, text=self._t("chip_text"), bg=self.ACCENT,
                                  fg=self.ACCENT_FG,
                                  font=self._chip_font(), padx=6, pady=0,
                                  highlightthickness=0, bd=0, cursor="hand2")
        self.jump_chip.bind("<Enter>", lambda e: self._on_chip_enter())
        self.jump_chip.bind("<Leave>", lambda e: self._clear_hover())
        self.jump_chip.bind("<Button-1>", lambda e: self._on_chip_click())

        self._init_match_data()
        self._populate_preview_tree(tree)
        self.update_checked_count()

        tree.bind("<Button-1>", lambda e: self.on_tree_click(e, tree))
        tree.bind("<Double-Button-1>", lambda e: self.on_tree_double_click(e, tree))
        tree.bind("<<TreeviewOpen>>", lambda e: self._refresh_parent_hints(tree))
        tree.bind("<<TreeviewClose>>", lambda e: self._refresh_parent_hints(tree))
        # 鼠标划到哪一行，底部就说清那行双击会发生什么
        tree.bind("<Motion>", lambda e: self.on_preview_motion(e, tree))
        tree.bind("<Leave>", lambda e: self._on_tree_leave())
        # 滚动 / 改变窗口大小后行位置会变，先收起行尾标签
        tree.bind("<MouseWheel>", lambda e: self._clear_hover())
        tree.bind("<Configure>", lambda e: self._clear_hover())
        # 窗口关闭时撤掉待执行的收起任务，避免对已销毁控件操作
        preview_window.bind("<Destroy>", lambda e: self._cancel_hover_clear(), add="+")
        # 命中着色贴片的重摆时机：首次映射、滚动（yscrollcommand）、展开折叠、
        # 窗口 / 列宽变化、悬停换行（底色要跟着换）。全部经 after_idle 合并。
        preview_window.bind("<Map>", lambda e: self._ctx_schedule_sync(), add="+")
        tree.bind("<Configure>", lambda e: self._ctx_schedule_sync(), add="+")
        tree.bind("<<TreeviewOpen>>", lambda e: self._ctx_schedule_sync(), add="+")
        tree.bind("<<TreeviewClose>>", lambda e: self._ctx_schedule_sync(), add="+")
        tree.bind("<ButtonRelease-1>", lambda e: self._ctx_schedule_sync(), add="+")

    # ========== 预览窗口提示（发现性增强） ==========
    def _build_preview_hint_bar(self, parent, is_filename):
        """列表正上方的常驻提示条：与主界面执行方式说明同一套可关闭样式
        （同底色、同 1px 描边、同 × 悬停变深），不看按钮出处就知道怎么关。"""
        if self.is_chain_mode:
            hint_text = self._t("chain_hint")
        else:
            hint_text = self._t("hint_bar_filename" if is_filename
                                else "hint_bar_content")

        bar = tk.Frame(parent, bg=self.HINT_BG, highlightthickness=1,
                       highlightbackground=self.LINE)
        bar.pack(fill="x", side="top")
        tk.Label(bar, text=hint_text, bg=self.HINT_BG, fg=self.INFO_FG,
                 anchor="w").pack(side="left", padx=(10, 0), pady=6)

        close = tk.Label(bar, text="\u00d7", bg=self.HINT_BG, fg=self.HINT_CLOSE_FG,
                         cursor="hand2", padx=6, pady=6)
        close.pack(side="right")
        close.bind("<Button-1>", lambda e: self._dismiss_preview_hint())
        close.bind("<Enter>", lambda e: close.config(fg=self.TXT))
        close.bind("<Leave>", lambda e: close.config(fg=self.HINT_CLOSE_FG))

        self.preview_hint_bar = bar

    def _dismiss_preview_hint(self):
        """关闭提示条：只影响当前这个预览窗，下次搜索照常出现。"""
        self.preview_hint_var.set(False)
        if self.preview_hint_bar is not None:
            self.preview_hint_bar.destroy()
            self.preview_hint_bar = None

    def _preview_hint_default(self):
        """底部说明的默认文案（链式另说：勾选粒度与轮次规则都变了）。"""
        if self.is_chain_mode:
            return self._t("chain_hint")
        return self._t("hint_default_filename" if getattr(self, "is_filename_mode", False)
                       else "hint_default_content")

    def _reset_preview_hint_label(self):
        """只把底部文案恢复默认，不动当前悬停行。"""
        if getattr(self, "preview_hint_label", None) is not None:
            self.preview_hint_label.config(text=self._preview_hint_default())

    def on_preview_motion(self, event, tree):
        """鼠标落在某行上时：行底色高亮 + 行尾浮出「↗ 双击跳转」+ 底部说明动作。"""
        if getattr(self, "preview_hint_label", None) is None:
            return
        self._cancel_hover_clear()
        item_id = tree.identify_row(event.y)
        if item_id == self._hover_iid:
            return
        self._set_hover_row(tree, item_id)
        self._apply_hint_for(tree, item_id)

    def _apply_hint_for(self, tree, item_id):
        """按行类型更新底部文案与行尾标签。"""
        label = getattr(self, "preview_hint_label", None)
        if label is None:
            return
        pos = self._iid_pos(item_id)
        if pos is None:
            label.config(text=self._preview_hint_default())
            self._hide_jump_chip()
            return
        file_idx, match_idx = pos
        try:
            if match_idx is None:
                label.config(text=self._t("hint_expand_row"))
                self._hide_jump_chip()
            else:
                fd = self.preview_data[file_idx]
                loc = fd["match_data"][match_idx]["match"]["location"]
                if not self._chain_can_jump(item_id):
                    label.config(text=self._t(
                        "hint_no_jump_row", n=self._chain_round_of.get(item_id, 1)))
                    self._hide_jump_chip()
                    return
                label.config(text=self._t("hint_jump_row", file=fd["file_name"], loc=loc))
                self._show_jump_chip(tree, item_id)
        except (IndexError, KeyError, ValueError):
            label.config(text=self._preview_hint_default())
            self._hide_jump_chip()

    def _iid_pos(self, iid):
        """把 iid 解析成 (文件序号, 匹配序号或 None)。不是预览行则返回 None。

        取代散落各处的 split("_match_") / split("_") 反查，见 _iid_index。
        """
        if not iid:
            return None
        if not hasattr(self, "_iid_index"):
            return None
        return self._iid_index.get(iid)

    # ========== 悬停高亮与行尾「↗ 双击跳转」标签 ==========
    def _set_hover_row(self, tree, iid):
        """把「悬停」tag 从上一行挪到当前行（Treeview 无原生悬停，用 tag 换底色）。"""
        for old in (getattr(self, "_hover_iid", None), iid):
            if old:
                try:
                    tags = tuple(t for t in tree.item(old, "tags") if t != "hover")
                    tree.item(old, tags=tags + (("hover",) if old == iid else ()))
                except tk.TclError:
                    pass
        self._hover_iid = iid
        # 悬停行换底色，该行上下文贴片的底色也要跟着换
        self._ctx_schedule_sync()

    def _show_jump_chip(self, tree, iid):
        """把「↗ 双击跳转」标签贴到该行右端。"""
        chip = getattr(self, "jump_chip", None)
        if chip is None or tree is None or not iid:
            return
        bbox = tree.bbox(iid)
        if not bbox:
            self._hide_jump_chip()
            return
        self._chip_iid = iid
        _x, y, _w, h = bbox
        chip_w = chip.winfo_reqwidth()
        chip_h = chip.winfo_reqheight()
        chip.place(x=max(0, tree.winfo_width() - chip_w - 22),
                   y=y + max(0, (h - chip_h) // 2))
        chip.lift()

    def _hide_jump_chip(self):
        chip = getattr(self, "jump_chip", None)
        if chip is not None:
            try:
                chip.place_forget()
            except tk.TclError:
                pass
        self._chip_iid = None

    def _on_chip_enter(self):
        """鼠标停在行尾标签上：留住行高亮与标签，底部说明也别退回默认。"""
        self._cancel_hover_clear()
        self._hover_iid = self._chip_iid
        self._apply_hint_for(getattr(self, "preview_tree", None), self._chip_iid)

    def _on_chip_click(self):
        """点行尾标签 = 双击该行：跳转到文件对应位置。"""
        iid = getattr(self, "_chip_iid", None)
        self._clear_hover()
        if iid:
            self.open_and_jump_to_match(iid)

    def _on_tree_leave(self):
        """鼠标离开列表：先恢复底部文案，稍后再收高亮——
        留出把鼠标移到行尾标签上的时间（否则标签刚出现就消失，点不到）。"""
        self._reset_preview_hint_label()
        self._schedule_hover_clear()

    def _schedule_hover_clear(self):
        self._cancel_hover_clear()
        try:
            self._chip_hide_job = self.root.after(180, self._clear_hover)
        except tk.TclError:
            self._chip_hide_job = None

    def _cancel_hover_clear(self):
        job = getattr(self, "_chip_hide_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
            self._chip_hide_job = None

    def _clear_hover(self):
        """撤掉行高亮与行尾标签，底部回到默认文案。"""
        self._chip_hide_job = None
        tree = getattr(self, "preview_tree", None)
        try:
            if tree is not None:
                self._set_hover_row(tree, None)
            self._hide_jump_chip()
            if getattr(self, "preview_hint_label", None) is not None:
                self.preview_hint_label.config(text=self._preview_hint_default())
        except tk.TclError:
            pass  # 预览窗已关闭，控件随之销毁
        # 悬停行撤了底色，该行上下文贴片的底色也要跟着换
        self._ctx_schedule_sync()

    def _refresh_parent_hints(self, tree):
        """折叠的文件行显示「展开可见 N 处匹配」，展开后清空（仅内容模式）。"""
        if getattr(self, "is_filename_mode", False):
            return
        for idx, file_data in enumerate(self.preview_data):
            file_iid = _file_iid(idx)
            try:
                values = list(tree.item(file_iid, "values"))
            except tk.TclError:
                continue
            if len(values) < 4:
                continue
            if tree.item(file_iid, "open"):
                values[3] = ""
            else:
                values[3] = self._t("parent_hint", n=sum(
                    match_hit_count(m) for m in file_data.get("matches", [])))
            tree.item(file_iid, values=tuple(values))

    # ========== 上下文列命中着色（贴片） ==========
    # Treeview 的 tag 只能整行着色，单元格内部没法给命中段单独换色——上下文
    # 列的着色由贴在列表上层的画布接管：每条可见命中行一片画布，底色跟随该行
    # 当前状态（默认 / 链式产出 / 悬停），前文后文照常正文色，命中段用主题的
    # HIT_FG。深浅主题各有自己的色值，都按「与该主题全部行底色拉开对比」挑的。

    def _ctx_schedule_sync(self):
        """合并同一轮事件里的多次重摆请求，空闲时一次摆完。"""
        if getattr(self, "_ctx_sync_job", None) is not None:
            return
        try:
            self._ctx_sync_job = self.root.after_idle(self._ctx_sync)
        except tk.TclError:
            self._ctx_sync_job = None

    def _ctx_sync(self):
        """按当前可见的命中行重摆贴片：少了补、多了收进备用池复用。"""
        self._ctx_sync_job = None
        tree = getattr(self, "preview_tree", None)
        if tree is None or not tree.winfo_exists():
            return
        if getattr(self, "is_filename_mode", False):
            return
        if getattr(self, "_ctx_overlays", None) is None \
                or getattr(self, "_ctx_text", None) is None:
            return
        try:
            texts = self._ctx_text
            wanted = {}

            def walk(parent):
                for iid in tree.get_children(parent):
                    if iid in texts:
                        bbox = tree.bbox(iid, "context")
                        # 折叠或滚出视野的行 bbox 为空，天然跳过
                        if bbox:
                            wanted[iid] = bbox
                    elif tree.item(iid, "open"):
                        walk(iid)

            walk("")
            pool = self._ctx_pool
            for iid in list(self._ctx_overlays):
                if iid not in wanted:
                    cv = self._ctx_overlays.pop(iid)
                    cv.place_forget()
                    if len(pool) < 64:
                        pool.append(cv)
                    else:
                        cv.destroy()
            tree_w = tree.winfo_width()
            for iid, (x, y, w, h) in wanted.items():
                if x >= tree_w:
                    continue
                cv = self._ctx_overlays.get(iid)
                if cv is None:
                    cv = self._ctx_make_overlay(tree)
                    self._ctx_overlays[iid] = cv
                tags = tree.item(iid, "tags")
                if "hover" in tags:
                    bg = self.HOVER_BG
                elif "produced" in tags:
                    bg = self.PRODUCED_BG
                else:
                    bg = self.SURFACE
                self._ctx_paint(cv, iid, bg, w, h)
                # 列被窗口右缘裁掉时贴片跟着裁，不越出列表
                cv.place(x=x, y=y, width=w, height=min(h, max(1, tree_w - x)))
            chip = getattr(self, "jump_chip", None)
            if chip is not None:
                chip.lift()      # 行尾标签永远压在贴片之上
        except tk.TclError:
            pass  # 摆放途中窗口被关了，控件随窗口销毁

    def _ctx_make_overlay(self, tree):
        cv = tk.Canvas(tree, highlightthickness=0, bd=0)
        # 贴片盖在上下文列上，悬停 / 双击 / 滚轮按树坐标转给树的处理器——
        # 行为与贴片不存在时完全一致
        cv.bind("<Motion>", lambda e, c=cv: self._ctx_event(e, c, "motion"))
        cv.bind("<Button-1>", lambda e, c=cv: self._ctx_event(e, c, "click"))
        cv.bind("<Double-Button-1>", lambda e, c=cv: self._ctx_event(e, c, "double"))
        cv.bind("<MouseWheel>", lambda e, c=cv: self._ctx_event(e, c, "wheel"))
        cv.bind("<Leave>", lambda e: self._on_tree_leave())
        return cv

    def _ctx_event(self, event, cv, kind):
        tree = getattr(self, "preview_tree", None)
        if tree is None:
            return
        ev = type("E", (), {"x": cv.winfo_x() + event.x,
                            "y": cv.winfo_y() + event.y})
        if kind == "motion":
            self.on_preview_motion(ev, tree)
        elif kind == "click":
            self.on_tree_click(ev, tree)
        elif kind == "double":
            self.on_tree_double_click(ev, tree)
        else:
            try:
                tree.event_generate("<MouseWheel>", delta=event.delta)
            except tk.TclError:
                pass
        return "break"

    def _ctx_paint(self, cv, iid, bg, w, h):
        pre, hit, post = self._ctx_text.get(iid, ("", "", ""))
        cv.configure(bg=bg)
        cv.delete("all")
        f = self._ctx_font
        # 4 是树画单元格文字的左边距（逐像素实测，恒定值，不随缩放变）；
        # 文字垂直方向 anchor="w" 自带居中
        x = 4
        y = h // 2
        for text, fill in ((pre, self.TXT), (hit, self.HIT_FG), (post, self.TXT)):
            if not text:
                continue
            cv.create_text(x, y, text=text, font=f, fill=fill, anchor="w")
            x += f.measure(text)

    def _init_match_data(self):
        """为所有 preview_data 初始化 match_data"""
        for file_data in self.preview_data:
            if "match_data" not in file_data:
                file_data["match_data"] = [
                    {"match": m, "checked": self._var(tk.BooleanVar,
                                                      value=True)}
                    for m in file_data["matches"]
                ]

    def _populate_preview_tree(self, tree):
        """填充 Treeview（始终插入子行，让原生 +/- 指示器显示）"""
        self._iid_index = {}
        self._round_index = {}
        self._chain_round_of = {}
        # 命中着色贴片要画的（前文, 命中, 后文），iid → 三元组；贴片机制见
        # _ctx_sync。树单元格里不画 context（置空），整格由贴片接管。
        self._ctx_text = {}
        for idx, file_data in enumerate(self.preview_data):
            file_iid = _file_iid(idx)
            self._iid_index[file_iid] = (idx, None)
            check_img = self._check_image(file_data["checked"].get())

            if self.is_filename_mode:
                context = file_data["matches"][0]["context"] if file_data["matches"] else ""
                tree.insert("", "end", iid=file_iid, text=file_data["file_name"],
                            image=check_img, values=(context,), tags=("file",))
            else:
                match_count = len(file_data["matches"])
                tree.insert("", "end", iid=file_iid, text="", image=check_img,
                            values=(file_data["file_name"], str(match_count), "", ""),
                            tags=("file",))
                if self.is_chain_mode:
                    self._insert_chain_children(tree, file_iid, file_data, idx)
                else:
                    self._insert_match_children(tree, file_iid, file_data, idx)
                # 默认展开第一个文件，让"文件行下面还藏着明细"第一眼就看得见
                tree.item(file_iid, open=(idx == 0 and match_count > 0))
        self._refresh_parent_hints(tree)

    def _insert_chain_children(self, tree, file_iid, file_data, file_idx):
        """链式预览：先按轮次分组，再列每轮的命中。

        轮次行只是分组标题，勾选框留白；命中行的勾选框一律画成禁用态——
        链式的轮次是连在一起的，只替换其中一轮没有意义，勾选只按整个文件生效。
        """
        flat = 0
        for ridx, rnd in enumerate(file_data.get("rounds", []), 1):
            round_iid = _round_iid(file_iid, ridx)
            self._round_index[round_iid] = (file_idx, ridx)
            text = self._t("chain_round", n=ridx, f=rnd["find_text"],
                           r=rnd["replace_text"])
            tree.insert(file_iid, "end", iid=round_iid, tags=("round",),
                        values=(text, str(len(rnd["matches"])), "", ""))
            for m in rnd["matches"]:
                match_iid = _match_iid(round_iid, flat)
                self._iid_index[match_iid] = (file_idx, flat)
                # 轮次号单独记一份：第 2 轮起的命中在文件里还不存在（链还没跑），
                # 跳转与"双击可跳"的提示都要按它区别对待
                self._chain_round_of[match_iid] = ridx
                location = m["location"]
                if m.get("from_prev"):
                    # 位置列加后缀说明"这一处是上一轮改出来的"：颜色看不出来
                    # 的人（色弱、黑白打印）看字就懂；跳转用的是
                    # type/para_idx/occurrence，不读这串，改它不影响跳转
                    location = f"{location} · {self._t('loc_from_prev')}"
                tree.insert(round_iid, "end", iid=match_iid, text="",
                            image=self._check_image(False, disabled=True),
                            values=("", "", location, ""),
                            tags=("produced" if m.get("from_prev") else "match",))
                self._ctx_text[match_iid] = _ctx_parts(m)
                flat += 1

    def _insert_match_children(self, tree, file_iid, file_data, file_idx):
        """插入匹配子项"""
        multi = len(getattr(self, "active_conditions", []) or []) > 1
        for midx, md in enumerate(file_data["match_data"]):
            match_iid = _match_iid(file_iid, midx)
            self._iid_index[match_iid] = (file_idx, midx)
            m = md["match"]
            check_img = self._check_image(md["checked"].get())
            # 多条件时在「位置」列标出这条属于哪个条件——不标的话，用户看到
            # 一堆命中却分不清哪条会被替换成什么
            location = m["location"]
            if multi and m.get("cond_idx"):
                location = f"{self._t('cond_prefix', n=m['cond_idx'])} · {location}"
            tree.insert(file_iid, "end", iid=match_iid, text="", image=check_img,
                        values=("", "", location, ""),
                        tags=("match",))
            self._ctx_text[match_iid] = _ctx_parts(m)

    # ====================== 事件处理 ======================
    def on_tree_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        item_id = tree.identify_row(event.y)
        pos = self._iid_pos(item_id)
        if pos is None:
            return
        # 勾选态是 #0 列的图标，但 #0 列同时住着展开箭头——点在箭头上要留给
        # Tk 自己的展开功能，抢过来会变成"点箭头改勾选"。
        if tree.identify_column(event.x) != "#0":
            return
        if "indicator" in tree.identify_element(event.x, event.y):
            return
        file_idx, match_idx = pos
        if match_idx is None:
            self.toggle_file_check(item_id, tree)
        elif self.is_chain_mode:
            # 链式里轮次连在一起，只换其中一轮没有意义——行级勾选整体停用，
            # 勾选框也画成禁用态，不给"点了会有反应"的错觉
            return
        else:
            self.toggle_match_check(item_id, tree)

    def on_tree_double_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        item_id = tree.identify_row(event.y)
        pos = self._iid_pos(item_id)
        if pos is None:
            return
        file_idx, match_idx = pos
        if match_idx is None:
            cur = tree.item(item_id, "open")
            tree.item(item_id, open=not cur)
        elif getattr(self, "_round_index", {}).get(item_id) is not None:
            # 轮次行同样能双击折叠：它下面带着该轮的全部命中
            cur = tree.item(item_id, "open")
            tree.item(item_id, open=not cur)
        elif not self._chain_can_jump(item_id):
            return
        else:
            self.open_and_jump_to_match(item_id)

    def _chain_can_jump(self, match_iid):
        """该命中现在真的在文件里吗？

        链式预览里第 2 轮起的命中，是前几轮跑完才会出现的文本——文件此刻
        还是原始的，跳过去要不落空、要不跳到别处，不如直接不跳。
        """
        return self._chain_round_of.get(match_iid, 1) == 1

    # ====================== 勾选联动 ======================
    def toggle_file_check(self, file_iid, tree):
        pos = self._iid_pos(file_iid)
        if pos is None:
            return
        idx = pos[0]
        file_data = self.preview_data[idx]
        new_state = not file_data["checked"].get()
        file_data["checked"].set(new_state)

        for md in file_data.get("match_data", []):
            md["checked"].set(new_state)

        self._update_node_check(tree, file_iid, new_state)
        # 链式的子行是轮次行 + 禁用态命中行，勾选用不上也不该跟着翻
        if not self.is_chain_mode:
            for child in tree.get_children(file_iid):
                self._update_node_check(tree, child, new_state)
        self.update_checked_count()

    def toggle_match_check(self, match_iid, tree):
        # 行级勾选在链式下整体停用（不只是挡住点击）——轮次连在一起，
        # 只替换其中一轮没有意义，从入口到动作都不给留下缝
        if self.is_chain_mode:
            return
        pos = self._iid_pos(match_iid)
        if pos is None:
            return
        file_idx, match_idx = pos
        file_data = self.preview_data[file_idx]
        if "match_data" not in file_data:
            return
        md = file_data["match_data"][match_idx]
        new_state = not md["checked"].get()
        md["checked"].set(new_state)

        self._update_node_check(tree, match_iid, new_state)

        all_checked = all(m["checked"].get() for m in file_data["match_data"])
        file_data["checked"].set(all_checked)
        self._update_node_check(tree, _file_iid(file_idx), all_checked)
        self.update_checked_count()

    def _update_node_check(self, tree, iid, checked):
        tree.item(iid, image=self._check_image(checked))

    def _sync_all_check_states(self):
        """原地同步所有勾选状态（不重建树）"""
        tree = self.preview_tree
        if not tree:
            return
        for idx, fd in enumerate(self.preview_data):
            file_iid = _file_iid(idx)
            self._update_node_check(tree, file_iid, fd["checked"].get())
            for midx, md in enumerate(fd.get("match_data", [])):
                match_iid = _match_iid(file_iid, midx)
                try:
                    self._update_node_check(tree, match_iid, md["checked"].get())
                except tk.TclError:
                    pass

    def update_checked_count(self):
        if not hasattr(self, "checked_count_label") or not self.checked_count_label:
            return
        count = 0
        if self.is_chain_mode:
            # 链式只按整文件勾选，计数也按文件——数成一串命中数看着像能逐条选
            count = sum(1 for fd in self.preview_data if fd["checked"].get())
        else:
            for fd in self.preview_data:
                if fd.get("matches") and fd["matches"][0].get("type") == "filename":
                    if fd["checked"].get():
                        count += 1
                else:
                    for md in fd.get("match_data", []):
                        if md["checked"].get():
                            count += 1
        self.checked_count_label.config(text=self._t("checked_count", n=count))
        self._schedule_conflict_gate()

    def _schedule_conflict_gate(self):
        """勾选后合并刷新冲突闸门：连点 / 全选时只跑一次检测，不每次都卡界面。

        执行前拦截（execute_replace_from_preview）仍同步重算，防抖只影响
        「按钮按死 + 状态栏提示」的展示时机的 120ms，不会放行冲突批次。
        """
        prev = getattr(self, "_conflict_gate_after", None)
        if prev is not None:
            try:
                self.root.after_cancel(prev)
            except tk.TclError:
                pass
        self._conflict_gate_after = self.root.after(120, self._run_conflict_gate)

    def _run_conflict_gate(self):
        self._conflict_gate_after = None
        self._refresh_conflict_gate()

    def selected_conflicts(self):
        """当前**勾选中的**那批匹配里还剩哪些重叠命中，返回提示文案列表（以勾选状态为准，取消勾选即放行）。"""
        notes = []
        if self.is_chain_mode:
            # 跨轮的命中区间来自不同文本状态，拿来做重叠判断会凭空报冲突把
            # 整批锁死；链式本就不存在"同一位置被两条条件抢"这回事
            return notes
        for fd in self.preview_data:
            is_filename = bool(fd.get("matches")) and \
                fd["matches"][0].get("type") == "filename"
            if is_filename:
                if fd["checked"].get():
                    notes.extend(fd.get("conflicts") or [])
                continue
            hits = [md["match"] for md in fd.get("match_data", [])
                    if md["checked"].get()]
            if not hits:
                continue
            for hit, rival in detect_condition_conflicts(hits):
                notes.append(format_conflict_note(
                    f'{fd.get("file_name", "")} {hit.get("location", "?")}',
                    hit, rival))
        return notes

    def _refresh_conflict_gate(self):
        """有重叠命中就把「替换勾选项」按死，并在状态栏说明原因（不自动裁决、不部分执行）。"""
        label = getattr(self, "conflict_label", None)
        btn = getattr(self, "preview_exec_btn", None)
        if label is None or btn is None:
            return
        try:
            if not label.winfo_exists():
                return
        except tk.TclError:
            return
        conflicts = self.selected_conflicts()
        if conflicts:
            label.config(text=self._t("conflict_gate", n=len(conflicts)), fg=self.DANGER_FG)
            btn.config(state="disabled")
            btn.configure(bg=self.ACCENT_OFF_BG, fg=self.ACCENT_OFF_FG)
        else:
            label.config(text="", fg=self.DANGER_FG)
            btn.config(state="normal")

    # ====================== 批量操作 ======================
    def select_all_rows(self):
        for fd in self.preview_data:
            fd["checked"].set(True)
            for md in fd.get("match_data", []):
                md["checked"].set(True)
        self._sync_all_check_states()
        self.update_checked_count()

    def deselect_all_rows(self):
        for fd in self.preview_data:
            fd["checked"].set(False)
            for md in fd.get("match_data", []):
                md["checked"].set(False)
        self._sync_all_check_states()
        self.update_checked_count()

    def expand_all_files(self):
        tree = self.preview_tree
        if not tree:
            return
        # 链式预览的命中挂在轮次行下，只展开文件行等于没展开——两层都要动
        for idx in range(len(self.preview_data)):
            file_iid = _file_iid(idx)
            tree.item(file_iid, open=True)
            for child in tree.get_children(file_iid):
                tree.item(child, open=True)
        self._refresh_parent_hints(tree)

    def collapse_all_files(self):
        tree = self.preview_tree
        if not tree:
            return
        for idx in range(len(self.preview_data)):
            file_iid = _file_iid(idx)
            tree.item(file_iid, open=False)
            for child in tree.get_children(file_iid):
                tree.item(child, open=False)
        self._refresh_parent_hints(tree)

    # ========== 替换执行 ==========
    def execute_replace_from_preview(self, preview_window, conditions):
        # 先校验、后关窗：若第一行就 destroy 预览窗，"一条都没勾就点替换"时
        # 警告框没有父窗口、预览窗也已经没了——用户丢失全部上下文，只能重搜一遍。
        checked_rows = []
        # 链式只按整个文件勾选：轮次连在一起，单独替换一轮没有意义，
        # 文件行的勾选状态就是全部入参。
        if self.is_chain_mode:
            for fd in self.preview_data:
                if not fd["checked"].get() or not fd.get("matches"):
                    continue
                checked_rows.append({
                    "file_path": fd["file_path"],
                    "file_name": fd["file_name"],
                    "match": fd["matches"][0]
                })
        else:
            for fd in self.preview_data:
                if fd.get("matches") and fd["matches"][0].get("type") == "filename":
                    if fd["checked"].get():
                        checked_rows.append({
                            "file_path": fd["file_path"],
                            "file_name": fd["file_name"],
                            "match": fd["matches"][0]
                        })
                else:
                    for md in fd.get("match_data", []):
                        if md["checked"].get():
                            checked_rows.append({
                                "file_path": fd["file_path"],
                                "file_name": fd["file_name"],
                                "match": md["match"]
                            })

        if not checked_rows:
            messagebox.showwarning(self._t("t_warning"), self._t("m_select_match"),
                                   parent=preview_window)
            return

        # 替换批次进行中禁止再发起一轮：预览窗关掉前快速连点两次、或从其他
        # 入口触发都会走进来（理由见 _replace_in_progress）。
        if self._replace_in_progress():
            messagebox.showwarning(self._t("t_warning"), self._t("m_replacing"),
                                   parent=preview_window)
            return

        has_filename = any(r["match"].get("type") == "filename" for r in checked_rows)
        has_content = any(r["match"].get("type") != "filename" for r in checked_rows)

        # 混合批次（既改名又改内容）里，先执行的分支会让另一分支手里的路径
        # 失效——一次搜索只产一种模式，真出现就是勾选数据异常，整批拦下。
        if has_filename and has_content:
            messagebox.showerror(self._t("t_error"), self._t("m_mixed_batch"),
                                 parent=preview_window)
            self.log(self._t("m_mixed_batch"))
            return

        # xlsx 内嵌对象预警放在建备份文件夹之前——用户此时取消不会留下空备份文件夹。
        # openpyxl 整车 load/save 会重写整个包，内嵌图片 / 图表 / 绘图 / 嵌入对象
        # 可能丢失，保存前先查 zip 告知并给取消机会，不做静默损坏。
        risky_files = []
        seen_risky = set()
        for row in checked_rows:
            fp = row["file_path"]
            if fp in seen_risky or not str(fp).lower().endswith(".xlsx"):
                continue
            seen_risky.add(fp)
            kinds = xlsx_embedded_objects(fp)
            if kinds:
                risky_files.append((fp, kinds))
        if risky_files:
            items = "\n".join(
                self._t("m_embedded_item", name=os.path.basename(fp), kinds="、".join(kinds))
                for fp, kinds in risky_files)
            if not messagebox.askyesno(self._t("t_notice"),
                                       self._t("m_embedded", files=items),
                                       parent=preview_window):
                self.log(self._t("lg_cancelled_embedded"))
                return

        # 重叠命中一律整批停下：不自动裁决、不做部分替换。位置必须早于建备份夹，
        # 否则用户吃中止提示时还多一个空文件夹。这里与预览按钮的禁用共用
        # selected_conflicts，"按钮灰着"与"点了也拦"不会给出两种结论。
        conflicts = self.selected_conflicts()
        if conflicts:
            messagebox.showerror(self._t("t_overlap_stopped"),
                                 conflict_block_message(conflicts),
                                 parent=preview_window)
            self.log(self._t("lg_overlap_abort", n=len(conflicts)))
            for note in conflicts[:200]:
                self.log(self._t("lg_overlap_blocked", note=note))
            return

        # 多条件时用「多条件N项」当查找词——把 N 条规则拼进文件夹名会又长又乱，
        # 规则明细在升级日志与预览窗里都能查到。
        backup_enabled = self.backup_var.get()
        backup_path = self.backup_path.get().strip()
        backup_folder = None
        # 开了备份却没有路径就当场中止：界面还显示「备份开启」，用户以为有
        # 退路其实没有，是最坏的一种。
        if backup_enabled and not backup_path:
            messagebox.showerror(self._t("t_backup_no_path"),
                                 self._t("m_backup_no_path"),
                                 parent=preview_window)
            self.log(self._t("lg_backup_no_path"))
            return
        if backup_enabled and backup_path:
            if len(conditions) == 1:
                b_find, b_replace = conditions[0]
            else:
                b_find, b_replace = self._t("multi_cond_tag", n=len(conditions)), ""
            try:
                backup_folder = create_backup_folder(
                    backup_path, b_find, b_replace,
                    self.replace_filename_var.get())
                self.log(self._t("lg_backup_folder", path=backup_folder))
            except Exception as e:
                # 备份建不出来就中止，不能"只打一行警告然后无备份替换"。
                # 本工具的安全立场是强制备份，做不到备份就不该动用户的文件。
                messagebox.showerror(
                    self._t("t_backup_failed"),
                    self._t("m_backup_failed", e=str(e)),
                    parent=preview_window)
                self.log(self._t("lg_backup_failed", e=str(e)))
                return

        # 二次确认必须覆盖所有不可逆操作：只给内容替换加确认框是不够的，
        # 纯文件名替换点了就直接改名——改名同样不可逆。混合场景（理论上不会
        # 出现，因为一次搜索只有一种模式）也一并说清楚，不会漏确认。
        # checked_rows 非空时 has_filename / has_content 必居其一，确认无条件弹。
        backup_status = self._t("backup_enabled_word" if backup_enabled
                                else "backup_disabled_word")
        lines = []
        if len(conditions) == 1:
            lines.append(self._t("confirm_line_find", v=conditions[0][0]))
            lines.append(self._t("confirm_line_replace", v=conditions[0][1]))
        elif self.is_chain_mode:
            # 链式要额外说清"顺序有意义"：条件表在这里的排列就是执行顺序
            lines.append(self._t("confirm_line_chain", n=len(conditions)))
            lines += [f"  {i}. {f} → {r}" for i, (f, r) in enumerate(conditions, 1)]
        else:
            lines.append(self._t("confirm_line_multi", n=len(conditions)))
            lines += [f"  {i}. {f} → {r}" for i, (f, r) in enumerate(conditions, 1)]
        lines.append(self._t("confirm_line_backup", v=backup_status))
        if has_filename:
            lines.append(self._t("confirm_line_rename", n=len(
                [r for r in checked_rows if r["match"].get("type") == "filename"])))
        if has_content:
            content_rows = [r for r in checked_rows if r["match"].get("type") != "filename"]
            if self.is_chain_mode:
                # 链式一条 = 一个文件（勾选粒度），命中数要按轮次展开数才准
                n = sum(match_hit_count(m) for fd in self.preview_data
                        if fd["checked"].get() and fd.get("matches")
                        and fd["matches"][0].get("type") != "filename"
                        for m in fd["matches"])
            else:
                n = sum(match_hit_count(r["match"]) for r in content_rows)
            lines.append(self._t("confirm_line_matches", n=n, m=len(
                set(r["file_path"] for r in content_rows))))
        msg = self._t("m_confirm_replace") + "\n\n" + "\n".join(lines)
        if not messagebox.askyesno(self._t("t_confirm"), msg, parent=preview_window):
            return

        # 校验与确认全部通过之后，才关掉预览窗
        preview_window.destroy()

        if has_filename:
            filename_rows = [r for r in checked_rows if r["match"]["type"] == "filename"]
            # 先预检、再动手：直接 os.rename 的话，非法字符要等 rename 才报错，
            # 批内撞名、目标文件已存在都不预检（POSIX 上 os.rename 还会静默覆盖
            # 别人的文件）。
            plan, warnings = plan_filename_renames([
                {"file_path": r["file_path"], "new_name": r["match"].get("new_name")}
                for r in filename_rows])
            for w in warnings:
                self.log(self._t("lg_fname_precheck_item", w=w))
            if warnings:
                self.log(self._t("lg_fname_precheck_sum", n=len(warnings)))
            success_count = 0
            error_count = 0
            for file_path, new_path in plan:
                try:
                    if backup_folder:
                        create_backup(file_path, backup_folder)
                    os.rename(file_path, new_path)
                    # OK / Error 是跨语言通用的状态前缀，不翻译
                    self.log(f"OK: {os.path.basename(file_path)} -> {os.path.basename(new_path)}")
                    success_count += 1
                    # 主列表跟着换行（新文件名 / 新状态），见 _rename_file_row
                    self._rename_file_row(file_path, new_path)
                except Exception as e:
                    self.log(f"Error: {file_path} - {e}")
                    error_count += 1
            self.log(self._t("lg_fname_done", ok=success_count, fail=error_count))

        if has_content:
            content_rows = [r for r in checked_rows if r["match"].get("type") != "filename"]
            if self.is_chain_mode:
                # 链式不传命中清单：替换时把整条链再跑一遍（与预览取样同一套
                # 过程），结果与预览逐轮一致
                chain_files = list(dict.fromkeys(r["file_path"] for r in content_rows))
                thread = threading.Thread(
                    target=self.replace_chain_files,
                    args=(chain_files, backup_folder, self.case_sensitive_var.get(),
                          backup_enabled, backup_path, conditions))
                thread.daemon = True
                self._replace_thread = thread
                thread.start()
                return
            # 按文件归组：同一文件的全部条件命中一起交给引擎，让它在原始内容上
            # 一次替换完（多条件不能拆成多轮跑，见「多条件替换」一节的语义 1）。
            file_groups = {}
            for row in content_rows:
                file_path = row["file_path"]
                file_groups.setdefault(file_path, []).append(row["match"])

            # 工作线程需要的一切都在主线程先读成普通值传进去，线程里不再触碰
            # 任何 tk 对象（含 BooleanVar.get()）。同时记下线程引用——搜索
            # 重入守卫与关窗确认都看它。
            thread = threading.Thread(
                target=self.replace_files_from_preview,
                args=(file_groups, backup_folder,
                      self.case_sensitive_var.get(), backup_enabled, backup_path,
                      conditions))
            thread.daemon = True
            self._replace_thread = thread
            thread.start()

    def _replace_batch_prologue(self, start_msg, total_files, case_sensitive,
                                backup_enabled, backup_path):
        """替换工作线程公共开局：锁住入口 + 打印批次头（条件 / 文件数 / 备份 / 大小写）。

        跑起来之前先锁入口、收尾才解——不锁的话替换途中主按钮与文件清单仍是
        活的（点了会被重入守卫挡下并弹提示，但界面看着像还能动）。
        """
        self._ui(self._set_primary_state, "disabled")
        self._ui(self._set_file_ops_state, False)
        self._ui(self.log, "=" * 60)
        self._ui(self.log, start_msg)
        self._ui(self.log, self._t("lg_file_count", n=total_files))
        self._ui(self.log, self._t("lg_backup_state",
                                   v=backup_path if backup_enabled
                                   else self._t("lg_backup_disabled_risk")))
        self._ui(self.log, self._t("lg_case_state",
                                   v=self._t("yes") if case_sensitive else self._t("no")))
        self._ui(self.log, "=" * 60)

    def _replace_batch_epilogue(self, success, fail, total_files, backup_folder,
                                conditions, chain, manifest_files,
                                case_sensitive, done_files):
        """替换工作线程公共收尾：进度复位 + 汇总 + 回滚 manifest + 解锁 + 完成框。

        完成框收敛到 _show_replace_done：直接 _ui(messagebox.showinfo) 不传
        parent，主窗最小化时框不一定到前台，用户会以为卡死。
        """
        self._ui(self._set_replace_progress, None, total_files, "")
        self._ui(self.log, "=" * 60)
        self._ui(self.log, self._t("lg_done_summary", ok=success, fail=fail))
        self._ui(self.log, "=" * 60)
        self._write_rollback_manifest(backup_folder, conditions, chain,
                                      manifest_files, case_sensitive)
        self._ui(self._mark_files_replaced, done_files)
        self._ui(self._set_primary_state, "normal")
        self._ui(self._set_file_ops_state, True)
        self._ui(self._show_replace_done, success, fail)

    def replace_chain_files(self, files, backup_folder=None, case_sensitive=False,
                            backup_enabled=True, backup_path="", conditions=None):
        """链式替换工作线程：每个文件在临时副本上跑完整条链，成了才搬回原位。

        commit 前原文件一个字节都没动——链中途任何一轮出错，都会退化成
        "什么都没改"，而不是留下一个改到一半的文件。
        """
        conditions = conditions or []
        total_files = len(files)
        self._replace_batch_prologue(
            self._t("lg_batch_start_chain", n=len(conditions),
                    summary=conditions_summary(conditions)),
            total_files, case_sensitive, backup_enabled, backup_path)

        success, fail = 0, 0
        done_files = []
        manifest_files = []
        for done, file_path in enumerate(files, 1):
            self._ui(self._set_replace_progress, done, total_files,
                     os.path.basename(file_path))
            # 前置换行让每个文件的处理块在日志里分开，便于扫读
            self._ui(self.log, "\n" + self._t("lg_processing",
                                              name=os.path.basename(file_path)))
            try:
                backup_target = None
                if backup_folder:
                    backup_target = create_backup(file_path, backup_folder)
                rounds, skipped = run_replace_chain(
                    file_path, conditions, case_sensitive, commit=True)
                for loc, err in skipped:
                    self._ui(self.log, self._t("lg_skip_warn", loc=loc, err=err))
                self._ui(self.log, self._t(
                    "lg_success_one", name=os.path.basename(file_path),
                    n=sum(match_hit_count(m) for r in rounds
                          for m in r["matches"])))
                success += 1
                done_files.append(file_path)
                if backup_target:
                    manifest_files.append((file_path, backup_target))
            except Exception as e:
                self._ui(self.log, self._t("lg_fail_one", e=str(e)))
                fail += 1

        self._replace_batch_epilogue(success, fail, total_files, backup_folder,
                                     conditions, True, manifest_files,
                                     case_sensitive, done_files)

    def replace_files_from_preview(self, file_groups, backup_folder=None,
                                   case_sensitive=False, backup_enabled=True,
                                   backup_path="", conditions=None):
        """工作线程：执行替换。

        界面操作一律经 self._ui 投递回主线程——子线程直接调 self.log() 属 tkinter 未定义行为。
        """
        total_files = len(file_groups)
        if conditions and len(conditions) == 1:
            start_msg = self._t("lg_batch_start_single",
                                f=conditions[0][0], r=conditions[0][1])
        elif conditions:
            start_msg = self._t("lg_batch_start_multi",
                                n=len(conditions),
                                summary=conditions_summary(conditions))
        else:
            start_msg = self._t("lg_batch_start_plain")
        self._replace_batch_prologue(start_msg, total_files, case_sensitive,
                                     backup_enabled, backup_path)

        success, fail = 0, 0
        done_files = []
        manifest_files = []
        for done, (file_path, matches) in enumerate(file_groups.items(), 1):
            # 替换进度进状态栏（日志面板默认收起，进度必须常驻可见）
            self._ui(self._set_replace_progress, done, total_files, os.path.basename(file_path))
            try:
                # 前置换行让每个文件的处理块在日志里分开，便于扫读
                self._ui(self.log, "\n" + self._t("lg_processing",
                                                  name=os.path.basename(file_path)))
                backup_target = None
                if backup_folder:
                    backup_target = create_backup(file_path, backup_folder)
                ext = Path(file_path).suffix.lower()
                engine = _engine_for(ext)
                if engine is None:
                    raise Exception(self._t("e_unsupported_type", ext=ext))
                if not engine["available"]:
                    raise Exception(self._t(engine["missing_replace"]))
                skipped = engine["replace"](file_path, matches, case_sensitive)
                for loc, err in skipped:
                    self._ui(self.log, self._t("lg_skip_warn", loc=loc, err=err))
                self._ui(self.log, self._t("lg_success_one",
                                           name=os.path.basename(file_path),
                                           n=sum(match_hit_count(m)
                                                 for m in matches)))
                success += 1
                done_files.append(file_path)
                # 只记实际替换成功的文件：跳过 / 失败的不入 manifest，否则
                # 回滚会去恢复一个根本没被改过的文件
                if backup_target:
                    manifest_files.append((file_path, backup_target))
            except Exception as e:
                self._ui(self.log, self._t("lg_fail_one", e=str(e)))
                fail += 1

        self._replace_batch_epilogue(success, fail, total_files, backup_folder,
                                     conditions, False, manifest_files,
                                     case_sensitive, done_files)

    def _write_rollback_manifest(self, backup_folder, conditions, chain, files,
                                 case_sensitive):
        """内容模式替换成功后，把 manifest 写进当次备份夹（回溯的依据）。

        chain 记下这批条件是链式还是并行跑的——回溯列表的标签要说清执行
        方式，否则用户按并行的预期去理解一条链式操作的结果会对不上。
        没开备份就没有备份夹，回溯本就无从做起，这里直接跳过。
        """
        if not backup_folder or not files:
            return
        conditions = [[str(f), str(r)] for f, r in (conditions or [])]
        data = {
            "version": MANIFEST_VERSION,
            "ts": datetime.now().strftime(TS_FMT),
            "mode": "content",
            "chain": bool(chain),
            "multi": len(conditions) > 1,
            "conditions": conditions,
            "case_sensitive": bool(case_sensitive),
            "files": [{"orig": orig, "orig_canon": canonical_path(orig),
                       "backup_rel": os.path.basename(backup_target)}
                      for orig, backup_target in files],
        }
        try:
            write_manifest(backup_folder, data)
            self._ui(self.log, self._t("lg_manifest_written", n=len(files)))
        except Exception as e:
            # 替换本身已成功、备份也在，但这一次操作从此无法回溯——必须说清
            self._ui(self.log, self._t("lg_manifest_failed", e=str(e)))

    def _show_replace_done(self, success, fail):
        """在主线程弹完成框：先提主窗到前台，再挂 parent 显示。"""
        try:
            self.root.lift()
            self.root.focus_force()
        except tk.TclError:
            pass
        messagebox.showinfo(self._t("t_done"),
                            self._t("m_done_replace", ok=success, fail=fail),
                            parent=self.root)

    def _mark_files_replaced(self, paths):
        """替换完成后刷新主列表：已处理文件的匹配数归零、状态改为「已替换」。"""
        tree = getattr(self, "files_tree", None)
        if tree is None:
            return
        for path in paths:
            try:
                if tree.exists(path):
                    tree.set(path, "match", "0")
                    tree.set(path, "status", self._t("st_replaced"))
            except tk.TclError:
                pass
        # 重算总匹配数，让状态栏与实际列表一致
        self._recalc_total_matches()

    def open_and_jump_to_match(self, match_iid):
        pos = self._iid_pos(match_iid)
        if pos is None:
            return
        file_idx, match_idx = pos
        if match_idx is None:
            # 文件行没有"某一个匹配位置"可跳，直接返回。
            # 若不提前返回，fd["match_data"][None] 会抛 TypeError，再被下方
            # except 兜底成"降级打开文件"——日志里平白多一条误导性错误，
            # 行为也从"无操作"变成了打开文件。
            return
        fd = self.preview_data[file_idx]
        file_path = fd["file_path"]
        match = fd["match_data"][match_idx]["match"]
        ext = os.path.splitext(file_path)[1].lower()

        try:
            engine = _engine_for(ext)
            if engine is None:
                messagebox.showinfo(self._t("t_notice"),
                                    self._t("m_unsupported_type", ext=ext))
            else:
                getattr(self, engine["jump"])(file_path, match)
        except Exception as e:
            # win32com 失败时降级为直接打开文件
            self.log(self._t("lg_jump_fallback", e=str(e)))
            try:
                os.startfile(file_path)
                messagebox.showinfo(self._t("t_notice"), self._t("m_opened_no_loc", e=str(e)))
            except Exception as e2:
                messagebox.showerror(self._t("t_jump_failed"),
                                     self._t("m_jump_failed", e=str(e2)))

    def _dispatch_com(self, prog_id):
        """获取 Office COM 实例，返回 (app, reused)。

        优先复用已在运行的实例（否则每文件叠一个新进程）；reused=True 时
        绝不能 Quit()——那是用户自己的 Word/Excel。
        """
        try:
            return win32com.client.GetActiveObject(prog_id), True
        except Exception:
            return win32com.client.Dispatch(prog_id), False

    def jump_to_word_location(self, file_path, match):
        """增强Word跳转 - 打开后定位到段落并高亮查找内容"""
        if not WIN32COM_AVAILABLE:
            raise Exception(self._t("e_win32com"))

        word, reused = self._dispatch_com("Word.Application")
        word.Visible = True
        doc = None
        opened = False

        try:
            doc = word.Documents.Open(os.path.abspath(file_path))
            opened = True  # 文件打开成功
            container = None  # 定位到的容器 Range，交给高亮在其中精确选词

            if match["type"] == "paragraph":
                para_idx = match["para_idx"]
                if para_idx <= doc.Paragraphs.Count:
                    para = doc.Paragraphs(para_idx)
                    container = para.Range
                    container.Select()
                else:
                    raise Exception(self._t("e_para_out_of_range", i=para_idx, n=doc.Paragraphs.Count))

            elif match["type"] == "table":
                table_idx = match["table_idx"]
                row_idx = match["row_idx"]
                col_idx = match["col_idx"]
                if table_idx <= doc.Tables.Count:
                    table = doc.Tables(table_idx)
                    if row_idx <= table.Rows.Count and col_idx <= table.Columns.Count:
                        cell = table.Cell(row_idx, col_idx)
                        container = cell.Range
                        container.Select()
                    else:
                        raise Exception(self._t("e_table_coord", r=row_idx, c=col_idx))
                else:
                    raise Exception(self._t("e_table_out_of_range", i=table_idx, n=doc.Tables.Count))

            elif match["type"] in ("header", "footer"):
                section_idx = match["section_idx"]
                para_idx = match["para_idx"]
                if section_idx <= doc.Sections.Count:
                    section = doc.Sections(section_idx)
                    if match["type"] == "header":
                        loc = section.Header
                    else:
                        loc = section.Footer
                    if para_idx <= loc.Range.Paragraphs.Count:
                        para = loc.Range.Paragraphs(para_idx)
                        container = para.Range
                        container.Select()
                    else:
                        raise Exception(self._t("e_hf_para_out_of_range", kind=match["type"], i=para_idx))
                else:
                    raise Exception(self._t("e_section_out_of_range", i=section_idx))
            else:
                raise Exception(self._t("e_match_type", t=match["type"]))

            # 定位成功后才尝试高亮（高亮失败不触发关闭）；词级高亮失败时，
            # 容器（整段 / 单元格）仍是选中状态，用户不会连落点都找不到。
            self._word_highlight_match(container, match)
            return  # ① 定位成功：文档保持打开状态，交还给用户

        except Exception:
            # 三种结局分清楚：文档已打开但定位失败 → 只关这一篇；实例是自己
            # 新建的 → 才 Quit。opened 标志若在 Open 成功后立刻置 True 再据此判断，
            # "定位失败"反而会把刚打开的用户文档一起 Quit 掉，与意图相反。
            if opened and doc is not None:
                # ③ 文档已打开、但定位失败：关掉这篇（是我们刚打开的，别留半截状态）
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            if not reused:
                # ② 实例是我们自己新建的（含 Open 就失败的情况）：退出它
                #    复用的实例绝不能 Quit —— 那是用户自己的 Word
                try:
                    word.Quit()
                except Exception:
                    pass
            raise

    def _word_highlight_match(self, container, match):
        """在已定位的容器（段落 / 单元格）内，选中用户点的那一条匹配。

        按 match 里记下的出现序号逐次 Find，不停在第一个（同段同词多次会选错
        位置）；也不用字符偏移——制表符、软换行使 docx 侧偏移与 Word 文本错位。
        """
        text = match.get("find_text") or self.find_entry.get()
        if not text:
            return
        if "^" in text:
            # ^ 在 Word 的查找文本里是特殊代码前缀（^p 段落标记、^t 制表符…），
            # 与提取侧的纯文本匹配不是同一套语义，出现序号会对不上。宁可不做
            # 词级高亮（用户仍能看到整段被选中），也不亮到错误的位置。
            self.log(self._t("lg_caret_skip"))
            return
        occurrence = match.get("occurrence", 1)
        match_case = bool(match.get("case_sensitive", False))
        try:
            hit = _word_find_nth(container, text, occurrence, match_case)
            if hit is not None:
                hit.Select()
        except Exception as e:
            # 高亮失败不影响主流程，但要在日志里留痕——不能静默 pass
            self.log(self._t("lg_highlight_failed", e=e))

    def jump_to_excel_location(self, file_path, match):
        """增强Excel跳转 - 打开后定位单元格"""
        if not WIN32COM_AVAILABLE:
            raise Exception(self._t("e_win32com"))

        excel, reused = self._dispatch_com("Excel.Application")
        excel.Visible = True
        wb = None
        opened = False

        try:
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            opened = True

            sheet_name = match["sheet_name"]
            cell_coord = match["cell_coordinate"]

            sheet = None
            for s in wb.Worksheets:
                if s.Name == sheet_name:
                    sheet = s
                    break

            if sheet is None:
                available = ", ".join(s.Name for s in wb.Worksheets)
                raise Exception(self._t("e_sheet_missing", name=sheet_name, avail=available))

            sheet.Activate()
            rng = sheet.Range(cell_coord)
            rng.Select()
            excel.ActiveWindow.ScrollRow = rng.Row
            excel.ActiveWindow.ScrollColumn = rng.Column
            return  # ① 定位成功：文件保持打开状态

        except Exception:
            # 同 Word：定位失败先关掉我们自己打开的那份，再按需退出实例；
            # 复用的实例（用户本来就开着的 Excel）绝不 Quit。
            if opened and wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if not reused:
                try:
                    excel.Quit()
                except Exception:
                    pass
            raise

    def jump_to_text_location(self, file_path, match):
        """txt 无法在记事本里精确定位到字符，直接打开文件并提示命中行/列。"""
        try:
            os.startfile(file_path)
            # 记事本启动后会抢走焦点，提示框若不置顶就被压在它下面，看起来像没反应。
            # 临时置顶主窗口让提示框浮到最前，点掉后立即还原，主窗不常驻置顶。
            self.root.attributes("-topmost", True)
            try:
                messagebox.showinfo(self._t("t_notice"),
                                    self._t("m_txt_jump", loc=match.get("location", "")),
                                    parent=self.root)
            finally:
                self.root.attributes("-topmost", False)
        except Exception as e:
            messagebox.showerror(self._t("t_jump_failed"),
                                 self._t("m_jump_failed", e=str(e)))

    def _open_url(self, url):
        """用系统默认浏览器打开链接；打不开就静默跳过。"""
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def _make_theme_badge(self, parent):
        """工具栏主题切换圆钮：太阳/月亮符号随当前主题变。"""
        cv = tk.Canvas(parent, width=26, height=26, bg=parent.cget("bg"),
                       highlightthickness=0, bd=0, cursor="hand2")
        f = tkfont.nametofont("TkDefaultFont").copy()
        f.configure(size=11)

        def repaint():
            cv.delete("all")
            cv._br_img = self._aa_round_image(26, 26, 13, self.BTN_BG, self.BTN_BD,
                                              master=cv)
            cv.create_image(13, 13, image=cv._br_img)
            cv.create_text(13, 13, text="\u2600" if self.theme == "light" else "\u263d",
                           fill=self.TXT_DIM, font=f)

        cv._br_repaint = repaint
        cv.bind("<Button-1>", lambda _e: self._toggle_theme())
        repaint()
        cv.pack(side="right", padx=2, pady=5)
        # 留个句柄：这颗钮曾漏过 pack，建了却永不出现，只有从外部查得到
        # 控件、才谈得上断言"它真的在工具栏上"。
        self.theme_badge = cv
        return cv

    def _toggle_theme(self):
        """深浅主题切换：捕获运行时状态 → 按新主题重建界面 → 恢复状态。

        走重建而非逐控件改色——色板同时影响 ttk 样式、Treeview tag、自绘
        Canvas 与全部 tk 控件，逐件刷漏一个就是一块旧主题色。
        """
        if self._search_in_progress() or self._replace_in_progress():
            return
        if getattr(self, "preview_window", None) is not None:
            messagebox.showinfo(self._t("t_notice"), self._t("m_theme_close_preview"))
            return
        state = self._capture_state()
        self.theme = "dark" if self.theme == "light" else "light"
        self.save_config()
        # 悬停提示与它的延时句柄先收掉：提示窗虽是 root 的子窗体、会随下面的
        # 销毁循环一起没，但句柄不取消的话，重建后到点仍会凭旧引用弹一个出来
        self._hide_chain_tip()
        for w in self.root.winfo_children():
            w.destroy()
        self._apply_theme(self.theme)
        self.setup_ui()
        self._restore_state(state)

    def _capture_state(self):
        """主题重建前的运行时状态快照。"""
        rows = []
        tree = getattr(self, "files_tree", None)
        if tree is not None:
            for iid in tree.get_children():
                try:
                    rows.append((iid, tree.set(iid, "type"), tree.set(iid, "size"),
                                 tree.set(iid, "match"), tree.set(iid, "status")))
                except tk.TclError:
                    pass
        log_text = ""
        if hasattr(self, "log_text"):
            try:
                log_text = self.log_text.get("1.0", "end-1c")
            except tk.TclError:
                pass
        return {
            "files": list(self.selected_files),
            "rows": rows,
            "conditions": [list(c) for c in self.conditions],
            "multi": self.multi_condition_var.get(),
            "multi_mode": self._multi_mode,
            "case": self.case_sensitive_var.get(),
            "fname": self.replace_filename_var.get(),
            "backup": self.backup_var.get(),
            "log": log_text,
            "log_visible": self.log_visible,
        }

    def _restore_state(self, st):
        """重建后把快照写回界面。"""
        for path, typ, size, match, status in st["rows"]:
            stripe = "even" if len(self.files_tree.get_children()) % 2 else "odd"
            self.files_tree.insert("", "end", iid=path, text=os.path.basename(path),
                                   values=(typ, size, match, status), tags=(stripe,))
        self.selected_files = list(st["files"])
        self.conditions = [list(c) for c in st["conditions"]]
        self._refresh_condition_tree()
        self.multi_condition_var.set(st["multi"])
        self._multi_mode = st.get("multi_mode", "parallel")
        # 内部模式是直接赋值的，显示串不会跟着变（换主题重建后下拉框会仍旧显示
        # 旧文案，实际执行的却是另一条）——按内部模式重写一遍显示串
        self._sync_multi_mode_texts()
        self.case_sensitive_var.set(st["case"])
        self.replace_filename_var.set(st["fname"])
        self.backup_var.set(st["backup"])
        if st["log"]:
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, st["log"] + "\n")
            self.log_text.config(state="disabled")
        if st["log_visible"] != self.log_visible:
            self.toggle_log()
        self.toggle_multi_condition(announce=False)
        self._refresh_find_replace_labels()
        self._update_mode_hint()
        self._recalc_total_matches()

    def _make_about_badge(self, parent):
        """右上角圆形「!」钮：关于窗入口。tk 没有圆形按钮，用 Canvas 自绘。"""
        cv = tk.Canvas(parent, width=26, height=26, bg=self.BG,
                       highlightthickness=0, bd=0, cursor="hand2")
        f = tkfont.nametofont("TkDefaultFont").copy()
        f.configure(weight="bold")

        def repaint(hover=False):
            cv.delete("all")
            cv._br_img = self._aa_round_image(
                26, 26, 13, self.BTN_BG,
                self.BTN_BD_HOVER if hover else self.BTN_BD, master=cv)
            cv.create_image(13, 13, image=cv._br_img)
            cv.create_text(13, 13, text="!", fill=self.TXT_DIM, font=f)

        cv._br_repaint = repaint
        repaint()

        cv.bind("<Enter>", lambda _e: repaint(hover=True))
        cv.bind("<Leave>", lambda _e: repaint(hover=False))
        cv.bind("<Button-1>", lambda _e: self.show_about())
        cv.pack(side="right", padx=(2, 8), pady=5)
        return cv

    def silent_check_update(self):
        """启动时自动检查一次，复用 check_for_update 的清单校验与三选一流程。"""
        self.check_for_update(self.root, silent=True)

    def ask_update_choice(self, parent, latest):
        """三选一：直接下载 / 发布页 / 暂不更新。返回 download/page/skip。"""
        chosen = {"value": "skip"}
        dlg = tk.Toplevel(parent)
        dlg.title(self._t("t_update_found"))
        dlg.transient(parent)
        dlg.resizable(False, False)

        def close(choice):
            chosen["value"] = choice
            dlg.destroy()

        # grab 放在控件建完之后：建到一半抛异常时窗口还没抓输入，异常照常
        # 上抛，用户只是看到一次报错，不会整窗僵死。反过来（先 grab 再建）
        # 一旦中途失败，空壳窗会一直抓着键盘、主界面点不动。
        try:
            page = ttk.Frame(dlg, padding=(20, 16))
            page.pack(fill="both", expand=True)
            ttk.Label(page, anchor="w",
                      text=self._t("m_update_found_short", new=latest,
                                   cur="v" + VERSION),
                      font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")
            ttk.Label(page, text=self._t("l_update_choose"),
                      anchor="w").pack(anchor="w")
            row = ttk.Frame(page)
            row.pack(pady=(14, 0), fill="x")
            ttk.Button(row, text=self._t("b_update_page"), width=10,
                       command=lambda: close("page")).pack(side="left")
            ttk.Button(row, text=self._t("b_update_download"), width=12,
                       command=lambda: close("download")).pack(side="left", padx=(8, 0))
            ttk.Button(row, text=self._t("b_update_skip"), width=10,
                       command=lambda: close("skip")).pack(side="left", padx=(8, 0))

            dlg.protocol("WM_DELETE_WINDOW", lambda: close("skip"))
            dlg.bind("<Escape>", lambda _e: close("skip"))
            # 摆位走公共路径：它先 update 再取尺寸。只 update_idletasks 的话
            # 窗口还没拿到 WM 给的尺寸，winfo_width 是 1，位置会算到左上角。
            self._center_dialog(dlg, parent)
        except Exception:
            # 建窗失败就把半成品拆掉，别把没有关闭按钮的空窗留给用户。
            try:
                dlg.destroy()
            except tk.TclError:
                pass
            raise
        self._safe_grab(dlg)
        dlg.wait_window()
        return chosen["value"]

    def _handle_update_found(self, parent, manifest, latest):
        """发现新版：三选一后分派。"""
        choice = self.ask_update_choice(parent, latest)
        if choice == "download":
            self.log(self._t("lg_update_download", new=latest))
            self._download_and_launch(parent, manifest, latest)
        elif choice == "page":
            target = manifest.get("page") or RELEASE_PAGE_URL
            self.log(self._t("lg_update_page", url=target))
            self._open_url(target)
        else:
            self.log(self._t("lg_update_skip", new=latest))

    def check_for_update(self, parent=None, silent=False):
        """检查更新：手动按钮与启动检查共用。清单校验 → 三选一。

        「没有新版」「检查失败」只写日志、不弹说明窗：这两种弹窗是纯噪音。
        silent=True（启动检查）不影响发现新版的处理——发现新版照常弹三选一，
        启动时也得让用户知道有新版并选择怎么处理。
        """
        parent = parent or self.root
        box = queue.Queue()

        def worker():
            box.put(fetch_update_manifest())

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                manifest, err = box.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            if err:
                self.log(self._t("lg_update_fail", e=err), quiet=silent)
                return
            latest = manifest["version"]
            if compare_versions(latest, VERSION) <= 0:
                self.log(self._t("lg_update_none", cur="v" + VERSION), quiet=silent)
                return
            self.log(self._t("lg_update_found", new=latest, cur="v" + VERSION))
            self._handle_update_found(parent, manifest, latest)

        self.root.after(100, poll)

    def _download_and_launch(self, parent, manifest, latest):
        """后台下载校验，完成后回主线程启动新版。"""
        box = queue.Queue()

        def worker():
            def progress(done, total):
                self.root.after(0, lambda: self.log(
                    self._t("lg_update_progress", done=done, total=total)))
            box.put(download_verified_exe(
                manifest,
                os.path.dirname(os.path.abspath(sys.executable
                                                if getattr(sys, "frozen", False) else __file__)),
                on_progress=progress))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                path, err = box.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            if err:
                self.log(self._t("lg_update_dl_fail", e=err))
                messagebox.showerror(
                    self._t("t_update_dl_fail"),
                    self._t("m_update_dl_fail", e=err), parent=parent)
                return
            self.log(self._t("lg_update_verified", path=path))
            ok, err = launch_verified_exe(path)
            if not ok:
                messagebox.showerror(
                    self._t("t_update_dl_fail"),
                    self._t("m_update_launch_fail", e=err, path=path), parent=parent)
                return
            messagebox.showinfo(
                self._t("t_update_done"),
                self._t("m_update_done", new=latest), parent=parent)
            # 新版已经起来，旧版自己退场：两个窗口并排摆着用户分不清哪个在跑，
            # 旧版留在后台还会继续读写同一份配置。
            self.on_closing()

        self.root.after(100, poll)


    def show_about(self):
        """关于窗：单页，外链走底部按钮跳转。"""
        dlg = tk.Toplevel(self.root)
        dlg.title(self._t("about_title"))
        dlg.transient(self.root)
        dlg.resizable(False, False)
        self._safe_grab(dlg)
        dlg.configure(bg=self.BG)

        btns = tk.Frame(dlg, bg=self.BG)
        btns.pack(fill="x", side="bottom", padx=14, pady=(4, 12))
        btn = self._make_outline_button(
            btns, self._t("about_release"), lambda: self._open_url(RELEASE_PAGE_URL))
        btn.pack(side="left", padx=3)
        self._reg_i18n(btn, "about_release")
        btn = self._make_outline_button(
            btns, self._t("check_update"), lambda: self.check_for_update(dlg))
        btn.pack(side="left", padx=3)
        self._reg_i18n(btn, "check_update")

        page = tk.Frame(dlg, bg=self.BG)
        page.pack(fill="both", expand=True, padx=16, pady=(6, 4))

        tk.Label(page, text=self._t("about_app", ver=VERSION), bg=self.BG,
                 fg=self.TXT, font=self._title_font()).pack(pady=(10, 0))
        tk.Label(page, text=self._t("about_years", start=START_YEAR,
                                    end=datetime.now().year,
                                    name=AUTHOR_NAME),
                 bg=self.BG, fg=self.TXT_DIM).pack(pady=(6, 0))
        tk.Label(page, text=self._t("contact", qq=CONTACT_QQ, email=CONTACT_EMAIL),
                 bg=self.BG, fg=self.TXT_DIM).pack(pady=(4, 0))
        tk.Label(page, text=self._t("about_free"), bg=self.BG,
                 fg=self.TXT).pack(pady=(10, 0))
        tk.Label(page, text=self._t("about_note"), bg=self.BG,
                 fg=self.TXT_DIM, wraplength=400, justify="center").pack(pady=(8, 0))
        tk.Label(page, text=self._t("about_no_warranty"), bg=self.BG,
                 fg=self.TXT_DIM, wraplength=400, justify="center").pack(pady=(8, 0))

        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dlg.focus_set()
        self._center_dialog(dlg)
        self._wait_dialog(dlg)


def _run_selfcheck():
    """打包完整性自检：`BatchReplace.exe --selfcheck`，秒级返回。

    冻结版里逐个 import 清单模块，漏带立刻报错退出——替代「把自检脚本
    再打一个 exe」的旧流程（那要花一次完整打包的时长）。再实际画一张
    PIL 位图：_imaging 缺失时 import PIL 不报错，一画才炸，必须真画。
    输出用英文：这里是控制台诊断信息，不进界面文案（TEXTS）体系。
    """
    frozen = bool(getattr(sys, "frozen", False))
    print("BatchReplace v%s selfcheck (frozen=%s)" % (VERSION, frozen))
    missing = []
    for name in SELFCHECK_IMPORTS:
        try:
            __import__(name)
        except Exception as exc:
            missing.append("  %s: %s" % (name, exc))
    try:
        from PIL import Image as _Img, ImageDraw as _Draw
        _probe = _Img.new("RGB", (8, 8))
        _Draw.Draw(_probe).rectangle((0, 0, 7, 7), fill=(255, 0, 0))
        _probe.load()
    except Exception as exc:
        missing.append("  PIL bitmap draw: %s" % exc)
    if missing:
        print("selfcheck FAILED (%d missing):" % len(missing))
        print("\n".join(missing))
        return 1
    print("selfcheck OK: %d modules importable, PIL draw works"
          % len(SELFCHECK_IMPORTS))
    return 0


def _run_selftest_script(script, log_path=None):
    """在交付 exe 自身进程里执行外部自检脚本（`--selftest`，完整冻结自检用）。

    旧做法是按同一份 spec 把自检脚本再打成 br_selftest.exe：每次多花一整轮
    打包，而且那个兄弟 exe 的模块图和交付的这个并不相同——「兄弟有、正品
    没有」的缺口它照不出来。改成在正品进程里执行，打包只做一次，验的就是
    要交出去的那个 exe。
    """
    stream = None
    if log_path:
        # 窗口版 exe 从资源管理器启动时没有 stdout，结果只能落文件
        try:
            stream = open(log_path, "w", encoding="utf-8", buffering=1)
        except OSError as exc:
            print("selftest log not writable (%r), reporting to stdout only" % (exc,))
        else:
            sys.stdout = sys.stderr = stream
    # 冻结版把入口脚本挂在 __main__ 名下，不会注册成 BatchReplace 模块；
    # 而自检脚本头一句就是 import BatchReplace。指过去复用同一个模块对象，
    # 否则会再加载出一份独立副本，两份全局状态各改各的。
    main_module = sys.modules.get("__main__")
    if main_module is not None:
        sys.modules.setdefault("BatchReplace", main_module)
    try:
        with open(script, "rb") as fh:
            code = compile(fh.read(), script, "exec")
        exec(code, {"__name__": "__main__", "__file__": script})
        return 0
    except SystemExit as exc:
        if isinstance(exc.code, int):
            return exc.code
        return 0 if exc.code is None else 1
    except Exception as exc:
        # 逐项检查的异常在自检脚本内部已各自记下，走到这里只剩脚本自身加载
        # 失败（语法 / 导入），异常自带位置与模块名。不引 traceback：正文
        # 不用它，冻结版包里也没有。
        print("selftest script failed: %r" % (exc,))
        return 1
    finally:
        if stream is not None:
            stream.close()
            sys.stdout, sys.stderr = sys.__stdout__, sys.__stderr__


def main():
    if "--selfcheck" in sys.argv:
        sys.exit(_run_selfcheck())
    if "--selftest" in sys.argv:
        argv = sys.argv
        at = argv.index("--selftest")
        script = argv[at + 1] if len(argv) > at + 1 else ""
        log_path = ""
        if "--log" in argv:
            at = argv.index("--log")
            log_path = argv[at + 1] if len(argv) > at + 1 else ""
        if not script or not os.path.isfile(script):
            print("--selftest needs an existing script path: %r" % script)
            sys.exit(2)
        sys.exit(_run_selftest_script(script, log_path or None))
    # 命令行更新入口：--update-manifest <url> [--update-dry-run]
    argv = sys.argv[1:]
    if "--update-manifest" in argv:
        idx = argv.index("--update-manifest")
        target = argv[idx + 1] if idx + 1 < len(argv) else ""
        if not target:
            print("FAIL: --update-manifest needs a URL")
            sys.exit(1)
        dry = "--update-dry-run" in argv
        print(run_one_click_update(target, dry_run=dry))
        sys.exit(0)

    # 进程身份要在建窗之前定下来：任务栏按进程归组，晚设的话第一次启动的
    # 按钮仍挂在 python.exe / 默认身份下
    _set_app_user_model_id()
    # 图标在建窗前先画好落盘：PIL 超采样画一张 1024 画布要上百毫秒，留到
    # 窗口显示之后画，任务栏按钮会先以空白图标出现、随后才刷新。
    _app_icon_path()
    root = tk.Tk()
    # 先隐藏：任务栏样式在窗口首次可见之前就位，Shell 不必事后重建按钮。
    root.withdraw()
    BatchReplaceApp(root, defer_show=True)
    _apply_taskbar_identity(root)
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
BatchReplace 核心逻辑的回归测试。

这些用例覆盖的正是历史上反复出问题的地方：
  - 数字类型被替换后必须仍是数字，不能变成字符串
  - 备份不得覆盖同名文件
  - 大小写敏感 / 不敏感两种模式的匹配与替换
  - 逻辑函数与界面同处一个文件，测试直接导入 BatchReplace 本体验证

注意：逻辑与界面同在 BatchReplace.py，导入它会连带导入 tkinter，
因此必须用带 tkinter 的解释器运行（见 pytest.ini 上方的说明）。

注释约定：注释只写「这条用例守什么」，不记录改动出在哪个版本——
版本沿革看 git 提交记录。

运行：pytest -q
"""

import ast
import ctypes
import json
import os
import queue
import re
import shutil
import sys
import threading
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import BatchReplace as core  # noqa: E402

pytest.importorskip("docx")
pytest.importorskip("openpyxl")

from docx import Document  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402


# ---------- 工具函数 ----------

def test_sanitize_filename_replaces_invalid_chars():
    assert core.sanitize_filename('a/b\\c:d*e?f"g<h>i|j') == "a_b_c_d_e_f_g_h_i_j"


def test_sanitize_filename_strips_whitespace():
    assert core.sanitize_filename("  name  ") == "name"


# ---------- Word ----------

def make_docx(path, text="Hello World"):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(path)
    return path


def test_extract_word_finds_text(tmp_path):
    p = make_docx(tmp_path / "a.docx", "Hello World")
    matches, skipped = core.extract_word_matches(str(p), "World", True)
    assert len(matches) == 1
    assert matches[0]["type"] == "paragraph"
    assert "World" in matches[0]["context"]
    assert skipped == []


def test_extract_word_case_insensitive_by_default(tmp_path):
    p = make_docx(tmp_path / "b.docx", "Hello World")
    matches, _ = core.extract_word_matches(str(p), "world", False)
    assert len(matches) == 1


def test_extract_word_case_sensitive_miss(tmp_path):
    p = make_docx(tmp_path / "c.docx", "Hello World")
    matches, _ = core.extract_word_matches(str(p), "world", True)
    assert matches == []


def test_extract_word_no_match(tmp_path):
    p = make_docx(tmp_path / "d.docx", "Hello World")
    matches, _ = core.extract_word_matches(str(p), "不存在", True)
    assert matches == []


def test_word_paragraph_numbers_align_with_word_collection(tmp_path):
    """顶层段落编号必须按 Word doc.Paragraphs 的口径算（期望值与真 Word 实测核对过）。

    Word 的集合含表格单元格内的段落，且每个表格行尾还有一个行标记幻影段；
    python-docx 的 doc.paragraphs 只数顶层段落，两者差着整个表格的段数。
    """
    p = tmp_path / "nums.docx"
    doc = Document()
    doc.add_paragraph("表前文字")
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "格零甲"
    t.cell(0, 1).text = "格零乙"
    t.cell(1, 0).text = "格一甲"
    t.cell(1, 1).text = "格一乙"
    doc.add_paragraph("表后文字")
    doc.save(p)

    d = Document(str(p))
    numbers = core._word_paragraph_numbers(d.element.body)
    assert numbers[d.paragraphs[0]._p] == 1, "表前段落编号"
    # 行0 两格（2、3）+ 行尾幻影（4）+ 行1 两格（5、6）+ 行尾幻影（7）
    assert numbers[d.paragraphs[1]._p] == 8, "表格后的顶层段落必须按 Word 口径编号"
    assert numbers[d.tables[0].rows[0].cells[0].paragraphs[0]._p] == 2
    assert numbers[d.tables[0].rows[1].cells[1].paragraphs[0]._p] == 6


def test_extract_word_paragraph_after_table_gets_word_number(tmp_path):
    """表格后的顶层段落：para_idx 必须是 Word 口径，否则双击跳转落进表格里。

    Word 的 doc.Paragraphs(n) 含表格内段落与行标记，顶层编号比它小——差多少
    取决于前面有几张表、每张表几行。含表格的文档必现，无表格文档不受影响。
    """
    p = tmp_path / "after.docx"
    doc = Document()
    doc.add_paragraph("表前文字")
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "格内文字"
    doc.add_paragraph("表后文字")
    doc.save(p)

    matches, _ = core.extract_word_matches(str(p), "表后文字", True)
    assert len(matches) == 1
    assert matches[0]["type"] == "paragraph"
    # 表前(1) + 格内(2) + 行标记幻影(3) 之后的 4
    assert matches[0]["para_idx"] == 4, "表格后段落的编号不是 Word 口径: %r" % matches[0]


def test_replace_word_updates_and_saves(tmp_path):
    p = make_docx(tmp_path / "e.docx", "Hello World")
    matches, _ = core.extract_word_matches(str(p), "World", True)
    skipped = core.replace_word_multi(str(p), _annotate(matches, "World", "YehOlt"), True)
    assert skipped == []
    assert Document(str(p)).paragraphs[0].runs[0].text == "Hello YehOlt"


def test_replace_word_case_insensitive(tmp_path):
    p = make_docx(tmp_path / "f.docx", "Hello World")
    matches, _ = core.extract_word_matches(str(p), "world", False)
    core.replace_word_multi(str(p), _annotate(matches, "world", "X"), False)
    assert Document(str(p)).paragraphs[0].runs[0].text == "Hello X"


# ---------- Excel ----------

def make_xlsx(path, cells):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for coord, val in cells.items():
        ws[coord] = val
    wb.save(path)
    return path


def test_extract_excel_finds_cell(tmp_path):
    p = make_xlsx(tmp_path / "a.xlsx", {"A1": "订单ABC编号", "B1": 123})
    matches, skipped = core.extract_excel_matches(str(p), "ABC", True)
    assert len(matches) == 1
    assert matches[0]["cell_coordinate"] == "A1"
    assert skipped == []


def test_replace_excel_text_cell(tmp_path):
    p = make_xlsx(tmp_path / "b.xlsx", {"A1": "订单ABC编号"})
    matches, _ = core.extract_excel_matches(str(p), "ABC", True)
    core.replace_excel_multi(str(p), _annotate(matches, "ABC", "XYZ"), True)
    assert load_workbook(str(p)).active["A1"].value == "订单XYZ编号"


def test_replace_excel_preserves_int_type(tmp_path):
    """核心回归点：数字被替换后必须还是数字，不能变成字符串。"""
    p = make_xlsx(tmp_path / "c.xlsx", {"A1": 12345})
    matches, _ = core.extract_excel_matches(str(p), "234", True)
    assert len(matches) == 1
    core.replace_excel_multi(str(p), _annotate(matches, "234", "999"), True)
    val = load_workbook(str(p)).active["A1"].value
    assert val == 19995
    assert isinstance(val, int)


def test_replace_excel_preserves_float_type(tmp_path):
    p = make_xlsx(tmp_path / "d.xlsx", {"A1": 1234.5})
    matches, _ = core.extract_excel_matches(str(p), "234", True)
    assert len(matches) == 1
    core.replace_excel_multi(str(p), _annotate(matches, "234", "999"), True)
    val = load_workbook(str(p)).active["A1"].value
    assert val == 1999.5
    assert isinstance(val, float)


def test_replace_excel_non_numeric_result_falls_back_to_text(tmp_path):
    """数字替换后不再是合法数字时，应降级为文本而不是抛异常。"""
    p = make_xlsx(tmp_path / "e.xlsx", {"A1": 100})
    matches, _ = core.extract_excel_matches(str(p), "0", True)
    core.replace_excel_multi(str(p), _annotate(matches, "0", "A"), True)
    assert load_workbook(str(p)).active["A1"].value == "1AA"


# ---------- 纯文本 (.txt) ----------

def make_txt(path, text, encoding="utf-8", bom=False):
    data = text.encode(encoding)
    if bom:
        data = b"\xef\xbb\xbf" + data
    path.write_bytes(data)
    return path


def test_extract_text_finds_occurrence(tmp_path):
    p = make_txt(tmp_path / "a.txt", "Hello World")
    matches, skipped = core.extract_text_matches(str(p), "World", True)
    assert len(matches) == 1
    assert matches[0]["type"] == "text"
    assert "World" in matches[0]["context"]
    assert skipped == []


def test_extract_text_case_insensitive(tmp_path):
    p = make_txt(tmp_path / "b.txt", "Hello World")
    matches, _ = core.extract_text_matches(str(p), "world", False)
    assert len(matches) == 1


def test_extract_text_no_match(tmp_path):
    p = make_txt(tmp_path / "c.txt", "Hello World")
    matches, _ = core.extract_text_matches(str(p), "不存在", True)
    assert matches == []


def test_extract_text_location_line_col(tmp_path):
    # 第 1 行: abc；第 2 行: def（命中起点在 d，即第 2 行第 1 列）
    p = make_txt(tmp_path / "d.txt", "abc\ndef")
    matches, _ = core.extract_text_matches(str(p), "def", True)
    assert len(matches) == 1
    assert matches[0]["location"] == "第 2 行, 第 1 列"


def test_read_text_auto_detects_gbk(tmp_path):
    p = make_txt(tmp_path / "i.txt", "中文测试", encoding="gbk")
    text, enc = core._read_text_auto(str(p))
    assert enc == "gbk"
    assert text == "中文测试"


def test_replace_text_utf8_roundtrip(tmp_path):
    p = make_txt(tmp_path / "e.txt", "订单ABC编号")
    matches, _ = core.extract_text_matches(str(p), "ABC", True)
    for m in matches:
        m["replace_text"] = "XYZ"
        m["cond_idx"] = 1
    skipped = core.replace_text_multi(str(p), matches, True)
    assert skipped == []
    assert p.read_text(encoding="utf-8") == "订单XYZ编号"


def test_replace_text_gbk_no_mojibake(tmp_path):
    """核心回归点：GBK 编码的 .txt 不能乱码，替换后仍以 GBK 写回。"""
    p = make_txt(tmp_path / "f.txt", "订单ABC编号", encoding="gbk")
    matches, _ = core.extract_text_matches(str(p), "ABC", True)
    for m in matches:
        m["replace_text"] = "XYZ"
        m["cond_idx"] = 1
    core.replace_text_multi(str(p), matches, True)
    assert p.read_bytes().decode("gbk") == "订单XYZ编号"


def test_replace_text_utf8_bom_preserved(tmp_path):
    """带 BOM 的 UTF-8 文件替换后 BOM 仍在。"""
    p = make_txt(tmp_path / "g.txt", "Hello World", bom=True)
    matches, _ = core.extract_text_matches(str(p), "World", True)
    for m in matches:
        m["replace_text"] = "YehOlt"
        m["cond_idx"] = 1
    core.replace_text_multi(str(p), matches, True)
    raw = p.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert raw.decode("utf-8-sig") == "Hello YehOlt"


def test_replace_text_multi_condition(tmp_path):
    """多条件：两类查找词共存，逐一替换、不链式覆盖。"""
    p = make_txt(tmp_path / "h.txt", "苹果 香蕉 苹果")
    m1, _ = core.extract_text_matches(str(p), "苹果", True)
    m2, _ = core.extract_text_matches(str(p), "香蕉", True)
    for m in m1:
        m["replace_text"] = "橙子"
        m["cond_idx"] = 1
    for m in m2:
        m["replace_text"] = "西瓜"
        m["cond_idx"] = 2
    core.replace_text_multi(str(p), m1 + m2, True)
    assert p.read_text(encoding="utf-8") == "橙子 西瓜 橙子"


def test_replace_text_multi_does_not_feed_result_to_next_condition(tmp_path):
    """上一条的替换结果不得成为下一条的输入——预览承诺并行，写盘就得并行。

    原文 'ab'：a→b 与 b→X 两命中不重叠、查找词也不重复，任何弹窗都不该拦；
    旧实现却按条件顺序在**替换后的文本**上重找，写出 'XX'（预览说的是 'bX'）。
    """
    p = make_txt(tmp_path / "chain.txt", "ab")
    m1, _ = core.extract_text_matches(str(p), "a", True)
    m2, _ = core.extract_text_matches(str(p), "b", True)
    for m in m1:
        m["replace_text"] = "b"
        m["cond_idx"] = 1
    for m in m2:
        m["replace_text"] = "X"
        m["cond_idx"] = 2
    skipped = core.replace_text_multi(str(p), m1 + m2, True)
    assert skipped == []
    assert p.read_text(encoding="utf-8") == "bX"


def test_replace_text_multi_skips_when_text_changed_since_search(tmp_path):
    """搜索之后文件被外部改动：区间对不上就跳过并报告，绝不盲写。"""
    p = make_txt(tmp_path / "stale.txt", "甲乙")
    matches, _ = core.extract_text_matches(str(p), "甲", True)
    for m in matches:
        m["replace_text"] = "丙"
        m["cond_idx"] = 1
    p.write_text("乙甲", encoding="utf-8")   # 区间已不是当初搜到的位置
    skipped = core.replace_text_multi(str(p), matches, True)
    assert skipped and "已变化" in skipped[0][1], "文本已变未被报告: %r" % (skipped,)
    assert p.read_text(encoding="utf-8") == "乙甲", " stale 区间不该被写"


def test_replace_text_multi_reports_overlap_instead_of_crashing(tmp_path):
    """绕过界面直接调引擎且命中重叠：报告冲突，不是 AttributeError。

    界面侧 detect_condition_conflicts 会把重叠整批拦下，这条路径只在绕过
    界面时出现——但"兜底路径自己先崩"等于没有兜底。rival 必须是一条已应用
    的命中字典，不是已应用区间的 start。
    """
    p = make_txt(tmp_path / "ov.txt", "abcd")
    m1, _ = core.extract_text_matches(str(p), "abc", True)
    m2, _ = core.extract_text_matches(str(p), "bcd", True)
    for m in m1:
        m["replace_text"] = "X"
        m["cond_idx"] = 1
    for m in m2:
        m["replace_text"] = "Y"
        m["cond_idx"] = 2
    skipped = core.replace_text_multi(str(p), m1 + m2, True)
    assert len(skipped) == 1, "重叠命中应恰有一条被拦下: %r" % (skipped,)
    loc, reason = skipped[0]
    assert "abc" in reason and "bcd" in reason, \
        "冲突原因未点出双方查找词: %r" % (reason,)
    # 先应用的那条照常写盘（从后往前，位置靠后的 bcd 先写）
    assert p.read_text(encoding="utf-8") == "aY"


# ---------- 备份 ----------

def test_build_backup_folder_name_content_mode():
    from datetime import datetime
    now = datetime(2026, 10, 10, 15, 4, 5)
    name = core.build_backup_folder_name("旧名", "新名", False, now)
    # 查找→替换 用箭头连接，方向一目了然
    assert name == "backup_20261010_旧名→新名_内容替换_150405"


def test_build_backup_folder_name_uses_arrow_and_keeps_time():
    """箭头方向必须正确（备份的是替换前的文件），且时间戳不能丢。"""
    from datetime import datetime
    now = datetime(2026, 10, 10, 16, 48, 27)
    name = core.build_backup_folder_name("哈哈", "银行", False, now)
    assert name == "backup_20261010_哈哈→银行_内容替换_164827"
    assert name.startswith("backup_")
    # 箭头左边是替换前（备份内容），右边是替换后
    assert name.split("→")[0].endswith("哈哈")
    assert "164827" in name


def test_build_backup_folder_name_filename_mode():
    from datetime import datetime
    now = datetime(2026, 10, 10, 15, 4, 5)
    name = core.build_backup_folder_name("a/b", "", True, now)
    assert name == "backup_20261010_a_b_文件名替换_150405"


def test_create_backup_does_not_overwrite(tmp_path):
    """回归点：同名文件必须并存，不能互相覆盖。"""
    src = tmp_path / "report.docx"
    src.write_text("v1", encoding="utf-8")
    backup_dir = tmp_path / "backup"

    first = core.create_backup(str(src), str(backup_dir))
    src.write_text("v2", encoding="utf-8")
    second = core.create_backup(str(src), str(backup_dir))

    assert first != second
    assert Path(first).read_text(encoding="utf-8") == "v1"
    assert Path(second).read_text(encoding="utf-8") == "v2"


def test_create_backup_folder_creates_named_dir(tmp_path):
    from datetime import datetime
    now = datetime(2026, 10, 10, 16, 48, 27)
    folder = core.create_backup_folder(str(tmp_path), "哈哈", "银行", False, now)
    assert Path(folder).is_dir()
    assert Path(folder).name == "backup_20261010_哈哈→银行_内容替换_164827"


def test_create_backup_folder_dedupes_existing(tmp_path):
    """同一秒重复执行时，第二个备份文件夹必须另起名字，不能复用/覆盖。"""
    from datetime import datetime
    now = datetime(2026, 10, 10, 16, 48, 27)
    first = core.create_backup_folder(str(tmp_path), "哈哈", "银行", False, now)
    second = core.create_backup_folder(str(tmp_path), "哈哈", "银行", False, now)
    assert first != second
    assert Path(first).is_dir() and Path(second).is_dir()


# ---------- 原子写 ----------

def test_atomic_save_replaces_file_on_success(tmp_path):
    target = tmp_path / "a.txt"
    target.write_text("old", encoding="utf-8")

    core.atomic_save(lambda p: Path(p).write_text("new", encoding="utf-8"), str(target))

    assert target.read_text(encoding="utf-8") == "new"
    assert not list(tmp_path.glob(".batchreplace_tmp_*"))


def test_atomic_save_keeps_original_when_save_fails(tmp_path):
    """保存中途失败时原文件必须原封不动，且不留临时文件。"""
    target = tmp_path / "a.txt"
    target.write_text("old", encoding="utf-8")

    def bad_save(p):
        Path(p).write_text("partial", encoding="utf-8")
        raise RuntimeError("模拟磁盘写满")

    with pytest.raises(RuntimeError):
        core.atomic_save(bad_save, str(target))

    assert target.read_text(encoding="utf-8") == "old"
    assert not list(tmp_path.glob(".batchreplace_tmp_*"))


# ---------- 防复发：逻辑只能有一份实现（单文件） ----------

UI_SOURCE = Path(__file__).resolve().parents[1] / "BatchReplace.py"

# 这些逻辑只允许在模块顶层定义一次，且必须与界面同处一个文件。
SHARED_LOGIC = [
    "sanitize_filename", "extract_word_matches", "extract_excel_matches",
    "atomic_save",
    "build_backup_folder_name", "create_backup_folder", "create_backup",
    "restore_cell_value",
    # 多条件替换的核心：界面与测试共用同一份实现
    "replace_word_multi", "replace_excel_multi", "apply_filename_conditions",
    "plan_span_replacements", "validate_conditions", "detect_condition_conflicts",
    "conditions_summary", "conflict_block_message", "normalize_history_item",
]


def test_no_split_module():
    """源码必须保持单文件：不得再拆出 batchreplace_core 之类的搭档模块。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert "import batchreplace_core" not in src
    assert "batchreplace_core" not in src, "源码里残留了对已删除模块的引用"


@pytest.mark.parametrize("name", SHARED_LOGIC)
def test_logic_defined_exactly_once_at_top_level(name):
    """
    历史上界面在类里手抄了一份与核心模块重复的实现，导致「测试全绿但 exe 仍带 bug」。

    这条用例守住边界：每个逻辑函数在模块顶层恰好定义一次。
    """
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert src.count(f"\ndef {name}(") == 1, \
        f"{name} 在源码里不是恰好定义一次（顶层 1 次），可能又出现了重复实现"


@pytest.mark.parametrize("name", SHARED_LOGIC)
def test_no_duplicate_inside_class(name):
    """类内不得再出现同名方法副本——那正是当初「两套代码」的形态。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert f"\n    def {name}(" not in src, \
        f"类内出现了重复实现 {name}，测试将不再能保护 exe"


def test_tested_module_is_the_shipped_one():
    """测试导入的模块必须就是打包交付的那份源码文件。"""
    assert Path(core.__file__).resolve() == UI_SOURCE.resolve()


# =====================================================================
# P0/P1 整改的护航测试
#
# 这些用例守的是「用户看不到但会出事」的地方：线程安全、失败时是否真的
# 中止、窗口什么时候关、规则是否完整还原、文件名里带特殊串还对不对。
# =====================================================================

import tkinter as tk  # noqa: E402
from tkinter import ttk  # noqa: E402


class _FakeWidget:
    """只有一个 destroy() 的假窗口，用来替代预览窗。"""

    def __init__(self):
        self.destroyed = False

    def destroy(self):
        self.destroyed = True


class _FakeRoot:
    """记录 after / after_cancel 的极简 root：证明「工作线程不碰 Tcl」。"""

    def __init__(self):
        self.after_calls = []
        self.cancelled = []
        self.bg = None

    def after(self, ms, fn=None, *args):
        self.after_calls.append((ms, fn))
        return f"job-{len(self.after_calls)}"

    def after_cancel(self, job):
        self.cancelled.append(job)

    def configure(self, **kw):
        # _apply_theme 会铺 root 底色；假 root 只记录，供断言
        self.bg = kw.get("bg", self.bg)


class _FakeThread:
    def __init__(self, *args, **kwargs):
        self.daemon = False

    def start(self):
        pass


class _FakeThreading:
    """替身 threading 模块：记录是否真的起了线程，但不真起。"""

    def __init__(self, sink):
        self.sink = sink

    def Thread(self, *args, **kwargs):
        self.sink.append(kwargs.get("target"))
        return _FakeThread()


@pytest.fixture(scope="function")
def tk_root():
    """真 Tk 根窗口（隐藏不弹窗）；无图形环境则整体跳过。

    每个测试单独一个 root：BatchReplaceApp 的控件直接挂在传入的 root 上，
    模块级共享会让后跑的测试摸到前一个 app 遗留的同文案控件（按文案找
    勾选框 / 按钮的查找函数都会撞上），断言的是上一个实例的状态。
    """
    try:
        root = tk.Tk()
    except Exception as e:  # pragma: no cover
        pytest.skip(f"无可用图形环境: {e}")
    root.withdraw()
    yield root
    try:
        root.destroy()
    except Exception:
        pass


class _FakeVar:
    """只有 get/set 的假 tk 变量：给不接 tk_root 的辅助实例用。"""

    def __init__(self, value=None):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


def _bare_app():
    """只挂 UI 通道所需属性的最小实例（不建界面、不起心跳、不读配置）。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app.root = _FakeRoot()
    app._ui_queue = queue.Queue()
    app._ui_pump_id = None
    app._log_lines = []
    app.log = app._log_lines.append
    app.theme = "light"
    app._apply_theme("light")
    # 状态栏要读的备份状态。用假变量而非 tk.BooleanVar：本辅助函数不接
    # tk_root，建真变量会要求解释器，把无关测试拖上 Tk 依赖。
    app.backup_var = _FakeVar(True)
    app.backup_path = _FakeVar(core.DEFAULT_BACKUP_DIR)
    return app


def _annotate(matches, find_text, replace_text, cond_idx=0):
    """给匹配补上替换词与条件序号：替换引擎按每条匹配自带的词替换。

    提取出的匹配不带这两个字段，直接调替换引擎或替换 worker 时必须由
    调用方补上。
    """
    out = []
    for m in matches:
        m2 = dict(m)
        m2.setdefault("find_text", find_text)
        m2["replace_text"] = replace_text
        m2.setdefault("cond_idx", cond_idx)
        out.append(m2)
    return out


def _make_replace_app(master, tmp_path, backup_dir=None):
    """替换流程所需的最小实例：tk 变量是真的，界面控件不建。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app.root = _FakeRoot()
    app._ui_queue = queue.Queue()
    app._ui_pump_id = None
    app._log_lines = []
    app.log = app._log_lines.append
    app.selected_files = []
    app.preview_data = []
    # 搜索重入守卫的基线状态（真实实例由 __init__ 设置）
    app._search_thread = None
    app._search_gen = 0
    app.theme = "light"
    app._apply_theme("light")
    # 替换互斥守卫的基线状态
    app._replace_thread = None
    app.backup_var = tk.BooleanVar(value=backup_dir is not None)
    app.backup_path = tk.StringVar(value=str(backup_dir) if backup_dir else "")
    app.case_sensitive_var = tk.BooleanVar(value=False)
    app.replace_filename_var = tk.BooleanVar(value=False)
    app.multi_condition_var = tk.BooleanVar(value=False)
    app.conditions = []
    app._set_primary_state = lambda state: None
    app.find_entry = tk.Entry(master)
    app.replace_entry = tk.Entry(master)
    return app


def _content_preview_data(path, text="Hello World"):
    # 匹配带段落级字符区间（start/end），替换按区间映射回 w:t
    match = {"type": "paragraph", "para_idx": 1, "run_idx": 1,
             "start": 0, "end": len(text),
             "location": "第1段", "context": text, "text": text}
    return [{
        "file_path": str(path), "file_name": Path(path).name,
        "matches": [match],
        "match_data": [{"match": match, "checked": tk.BooleanVar(value=True)}],
        "checked": tk.BooleanVar(value=True),
    }]


# ---------- P0-1：工作线程与 UI 之间只能走队列 ----------

def test_ui_channel_marshals_calls_to_main_thread_pump():
    app = _bare_app()
    out = []
    app._ui(out.append, "a")
    app._ui(out.extend, ["b", "c"])
    assert out == [], "投递时不应立刻执行——那等于还在工作线程里碰 UI"
    app._pump_ui_queue()
    assert out == ["a", "b", "c"]


def test_ui_channel_never_touches_tcl_from_worker_thread():
    """P0-1 的核心：子线程投递 UI 任务时，一次都不许碰 Tcl。

    注意 root.after 本身也是 Tcl 调用，所以在工作线程里调它同样不安全；
    真正安全的只有「纯 Python 队列 + 主线程排空」。
    """
    app = _bare_app()
    out = []
    t = threading.Thread(target=lambda: [app._ui(out.append, i) for i in range(50)])
    t.start()
    t.join()

    assert out == [], "子线程不得直接执行 UI 操作"
    assert app.root.after_calls == [], "子线程不得调用 after（那仍是 Tcl 调用）"
    app._pump_ui_queue()
    assert out == list(range(50)), "主线程排空后应全部执行"


def test_ui_pump_reschedules_and_survives_bad_callback():
    app = _bare_app()
    done = []

    def bad():
        raise RuntimeError("模拟某个回调出错")

    app._ui(bad)
    app._ui(done.append, "ok")
    app._pump_ui_queue()
    assert done == ["ok"], "单个回调出错不得打断后面的任务"
    assert app.root.after_calls, "排空后必须重新挂上下一次心跳"


def test_worker_replace_touches_ui_only_through_channel():
    """静态守门：工作线程函数体内不得直接写日志 / 改按钮 / 碰 root。

    用 AST 而不是文本匹配——否则 docstring 和注释里的字样会误报。
    """
    import ast
    tree = ast.parse(UI_SOURCE.read_text(encoding="utf-8"))
    fn = next((n for n in ast.walk(tree)
               if isinstance(n, ast.FunctionDef) and n.name == "replace_files_from_preview"),
              None)
    assert fn is not None, "找不到 replace_files_from_preview"

    ui_names = {"log", "_set_primary_state", "update_status", "_mark_files_replaced"}
    bad = []
    for node in ast.walk(fn):
        # self.xxx(...) 直接调用
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and isinstance(node.func.value, ast.Name)
                and node.func.value.id == "self" and node.func.attr in ui_names):
            bad.append((node.lineno, ast.unparse(node)))
        # self.root.xxx
        if (isinstance(node, ast.Attribute) and isinstance(node.value, ast.Attribute)
                and isinstance(node.value.value, ast.Name)
                and node.value.value.id == "self" and node.value.attr == "root"):
            bad.append((node.lineno, ast.unparse(node)))
    assert bad == [], "工作线程里仍有直接触碰 UI 的代码: %r" % bad


# ---------- P0-2 / P0-3：先校验、后关窗；备份失败必须中止 ----------

def test_empty_selection_keeps_preview_open(tk_root, tmp_path, monkeypatch):
    """什么都没勾就点替换：预览窗必须还在，用户不至于丢掉全部上下文。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.preview_data = []
    warned = []
    monkeypatch.setattr(core.messagebox, "showwarning", lambda *a, **k: warned.append(a))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["a", "b"]])

    assert warned, "应弹出警告"
    assert not win.destroyed, "P0-2：空勾选时预览窗必须保持打开"


def test_aborts_without_touching_files_when_backup_unavailable(tk_root, tmp_path, monkeypatch):
    """P0-3：备份建不出来就中止，不做「无备份替换」。"""
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    blocker = tmp_path / "blocker"
    blocker.write_text("x", encoding="utf-8")  # 用文件占住路径，makedirs 必失败

    app = _make_replace_app(tk_root, tmp_path, backup_dir=blocker)
    app.preview_data = _content_preview_data(src)

    errors, started = [], []
    monkeypatch.setattr(core.messagebox, "showerror", lambda *a, **k: errors.append(a))
    monkeypatch.setattr(core, "threading", _FakeThreading(started))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["World", "YehOlt"]])

    assert errors, "应弹出错误框说明备份失败"
    assert started == [], "备份失败时不得进入替换流程"
    assert not win.destroyed, "中止时预览窗应保持打开"
    assert Document(str(src)).paragraphs[0].runs[0].text == "Hello World", "原文件必须零改动"


def test_aborts_when_backup_enabled_but_path_empty(tk_root, tmp_path, monkeypatch):
    """备份开着却没有路径：中止且文件零改动。

    曾经的缺陷是静默跳过备份直接改文件，同时界面还显示「备份开启」——
    用户以为有退路其实没有。做不到备份就不该动文件。
    """
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    app = _make_replace_app(tk_root, tmp_path, backup_dir=tmp_path / "bk")
    app.backup_path.set("")            # 用户把路径清空了
    app.preview_data = _content_preview_data(src)

    errors, started = [], []
    monkeypatch.setattr(core.messagebox, "showerror", lambda *a, **k: errors.append(a))
    monkeypatch.setattr(core, "threading", _FakeThreading(started))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["World", "YehOlt"]])

    assert errors, "应弹出错误框说明备份路径为空"
    assert started == [], "没有备份路径时不得进入替换流程"
    assert not win.destroyed, "中止时预览窗应保持打开"
    assert Document(str(src)).paragraphs[0].runs[0].text == "Hello World", "原文件必须零改动"


def test_status_bar_reports_missing_backup_path(tk_root, monkeypatch):
    """备份开着但路径为空时，状态栏不能说「备份开启」。"""
    app = _make_replace_app(tk_root, tmp_path=None, backup_dir=None)
    app.backup_var.set(True)           # 开着
    app.backup_path.set("")            # 但没有路径
    app.status_label = tk.Label(tk_root)
    app.selected_files = []
    app._total_matches = 0
    app.update_status()
    assert core.TEXTS["backup_no_path"]["zh"] in app.status_label.cget("text")
    assert app.status_label.cget("fg") == app.DANGER_FG, "空路径要用告警色"

    app.backup_path.set(str(tk_root.winfo_toplevel()))   # 填回路径
    app.update_status()
    assert core.TEXTS["backup_on"]["zh"] in app.status_label.cget("text")
    assert app.status_label.cget("fg") == app.TXT



def test_preview_closed_only_after_confirmation(tk_root, tmp_path, monkeypatch):
    """P0-2：用户在确认框点「否」时，预览窗仍在。"""
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    app = _make_replace_app(tk_root, tmp_path)  # 备份关闭 → 跳过备份分支
    app.preview_data = _content_preview_data(src)

    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: False)

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["World", "YehOlt"]])

    assert not win.destroyed, "点「否」后预览窗必须保持打开"
    assert Document(str(src)).paragraphs[0].runs[0].text == "Hello World"


# ---------- P0-5：规则保存与加载必须完整往返 ----------

def test_save_then_load_rule_roundtrip(tk_root, tmp_path, monkeypatch):
    rules_dir = tmp_path / "rules"
    rules_dir.mkdir()
    monkeypatch.setattr(core, "RULES_DIR", str(rules_dir))
    monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: None)

    app = _make_replace_app(tk_root, tmp_path)
    app.find_entry.insert(0, "哈哈")
    app.replace_entry.insert(0, "银行")
    app.case_sensitive_var.set(True)
    app.replace_filename_var.set(True)
    app.backup_var.set(False)
    app.save_rule()

    saved = list(rules_dir.glob("*.json"))
    assert len(saved) == 1, "规则文件未写出"
    payload = json.loads(saved[0].read_text(encoding="utf-8"))
    assert payload["replace_filename"] is True

    # 改乱状态后加载，看是否逐项还原
    app.find_entry.delete(0, tk.END)
    app.replace_entry.delete(0, tk.END)
    app.case_sensitive_var.set(False)
    app.replace_filename_var.set(False)
    monkeypatch.setattr(core.filedialog, "askopenfilename", lambda **k: str(saved[0]))
    app.load_rule()

    assert app.find_entry.get() == "哈哈"
    assert app.replace_entry.get() == "银行"
    assert app.case_sensitive_var.get() is True
    assert app.replace_filename_var.get() is True, "P0-5：replace_filename 没有被加载回来"
    assert app.backup_var.get() is False


# ---------- P1-6：新文件名不靠解析显示串 ----------

def test_search_worker_sets_new_name_field(tmp_path):
    p = tmp_path / "abc.docx"
    p.write_bytes(b"x")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    app._search_worker([str(p)], [["abc", "xyz"]], True, True)

    assert captured, "工作线程没有把结果投递回主线程"
    # worker 还会投递搜索进度，收尾调用按身份找
    finish = [c for c in captured if c[0] == app._finish_search]
    assert finish, "收尾结果未按 UI 通道投递"
    raw = finish[0][1][0]
    assert raw[0]["matches"][0]["new_name"] == "xyz.docx"
    # 显示串只负责显示
    assert raw[0]["matches"][0]["context"] == "abc.docx -> xyz.docx"


def test_filename_replace_reads_field_not_display_string(tk_root, tmp_path, monkeypatch):
    """P1-6：执行重命名时必须读 new_name 字段，而不是回头解析显示串。

    说明：Windows 的文件名不允许含 ">"，所以「文件名里带 ->」在本地盘上
    其实造不出来；但"替换为"这一侧可以是任何文本，一旦显示串里多出一个
    " -> "，"取第 2 段"的老写法就会切错。这里用「字段与显示串故意不一致」
    的样本来证明代码走的是字段。
    """
    src = tmp_path / "abc.docx"
    src.write_bytes(b"x")
    app = _make_replace_app(tk_root, tmp_path)
    app.replace_filename_var.set(True)
    app.preview_data = [{
        "file_path": str(src), "file_name": "abc.docx",
        "matches": [{"type": "filename", "location": "文件名",
                     "new_name": "renamed.docx",           # 字段：最终答案
                     "context": "abc.docx -> 陷阱 -> 串.docx"}],  # 显示串：切出来是错的
        "checked": tk.BooleanVar(value=True),
    }]

    # 文件名替换也有二次确认，这里直接放行
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: True)

    app.execute_replace_from_preview(_FakeWidget(), [["abc", "renamed"]])

    assert (tmp_path / "renamed.docx").exists(), "应严格按 new_name 重命名"
    assert not src.exists()


# ---------- P1-7：iid 反查统一走索引 ----------

def test_iid_index_replaces_string_parsing(tk_root):
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app.root = tk_root
    app.theme = "light"
    app._apply_theme("light")     # 勾选图标要从色板取色
    app.is_filename_mode = False
    app._iid_index = {}
    app.preview_data = [
        {"file_path": "f1", "file_name": "f1.docx",
         "matches": [{"location": "第1段", "context": "x1"},
                     {"location": "第2段", "context": "x2"}],
         "checked": tk.BooleanVar(value=True)},
        {"file_path": "f2", "file_name": "f2.docx",
         "matches": [{"location": "第1段", "context": "y1"}],
         "checked": tk.BooleanVar(value=True)},
    ]
    tree = ttk.Treeview(tk_root)
    try:
        app._init_match_data()
        app._populate_preview_tree(tree)
        assert app._iid_pos("file_0") == (0, None)
        assert app._iid_pos("file_1") == (1, None)
        assert app._iid_pos("file_1_match_0") == (1, 0)
        assert app._iid_pos("file_9_match_9") is None
        assert app._iid_pos("") is None
        assert app._iid_pos(None) is None
    finally:
        tree.destroy()


def test_no_iid_string_parsing_anywhere():
    """反查处不得再出现基于 iid 字符串协议的 split 解析（注释里的说明不算）。"""
    import ast
    tree = ast.parse(UI_SOURCE.read_text(encoding="utf-8"))
    bad = []
    for node in ast.walk(tree):
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and node.func.attr == "split" and node.args
                and isinstance(node.args[0], ast.Constant)
                and node.args[0].value in ("_match_", "_")):
            bad.append((node.lineno, ast.unparse(node)))
    assert bad == [], "仍有用字符串协议反推 iid 的代码: %r" % bad


# ---------- P2-14：Excel 类型保留 ----------

def test_restore_cell_value_prefers_bool_before_int():
    """bool 必须排在 int 前面判断——isinstance(True, int) 是真。"""
    assert core.restore_cell_value(True, "False") is False
    assert core.restore_cell_value(True, "TRUE") is True
    assert core.restore_cell_value(True, "不是布尔") == "不是布尔"
    assert core.restore_cell_value(123, "456") == 456
    assert core.restore_cell_value(1.5, "abc") == "abc"
    assert core.restore_cell_value("x", "y") == "y"


def test_replace_excel_preserves_bool(tmp_path):
    p = make_xlsx(tmp_path / "b.xlsx", {"A1": True})
    matches, _ = core.extract_excel_matches(str(p), "Tru", True)
    assert len(matches) == 1
    core.replace_excel_multi(str(p), _annotate(matches, "Tru", "Fals"), True)
    val = load_workbook(str(p)).active["A1"].value
    assert val is False, "布尔单元格被替换后退化成了字符串: %r" % (val,)


def test_replace_excel_preserves_datetime(tmp_path):
    from datetime import datetime as dt
    p = make_xlsx(tmp_path / "dt.xlsx", {"A1": dt(2024, 1, 1, 0, 0, 0)})
    matches, _ = core.extract_excel_matches(str(p), "2024", True)
    assert len(matches) == 1
    core.replace_excel_multi(str(p), _annotate(matches, "2024", "2025"), True)
    val = load_workbook(str(p)).active["A1"].value
    assert val == dt(2025, 1, 1, 0, 0, 0), "日期单元格被替换后退化成了字符串: %r" % (val,)


# =====================================================================
# R1-R6 六项修复的护航测试
#
# 守的是「同一个入口有两个触发源」这类结构性漏洞：搜索入口必须唯一且能自我
# 封锁、文件名与内容替换都要过确认、完成框要能找到父窗口。
# =====================================================================

class _AliveThread:
    """is_alive() 恒为真的假线程：模拟"搜索正在跑"。"""

    def is_alive(self):
        return True


def _filename_preview_data(path, new_name="xyz.docx"):
    return [{
        "file_path": str(path), "file_name": Path(path).name,
        "matches": [{"type": "filename", "location": "文件名",
                     "new_name": new_name,
                     "context": f"{Path(path).name} -> {new_name}"}],
        "checked": tk.BooleanVar(value=True),
    }]


def _make_files_tree(master):
    """与主界面同构的文件列表 Treeview——列定义一致，tree.set 才能按列名写值。"""
    tree = ttk.Treeview(master, columns=("type", "size", "match", "status"),
                        show="tree headings")
    tree.heading("#0", text="名称")
    for col in ("type", "size", "match", "status"):
        tree.heading(col, text=col)
    return tree


# ---------- R1：搜索重入守卫 ----------

def test_start_search_refuses_reentry_while_worker_alive(tk_root, tmp_path, monkeypatch):
    """R1：worker 还活着时，再次触发「开始搜索」必须被拒绝。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.selected_files = [str(tmp_path / "a.docx")]
    app._search_thread = _AliveThread()

    infos = []
    monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: infos.append(a))

    app.start_search()

    assert infos, "运行中再次触发搜索应被拒绝并提示"
    assert getattr(app, "_search_gen", 0) == 0, "被拒绝时不得开启新一轮搜索"


def test_start_search_disables_primary_entry(tk_root, tmp_path, monkeypatch):
    """R1：搜索启动后，唯一的搜索入口（主按钮）必须立即置灰。

    菜单栏已在 v2.0.0 双语版删除，不存在「按钮灰着、菜单却能点」的第二入口，
    故断言对象从菜单项改为 _set_primary_state（主按钮）。搜索代号与 worker
    线程仍照旧要推进——置灰只是表面，重入守卫才是实质。
    """
    app = _make_replace_app(tk_root, tmp_path)
    app.selected_files = [str(tmp_path / "a.docx")]
    app.find_entry.insert(0, "x")
    states = []
    app._set_primary_state = states.append
    started = []
    monkeypatch.setattr(core, "threading", _FakeThreading(started))

    app.start_search()

    assert states == ["disabled"], "搜索启动后主按钮未置灰: %r" % states
    assert app._search_gen == 1, "搜索代号未递增"
    assert started, "worker 线程未创建"


def test_finish_search_discards_stale_generation():
    """R1：代号对不上的旧结果必须整包丢弃，不得刷界面、不得改按钮状态。"""
    app = _bare_app()
    app._search_gen = 5
    app.preview_data = ["哨兵——不该被清空"]
    states = []
    app._set_primary_state = states.append

    app._finish_search([{"file_path": "f"}], {"f": 1}, 1, [], [["a", "b"]], False, gen=4)

    assert app.preview_data == ["哨兵——不该被清空"], "过期结果不得写进 preview_data"
    assert states == [], "过期结果不得改按钮状态"
    assert any("过期" in line for line in app._log_lines), "应留一行日志说明丢弃"


# ---------- R2：_finish_search 的 try/finally ----------

def test_finish_search_restores_ui_even_on_error():
    """R2：刷新中途抛异常，按钮也必须恢复可操作。"""
    app = _bare_app()
    app._search_gen = 1
    app.preview_data = []
    states = []
    app._set_primary_state = states.append
    app.update_status = lambda: None

    def boom(counts):
        raise RuntimeError("模拟刷列表时炸了")

    app._apply_search_status = boom

    with pytest.raises(RuntimeError):
        app._finish_search([], {}, 1, [], [["a", "b"]], False, gen=1)

    assert states == ["normal"], "异常路径下按钮未恢复（会被永久卡在禁用态）"


# ---------- R3：文件名替换后主列表整行刷新 ----------

def test_filename_replace_refreshes_main_list_row(tk_root, tmp_path, monkeypatch):
    """R3：重命名成功后，主列表要换成新文件名/新状态，selected_files 同步。"""
    src = tmp_path / "abc.docx"
    src.write_bytes(b"x")
    app = _make_replace_app(tk_root, tmp_path)
    app.replace_filename_var.set(True)
    app.files_tree = _make_files_tree(tk_root)
    app._insert_file_row(str(src))
    app.selected_files = [str(src)]
    app.preview_data = _filename_preview_data(src)

    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: True)

    app.execute_replace_from_preview(_FakeWidget(), [["abc", "xyz"]])

    new_path = str(tmp_path / "xyz.docx")
    assert (tmp_path / "xyz.docx").exists(), "文件应被重命名"
    assert app.files_tree.exists(new_path), "R3：主列表应换成新文件名（新 iid）"
    assert not app.files_tree.exists(str(src)), "旧行应删除"
    assert app.files_tree.set(new_path, "status") == "已替换", "状态应刷成已替换"
    assert app.selected_files == [new_path], "selected_files 应跟随新路径"


# ---------- R4：文件行不得进入跳转 ----------

def test_jump_ignores_file_row_iid():
    """R4：文件行（match_idx=None）调跳转应直接返回，不走降级打开文件。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app._iid_index = {"file_0": (0, None)}
    app.preview_data = [{"file_path": "f.docx", "file_name": "f.docx",
                         "matches": [], "match_data": []}]
    jumped = []
    app.jump_to_word_location = lambda *a: jumped.append(("word", a))
    app.jump_to_excel_location = lambda *a: jumped.append(("excel", a))
    app.log = lambda m: None

    app.open_and_jump_to_match("file_0")

    assert jumped == [], "R4：文件行不该触发任何跳转 / 降级打开"


# ---------- R5：纯文件名替换也要二次确认 ----------

def test_filename_replace_requires_confirmation(tk_root, tmp_path, monkeypatch):
    """R5：纯文件名替换路径必须有二次确认框，点「否」不得改名、不得关预览窗。"""
    src = tmp_path / "abc.docx"
    src.write_bytes(b"x")
    app = _make_replace_app(tk_root, tmp_path)
    app.replace_filename_var.set(True)
    app.preview_data = _filename_preview_data(src)

    asks = []
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: asks.append(a) or False)

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["abc", "xyz"]])

    assert asks, "R5：纯文件名替换也必须弹确认框"
    assert src.exists(), "点「否」后不得重命名"
    assert not win.destroyed, "点「否」后预览窗应保持打开"


# ---------- R6：完成框挂 parent ----------

def test_completion_dialog_has_parent_and_goes_through_channel():
    """R6：完成框必须经 UI 通道投递，且带上 parent（主窗最小化时才能到前台）。"""
    import ast
    tree = ast.parse(UI_SOURCE.read_text(encoding="utf-8"))

    worker = next((n for n in ast.walk(tree)
                   if isinstance(n, ast.FunctionDef)
                   and n.name == "replace_files_from_preview"), None)
    assert worker is not None, "找不到 replace_files_from_preview"
    direct = [n.lineno for n in ast.walk(worker)
              if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
              and isinstance(n.func.value, ast.Name)
              and n.func.value.id == "messagebox"]
    assert direct == [], "工作线程里不得直接弹框（必须经 UI 通道投递）"

    dialog = next((n for n in ast.walk(tree)
                   if isinstance(n, ast.FunctionDef)
                   and n.name == "_show_replace_done"), None)
    assert dialog is not None, "完成框应收敛到独立方法 _show_replace_done"
    assert "parent=" in ast.unparse(dialog), "完成框必须传 parent"


# =====================================================================
# P0 数据安全 + P1 功能缺口的护航测试
#
# 守的是两类问题：
#   数据安全——公式格被静默改写、xlsx 内嵌对象丢失、run 内分页符/图片
#              被抹掉、历史对话框删除即崩；
#   功能缺口——跨 run / 超链接 / 修订文本漏搜、同 run 多命中只记一笔、
#              替换无互斥、新文件名不预检、页眉重复枚举、长任务无进度。
# =====================================================================

from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402


def _docx_with_hyperlinked_run(path, link_text="链接里的目标文本"):
    """段落里追加一个超链接，目标文本在超链接内部的 run 中。"""
    doc = Document()
    p = doc.add_paragraph()
    r_id = p.part.relate_to("https://example.com", RT.HYPERLINK, is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = link_text
    r.append(t)
    hl.append(r)
    p._p.append(hl)
    doc.save(path)
    return path


def _docx_with_revision_run(path, rev_text="修订插入的目标文本"):
    """段落里追加一段修订插入（w:ins），目标文本在修订内部的 run 中。"""
    doc = Document()
    p = doc.add_paragraph()
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), "100")
    ins.set(qn("w:author"), "tester")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = rev_text
    r.append(t)
    ins.append(r)
    p._p.append(ins)
    doc.save(path)
    return path


def _docx_with_pagebreak_in_run(path):
    """一个 run 内同时含文本与分页符：旧引擎 run.text= 会把 br 一起抹掉。"""
    doc = Document()
    p1 = doc.add_paragraph()
    r = p1.add_run("前半段")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r._r.append(br)
    r.add_text("后半段")
    doc.add_paragraph("这是干净段落里的目标文本")
    doc.save(path)
    return path


def _docx_with_split_runs(path, parts):
    """一个段落里文本被拆进多个 run（Word 里极常见：中英文/加粗边界）。"""
    doc = Document()
    p = doc.add_paragraph()
    for part in parts:
        p.add_run(part)
    doc.save(path)
    return path


def _docx_with_multi_section_headers(path):
    """三节结构：第 1 节有自己的页眉，第 2 节也有自己的，第 3 节链接继承。"""
    doc = Document()
    doc.sections[0].header.paragraphs[0].add_run("甲处页眉文字")
    doc.add_paragraph("正文")
    doc.add_section()
    doc.sections[1].header.is_linked_to_previous = False  #  setter 会建独立定义
    doc.sections[1].header.paragraphs[0].add_run("乙处页眉文字")
    doc.add_paragraph("正文二")
    doc.add_section()
    doc.add_paragraph("正文三")
    doc.save(path)
    return path


def _make_xlsx_with_media(path):
    """造一个含 xl/media 的 xlsx（只加目录条目，检测函数只看名字）。"""
    make_xlsx(path, {"A1": "目标文本"})
    import zipfile
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/media/image1.png", b"\x89PNG\r\n\x1a\n dummy")
    os.replace(tmp, path)
    return path


def _cell_preview_data(path, location="Sheet1!A1"):
    match = {"type": "cell", "sheet_name": "Sheet1", "cell_coordinate": "A1",
             "location": location, "context": "目标文本", "value": "目标文本"}
    return [{
        "file_path": str(path), "file_name": Path(path).name,
        "matches": [match],
        "match_data": [{"match": match, "checked": tk.BooleanVar(value=True)}],
        "checked": tk.BooleanVar(value=True),
    }]


def _para_texts(path):
    """读取 docx 全部段落的纯文本（按文档顺序拼 w:t）。"""
    doc = Document(str(path))
    return [p.text for p in doc.paragraphs]


# ---------- P0-1：公式格跳过 ----------

def test_extract_excel_skips_formula_cells(tmp_path):
    """P0-1：公式格一律不进匹配——公式串里含查找词时旧实现会改写公式。"""
    p = make_xlsx(tmp_path / "f.xlsx", {"A1": "目标文本", "A2": '=CONCATENATE("x","目标文本")'})
    matches, skipped = core.extract_excel_matches(str(p), "目标文本", False)
    coords = [m["cell_coordinate"] for m in matches]
    assert "A1" in coords, "普通单元格应正常命中"
    assert "A2" not in coords, "P0-1：公式格不得进入匹配"
    assert any("公式" in s for s in skipped), "应留下公式格跳过的说明"


def test_replace_excel_refuses_formula_match(tmp_path):
    """P0-1 双保险：万一带公式的匹配项混进替换，也绝不改写公式。"""
    p = make_xlsx(tmp_path / "f2.xlsx", {"A1": '=CONCATENATE("x","目标文本")'})
    formula_match = [{"type": "cell", "sheet_name": "Sheet1",
                      "cell_coordinate": "A1", "location": "Sheet1!A1",
                      "context": "目标文本", "value": '=CONCATENATE("x","目标文本")'}]
    skipped = core.replace_excel_multi(
        str(p), _annotate(formula_match, "目标文本", "改掉了"), False)
    assert skipped, "公式匹配项应被跳过并报告"
    val = load_workbook(str(p)).active["A1"].value
    assert val == '=CONCATENATE("x","目标文本")', "公式被改写了: %r" % val


# ---------- P0-2：xlsx 内嵌对象预警 ----------

def test_xlsx_embedded_objects_detects_media(tmp_path):
    """P0-2：保存前查 zip，含 xl/media 必须被发现。"""
    p = _make_xlsx_with_media(tmp_path / "img.xlsx")
    kinds = core.xlsx_embedded_objects(str(p))
    assert "内嵌图片" in kinds, "未检出 xl/media: %r" % kinds
    # 这是给用户看的风险提示，要跟着界面语言走
    core.CURRENT_LANG = "en"
    assert "embedded image" in core.xlsx_embedded_objects(str(p))
    core.CURRENT_LANG = "zh"
    # 普通 xlsx 不应误报
    plain = make_xlsx(tmp_path / "plain.xlsx", {"A1": "x"})
    assert core.xlsx_embedded_objects(str(plain)) == []


def test_execute_replace_warns_and_cancels_on_embedded_objects(tk_root, tmp_path, monkeypatch):
    """P0-2：目标 xlsx 含内嵌对象时，弹窗告知并可取消；点否必须中止且不动文件。"""
    p = _make_xlsx_with_media(tmp_path / "img2.xlsx")
    app = _make_replace_app(tk_root, tmp_path)  # 备份关闭
    app.preview_data = _cell_preview_data(p)

    asks = []
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: asks.append(a) or False)
    started = []
    monkeypatch.setattr(core, "threading", _FakeThreading(started))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["目标文本", "替换后"]])

    assert asks, "含内嵌对象时必须弹窗告知"
    assert started == [], "用户取消后不得进入替换流程"
    assert not win.destroyed, "取消时预览窗应保持打开"
    assert load_workbook(str(p)).active["A1"].value == "目标文本", "原文件被改动了"


# ---------- P0-3：run 内分页符/图片跳过 ----------

def test_protected_run_match_is_skipped_with_note(tmp_path):
    """P0-3：命中 run 内含分页符时，该命中不进匹配，并留下说明。"""
    p = _docx_with_pagebreak_in_run(tmp_path / "pb.docx")
    matches, skipped = core.extract_word_matches(str(p), "前半段", False)
    assert matches == [], "含分页符的 run 不应进入匹配"
    assert any("分页符" in s or "run" in s for s in skipped), "应留下跳过说明: %r" % skipped


def test_replace_word_preserves_pagebreak_of_untouched_run(tmp_path):
    """P0-3：同文档里干净段落的替换照常完成，含分页符 run 的分页符必须还在。"""
    p = _docx_with_pagebreak_in_run(tmp_path / "pb2.docx")
    matches, _ = core.extract_word_matches(str(p), "干净段落", False)
    assert len(matches) == 1, "干净段落应正常命中"
    core.replace_word_multi(str(p), _annotate(matches, "干净段落", "已替换段落"), False)

    doc = Document(str(p))
    # 分页符还在
    brs = doc.paragraphs[0]._p.findall(".//" + qn("w:br"))
    assert brs, "P0-3：无辜 run 里的分页符被抹掉了"
    assert "已替换段落" in doc.paragraphs[1].text


# ---------- P0-4：历史对话框稳定键 ----------

def test_history_row_keys_unique_and_stable():
    """P0-4：行键用时间戳，同秒多条自动消歧；同一数据多次生成结果一致。"""
    history = [
        {"find": "a", "replace": "b", "timestamp": "2026-10-13 10:00:00"},
        {"find": "c", "replace": "d", "timestamp": "2026-10-13 10:00:00"},
        {"find": "e", "replace": "f", "timestamp": "2026-10-13 10:00:01"},
    ]
    keys = core.history_row_keys(history)
    assert len(set(keys)) == 3, "同秒记录的行键冲突: %r" % keys
    assert keys == core.history_row_keys(history), "行键必须稳定"


def test_history_delete_keeps_rows_aligned_with_data():
    """P0-4：删除一条后重建，每一行仍能取到正确记录（旧实现点最后一行即 IndexError）。"""
    history = [
        {"find": "a", "replace": "b", "timestamp": "2026-10-13 10:00:00"},
        {"find": "c", "replace": "d", "timestamp": "2026-10-13 10:00:01"},
        {"find": "e", "replace": "f", "timestamp": "2026-10-13 10:00:02"},
    ]
    # 旧行为复形：iid=下标，删除中间条后 history 变短，旧 iid "2" 越界
    old_iids = [str(i) for i in range(len(history))]
    del history[1]
    with pytest.raises(IndexError):
        _ = history[int(old_iids[-1])]

    # 新行为：稳定键 + 重建，行与数据永远对齐
    keys = core.history_row_keys(history)
    mapping = dict(zip(keys, history))
    assert len(keys) == len(history) == 2
    assert [mapping[k]["find"] for k in keys] == ["a", "e"], "行与记录错位"


# ---------- P1-5：跨 run 匹配 ----------

def test_cross_run_text_is_found_and_replaced(tmp_path):
    """P1-5：文本拆进两个 run 也要能搜到、能替换（旧引擎 per-run 匹配 0 命中）。"""
    p = _docx_with_split_runs(tmp_path / "split.docx", ["查找", "内容"])
    matches, skipped = core.extract_word_matches(str(p), "查找内容", False)
    assert len(matches) == 1, "跨 run 文本没搜到: %r" % matches
    assert matches[0]["start"] == 0 and matches[0]["end"] == 4
    assert matches[0]["text"] == "查找内容"

    core.replace_word_multi(str(p), _annotate(matches, "查找内容", "替换"), False)
    assert _para_texts(p) == ["替换"], _para_texts(p)


def test_cross_run_case_insensitive_replace(tmp_path):
    """P1-5：大小写不敏感模式下跨 run 替换同样走区间，不靠正则重扫。"""
    p = _docx_with_split_runs(tmp_path / "split2.docx", ["Hello", "World"])
    matches, _ = core.extract_word_matches(str(p), "helloworld", False)
    assert len(matches) == 1
    core.replace_word_multi(str(p), _annotate(matches, "helloworld", "X"), False)
    assert _para_texts(p) == ["X"]


# ---------- P1-6：超链接 / 修订内文本 ----------

def test_hyperlink_inner_text_is_matched(tmp_path):
    """P1-6：超链接内部的 run 也要能搜到（para.runs 取不到它们）。"""
    p = _docx_with_hyperlinked_run(tmp_path / "hl.docx")
    matches, _ = core.extract_word_matches(str(p), "目标文本", False)
    assert len(matches) == 1, "超链接内文本漏搜"
    core.replace_word_multi(str(p), _annotate(matches, "目标文本", "新文字"), False)
    doc = Document(str(p))
    inner = "".join(t.text or "" for t in doc.paragraphs[0]._p.iter(qn("w:t")))
    assert "新文字" in inner and "目标文本" not in inner, inner


def test_revision_inner_text_is_matched(tmp_path):
    """P1-6：修订插入（w:ins）内部的 run 也要能搜到。"""
    p = _docx_with_revision_run(tmp_path / "ins.docx")
    matches, _ = core.extract_word_matches(str(p), "目标文本", False)
    assert len(matches) == 1, "修订内文本漏搜"
    core.replace_word_multi(str(p), _annotate(matches, "目标文本", "新文字"), False)
    doc = Document(str(p))
    inner = "".join(t.text or "" for t in doc.paragraphs[0]._p.iter(qn("w:t")))
    assert "新文字" in inner and "目标文本" not in inner, inner


# ---------- P1-14：同 run 多命中 ----------

def test_multiple_hits_in_same_run_all_recorded(tmp_path):
    """P1-14：同一 run 内多次出现各记一条，且全部被替换（旧实现只记第一处）。"""
    p = make_docx(tmp_path / "multi.docx", "aa bb aa")
    matches, _ = core.extract_word_matches(str(p), "aa", False)
    assert len(matches) == 2, "同 run 多命中只记了 %d 笔" % len(matches)
    core.replace_word_multi(str(p), _annotate(matches, "aa", "X"), False)
    assert _para_texts(p) == ["X bb X"], _para_texts(p)


# ---------- P1-9：页眉重复枚举 ----------

def test_linked_header_not_enumerated_twice(tmp_path):
    """P1-9：第 2 节链接继承时，同一页眉不得被扫两遍（匹配数虚胖）。"""
    doc = Document()
    doc.sections[0].header.paragraphs[0].add_run("页眉目标文本")
    doc.add_paragraph("正文")
    doc.add_section()
    p = tmp_path / "linked.docx"
    doc.save(p)

    assert Document(str(p)).sections[1].header.is_linked_to_previous is True
    matches, _ = core.extract_word_matches(str(p), "页眉目标文本", False)
    assert len(matches) == 1, "链接页眉被重复枚举: %d 处" % len(matches)
    assert matches[0]["section_idx"] == 1


def test_multi_section_headers_each_counted_once(tmp_path):
    """P1-9：各自独立定义的页眉各记一次，链接节不重复。"""
    p = _docx_with_multi_section_headers(tmp_path / "ms.docx")
    matches, _ = core.extract_word_matches(str(p), "页眉文字", False)
    locations = sorted(m["location"] for m in matches)
    assert locations == ["第1节 页眉 第1段", "第2节 页眉 第1段"], locations


# ---------- P1-7：替换互斥 ----------

def test_start_search_refused_while_replace_running(tk_root, tmp_path, monkeypatch):
    """P1-7：替换批次进行中，搜索必须被拒绝（两批次同开文件会撞权限错）。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.selected_files = [str(tmp_path / "a.docx")]
    app._replace_thread = _AliveThread()

    infos = []
    monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: infos.append(a))

    app.start_search()

    assert infos, "替换进行中应拒绝再搜索"
    assert app._search_gen == 0, "被拒绝时不得开启新一轮搜索"


def test_execute_replace_refused_while_replace_running(tk_root, tmp_path, monkeypatch):
    """P1-7：替换进行中禁止再从预览窗发起一轮。"""
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    app = _make_replace_app(tk_root, tmp_path)
    app.preview_data = _content_preview_data(src)
    app._replace_thread = _AliveThread()

    warned = []
    monkeypatch.setattr(core.messagebox, "showwarning", lambda *a, **k: warned.append(a))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["World", "X"]])

    assert warned, "替换进行中应拒绝再执行"
    assert not win.destroyed, "被拒绝时预览窗应保持打开"
    assert Document(str(src)).paragraphs[0].runs[0].text == "Hello World"


def test_on_closing_asks_before_interrupting_replace(tk_root, tmp_path, monkeypatch):
    """P1-7：替换进行中关窗必须先确认，用户点否则窗口不能关。"""
    app = _make_replace_app(tk_root, tmp_path)
    app._replace_thread = _AliveThread()
    app.root.destroyed = False

    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: False)

    app.on_closing()

    assert not app.root.destroyed, "用户点「否」后窗口被关掉了"


def test_replace_worker_disables_search_during_batch(tk_root, tmp_path):
    """P1-7 + P1-10：替换 worker 全程禁用搜索入口，并按文件上报进度。"""
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    matches, _ = core.extract_word_matches(str(src), "World", True)
    app.replace_files_from_preview(
        {str(src): _annotate(matches, "World", "X")},
        conditions=[["World", "X"]])

    # 菜单删除后，「搜索入口」就是主按钮本身（_set_primary_state 切它）。
    fns = [c[0] for c in captured]
    assert app._set_primary_state in fns, "替换期间应禁用搜索入口"
    states = [c[1][0] for c in captured if c[0] == app._set_primary_state]
    assert states[0] == "disabled" and states[-1] == "normal", \
        "搜索入口应最后恢复: %r" % states

    progress = [c[1] for c in captured if c[0] == app._set_replace_progress]
    assert progress and progress[0][:2] == (1, 1), "应按文件上报进度: %r" % progress
    assert progress[-1][0] is None, "批次结束必须清空进度"


# ---------- P1-8：新文件名预检 ----------

def test_validate_filename_fixes_illegal_reserved_and_trailing():
    """P1-8：非法字符/保留设备名/尾点尾空格都要修正，且每项都有说明。"""
    name, problems = core.validate_filename("a:b*.txt")
    assert name == "a_b_.txt"
    assert any("非法字符" in p for p in problems)

    name, problems = core.validate_filename("CON.txt")
    assert name.startswith("CON") and name != "CON.txt"
    assert any("保留设备名" in p for p in problems)

    name, problems = core.validate_filename("abc. ")
    assert name == "abc"
    assert any("末尾" in p for p in problems)

    name, problems = core.validate_filename("  ")
    assert name == "" and problems


def test_plan_filename_renames_skips_collisions(tmp_path):
    """P1-8：批内撞名、目标已存在都要跳过并说明（不再等 os.rename 才炸）。"""
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    c = tmp_path / "c.docx"
    other = tmp_path / "same.docx"
    for f in (a, b, c, other):
        f.write_bytes(b"x")

    plan, warnings = core.plan_filename_renames([
        {"file_path": str(a), "new_name": "new.docx"},    # 正常执行
        {"file_path": str(b), "new_name": "new.docx"},    # 批内与上一条撞名
        {"file_path": str(c), "new_name": "same.docx"},   # 与磁盘上 other 撞名
    ])
    assert [p[1] for p in plan] == [str(tmp_path / "new.docx")], plan
    assert any("重复" in w for w in warnings), warnings
    assert any("已存在" in w for w in warnings), warnings


def test_filename_replace_sanitizes_new_name(tk_root, tmp_path, monkeypatch):
    """P1-8：执行路径上的新文件名先修正再改名，修正项写进日志。"""
    src = tmp_path / "abc.docx"
    src.write_bytes(b"x")
    app = _make_replace_app(tk_root, tmp_path)
    app.replace_filename_var.set(True)
    app.preview_data = [{
        "file_path": str(src), "file_name": "abc.docx",
        "matches": [{"type": "filename", "location": "文件名",
                     "new_name": "a:b.docx",               # 含非法字符
                     "context": "abc.docx -> a:b.docx"}],
        "checked": tk.BooleanVar(value=True),
    }]

    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: True)

    app.execute_replace_from_preview(_FakeWidget(), [["abc", "a:b"]])

    assert (tmp_path / "a_b.docx").exists(), "新文件名应被修正后改名"
    assert not src.exists()
    assert any("非法字符" in line for line in app._log_lines), \
        "修正项应写进日志: %r" % app._log_lines


# ---------- P1-10：进度反馈 ----------

def test_show_progress_updates_label_and_tolerates_missing(tk_root):
    """P1-10：进度写进状态栏标签；控件不存在时不得炸。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app._show_progress("正在替换", 2, 5, "a.docx")  # 无标签 → 静默返回
    app.progress_label = tk.Label(tk_root)
    app._show_progress("正在替换", 2, 5, "a.docx")
    assert app.progress_label.cget("text") == "正在替换 2/5 · a.docx"
    app._show_progress("正在替换", None, 5, "")
    assert app.progress_label.cget("text") == ""


# =====================================================================
# 依赖缺失不再静默
#
# 已实测复现的缺陷：全局 Python 缺 python-docx 时直接跑 .py，.docx 全部
# 落入"无匹配"分支——"未找到匹配内容"和"真的没有匹配"提示一模一样，用户
# 完全无法分辨（exe 把依赖打包进去了所以正常）。这三组用例守住：
# 缺库必须在搜索时报出来、启动时要说清楚、零匹配对话框要能区分两种原因。
# =====================================================================

def test_search_worker_reports_missing_docx_library(tmp_path, monkeypatch):
    """缺 python-docx 时，.docx 要报明确错误，不能静默返回空匹配。"""
    p = tmp_path / "a.docx"
    p.write_bytes(b"x")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    monkeypatch.setattr(core, "DOCX_AVAILABLE", False)
    app._search_worker([str(p)], [["x", "y"]], True, False)

    finish = [c for c in captured if c[0] == app._finish_search]
    assert finish, "收尾结果未投递"
    errors = finish[0][1][3]
    assert any("未安装 python-docx" in e for e in errors), \
        "缺库时应在错误里说清楚: %r" % (errors,)
    raw, counts, total = finish[0][1][0], finish[0][1][1], finish[0][1][2]
    assert raw == [] and counts[str(p)] == 0 and total == 0, "缺库时不应产生匹配"


def test_search_worker_reports_missing_openpyxl_library(tmp_path, monkeypatch):
    """缺 openpyxl 时，.xlsx 同样要报明确错误。"""
    p = tmp_path / "a.xlsx"
    p.write_bytes(b"x")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    monkeypatch.setattr(core, "EXCEL_AVAILABLE", False)
    app._search_worker([str(p)], [["x", "y"]], True, False)

    finish = [c for c in captured if c[0] == app._finish_search]
    errors = finish[0][1][3]
    assert any("未安装 openpyxl" in e for e in errors), \
        "缺库时应在错误里说清楚: %r" % (errors,)
    assert finish[0][1][0] == [] and finish[0][1][2] == 0


def test_check_dependencies_dialog_lists_missing_libs(monkeypatch):
    """启动检测要说出缺哪个库、影响哪种格式、怎么装。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    infos = []
    monkeypatch.setattr(core, "DOCX_AVAILABLE", False)
    monkeypatch.setattr(core.messagebox, "showinfo",
                        lambda *a, **k: infos.append((a, k)))

    app._check_dependencies()

    assert infos, "缺库时必须弹窗告知"
    msg = infos[0][0][1]
    assert "python-docx" in msg and ".docx" in msg, "未说清缺什么: %r" % msg
    assert "pip install" in msg, "未给出安装出路: %r" % msg


def test_check_dependencies_silent_when_all_present(monkeypatch):
    """依赖齐全时启动不弹窗（exe 版场景）。"""
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    infos = []
    monkeypatch.setattr(core.messagebox, "showinfo",
                        lambda *a, **k: infos.append((a, k)))
    monkeypatch.setattr(core, "DOCX_AVAILABLE", True)
    monkeypatch.setattr(core, "EXCEL_AVAILABLE", True)

    app._check_dependencies()

    assert infos == [], "依赖齐全时不应弹窗"


def test_finish_search_zero_match_names_dependency_cause(monkeypatch):
    """零匹配若因缺库，对话框必须说"搜不了"，不能冒充"没找到"。"""
    app = _bare_app()
    app._search_gen = 1
    app.preview_data = []
    app._search_counts = {}
    app._total_matches = 0
    states = []
    app._set_primary_state = states.append
    app.update_status = lambda: None
    app._apply_search_status = lambda counts: None
    app.add_to_history = lambda *a: None

    infos = []
    monkeypatch.setattr(core.messagebox, "showinfo",
                        lambda *a, **k: infos.append((a, k)))

    dep_error = "a.docx: 未安装 python-docx，.docx 无法搜索（请使用 exe 版...）"
    # 依赖缺失由 worker 单独上报（dep_errors），不再靠翻文案猜——文案多语言后
    # 英文界面里永远搜不到"未安装"三个字。
    app._finish_search([], {}, 0, [], [["x", "y"]], False, gen=1,
                       dep_errors=[dep_error])

    assert infos, "应弹提示框"
    msg = infos[0][0][1]
    assert "未安装" in msg, "零匹配对话框未说明依赖缺失: %r" % msg
    assert "未找到匹配内容" not in msg, "不能冒充'没找到': %r" % msg
    assert states and states[-1] == "normal", "按钮应恢复（finally 幂等重放一次属正常）"


def test_finish_search_zero_match_still_says_no_match_without_dep_errors(monkeypatch):
    """真的没有匹配时，提示保持原样（不误报依赖问题）。"""
    app = _bare_app()
    app._search_gen = 1
    app.preview_data = []
    app._search_counts = {}
    app._total_matches = 0
    app._set_primary_state = lambda s: None
    app.update_status = lambda: None
    app._apply_search_status = lambda counts: None
    app.add_to_history = lambda *a: None

    infos = []
    monkeypatch.setattr(core.messagebox, "showinfo",
                        lambda *a, **k: infos.append((a, k)))

    app._finish_search([], {}, 0, [], [["x", "y"]], False, gen=1)

    assert infos and "未找到匹配内容" in infos[0][0][1], \
        "无依赖问题时应保持原提示: %r" % (infos,)


# ---------- 双击跳转的精确定位 ----------


def test_word_match_records_occurrence_for_jump(tmp_path):
    """同一段里同一个词出现多次时，每条匹配都记下「它在本段是第几次出现」。

    跳转时靠这个序号在 Word 里选中「用户双击的那一条」；没有它就只能永远
    停在第一处。
    """
    path = tmp_path / "occurrence.docx"
    doc = Document()
    doc.add_paragraph("Dept-1 与 dept-2 和 DEPT-3")
    doc.save(str(path))

    matches, _ = core.extract_word_matches(str(path), "dept", False)

    assert [m["occurrence"] for m in matches] == [1, 2, 3], \
        "同段三次出现的序号应依次为 1/2/3: %r" % [m["occurrence"] for m in matches]
    assert [m["text"] for m in matches] == ["Dept", "dept", "DEPT"], \
        "每条匹配要记下自己命中的原文（保留大小写）"
    assert all(m["find_text"] == "dept" and m["case_sensitive"] is False
               for m in matches), "匹配要带上生成它的查找词与大小写语义"


def test_word_highlight_uses_recorded_match_not_search_box(monkeypatch):
    """高亮用的是「匹配记录里的查找词 + 出现序号」，不是搜索框的当前内容。

    搜索框内容属于下一次搜索；用户双击的却是列表里那一条旧记录，两者可能
    已经不是一回事。
    """
    app = _bare_app()
    calls = []
    monkeypatch.setattr(core, "_word_find_nth",
                        lambda container, text, occurrence, match_case:
                        calls.append((text, occurrence, match_case)))
    match = {"find_text": "部门", "occurrence": 3, "case_sensitive": True}

    app._word_highlight_match(object(), match)

    assert calls == [("部门", 3, True)], "应按记录定位第 3 次出现: %r" % (calls,)


def test_word_highlight_skips_caret_find_text():
    """查找词含 ^ 时不做词级高亮：Word 把 ^ 当特殊代码，序号会对不上。

    宁可保住"整段被选中"，也不能亮到错误的位置。
    """
    app = _bare_app()
    app._word_highlight_match(object(), {"find_text": "^p", "occurrence": 1})

    assert app._log_lines and "^" in app._log_lines[0], \
        "跳过词级高亮要留痕: %r" % (app._log_lines,)


# ---------- 配置与历史的读取失败要留痕 ----------


def test_load_config_failure_is_reported_in_startup_notes(tk_root, monkeypatch, tmp_path):
    """坏掉的配置文件不该静默：启动日志要能说出原因。

    静默的话，用户只会经历"设置没生效"，而没有任何可查的线索。
    """
    bad = tmp_path / "config.json"
    bad.write_text("{ 坏掉的 json", encoding="utf-8")
    monkeypatch.setattr(core, "CONFIG_FILE", str(bad))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = _bare_app()
    app.startup_notes = []

    app.load_config()

    # startup_notes 存的是 (文案键, 格式化参数)：读配置时语言还没读出来，只能
    # 先记键，等界面就位后再按当时的语言渲染。故这里断言键、并顺带验一次渲染。
    notes = dict(app.startup_notes)
    assert "note_config_failed" in notes, "读取失败必须留痕: %r" % (app.startup_notes,)
    assert "配置" in core.T("note_config_failed", **notes["note_config_failed"]), \
        "渲染后的文案应说明是配置问题"


def test_save_config_does_not_persist_backup_switch(tk_root, monkeypatch, tmp_path):
    """配置文件里只留「备份路径 / 匹配大小写」。

    以下四项都不落盘，启动时各有各的来源：
    - 备份开关启动时一律强制开启，读进去也不会生效；
    - 预览提示开关随菜单一起删了，没有第二个入口能把它打开，持久化只会
      变成一个关不回来的隐藏状态；
    - 语言与主题每次启动取系统当前值（见 load_config），存进来也不会被读。
    """
    cfg = tmp_path / "config.json"
    monkeypatch.setattr(core, "CONFIG_FILE", str(cfg))
    app = _bare_app()
    app.backup_var = tk.BooleanVar(master=tk_root, value=False)
    app.backup_path = tk.StringVar(master=tk_root, value=str(tmp_path))
    app.case_sensitive_var = tk.BooleanVar(master=tk_root, value=False)

    app.save_config()

    saved = json.loads(cfg.read_text(encoding="utf-8"))
    assert set(saved) == {"backup_path", "case_sensitive"}, \
        "配置里不该再有备份开关 / 预览提示开关 / 语言 / 主题: %r" % (saved,)


# ---------- 运行中禁用文件操作 ----------

class _FakeButton:
    """圆角按钮（Canvas 自绘）的测试替身：只记录启用态与重绘次数。"""

    def __init__(self):
        self.repaints = 0
        self._br_enabled = True

    def _br_repaint(self):
        self.repaints += 1


FILE_OP_BUTTONS = ("add_files_btn", "remove_files_btn", "clear_files_btn")


def _btn_text(btn):
    """圆角按钮（Canvas 自绘）上的文字：存在 brtext 标签的 item 里。"""
    for item in btn.find_withtag("brtext"):
        return btn.itemcget(item, "text")
    return ""


def _app_with_file_ops():
    """挂上三颗假按钮：本组只关心"状态有没有切到全部文件操作入口"。"""
    app = _bare_app()
    for name in FILE_OP_BUTTONS:
        setattr(app, name, _FakeButton())
    return app


def test_set_file_ops_state_covers_all_buttons():
    """工具栏三颗按钮必须一起切（菜单删除后，这就是全部文件操作入口）。"""
    app = _app_with_file_ops()

    app._set_file_ops_state(False)

    for name in FILE_OP_BUTTONS:
        assert getattr(app, name)._br_enabled is False, f"{name} 未随运行中禁用"

    app._set_file_ops_state(True)

    for name in FILE_OP_BUTTONS:
        assert getattr(app, name)._br_enabled is True, f"{name} 未恢复可点"


def test_set_file_ops_state_tolerates_missing_widgets():
    """控件尚未建立（或已销毁）时不得抛异常——启动早期与关窗后都会走到。"""
    _bare_app()._set_file_ops_state(False)


def test_start_search_disables_file_ops(tk_root, tmp_path, monkeypatch):
    """搜索启动后：文件操作三颗按钮一并置灰。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.selected_files = [str(tmp_path / "a.docx")]
    app.find_entry.insert(0, "x")
    for name in FILE_OP_BUTTONS:
        setattr(app, name, _FakeButton())
    monkeypatch.setattr(core, "threading", _FakeThreading([]))

    app.start_search()

    assert app.add_files_btn._br_enabled is False, "搜索中「添加文件」仍可点"


def test_finish_search_restores_file_ops_even_on_error():
    """收尾必须恢复文件操作入口，异常路径也不例外（否则永久卡灰）。"""
    app = _app_with_file_ops()
    app._search_gen = 1
    app.preview_data = []
    app._set_primary_state = lambda state: None
    app.update_status = lambda: None

    def boom(counts):
        raise RuntimeError("模拟刷列表时炸了")

    app._apply_search_status = boom

    with pytest.raises(RuntimeError):
        app._finish_search([], {}, 1, [], [["a", "b"]], False, gen=1)

    for name in FILE_OP_BUTTONS:
        assert getattr(app, name)._br_enabled is True, f"{name} 异常路径下未恢复"


def test_finish_search_stale_result_keeps_file_ops_disabled():
    """过期结果早退不得恢复：新一轮搜索正持有禁用态，恢复等于把它放开。"""
    app = _app_with_file_ops()
    app._search_gen = 5
    app.preview_data = []
    app._set_primary_state = lambda state: None

    app._finish_search([{"file_path": "f"}], {"f": 1}, 1, [], [["a", "b"]], False, gen=4)

    for name in FILE_OP_BUTTONS:
        assert getattr(app, name).repaints == 0, f"{name} 被过期结果误恢复"


def test_replace_worker_disables_and_restores_file_ops(tmp_path):
    """替换全程同样禁用文件操作，收尾恢复——与搜索运行中的置灰对称。"""
    src = tmp_path / "a.docx"
    make_docx(src, "Hello World")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    matches, _ = core.extract_word_matches(str(src), "World", True)
    app.replace_files_from_preview(
        {str(src): _annotate(matches, "World", "X")},
        conditions=[["World", "X"]])

    states = [c[1][0] for c in captured if c[0] == app._set_file_ops_state]
    assert states[:1] == [False], "替换期间未禁用文件操作: %r" % (states,)
    assert states[-1] is True, "替换结束未恢复: %r" % (states,)


# ---------- 状态栏不得停在过期状态 ----------

def _status_app(master):
    """只挂状态栏刷新所需属性的最小实例。"""
    app = _bare_app()
    app.backup_var = tk.BooleanVar(master=master, value=True)
    app.selected_files = []
    app._total_matches = 0
    app.status_label = tk.Label(master)
    return app


def test_toggle_backup_refreshes_status_bar(tk_root, monkeypatch):
    """关掉备份后状态栏必须立刻改口。

    状态栏那行字只在别的操作触发时才重写；备份状态是安全信号，"已关却显示开启"
    比不显示更糟。
    """
    app = _status_app(tk_root)
    refreshed = []
    app.update_status = lambda: refreshed.append(app.backup_var.get())
    monkeypatch.setattr(core.messagebox, "showwarning", lambda *a, **k: "yes")

    app.backup_var.set(False)   # 勾选动作本身改的是变量，回调只负责警告与刷新
    app.toggle_backup()

    assert app.backup_var.get() is False
    assert refreshed == [False], "关闭备份后状态栏未刷新"


def test_toggle_backup_rollback_also_refreshes(tk_root, monkeypatch):
    """警告框点「否」回滚成开启，状态栏同样要刷新（那时确实仍是开启）。"""
    app = _status_app(tk_root)
    refreshed = []
    app.update_status = lambda: refreshed.append(app.backup_var.get())
    monkeypatch.setattr(core.messagebox, "showwarning", lambda *a, **k: "no")

    app.backup_var.set(False)
    app.toggle_backup()

    assert app.backup_var.get() is True, "点否应回滚为开启"
    assert refreshed == [True], "回滚后状态栏未刷新"


def test_load_rule_refreshes_status_bar(tk_root, tmp_path, monkeypatch):
    """加载规则会改掉备份开关，状态栏必须跟着走。"""
    rule = tmp_path / "r.json"
    rule.write_text(json.dumps({"find": "a", "replace": "b", "case_sensitive": True,
                                "replace_filename": False, "backup_enabled": False}),
                    encoding="utf-8")
    app = _make_replace_app(tk_root, tmp_path)
    refreshed = []
    app.update_status = lambda: refreshed.append(app.backup_var.get())
    monkeypatch.setattr(core.filedialog, "askopenfilename", lambda *a, **k: str(rule))
    monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(core.messagebox, "showerror", lambda *a, **k: None)

    app.load_rule()

    assert app.backup_var.get() is False, "规则里的备份开关未生效"
    assert refreshed == [False], "加载规则后状态栏未刷新"


def _tree_with_rows(master, tmp_path, counts):
    """造一个带行的真文件列表：iid 是路径，match 列写死匹配数。"""
    app = _status_app(master)
    app.files_tree = _make_files_tree(master)
    paths = []
    for i, count in enumerate(counts, 1):
        p = tmp_path / f"f{i}.docx"
        p.write_bytes(b"x")
        app._insert_file_row(str(p))
        app.files_tree.set(str(p), "match", str(count))
        paths.append(str(p))
    app.selected_files = list(paths)
    app._total_matches = sum(counts)
    return app, paths


def test_clear_files_resets_match_count(tk_root, tmp_path):
    """清空列表后「匹配 N 处」必须归零。

    它读的是搜索当时存下的数字，删行不会让它自己变小——只刷状态栏是不够的，
    必须按列表实际内容重算。
    """
    app, _ = _tree_with_rows(tk_root, tmp_path, [247, 248])

    app.clear_files()

    assert app.files_tree.get_children() == (), "列表应被清空"
    assert app._total_matches == 0, "清空后匹配数仍挂着旧值: %r" % (app._total_matches,)
    assert "匹配 0 处" in app.status_label.cget("text"), app.status_label.cget("text")


def test_remove_files_recomputes_match_count(tk_root, tmp_path):
    """移除单个文件同样要重算：不同步就会出现"文件没了、匹配数还在"。"""
    app, paths = _tree_with_rows(tk_root, tmp_path, [100, 50])
    app.files_tree.selection_set(paths[0])

    app.remove_files()

    assert app._total_matches == 50, "移除后匹配数未按剩余文件重算: %r" % (app._total_matches,)
    assert "匹配 50 处" in app.status_label.cget("text"), app.status_label.cget("text")


# =====================================================================
# 多条件替换
#
# 三条语义（与用户逐条定稿），每条都要有机器守门：
#   1. 所有条件在**原始内容**上匹配，绝不链式叠加
#      （a→b 之后再拿结果跑 b→c 是错误语义：新生成的 b 不该被改）
#   2. 查找内容不允许重复，替换为可以重复
#   3. 同一位置被多条件命中时**整批停下**，被挤到的那条必须报给用户
# =====================================================================

def test_plan_span_replacements_overlap_keeps_one():
    """重叠只保留一条（"ab" 与 "a" 同时命中同一处，只留一条进 accepted），
    另一条进 conflicts 报给用户——但界面根本不执行，只据此拦下整批。"""
    hits = [
        {"start": 0, "end": 1, "cond_idx": 0, "find_text": "a", "replace_text": "b"},
        {"start": 0, "end": 2, "cond_idx": 1, "find_text": "ab", "replace_text": "c"},
    ]
    accepted, conflicts = core.plan_span_replacements(hits)
    assert [(h["find_text"], h["replace_text"]) for h in accepted] == [("ab", "c")]
    assert len(conflicts) == 1
    hit, rival = conflicts[0]
    assert hit["find_text"] == "a" and rival["find_text"] == "ab"


def test_plan_span_replacements_equal_length_keeps_earlier_condition():
    """等长部分重叠（"ab" 与 "bc"）同样只保留一条，另一条照旧上报。"""
    hits = [
        {"start": 0, "end": 2, "cond_idx": 0, "find_text": "ab", "replace_text": "X"},
        {"start": 1, "end": 3, "cond_idx": 1, "find_text": "bc", "replace_text": "Y"},
    ]
    accepted, conflicts = core.plan_span_replacements(hits)
    assert [h["find_text"] for h in accepted] == ["ab"]
    assert conflicts and conflicts[0][1]["find_text"] == "ab"


def test_plan_span_replacements_returns_desc_and_no_false_conflict():
    hits = [
        {"start": 0, "end": 1, "cond_idx": 0, "find_text": "a", "replace_text": "b"},
        {"start": 5, "end": 6, "cond_idx": 1, "find_text": "c", "replace_text": "d"},
    ]
    accepted, conflicts = core.plan_span_replacements(hits)
    assert [h["start"] for h in accepted] == [5, 0], "应按 start 降序返回（从后往前应用）"
    assert conflicts == []


def test_apply_filename_conditions_matches_original_only():
    """a→b 与 b→c 同时给出：原名里的 a、b 各改一次，新生成的 b 不再被改。"""
    new_name, notes = core.apply_filename_conditions(
        "ab.docx", [["a", "b"], ["b", "c"]], True)
    assert new_name == "bc.docx", new_name
    assert notes == []


def test_apply_filename_conditions_reports_overlap():
    """重叠必须带回一条说明（不能静默丢弃），文件名模式也一样。"""
    new_name, notes = core.apply_filename_conditions(
        "abc.txt", [["a", "X"], ["ab", "Y"]], True)
    assert new_name == "Yc.txt", new_name
    assert notes and "重叠" in notes[0]


def test_validate_conditions_rejects_duplicate_find_and_empty():
    cleaned, errors = core.validate_conditions([["A", "1"], ["A", "2"]], True)
    assert cleaned == [["A", "1"]] and any("重复" in e for e in errors)

    cleaned, errors = core.validate_conditions([["", "x"], ["a", ""]], False)
    assert cleaned == [["a", ""]] and any("不能为空" in e for e in errors)


def test_validate_conditions_case_flag_decides_variants():
    """大小写不敏感时 A 与 a 是同一条查找内容；敏感时是两条。"""
    cleaned, errors = core.validate_conditions([["A", "1"], ["a", "2"]], False)
    assert len(cleaned) == 1 and errors
    cleaned, errors = core.validate_conditions([["A", "1"], ["a", "2"]], True)
    assert len(cleaned) == 2 and not errors


def test_validate_conditions_allows_duplicate_replace():
    """替换为可以重复。"""
    cleaned, errors = core.validate_conditions([["a", "X"], ["b", "X"]], True)
    assert cleaned == [["a", "X"], ["b", "X"]] and errors == []


def test_replace_word_multi_matches_original_only(tmp_path):
    """核心语义：a→b 与 b→c 一次跑完，docx 里的 "ab" 应得 "bc"（不是 "cc"）。"""
    p = make_docx(tmp_path / "multi.docx", "ab")
    matches = []
    for idx, (find_text, replace_text) in enumerate([("a", "b"), ("b", "c")], 1):
        ms, _ = core.extract_word_matches(str(p), find_text, True)
        matches += _annotate(ms, find_text, replace_text, idx)

    skipped = core.replace_word_multi(str(p), matches, True)

    assert skipped == []
    assert Document(str(p)).paragraphs[0].text == "bc"


def test_replace_word_multi_longer_wins_and_reports(tmp_path):
    """同段落内 "ab" 与 "a" 都命中：只改一处，另一处写进 skipped 说明。"""
    p = make_docx(tmp_path / "wins.docx", "ab")
    matches = []
    for idx, (find_text, replace_text) in enumerate([("a", "Z"), ("ab", "Y")], 1):
        ms, _ = core.extract_word_matches(str(p), find_text, True)
        matches += _annotate(ms, find_text, replace_text, idx)

    skipped = core.replace_word_multi(str(p), matches, True)

    assert Document(str(p)).paragraphs[0].text == "Y"
    assert any("重叠" in reason for _loc, reason in skipped), skipped


def test_detect_condition_conflicts_reports_overlap(tmp_path):
    p = make_docx(tmp_path / "conf.docx", "ab")
    matches = []
    for idx, (find_text, replace_text) in enumerate([("a", "Z"), ("ab", "Y")], 1):
        ms, _ = core.extract_word_matches(str(p), find_text, True)
        matches += _annotate(ms, find_text, replace_text, idx)

    conflicts = core.detect_condition_conflicts(matches)

    assert len(conflicts) == 1
    hit, rival = conflicts[0]
    assert hit["find_text"] == "a" and rival["find_text"] == "ab"


def test_detect_condition_conflicts_no_false_positive(tmp_path):
    """互不重叠的多条件不得报冲突（否则用户每次都被无谓弹窗拦住）。"""
    p = make_docx(tmp_path / "clean.docx", "苹果和香蕉")
    matches = []
    for idx, (find_text, replace_text) in enumerate([("苹果", "梨"), ("香蕉", "桃")], 1):
        ms, _ = core.extract_word_matches(str(p), find_text, True)
        matches += _annotate(ms, find_text, replace_text, idx)

    assert core.detect_condition_conflicts(matches) == []


def test_extract_excel_match_carries_spans_and_find_text(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "abab"
    p = tmp_path / "spans.xlsx"
    wb.save(p)

    matches, _ = core.extract_excel_matches(str(p), "ab", True)

    assert matches[0]["find_text"] == "ab"
    assert matches[0]["spans"] == [(0, 2), (2, 4)]


def test_replace_excel_multi_matches_original_only(tmp_path):
    """同一单元格 "ab" + 条件 a→b、b→c：一次替换应得 "bc"，不是链式的 "cc"。"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "ab"
    p = tmp_path / "multi.xlsx"
    wb.save(p)
    matches = []
    for idx, (find_text, replace_text) in enumerate([("a", "b"), ("b", "c")], 1):
        ms, _ = core.extract_excel_matches(str(p), find_text, True)
        matches += _annotate(ms, find_text, replace_text, idx)

    core.replace_excel_multi(str(p), matches, True)

    assert load_workbook(str(p)).active["A1"].value == "bc"


def test_replace_excel_still_replaces_all_occurrences(tmp_path):
    """单条件行为不变：一个单元格里的多处出现全部替换。"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "abab"
    p = tmp_path / "single.xlsx"
    wb.save(p)

    matches, _ = core.extract_excel_matches(str(p), "ab", True)
    core.replace_excel_multi(str(p), _annotate(matches, "ab", "XY"), True)

    assert load_workbook(str(p)).active["A1"].value == "XYXY"


def test_search_worker_tags_matches_with_condition(tmp_path):
    """每条命中要带上它所属的条件与替换词，预览/执行都靠它。"""
    src = make_docx(tmp_path / "tagged.docx", "苹果和香蕉")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    app._search_worker([str(src)], [["苹果", "梨"], ["香蕉", "桃"]], True, False)

    finish = [c for c in captured if c[0] == app._finish_search][-1]
    (raw, _counts, total, _errors, _conditions, _is_filename, _gen,
     conflicts, _dep_errors) = finish[1]
    assert total == 2
    matched = raw[0]["matches"]
    assert {m["replace_text"] for m in matched} == {"梨", "桃"}
    assert {m["cond_idx"] for m in matched} == {1, 2}
    assert conflicts == []


def test_search_worker_reports_overlap_conflict(tmp_path):
    """重叠命中要在搜索阶段就报出来，不能等执行完才发现少改了。"""
    src = make_docx(tmp_path / "overlap.docx", "ab")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    app._search_worker([str(src)], [["a", "Z"], ["ab", "Y"]], True, False)

    finish = [c for c in captured if c[0] == app._finish_search][-1]
    conflicts = finish[1][7]
    assert conflicts and "重叠" in conflicts[0]


def test_search_worker_keeps_single_condition_behaviour(tmp_path):
    """单条件走同一条 worker：命中仍带 find_text，区间照旧。"""
    src = make_docx(tmp_path / "single.docx", "Hello World")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a, **k: captured.append((fn, a, k))

    app._search_worker([str(src)], [["World", "YehOlt"]], True, False)

    finish = [c for c in captured if c[0] == app._finish_search][-1]
    (raw, _counts, total, _errors, _conditions, _is_filename, _gen,
     _conflicts, _dep_errors) = finish[1]
    assert total == 1
    assert raw[0]["matches"][0]["find_text"] == "World"


def test_finish_search_records_multi_conditions_as_one_group(tk_root):
    """多条件现在也进历史：一组条件就是一条记录，整组一起记。"""
    app = _bare_app()
    app._search_gen = 1
    app.preview_data = []
    app._search_counts = {}
    app._total_matches = 0
    app.case_sensitive_var = tk.BooleanVar(value=False)
    app.replace_filename_var = tk.BooleanVar(value=False)
    app._apply_search_status = lambda counts: None
    app.update_status = lambda: None
    app._set_primary_state = lambda s: None
    app._set_search_menu_state = lambda v: None
    app._set_file_ops_state = lambda v: None
    app.create_preview_window = lambda *a: None
    seen = []
    app.add_to_history = lambda *a: seen.append(a)
    raw = [{"file_path": "f", "file_name": "f",
            "matches": [{"type": "paragraph", "location": "第1段", "context": "x",
                         "find_text": "a", "cond_idx": 1}]}]

    app._finish_search(raw, {"f": 1}, 1, [], [["a", "b"], ["c", "d"]], False, gen=1)
    assert seen == [([["a", "b"], ["c", "d"]], False, False)], \
        "多条件必须整组写进历史，而不是被跳过"

    app._finish_search(raw, {"f": 1}, 1, [], [["a", "b"]], False, gen=1)
    assert seen[-1] == ([["a", "b"]], False, False), "单条件同样按组结构记录"

    app.replace_filename_var.set(True)
    app._finish_search(raw, {"f": 1}, 1, [], [["a", "b"]], True, gen=1)
    assert seen[-1] == ([["a", "b"]], False, True), \
        "作用对象（内容 / 文件名）要一起进历史，否则载入出来的是另一回事"


def test_multi_condition_checkbox_placed_after_filename_checkbox():
    """多条件勾选框紧跟在「替换文件名（而非内容）」右边。

    文案已走 _t（双语），故按 key 出现顺序断言，不再匹配写死的中文；
    勾选框改自绘后 key 由 _make_check 传入，仍按源码里 key 的先后比。
    """
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert (src.index('_make_check(row2, "replace_filename_opt"')
            < src.index('_make_check(row2, "multi_condition"')), \
        "多条件勾选框应排在替换文件名勾选框右边"


# =====================================================================
# 重叠命中 → 整批停下（不自动裁决）
#
# 原规则是工具替用户在冲突位置挑一条执行，用户否决了——挑哪条是内容判断，
# 工具无权代劳。现在冲突一律拦下整批，一个文件都不改。
# =====================================================================

def _conflicting_preview_data(path, text="ab"):
    """同一段落**同一区间**被两条条件命中的预览数据。"""
    def mk(find, repl, cond_idx):
        return {"type": "paragraph", "para_idx": 1, "run_idx": 1,
                "start": 0, "end": 2,
                "location": "第1段", "context": text, "text": text,
                "find_text": find, "replace_text": repl, "cond_idx": cond_idx}

    long_hit, short_hit = mk("ab", "X", 1), mk("a", "Y", 2)
    return [{
        "file_path": str(path), "file_name": Path(path).name,
        "matches": [long_hit, short_hit],
        "match_data": [{"match": long_hit, "checked": tk.BooleanVar(value=True)},
                       {"match": short_hit, "checked": tk.BooleanVar(value=True)}],
        "checked": tk.BooleanVar(value=True),
        "conflicts": [],
    }]


def test_conflict_reason_no_longer_promises_auto_skip():
    """文案不得再承诺"已由某某优先、本条跳过"——那正是被否决的自动裁决。"""
    hit = {"start": 0, "end": 1, "find_text": "a"}
    rival = {"start": 0, "end": 2, "find_text": "ab"}
    text = core.conflict_reason(hit, rival)
    assert "重叠" in text
    assert "跳过" not in text, "不得再写「本条跳过」，用户会以为工具已经替他决定好了"
    # 既然一律拦下、根本不执行，"谁排在前面"已无意义，文案里也不该出现
    assert "更长者" not in text and "排在前面" not in text and "优先" not in text, text
    assert "a" in text and "ab" in text, "两条条件都要指名，用户才知道是哪两条撞上了"


def test_conflict_block_message_states_whole_batch_stops():
    """阻断提示必须说清"整批停下、一个文件都不改"，并给出两条出路。"""
    msg = core.conflict_block_message(["f.docx 第1段：重叠"])
    assert "整批停下" in msg
    assert "一个文件都不会改动" in msg
    assert "取消勾选" in msg, "必须告诉用户可以取消勾选冲突行来放行"
    assert "f.docx" in msg


def test_conflict_block_message_folds_long_lists():
    """冲突很多时不能把弹窗撑成一面墙：前 8 条明列，其余折叠并指向日志。"""
    msg = core.conflict_block_message([f"第{i}处" for i in range(20)])
    assert "第0处" in msg and "第7处" in msg
    assert "第8处" not in msg
    assert "另有 12 处" in msg


def test_selected_conflicts_only_counts_checked_rows(tk_root, tmp_path):
    """取消勾选冲突行后必须放行：按搜索总数一刀切会把整批永久锁死。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.preview_data = _conflicting_preview_data(tmp_path / "a.docx")

    assert len(app.selected_conflicts()) == 1, "两条条件命中同一区间应报 1 处冲突"

    app.preview_data[0]["match_data"][1]["checked"].set(False)
    assert app.selected_conflicts() == [], "重叠的那条取消勾选后冲突随之消失"


def test_selected_conflicts_covers_filename_mode(tk_root, tmp_path):
    """文件名模式一个文件就是一整条记录，冲突清单来自搜索阶段挂上的字段。"""
    app = _make_replace_app(tk_root, tmp_path)
    note = "文件名：查找「a」的命中与「ab」重叠"
    app.preview_data = [{
        "file_path": str(tmp_path / "ab.docx"), "file_name": "ab.docx",
        "matches": [{"type": "filename", "location": "文件名", "new_name": "Xc.docx"}],
        "checked": tk.BooleanVar(value=True),
        "conflicts": [note],
    }]
    assert app.selected_conflicts() == [note]

    app.preview_data[0]["checked"].set(False)
    assert app.selected_conflicts() == [], "取消勾选该文件后不再拦"


def test_refresh_conflict_gate_disables_and_restores_exec_button(tk_root, tmp_path):
    """有冲突就按死「替换勾选项」并在状态栏说明；冲突消失后自动恢复。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.preview_data = _conflicting_preview_data(tmp_path / "a.docx")

    class _Label:
        def __init__(self):
            self.text, self.fg = "", None

        def winfo_exists(self):
            return True

        def config(self, **kw):
            self.text, self.fg = kw.get("text", self.text), kw.get("fg", self.fg)

    class _Button:
        def __init__(self):
            self.state, self.bg, self.fg = "normal", None, None

        def config(self, **kw):
            self.state = kw.get("state", self.state)

        def configure(self, **kw):
            self.bg, self.fg = kw.get("bg", self.bg), kw.get("fg", self.fg)

    app.conflict_label, app.preview_exec_btn = _Label(), _Button()
    app._refresh_conflict_gate()
    assert app.preview_exec_btn.state == "disabled"
    assert "重叠" in app.conflict_label.text

    app.preview_data[0]["match_data"][1]["checked"].set(False)
    app._refresh_conflict_gate()
    assert app.preview_exec_btn.state == "normal"
    assert app.conflict_label.text == ""


def test_execute_replace_aborts_whole_batch_on_conflict(tk_root, tmp_path, monkeypatch):
    """核心行为：有重叠就整批停下——不建备份夹、不起替换线程、原文件零改动。"""
    src = tmp_path / "a.docx"
    make_docx(src, "ab")
    app = _make_replace_app(tk_root, tmp_path)
    app.preview_data = _conflicting_preview_data(src)

    errors, started = [], []
    monkeypatch.setattr(core.messagebox, "showerror", lambda *a, **k: errors.append(a))
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(core, "threading", _FakeThreading(started))
    monkeypatch.setattr(core, "create_backup_folder",
                        lambda *a, **k: pytest.fail("冲突拦下时不该去建备份文件夹"))

    win = _FakeWidget()
    app.execute_replace_from_preview(win, [["ab", "X"], ["a", "Y"]])

    assert errors, "应弹出错误框说明整批停下"
    assert started == [], "整批停下时不得起替换线程"
    assert not win.destroyed, "拦下时预览窗保持打开，用户才好去取消勾选"
    assert Document(str(src)).paragraphs[0].runs[0].text == "ab", "原文件必须零改动"
    assert any("整批中止" in line for line in app._log_lines)


def test_execute_replace_conflict_gate_runs_before_backup_creation(tk_root, tmp_path, monkeypatch):
    """拦截位置必须早于建备份夹，否则用户吃中止时还多一个空备份文件夹。"""
    src = tmp_path / "a.docx"
    make_docx(src, "ab")
    backup_dir = tmp_path / "bak"
    backup_dir.mkdir()
    app = _make_replace_app(tk_root, tmp_path, backup_dir=backup_dir)
    app.preview_data = _conflicting_preview_data(src)

    monkeypatch.setattr(core.messagebox, "showerror", lambda *a, **k: None)
    monkeypatch.setattr(core.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(core, "threading", _FakeThreading([]))

    app.execute_replace_from_preview(_FakeWidget(), [["ab", "X"], ["a", "Y"]])

    assert list(backup_dir.iterdir()) == [], "整批停下时不应留下空备份文件夹"


def test_search_worker_attaches_per_file_conflicts(tmp_path):
    """冲突清单要挂在文件记录上，预览窗才判得出"这一批还能不能执行"。"""
    p = tmp_path / "a.docx"
    make_docx(p, "ab")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a: captured.append((fn, a))

    app._search_worker([str(p)], [["ab", "X"], ["a", "Y"]], True, False)

    # _search_worker 先投递进度、再投递收尾，不能假设 captured[0] 就是收尾
    finish = [c for c in captured if c[0] == app._finish_search]
    raw, conflicts = finish[0][1][0], finish[0][1][7]
    assert len(raw) == 1
    assert raw[0]["conflicts"], "文件记录上必须带上自己的冲突清单"
    assert conflicts and all("重叠" in c for c in conflicts), "全局清单照旧要给"
    assert raw[0]["conflicts"][0] in conflicts[0], "两份清单应是同一批内容"


def test_search_worker_no_conflict_when_filename_unchanged(tmp_path):
    """文件名没变＝这个文件没有待改项，条件间有重叠也不该报冲突拦整批。"""
    p = tmp_path / "ab.txt"
    p.write_text("x", encoding="utf-8")
    app = _bare_app()
    captured = []
    app._ui = lambda fn, *a: captured.append((fn, a))

    # "ab"→"ab" 与 "a"→"a" 都命中同一区间，但改完名字不变
    app._search_worker([str(p)], [["ab", "ab"], ["a", "a"]], True, True)

    finish = [c for c in captured if c[0] == app._finish_search]
    raw, conflicts = finish[0][1][0], finish[0][1][7]
    assert raw == [], "名字没变就不该进结果列表"
    assert conflicts == [], "没有待改项时不该报冲突（否则会白拦一整批）"


# =====================================================================
# 查找历史：按组展示
# =====================================================================

def test_normalize_history_item_upgrades_legacy_single_record():
    """旧格式的单条记录必须能读进来并升级成组结构，否则已有历史全废。"""
    item = core.normalize_history_item(
        {"find": "a", "replace": "b", "case_sensitive": True,
         "timestamp": "2026-10-14 10:00:00"})
    assert item == {
        "conditions": [["a", "b"]],
        "case_sensitive": True,
        "replace_filename": False,
        "timestamp": "2026-10-14 10:00:00",
    }


def test_normalize_history_item_keeps_group_and_rejects_garbage():
    item = core.normalize_history_item({
        "conditions": [["a", "b"], ["c", "d"]],
        "case_sensitive": False, "replace_filename": True, "timestamp": "t"})
    assert item["conditions"] == [["a", "b"], ["c", "d"]]
    assert item["replace_filename"] is True

    assert core.normalize_history_item("不是字典") is None
    assert core.normalize_history_item({}) is None
    assert core.normalize_history_item({"conditions": []}) is None
    assert core.normalize_history_item({"conditions": [["只有一项"]]}) is None


def test_add_to_history_dedupes_by_whole_group(tk_root, tmp_path, monkeypatch):
    """去重要比整组：只比第一条会把两组真正不同的条件判成同一组。"""
    app = _make_replace_app(tk_root, tmp_path)
    app.find_history = []
    monkeypatch.setattr(app, "save_history", lambda: None)

    app.add_to_history([["a", "b"], ["c", "d"]], False, False)
    app.add_to_history([["a", "b"], ["e", "f"]], False, False)
    assert len(app.find_history) == 2, "第一条相同、第二条不同＝两组不同条件"

    app.add_to_history([["a", "b"], ["c", "d"]], False, False)
    assert len(app.find_history) == 2, "整组（含顺序）完全相同才算同一组"
    assert app.find_history[0]["conditions"] == [["a", "b"], ["c", "d"]], \
        "重复的那组应被提到最前，而不是新增一条"

    app.add_to_history([["a", "b"], ["c", "d"]], False, True)
    assert len(app.find_history) == 3, "作用对象不同也是不同的一组"


def test_add_to_history_limits_groups_and_ignores_empty(tk_root, tmp_path, monkeypatch):
    app = _make_replace_app(tk_root, tmp_path)
    app.find_history = []
    monkeypatch.setattr(app, "save_history", lambda: None)

    for i in range(core.MAX_HISTORY + 5):
        app.add_to_history([[f"k{i}", "v"]], False, False)
    assert len(app.find_history) == core.MAX_HISTORY

    before = list(app.find_history)
    app.add_to_history([], False, False)
    assert app.find_history == before, "空条件不该写进历史"


def test_find_history_naming_matches_what_it_actually_stores():
    """名字必须对齐实际行为：这里只存查找条件，不是"动过哪些文件"的审计轨迹。

    「操作」会被读成"对文件做过什么"，而记录里既没有文件也没有结果，
    只有条件本身。措辞回退到「操作 / 记录」就是缺陷，故在此钉住。
    """
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert core.TEXTS["history"]["zh"] == "查找历史"
    for key in ("history_title", "history_header", "history_note",
                "history_group", "col_cond_group", "m_clear_all_history"):
        assert key in core.TEXTS, "缺少历史对话框文案键：%s" % key
    assert "操作" not in core.TEXTS["history_note"]["zh"], "说明行不该再说「操作」"
    # 按钮在循环里按 key 建，故断言"键与命令配对"而非写死文案
    assert '"history", self.show_history_dialog' in src, "按钮未挂 i18n 键"
    assert 'self._t("history_title"' in src, "历史对话框标题未走 i18n 键"
    assert 'self._t("history_note")' in src, "说明行没挂到对话框上"


def test_history_dialog_groups_conditions():
    """对话框必须按组展示：树形 + 组头加粗 + 组头/条件行两级 iid。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    assert 'show="tree headings"' in src, "按组展示必须用树形（有行首可展开）"
    assert '"group"' in src and "_history_group_font" in src, "组头要有独立样式"
    assert 'cid = f"{key}::{i}"' in src, "条件行 iid 由组键 + 序号构成，不带列表下标"
    assert "history_row_keys(self.find_history)" in src


# =====================================================================
# 首个公开发布版：关于窗口里的发布页入口
#
# 地址写死在源码里是没办法的事（exe 要能离线打开），但必须只写一处：
# 换仓库时漏改任何一处，点了就会打开一个不存在的页面。
# =====================================================================

def test_release_page_url_is_single_source_of_truth():
    """仓库根地址只写一处，发布页 / API / raw 全部从它推导，换仓库不会漏改。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    # 源码里字面出现的 https:// 恰为两处：发布页常量与 https 协议校验。
    # API / raw 两路的仓库根同样来自 RELEASE_PAGE_URL 推导，不另写域名。
    assert src.count("https://github.com/YehOlt/BatchReplace") == 1, \
        "仓库根地址只许写一处"
    assert src.count("https://") == 2, "外部地址只许发布页 / 协议校验两处"
    assert core.RELEASE_PAGE_URL == "https://github.com/YehOlt/BatchReplace/releases"
    assert core.UPDATE_API == "https://api.github.com/repos/YehOlt/BatchReplace/releases/latest"
    assert core.TAGS_API == "https://api.github.com/repos/YehOlt/BatchReplace/tags"
    assert core.UPDATE_MANIFEST == \
        "https://raw.githubusercontent.com/YehOlt/BatchReplace/refs/heads/main/update.json"
    assert core.MANIFEST_API == \
        "https://api.github.com/repos/YehOlt/BatchReplace/contents/update.json?ref=main"


def test_open_url_opens_system_browser(monkeypatch):
    """外链走系统默认浏览器；本机打不开也不该把工具弄崩。"""
    opened = []
    monkeypatch.setattr(core.webbrowser, "open", lambda url: opened.append(url))
    app = object.__new__(core.BatchReplaceApp)   # 只测这一个方法，不建窗口
    app.root = _FakeRoot()
    app._apply_theme("light")
    app._open_url(core.RELEASE_PAGE_URL)
    assert opened == [core.RELEASE_PAGE_URL]

    def boom(_url):
        raise OSError("没有默认浏览器")

    monkeypatch.setattr(core.webbrowser, "open", boom)
    app._open_url(core.RELEASE_PAGE_URL)         # 不该抛出去


def test_about_window_links_are_buttons():
    """关于窗的外链一律是按钮：URL 不当文本显示，点了直接开浏览器。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    about = src[src.index("    def show_about(self):"):src.index("\ndef main():")]
    assert "RELEASE_PAGE_URL" in about, "关于窗口里没有发布页地址"
    assert "_open_url" in about, "外链没走跳转方法"
    assert '"contact"' in about, "关于页没有联系方式"
    assert "messagebox.showinfo" not in about, "系统消息框给不了可点的链接"


# =====================================================================
# i18n：中英双语
#
# 一句话：要么整份做、要么不做——不允许出现"切到英文，界面里还夹着中文"。
# 这一组守四件事：
#   1. 文案表本身完整、中英占位符对齐（否则某语言下会 KeyError 或静默少信息）
#   2. 切换真的换了界面（包括带符号前缀的按钮、Treeview 表头、状态列）
#   3. 引擎层（纯函数）的中文写法与旧版逐字一致（日志 / 测试都按字面断言）
#   4. 源码里没有漏网的写死中文——这条把"下次新增控件忘了走 _t"钉死在 CI 里
# =====================================================================


@pytest.fixture(autouse=True)
def _restore_language():
    """CURRENT_LANG 是模块级全局：任何用例改了它都必须还原。

    语言状态放模块级是有意的（校验、冲突说明这些纯函数在界面类之外，也要出
    文案）。代价是它跨用例存活——一个把界面切到英文的用例会污染后面所有
    用例（它们断言的都是中文），失败会成片出现且完全指不到真凶。用 autouse
    fixture 兜住这个代价。
    """
    prev = core.CURRENT_LANG
    yield
    core.CURRENT_LANG = prev


def test_texts_table_has_both_languages():
    """每条文案都要有 zh 与 en，且都非空。"""
    missing = [(key, lang) for key, entry in core.TEXTS.items()
               for lang in ("zh", "en") if not entry.get(lang)]
    assert missing == [], "文案表缺条目: %r" % (missing,)


def test_texts_placeholders_agree_between_languages():
    """同一 key 的中英模板必须用同一批 {占位符}。

    对不上会出两种事故，而且都不报错：
      - 少一个占位符 → 该语言下 T(key, **kw) 抛 KeyError；
      - 多一个占位符 → str.format 对多余参数静默忽略，界面上就少一段信息。
    """
    bad = []
    for key, entry in core.TEXTS.items():
        zh = set(re.findall(r"\{(\w+)\}", entry.get("zh", "")))
        en = set(re.findall(r"\{(\w+)\}", entry.get("en", "")))
        if zh != en:
            bad.append((key, sorted(zh), sorted(en)))
    assert bad == [], "中英占位符不一致 (key, zh, en): %r" % (bad,)


def test_t_reads_module_level_language():
    """T() 读的是模块级语言，不是某个实例的状态——纯函数区也要出文案。"""
    assert core.T("about") == core.TEXTS["about"]["zh"]
    core.CURRENT_LANG = "en"
    assert core.T("about") == core.TEXTS["about"]["en"]


def test_t_falls_back_to_zh_then_to_key():
    """英文缺条目回退中文；键都不存在时回退成键名，绝不抛异常。

    界面不该因为漏了一条文案就崩——漏文案是编辑疏漏，不是运行时错误。
    """
    core.CURRENT_LANG = "en"
    core.TEXTS["__tmp_zh_only__"] = {"zh": "只有中文"}
    try:
        assert core.T("__tmp_zh_only__") == "只有中文"
    finally:
        core.TEXTS.pop("__tmp_zh_only__", None)
    assert core.T("__no_such_key_at_all__") == "__no_such_key_at_all__"


def test_status_column_text_survives_language_round_trip():
    """状态列存的是显示串：切语言靠反查回键，中↔英来回切都不能丢。

    Treeview 只存文本、没有"键"的位置，所以 _status_text 按两个语言的字面值
    反查。这条用例把"新增状态时两个语言都要能被认出来"钉死。
    """
    app = object.__new__(core.BatchReplaceApp)
    app.root = _FakeRoot()
    app._apply_theme("light")
    for lang in ("zh", "en"):
        core.CURRENT_LANG = lang
        for key in core.BatchReplaceApp._STATUS_KEYS:
            shown = core.TEXTS[key][lang]
            assert app._status_text(shown) == core.TEXTS[key][lang], \
                "状态串认不出: %r（键 %s）" % (shown, key)
    core.CURRENT_LANG = "zh"
    assert app._status_text("") == "", "空状态不该被改写"
    assert app._status_text("我自己填的") == "我自己填的", "未知串应原样返回"


def test_root_bg_follows_theme():
    """root 底色必须随主题重铺：框架间隙（列表四周的 padx）直接透出它，
    只铺属性不铺窗口，切主题后窗口边缘仍留着上一个主题的底色（白边）。"""
    app = object.__new__(core.BatchReplaceApp)
    app.root = _FakeRoot()
    for theme in ("light", "dark"):
        app._apply_theme(theme)
        assert app.root.bg == app.BG, \
            "root 底色没跟着 %s 铺: %r vs %r" % (theme, app.root.bg, app.BG)


def test_window_chrome_self_drawn_follows_theme(tk_root, monkeypatch, tmp_path):
    """标题栏自绘：文字居中、栏底随主题、三个窗口钮真的挂着且命令可用。

    标题居中与栏底随主题是系统标题栏做不到的两件事，故窗口装饰自绘；
    无边框随之带来配套机制（任务栏身份、最小化交还标志、边缘热区），
    窗口按钮命令直接调用验证（最大化 / 还原 / 最小化都会真实改窗口状态）。
    文件列表的滚动条仍自绘——ttk 滚动条的水槽吃不进主题色，深浅主题下
    会露白边。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    monkeypatch.setattr(core, "DEFAULT_BACKUP_DIR", str(tmp_path / "backups"))
    app = core.BatchReplaceApp(tk_root)
    try:
        assert tk_root.overrideredirect() == 1, "窗口应无边框（装饰自绘）"
        bar = app._titlebar_frame
        assert bar is not None and bar.winfo_manager() == "pack", "标题栏未布局"
        # 标题水平居中：place(relx=0.5, anchor=center) 的几何语义
        lbl = app._title_label
        info = lbl.place_info()
        assert lbl.winfo_manager() == "place", "标题未用 place 布局: %r" % lbl.winfo_manager()
        assert float(info["relx"]) == 0.5 and info["anchor"] == "center", \
            "标题 place 参数不是水平居中: %r" % (info,)
        assert lbl.cget("text") == "BatchReplace v%s" % core.VERSION
        # 栏底与标题字色取当前主题色板
        assert bar.cget("bg") == app.BG, "标题栏底色没随主题"
        assert lbl.cget("fg") == app.TXT_DIM, "标题字色没随主题"
        # 恰好三个窗口钮，命令分别挂着
        btns = [c for c in bar.winfo_children()
                if isinstance(c, tk.Button) and getattr(c, "_kind", None)]
        assert len(btns) == 3, "标题栏应恰好三个窗口按钮，实际 %d" % len(btns)
        assert sorted(b._kind for b in btns) == ["close", "max", "min"]
        for b in btns:
            assert b.cget("command"), "窗口按钮没挂命令: %s" % b._kind
        # 窗口钮底色随主题：关闭钮悬停红是常量，其余钮悬停色取 TITLE_HOVER
        hover = [b for b in btns if b._kind == "max"][0]
        hover_bg = hover._img_hover
        assert hover_bg, "悬停符号图没生成"
        # 文件列表滚动条仍自绘
        vsb = [c for c in app.files_tree.master.winfo_children()
               if isinstance(c, core._ZScrollbar)]
        assert len(vsb) == 1, "文件列表应挂一个自绘滚动条"
        assert str(vsb[0].cget("bg")).lower() == str(app.SURFACE).lower(), \
            "水槽色应取主题 SURFACE"
        # 窗口按钮命令可用：最大化 → 还原（真实改窗口几何与状态）
        tk_root.deiconify()
        tk_root.update()
        geo0 = tk_root.geometry()
        app._toggle_maximize()
        tk_root.update()
        assert app._maximized is True, "最大化命令没生效"
        assert tk_root.geometry() != geo0, "最大化后几何没变"
        app._toggle_maximize()
        tk_root.update()
        assert app._maximized is False, "还原命令没生效"
        assert tk_root.geometry() == geo0, "还原后几何没回去"
        # 最小化只验命令不抛（会真的最小化，随即还原）
        app._pre_min_geo = tk_root.geometry()
        app._minimize_window()
        tk_root.update()
        assert tk_root.state() == "iconic", "最小化命令没生效: %r" % tk_root.state()
        tk_root.deiconify()
        tk_root.update()
    finally:
        core.CURRENT_LANG = core.DEFAULT_LANG


def test_dialog_titlebar_follows_app_theme(tk_root, monkeypatch):
    """对话框原生标题栏要跟随应用主题（不是系统主题）。

    对话框保持原生装饰，但原生栏只认系统深色模式：系统深 + 应用浅时
    白界面顶黑栏。_center_dialog 是全部对话框的必经之路，在这里统一
    打 DWM 深浅标志；标志值随 self.theme 翻转，双向都要跟。
    """
    calls = []
    monkeypatch.setattr(core, "_apply_dialog_titlebar_theme",
                        lambda win, dark: calls.append(dark))
    app = object.__new__(core.BatchReplaceApp)
    app.root = tk_root
    dlg = tk.Toplevel(tk_root)
    try:
        app.theme = "dark"
        app._center_dialog(dlg)
        app.theme = "light"
        app._center_dialog(dlg)
        assert calls == [True, False], \
            "对话框标题栏深浅标志没随应用主题翻转: %r" % calls
    finally:
        dlg.destroy()
    # 真函数冒烟：真 Windows 上属性设没设成都不该抛（旧系统不认就放弃）
    dlg2 = tk.Toplevel(tk_root)
    try:
        core._apply_dialog_titlebar_theme(dlg2, True)
        core._apply_dialog_titlebar_theme(dlg2, False)
    finally:
        dlg2.destroy()


def test_engine_location_strings_stay_identical_in_chinese():
    """位置显示串的中文写法与旧版逐字一致。

    日志、文件列表"位置"列、以及一批用例都按这些字面量断言，改一个字就是
    一次无声的界面回归。
    """
    core.CURRENT_LANG = "zh"
    assert core._word_location("paragraph", {"para_idx": 1}) == "第1段"
    assert core._word_location(
        "table", {"table_idx": 2, "row_idx": 3, "col_idx": 4}) == "表格2 第3行第4列"
    assert core._word_location("header", {"section_idx": 1, "para_idx": 2}) \
        == "第1节 页眉 第2段"
    assert core._word_location("footer", {"section_idx": 2, "para_idx": 1}) \
        == "第2节 页脚 第1段"


def test_backup_folder_name_follows_language():
    """备份文件夹名跟着语言走：英文界面下不该在磁盘上留下中文目录名。

    规则目录 / 默认备份目录相反——那两个是固定的数据位置，不随语言变（见
    _scan_chinese 的白名单）。这里守的是"每次替换新建的那个子文件夹"。
    """
    from datetime import datetime
    now = datetime(2026, 10, 10, 15, 4, 5)
    core.CURRENT_LANG = "en"
    assert core.build_backup_folder_name("Old", "New", False, now) \
        == "backup_20261010_Old→New_ContentReplace_150405"
    assert core.build_backup_folder_name("a/b", "", True, now) \
        == "backup_20261010_a_b_FilenameRename_150405"
    core.CURRENT_LANG = "zh"
    assert core.build_backup_folder_name("旧名", "新名", False, now) \
        == "backup_20261010_旧名→新名_内容替换_150405"


def _check_label(app, key):
    """按文案找勾选框的文字控件（自绘勾选框 = Canvas + 独立 Label）。

    只认「父容器第一个孩子是 Canvas」的 Label：界面别处可能出现同一句文案
    （提示条、状态栏），那些没有绑 toggle，取错了会让用例假失败。
    """
    want = app._t(key)

    def is_check_label(c):
        if not isinstance(c, tk.Label) or c.cget("text") != want:
            return False
        kids = c.master.winfo_children()
        return bool(kids) and isinstance(kids[0], tk.Canvas)

    def walk(w):
        for c in w.winfo_children():
            if is_check_label(c):
                return c
            found = walk(c)
            if found is not None:
                return found
        return None

    return walk(app.root)


def test_checkbox_label_click_toggles(tk_root, monkeypatch, tmp_path):
    """点文字也要能勾上，不只是点那个 16px 方框。

    勾选框是自绘的（Canvas 方框 + 独立 Label），文字与方框必须绑同一个
    toggle——只给 Canvas 绑的话，用户点文字会以为"点了没反应"。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    app.save_config = lambda: None
    tk_root.update()          # 控件完成布局后再发点击，接近真实操作
    try:
        # 只验没有副作用的两个开关：「自动备份」关掉时会弹确认框，留给别的用例
        for key, var in (("match_case", app.case_sensitive_var),
                         ("replace_filename_opt", app.replace_filename_var)):
            label = _check_label(app, key)
            assert label is not None, "找不到勾选框文字: %s" % key
            before = bool(var.get())
            label.event_generate("<Button-1>", x=2, y=2, when="now")
            label.event_generate("<ButtonRelease-1>", x=2, y=2, when="now")
            assert bool(var.get()) is (not before), \
                "点文字没切换勾选状态: %s" % key
    finally:
        core.CURRENT_LANG = core.DEFAULT_LANG


def test_checkbox_repaints_when_variable_set_programmatically(tk_root, monkeypatch, tmp_path):
    """外部直接 set 变量（加载规则 / 恢复状态）时方框要跟着重绘。

    重绘由变量上的 trace 驱动，这里分开验两件事：trace 确实挂在变量上、
    重绘闭包本身能换图。直接驱动闭包是为了不依赖 Tk 事件循环的时序——
    trace 到重绘之间的派发顺序在连续建窗的测试进程里并不稳定。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    app.save_config = lambda: None
    tk_root.update()
    try:
        label = _check_label(app, "match_case")
        canvas = label.master.winfo_children()[0]
        app.case_sensitive_var.set(True)       # 变量被程序改动，不经过点击
        label.master._br_repaint()
        # 断言对象身份而不是图元里的图像名：Tk 的图像计数随解释器重建从头开始，
        # 不同对象完全可能同名，比名字会把"已重绘"误判成"没变"。
        assert canvas._img is app._check_image(True), "方框没换成选中图标"
        assert app.case_sensitive_var.trace_info(), \
            "变量没挂重绘回调，程序改动后界面不会刷新"
    finally:
        core.CURRENT_LANG = core.DEFAULT_LANG


def test_switch_lang_retranslates_widgets_without_persisting(tk_root, monkeypatch, tmp_path):
    """点语言按钮：界面真的换语言，但不写进配置（下次启动回系统语言）。

    覆盖四种容易漏掉的形态：带符号前缀的主按钮、语言按钮自身、Treeview 表头
    （不能走 config(text=)）、以及参数区随模式变化的标签。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    try:
        assert app.lang == "zh", "全新配置应以中文启动"
        assert _btn_text(app.replace_btn) == "\u25b6  " + core.TEXTS["start_search"]["zh"]
        assert _btn_text(app.lang_btn) == core.TEXTS["lang_to_en"]["zh"]
        assert app.find_label.cget("text") == core.TEXTS["find_content"]["zh"]
        assert app.files_tree.heading("status")["text"] == core.TEXTS["col_status"]["zh"]

        app.switch_lang()

        assert app.lang == "en"
        # ▶ 前缀必须还在：建控件时写了前缀，刷新时只写文案就会把它抹掉
        assert _btn_text(app.replace_btn) == "\u25b6  " + core.TEXTS["start_search"]["en"]
        assert _btn_text(app.add_files_btn) == core.TEXTS["add_files"]["en"]
        assert _btn_text(app.lang_btn) == core.TEXTS["lang_to_zh"]["en"], \
            "语言按钮要指向「切过去会变成哪种语言」"
        assert app.find_label.cget("text") == core.TEXTS["find_content"]["en"]
        assert app.files_tree.heading("status")["text"] == core.TEXTS["col_status"]["en"]
        assert app.files_tree.heading("#0")["text"] == core.TEXTS["col_name"]["en"]
        assert app.log_arrow.cget("text").endswith(core.TEXTS["process_log"]["en"])

        saved = json.loads((tmp_path / "cfg.json").read_text(encoding="utf-8"))
        assert "lang" not in saved, "语言不该落盘：下次启动要回系统语言"

        app.switch_lang()
        assert app.lang == "zh"
        assert _btn_text(app.replace_btn) == "\u25b6  " + core.TEXTS["start_search"]["zh"]
    finally:
        core.CURRENT_LANG = core.DEFAULT_LANG


def test_switch_language_after_closing_preview_does_not_crash(tk_root, monkeypatch, tmp_path):
    """预览窗关掉之后再切语言，不能崩。

    预览窗把它内部的 Label 挂在实例上（`checked_count_label` 等），窗一关这些
    引用就成了死引用；而 `apply_lang` 会把动态文案重刷一遍（「已选: N 项」经
    `update_checked_count`），碰到死引用即抛 `TclError: invalid command name`。
    用户实测过这条路径：「如果开始搜索过，再切换英文就会报错」。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    core.CURRENT_LANG = "zh"
    app = core.BatchReplaceApp(tk_root)
    app.preview_data = []          # 要的是「预览窗建过又关掉」，结果为空也够
    app.create_preview_window([["a", "b"]], 0, False)
    tk_root.update()

    preview = app.preview_window
    assert preview is not None, "预览窗要记在实例上，关闭时才清得掉引用"
    preview.destroy()
    tk_root.update()

    # 关窗后这些引用必须已清空 —— 各处原有的 is not None 守卫才会生效
    assert app.preview_window is None
    assert app.checked_count_label is None
    assert app.conflict_label is None
    assert app.preview_hint_label is None

    app.switch_lang()              # 修复前这里抛 TclError: invalid command name
    assert app.lang == "en"


def test_lang_from_lcid_maps_chinese_family_to_zh():
    """LCID → 语言：中文系（简 / 繁 / 港 / 新）出 zh，其余语言出 en。"""
    for lcid in (0x0804, 0x0404, 0x0C04, 0x1004):
        assert core.lang_from_lcid(lcid) == "zh", hex(lcid)
    for lcid in (0x0409, 0x0411, 0x0407, 0x0419):
        assert core.lang_from_lcid(lcid) == "en", hex(lcid)


def test_detect_default_lang_follows_system_ui_language(monkeypatch):
    """默认语言跟着系统界面语言走，不能在源码里写死。

    目标：中文系统打开就是中文，非中文系统打开就是英文。
    """
    monkeypatch.setattr(core, "_system_ui_lcid", lambda: 0x0804)
    assert core.detect_default_lang() == "zh"
    monkeypatch.setattr(core, "_system_ui_lcid", lambda: 0x0409)
    assert core.detect_default_lang() == "en"

    def boom():
        raise OSError("api 不可用")

    monkeypatch.setattr(core, "_system_ui_lcid", boom)
    assert core.detect_default_lang() in ("zh", "en"), \
        "取不到系统语言时要给一个受支持的答案，而不是崩"


def test_migrate_backup_path_replaces_legacy_chinese_default():
    """旧配置里的中文默认备份目录，要迁移成新的英文默认目录。

    这个值会被关窗时的保存动作原样写回配置文件，不迁移就等于永久锁死——
    界面上看着像"改名没生效"。

    正斜杠与反斜杠两种写法都得认：tkinter 的目录选择框返回的是正斜杠路径，
    而源码拼出来的是反斜杠路径，不归一化就会漏掉一半。
    """
    legacy = core.LEGACY_DEFAULT_BACKUP_DIR
    assert core.migrate_backup_path(legacy) == core.DEFAULT_BACKUP_DIR
    assert core.migrate_backup_path(legacy.replace("\\", "/")) == core.DEFAULT_BACKUP_DIR

    custom = "D:/我自己的备份目录"
    assert core.migrate_backup_path(custom) == custom, "用户自选的路径不许动"
    assert core.migrate_backup_path(core.DEFAULT_BACKUP_DIR) == core.DEFAULT_BACKUP_DIR


def test_load_config_migrates_legacy_backup_path(tk_root, monkeypatch, tmp_path):
    """端到端：配置里存着旧的中文默认值，启动后界面显示新英文路径。"""
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({
        "backup_path": core.LEGACY_DEFAULT_BACKUP_DIR.replace("\\", "/"),
        "case_sensitive": False,
        "lang": "zh",
    }, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(core, "CONFIG_FILE", str(cfg))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = _bare_app()
    app.startup_notes = []

    app.load_config()

    assert app.backup_path.get() == core.DEFAULT_BACKUP_DIR, \
        "老配置里的中文默认目录必须迁到新默认目录"


def test_startup_follows_system_theme_and_lang(tk_root, monkeypatch, tmp_path):
    """配置里存着上次的语言 / 主题也不读：每次启动都取系统当前值。"""
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"lang": "en", "theme": "dark"}),
                   encoding="utf-8")
    monkeypatch.setattr(core, "CONFIG_FILE", str(cfg))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    monkeypatch.setattr(core, "DEFAULT_LANG", "zh")
    monkeypatch.setattr(core, "detect_default_theme", lambda: "light")
    app = _bare_app()
    app.startup_notes = []

    app.load_config()

    assert app.lang == "zh", "语言应跟系统走，不许被配置里的 en 拉走"
    assert app.theme == "light", "主题应跟系统走，不许被配置里的 dark 拉走"


def test_detect_default_theme_matches_registry():
    """主题探测与注册表一致：AppsUseLightTheme 为 0 时出深色。"""
    theme = core.detect_default_theme()
    assert theme in ("light", "dark")
    if os.name == "nt":
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            light = winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
        assert theme == ("light" if light else "dark")


def test_load_config_creates_default_backup_dir(tk_root, monkeypatch, tmp_path):
    """默认备份目录不存在时，启动要把它建出来。

    目录名改英文后，已有机器上并没有这个目录，点「打开备份文件夹」会被
    "路径不存在"拦下——首次启动就该建好，而不是等用户撞上弹窗。
    """
    backup_dir = tmp_path / "backups"
    monkeypatch.setattr(core, "DEFAULT_BACKUP_DIR", str(backup_dir))
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = _bare_app()
    app.startup_notes = []

    app.load_config()

    assert backup_dir.is_dir(), "默认备份目录应被自动创建"


def test_load_config_never_creates_custom_backup_dir(tk_root, monkeypatch, tmp_path):
    """用户自选的备份路径不存在时保持原样提示，不替用户在任意位置建目录。"""
    custom = tmp_path / "自选目录"
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({
        "backup_path": str(custom), "case_sensitive": False, "lang": "zh",
    }, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(core, "CONFIG_FILE", str(cfg))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = _bare_app()
    app.startup_notes = []

    app.load_config()

    assert app.backup_path.get() == str(custom), "自选路径必须原样加载"
    assert not custom.exists(), "自选路径不存在时不该替用户创建"


# 有意保留的中文串：
# - 字体名，以及旧版默认备份目录的名字——迁移老配置时必须认得这个原文。
# - 一键更新模块的校验/下载日志与错误提示：这些提示写死中文
#   （对外只经 log() 输出，不进可切换语言的界面控件）。
_ALLOWED_CHINESE_LITERALS = {
    "微软雅黑",
    "批量替换备份",
    "检查更新：正在获取版本清单…",
    "已启动新版本 %s；新版文件已下载到同一目录。",
    "清单必须是 JSON 对象",
    "清单 version 非法",
    "清单 exe 必须以 .exe 结尾",
    "清单 exe 不得含路径分隔或 ..",
    "清单 exe 文件名必须含可解析版本号",
    "清单 sha256 必须是 64 位十六进制",
    "清单缺少 url",
    "清单 size 必须是非负整数",
    "清单缺少下载地址",
    "下载地址必须是 https",
    "失败：%s",
    "当前已是最新版 v%s。",
    "发现新版本 %s（当前 v%s），开始下载…",
    "下载完成并通过 SHA-256 校验：%s",
    "试运行完成：已下载并校验 %s（未启动）。",
    "失败：新版已下载并校验，但启动失败：%s（文件：%s）",
    "清单异常大",
    "SHA-256 校验不符",
    "大小与清单不符（收到 %d，清单 %s）",
    "文件超过大小上限",
    "下载内容不是 exe（缺少 MZ 魔数）",
}


def _chinese_string_literals(src):
    """源码里所有含中文的字符串字面量 → [(行号, 内容)]。

    跳过 TEXTS 文案表本身（它就是中文的存放处）与 docstring（注释性文字）。
    用 AST 而不是正则：注释天然不是字面量，也就不会被误报。
    """
    tree = ast.parse(src)
    texts_range = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == "TEXTS" for t in node.targets):
            texts_range = (node.lineno, node.end_lineno)
            break

    doc_ranges = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef,
                             ast.AsyncFunctionDef)) and node.body:
            first = node.body[0]
            if (isinstance(first, ast.Expr) and isinstance(first.value, ast.Constant)
                    and isinstance(first.value.value, str)):
                doc_ranges.append((first.value.lineno,
                                   first.value.end_lineno or first.value.lineno))

    out = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Constant) and isinstance(node.value, str)):
            continue
        if not re.search(r"[\u4e00-\u9fff]", node.value):
            continue
        if texts_range and texts_range[0] <= node.lineno <= texts_range[1]:
            continue
        if any(s <= node.lineno <= e for s, e in doc_ranges):
            continue
        out.append((node.lineno, node.value))
    return out


def test_no_hardcoded_chinese_left_in_source():
    """源码里不该再有写死的界面中文（文案表、docstring、注释除外）。

    写死的中文绕过了 i18n：切到英文时它不会变，界面就出现"英文夹中文"。这条
    用例把"新增控件忘了走 _t"变成一个立刻失败的测试，而不是等用户截图发现。
    """
    leftovers = [(line, text)
                 for line, text in _chinese_string_literals(UI_SOURCE.read_text(encoding="utf-8"))
                 if text not in _ALLOWED_CHINESE_LITERALS]
    assert leftovers == [], "源码里还有写死的界面中文（应改走 TEXTS）: %r" % (leftovers,)


# =====================================================================
# 链式替换：上一条的输出是下一条的输入
# =====================================================================

def test_run_replace_chain_feeds_each_round_into_the_next(tmp_path):
    """链式：第 1 条改完的结果，是第 2 条要找的文本。

    对照是并行语义下同样条件的结果——两条语义必须不同，否则这个功能不存在。
    """
    chained = make_txt(tmp_path / "chain.txt", "X 甲 X")
    parallel = make_txt(tmp_path / "para.txt", "X 甲 X")
    conditions = [["X", "Y"], ["Y", "Z"]]

    rounds, skipped = core.run_replace_chain(str(chained), conditions, True)
    assert [len(r["matches"]) for r in rounds] == [2, 2], "每轮的命中数: %r" % rounds
    assert skipped == []
    core.run_replace_chain(str(chained), conditions, True, commit=True)
    assert chained.read_text(encoding="utf-8") == "Z 甲 Z"

    # 并行：两条条件都在原文上找，第 2 条在原文里根本没命中
    matches = []
    for i, (find_text, replace_text) in enumerate(conditions, 1):
        found, _ = core.extract_text_matches(str(parallel), find_text, True)
        for m in found:
            m["replace_text"] = replace_text
            m["cond_idx"] = i
        matches += found
    core.replace_text_multi(str(parallel), matches, True)
    assert parallel.read_text(encoding="utf-8") == "Y 甲 Y"


def test_run_replace_chain_marks_text_made_by_previous_round(tmp_path):
    """from_prev 必须只标"上一轮改出来的"，原文里本来就有的不能误标。

    第 2 段原本就写着「乙公司」，与第 1 轮改出来的那个「乙公司」同形，
    只有按区间追踪才分得开。
    """
    import docx
    p = tmp_path / "b.docx"
    doc = docx.Document()
    doc.add_paragraph("甲公司在杭州")
    doc.add_paragraph("乙公司")
    doc.save(str(p))

    rounds, _ = core.run_replace_chain(
        str(p), [["甲公司", "乙公司"], ["乙公司", "丙公司"]], True)
    first, second = rounds
    assert [m["from_prev"] for m in first["matches"]] == [False], "第 1 轮不可能来自上一轮"
    assert [m["from_prev"] for m in second["matches"]] == [True, False], \
        "第 2 段是原文就有的，不该被标成上一轮产出"


def test_run_replace_chain_excel_multi_occurrence_cell(tmp_path):
    """Excel 一条命中含同一单元格内的多处出现：区间口径要与替换引擎一致。"""
    import openpyxl
    p = tmp_path / "c.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "甲公司 甲公司"
    wb.save(str(p))

    rounds, _ = core.run_replace_chain(
        str(p), [["甲公司", "乙公司"], ["乙公司", "丙公司"]], True)
    assert [m["from_prev"] for m in rounds[0]["matches"]] == [False]
    assert [m["from_prev"] for m in rounds[1]["matches"]] == [True]
    core.run_replace_chain(str(p), [["甲公司", "乙公司"], ["乙公司", "丙公司"]],
                           True, commit=True)
    assert openpyxl.load_workbook(str(p)).active["A1"].value == "丙公司 丙公司"


def test_run_replace_chain_sampling_does_not_touch_the_file(tmp_path):
    """预览取样（commit=False）绝不能改动原文件——预览是"看"，不是"做"。"""
    p = make_txt(tmp_path / "s.txt", "X")
    before = p.read_bytes()
    rounds, _ = core.run_replace_chain(str(p), [["X", "Y"], ["Y", "Z"]], True)
    assert [len(r["matches"]) for r in rounds] == [1, 1]
    assert p.read_bytes() == before, "取样把文件改了"


def test_run_replace_chain_leaves_original_untouched_when_a_round_fails(tmp_path, monkeypatch):
    """链式半途出错必须退化成"什么都没改"，而不是留一个改到一半的文件。"""
    p = make_txt(tmp_path / "x.txt", "X")
    before = p.read_text(encoding="utf-8")
    real = core.extract_text_matches
    seen = {"n": 0}

    def flaky(path, find_text, case_sensitive):
        seen["n"] += 1
        if seen["n"] == 2:
            raise RuntimeError("boom")
        return real(path, find_text, case_sensitive)

    monkeypatch.setattr(core, "extract_text_matches", flaky)
    with pytest.raises(RuntimeError):
        core.run_replace_chain(str(p), [["X", "Y"], ["Y", "Z"]], True, commit=True)
    assert p.read_text(encoding="utf-8") == before


def test_run_filename_chain_applies_conditions_in_order():
    """文件名链式：每条在上一条的输出上找，命中同样标出上一轮产出。"""
    rounds, final = core.run_filename_chain(
        "甲公司_合同.docx", [["甲公司", "乙公司"], ["乙公司", "丙公司"]], True)
    assert final == "丙公司_合同.docx"
    assert [len(r["matches"]) for r in rounds] == [1, 1]
    assert [m["from_prev"] for m in rounds[0]["matches"]] == [False]
    assert [m["from_prev"] for m in rounds[1]["matches"]] == [True]




# =====================================================================
# 多条件执行方式：并行（默认） / 链式
#
# 二者区别在于"在谁的身上找"：并行全都对着原文，链式让上一条的输出成为
# 下一条的输入。名字一旦退回「操作 / 记录」就会被读成审计轨迹，故钉住措辞。
# =====================================================================

def test_multi_mode_combo_appears_only_with_multi_condition(tk_root, tmp_path, monkeypatch):
    """下拉框：勾选多条件替换才显现，且默认是并行——单条件没有"链"可言。"""
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    tk_root.update()

    # 断言布局管理器而不是 winfo_ismapped：tk_root 夹是隐藏窗口，子控件
    # 一律 ismapped=0；而这里要防的正是"控件建了没布局"
    def visible():
        return app.multi_mode_combo.winfo_manager() == "pack"

    assert not app.multi_condition_var.get(), "默认不该是多条件"
    assert not visible(), "没勾多条件时下拉框不该出现"
    app.multi_condition_var.set(True)
    app.toggle_multi_condition()
    tk_root.update()
    assert visible(), "勾了多条件才出现下拉框"
    assert app.multi_mode_var.get() == core.TEXTS["multi_mode_parallel"]["zh"], \
        "默认必须是并行: %r" % app.multi_mode_var.get()
    assert app._multi_mode == "parallel"

    app.multi_condition_var.set(False)
    app.toggle_multi_condition()
    tk_root.update()
    assert not visible(), "取消多条件后下拉框要跟着收起"


def test_mode_tip_pops_on_switch_and_closes_manually(tk_root, tmp_path, monkeypatch):
    """执行方式说明不靠悬停：切换方式 / 刚勾多条件时自动摆出，点 × 才收。

    摆的是四段：标题（方式名，加粗）→ 说明 → 示例 → 补充，逐段比对。
    换语言时模式没变，不该凭空弹窗（主题恢复同理）。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    tk_root.update()

    def tip_parts():
        assert app._chain_tip is not None, "说明未弹出"
        parts = []

        def walk(w):
            for c in w.winfo_children():
                if isinstance(c, tk.Label) and not getattr(c, "_tip_closer", False):
                    parts.append(c.cget("text"))
                walk(c)

        walk(app._chain_tip)
        return parts

    def want(mode):
        return [core.TEXTS["multi_mode_chain" if mode == "chain"
                           else "multi_mode_parallel"][core.CURRENT_LANG],
                core.TEXTS["multi_mode_tip_%s" % mode][core.CURRENT_LANG],
                core.TEXTS["multi_mode_tip_%s_example" % mode][core.CURRENT_LANG],
                core.TEXTS["multi_mode_tip_%s_note" % mode][core.CURRENT_LANG]]

    def closer_widget():
        """标题行右侧的 × 关闭钮。"""
        head = [c for c in app._chain_tip.winfo_children()[0].winfo_children()[0]
                .winfo_children() if isinstance(c, tk.Frame)]
        assert head, "说明里没有标题行"
        return [c for c in head[0].winfo_children()
                if isinstance(c, tk.Label)][-1]

    try:
        # ① 刚勾上「多条件替换」：默认并行，说明直接摆在眼前
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        tk_root.update()
        assert tip_parts() == want("parallel"), \
            "勾上多条件应弹出并行四段: %r" % (tip_parts(),)

        # ② 点 × 手动收
        closer_widget().event_generate("<Button-1>")
        tk_root.update()
        assert app._chain_tip is None, "点 × 后说明未收"

        # ③ 切到链式：自动换文重摆（四段全是链式的）
        app.multi_mode_var.set(core.TEXTS["multi_mode_chain"][core.CURRENT_LANG])
        tk_root.update()
        assert app._multi_mode == "chain"
        assert tip_parts() == want("chain"), \
            "切到链式应弹出链式四段: %r" % (tip_parts(),)

        # ④ 切回并行：再摆一次并行的
        app.multi_mode_var.set(core.TEXTS["multi_mode_parallel"][core.CURRENT_LANG])
        tk_root.update()
        assert tip_parts() == want("parallel"), \
            "切回并行应弹出并行四段: %r" % (tip_parts(),)

        # ⑤ 取消多条件：下拉框收起，说明跟着收
        app.multi_condition_var.set(False)
        app.toggle_multi_condition()
        tk_root.update()
        assert app._chain_tip is None, "取消多条件后说明未收"

        # ⑥ 换语言（模式没变）：不弹
        app.switch_lang()
        tk_root.update()
        assert app._chain_tip is None, "换语言不该弹说明"
    finally:
        app._hide_chain_tip()
        core.CURRENT_LANG = core.DEFAULT_LANG


def test_multi_mode_tip_words_avoid_audit_trail_reading():
    """两种模式的悬停说明必须说的是"在哪段文本上找"，不能是"操作 / 记录"。"""
    for key in ("multi_mode_parallel", "multi_mode_chain",
                "multi_mode_tip_parallel", "multi_mode_tip_parallel_example",
                "multi_mode_tip_parallel_note",
                "multi_mode_tip_chain", "multi_mode_tip_chain_example",
                "multi_mode_tip_chain_note"):
        assert key in core.TEXTS, "缺少模式文案键：%s" % key
    for lang in ("zh", "en"):
        intro = core.TEXTS["multi_mode_tip_parallel"][lang]
        assert "原始文本" in intro or "original" in intro, \
            "并行说明要说到只作用于原始文本: %r" % intro
        note = core.TEXTS["multi_mode_tip_parallel_note"][lang]
        assert ("不会被后面的规则再命中" in note
                or "never re-matched" in note), \
            "并行补充要说刚改出的内容不再被后面的规则命中: %r" % note
        intro = core.TEXTS["multi_mode_tip_chain"][lang]
        assert ("上一" in intro or "顺序" in intro
                or "previous" in intro or "order" in intro.lower()), \
            "链式说明要说到上一条的输出 / 顺序: %r" % intro
        note = core.TEXTS["multi_mode_tip_chain_note"][lang]
        assert ("接力" in note or "chain" in note.lower()), \
            "链式补充要说到交叉规则需要接力: %r" % note


def test_chain_preview_groups_matches_into_rounds(tk_root, tmp_path, monkeypatch):
    """链式预览按轮次分组：轮次是分组标题，命中行的勾选用不了（行级禁用）。

    勾选只按整个文件生效——轮次连在一起，只替换其中一轮没有意义。
    """
    monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
    monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
    monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
    app = core.BatchReplaceApp(tk_root)
    app.is_chain_mode = True
    app.preview_data = [{
        "file_path": "/tmp/x.txt", "file_name": "x.txt",
        "checked": tk.BooleanVar(value=True),
        "matches": [
            {"type": "text", "location": "第1行, 第1列", "context": "X",
             "round": 1, "from_prev": False},
            {"type": "text", "location": "第1行, 第1列", "context": "Y",
             "round": 2, "from_prev": True},
        ],
        "rounds": [
            {"cond_idx": 1, "find_text": "X", "replace_text": "Y",
             "matches": [{"type": "text", "location": "第1行, 第1列", "context": "X",
                          "round": 1, "from_prev": False}]},
            {"cond_idx": 2, "find_text": "Y", "replace_text": "Z",
             "matches": [{"type": "text", "location": "第1行, 第1列", "context": "Y",
                          "round": 2, "from_prev": True}]},
        ],
    }]
    app.create_preview_window([["X", "Y"], ["Y", "Z"]], 2, False)
    tk_root.update()
    tree = app.preview_tree

    rounds = tree.get_children("file_0")
    assert len(rounds) == 2, "两个轮次各一行: %r" % rounds
    assert "第 1 轮" in tree.item(rounds[0], "values")[0], \
        "轮次行要说清是第几轮: %r" % tree.item(rounds[0], "values")[0]
    # 命中挂在各自轮次下，位置列还要标出"上一轮产出"
    hit = tree.get_children(rounds[1])[0]
    assert "上一轮产出" in tree.item(hit, "values")[2], \
        "链式产出要标出来: %r" % tree.item(hit, "values")[2]
    assert "produced" in tree.item(hit, "tags"), "链式产出要整行标色"

    # 行级勾选停用：点击不改状态，也不改变计数
    before = app.checked_count_label.cget("text")
    app.on_tree_click(type("E", (), {"x": 1, "y": 10})(), tree)
    app.toggle_match_check(hit, tree)
    assert app.checked_count_label.cget("text") == before, "链式不该有行级勾选"

    tree.winfo_toplevel().destroy()


# =====================================================================
# 回溯：manifest · 扫描 · 回滚
# =====================================================================

def _mk_manifest_entry(base_dir, ts, files, chain=False, multi=False,
                       conditions=(("甲", "乙"),)):
    """造一条回溯记录，返回备份夹路径。

    存在的源文件会照替换前的样子复制进备份夹——回溯恢复的就是这些副本。
    """
    folder = core.create_backup_folder(str(base_dir), "甲", "乙", False)
    for p in files:
        if os.path.isfile(p):
            shutil.copy2(p, os.path.join(folder, os.path.basename(p)))
    core.write_manifest(folder, {
        "version": core.MANIFEST_VERSION, "ts": ts, "mode": "content",
        "chain": chain, "multi": multi, "conditions": [list(c) for c in conditions],
        "case_sensitive": True,
        "files": [{"orig": p, "orig_canon": core.canonical_path(p),
                   "backup_rel": os.path.basename(p)} for p in files],
    })
    return folder


def test_manifest_roundtrip_and_atomic_write(tmp_path):
    """manifest 写入后能原样读回，且不残留临时文件。"""
    folder = _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00",
                                [str(tmp_path / "a.txt")])
    data = core.read_manifest(folder)
    assert data["ts"] == "2026-10-16 10:00:00"
    assert data["mode"] == "content"
    assert data["files"][0]["backup_rel"] == "a.txt"
    assert os.listdir(folder) == ["manifest.json"], "原子写不许留 .tmp 残留"


def test_scan_backups_orders_desc_and_excludes_filename_mode(tmp_path):
    """扫描按时间倒序；文件名替换模式（无 manifest）的备份夹被排除。"""
    _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00", [str(tmp_path / "a.txt")])
    _mk_manifest_entry(tmp_path, "2026-10-16 12:00:00", [str(tmp_path / "a.txt")])
    _mk_manifest_entry(tmp_path, "2026-10-16 09:00:00", [str(tmp_path / "b.txt")])
    os.makedirs(tmp_path / "backup_rename_only", exist_ok=True)  # 无 manifest

    entries, problems = core.scan_backups(str(tmp_path))
    assert len(entries) == 3, "文件名替换夹应被排除"
    assert [e["ts"] for e in entries] == ["2026-10-16 12:00:00",
                                          "2026-10-16 10:00:00",
                                          "2026-10-16 09:00:00"], "必须倒序"
    assert problems == []


def test_scan_backups_skips_corrupt_manifest(tmp_path):
    """manifest 损坏只跳过那一夹，其余记录照常读出。"""
    _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00", [str(tmp_path / "a.txt")])
    bad = tmp_path / "bad"
    bad.mkdir()
    (bad / core.MANIFEST_NAME).write_text("{oops", encoding="utf-8")

    entries, problems = core.scan_backups(str(tmp_path))
    assert len(entries) == 1
    assert len(problems) == 1, "坏夹要报出来，不能静默吞掉"


def test_detect_later_operations_finds_only_newer_runs(tmp_path):
    """同一文件在更晚操作里也改过，才需要警告；更早的不算。"""
    early = _mk_manifest_entry(tmp_path, "2026-10-16 09:00:00",
                               [str(tmp_path / "a.txt")])
    late = _mk_manifest_entry(tmp_path, "2026-10-16 12:00:00",
                              [str(tmp_path / "a.txt"), str(tmp_path / "b.txt")])
    entries, _ = core.scan_backups(str(tmp_path))
    by_folder = {e["folder"]: e for e in entries}
    canon = core.canonical_path(str(tmp_path / "a.txt"))

    later = core.detect_later_operations(entries, by_folder[early], canon)
    assert [e["folder"] for e in later] == [late], "只有 12:00 那次更晚"
    later2 = core.detect_later_operations(entries, by_folder[late], canon)
    assert later2 == [], "最新一次没有更晚操作"


def test_rollback_restores_and_snapshots_current_file(tmp_path):
    """回滚把文件恢复成替换前；被覆盖的当前文件先进 rollback_ 快照。"""
    target = tmp_path / "work" / "a.txt"
    target.parent.mkdir()
    target.write_text("甲公司", encoding="utf-8")
    folder = _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00", [str(target)])
    # 备份夹里的就是替换前原文；模拟替换已发生
    target.write_text("乙公司", encoding="utf-8")
    entries, _ = core.scan_backups(str(tmp_path))

    ok, bad = core.rollback_files(entries[0], entries[0]["files"])
    assert bad == []
    assert target.read_text(encoding="utf-8") == "甲公司", "回滚后应是替换前内容"
    snap_dir = os.path.join(folder, "rollback")
    snaps = os.listdir(snap_dir)
    assert snaps == ["a.txt"], "当前文件应先快照再覆盖"
    assert (Path(snap_dir) / "a.txt").read_text(encoding="utf-8") == "乙公司", \
        "快照存的是回滚前的当前文件，误回滚也能再捞回来"


def test_rollback_reports_missing_backup_without_touching_file(tmp_path):
    """备份文件缺失：报错、不动当前文件、不留临时残留。"""
    target = tmp_path / "a.txt"
    target.write_text("乙公司", encoding="utf-8")
    folder = core.create_backup_folder(str(tmp_path), "甲", "乙", False)
    core.write_manifest(folder, {
        "version": core.MANIFEST_VERSION, "ts": "2026-10-16 10:00:00",
        "mode": "content", "chain": False, "multi": False,
        "conditions": [["甲", "乙"]], "case_sensitive": True,
        "files": [{"orig": str(target), "orig_canon": str(target),
                   "backup_rel": "gone.txt"}],
    })
    entries, _ = core.scan_backups(str(tmp_path))
    ok, bad = core.rollback_files(entries[0], entries[0]["files"])
    assert ok == [] and len(bad) == 1
    assert target.read_text(encoding="utf-8") == "乙公司", "失败也不许动文件"
    assert not list(tmp_path.rglob("*.rollback_tmp")), "不许留临时文件"


def test_rollback_snapshot_name_collides_safely(tmp_path):
    """同一文件回滚两次：第二次的快照不许覆盖第一次的（unique_path 生效）。"""
    target = tmp_path / "a.txt"
    target.write_text("v0", encoding="utf-8")
    folder = _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00", [str(target)])
    entries, _ = core.scan_backups(str(tmp_path))

    target.write_text("v1", encoding="utf-8")
    core.rollback_files(entries[0], entries[0]["files"])
    target.write_text("v2", encoding="utf-8")
    core.rollback_files(entries[0], entries[0]["files"])

    snaps = sorted(os.listdir(os.path.join(folder, "rollback")))
    assert snaps == ["a.txt", "a_1.txt"], "第二次快照应加后缀而不是覆盖: %r" % snaps
    contents = sorted((Path(folder) / "rollback" / s).read_text(encoding="utf-8")
                      for s in snaps)
    assert contents == ["v1", "v2"], "两份快照都要留下"


def test_is_cloud_placeholder_returns_bool_for_missing_file(tmp_path):
    """占位判定对不存在的文件返回 False，不炸。"""
    assert core.is_cloud_placeholder(str(tmp_path / "nope.txt")) is False


# =====================================================================
# 条件表拖拽调序
# =====================================================================

def test_condition_drag_reorders_without_dropping_duplicates(tmp_path, monkeypatch):
    """拖拽只改顺序；即使列表里有重复查找词，也不能被顺手剔掉。"""
    root = tk.Tk()
    root.deiconify()   # bbox 要真实绘制的行，withdraw 下取不到
    root.geometry("900x600")
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        # 故意放两条查找词相同——拖拽路径不该触发去重校验
        app.conditions = [["甲", "A"], ["甲", "B"], ["乙", "C"]]
        app._refresh_condition_tree()
        root.update()

        class E:
            def __init__(self, y):
                self.y = y

        tree = app.cond_tree
        y_of = lambda iid: (lambda b: b[1] + b[3] // 2)(tree.bbox(iid, "find"))
        # 第 3 行拖到第 1 行
        app._cond_drag_press(E(y_of(tree.get_children()[2])))
        app._cond_drag_motion(E(y_of(tree.get_children()[0])))
        app._cond_drag_release(E(y_of(tree.get_children()[0])))
        root.update()
        assert app.conditions == [["乙", "C"], ["甲", "A"], ["甲", "B"]], \
            "顺序变了、三条都还在（重复项不被剔除）"
        assert tree.selection() == ("0",), "拖动后选中跟着目标行"
    finally:
        root.destroy()


def test_condition_drag_ignores_click_outside_rows(tmp_path, monkeypatch):
    """点在行外空白处不启动拖拽， Motion 来了也不许改顺序。"""
    root = tk.Tk()
    root.deiconify()
    root.geometry("900x600")
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        app.conditions = [["甲", "A"], ["乙", "B"]]
        app._refresh_condition_tree()
        root.update()

        class E:
            def __init__(self, y):
                self.y = y

        app._cond_drag_press(E(-999))          # 空白处
        app._cond_drag_motion(E(0))            # 随便动
        app._cond_drag_release(E(0))
        root.update()
        assert app.conditions == [["甲", "A"], ["乙", "B"]]
    finally:
        root.destroy()


# =====================================================================
# 条件表右键菜单：调序 / 插入
# =====================================================================

def test_cond_context_menu_move_up_and_to_bottom(tmp_path, monkeypatch):
    """右键「上移」「移至底部」改顺序并选中跟着移动，且边界不动。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        app.conditions = [["a", "1"], ["b", "2"], ["c", "3"], ["d", "4"]]
        app._refresh_condition_tree()

        app.cond_tree.selection_set("1")
        app._cond_move("up")
        assert app.conditions == [["b", "2"], ["a", "1"], ["c", "3"], ["d", "4"]]
        assert app.cond_tree.selection() == ("0",), "上移后选中跟着目标行"

        app.cond_tree.selection_set("0")
        app._cond_move("bottom")
        assert app.conditions[-1] == ["b", "2"]
        assert app.cond_tree.selection() == (str(len(app.conditions) - 1),)

        # 已在最底再「移至底部」不许越界、不许改顺序
        app._cond_move("bottom")
        assert app.conditions[-1] == ["b", "2"]
    finally:
        root.destroy()


def test_cond_context_menu_insert_above_selected(tmp_path, monkeypatch):
    """「插入」插在选中行上方；无选中则追加到末尾（跳过真实弹窗）。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        # 条件校验失败会弹模态框，测试里必须屏蔽，否则会打断运行
        monkeypatch.setattr(core.messagebox, "showwarning", lambda *a, **k: None)
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        app.conditions = [["a", "1"], ["b", "2"]]
        app._refresh_condition_tree()
        app._condition_dialog = lambda *a, **k: ("x", "9")

        app.cond_tree.selection_set("1")
        app._cond_insert()
        assert app.conditions == [["a", "1"], ["x", "9"], ["b", "2"]]
        assert app.cond_tree.selection() == ("1",)

        # 无选中：追加到末尾（用不同的查找词，避免被去重拦截）
        app.cond_tree.selection_remove("1")
        app._condition_dialog = lambda *a, **k: ("z", "0")
        app._cond_insert()
        assert len(app.conditions) == 4
        assert app.conditions[-1] == ["z", "0"]
    finally:
        root.destroy()


def test_cond_header_drops_unique_marker(tmp_path, monkeypatch):
    """表头只写「查找内容」，重复由弹窗拦截，不在表头提示。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        assert "不可重复" not in app.cond_tree.heading("find", "text")
        app.lang = "en"
        app.apply_lang()
        assert "Unique" not in app.cond_tree.heading("find", "text")
    finally:
        root.destroy()


# =====================================================================
# 多条件执行方式：属性与方法同名遮蔽的回归
# =====================================================================

def test_multi_mode_switch_and_lang_switch_keep_in_sync(tmp_path, monkeypatch):
    """选链式 → 切语言 → 模式值不变；下拉显示串随语言刷新。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        root.update()
        app.multi_mode_var.set(app._t("multi_mode_chain"))
        root.update()
        assert app._multi_mode == "chain"

        app.lang = "en"
        app.apply_lang()
        assert app._multi_mode == "chain", "切语言不许把内部模式洗掉"
        assert app.multi_mode_var.get() == app._t("multi_mode_chain")

        app.multi_mode_var.set(app._t("multi_mode_parallel"))
        root.update()
        assert app._multi_mode == "parallel"
    finally:
        root.destroy()


# =====================================================================
# 规则文件名：多条件摘要口径
# =====================================================================

def test_rule_filename_uses_condition_summary_for_multi(tmp_path, monkeypatch):
    """多条件规则文件名带上前两条条件 + 等N项，且不含引号。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        conds = [["甲公司", "乙公司"], ["2024", "2025"], ["甲", "乙"]]
        name = app.generate_rule_filename(core.conditions_rule_base(conds), "")
        assert "甲公司→乙公司" in name and "2024→2025" in name
        assert "等 3 项" in name
        assert '"' not in name, "引号会被文件名清洗成下划线，不能进文件名"
        assert name.endswith(".json")

        # 单条件保持「查找 替换」老口径（箭头两侧的 > 是文件名非法字符，
        # 落盘时会被清洗成 _，这是既有行为，不断言箭头本身）
        single = app.generate_rule_filename("甲公司", "乙公司")
        assert "甲公司" in single and "乙公司" in single
        assert not set(single) & set('/\\:*?"<>|')
    finally:
        root.destroy()


def test_rule_filename_stays_safe_with_illegal_chars(tmp_path, monkeypatch):
    """查找词含文件名非法字符时，落盘名仍安全。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        name = app.generate_rule_filename('a/b:c*d?', "x")
        assert not set(name) & set('/\\:*?"<>|'), "非法字符须被替换"
    finally:
        root.destroy()


def test_rollback_multi_file_flow_end_to_end(tmp_path, monkeypatch):
    """多文件回溯全流程：确认 → 文件勾选窗 → 执行。

    勾选窗的「确定」曾因 checked 键是 iid 字符串却拿去索引 list 而抛
    TypeError——异常被 Tk 吞掉，表现就是「点了确定什么都没反应」。这里走
    完整 UI 链路，确保那条路径真的能跑到底。
    """
    root = tk.Tk()
    root.deiconify()
    root.geometry("1000x700")
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: None)
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None

        work = tmp_path / "work"
        work.mkdir()
        targets = []
        for i in range(3):
            p = work / ("f%d.txt" % i)
            p.write_text("甲公司", encoding="utf-8")
            targets.append(p)
        folder = _mk_manifest_entry(tmp_path, "2026-10-16 10:00:00",
                                    [str(p) for p in targets])
        for p in targets:                      # 模拟替换已发生
            p.write_text("乙公司", encoding="utf-8")

        app._confirm_dialog = lambda *a, **k: True     # 跳过模态确认
        entries, _ = core.scan_backups(str(tmp_path))

        # 文件勾选窗是模态的：用 after 在 wait_window 期间自动点「确定」
        def auto_ok():
            tops = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]

            def walk(w):
                yield w
                for c in w.winfo_children():
                    yield from walk(c)

            for top in tops:
                btns = [w for w in walk(top)
                        if getattr(w, "_br_text", None) == app._t("ok")]
                if btns:
                    btns[0].event_generate("<Button-1>")
                    return

        root.after(500, auto_ok)
        app._rollback_entry(entries[0], entries, root)
        root.update()

        for p in targets:
            assert p.read_text(encoding="utf-8") == "甲公司", \
                "多文件回溯后应全部恢复: %s" % p
        assert os.path.isdir(os.path.join(folder, "rollback")), "回滚前应留快照"
    finally:
        root.destroy()


# =====================================================================
# 状态栏模式提示：「当前：替换文件内容」住在状态栏汇总句右边
# =====================================================================

def test_mode_hint_lives_in_status_bar(tmp_path, monkeypatch):
    """模式提示不再是参数区第二行的标签，而是状态栏上紧随汇总句的标签。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        root.update()
        assert not hasattr(app, "mode_hint"), "参数区不应再留模式提示标签"
        assert app.mode_label.winfo_ismapped() or app.mode_label.winfo_manager() == "pack"
        assert app.mode_label.winfo_rootx() >= (
            app.status_label.winfo_rootx() + app.status_label.winfo_width() - 2), \
            "模式提示应紧跟汇总句右边"
        assert "当前" in app.mode_label.cget("text")
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        assert "多条件" in app.mode_label.cget("text")
    finally:
        root.destroy()


def test_preview_title_has_no_version():
    """搜索结果窗标题不带版本号。"""
    assert core.TEXTS["preview_title"]["zh"] == "预览替换结果"
    assert core.TEXTS["preview_title"]["en"] == "Preview Results"


def test_cond_divider_between_columns(tmp_path, monkeypatch):
    """条件表两列之间盖着 1px 分界线，位置落在「查找内容」列宽处。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        app.conditions = [["甲公司", "乙公司"], ["2024", "2025"]]
        app._refresh_condition_tree()
        root.update()
        root.update_idletasks()
        assert app.cond_divider.winfo_manager() == "place", "分界线未摆放"
        assert app.cond_divider.winfo_x() == int(app.cond_tree.column("find", "width")), \
            "分界线应在两列交界上"
        assert app.cond_divider.winfo_width() == 1
    finally:
        root.destroy()


def test_dark_theme_checkbox_uses_palette(tmp_path, monkeypatch):
    """深色主题下自绘勾选框取深色底板与深色文字，不用系统白框。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.theme = "dark"
        app._apply_theme("dark")
        root.update()
        label = _check_label(app, "match_case")
        assert label is not None
        assert str(label.cget("bg")).upper() == app.BG.upper(), \
            "勾选框文字底色没跟主题"
        assert str(label.cget("fg")).upper() == app.TXT.upper(), \
            "勾选框文字颜色没跟主题"
        canvas = label.master.winfo_children()[0]
        assert str(canvas.cget("bg")).upper() == app.BG.upper()
    finally:
        root.destroy()


@pytest.mark.skipif(os.name != "nt", reason="任务栏身份只在 Windows 上有意义")
def test_window_icon_applied(tmp_path, monkeypatch):
    """窗口图标要真的挂上去：标题栏左上角与任务栏按钮都读它。

    图标先画到临时 .ico 再走 iconbitmap；这条用例盯住"路径拿到了但没设成功"
    ——那种情况下窗口能开，只是任务栏和标题栏都是一块空白。
    """
    root = tk.Tk()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        core.BatchReplaceApp(root)
        root.update()
        user32 = ctypes.windll.user32
        user32.GetClassLongPtrW.restype = ctypes.c_ssize_t
        user32.GetClassLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
        hwnd = user32.GetAncestor(ctypes.c_void_p(root.winfo_id()), 2)
        # iconbitmap 的 default 形式设的是窗口类图标（GCLP_HICON / GCLP_HICONSM），
        # 不是窗口级 WM_SETICON —— 后者为空时系统正是回退到类图标来画标题栏
        # 与任务栏按钮，所以这里查类图标。
        assert user32.GetClassLongPtrW(hwnd, -14), "大图标未设置"      # GCLP_HICON
        assert user32.GetClassLongPtrW(hwnd, -34), "小图标未设置"      # GCLP_HICONSM
    finally:
        root.destroy()


def test_theme_rebuild_syncs_multi_mode_display(tmp_path, monkeypatch):
    """换主题重建界面后，执行方式下拉的显示串必须跟着内部模式走。"""
    root = tk.Tk()
    root.withdraw()
    try:
        monkeypatch.setattr(core, "CONFIG_FILE", str(tmp_path / "cfg.json"))
        monkeypatch.setattr(core, "HISTORY_FILE", str(tmp_path / "hist.json"))
        monkeypatch.setattr(core, "RULES_DIR", str(tmp_path / "rules"))
        app = core.BatchReplaceApp(root)
        app.save_config = lambda: None
        app.multi_condition_var.set(True)
        app.toggle_multi_condition()
        app.multi_mode_var.set(app._t("multi_mode_chain"))
        root.update()
        app._toggle_theme()
        root.update()
        assert app._multi_mode == "chain", "换主题不许洗掉内部模式"
        assert app.multi_mode_var.get() == app._t("multi_mode_chain"), \
            "换主题后下拉显示串应与内部模式一致"
    finally:
        root.destroy()


# =====================================================================
# 更新检查与一键升级
# =====================================================================

def test_version_tuple_parses_tag_and_exe_name():
    assert core.version_tuple("v2.0.0") == (2, 0, 0)
    assert core.version_tuple("BatchReplace_v2.1.3.exe") == (2, 1, 3)
    assert core.version_tuple("junk") == (0, 0, 0)


def test_compare_versions_three_branches():
    assert core.compare_versions("v2.0.1", "v2.0.0") == 1
    assert core.compare_versions("v1.9.9", "v2.0.0") == -1
    assert core.compare_versions("2.0.0", "v2.0.0") == 0


def test_silent_update_target_only_new_version():
    assert core.silent_update_target("v2.0.1", "url", "", "2.0.0") == "url"
    assert core.silent_update_target("v2.0.1", "", "", "2.0.0") == core.RELEASE_PAGE_URL
    assert core.silent_update_target("v2.0.0", "url", "", "2.0.0") == ""
    assert core.silent_update_target("v1.9.9", "url", "", "2.0.0") == ""
    assert core.silent_update_target("v2.0.1", "url", "network error", "2.0.0") == ""
    assert core.silent_update_target("", "url", "", "2.0.0") == ""


def test_fetch_latest_release_prefers_release(monkeypatch):
    payload = (b'{"tag_name":"v2.0.1","html_url":"https://example.test/release"}')
    monkeypatch.setattr(core, "_github_get", lambda url, timeout=8: payload)
    assert core.fetch_latest_release() == (
        "v2.0.1", "https://example.test/release", None)


def test_fetch_latest_release_falls_back_to_tags(monkeypatch):
    import urllib.error
    err = urllib.error.HTTPError(core.UPDATE_API, 404, "Not Found", None, None)

    def fake_get(url, timeout=8):
        if url in (core.MANIFEST_API, core.UPDATE_MANIFEST):
            raise OSError("manifest unavailable")
        if url == core.UPDATE_API:
            raise err
        return b'[{"name":"v1.9.9"},{"name":"v2.0.2"},{"name":"v2.0.1"}]'

    monkeypatch.setattr(core, "_github_get", fake_get)
    assert core.fetch_latest_release() == (
        "v2.0.2", core.RELEASE_PAGE_URL, None)


def test_fetch_latest_release_reports_error(monkeypatch):
    def fake_get(url, timeout=8):
        raise OSError("offline")

    monkeypatch.setattr(core, "_github_get", fake_get)
    tag, url, err = core.fetch_latest_release()
    assert tag == "" and url == "" and "offline" in err


def test_fetch_latest_release_reads_manifest_via_api(monkeypatch):
    """API contents 主路：base64 envelope 解出清单，只读 version/page。"""
    import base64
    import json
    info = {"version": "v2.0.1", "page": "https://example.test/tag/v2.0.1",
            "pub_date": "2026-11-24T11:10:00Z"}
    envelope = json.dumps({"content": base64.b64encode(
        json.dumps(info).encode("utf-8")).decode("ascii")}).encode("utf-8")
    seen = []

    def fake_get(url, timeout=8):
        seen.append(url)
        return envelope

    monkeypatch.setattr(core, "_github_get", fake_get)
    assert core.fetch_latest_release() == (
        "v2.0.1", "https://example.test/tag/v2.0.1", None)
    assert seen == [core.MANIFEST_API]


def test_fetch_latest_release_api_falls_back_to_raw(monkeypatch):
    """API 读不到时退 raw，清单即使带 asset 字段也不再解析。"""
    seen = []

    def fake_get(url, timeout=8):
        seen.append(url)
        if url == core.MANIFEST_API:
            raise OSError("api blocked")
        return (b'{"version":"v2.0.1","page":"https://example.test/tag/v2.0.1",'
                b'"asset":"https://example.test/exe","sha256":"' + b"b" * 64 + b'"}')

    monkeypatch.setattr(core, "_github_get", fake_get)
    assert core.fetch_latest_release() == (
        "v2.0.1", "https://example.test/tag/v2.0.1", None)
    assert seen == [core.MANIFEST_API, core.UPDATE_MANIFEST]


def test_one_click_update_module_present():
    """一键更新链路完整：清单校验 → 下载验签 → 校验即启动。"""
    src = UI_SOURCE.read_text(encoding="utf-8")
    for name in ("UPDATE_MANIFEST", "validate_update_manifest",
                 "fetch_update_manifest", "download_verified_exe",
                 "launch_verified_exe", "run_one_click_update",
                 "silent_update_target", "subprocess"):
        assert name in src, "一键更新模块缺失：%s" % name
    # 已废弃的 bat 脚本式升级不得回潮
    for name in ("_bat_quote", "write_upgrade_bat", "start_upgrade_bat", "run_upgrade"):
        assert name not in src, "已废弃的 bat 升级残留：%s" % name


class _SyncThread:
    def __init__(self, target=None, args=(), **_kwargs):
        self.target = target
        self.args = args

    def start(self):
        self.target(*self.args)


class _SyncThreading:
    Thread = _SyncThread


def _fast_after(monkeypatch, root):
    """让 100ms 的 UI 回投立即执行，其余定时器保持原样。"""
    original = root.after

    def patched(ms, callback, *args):
        if ms == 100:
            callback(*args)
            return "immediate"
        return original(ms, callback, *args)

    monkeypatch.setattr(root, "after", patched)


def _update_app(monkeypatch, root, logs):
    _fast_after(monkeypatch, root)
    app = core.BatchReplaceApp.__new__(core.BatchReplaceApp)
    app.root = root
    app.log = lambda message: logs.append(message)
    return app


def test_silent_check_does_not_popup_without_new_version(monkeypatch, tk_root):
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "fetch_update_manifest",
                        lambda: ({"version": "v2.0.0", "page": ""}, None))
    monkeypatch.setattr(core.BatchReplaceApp, "ask_update_choice",
                        lambda *a, **k: pytest.fail("无新版不该弹三选一"))
    logs = []
    _update_app(monkeypatch, tk_root, logs).silent_check_update()
    assert any("未发现新版本" in line for line in logs)


def test_silent_check_reports_new_version_in_log_only(monkeypatch, tk_root):
    """启动检查发现新版照常弹三选一：静默只压「没有新版」「检查失败」两种噪音。

    启动时用户也需要知道有新版并决定怎么处理；「已是最新」这类弹窗才是纯噪音。
    """
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "fetch_update_manifest",
                        lambda: ({"version": "v2.0.9",
                                  "page": "https://example.test/v209"}, None))
    asked = []
    monkeypatch.setattr(core.BatchReplaceApp, "ask_update_choice",
                        lambda self, parent, latest: asked.append(latest) or "skip")
    opened = []
    monkeypatch.setattr(core.BatchReplaceApp, "_open_url",
                        lambda self, url: opened.append(url))
    logs = []
    _update_app(monkeypatch, tk_root, logs).silent_check_update()
    assert asked == ["v2.0.9"], "启动检查发现新版没弹三选一"
    assert opened == [], "选「暂不更新」不该打开发布页"
    assert any("发现新版本" in line for line in logs)


def test_update_launch_hands_over_and_closes_old_window(monkeypatch, tk_root):
    """新版启动成功后旧版自动关闭；启动失败则留着旧版并报错。"""
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "download_verified_exe",
                        lambda *a, **k: ("C:/tmp/BatchReplace_v2.0.9.exe", None))
    monkeypatch.setattr(core.messagebox, "showinfo", lambda *a, **k: None)
    errs = []
    monkeypatch.setattr(core.messagebox, "showerror",
                        lambda *a, **k: errs.append(a))

    closed = []
    app = _update_app(monkeypatch, tk_root, [])
    app.on_closing = lambda: closed.append(True)
    monkeypatch.setattr(core, "launch_verified_exe", lambda _p: (True, None))
    app._download_and_launch(tk_root, {"version": "v2.0.9"}, "v2.0.9")
    assert closed == [True], "新版启动后旧版没关，两个窗口并存"

    closed.clear()
    monkeypatch.setattr(core, "launch_verified_exe", lambda _p: (False, "boom"))
    app._download_and_launch(tk_root, {"version": "v2.0.9"}, "v2.0.9")
    assert closed == [] and errs, "启动失败时应留着旧版并报错"


def test_manual_check_reports_latest_and_logs(monkeypatch, tk_root):
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "fetch_update_manifest",
                        lambda: ({"version": "v2.0.0", "page": ""}, None))
    shown = []
    monkeypatch.setattr(
        core.messagebox, "showinfo", lambda *a, **k: shown.append(a) or None)
    logs = []
    _update_app(monkeypatch, tk_root, logs).check_for_update(tk_root)
    assert len(shown) == 1 and "2.0.0" in shown[0][1]
    assert any("未发现新版本" in line for line in logs)


def test_manual_check_skip_does_not_open(monkeypatch, tk_root):
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "fetch_update_manifest",
                        lambda: ({"version": "v2.0.1", "page": "url"}, None))
    monkeypatch.setattr(core.BatchReplaceApp, "ask_update_choice", lambda *a, **k: "skip")
    opened = []
    monkeypatch.setattr(core.BatchReplaceApp, "_open_url",
                        lambda self, url: opened.append(url))
    logs = []
    _update_app(monkeypatch, tk_root, logs).check_for_update(tk_root)
    assert opened == []


def test_manual_check_confirm_opens_release_page(monkeypatch, tk_root):
    monkeypatch.setattr(core, "threading", _SyncThreading)
    monkeypatch.setattr(core, "fetch_update_manifest",
                        lambda: ({"version": "v2.0.1", "page": "https://example.test/v201"}, None))
    monkeypatch.setattr(core.BatchReplaceApp, "ask_update_choice", lambda *a, **k: "page")
    opened = []
    monkeypatch.setattr(core.BatchReplaceApp, "_open_url",
                        lambda self, url: opened.append(url))
    logs = []
    _update_app(monkeypatch, tk_root, logs).check_for_update(tk_root)
    assert opened == ["https://example.test/v201"]

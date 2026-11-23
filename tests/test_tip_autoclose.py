"""悬浮说明弹窗默认 N 秒自动关闭的回归测试。

需求：界面上的悬浮说明弹窗（多条件执行方式说明、历史窗口表头说明）默认
存活 10 秒后自动收起，不需要用户手动点 × 或移开鼠标。测试把
TIP_AUTO_CLOSE_MS 临时调小（几十毫秒），泵事件循环验证弹窗确实被定时器
收掉，避免真等 10 秒。
"""

import time
import tkinter as tk

import pytest

import BatchReplace as core


@pytest.fixture
def tk_root():
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


def _pump(root, predicate, timeout=3.0, step=0.01):
    """泵事件循环直到 predicate() 为真或超时；返回 predicate 最终值。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if predicate():
            return True
        time.sleep(step)
    return predicate()


def test_mode_tip_auto_closes(tk_root):
    """执行方式说明弹窗：无人理会也应在 TIP_AUTO_CLOSE_MS 后自动关闭。"""
    app = core.BatchReplaceApp(tk_root)
    app.TIP_AUTO_CLOSE_MS = 80  # 临时调小，免等 10 秒

    app.multi_condition_var.set(True)
    app.toggle_multi_condition(announce=True)
    assert app._chain_tip is not None, "勾选多条件后应弹出说明"

    closed = _pump(tk_root, lambda: app._chain_tip is None, timeout=3.0)
    assert closed, "执行方式说明弹窗未在 TIP_AUTO_CLOSE_MS 后自动关闭"


def _find_label(widget, text):
    """递归找 text 匹配的 Label。"""
    found = []
    for child in widget.winfo_children():
        if isinstance(child, tk.Label) and child.cget("text") == text:
            found.append(child)
        found.extend(_find_label(child, text))
    return found


def test_history_header_tip_auto_closes(tk_root):
    """历史窗口表头悬停说明：同样应在 TIP_AUTO_CLOSE_MS 后自动关闭。"""
    app = core.BatchReplaceApp(tk_root)
    app.TIP_AUTO_CLOSE_MS = 80

    before = set(tk_root.winfo_children())
    app.show_history_dialog()
    new_wins = [w for w in tk_root.winfo_children() if w not in before]
    history = [w for w in new_wins if isinstance(w, tk.Toplevel)]
    assert history, "应打开历史窗口"
    hw = history[0]

    header_text = app._t("history_header", n=core.MAX_HISTORY)
    labels = _find_label(hw, header_text)
    assert labels, "应找到历史表头"
    header = labels[0]

    # 触发悬停：400ms 后弹提示，再过 TIP_AUTO_CLOSE_MS 自动关
    header.event_generate("<Enter>")

    def any_tip():
        return [w for w in hw.winfo_children()
                if isinstance(w, tk.Toplevel) and w.overrideredirect()]

    # 先等提示弹出
    appeared = _pump(tk_root, lambda: bool(any_tip()), timeout=3.0, step=0.02)
    assert appeared, "悬停表头后应弹出说明"

    # 再等它自动关闭
    gone = _pump(tk_root, lambda: not any_tip(), timeout=3.0, step=0.02)
    assert gone, "历史表头说明未在 TIP_AUTO_CLOSE_MS 后自动关闭"

    try:
        hw.destroy()
    except Exception:
        pass


def _open_history(tk_root, app):
    before = set(tk_root.winfo_children())
    app.show_history_dialog()
    new_wins = [w for w in tk_root.winfo_children() if w not in before]
    hw = [w for w in new_wins if isinstance(w, tk.Toplevel)][0]
    labels = _find_label(hw, app._t("history_header", n=core.MAX_HISTORY))
    return hw, labels[0]


def _hover_tip(hw):
    return [w for w in hw.winfo_children()
            if isinstance(w, tk.Toplevel) and w.overrideredirect()]


def test_history_header_sweep_leaves_no_ghost_tip(tk_root):
    """鼠标扫过表头（Enter 后立刻 Leave）不得留下「离开后仍弹出」的幽灵窗。

    延迟弹出任务不撤时，扫 N 次就会在鼠标早已离开后先后冒出 N 个窗——
    实际使用中表现为「不停弹窗」。
    """
    app = core.BatchReplaceApp(tk_root)
    app.TIP_AUTO_CLOSE_MS = 80
    hw, header = _open_history(tk_root, app)

    for _ in range(4):
        header.event_generate("<Enter>")
        header.event_generate("<Leave>")

    _pump(tk_root, lambda: False, timeout=1.0)
    assert not _hover_tip(hw), "鼠标已离开，不应再弹出说明窗"

    try:
        hw.destroy()
    except Exception:
        pass


def test_history_tip_pending_show_dropped_when_dialog_closes(tk_root):
    """悬停后 0.4 秒内关掉窗口：待弹任务不能再对已销毁的父窗建 Toplevel。"""
    app = core.BatchReplaceApp(tk_root)
    app.TIP_AUTO_CLOSE_MS = 80
    hw, header = _open_history(tk_root, app)

    errors = []
    tk_root.report_callback_exception = lambda *a: errors.append(a)
    header.event_generate("<Enter>")
    hw.destroy()
    _pump(tk_root, lambda: False, timeout=1.0)
    assert not errors, "关窗后仍触发了回调异常: %r" % (errors[:1],)


def test_loading_rule_does_not_pop_mode_tip(tk_root, tmp_path, monkeypatch):
    """从规则文件恢复输入区不算「刚勾上多条件」，不应弹执行方式说明。"""
    import json

    app = core.BatchReplaceApp(tk_root)
    app.TIP_AUTO_CLOSE_MS = 80
    rule = {"multi": True,
            "conditions": [["北京", "上海"], ["广州", "深圳"]],
            "case_sensitive": False}
    path = tmp_path / "rule.json"
    path.write_text(json.dumps(rule, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(core.filedialog, "askopenfilename", lambda **kw: str(path))

    app.multi_condition_var.set(False)
    app.load_rule()

    assert app.multi_condition_var.get()
    assert app._chain_tip is None, "加载规则不该弹执行方式说明"


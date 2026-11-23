"""多条件替换·执行方式说明悬浮窗的定位回归测试。

现象（用户报）：第一次勾选「多条件替换」时，并行/链式说明悬浮窗没有出现在
正确位置；显示过一次后正常。根因：toggle_multi_condition 刚把下拉框 pack
显形，同一事件处理里 _chain_tip_show 立刻读 combo.winfo_rootx/rooty/height
给悬浮窗定位，但 pack 只登记布局、几何要等 Tk 跑一轮事件才落地，此刻读到的是
旧值（甚至 0）。修复：定位前先 self.root.update_idletasks() 把布局算完。

本测试不依赖无头环境的绝对坐标值（无头下 rootx 恒为 0，无法区分两种状态），
改为捕获「调用顺序」：断言 _chain_tip_show 在读取下拉框几何之前先调用了
update_idletasks——这正是修复的本质，顺序对了定位才对。
"""

import tkinter as tk
import pytest

import BatchReplace as core


@pytest.fixture
def tk_root():
    """真 Tk 根窗口（隐藏不弹窗）；无图形环境则整体跳过。"""
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


def test_multi_mode_tip_flushes_layout_before_reading_geometry(tk_root, monkeypatch):
    """修复后：_chain_tip_show 必须先 update_idletasks，再读下拉框几何。"""
    app = core.BatchReplaceApp(tk_root)

    calls = []
    orig_update = app.root.update_idletasks

    def update_spy():
        calls.append("update")
        return orig_update()

    app.root.update_idletasks = update_spy

    combo = app.multi_mode_combo
    for meth in ("winfo_rootx", "winfo_rooty", "winfo_height"):
        orig = getattr(combo, meth)

        def geom_spy(_orig=orig, _meth=meth):
            calls.append("geom:%s" % _meth)
            return _orig()

        monkeypatch.setattr(combo, meth, geom_spy)

    app.multi_condition_var.set(True)
    app.toggle_multi_condition(announce=True)

    assert app._chain_tip is not None, "勾选多条件后应弹出说明悬浮窗"
    assert "update" in calls, "_chain_tip_show 必须先调用 update_idletasks 算布局"

    update_idx = calls.index("update")
    geom_idxs = [i for i, c in enumerate(calls) if c.startswith("geom:")]
    assert geom_idxs, "应读取下拉框几何用于定位"
    assert min(geom_idxs) > update_idx, (
        "几何读取发生在 update_idletasks 之前（修复前的 bug）：%r" % calls
    )


def test_multi_mode_tip_anchored_below_combo(tk_root):
    """说明悬浮窗应贴在下拉框正下方（y = 下拉框底 + 4）。"""
    app = core.BatchReplaceApp(tk_root)

    recorded = {}
    orig_geo = tk.Toplevel.wm_geometry

    def spy(self_, arg=None):
        if arg is not None:
            recorded["geo"] = arg
        return orig_geo(self_, arg)

    # 让下拉框几何返回可控值，验证定位公式
    combo = app.multi_mode_combo
    combo.winfo_rootx = lambda: 100
    combo.winfo_rooty = lambda: 200
    combo.winfo_height = lambda: 30
    tk.Toplevel.wm_geometry = spy

    app.multi_condition_var.set(True)
    app.toggle_multi_condition(announce=True)

    assert recorded["geo"] == "+100+234", (
        "悬浮窗应贴在下拉框正下方(+rootx, +rooty+height+4)，实际 %r"
        % recorded["geo"]
    )

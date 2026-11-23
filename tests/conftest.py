# -*- coding: utf-8 -*-
"""测试补丁：Tk 实例重试 + 测试窗口不打扰桌面。

1) 重试：Windows 上测试反复创建 / 销毁 Tk 实例时，Tcl 初始化偶尔读不到
   tcl/tk 资源文件（杀毒实时扫描对 .tcl 的瞬时锁），表现为
   ``_tkinter.TclError: Can't find a usable tk.tcl``。产品运行期只建一个 Tk
   实例，不存在此问题；只在测试里高频 new Tk 才偶发撞上。这里重试几次，
   消除偶发失败，不影响任何产品逻辑。
2) 隐形：几百项用例各自建窗 / 建对话框再销毁，桌面上会连续闪空窗。
3) 不抢焦点：光把窗口调透明不够——Windows 映射新窗口时会把它提到最前并
   夺走键盘焦点，跑测试期间别人打字会被打断。加 ``WS_EX_NOACTIVATE``
   后窗口不会成为前台窗口，再配 ``WS_EX_TOOLWINDOW`` 不进任务栏。
4) 打标志的时机：必须在窗口**第一次映射之前**。先 update 再补标志的话，
   那一次映射已经把窗口映射到桌面、任务栏加了按钮、焦点被抢走了，事后
   补只对下一次映射有效。``overrideredirect`` 会重建 HWND 并丢掉扩展样式，
   所以那次调用之后要再打一遍。
"""
import ctypes
import tempfile
import tkinter as tk
import time

_ORIG_TK = tk.Tk
_ORIG_TOPLEVEL = tk.Toplevel
_ORIG_OVERRIDEREDIRECT = tk.Wm.overrideredirect
_ORIG_BIND = tk.Misc.bind

_GWL_EXSTYLE = -20
_WS_EX_NOACTIVATE = 0x08000000
_WS_EX_TOOLWINDOW = 0x00000080
_WS_EX_APPWINDOW = 0x00040000


def _no_activate(win):
    """给窗口句柄打上「不可激活 + 不进任务栏」。

    只建窗、不触发映射：Toplevel 的外框 HWND 要到第一次映射才由 Tk 建出来，
    建窗这一刻它还不存在（GetParent 返回 0），所以首次调用只能标到内层窗；
    真正的开关在外框上，必须等映射后由 ``_restyle_on_map`` 再补一次。

    ``WS_EX_APPWINDOW`` 必须清掉：带这一位的窗口即使有 owner、即使标了
    TOOLWINDOW，任务栏照样给它一个按钮。
    """
    try:
        hwnd = win.winfo_id()          # 需要就建 HWND，但不映射到桌面
        user32 = ctypes.windll.user32
        for target in (hwnd, user32.GetParent(hwnd)):
            if not target:
                continue
            style = user32.GetWindowLongW(target, _GWL_EXSTYLE)
            style = (style | _WS_EX_NOACTIVATE | _WS_EX_TOOLWINDOW) \
                & ~_WS_EX_APPWINDOW
            user32.SetWindowLongW(target, _GWL_EXSTYLE, style)
    except Exception:
        # 补丁失败只是少一层保护，不影响用例本身。
        pass


def _restyle_on_map(event):
    """映射完成后补标志：此刻外框 HWND 才存在，任务栏按钮随即被摘掉。"""
    _no_activate(event.widget)


def _bind_keeps_quiet(self, sequence=None, func=None, add=None):
    """产品代码会 ``root.bind("<Map>", self._on_root_map)``——不带 add 的绑定
    会把已有处理器整体替换掉，补丁就失效了。这里在每次绑定 <Map> 之后追加
    自己的处理器，顺序在业务之后，不影响业务逻辑。"""
    result = _ORIG_BIND(self, sequence, func, add)
    if sequence == "<Map>" and func is not None and func is not _restyle_on_map:
        _ORIG_BIND(self, "<Map>", _restyle_on_map, "+")
    return result


def _invisible(win):
    try:
        win.attributes("-alpha", 0.0)
    except tk.TclError:
        pass
    _no_activate(win)
    try:
        win.bind("<Map>", _restyle_on_map, add="+")
    except tk.TclError:
        pass
    return win


def _tk_with_retry(*args, **kwargs):
    last = None
    for attempt in range(3):
        try:
            return _invisible(_ORIG_TK(*args, **kwargs))
        except tk.TclError as exc:
            last = exc
            time.sleep(0.25 * (attempt + 1))
    raise last


class _InvisibleToplevel(_ORIG_TOPLEVEL):
    """对话框同样不上桌面。用子类而不是函数包装——产品代码里若有
    ``class X(tk.Toplevel)``，函数会当场 TypeError。"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        _invisible(self)


def _override_keeps_quiet(self, *args, **kwargs):
    """overrideredirect 重建 HWND，扩展样式随之丢失，重建后立刻补回。"""
    result = _ORIG_OVERRIDEREDIRECT(self, *args, **kwargs)
    _no_activate(self)
    return result


tk.Tk = _tk_with_retry
tk.Toplevel = _InvisibleToplevel
tk.Wm.overrideredirect = _override_keeps_quiet
tk.Misc.bind = _bind_keeps_quiet


def pytest_configure(config):
    # 临时目录固定到独立路径：默认的 pytest-of-<user> 全局根会被 pytest 的
    # 跨会话清扫机制整目录删除，在受控环境下大批量删除会被安全策略拦截，
    # 把整轮已通过的测试带成非零退出码。独立根目录不在清扫范围内。
    # 每轮再用一个全新目录：pytest 启动时会清空 basetemp，若沿用同一个固定
    # 路径，上一轮留下的几百个临时文件会在本轮开头被一次性删除——同样触发
    # 上述拦截，fixture 收尾状态被打乱后，后续上百个用例连带报 setup 错误。
    if not config.option.basetemp:
        config.option.basetemp = tempfile.mkdtemp(prefix="br_pytest_")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把一键更新链路注入 BatchReplace v2.0.0 源码（结构锚点，不用行号）。"""
import io
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "BatchReplace.py")


def between(text, start, end):
    i = text.index(start)
    j = text.index(end, i)
    return text[i:j]


MODULE = '''# ---- 一键更新：清单校验 + 下载校验 + 校验即启动 ----
# SHA-256 只防下载损坏/截断，不防恶意替换（哈希与文件同源）。单一 GitHub 来源 + HTTPS 下足够；
# 若将来接入第三方镜像，必须补清单签名。
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/YehOlt/BatchReplace/refs/tags/updater/update.json"
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
    if re.search(r"[\\\\/]", exe) or ".." in exe:
        raise ValueError("清单 exe 不得含路径分隔或 ..")
    if not re.search(r"[_-]v?(\\d+\\.\\d+\\.\\d+)", exe, re.I):
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


def fetch_update_manifest(url=None, timeout=8):
    """取清单 → 字段校验。返回 (清单, 错误说明)。"""
    target = url or UPDATE_MANIFEST_URL
    try:
        req = urllib.request.Request(target, headers={"User-Agent": "BatchReplace/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(UPDATE_MAX_MANIFEST + 1)
        if len(raw) > UPDATE_MAX_MANIFEST:
            return None, "清单异常大"
        m = json.loads(raw.decode("utf-8"))
        validate_update_manifest(m)
        return m, None
    except Exception as exc:
        return None, str(exc)


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


'''

UI = '''    def silent_check_update(self):
        """启动时自动检查一次，复用 check_for_update 的清单校验与三选一流程。"""
        self.check_for_update(self.root, silent=True)

    def ask_update_choice(self, parent, latest):
        """三选一：直接下载 / 发布页 / 暂不更新。返回 download/page/skip。"""
        chosen = {"value": "skip"}
        dlg = tk.Toplevel(parent)
        dlg.title(self._t("t_update_found"))
        dlg.transient(parent)
        dlg.resizable(False, False)
        dlg.grab_set()

        def close(choice):
            chosen["value"] = choice
            dlg.destroy()

        page = ttk.Frame(dlg, padding=(20, 16))
        page.pack(fill="both", expand=True)
        ttk.Label(page,
                  self._t("m_update_found_short", new=latest, cur="v" + VERSION),
                  font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")
        ttk.Label(page, text=self._t("l_update_choose"), anchor="w").pack(anchor="w")
        row = ttk.Frame(page)
        row.pack(pady=(14, 0), fill="x")
        ttk.Button(row, text=self._t("b_update_download"), width=12,
                   command=lambda: close("download")).pack(side="left")
        ttk.Button(row, text=self._t("b_update_page"), width=10,
                   command=lambda: close("page")).pack(side="left", padx=(8, 0))
        ttk.Button(row, text=self._t("b_update_skip"), width=10,
                   command=lambda: close("skip")).pack(side="left", padx=(8, 0))

        dlg.protocol("WM_DELETE_WINDOW", lambda: close("skip"))
        dlg.bind("<Escape>", lambda _e: close("skip"))
        dlg.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - dlg.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
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

        silent=True 时抑制「已是最新版」与失败弹窗，发现新版照常三选一。
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
                self.log(self._t("lg_update_fail", e=err))
                if silent:
                    return
                if messagebox.askyesno(
                        self._t("t_update_fail"),
                        self._t("m_update_fail", e=err), parent=parent):
                    self._open_url(RELEASE_PAGE_URL)
                return
            latest = manifest["version"]
            if compare_versions(latest, VERSION) <= 0:
                self.log(self._t("lg_update_none", cur="v" + VERSION))
                if not silent:
                    messagebox.showinfo(
                        self._t("check_update"),
                        self._t("m_update_latest", cur="v" + VERSION), parent=parent)
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

        self.root.after(100, poll)


'''

I18N = '''    "m_update_found_short": {"zh": "发现新版本 {new}，当前版本 {cur}。",
                             "en": "New version {new} available (current: {cur})."},
    "l_update_choose": {"zh": "\\n请选择处理方式：", "en": "\\nChoose an action:"},
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
    "m_update_launch_fail": {"zh": "新版已下载并校验，但启动失败：{e}\\n文件：{path}",
                             "en": "Update verified but failed to launch: {e}\\nFile: {path}"},
    "t_update_done": {"zh": "更新完成", "en": "Update complete"},
    "m_update_done": {"zh": "新版 {new} 已下载并启动。\\n\\n新版文件已保存到同一目录，旧版文件保留不变。",
                      "en": "Version {new} has been downloaded and launched.\\n\\nThe new exe is in the same folder; the old file is kept."},
'''

CLI = '''    # 命令行更新入口：--update-manifest <url> [--update-dry-run]
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

'''


def main():
    s = io.open(SRC, encoding="utf-8").read()
    orig = s

    # 1) import 补齐（SELFCHECK_IMPORTS 也要同步，否则 --selfcheck 失败）
    assert "import hashlib" not in s, "hashlib 已存在"
    assert "import subprocess" not in s, "subprocess 已存在"
    s = s.replace("import base64\nimport json\n",
                  "import base64\nimport hashlib\nimport json\n", 1)
    s = s.replace("import shutil\nimport tempfile\n",
                  "import shutil\nimport subprocess\nimport tempfile\n", 1)
    s = s.replace('"base64", "ctypes", "datetime", "docx", "json", "openpyxl", "os", "pathlib",',
                  '"base64", "ctypes", "datetime", "docx", "hashlib", "json", "openpyxl", "os", "pathlib",', 1)
    s = s.replace('"queue", "re", "shutil", "sys", "tempfile", "threading",',
                  '"queue", "re", "shutil", "subprocess", "sys", "tempfile", "threading",', 1)

    # 2) 更新模块插在 silent_update_target 之前
    anchor = "def silent_update_target("
    assert anchor in s
    s = s.replace(anchor, MODULE + anchor, 1)

    # 3) UI 段整体替换（silent_check_update → show_about）
    start = "    def silent_check_update(self):"
    end = "    def show_about(self"
    assert start in s and end in s
    ui_old = between(s, start, end)
    s = s.replace(ui_old, UI, 1)

    # 4) i18n 文案
    anchor = '    "m_theme_close_preview": '
    assert anchor in s
    s = s.replace(anchor, I18N + anchor, 1)

    # 5) main 命令行入口（在 _set_app_user_model_id 前）
    anchor = "    # 进程身份要在建窗之前定下来"
    assert anchor in s
    s = s.replace(anchor, CLI + anchor, 1)

    io.open(SRC, "w", encoding="utf-8").write(s)
    print("注入完成：%d → %d 字节" % (len(orig), len(s)))


if __name__ == "__main__":
    main()

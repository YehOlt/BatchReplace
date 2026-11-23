#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BatchReplace v2.0.0 一键更新链路测试（纯标准库，无需 pytest）。"""
import hashlib
import http.server
import importlib.util
import json
import os
import socketserver
import tempfile
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "BatchReplace.py")

spec = importlib.util.spec_from_file_location("br_update", SRC)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

PASS = []


def check(name, ok):
    assert ok, name
    PASS.append(name)
    print("PASS %s" % name)


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            h.update(b)
    return h.hexdigest()


def make_fake(path, tag=b"MZ" + b"fake-batchreplace-new-version" * 1000):
    with open(path, "wb") as f:
        f.write(tag)


class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass


def main():
    tmp = tempfile.mkdtemp(prefix="br-update-")
    web = os.path.join(tmp, "web")
    dest = os.path.join(tmp, "dest")
    os.makedirs(web)
    os.makedirs(dest)
    srv = socketserver.TCPServer(("127.0.0.1", 0), lambda *a, **k: Quiet(*a, directory=web, **k))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = "http://127.0.0.1:%d" % srv.server_address[1]
    try:
        exe = "BatchReplace_v9.9.9.exe"
        exe_web = os.path.join(web, exe)
        make_fake(exe_web)
        good = {
            "version": "9.9.9",
            "exe": exe,
            "sha256": sha(exe_web),
            "size": os.path.getsize(exe_web),
            "url": base + "/" + exe,
        }
        with open(os.path.join(web, "update.json"), "w", encoding="utf-8") as f:
            json.dump(good, f)

        check("合法清单字段校验", m.validate_update_manifest(good) is good)
        for bad in [dict(good, exe="../evil.exe"), dict(good, exe="BatchReplace.exe"),
                    dict(good, sha256="xyz"), dict(good, url="")]:
            try:
                m.validate_update_manifest(bad)
                raise AssertionError("非法清单未拒绝: %s" % bad)
            except ValueError:
                pass
        check("非法字段被拒绝（路径/版本名/sha/url）", True)

        got, err = m.fetch_update_manifest(base + "/update.json")
        check("本地 HTTP 取清单并校验", err is None and got["version"] == "9.9.9")

        path, err = m.download_verified_exe(got, dest)
        check("下载校验成功并落到指定文件名", err is None and os.path.basename(path) == exe)
        check("下载内容哈希一致", sha(path) == good["sha256"])

        badsha = dict(got, sha256="0" * 64)
        p2, e2 = m.download_verified_exe(badsha, dest)
        check("哈希不符被拒", p2 is None and "SHA-256" in e2)
        check("失败不留 .part", not os.path.exists(os.path.join(dest, exe + ".part")))
        check("失败不覆盖已校验文件", sha(os.path.join(dest, exe)) == good["sha256"])

        badsize = dict(got, size=1)
        p3, e3 = m.download_verified_exe(badsize, dest)
        check("大小不符被拒", p3 is None and "大小" in e3)

        fake_path = os.path.join(web, "BatchReplace_v9.9.8.exe")
        with open(fake_path, "wb") as f:
            f.write(b"PK\x03\x04not-an-exe" * 100)
        fake_m = dict(good, exe="BatchReplace_v9.9.8.exe", sha256=sha(fake_path),
                      size=os.path.getsize(fake_path), url=base + "/BatchReplace_v9.9.8.exe")
        p4, e4 = m.download_verified_exe(fake_m, dest)
        check("非 MZ 魔数被拒", p4 is None and "MZ" in e4)

        out = []
        r = m.run_one_click_update(base + "/update.json", log=out.append,
                                   dry_run=True, current="2.0.0")
        check("命令行 dry-run 发现新版并校验", "试运行完成" in r and "9.9.9" in r)

        out2 = []
        r2 = m.run_one_click_update(base + "/update.json", log=out2.append,
                                    dry_run=True, current="9.9.9")
        check("已是最新版不下载", "当前已是最新版" in r2)

        src = open(SRC, encoding="utf-8").read()
        check("三选一三按钮齐全",
              all(k in src for k in ('"b_update_download"', '"b_update_page"', '"b_update_skip"')))
        check("静默与手动共用同一实现",
              "self.check_for_update(self.root, silent=True)" in src
              and "def check_for_update(self, parent=None, silent=False):" in src)
        check("启动新版不接管旧版",
              "新版文件已保存到同一目录" in src
              and "旧版窗口即将被新版接管" not in src)
        check("命令行入口存在", '"--update-manifest" in argv' in src)
        check("清单 URL 为 refs/tags 形式",
              "raw.githubusercontent.com/YehOlt/BatchReplace/refs/tags/updater/update.json" in src)
    finally:
        srv.shutdown()
        for stray in ("BatchReplace_v9.9.9.exe", "BatchReplace_v9.9.8.exe"):
            p = os.path.join(ROOT, stray)
            if os.path.exists(p):
                os.remove(p)
        print("\nTOTAL PASS: %d" % len(PASS))


if __name__ == "__main__":
    main()

# BatchReplace 第三方声明

本文件适用于 BatchReplace v2.0.0 及后续版本。BatchReplace 自身采用 MIT License，见同目录 `LICENSE`。

下列组件可能被打包进 `BatchReplace_v2.0.0.exe` 或随发行包提供。完整许可文本放在 `third_party_licenses/`，保留本文件及该目录即视为保留相应版权与许可声明。

---

## 运行时依赖

| 组件 | 版本 | 许可证 | 许可文本 |
|---|---:|---|---|
| openpyxl | 3.1.5 | MIT | `third_party_licenses/openpyxl/LICENCE.rst` |
| Pillow | 12.3.0 | MIT-CMU (HPND) | `third_party_licenses/pillow/LICENSE` |
| python-docx | 1.2.0 | MIT | `third_party_licenses/python-docx/LICENSE` |
| pywin32 | 312 | PSF（含多个子许可证） | `third_party_licenses/pywin32/` |
| et-xmlfile | 2.0.0 | MIT | `third_party_licenses/et_xmlfile/` |
| lxml | 6.1.1 | BSD-3-Clause | `third_party_licenses/lxml/` |
| typing-extensions | 4.16.0 | PSF-2.0 | `third_party_licenses/typing_extensions/LICENSE` |
| Python | 3.12.10 | PSF License Agreement | `third_party_licenses/python/LICENSE.txt` |

BatchReplace 未修改上述库的源代码，仅将其作为依赖调用或随 exe 打包。

---

## 构建工具

| 组件 | 版本 | 许可证 | 说明 |
|---|---:|---|---|
| PyInstaller | 6.22.3 | GPLv2-or-later with distribution exception | 分发例外允许用它构建并分发专有/商业程序；BatchReplace 自身代码不受 GPL 约束。完整文本见 `third_party_licenses/pyinstaller/COPYING.txt` |

PyInstaller 仅用于本机构建，最终用户无需安装。

---

## 系统组件

| 组件 | 说明 | 许可文本 |
|---|---|---|
| Microsoft Visual C++ Runtime / Universal CRT | 包含 `VCRUNTIME140.dll`、`VCRUNTIME140_1.dll`、`ucrtbase.dll`、`api-ms-win-*.dll`；按 Microsoft 再分发条款随程序分发 | 在线条款：https://learn.microsoft.com/cpp/windows/latest-supported-vc-redist |
| Tcl / Tk | `_tk_data\license.terms` 等许可文本已随 Python 标准库打入 exe，未额外修改 | 已内嵌于 exe |
| libffi | `libffi-8.dll`；MIT-style | `third_party_licenses/libffi/LICENSE.txt` |
| zlib | `zlib1.dll`；zlib License | `third_party_licenses/zlib/LICENSE.txt` |

bzip2、expat 等 Python 内置组件的许可文本已包含在 `third_party_licenses/python/LICENSE.txt` 中。

---

本文件仅用于说明第三方组件许可状态，不构成法律意见。

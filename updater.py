# -*- coding: utf-8 -*-
"""
updater.py — 客户端自动更新

工作流：
  1. 启动 1.5 秒后后台线程 GET PRIMARY_URL（VPS）→ 失败 fallback FALLBACK_URL（GitHub Releases）
  2. 拿到 latest.json，与本地 __version__ 比较；新版 → 主线程弹窗
  3. 用户确认 → 下载 Inno 安装包到 %TEMP% → SHA256 校验 → ShellExecute 启动 → 当前 app 退出

设计取舍：
  · 不引入 requests / packaging 依赖，全用 stdlib（urllib + hashlib），避免 PyInstaller 排除列表越来越长
  · 仅在 frozen 产物里默认启用；开发模式 (python main.py) 静默跳过，加 env MARK123_UPDATE_CHECK_DEV=1 可强制开
  · 网络 IO 全在 daemon 线程，UI 全在主线程，靠 pyqtSignal 跨线程传递结果
  · 任何网络/解析错误都静默吞掉，绝不打扰用户

落地前要改的地方：
  · PRIMARY_URL —— 改成你的 VPS 真实域名（建议 HTTPS）
  · GITHUB_REPO —— 已对，无需改
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import threading
import urllib.request
import urllib.error
from dataclasses import dataclass
from typing import Optional

from PyQt6.QtCore import QObject, QSettings, Qt, pyqtSignal
from PyQt6.QtWidgets import QMessageBox, QProgressDialog, QWidget, QApplication

# ─────────────────────── 配置 ───────────────────────
# 主源：自定义域名 → GitHub Pages（由 GitHub Actions 自动推送到 gh-pages 分支）
PRIMARY_URL = "https://zlzs.cc.cd/latest.json"

# 备源：GitHub Releases 的 "latest" 别名 —— 主域名 DNS 没生效 / GitHub Pages 故障时兜底
GITHUB_REPO = "vvangpc/mark123"
FALLBACK_URL = f"https://github.com/{GITHUB_REPO}/releases/latest/download/latest.json"

HTTP_TIMEOUT = 5            # 检查阶段
DOWNLOAD_TIMEOUT = 60        # 下载阶段（连接超时；读不超时由 read 阻塞）
USER_AGENT = "PatentMarker-Updater/1.0"

# QSettings 命名空间，必须与 config_manager.py 同步
_ORG = "PatentMarker"
_APP = "MarkAssistant"


# ─────────────────────── 数据模型 ───────────────────────

@dataclass
class UpdateInfo:
    version: str
    url: str          # VPS 直链
    url_github: str   # GitHub Releases 直链（备用）
    sha256: str
    size: int
    notes: str
    released_at: str

    @classmethod
    def from_json(cls, raw: dict) -> "UpdateInfo":
        return cls(
            version=str(raw["version"]).strip(),
            url=str(raw.get("url", "")).strip(),
            url_github=str(raw.get("url_github", "")).strip(),
            sha256=str(raw.get("sha256", "")).strip().lower(),
            size=int(raw.get("size", 0)),
            notes=str(raw.get("notes", "")).strip(),
            released_at=str(raw.get("released_at", "")).strip(),
        )


# ─────────────────────── 版本比较 ───────────────────────

def _parse_version(v: str) -> tuple:
    """'3.5' → (3, 5)；'3.5.1' → (3, 5, 1)；非数字段补 0 防崩。"""
    parts = []
    for seg in v.strip().lstrip("vV").split("."):
        try:
            parts.append(int(seg))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _is_newer(remote: str, local: str) -> bool:
    """remote 是否严格大于 local（即需要更新）"""
    return _parse_version(remote) > _parse_version(local)


# ─────────────────────── 后台检查 worker ───────────────────────

class _CheckWorker(QObject):
    """跑在后台线程，结果通过 signal 回主线程"""
    found = pyqtSignal(object)   # UpdateInfo
    no_update = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, current_version: str):
        super().__init__()
        self._current = current_version

    def run(self):
        info = _fetch_latest()
        if info is None:
            self.failed.emit("无法获取更新信息（网络不可达或服务器无响应）")
            return
        if not _is_newer(info.version, self._current):
            self.no_update.emit()
            return
        self.found.emit(info)


def _fetch_latest() -> Optional[UpdateInfo]:
    """先试 PRIMARY_URL，失败 fallback GitHub。任何异常都吞掉返回 None。"""
    for url in (PRIMARY_URL, FALLBACK_URL):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            return UpdateInfo.from_json(data)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError,
                json.JSONDecodeError, KeyError, ValueError, OSError):
            continue
    return None


# ─────────────────────── 主入口 ───────────────────────

class UpdateChecker(QObject):
    """生命周期挂在主窗口上，避免 worker / thread 提前 gc 掉。"""

    def __init__(self, parent_widget: QWidget, current_version: str):
        super().__init__(parent_widget)
        self._parent_widget = parent_widget
        self._current = current_version
        self._worker: Optional[_CheckWorker] = None
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        """启动后台检查。已经在跑则忽略。"""
        if self._thread is not None and self._thread.is_alive():
            return
        self._worker = _CheckWorker(self._current)
        # AutoConnection：worker 由 bg 线程 emit，slot 跑在主线程（parent 所在线程）
        self._worker.found.connect(self._on_found)
        self._worker.no_update.connect(self._on_no_update)
        self._worker.failed.connect(self._on_failed)
        self._thread = threading.Thread(
            target=self._worker.run, name="UpdateChecker", daemon=True
        )
        self._thread.start()

    # ── 主线程槽 ──
    def _on_found(self, info: UpdateInfo) -> None:
        # 用户曾点过「跳过此版本」？同版本静默
        settings = QSettings(_ORG, _APP)
        skipped = settings.value("updater/skip_version", "", type=str)
        if skipped == info.version:
            return
        self._prompt(info)

    def _on_no_update(self) -> None:
        # 静默，不打扰
        pass

    def _on_failed(self, _reason: str) -> None:
        # 静默，不打扰；如需排障可改成 print
        pass

    def _prompt(self, info: UpdateInfo) -> None:
        box = QMessageBox(self._parent_widget)
        box.setIcon(QMessageBox.Icon.Information)
        box.setWindowTitle("发现新版本")
        box.setText(f"<b>专利标记助手 V{info.version}</b> 已发布<br/>"
                    f"<span style='color:#888'>当前版本 V{self._current}</span>")
        if info.notes:
            box.setInformativeText(f"更新内容：\n{info.notes}")
        update_btn = box.addButton("立即更新", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("稍后提醒", QMessageBox.ButtonRole.RejectRole)
        skip_btn = box.addButton("跳过此版本", QMessageBox.ButtonRole.DestructiveRole)
        box.setDefaultButton(update_btn)
        box.exec()

        clicked = box.clickedButton()
        if clicked is update_btn:
            _download_and_launch(self._parent_widget, info)
        elif clicked is skip_btn:
            QSettings(_ORG, _APP).setValue("updater/skip_version", info.version)
        # 「稍后提醒」什么都不做，下次启动还会再问


# ─────────────────────── 下载 & 启动安装 ───────────────────────

def _download_and_launch(parent: QWidget, info: UpdateInfo) -> None:
    """阻塞式（带进度条）下载到 %TEMP%，校验 SHA256，启动 Inno 安装包，退出 app。"""
    # 仅 Windows 走自动下载/拉安装；其他平台只能提示用户
    if sys.platform != "win32":
        QMessageBox.information(
            parent, "请手动更新",
            f"自动更新仅支持 Windows。请到 {info.url_github or info.url} 手动下载。"
        )
        return

    target = os.path.join(
        tempfile.gettempdir(), f"专利标记助手V{info.version}-安装版.exe"
    )

    progress = QProgressDialog(
        f"正在下载 V{info.version} 安装包...", "取消", 0, 100, parent
    )
    progress.setWindowTitle("更新中")
    progress.setWindowModality(Qt.WindowModality.ApplicationModal)
    progress.setMinimumDuration(0)
    progress.setAutoClose(False)
    progress.setAutoReset(False)
    progress.show()

    sha = hashlib.sha256()
    bytes_done = 0
    total = info.size or 0
    cancelled = False

    # 优先 url（latest.json 里写的主下载源），失败 fallback url_github；去重避免重试同一个地址
    seen: set = set()
    urls: list = []
    for u in (info.url, info.url_github):
        if u and u not in seen:
            urls.append(u)
            seen.add(u)
    last_err: Optional[Exception] = None
    success = False

    for url in urls:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
                # 服务器没给 size 的话用 Content-Length 兜底
                if total <= 0:
                    cl = resp.headers.get("Content-Length")
                    if cl and cl.isdigit():
                        total = int(cl)
                        progress.setMaximum(total)
                else:
                    progress.setMaximum(total)

                with open(target, "wb") as f:
                    while True:
                        if progress.wasCanceled():
                            cancelled = True
                            break
                        chunk = resp.read(64 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        sha.update(chunk)
                        bytes_done += len(chunk)
                        if total > 0:
                            progress.setValue(min(bytes_done, total))
                        QApplication.processEvents()
            success = True
            break
        except Exception as e:
            last_err = e
            continue

    progress.close()

    if cancelled:
        try:
            os.remove(target)
        except OSError:
            pass
        return

    if not success:
        QMessageBox.warning(
            parent, "下载失败",
            f"无法下载更新包：{last_err}\n请稍后重试，或访问\n{info.url_github or info.url}"
        )
        try:
            os.remove(target)
        except OSError:
            pass
        return

    # SHA256 校验
    if info.sha256:
        digest = sha.hexdigest().lower()
        if digest != info.sha256:
            QMessageBox.critical(
                parent, "更新包校验失败",
                f"SHA256 不匹配，可能下载被中间人篡改或服务器文件损坏。\n"
                f"预期: {info.sha256}\n实际: {digest}\n\n"
                f"已删除可疑文件，请稍后重试。"
            )
            try:
                os.remove(target)
            except OSError:
                pass
            return

    # 拉起安装程序后立即退出当前进程，让 Inno 覆盖文件
    try:
        os.startfile(target)  # type: ignore[attr-defined]  # Windows only
    except OSError as e:
        QMessageBox.critical(
            parent, "启动安装程序失败",
            f"已下载到：\n{target}\n但无法自动启动：{e}\n请手动双击该文件继续安装。"
        )
        return

    QApplication.quit()


# ─────────────────────── 对外便捷函数 ───────────────────────

def should_check() -> bool:
    """frozen 产物默认开；开发模式默认关，可用 MARK123_UPDATE_CHECK_DEV=1 强开。"""
    if getattr(sys, "frozen", False):
        return True
    return os.environ.get("MARK123_UPDATE_CHECK_DEV") == "1"

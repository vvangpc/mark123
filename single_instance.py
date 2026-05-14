# -*- coding: utf-8 -*-
"""single_instance.py — 单实例守卫

启动入口分两步：
  1) try_send_to_running(path) → bool
     若已有实例在跑，把路径发过去并返回 True；调用方应立即 sys.exit(0)
  2) install_listener(window) → QLocalServer
     成为第一实例时挂监听，新连接到达时调用 window._receive_remote_file(path)

依赖 Qt 自带的 QLocalServer / QLocalSocket（Windows 底层是命名管道），无新增 pip 依赖。
"""
import os
import getpass

from PyQt6.QtNetwork import QLocalServer, QLocalSocket


def _user_id() -> str:
    try:
        return os.getlogin()
    except OSError:
        return getpass.getuser() or "default"


# 用户级唯一 server name，避免多用户 / Win 多 session 互踩
SERVER_NAME = f"PatentMarker_ipc_{_user_id()}"
_CONNECT_TIMEOUT_MS = 500
_WRITE_TIMEOUT_MS = 1000


def try_send_to_running(file_path: str) -> bool:
    """尝试把 file_path 发给已运行的实例。成功返回 True，调用方应立即退出。

    注意：不主动 disconnectFromServer —— 在 Windows 上，发送端立即断开会让
    服务端来不及把字节从管道里读出来就先收到 disconnected 信号。让 sys.exit
    / 进程退出去清理管道；只要数据已经 flush + waitForBytesWritten 完成，
    服务端就一定能读到。
    """
    sock = QLocalSocket()
    sock.connectToServer(SERVER_NAME)
    if not sock.waitForConnected(_CONNECT_TIMEOUT_MS):
        return False
    try:
        payload = (file_path or "").encode("utf-8") + b"\n"
        sock.write(payload)
        sock.flush()
        if not sock.waitForBytesWritten(_WRITE_TIMEOUT_MS):
            return False
        return True
    except Exception:
        return False


def install_listener(window) -> QLocalServer:
    """挂监听。新连接到达时调用 window._receive_remote_file(path)。"""
    # 清掉同名残留 socket（进程异常退出时偶有遗留）
    QLocalServer.removeServer(SERVER_NAME)
    server = QLocalServer(window)

    def _handle_payload(client: QLocalSocket):
        try:
            data = bytes(client.readAll()).decode("utf-8", errors="replace").strip()
        except Exception:
            data = ""
        if data:
            try:
                window._receive_remote_file(data)
            except Exception:
                pass
        try:
            client.disconnectFromServer()
        except Exception:
            pass

    def _on_new_conn():
        client = server.nextPendingConnection()
        if client is None:
            return
        # 关键：数据可能已经在 buffer 里（发送端发完就 disconnect 了），
        # 直接 connect readyRead 会错过；用 waitForReadyRead 做一次同步读，
        # 同时也连上后续的 readyRead 以应对分片到达
        client.readyRead.connect(lambda c=client: _handle_payload(c))
        client.disconnected.connect(client.deleteLater)
        if client.bytesAvailable() > 0 or client.waitForReadyRead(200):
            _handle_payload(client)

    server.newConnection.connect(_on_new_conn)
    server.listen(SERVER_NAME)
    return server

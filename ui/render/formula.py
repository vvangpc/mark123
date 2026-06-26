# -*- coding: utf-8 -*-
"""
ui/render/formula.py — WMF / EMF 矢量图（公式预览）→ QImage

专利里的公式多为 OLE 对象（MathType / AxMath / MS Equation），但每个对象都带一张
WMF/EMF 矢量预览（`<v:imagedata>`），**这张预览就是 Word 显示的公式图**。Qt 不能直接
解码 WMF/EMF，这里用 Windows GDI（ctypes，零新依赖）把它播放到位图再转 QImage。

设计：
  · 仅 win32 走 GDI；其它平台或任何异常 → 返回 None（调用方降级为【公式】占位，绝不崩）。
  · 统一先得到 HENHMETAFILE（EMF 直接载入；WMF 经 SetWinMetaFileBits 转 EMF），
    再 PlayEnhMetaFile 到一块白底 32 位 DIB，拷贝像素深拷贝成 QImage。
  · 内部按 supersample 倍率渲染再平滑缩回自然尺寸，保证清晰。
"""
import sys
import struct
import ctypes

from PyQt6.QtGui import QImage

_WMF_PLACEABLE_KEY = 0x9AC6CDD7      # 可放置 WMF 的 Aldus 头魔数
_EMF_SIGNATURE = 0x464D4520          # ' EMF'，位于 EMF 头偏移 40 处
_MM_ANISOTROPIC = 8
_BI_RGB = 0
_WHITE_BRUSH = 0


def metafile_to_qimage(blob: bytes, supersample: int = 2):
    """把 WMF/EMF 二进制渲染成自然尺寸的 QImage；失败返回 None。"""
    if sys.platform != "win32" or not blob:
        return None
    try:
        return _render(blob, supersample)
    except Exception:
        return None


# ── Windows 结构体 ──
if sys.platform == "win32":
    from ctypes import wintypes

    class _RECTL(ctypes.Structure):
        _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                    ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

    class _SIZEL(ctypes.Structure):
        _fields_ = [("cx", wintypes.LONG), ("cy", wintypes.LONG)]

    class _ENHMETAHEADER(ctypes.Structure):
        _fields_ = [
            ("iType", wintypes.DWORD), ("nSize", wintypes.DWORD),
            ("rclBounds", _RECTL), ("rclFrame", _RECTL),
            ("dSignature", wintypes.DWORD), ("nVersion", wintypes.DWORD),
            ("nBytes", wintypes.DWORD), ("nRecords", wintypes.DWORD),
            ("nHandles", wintypes.WORD), ("sReserved", wintypes.WORD),
            ("nDescription", wintypes.DWORD), ("offDescription", wintypes.DWORD),
            ("nPalEntries", wintypes.DWORD),
            ("szlDevice", _SIZEL), ("szlMillimeters", _SIZEL),
        ]

    class _METAFILEPICT(ctypes.Structure):
        _fields_ = [("mm", wintypes.LONG), ("xExt", wintypes.LONG),
                    ("yExt", wintypes.LONG), ("hMF", wintypes.HANDLE)]

    class _BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    class _RECT(ctypes.Structure):
        _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                    ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

    _gdi32 = ctypes.windll.gdi32
    _user32 = ctypes.windll.user32

    _gdi32.SetEnhMetaFileBits.restype = wintypes.HANDLE
    _gdi32.SetEnhMetaFileBits.argtypes = [wintypes.UINT, ctypes.c_char_p]
    _gdi32.SetWinMetaFileBits.restype = wintypes.HANDLE
    _gdi32.SetWinMetaFileBits.argtypes = [
        wintypes.UINT, ctypes.c_char_p, wintypes.HDC, ctypes.c_void_p]
    _gdi32.GetEnhMetaFileHeader.restype = wintypes.UINT
    _gdi32.GetEnhMetaFileHeader.argtypes = [
        wintypes.HANDLE, wintypes.UINT, ctypes.c_void_p]
    _gdi32.PlayEnhMetaFile.restype = wintypes.BOOL
    _gdi32.PlayEnhMetaFile.argtypes = [
        wintypes.HDC, wintypes.HANDLE, ctypes.c_void_p]
    _gdi32.DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
    _gdi32.CreateCompatibleDC.restype = wintypes.HDC
    _gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
    _gdi32.CreateDIBSection.restype = wintypes.HANDLE
    _gdi32.CreateDIBSection.argtypes = [
        wintypes.HDC, ctypes.c_void_p, wintypes.UINT,
        ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD]
    _gdi32.SelectObject.restype = wintypes.HANDLE
    _gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HANDLE]
    _gdi32.DeleteObject.argtypes = [wintypes.HANDLE]
    _gdi32.DeleteDC.argtypes = [wintypes.HDC]
    _gdi32.GetStockObject.restype = wintypes.HANDLE
    _gdi32.GetStockObject.argtypes = [ctypes.c_int]
    _gdi32.GdiFlush.argtypes = []
    _user32.GetDC.restype = wintypes.HDC
    _user32.GetDC.argtypes = [wintypes.HWND]
    _user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    _user32.FillRect.argtypes = [wintypes.HDC, ctypes.c_void_p, wintypes.HANDLE]


def _emf_handle_from_blob(blob: bytes):
    """返回 (HENHMETAFILE, 自然像素尺寸 或 None)。"""
    # EMF：偏移 40 处签名 ' EMF'
    if len(blob) >= 44 and struct.unpack_from("<I", blob, 40)[0] == _EMF_SIGNATURE:
        h = _gdi32.SetEnhMetaFileBits(len(blob), blob)
        return h, None

    data = blob
    natural = None
    mfp_ref = None
    mfp = None
    # 可放置 WMF：22 字节 Aldus 头（魔数 + bbox(twips/单位) + inch）
    if len(blob) >= 22 and struct.unpack_from("<I", blob, 0)[0] == _WMF_PLACEABLE_KEY:
        left, top, right, bottom = struct.unpack_from("<hhhh", blob, 6)
        inch = struct.unpack_from("<H", blob, 14)[0] or 1440
        data = blob[22:]
        w_units = abs(right - left)
        h_units = abs(bottom - top)
        if w_units and h_units:
            natural = (max(1, round(w_units / inch * 96)),
                       max(1, round(h_units / inch * 96)))
            mfp = _METAFILEPICT()
            mfp.mm = _MM_ANISOTROPIC
            mfp.xExt = round(w_units / inch * 2540)   # .01mm
            mfp.yExt = round(h_units / inch * 2540)
            mfp.hMF = None
            mfp_ref = ctypes.byref(mfp)

    hdc = _user32.GetDC(None)
    try:
        h = _gdi32.SetWinMetaFileBits(len(data), data, hdc, mfp_ref)
    finally:
        _user32.ReleaseDC(None, hdc)
    return h, natural


def _render(blob: bytes, supersample: int):
    hemf, natural = _emf_handle_from_blob(blob)
    if not hemf:
        return None
    try:
        if natural is None:
            hdr = _ENHMETAHEADER()
            _gdi32.GetEnhMetaFileHeader(hemf, ctypes.sizeof(hdr), ctypes.byref(hdr))
            bw = hdr.rclBounds.right - hdr.rclBounds.left + 1
            bh = hdr.rclBounds.bottom - hdr.rclBounds.top + 1
            if bw <= 1 or bh <= 1 or bw > 20000 or bh > 20000:
                fw = (hdr.rclFrame.right - hdr.rclFrame.left) / 100.0   # mm
                fh = (hdr.rclFrame.bottom - hdr.rclFrame.top) / 100.0
                bw = max(1, round(fw / 25.4 * 96))
                bh = max(1, round(fh / 25.4 * 96))
            natural = (int(bw), int(bh))
        nw = max(1, min(natural[0], 4000))
        nh = max(1, min(natural[1], 4000))
        ss = max(1, int(supersample))
        img = _play_to_qimage(hemf, nw * ss, nh * ss)
        if img is None:
            return None
        if ss != 1:
            from PyQt6.QtCore import Qt
            img = img.scaled(nw, nh, Qt.AspectRatioMode.IgnoreAspectRatio,
                             Qt.TransformationMode.SmoothTransformation)
        return img
    finally:
        _gdi32.DeleteEnhMetaFile(hemf)


def _play_to_qimage(hemf, w: int, h: int):
    hscreen = _user32.GetDC(None)
    hdc = _gdi32.CreateCompatibleDC(hscreen)
    _user32.ReleaseDC(None, hscreen)
    if not hdc:
        return None
    hbmp = None
    try:
        bmi = _BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
        bmi.biWidth = w
        bmi.biHeight = -h            # 负 = top-down
        bmi.biPlanes = 1
        bmi.biBitCount = 32
        bmi.biCompression = _BI_RGB
        ppv = ctypes.c_void_p()
        hbmp = _gdi32.CreateDIBSection(hdc, ctypes.byref(bmi), 0,
                                       ctypes.byref(ppv), None, 0)
        if not hbmp or not ppv.value:
            return None
        old = _gdi32.SelectObject(hdc, hbmp)
        rect = _RECT(0, 0, w, h)
        _user32.FillRect(hdc, ctypes.byref(rect), _gdi32.GetStockObject(_WHITE_BRUSH))
        _gdi32.PlayEnhMetaFile(hdc, hemf, ctypes.byref(rect))
        _gdi32.GdiFlush()
        nbytes = w * h * 4
        raw = ctypes.string_at(ppv.value, nbytes)
        # DIB 为 BGRA；用 RGB32 忽略 alpha（GDI 不写 alpha，避免全透明）
        img = QImage(raw, w, h, w * 4, QImage.Format.Format_RGB32).copy()
        _gdi32.SelectObject(hdc, old)
        return img
    finally:
        if hbmp:
            _gdi32.DeleteObject(hbmp)
        _gdi32.DeleteDC(hdc)

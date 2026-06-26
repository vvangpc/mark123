# -*- coding: utf-8 -*-
"""
ui/render/media.py — 段落内容遍历 + 图片/公式预览解析 + 缩放

把一个 docx 段落按文档顺序拆成「文本 / 图片」序列，供内容区构建富文本：
  · 真图（PNG/JPEG/…）：`a:blip@r:embed` 或 `v:imagedata@r:id` → related_parts[rId].blob → QImage。
  · 公式预览（WMF/EMF）：同样取 blob，QImage 解不了再交给 formula.metafile_to_qimage（GDI）。
  · 解析不到位图的对象（如无预览的 OMML）→ 占位文本（【公式】/【图】）。
文本部分之和 == para.text（只取 w:t / w:tab），保证回写/高亮偏移一致。
"""
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QImage

from ui.render.formula import metafile_to_qimage

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_V = "{urn:schemas-microsoft-com:vml}"

_OBJECT_TAGS = (f"{_W}drawing", f"{_W}object", f"{_W}pict")


def _find_rid(el):
    """在元素子树里找图片关系 id：优先 a:blip@r:embed，其次 v:imagedata@r:id。"""
    blip = el.find(f".//{_A}blip")
    if blip is not None:
        rid = blip.get(f"{_R}embed") or blip.get(f"{_R}link")
        if rid:
            return rid
    imd = el.find(f".//{_V}imagedata")
    if imd is not None:
        rid = imd.get(f"{_R}id")
        if rid:
            return rid
    return None


def _decode_rid(rid, document):
    """rId → 二进制 → QImage（先按位图，再按 WMF/EMF 矢量）。失败返回 None。"""
    try:
        blob = document.part.related_parts[rid].blob
    except Exception:
        return None
    img = QImage.fromData(blob)
    if not img.isNull():
        return img
    return metafile_to_qimage(blob)


def _image_from_element(el, document, cache):
    """从一个 drawing/object/pict 元素解析出 QImage；按 rId 缓存（含负缓存）。"""
    rid = _find_rid(el)
    if not rid:
        return None
    if rid in cache:
        return cache[rid]
    img = _decode_rid(rid, document)
    cache[rid] = img
    return img


def has_renderable_object(para, document, cache) -> bool:
    """该段是否含可渲染的图片/公式预览（顺带填充缓存）。"""
    for el in para._p.iter():
        if el.tag in _OBJECT_TAGS:
            if _image_from_element(el, document, cache) is not None:
                return True
    return False


def iter_content(para, document, cache):
    """按文档顺序产出 (kind, payload)：
       ('text', str) / ('image', QImage) / ('placeholder', str)。
    """
    for child in para._p.iterchildren():
        tag = child.tag
        if tag == f"{_W}r":
            for rc in child.iterchildren():
                rtag = rc.tag
                if rtag == f"{_W}t":
                    yield ("text", rc.text or "")
                elif rtag == f"{_W}tab":
                    yield ("text", "\t")
                elif rtag in _OBJECT_TAGS:
                    img = _image_from_element(rc, document, cache)
                    if img is not None:
                        yield ("image", img)
                    else:
                        ph = "【图】" if rtag == f"{_W}drawing" else "【公式】"
                        yield ("placeholder", ph)
                # w:br / w:cr 忽略，避免给一个段引入额外换行
        elif tag in _OBJECT_TAGS:
            img = _image_from_element(child, document, cache)
            yield ("image", img) if img is not None else ("placeholder", "【图】")
        elif tag in (f"{_M}oMath", f"{_M}oMathPara"):
            # OMML 公式：无现成预览位图（本项目样例无此类）→ 占位
            yield ("placeholder", "【公式】")
        # pPr 等其它子元素忽略


def scale_to_width(img: QImage, max_w: int) -> QImage:
    """图片宽超出 max_w 时平滑缩放到 max_w（保持纵横比）；否则原图。"""
    if max_w > 0 and img.width() > max_w:
        return img.scaledToWidth(max_w, Qt.TransformationMode.SmoothTransformation)
    return img

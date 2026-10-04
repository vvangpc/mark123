# -*- coding: utf-8 -*-
"""
ui/render/media.py — 段落内容遍历 + 图片/公式预览解析 + 缩放

把一个 docx 段落按文档顺序拆成「文本 / 图片」序列，供内容区构建富文本：
  · 真图（PNG/JPEG/…）：`a:blip@r:embed` 或 `v:imagedata@r:id` → related_parts[rId].blob → QImage。
  · 公式预览（WMF/EMF）：同样取 blob，QImage 解不了再交给 formula.metafile_to_qimage（GDI）；
    这类"白底黑线"的灰阶预览会被 _recolor_lineart 转成透明底 + 主题色，免得在深色
    主题下变成正文里一排排白盒子（彩色图不动）。
  · OMML 公式（Word「插入→公式」）：**没有任何预览位图**，交给 omml.omml_to_qimage
    现场排版成透明底图片（见 ui/render/omml.py）。
  · 实在解析不出的对象 → 占位文本（【公式】/【图】）。
文本部分之和 == core.paragraph_edit.display_text，保证回写/高亮偏移一致。
"""
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QImage, QPainter

from core.paragraph_edit import token_text
from ui.render.formula import metafile_to_qimage
from ui.render.omml import omml_to_qimage

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_V = "{urn:schemas-microsoft-com:vml}"

_OBJECT_TAGS = (f"{_W}drawing", f"{_W}object", f"{_W}pict")
_OMATH_TAGS = (f"{_M}oMath", f"{_M}oMathPara")

# 公式排版的默认取色 / 字号；内容区按当前主题覆盖（见 ContentArea.set_math_color）
DEFAULT_MATH_COLOR = "#2c3e50"
DEFAULT_MATH_PX = 16.0


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
    """rId → 二进制 → (QImage, 是否矢量预览)。失败返回 (None, False)。"""
    try:
        blob = document.part.related_parts[rid].blob
    except Exception:
        return None, False
    img = QImage.fromData(blob)
    if not img.isNull():
        return img, False                     # PNG/JPEG 等真位图
    return metafile_to_qimage(blob), True     # WMF/EMF 矢量预览（公式多走这条）


# 逐字节取反表：用 bytes.translate 一次完成整幅图的"亮度 → 不透明度"反转，
# 比 Python 逐像素快几个数量级（110 个公式预览的量级下这是必须的）。
_INVERT_TABLE = bytes(range(255, -1, -1))


def _is_grayscale(img: QImage, samples: int = 24) -> bool:
    """抽样判断是否基本是灰阶（公式预览一律黑线白底；彩色附图必须原样保留）。"""
    w, h = img.width(), img.height()
    if w <= 0 or h <= 0:
        return False
    sx = max(1, w // samples)
    sy = max(1, h // samples)
    for y in range(0, h, sy):
        for x in range(0, w, sx):
            c = img.pixelColor(x, y)
            if max(c.red(), c.green(), c.blue()) - min(c.red(), c.green(), c.blue()) > 24:
                return False
    return True


def _recolor_lineart(img: QImage, color: str):
    """白底黑线的矢量预览 → 透明底 + 主题色线条；不适用时返回 None。

    WMF/EMF 预览是"白底黑字"的位图，深色主题下就成了正文里一排排刺眼的白盒子。
    这里把亮度反过来当不透明度（白 → 全透明，黑 → 全不透明），再整体填成主题色，
    观感与现场排版的 OMML 公式一致。只对灰阶图生效，彩色附图不碰。
    """
    try:
        if img is None or img.isNull() or not _is_grayscale(img):
            return None
        gray = img.convertToFormat(QImage.Format.Format_Grayscale8)
        buf = gray.bits()
        buf.setsize(gray.sizeInBytes())
        alpha_bytes = bytes(buf).translate(_INVERT_TABLE)
        alpha = QImage(alpha_bytes, gray.width(), gray.height(),
                       gray.bytesPerLine(), QImage.Format.Format_Alpha8).copy()
        out = QImage(img.width(), img.height(),
                     QImage.Format.Format_ARGB32_Premultiplied)
        out.fill(Qt.GlobalColor.transparent)
        p = QPainter(out)
        try:
            p.fillRect(0, 0, out.width(), out.height(), QColor(color))
            p.setCompositionMode(
                QPainter.CompositionMode.CompositionMode_DestinationIn)
            p.drawImage(0, 0, alpha)
        finally:
            p.end()
        return out
    except Exception:
        return None


def _image_from_element(el, document, cache, color: str = None):
    """从 drawing/object/pict 元素解析出 QImage。

    缓存两层：`rid` 存解码结果（与配色无关，最贵的一步只做一次），
    `(rid, color)` 存按主题重新着色后的成品（主题切换时只丢这一层）。
    """
    rid = _find_rid(el)
    if not rid:
        return None
    ckey = (rid, color)
    if ckey in cache:
        return cache[ckey]
    if rid in cache:
        raw, is_vector = cache[rid]
    else:
        raw, is_vector = _decode_rid(rid, document)
        cache[rid] = (raw, is_vector)
    img = raw
    if raw is not None and is_vector and color:
        img = _recolor_lineart(raw, color) or raw
    cache[ckey] = img
    return img


def _omml_image(el, cache, color: str, px: float):
    """OMML → QImage；按 (XML, 取色, 字号) 缓存，含负缓存。

    键里带 XML 而不是 id(el)：lxml 的元素代理是临时对象，id 会被复用，
    拿 id 当键会张冠李戴。
    """
    key = None
    try:
        from lxml import etree
        key = ("omml", etree.tostring(el), color, px)
    except Exception:
        pass
    if key is not None and key in cache:
        return cache[key]
    img = omml_to_qimage(el, color=color, px=px)
    if img is not None and img.isNull():
        img = None
    if key is not None:
        cache[key] = img
    return img


def has_renderable_object(para, document, cache) -> bool:
    """该段是否含可渲染的图片/公式（顺带填充图片缓存）。

    OMML 一律算「有」：它没有预览位图，但能现场排版，最差也会降级成一行文字，
    所以必须走富文本路径 —— 否则纯文本路径会让公式在 1框 里彻底消失
    （para.text 只收 w:t，不含公式里的 m:t）。
    """
    for el in para._p.iter():
        if el.tag in _OMATH_TAGS:
            return True
        if el.tag in _OBJECT_TAGS:
            if _image_from_element(el, document, cache) is not None:
                return True
    return False


def iter_content(para, document, cache,
                 math_color: str = DEFAULT_MATH_COLOR,
                 math_px: float = DEFAULT_MATH_PX):
    """按文档顺序产出 (kind, payload)：
       ('text', str) / ('image', QImage) / ('placeholder', str)。
    """
    for child in para._p.iterchildren():
        tag = child.tag
        if tag in (f"{_W}r", f"{_W}hyperlink"):
            runs = [child] if tag == f"{_W}r" else child.findall(f"{_W}r")
            for run in runs:
                for rc in run.iterchildren():
                    rtag = rc.tag
                    if rtag == f"{_W}t":
                        yield ("text", rc.text or "")
                    elif rtag in _OBJECT_TAGS:
                        img = _image_from_element(rc, document, cache, math_color)
                        if img is not None:
                            yield ("image", img)
                        else:
                            ph = "【图】" if rtag == f"{_W}drawing" else "【公式】"
                            yield ("placeholder", ph)
                    else:
                        # 制表符 / 软回车（显示为「↵」，不增行）等结构元素
                        tok = token_text(rc)
                        if tok:
                            yield ("text", tok)
        elif tag in _OBJECT_TAGS:
            img = _image_from_element(child, document, cache, math_color)
            yield ("image", img) if img is not None else ("placeholder", "【图】")
        elif tag in _OMATH_TAGS:
            # OMML 公式：Word 不给预览位图 → 现场排版（ui/render/omml.py）
            children = ([child] if tag == f"{_M}oMath"
                        else child.findall(f"{_M}oMath"))
            for om in children:
                img = _omml_image(om, cache, math_color, math_px)
                yield ("image", img) if img is not None else ("placeholder", "【公式】")
        # pPr 等其它子元素忽略


def scale_to_width(img: QImage, max_w: int) -> QImage:
    """图片宽超出 max_w 时平滑缩放到 max_w（保持纵横比）；否则原图。"""
    if max_w > 0 and img.width() > max_w:
        return img.scaledToWidth(max_w, Qt.TransformationMode.SmoothTransformation)
    return img

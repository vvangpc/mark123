# -*- coding: utf-8 -*-
"""
ui/render/omml.py — OMML（Office Math，m:oMath）→ QImage

Word 的「插入 → 公式」写出来的是 OMML XML，**不带任何预览位图** —— 这点和
MathType / 公式编辑器那类 OLE 对象正好相反（后者每个对象都附一张 WMF 预览，
所以 formula.py 用 GDI 播放一下就能显示）。media.py 原先对 OMML 只能给一个
【公式】占位，这正是「Office Math 公式无法预览」的根因。

这里用 QPainter 实现一个小型数学排版引擎：把 OMML 递归排成盒模型
（宽 w / 基线上高 asc / 基线下深 desc），再画到一张**透明底** QImage 上，
颜色由调用方按主题传入。覆盖专利文件里几乎全部写法：

    m:r,m:t 文本     m:sSub 下标      m:sSup 上标       m:sSubSup 上下标
    m:sPre 前置角标  m:f 分式         m:rad 根式        m:d 括号（随内容伸缩）
    m:nary 大算符    m:m 矩阵         m:func 函数名     m:limLow/limUpp 上下限
    m:acc 重音       m:bar 上下划线   m:groupChar       m:borderBox 框
    m:eqArr 多行     m:box/m:phant 透传

不认识的元素一律降级成「把里面的 m:t 文字排成一行」，绝不返回空白。
零新依赖：只用 PyQt6 自带的 QFont / QPainter。
"""
import math

from PyQt6.QtCore import Qt, QPointF, QRectF
from PyQt6.QtGui import (
    QColor, QFont, QFontDatabase, QFontMetricsF, QImage, QPainter,
    QPainterPath, QPen,
)

_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# 数学字体优先级：Cambria Math 是 Word 公式的默认字体，Windows 自带；
# 其余是常见的开源数学字体与兜底衬线字体。
_FAMILY_CANDIDATES = (
    "Cambria Math", "STIX Two Math", "XITS Math", "Latin Modern Math",
    "DejaVu Math TeX Gyre", "Times New Roman", "Georgia", "serif",
)
_family_cache = None


def _math_family() -> str:
    global _family_cache
    if _family_cache is None:
        try:
            available = set(QFontDatabase.families())
        except Exception:
            available = set()
        _family_cache = next(
            (f for f in _FAMILY_CANDIDATES if f in available), "serif"
        )
    return _family_cache


_font_cache = {}


def _font(sz: float, italic: bool = False, bold: bool = False) -> QFont:
    key = (round(sz, 1), italic, bold)
    f = _font_cache.get(key)
    if f is None:
        f = QFont(_math_family())
        f.setPixelSize(max(6, int(round(sz))))
        f.setItalic(italic)
        f.setBold(bold)
        _font_cache[key] = f
    return f


def _local(tag) -> str:
    """取无命名空间的标签名；注释 / PI 节点的 tag 不是 str，返回空串。"""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _val(el, name: str, default=None):
    """读 <m:xxx m:val="…"/> 的属性值；name 可为 'naryPr/chr' 这样的路径。"""
    node = el.find("/".join(_M + part for part in name.split("/")))
    if node is None:
        return default
    v = node.get(_M + "val")
    return default if v is None else v


def _has(el, name: str) -> bool:
    return el.find("/".join(_M + part for part in name.split("/"))) is not None


def _on(v) -> bool:
    """OMML 里布尔属性的几种写法。"""
    return v in ("1", "true", "on")


# ─────────────────────────────────────────
# 盒模型
# ─────────────────────────────────────────
class _Box:
    """排版盒：宽 w、基线以上高 asc、基线以下深 desc；draw 传入基线点。"""

    w = 0.0
    asc = 0.0
    desc = 0.0

    def draw(self, p: QPainter, x: float, y: float) -> None:
        pass


class _Space(_Box):
    def __init__(self, w: float, asc: float = 0.0, desc: float = 0.0):
        self.w, self.asc, self.desc = w, asc, desc


class _Glyph(_Box):
    """一段同字体的文字。

    用 fm.ascent()/descent() 而不是 tightBoundingRect —— 后者随字符不同而变，
    同一行里会把基线拉得高低不齐。
    """

    def __init__(self, text: str, font: QFont):
        self.text = text
        self.font = font
        fm = QFontMetricsF(font)
        self.w = fm.horizontalAdvance(text)
        self.asc = fm.ascent()
        self.desc = fm.descent()

    def draw(self, p, x, y):
        p.setFont(self.font)
        p.drawText(QPointF(x, y), self.text)


class _Row(_Box):
    """水平排列。"""

    def __init__(self, boxes, gap: float = 0.0):
        self.boxes = [b for b in boxes if b is not None]
        self.gap = gap
        n = len(self.boxes)
        self.w = sum(b.w for b in self.boxes) + gap * max(0, n - 1)
        self.asc = max((b.asc for b in self.boxes), default=0.0)
        self.desc = max((b.desc for b in self.boxes), default=0.0)

    def draw(self, p, x, y):
        for b in self.boxes:
            b.draw(p, x, y)
            x += b.w + self.gap


class _Stack(_Box):
    """若干盒子垂直堆叠，整体绕数学轴居中（矩阵行 / 上下限 / 多行公式共用）。"""

    def __init__(self, boxes, sz: float, vgap: float = None, align: str = "center"):
        self.boxes = [b for b in boxes if b is not None]
        self.vgap = sz * 0.32 if vgap is None else vgap
        self.align = align
        self.w = max((b.w for b in self.boxes), default=0.0)
        total = sum(b.asc + b.desc for b in self.boxes)
        total += self.vgap * max(0, len(self.boxes) - 1)
        axis = sz * 0.30
        self.asc = total / 2.0 + axis
        self.desc = max(0.0, total / 2.0 - axis)

    def draw(self, p, x, y):
        top = y - self.asc
        for b in self.boxes:
            if self.align == "left":
                bx = x
            elif self.align == "right":
                bx = x + self.w - b.w
            else:
                bx = x + (self.w - b.w) / 2.0
            b.draw(p, bx, top + b.asc)
            top += b.asc + b.desc + self.vgap


class _Frac(_Box):
    def __init__(self, num: _Box, den: _Box, sz: float, bar: bool = True):
        self.num, self.den, self.sz, self.bar = num, den, sz, bar
        self.pad = sz * 0.18
        self.axis = sz * 0.30                 # 分数线在基线上方的高度
        self.gap = sz * 0.16
        self.w = max(num.w, den.w) + 2 * self.pad
        self.asc = self.axis + self.gap + num.desc + num.asc
        self.desc = max(0.0, -self.axis + self.gap + den.asc + den.desc)

    def draw(self, p, x, y):
        cx = x + self.w / 2.0
        ybar = y - self.axis
        self.num.draw(p, cx - self.num.w / 2.0, ybar - self.gap - self.num.desc)
        self.den.draw(p, cx - self.den.w / 2.0, ybar + self.gap + self.den.asc)
        if self.bar:
            old = p.pen()
            pen = QPen(old)
            pen.setWidthF(max(1.0, self.sz * 0.06))
            p.setPen(pen)
            p.drawLine(QPointF(x + self.pad * 0.4, ybar),
                       QPointF(x + self.w - self.pad * 0.4, ybar))
            p.setPen(old)


class _Script(_Box):
    """base 带上/下标（sub / sup 任一可为 None）；pre=True 时角标画在左侧。"""

    def __init__(self, base: _Box, sub: _Box, sup: _Box, sz: float, pre: bool = False):
        self.base, self.sub, self.sup, self.pre = base, sub, sup, pre
        self.sup_dy = sz * 0.46
        self.sub_dy = sz * 0.22
        self.kern = sz * 0.06
        sw = max(sub.w if sub is not None else 0.0,
                 sup.w if sup is not None else 0.0)
        self.sw = sw
        self.w = base.w + sw + (self.kern if sw else 0.0)
        self.asc = base.asc
        self.desc = base.desc
        if sup is not None:
            self.asc = max(self.asc, self.sup_dy + sup.asc)
        if sub is not None:
            self.desc = max(self.desc, self.sub_dy + sub.desc)

    def draw(self, p, x, y):
        if self.pre:
            sx, bx = x, x + self.sw + (self.kern if self.sw else 0.0)
        else:
            bx = x
            sx = x + self.base.w + (self.kern if self.sw else 0.0)
        self.base.draw(p, bx, y)
        if self.sup is not None:
            self.sup.draw(p, sx, y - self.sup_dy)
        if self.sub is not None:
            self.sub.draw(p, sx, y + self.sub_dy)


class _Radical(_Box):
    def __init__(self, body: _Box, deg: _Box, sz: float):
        self.body, self.deg, self.sz = body, deg, sz
        self.gap_top = sz * 0.16
        self.lead = sz * 0.62 + (deg.w if deg is not None else 0.0)
        self.w = self.lead + body.w + sz * 0.14
        self.asc = body.asc + self.gap_top + sz * 0.10
        self.desc = body.desc

    def draw(self, p, x, y):
        top = y - self.asc
        bottom = y + self.desc
        h = bottom - top
        x0 = x + (self.deg.w if self.deg is not None else 0.0)
        path = QPainterPath()
        path.moveTo(x0, bottom - h * 0.42)
        path.lineTo(x0 + self.sz * 0.18, bottom - h * 0.50)
        path.lineTo(x0 + self.sz * 0.36, bottom)
        path.lineTo(x0 + self.sz * 0.60, top)
        path.lineTo(x + self.w, top)
        old = p.pen()
        pen = QPen(old)
        pen.setWidthF(max(1.0, self.sz * 0.065))
        pen.setJoinStyle(Qt.PenJoinStyle.MiterJoin)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawPath(path)
        p.setPen(old)
        if self.deg is not None:
            self.deg.draw(p, x, top + h * 0.42)
        self.body.draw(p, x + self.lead, y)


class _Fence(_Box):
    """可伸缩括号：按内容高度放大字号，绕数学轴垂直居中。"""

    def __init__(self, ch: str, height: float, axis: float, sz: float):
        self.ch = ch
        rect0 = QFontMetricsF(_font(sz)).tightBoundingRect(ch)
        gh = rect0.height() or sz
        scale = min(6.0, max(1.0, height / gh))
        self.font = _font(sz * scale)
        fm = QFontMetricsF(self.font)
        self.rect = fm.tightBoundingRect(ch)
        self.w = fm.horizontalAdvance(ch) or (sz * 0.35)
        self.axis = axis
        self.asc = height / 2.0 + axis
        self.desc = max(0.0, height / 2.0 - axis)

    def draw(self, p, x, y):
        p.setFont(self.font)
        cy = y - self.axis
        # tightBoundingRect.top() 是相对基线的负偏移 → 反推使字形视觉中心落在 cy
        by = cy - (self.rect.top() + self.rect.height() / 2.0)
        p.drawText(QPointF(x, by), self.ch)


class _Grid(_Box):
    """矩阵：按列宽 / 行高对齐的网格，整体绕数学轴居中。"""

    def __init__(self, rows, sz: float):
        self.rows = rows
        self.hgap = sz * 0.55
        self.vgap = sz * 0.30
        ncol = max((len(r) for r in rows), default=0)
        self.colw = [
            max((r[c].w for r in rows if c < len(r)), default=0.0)
            for c in range(ncol)
        ]
        self.rowm = [
            (max((c.asc for c in r), default=0.0),
             max((c.desc for c in r), default=0.0))
            for r in rows
        ]
        self.w = sum(self.colw) + self.hgap * max(0, ncol - 1)
        total = sum(a + d for a, d in self.rowm) + self.vgap * max(0, len(rows) - 1)
        axis = sz * 0.30
        self.asc = total / 2.0 + axis
        self.desc = max(0.0, total / 2.0 - axis)

    def draw(self, p, x, y):
        top = y - self.asc
        for r, (ra, rd) in zip(self.rows, self.rowm):
            cx = x
            for i, cell in enumerate(r):
                cell.draw(p, cx + (self.colw[i] - cell.w) / 2.0, top + ra)
                cx += self.colw[i] + self.hgap
            top += ra + rd + self.vgap


class _Line(_Box):
    """给内容加上划线 / 下划线（m:bar）。"""

    def __init__(self, body: _Box, sz: float, pos: str = "top"):
        self.body, self.sz, self.pos = body, sz, pos
        self.pad = sz * 0.16
        self.w = body.w
        self.asc = body.asc + (self.pad if pos == "top" else 0.0)
        self.desc = body.desc + (self.pad if pos != "top" else 0.0)

    def draw(self, p, x, y):
        self.body.draw(p, x, y)
        old = p.pen()
        pen = QPen(old)
        pen.setWidthF(max(1.0, self.sz * 0.06))
        p.setPen(pen)
        if self.pos == "top":
            ly = y - self.asc + self.sz * 0.04
        else:
            ly = y + self.desc - self.sz * 0.04
        p.drawLine(QPointF(x, ly), QPointF(x + self.w, ly))
        p.setPen(old)


class _Border(_Box):
    def __init__(self, body: _Box, sz: float):
        self.body, self.sz = body, sz
        self.pad = sz * 0.18
        self.w = body.w + 2 * self.pad
        self.asc = body.asc + self.pad
        self.desc = body.desc + self.pad

    def draw(self, p, x, y):
        self.body.draw(p, x + self.pad, y)
        old = p.pen()
        pen = QPen(old)
        pen.setWidthF(max(1.0, self.sz * 0.05))
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(QRectF(x, y - self.asc, self.w, self.asc + self.desc))
        p.setPen(old)


class _AccBox(_Box):
    """重音符号画在内容正上方（不改变基线）。"""

    def __init__(self, body: _Box, mark: _Box, sz: float):
        self.body, self.mark, self.sz = body, mark, sz
        self.w = max(body.w, mark.w)
        self.asc = body.asc + sz * 0.22
        self.desc = body.desc

    def draw(self, p, x, y):
        self.body.draw(p, x + (self.w - self.body.w) / 2.0, y)
        self.mark.draw(p, x + (self.w - self.mark.w) / 2.0,
                       y - self.body.asc - self.sz * 0.02)


# ─────────────────────────────────────────
# 文本：斜体规则 + 运算符间距
# ─────────────────────────────────────────
# 数学变量默认斜体（拉丁字母与小写希腊字母）；数字 / 汉字 / 运算符保持正体
_ITALIC_CHARS = set(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    "αβγδεζηθικλμνξοπρστυφχψω"
)
# 二元运算符两侧留细空 —— Word 自己也这么排，不留会挤成一团
_BINARY_OPS = set("=+-−±∓×÷⋅∙*/<>≤≥≈≠≡≅∝→←↔⇒⇔∈∉⊂⊃⊆⊇∪∩∧∨")


def _text_boxes(text: str, sz: float, italic: bool, bold: bool):
    """把一段公式文字拆成若干 _Glyph（按斜体 / 正体分段）+ 运算符两侧细空。"""
    out = []
    buf = []
    state = {"italic": False}

    def flush():
        if buf:
            out.append(_Glyph("".join(buf), _font(sz, state["italic"], bold)))
            del buf[:]

    for ch in text:
        if ch in _BINARY_OPS:
            flush()
            out.append(_Space(sz * 0.15))
            out.append(_Glyph(ch, _font(sz, False, bold)))
            out.append(_Space(sz * 0.15))
            continue
        it = italic and ch in _ITALIC_CHARS
        if buf and it != state["italic"]:
            flush()
        state["italic"] = it
        buf.append(ch)
    flush()
    return out


def _run(el, sz: float) -> _Box:
    """m:r —— 一个公式文本运行。

    斜体规则：默认变量斜体；m:nor 表示正体；m:sty 可显式指定 p/b/i/bi。
    """
    text = "".join(t.text or "" for t in el.findall(_M + "t"))
    if not text:
        return _Space(0.0)
    sty = _val(el, "rPr/sty")
    italic = not _has(el, "rPr/nor")
    bold = False
    if sty == "p":
        italic = False
    elif sty == "b":
        italic, bold = False, True
    elif sty == "i":
        italic = True
    elif sty == "bi":
        italic, bold = True, True
    return _Row(_text_boxes(text, sz, italic, bold))


# ─────────────────────────────────────────
# 各 OMML 结构的排版
# ─────────────────────────────────────────
def _scr_sz(sz: float) -> float:
    """角标字号：逐层缩小但不小于 7px，避免深层嵌套糊成一片。"""
    return max(7.0, sz * 0.72)


def _seq(el, sz: float) -> _Box:
    """按文档顺序排一个容器元素的子元素（oMath / e / num / den / lim / …）。"""
    boxes = []
    for child in el.iterchildren():
        name = _local(child.tag)
        if not name or name.endswith("Pr"):      # ctrlPr / fPr / rPr … 只带属性
            continue
        boxes.append(_layout(child, sz))
    return _Row(boxes)


def _arg(el, name: str, sz: float):
    """取子元素（m:e / m:sub / m:sup / m:lim …）并排版；不存在返回 None。"""
    node = el.find(_M + name)
    return None if node is None else _seq(node, sz)


def _f(el, sz):
    kind = _val(el, "fPr/type", "bar")
    num = _arg(el, "num", sz) or _Space(0.0)
    den = _arg(el, "den", sz) or _Space(0.0)
    if kind == "lin":                             # 线性分式 a/b
        return _Row([num, _Glyph("/", _font(sz)), den])
    return _Frac(num, den, sz, bar=(kind != "noBar"))


def _sSub(el, sz):
    return _Script(_arg(el, "e", sz) or _Space(0.0),
                   _arg(el, "sub", _scr_sz(sz)), None, sz)


def _sSup(el, sz):
    return _Script(_arg(el, "e", sz) or _Space(0.0),
                   None, _arg(el, "sup", _scr_sz(sz)), sz)


def _sSubSup(el, sz):
    return _Script(_arg(el, "e", sz) or _Space(0.0),
                   _arg(el, "sub", _scr_sz(sz)),
                   _arg(el, "sup", _scr_sz(sz)), sz)


def _sPre(el, sz):
    return _Script(_arg(el, "e", sz) or _Space(0.0),
                   _arg(el, "sub", _scr_sz(sz)),
                   _arg(el, "sup", _scr_sz(sz)), sz, pre=True)


def _rad(el, sz):
    body = _arg(el, "e", sz) or _Space(0.0)
    deg = None
    if not _on(_val(el, "radPr/degHide")):
        node = el.find(_M + "deg")
        if node is not None and "".join(t.text or "" for t in node.iter(_M + "t")):
            deg = _seq(node, _scr_sz(_scr_sz(sz)))
    return _Radical(body, deg, sz)


def _d(el, sz):
    beg = _val(el, "dPr/begChr", "(")
    end = _val(el, "dPr/endChr", ")")
    sep = _val(el, "dPr/sepChr", "|")
    args = [_seq(e, sz) for e in el.findall(_M + "e")] or [_Space(0.0)]
    inner = []
    for i, a in enumerate(args):
        if i:
            inner.append(_Glyph(sep, _font(sz)))
        inner.append(a)
    body = _Row(inner)
    axis = sz * 0.30
    height = max(body.asc + body.desc, sz * 1.0)
    out = []
    if beg:
        out.append(_Fence(beg, height, axis, sz))
    out.append(body)
    if end:
        out.append(_Fence(end, height, axis, sz))
    return _Row(out)


# 求和 / 求积 / 并交类算符的上下限默认画在算符上下方；积分类默认画成角标
_UNDOVR_CHARS = set("∑∏∐⋀⋁⋂⋃⨁⨂⨀")


def _nary(el, sz):
    ch = _val(el, "naryPr/chr", "∫")
    loc = _val(el, "naryPr/limLoc")
    body = _arg(el, "e", sz) or _Space(0.0)
    sub = None if _on(_val(el, "naryPr/subHide")) else _arg(el, "sub", _scr_sz(sz))
    sup = None if _on(_val(el, "naryPr/supHide")) else _arg(el, "sup", _scr_sz(sz))
    op = _Glyph(ch, _font(sz * 1.7))
    if loc == "undOvr" or (loc is None and ch in _UNDOVR_CHARS):
        parts = [b for b in (sup, op, sub) if b is not None]
        stacked = _Stack(parts, sz, vgap=sz * 0.06) if len(parts) > 1 else op
        return _Row([stacked, _Space(sz * 0.12), body])
    return _Row([_Script(op, sub, sup, sz), _Space(sz * 0.12), body])


def _m(el, sz):
    rows = []
    for mr in el.findall(_M + "mr"):
        rows.append([_seq(e, sz) for e in mr.findall(_M + "e")])
    return _Grid(rows, sz) if rows else _Space(0.0)


def _func(el, sz):
    name = el.find(_M + "fName")
    body = _arg(el, "e", sz) or _Space(0.0)
    if name is None:
        return body
    return _Row([_seq(name, sz), _Space(sz * 0.14), body])


def _limLow(el, sz):
    base = _arg(el, "e", sz) or _Space(0.0)
    lim = _arg(el, "lim", _scr_sz(sz))
    return _Stack([base, lim], sz, vgap=sz * 0.08)


def _limUpp(el, sz):
    base = _arg(el, "e", sz) or _Space(0.0)
    lim = _arg(el, "lim", _scr_sz(sz))
    return _Stack([lim, base], sz, vgap=sz * 0.08)


def _acc(el, sz):
    ch = _val(el, "accPr/chr", "^")
    body = _arg(el, "e", sz) or _Space(0.0)
    return _AccBox(body, _Glyph(ch, _font(sz * 0.9)), sz)


def _bar(el, sz):
    pos = _val(el, "barPr/pos", "bot")
    return _Line(_arg(el, "e", sz) or _Space(0.0), sz,
                 "top" if pos == "top" else "bot")


def _groupChar(el, sz):
    ch = _val(el, "groupChrPr/chr", "⏟")
    pos = _val(el, "groupChrPr/pos", "bot")
    body = _arg(el, "e", sz) or _Space(0.0)
    mark = _Glyph(ch, _font(sz))
    order = [mark, body] if pos == "top" else [body, mark]
    return _Stack(order, sz, vgap=sz * 0.02)


def _borderBox(el, sz):
    return _Border(_arg(el, "e", sz) or _Space(0.0), sz)


def _eqArr(el, sz):
    rows = [_seq(e, sz) for e in el.findall(_M + "e")]
    return _Stack(rows, sz, vgap=sz * 0.28, align="left") if rows else _Space(0.0)


_LAYOUT = {
    "oMath": _seq, "oMathPara": _seq, "e": _seq, "num": _seq, "den": _seq,
    "lim": _seq, "fName": _seq, "deg": _seq, "box": _seq, "phant": _seq,
    "r": _run,
    "f": _f, "sSub": _sSub, "sSup": _sSup, "sSubSup": _sSubSup, "sPre": _sPre,
    "rad": _rad, "d": _d, "nary": _nary, "m": _m, "func": _func,
    "limLow": _limLow, "limUpp": _limUpp, "acc": _acc, "bar": _bar,
    "groupChar": _groupChar, "borderBox": _borderBox, "eqArr": _eqArr,
}


def _fallback(el, sz) -> _Box:
    """不认识的结构：把里面所有 m:t 文字排成一行，至少让用户看到内容。"""
    text = "".join(t.text or "" for t in el.iter(_M + "t"))
    if not text:
        return _Space(0.0)
    return _Row(_text_boxes(text, sz, True, False))


def _layout(el, sz: float) -> _Box:
    fn = _LAYOUT.get(_local(el.tag))
    return _fallback(el, sz) if fn is None else fn(el, sz)


# ─────────────────────────────────────────
# 对外入口
# ─────────────────────────────────────────
def omml_to_qimage(el, color: str = "#202020", px: float = 16.0,
                   supersample: int = 2):
    """把一个 m:oMath / m:oMathPara 元素排版成透明底 QImage；失败返回 None。

    color 由调用方按当前主题传入（深色主题给浅色字），px 是公式正文字号。
    内部按 supersample 倍率渲染再平滑缩回，保证在 QTextEdit 里不发虚。
    """
    try:
        box = _layout(el, float(px))
        pad = max(1.0, px * 0.14)
        w = box.w + 2 * pad
        h = box.asc + box.desc + 2 * pad
        if w <= 1.0 or h <= 1.0 or w > 6000 or h > 6000:
            return None
        ss = max(1, int(supersample))
        img = QImage(int(math.ceil(w * ss)), int(math.ceil(h * ss)),
                     QImage.Format.Format_ARGB32_Premultiplied)
        if img.isNull():
            return None
        img.fill(Qt.GlobalColor.transparent)
        p = QPainter(img)
        try:
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            p.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
            p.scale(ss, ss)
            p.setPen(QPen(QColor(color)))
            box.draw(p, pad, pad + box.asc)
        finally:
            p.end()
        if ss != 1:
            img = img.scaled(int(round(w)), int(round(h)),
                             Qt.AspectRatioMode.IgnoreAspectRatio,
                             Qt.TransformationMode.SmoothTransformation)
        return img
    except Exception:
        return None


def omml_text(el) -> str:
    """公式的纯文字近似（渲染失败时的占位文本，也便于日志 / 调试）。"""
    return "".join(t.text or "" for t in el.iter(_M + "t"))

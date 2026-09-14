# -*- coding: utf-8 -*-
"""ui/render — 内容富显示渲染（阶段三）。

media.py：段落内容遍历 + 图片(PNG/JPEG/…)解析 + 缩放。
formula.py：WMF/EMF 矢量预览 → QImage（Windows GDI，零依赖，失败降级）——
            服务 MathType / 公式编辑器这类 OLE 对象自带的预览位图。
omml.py：  OMML（Word「插入→公式」）→ QImage（QPainter 小型数学排版引擎）——
            这类公式**没有**预览位图，只能现场排版。
"""

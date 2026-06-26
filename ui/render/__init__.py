# -*- coding: utf-8 -*-
"""ui/render — 内容富显示渲染（阶段三）。

media.py：段落内容遍历 + 图片(PNG/JPEG/…)解析 + 缩放。
formula.py：WMF/EMF 矢量预览 → QImage（Windows GDI，零依赖，失败降级）。
"""

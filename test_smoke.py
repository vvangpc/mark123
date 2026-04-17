# -*- coding: utf-8 -*-
"""
test_smoke.py — 主窗口冒烟测试
不依赖真实 docx 文件，仅验证主窗口可正常启动、4 个 Tab 均构建成功、
核心 worker 类可从 workers 模块正常导入。用于重构过程中的快速回归。

运行：
    uv run python test_smoke.py
"""
import os
import sys


def test_workers_import():
    from workers import (
        _longest_nonspace_run, _is_pycorrector_available,
        AnnotateWorker, CleanWorker, ToastWidget,
    )
    assert _longest_nonspace_run("hello world") == "hello"
    assert _longest_nonspace_run("") == ""
    assert _longest_nonspace_run("  abc  defg  ") == "defg"
    assert isinstance(_is_pycorrector_available(), bool)
    print("[OK] workers module imports & helpers work")


def test_main_window_construct():
    from PyQt6.QtWidgets import QApplication
    from main_window import MainWindow

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()

    # 4 个 Tab 存在
    assert hasattr(win, "tab_widget"), "MainWindow 应有 tab_widget 属性"
    tab_count = win.tab_widget.count()
    assert tab_count >= 4, f"预期至少 4 个 Tab，实际 {tab_count}"

    # 初始状态
    assert win.doc_data is None
    assert win.current_marks == {}

    # 关键属性
    for attr in ("typo_data", "dup_data", "history_entries"):
        assert hasattr(win, attr), f"MainWindow 缺少属性 {attr}"

    win.close()
    print(f"[OK] MainWindow constructed ({tab_count} tabs)")


def test_load_sample_docx():
    """如果项目根目录存在样本 docx，额外加载一份验证 doc_data 结构。"""
    from PyQt6.QtWidgets import QApplication
    from main_window import MainWindow

    samples = [f for f in os.listdir(".") if f.endswith(".docx") and "初稿" in f]
    if not samples:
        print("[SKIP] no sample docx found in cwd")
        return

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win._load_document(samples[0])
    assert win.doc_data is not None, "加载后 doc_data 应非空"
    assert "paragraphs" in win.doc_data and "sections" in win.doc_data
    print(f"[OK] loaded sample docx: {samples[0]} "
          f"({len(win.doc_data['paragraphs'])} paragraphs, "
          f"{len(win.doc_data['sections'])} sections)")
    win.close()


if __name__ == "__main__":
    test_workers_import()
    test_main_window_construct()
    test_load_sample_docx()
    print("\nAll smoke tests passed.")

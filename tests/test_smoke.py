# -*- coding: utf-8 -*-
"""
test_smoke.py — 主窗口冒烟测试
不依赖真实 docx 文件，仅验证主窗口可正常启动、4 个 Tab 均构建成功、
核心 worker 类可从 workers 模块正常导入。用于重构过程中的快速回归。

运行：
    python tests/test_smoke.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def test_workers_import():
    from ui.workers import (
        _longest_nonspace_run,
        AnnotateWorker, CleanWorker,
    )
    assert _longest_nonspace_run("hello world") == "hello"
    assert _longest_nonspace_run("") == ""
    assert _longest_nonspace_run("  abc  defg  ") == "defg"
    print("[OK] workers module imports & helpers work")


def test_main_window_construct():
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow

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
    from ui.main_window import MainWindow

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


def _all_texts(win):
    return [p.text for p in win.doc_data["paragraphs"]]


def test_spec_only_document():
    """仅说明书（无权利要求书）也能导入、提取标记，并对说明书正文标注，
    且不破坏「附图说明」里的标记定义。"""
    import tempfile
    from docx import Document
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow, AnnotateWorker  # 用主窗口内联 worker（含 scope='spec'）

    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, "spec_only.docx")
    doc = Document()
    for line in (
        "技术领域",
        "本实用新型涉及一种夹持装置。",
        "附图说明",
        "1-齿圈，2-夹指",
        "具体实施方式",
        "如图1所示，所述齿圈固定连接夹指。",
    ):
        doc.add_paragraph(line)
    doc.save(path)

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win._load_document(path)

    # 1) 正常加载、且不含权利要求书
    assert win.doc_data is not None, "仅说明书文档应能成功加载"
    assert "权利要求书" not in win.doc_data["sections"], "样例不应含权利要求书"
    assert win._claim_loaded is False, "无权项时权项检查应保持未加载状态"

    # 2) 从「附图说明」提取到 2 个标记
    assert len(win.current_marks) == 2, f"应提取到 2 个标记，实际 {win.current_marks}"
    assert "齿圈" in win.current_marks.values() and "夹指" in win.current_marks.values()

    # 3) scope='spec' 标注：具体实施方式被标注，附图说明定义不被破坏
    worker = AnnotateWorker(win.doc_data, win.current_marks, action="add", scope="spec")
    worker.run()  # 同步执行（不开线程）
    texts = _all_texts(win)
    impl_text = next(t for t in texts if "固定连接" in t)
    assert "齿圈1" in impl_text and "夹指2" in impl_text, f"说明书正文应被标注：{impl_text!r}"
    def_text = next(t for t in texts if "1-齿圈" in t)
    assert def_text.strip() == "1-齿圈，2-夹指", f"附图说明标记定义不应被标注破坏：{def_text!r}"

    win.close()
    try:
        os.remove(path)
        os.rmdir(tmpdir)
    except OSError:
        pass
    print("[OK] spec-only document loads, extracts marks, annotates spec body safely")


if __name__ == "__main__":
    test_workers_import()
    test_main_window_construct()
    test_load_sample_docx()
    test_spec_only_document()
    print("\nAll smoke tests passed.")

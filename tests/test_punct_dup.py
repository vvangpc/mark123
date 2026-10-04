# -*- coding: utf-8 -*-
"""
test_punct_dup.py — 重复标点检查 + 共用结果表回归测试

覆盖：
  · core.cleaner.check_duplicate_punct 的检出范围与边界（省略号不误报）
  · occurrence 为 1 基（同段同词第 n 处能被正确定位 / 高亮）
  · 3列 新模块（🔣 标点 / 🔍 孤立标）与 4列 按钮就位
  · 「重复标点检查 → 应用」端到端写回内存
  · 单条「修改」会把同段同词的兄弟行一并清掉（整段替换语义）

运行：
    python tests/test_punct_dup.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tests.isolation  # noqa: F401  测试不写用户真实的设置 / 配置目录


class _Shell:
    __slots__ = ("text",)

    def __init__(self, t):
        self.text = t


def test_check_duplicate_punct_scope():
    from core.cleaner import check_duplicate_punct

    paras = [_Shell("本发明的装置。。其中，，设有孔，、并且如图1。.所示...结束。")]
    got = {(r["wrong"], r["suggestion"]) for r in check_duplicate_punct(paras)}
    assert ("。。", "。") in got, got          # 同字符连写
    assert ("，，", "，") in got, got          # 同字符连写
    assert ("，、", "，") in got, got          # 逗号 / 顿号混写
    assert ("。.", "。") in got, got           # 全角 / 半角混写 → 取全角
    assert not any(w == "..." for w, _ in got), "英文省略号不应报重复标点"

    # 不同停顿的连写（，。）含义不明确，不在本检查范围
    assert not check_duplicate_punct([_Shell("如上，。下同")])
    # 干净文本不应有任何检出
    assert not check_duplicate_punct([_Shell("一种装置，其特征在于，包括壳体。")])
    print("[OK] 重复标点检出范围正确（同字符 / 全半角混写 / 顿逗混写；省略号不误报）")


def test_occurrence_is_one_based():
    """同段同词的第 n 处必须能被 1 基定位取到（修好的 B1）。"""
    from core.cleaner import check_typos_wordbank, check_duplicate_punct
    from ui.content_area import ContentArea

    text = "见权力要求1，也见权力要求2，还见权力要求3。"
    res = [r for r in check_typos_wordbank([_Shell(text)]) if r["wrong"] == "权力要求"]
    assert len(res) == 3, f"应报 3 处，实际 {len(res)}"
    occs = [r["occurrence"] for r in res]
    assert occs == [1, 2, 3], f"occurrence 应为 1 基递增，实际 {occs}"
    # 每条都能反查到互不相同的偏移 —— 0 基时第 3 处会取不到、第 1 处被取两次
    offsets = [ContentArea._nth_occurrence(text, "权力要求", o) for o in occs]
    assert all(o >= 0 for o in offsets) and len(set(offsets)) == 3, offsets

    p = check_duplicate_punct([_Shell("甲。。乙。。丙")])
    assert [r["occurrence"] for r in p] == [1, 2], p
    print("[OK] occurrence 为 1 基，同段多处可分别定位")


def _wait_worker(win, app, timeout_ms=10000):
    """等后台清洗线程跑完并把 finished 信号派发掉。

    不要用 `while not win._is_busy()` 轮询 —— QThread.start() 之后线程可能
    还没真正跑起来，isRunning() 先返回 False，轮询会立刻退出，断言就跑在
    改动落盘之前（间歇性假失败）。
    """
    w = getattr(win, "clean_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    app.processEvents()


def _make_docx(path, lines):
    from docx import Document
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)


def test_nav_modules_and_buttons():
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()

    labels = [win.nav_panel.col1.item(i).text() for i in range(win.nav_panel.col1.count())]
    assert len(labels) == 8, f"3列 应有 8 个模块，实际 {labels}"
    assert any("标点" in t for t in labels), labels
    assert any("孤立标" in t for t in labels), labels

    # 标点 / 孤立标 的操作按钮已经在 4列
    for attr in ("punct_btn", "punct_dup_btn", "punct_dup_apply_btn",
                 "orphan_btn", "orphan_clear_btn"):
        assert hasattr(win, attr), f"4列 缺少 {attr}"
    assert win.punct_btn.parent() is not None
    # 三类检查共用同一张结果表
    assert set(win.CHECK_KINDS) == {"typo", "dup", "punct_dup"}
    win._unsaved = False  # 测试不弹「未保存」确认
    win.close()
    print("[OK] 3列 8 模块 + 标点 / 孤立标 的 4列 按钮就位")


def test_punct_dup_end_to_end():
    """检查 → 结果进共用表 → 应用所有修改 → 内存真的被改。"""
    import tempfile
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow
    from ui.workers import CleanWorker

    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, "punct.docx")
    _make_docx(path, [
        "权利要求书",
        "1.一种夹持装置，其特征在于，包括齿圈。。",
        "说明书",
        "技术领域",
        "本实用新型涉及一种夹持装置，，属于机械领域。",
    ])

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win._load_document(path)
    assert win.doc_data is not None

    # 同步跑 worker（避免测试里起线程）
    got = []
    w = CleanWorker(win.doc_data, "punct_dup_check")
    w.typo_results.connect(got.append)
    w.run()
    results = got[0]
    wrongs = sorted(r["wrong"] for r in results)
    assert wrongs == ["。。", "，，"], f"应检出 2 处重复标点，实际 {wrongs}"
    assert all(r["kind"] == "punct_dup" for r in results)
    assert all(r["section"] for r in results), "每条都应带章节定位"

    # 进共用结果表 + 1框 内联高亮
    win._current_check_kind = "punct_dup"
    win.punct_dup_data = results
    win._render_table_from_data(results)
    assert win.typo_table.rowCount() == 2
    assert "重复标点" in win.typo_result_group.title()
    assert win.punct_dup_apply_btn.isEnabled()

    # 应用所有修改（同步跑）
    corrections = [
        {"para_idx": r["para_idx"], "wrong": r["wrong"], "confirmed_fix": r["suggestion"]}
        for r in results
    ]
    CleanWorker(win.doc_data, "typo_apply", corrections=corrections).run()
    texts = [p.text for p in win.doc_data["paragraphs"]]
    assert not any("。。" in t or "，，" in t for t in texts), texts
    assert any(t.endswith("包括齿圈。") for t in texts), texts

    win._unsaved = False  # 测试不弹「未保存」确认
    win.close()
    try:
        os.remove(path)
        os.rmdir(tmpdir)
    except OSError:
        pass
    print("[OK] 重复标点端到端：检查 → 共用结果表 → 应用写回内存")


def test_single_fix_only_touches_its_occurrence():
    """单条「修改」只改这一处：同段同词的另一行保留、occurrence 按改后文本重算，
    随后再点它也能改对位置（旧实现整段全改，忽略第 1 处、只改第 2 处会两处都改）。"""
    import tempfile
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow
    from ui.workers import CleanWorker

    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, "sib.docx")
    _make_docx(path, [
        "说明书",
        "技术领域",
        "本实用新型涉及夹具。。其结构简单。。成本低。",
    ])

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win._load_document(path)

    got = []
    w = CleanWorker(win.doc_data, "punct_dup_check")
    w.typo_results.connect(got.append)
    w.run()
    results = [r for r in got[0] if r["wrong"] == "。。"]
    assert [r["occurrence"] for r in results] == [1, 2], results

    win._current_check_kind = "punct_dup"
    win.punct_dup_data = results
    win._render_table_from_data(results)
    # 只改第 2 处
    win._apply_single_correction(1)
    text = next(p.text for p in win.doc_data["paragraphs"] if "夹具" in p.text)
    assert text == "本实用新型涉及夹具。。其结构简单。成本低。", text
    assert len(win.punct_dup_data) == 1 and win.punct_dup_data[0]["occurrence"] == 1
    assert win.history_entries and "1 处" in win.history_entries[-1]["summary"], win.history_entries
    # 剩下那条仍能准确命中第 1 处
    win._apply_single_correction(0)
    text = next(p.text for p in win.doc_data["paragraphs"] if "夹具" in p.text)
    assert text == "本实用新型涉及夹具。其结构简单。成本低。", text
    assert win.punct_dup_data == [] and win.typo_table.rowCount() == 0

    win._unsaved = False  # 测试不弹「未保存」确认
    win.close()
    try:
        os.remove(path)
        os.rmdir(tmpdir)
    except OSError:
        pass
    print("[OK] 单条修改只改对应的那一处，同段其余行重算后仍可定位")


def test_execute_buttons_switch_back_to_their_page():
    """4列 的两个执行按钮：先把 2框 切回自己的参数页，再执行。

    「🗑️ 执行删除」和「▶ 执行标点检查」都是"参数在 2框、按钮在 4列"，
    而 2框 可能正停在别的页（全文替换 / 重复标点结果表）。不切页就执行，
    用户看不见自己改的是哪几章 / 按什么规则改。
    """
    import tempfile
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow

    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, "exec.docx")
    _make_docx(path, [
        "权利要求书",
        "1.一种夹具，其特征在于，包括所述齿圈。。",
        "说明书",
        "技术领域",
        "本发明涉及夹具。",
        "具体实施方式",
        "所述夹具，，包括齿圈。。",
    ])

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win._load_document(path)

    # 被删掉的控件确实不在了
    assert not hasattr(win, "fix_punctuation_cb"), "「修正连续重复标点」勾选项应已删除"
    assert win.suoshu_btn.text() == "🗑️ 执行删除", win.suoshu_btn.text()
    assert win.confirm_marks_btn.text() == "✅ 修改标记字典", win.confirm_marks_btn.text()
    assert win.clear_marks_btn.text() == "🗑️ 清空标记字典", win.clear_marks_btn.text()
    # 「修改 / 清空标记字典」贴着编辑框放在 2框，与编辑框同一个父面板
    assert win.confirm_marks_btn.parent() is win.clear_marks_btn.parent()

    # 执行删除：从「全文替换」页点下去
    win.panel_stack.setCurrentIndex(6)
    win.suoshu_checkboxes["具体实施方式"].setChecked(True)
    win._on_clean_suoshu()
    app.processEvents()
    assert win.panel_stack.currentIndex() == 1, "执行删除应把 2框 切回章节勾选页"
    _wait_worker(win, app)
    assert not any("所述夹具" in p.text for p in win.doc_data["paragraphs"])

    # 执行标点检查：从「重复标点结果表」页点下去
    win.panel_stack.setCurrentIndex(4)
    win._on_clean_punct()
    app.processEvents()
    assert win.panel_stack.currentIndex() == 2, "执行标点检查应把 2框 切回勾选项页"
    _wait_worker(win, app)
    # 连续标点交给「重复标点检查」逐条确认，标点统一不再顺手改掉
    assert any("。。" in p.text for p in win.doc_data["paragraphs"]),         "「标点统一」不应再自动修正连续重复标点"

    win._unsaved = False  # 测试不弹「未保存」确认
    win.close()
    try:
        os.remove(path)
        os.rmdir(tmpdir)
    except OSError:
        pass
    print("[OK] 执行删除 / 执行标点检查：先切回参数页再执行；连续标点归重复标点检查")


def test_replace_page_preview():
    """全文替换的实时预览（留空＝删除）。"""
    from PyQt6.QtWidgets import QApplication
    from ui.main_window import MainWindow

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    assert "填写" in win.replace_preview_label.text()
    win.replace_from_edit.setText("发明")
    win.replace_to_edit.setText("实用新型")
    assert "「发明」" in win.replace_preview_label.text()
    assert "「实用新型」" in win.replace_preview_label.text()
    win.replace_to_edit.setText("")
    assert "删除" in win.replace_preview_label.text(), win.replace_preview_label.text()
    win._unsaved = False  # 测试不弹「未保存」确认
    win.close()
    print("[OK] 全文替换实时预览（含留空＝删除）")


if __name__ == "__main__":
    test_check_duplicate_punct_scope()
    test_occurrence_is_one_based()
    test_nav_modules_and_buttons()
    test_punct_dup_end_to_end()
    test_single_fix_only_touches_its_occurrence()
    test_execute_buttons_switch_back_to_their_page()
    test_replace_page_preview()
    print("\nAll punct-dup tests passed.")

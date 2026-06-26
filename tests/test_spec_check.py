# -*- coding: utf-8 -*-
"""
test_spec_check.py — 说明书检查回归测试（实施例编号 + 摘要字数 + 中文数字）

运行：
    python tests/test_spec_check.py
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.doc_parser import DocSection
from core.spec_check import (
    _cn_to_int, check_embodiment_numbering, check_abstract_length,
)


class _Shell:
    """合成段落：只有 .text，足够喂给纯函数检查。"""
    __slots__ = ("text",)

    def __init__(self, t):
        self.text = t


def _impl(lines):
    """把若干行包成 具体实施方式 段落 + sections（首行当章节标题）。"""
    shells = [_Shell(t) for t in lines]
    sections = {"具体实施方式": DocSection("具体实施方式", 0, len(shells))}
    return check_embodiment_numbering(shells, sections)


def _kinds(results):
    return [r["kind"] for r in results]


# ── 中文数字 ──────────────────────────────────────────────
def test_cn_to_int():
    cases = {"一": 1, "二": 2, "两": 2, "九": 9, "十": 10, "十一": 11,
             "十九": 19, "二十": 20, "二十一": 21, "九十九": 99,
             "1": 1, "12": 12, "１２": 12, "〇": 0}
    for s, want in cases.items():
        got = _cn_to_int(s)
        assert got == want, f"_cn_to_int({s!r})={got}, want {want}"
    for bad in ("", "甲", "一二", "实施"):
        assert _cn_to_int(bad) is None, f"_cn_to_int({bad!r}) should be None"
    print("[OK] _cn_to_int")


# ── 实施例编号 ────────────────────────────────────────────
def test_clean_sequence():
    r = _impl(["具体实施方式", "实施例一：一种装置", "正文……",
               "实施例二", "正文……", "实施例三"])
    assert r == [], f"clean seq should pass, got {_kinds(r)}"
    print("[OK] 干净序列无报错")


def test_gap():
    r = _impl(["具体实施方式", "实施例一", "实施例三"])
    assert _kinds(r) == ["emb_gap"], _kinds(r)
    assert "实施例二" in r[0]["message"]
    print("[OK] 缺号")


def test_duplicate():
    r = _impl(["具体实施方式", "实施例一", "实施例二", "实施例二"])
    assert _kinds(r) == ["emb_dup"], _kinds(r)
    # para_idx 指向第二个「实施例二」(index 3)
    assert r[0]["para_idx"] == 3, r[0]["para_idx"]
    print("[OK] 重号（定位到第二个）")


def test_out_of_order():
    r = _impl(["具体实施方式", "实施例一", "实施例三", "实施例二"])
    ks = _kinds(r)
    assert "emb_order" in ks, ks
    print("[OK] 顺序颠倒")


def test_wrong_start_no_extra_gap():
    r = _impl(["具体实施方式", "实施例二", "实施例三"])
    assert _kinds(r) == ["emb_start"], _kinds(r)   # 不得多报「缺一」gap
    print("[OK] 起始非一（无多余缺号）")


def test_single():
    assert _impl(["具体实施方式", "实施例一"]) == []
    assert _kinds(_impl(["具体实施方式", "实施例二"])) == ["emb_start"]
    print("[OK] 单个实施例")


def test_inline_not_counted():
    r = _impl(["具体实施方式", "如实施例一所述的结构……",
               "实施例二中的齿轮与……", "参见实施例三相同的方法"])
    assert r == [], f"行内引用不应计入，got {_kinds(r)}"
    print("[OK] 行内引用不计入")


def test_arabic_and_di_form():
    assert _impl(["具体实施方式", "实施例1", "实施例2", "实施例3"]) == []
    assert _impl(["具体实施方式", "第一实施例", "第二实施例"]) == []
    assert _impl(["具体实施方式", "实施例一", "实施例2"]) == []   # 混用只看序号
    print("[OK] 阿拉伯 / 第N实施例 / 混用")


def test_section_absent():
    shells = [_Shell("随便")]
    assert check_embodiment_numbering(shells, {}) == []
    assert check_embodiment_numbering(shells, {"权利要求书": DocSection("权利要求书", 0, 1)}) == []
    print("[OK] 无具体实施方式章节")


def test_dup_then_order_combo():
    # [2,2,1] → emb_start(2) + emb_dup(2) + emb_order(1)
    r = _impl(["具体实施方式", "实施例二", "实施例二", "实施例一"])
    assert _kinds(r) == ["emb_start", "emb_dup", "emb_order"], _kinds(r)
    print("[OK] 组合 [2,2,1]")


# ── 摘要字数 ──────────────────────────────────────────────
def _abstract(lines, limit=300):
    shells = [_Shell(t) for t in lines]
    sections = {"说明书摘要": DocSection("说明书摘要", 0, len(shells))}
    return check_abstract_length(shells, sections, limit=limit)


def test_abstract_ok():
    assert _abstract(["说明书摘要", "本发明公开了一种装置。"]) == []
    print("[OK] 摘要未超限")


def test_abstract_over():
    body = "甲" * 250
    r = _abstract(["说明书摘要", body, "乙" * 80])   # 250+80=330 > 300
    assert len(r) == 1 and r[0]["kind"] == "abstract_len", r
    assert "330" in r[0]["message"], r[0]["message"]
    # 越界起点：第301字落在第二段（乙串）第51个字 → wrong 为该段越界尾巴
    assert r[0]["para_idx"] == 2, r[0]["para_idx"]
    assert r[0]["wrong"] and r[0]["wrong"].startswith("乙"), r[0]["wrong"]
    print("[OK] 摘要超限（定位越界段）")


def test_abstract_absent():
    assert check_abstract_length([_Shell("x")], {}) == []
    print("[OK] 无摘要章节")


if __name__ == "__main__":
    test_cn_to_int()
    test_clean_sequence()
    test_gap()
    test_duplicate()
    test_out_of_order()
    test_wrong_start_no_extra_gap()
    test_single()
    test_inline_not_counted()
    test_arabic_and_di_form()
    test_section_absent()
    test_dup_then_order_combo()
    test_abstract_ok()
    test_abstract_over()
    test_abstract_absent()
    print("\nAll spec_check tests passed.")

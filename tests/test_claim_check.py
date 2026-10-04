# -*- coding: utf-8 -*-
"""
test_claim_check.py — 权利要求引用基础检查回归测试

聚焦修复过的"明明有引用基础却报缺失"误报：
当"所述<较短术语><动词/系词><真实术语>"出现时，固定 N 字滑窗的
"所述"掩码会越过较短术语，把紧随其后真实术语的【首字】也盖掉，
导致该真实术语无法登记为引用基础，于是它后续的"所述X"全部被误报。

典型现场（本仓库样例 SE26Y1385）：
    所述装置包括流量调节阀…      → "装置"(2字) 之后的 "流量调节阀" 首字被盖
    所述流量计为质量流量计…      → "流量计"(3字) 之后的 "质量流量计" 首字被盖

运行：
    python tests/test_claim_check.py
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.claim_check import parse_claims_ex, check_antecedent_basis


class _Shell:
    __slots__ = ("text",)

    def __init__(self, t):
        self.text = t


def _antecedent(lines, n, **kw):
    """把若干行权利要求文本喂进检查器，返回引用基础问题的术语集合。"""
    shells = [_Shell(t) for t in lines]
    claims, _ = parse_claims_ex(shells, 0, len(shells))
    res = check_antecedent_basis(claims, n, set(), **kw)
    return {r["message"] for r in res}


def test_term_after_short_suoshu_has_basis():
    """『所述装置包括流量调节阀』为流量调节阀提供引用基础，不应误报。"""
    lines = [
        "1.一种调节装置，其特征在于，所述装置包括流量调节阀、流量计；"
        "所述流量调节阀的入口与气源连通，所述流量调节阀的出口与所述流量计连通。",
    ]
    # 关键：N=5 时旧实现会把"装置"(2字)掩码延伸到"流量调节阀"的"流"，
    # 使其无法登记为引用基础 → 误报。
    for n in (2, 3, 4, 5, 6):
        msgs = _antecedent(lines, n)
        bad = [m for m in msgs if "流量调节阀" in m]
        assert not bad, f"n={n} 误报流量调节阀缺少引用基础：{bad}"
    print("[OK] 短'所述'后的真实术语正确获得引用基础（流量调节阀）")


def test_term_after_copula_has_basis():
    """『所述流量计为质量流量计』为质量流量计提供引用基础，不应误报。"""
    lines = [
        "1.一种装置，其特征在于，包括流量计；所述流量计为质量流量计，"
        "所述质量流量计为热式质量流量计。",
    ]
    for n in (5,):
        msgs = _antecedent(lines, n)
        bad = [m for m in msgs if "质量流量计" in m]
        assert not bad, f"n={n} 误报质量流量计缺少引用基础：{bad}"
    print("[OK] 系词'为'后的真实术语正确获得引用基础（质量流量计）")


def test_genuinely_missing_basis_still_flagged():
    """真正缺引用基础的术语仍应被检出（确保修复没有把检查放空）。"""
    # "齿圈"从未以"非所述"形式定义，"所述齿圈"应被标记。
    lines = [
        "1.一种装置，其特征在于，包括底座；所述齿圈套设于所述底座。",
    ]
    msgs = _antecedent(lines, 2)
    assert any("齿圈" in m for m in msgs), f"应检出'齿圈'缺少引用基础，实际：{msgs}"
    print("[OK] 真正缺引用基础的术语仍被检出（齿圈）")


def test_modes_consistent_on_basis():
    """截断 / 回退 各模式下，有基础的术语都不应被误报。"""
    lines = [
        "1.一种调节装置，其特征在于，所述装置包括流量调节阀；"
        "所述流量调节阀的入口与气源连通。",
    ]
    from config.config_manager import get_builtin_boundary_blacklist
    bl = get_builtin_boundary_blacklist()
    combos = [
        dict(use_dynamic_truncate=False, use_dynamic_fallback=False, boundary_blacklist=None),
        dict(use_dynamic_truncate=False, use_dynamic_fallback=True, boundary_blacklist=None),
        dict(use_dynamic_truncate=True, use_dynamic_fallback=False, boundary_blacklist=bl),
        dict(use_dynamic_truncate=True, use_dynamic_fallback=True, boundary_blacklist=bl),
    ]
    for kw in combos:
        msgs = _antecedent(lines, 5, **kw)
        bad = [m for m in msgs if "流量调节阀" in m]
        assert not bad, f"模式 {kw} 误报：{bad}"
    print("[OK] 全部模式下流量调节阀均有引用基础")


def test_dynamic_truncate_boundaries():
    """动态截断的边界判定：黑名单词打头 / 单字虚词 / 序数前缀 / 尾部方位词。"""
    from core.claim_check import (
        _extract_term_dynamic_truncate as _trunc,
        _trim_stopchars, _build_blacklist_lookup, DEFAULT_BOUNDARY_BLACKLIST,
    )
    f, w, l = _build_blacklist_lookup(DEFAULT_BOUNDARY_BLACKLIST)

    def term(t):
        return _trim_stopchars(_trunc(t, 0, f, w, l))

    cases = {
        # 黑名单词打头的常见部件名：旧实现截出空串 → 整条漏检
        "安装座上设有孔": "安装座",
        "连接杆的一端": "连接杆",
        "固定板": "固定板",
        "设置槽": "设置槽",
        # 序数前缀「第X」不算进 min_keep，否则截成"第一"
        "第一连接件的端部": "第一连接件",
        "第二安装座": "第二安装座",
        # 单字虚词是最可靠的右边界（黑名单里只有双字词）
        "流量调节阀的入口与气源连通": "流量调节阀",
        "上盖与下盖": "上盖",
        # 尾部单字方位词剪掉，但量词"件"不能剪
        "壳体上的孔": "壳体",
        "弹性件": "弹性件",
        # 黑名单词仍正常生效
        "齿轮安装在主轴上": "齿轮",
        "装置包括流量调节阀": "装置",
    }
    for text, expect in cases.items():
        got = term(text)
        assert got == expect, f"截断 {text!r}: 期望 {expect!r}，实际 {got!r}"
    print("[OK] 动态截断边界：黑名单打头 / 单字虚词 / 序数前缀 / 尾部方位词")


def test_short_term_channel():
    """"所述"后不足 n 字时仍参与检查，且已定义的短术语不误报。"""
    # n=6，而"所述壳体，"后面只有 2 个汉字就遇到标点 ——
    # 旧实现要求恰好 n 字，取不满就 continue，该处**完全不检查**。
    有基础 = ["1.一种装置，其特征在于，包括壳体；所述壳体，其上设有孔。"]
    assert not [m for m in _antecedent(有基础, 6) if "壳体" in m], "已定义的短术语不应误报"

    无基础 = ["1.一种装置，其特征在于，包括底板；所述壳体，其上设有孔。"]
    msgs = _antecedent(无基础, 6)
    assert any("壳体" in m for m in msgs), f"n=6 时短术语'壳体'应被检出，实际：{msgs}"
    print("[OK] 短术语通道：不足 n 字的'所述X'仍被检查，且不凭空误报")


def test_truncate_overrun_falls_back():
    """截断溢出（黑名单没盖住的动词）不应报出长串误判。"""
    from config.config_manager import get_builtin_boundary_blacklist
    lines = [
        "1.一种方法，其特征在于，包括受力分析单元；"
        "将所述受力分析单元简化为梁单元力学模型。",
    ]
    msgs = _antecedent(
        lines, 4,
        use_dynamic_truncate=True,
        boundary_blacklist=get_builtin_boundary_blacklist(),
    )
    bad = [m for m in msgs if "受力分析单元" in m]
    assert not bad, f"截断溢出应回落到 n 字术语而非报长串：{bad}"
    print("[OK] 截断溢出回落，不再报出整句长串")


_ALL_MODES = [
    dict(use_dynamic_truncate=False, use_dynamic_fallback=False),
    dict(use_dynamic_truncate=False, use_dynamic_fallback=True),
    dict(use_dynamic_truncate=True, use_dynamic_fallback=False),
    dict(use_dynamic_truncate=True, use_dynamic_fallback=True),
]


def _each_mode(lines, n=2):
    from config.config_manager import get_builtin_boundary_blacklist
    bl = get_builtin_boundary_blacklist()
    for kw in _ALL_MODES:
        yield kw, _antecedent(lines, n, boundary_blacklist=bl, **kw)


def test_alternative_dependency_needs_basis_in_every_branch():
    """择一引用「权利要求1或2」：只在权2 引入的术语，与权1 组合时没有引用基础。"""
    lines = [
        "1.一种装置，其特征在于，包括壳体。",
        "2.根据权利要求1所述的装置，其特征在于，还包括弹簧。",
        "3.根据权利要求1或2所述的装置，其特征在于，所述弹簧套设于所述壳体。",
    ]
    for kw, msgs in _each_mode(lines):
        hit = [m for m in msgs if "弹簧" in m]
        assert hit and "权利要求 1" in hit[0], f"{kw} 应指出引用权1 时缺少基础：{msgs}"
        assert not [m for m in msgs if "壳体" in m], f"{kw} 壳体在两条分支都有定义：{msgs}"
    print("[OK] 择一引用：术语须在每条被引分支都有定义")


def test_suoshu_de_form():
    """「所述的X」与「所述X」同等对待：有基础不报，无基础要报。"""
    lines = ["1.一种装置，其特征在于，包括壳体和底座；"
             "所述的壳体固定于所述的底座，所述的齿圈套设于所述的壳体。"]
    for kw, msgs in _each_mode(lines):
        assert msgs == {"『所述的齿圈』缺少引用基础"}, f"{kw}：{msgs}"
    print("[OK] 所述的X：有基础不误报，无基础照报")


def test_fixed_mask_does_not_swallow_following_term():
    """「所述多源合束器（5）为二向色镜」里首次出现的二向色镜应登记为定义。"""
    lines = ["1.一种系统，其特征在于，包括多源合束器（5）；"
             "所述多源合束器（5）为二向色镜，所述二向色镜设定为透射光束。"]
    for n in (2, 4):
        for kw, msgs in _each_mode(lines, n):
            assert not [m for m in msgs if "二向" in m], f"n={n} {kw} 误报：{msgs}"
    print("[OK] 被引术语之后紧跟的首次定义不再被掩码吞掉")


def test_inherited_short_term():
    """n 大于术语长度时，父权项里定义的短术语照样能被从属权项继承。"""
    lines = [
        "1.一种装置，其特征在于，包括壳体。",
        "2.根据权利要求1所述的装置，其特征在于，所述壳体。",
    ]
    for n in (3, 4, 6):
        assert not _antecedent(lines, n), f"n={n} 误报继承来的短术语"
    print("[OK] 短术语可跨权项继承")


def test_ordinal_term():
    """「所述第三弹簧」：只定义了第一 / 第二弹簧时要报，不能被「第三光路」的「第三」放过。"""
    lines = ["1.一种装置，其特征在于，包括第一弹簧、第二弹簧和第三光路；"
             "所述第一弹簧与所述第三弹簧连接。"]
    for kw, msgs in _each_mode(lines):
        assert any("第三弹簧" in m for m in msgs), f"{kw} 漏报第三弹簧：{msgs}"
        assert not any("第一弹簧" in m for m in msgs), f"{kw} 误报第一弹簧：{msgs}"
    print("[OK] 序数术语参与检查且回退不缩进序数")


def test_positional_head_term_checked_in_default_mode():
    """默认模式下方位字打头的部件名（上盖）也参与检查。"""
    msgs = _antecedent(["1.一种装置，其特征在于，包括壳体；所述上盖盖合于所述壳体。"], 2)
    assert msgs == {"『所述上盖』缺少引用基础"}, msgs
    print("[OK] 方位字打头的术语参与默认模式检查")


def test_quantifier_prefix():
    """「包括若干连接杆；所述多个连接杆」剥掉数量前缀后有引用基础。"""
    lines = ["1.一种装置，其特征在于，包括若干连接杆；所述多个连接杆均匀分布。"]
    for kw, msgs in _each_mode(lines):
        assert not msgs, f"{kw} 误报：{msgs}"
    print("[OK] 数量前缀剥离后匹配")


def test_missing_term_reported_once():
    """同一缺失术语：本权项多处 + 从属权项，只在首处报一条。"""
    lines = [
        "1.一种装置，其特征在于，包括壳体；所述齿圈套设于所述壳体，"
        "所述齿圈与所述壳体之间设有间隙，所述齿圈为钢制。",
        "2.根据权利要求1所述的装置，其特征在于，所述齿圈为钢制。",
    ]
    from config.config_manager import get_builtin_boundary_blacklist
    bl = get_builtin_boundary_blacklist()
    for n in (2, 3):
        for kw in _ALL_MODES:
            shells = [_Shell(t) for t in lines]
            claims, _ = parse_claims_ex(shells, 0, len(shells))
            res = check_antecedent_basis(claims, n, set(), boundary_blacklist=bl, **kw)
            assert len(res) == 1 and res[0]["claim_no"] == 1, f"n={n} {kw}：{[r['message'] for r in res]}"
    print("[OK] 缺失术语只报首处")


def test_result_points_at_actual_paragraph():
    """结果的 para_idx 指向「所述X」所在段，anchor 是原文里能搜到的字面串。"""
    shells = [_Shell(t) for t in [
        "1.一种装置，其特征在于，包括：",
        "壳体；",
        "所述的齿圈套设于所述壳体。",
    ]]
    claims, _ = parse_claims_ex(shells, 0, len(shells))
    res = check_antecedent_basis(claims, 2, set())
    assert len(res) == 1, res
    assert res[0]["para_idx"] == 2 and res[0]["anchor"] == "所述的齿圈" and res[0]["term"] == "齿圈", res
    assert check_antecedent_basis(claims, 2, {"齿圈"}) == [], "忽略词应生效"
    print("[OK] 结果定位到实际段落，anchor / term 字段正确")


def test_citation_formula_variants():
    """「之一」「中的任一项」引用语能被识别为从属关系；长编号串里的「所述」不当术语引用。"""
    from core.claim_check import run_all_checks
    lines = [
        "1.一种装置，其特征在于，包括壳体。",
        "2.根据权利要求1所述的装置，其特征在于，还包括盖板。",
        "3.根据权利要求1或2之一所述的装置，其特征在于，所述壳体为金属。",
        "4.根据权利要求1至3中的任一项所述的装置，其特征在于，所述壳体为金属。",
        "5.根据权利要求1、2、3或4中任一项所述的装置，其特征在于，所述壳体为金属。",
    ]
    shells = [_Shell(t) for t in lines]
    claims, _ = parse_claims_ex(shells, 0, len(shells))
    assert claims[3].cites == {1, 2} and claims[4].cites == {1, 2, 3}, {k: v.cites for k, v in claims.items()}
    res = run_all_checks(shells, 0, len(shells), n=2, use_dynamic_truncate=True,
                         use_dynamic_fallback=True, boundary_blacklist=[])
    assert not [r for r in res if r["kind"] == "antecedent"], res
    assert not [r for r in res if "任一项" in r["message"]], "「之一」已写明择一，不应再提示补「任一项」"
    print("[OK] 引用语变体（之一 / 中的任一项 / 长编号串）")


def test_vague_overlapping_words_report_once():
    """重叠的不确定用语（优选/优选地、基本/基本上、大约/约为）只报一条。"""
    from core.claim_check import check_vague_terms, ClaimInfo
    claims = {1: ClaimInfo(
        no=1, para_indices=[0],
        text="一种装置，优选地设有基本上水平的板，大约为10mm。", raw_text="",
    )}
    msgs = [r["message"] for r in check_vague_terms(claims)]
    assert len(msgs) == 3, f"3 处问题应报 3 条，实际 {len(msgs)} 条：{msgs}"
    assert any("优选地" in m for m in msgs) and not any(m.endswith("『优选』") for m in msgs)
    assert any("基本上" in m for m in msgs) and not any(m.endswith("『基本』") for m in msgs)
    print("[OK] 不确定用语重叠词只报最长的一条")


if __name__ == "__main__":
    test_term_after_short_suoshu_has_basis()
    test_term_after_copula_has_basis()
    test_genuinely_missing_basis_still_flagged()
    test_modes_consistent_on_basis()
    test_dynamic_truncate_boundaries()
    test_short_term_channel()
    test_truncate_overrun_falls_back()
    test_alternative_dependency_needs_basis_in_every_branch()
    test_suoshu_de_form()
    test_fixed_mask_does_not_swallow_following_term()
    test_inherited_short_term()
    test_ordinal_term()
    test_positional_head_term_checked_in_default_mode()
    test_quantifier_prefix()
    test_missing_term_reported_once()
    test_result_points_at_actual_paragraph()
    test_citation_formula_variants()
    test_vague_overlapping_words_report_once()
    print("\nAll claim_check antecedent tests passed.")

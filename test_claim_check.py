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
    python test_claim_check.py
"""
from claim_check import parse_claims_ex, check_antecedent_basis


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
    from config_manager import get_builtin_boundary_blacklist
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


if __name__ == "__main__":
    test_term_after_short_suoshu_has_basis()
    test_term_after_copula_has_basis()
    test_genuinely_missing_basis_still_flagged()
    test_modes_consistent_on_basis()
    print("\nAll claim_check antecedent tests passed.")

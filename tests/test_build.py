#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_eezo_template の採算計算・推定禁止挙動を検証する.

実行: python -m pytest tests/ -q
（pytest未導入でも `python tests/test_build.py` で簡易実行可能）
"""

from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import build_eezo_template as b  # noqa: E402

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "..", "config", "eezo_policy.json")


def _cfg() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def _umeya_a(extra: dict | None = None) -> dict:
    rec = {
        "仕入先": "梅屋", "出典ファイル": "0416", "商品名_見積準拠": "シューAセット",
        "温度帯": "冷凍", "仕入単価_値": "1912", "仕入単価_税区分": "税抜",
        "参考上代_値": "2549", "参考上代_税区分": "税抜",
        "仕入送料区分": "産直別途", "税率_明記": "8", "想定販売価格_税込": "2160",
    }
    if extra:
        rec.update(extra)
    return rec


_UMEYA_SHIP = {
    "梅屋": {
        "体系": "温度帯型",
        "rows": [{
            "仕入先": "梅屋", "送料体系": "温度帯型", "キー種別": "温度帯",
            "キー値": "冷蔵・冷凍", "地域": "関西", "送料_値": "1410",
            "送料_税区分": "税込", "クール付加": "",
        }],
    }
}


def test_売上原価_は仕入単価と仕入送料の和() -> None:
    """添付の検証値（売上原価3322）と一致すること."""
    out, _ = b.build_row(_umeya_a(), _UMEYA_SHIP, _cfg())
    assert out["仕入単価_税抜"] == 1912
    assert out["仕入送料_基準地域_税抜"] == 1410
    assert out["売上原価_税抜"] == 3322


def test_粗利は売上引く原価_決済手数料は粗利外() -> None:
    """売上＝販売税抜＋消費者送料税抜 / 原価＝仕入＋仕入送料 で粗利が出ること.

    決済手数料は販管費のため粗利には含めない（参考列にのみ計上）。
    """
    out, _ = b.build_row(_umeya_a(), _UMEYA_SHIP, _cfg())
    # 販売税抜=2000, 消費者送料(クール関西1200・税抜)=1090.909091
    assert abs(out["売上_税抜"] - 3090.909091) < 1e-6
    # 粗利 = 3090.909091 - 3322 = -231.090909（決済手数料は引かない）
    assert abs(out["粗利_税抜"] - (-231.090909)) < 1e-6
    assert abs(out["粗利率"] - (-231.090909 / 3090.909091)) < 1e-6
    # 送料PL（行ってこい確認）= 1090.909091 - 1410 = -319.090909
    assert abs(out["送料PL_通常時"] - (-319.090909)) < 1e-6
    # 決済手数料は粗利の外。参考列に税込×3.25%＝70.2 が載るのみ。
    assert abs(out["決済手数料_販管費"] - 70.2) < 1e-6
    assert abs(out["手数料控除後利益_参考"] - (-231.090909 - 70.2)) < 1e-6


def test_粗利率20パーセントが再現すること() -> None:
    """売上=原価/0.8 となる販売価格を入れると粗利率がちょうど20%になること.

    送料込（原価＝仕入税抜のみ）・常温関西で検算する。
    原価=5000 → 売上目標=6250。売上=販売税抜＋消費者送料税抜(1000/1.1=909.090909)
    → 販売税抜=5340.909091, 販売税込(8%)=5768.181818。
    """
    rec = _umeya_a({
        "温度帯": "常温", "仕入単価_値": "5000", "仕入単価_税区分": "税抜",
        "仕入送料区分": "送料込", "参考上代_値": "", "想定販売価格_税込": "5768.181818",
    })
    out, _ = b.build_row(rec, {}, _cfg())
    assert out["売上原価_税抜"] == 5000
    assert abs(out["売上_税抜"] - 6250) < 1e-3
    assert abs(out["粗利率"] - 0.2) < 1e-4


def test_送料テーブル未格納なら仕入送料は空欄_要確認() -> None:
    """推定禁止: 送料見積が無ければ仮置きせず空欄＋要確認."""
    out, flags = b.build_row(_umeya_a(), {}, _cfg())  # テーブル空
    assert out["仕入送料_基準地域_税抜"] == ""
    assert out["売上原価_税抜"] == ""
    assert any("送料見積が未格納" in x for x in flags)


def test_酒類疑いは税率を確定せず空欄() -> None:
    """推定禁止: 酒類疑いかつ税率明記なしは8/10を確定しない."""
    rec = _umeya_a({"商品名_見積準拠": "NEIRO 2024（2本）", "税率_明記": "", "温度帯": "常温"})
    out, flags = b.build_row(rec, {}, _cfg())
    assert out["tax_rate"] == ""
    assert any("酒類の可能性" in x for x in flags)


def test_税区分不明なら税抜換算しない() -> None:
    """推定禁止: 税込/税抜不明の金額は換算せず空欄."""
    rec = _umeya_a({"仕入単価_税区分": "不明"})
    out, flags = b.build_row(rec, _UMEYA_SHIP, _cfg())
    assert out["仕入単価_税抜"] == ""
    assert any("税区分が不明" in x for x in flags)


def test_送料込は仕入送料ゼロで計算継続() -> None:
    """送料込なら仕入送料0として採算が出ること."""
    rec = _umeya_a({"仕入送料区分": "送料込", "温度帯": "冷凍", "想定販売価格_税込": "9500",
                    "仕入単価_値": "8251"})
    out, _ = b.build_row(rec, {}, _cfg())
    assert out["仕入送料_基準地域_税抜"] == 0
    assert out["売上原価_税抜"] == 8251
    # 送料込でも消費者送料は受領 → 送料PL通常 = 消費者送料税抜(1200/1.1) - 0
    assert abs(out["送料PL_通常時"] - 1200 / 1.1) < 1e-6


def _run_all() -> None:
    """pytest非依存の簡易ランナー."""
    funcs = [v for k, v in globals().items() if k.startswith("test_") and callable(v)]
    for fn in funcs:
        fn()
        print(f"PASS: {fn.__name__}")
    print(f"\n{len(funcs)} passed")


if __name__ == "__main__":
    _run_all()

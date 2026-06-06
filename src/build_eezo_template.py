#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""仕入先の中間CSVを EEZO Shopify登録用テンプレートへ整形するビルドスクリプト.

processed/ 配下の中間CSV（商品・仕入送料）を読み込み、出力テンプレート・採算
サマリ・要確認一覧を1つのExcelに統合出力する。あわせて shopify-product-register
連携用のCSVを書き出す。

設計原則:
    - 推定で埋めない。見積に無い事実は空欄＋要確認とする。
    - 代理推定（他社送料での仮置き等）は禁止。
    - 許可する自動導出は「情報を創作しない決定論的変換」のみ。
    - 消費者負担送料は公開ポリシー（config/eezo_policy.json）を正とする。
    - 採算は基準地域（既定: 関西）着で1点に集約して評価する。

Example:
    python src/build_eezo_template.py \\
        --processed data/processed \\
        --config config/eezo_policy.json \\
        --out outputs/YYYYMMDD_DATA_EEZO_seibi_output.xlsx
"""

from __future__ import annotations

import argparse
import csv
import glob
import json
import os
from datetime import date
from typing import Any, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ------------------------------------------------------------------
# 体裁（xlsx-format-override 準拠: Meiryo UI / 白文字禁止 / 最小配色）
# ------------------------------------------------------------------
FONT = Font(name="Meiryo UI", size=10)
FONT_BOLD = Font(name="Meiryo UI", size=10, bold=True)
FONT_TITLE = Font(name="Meiryo UI", size=12, bold=True)
FILL_HEADER = PatternFill("solid", fgColor="D9D9D9")
FILL_INPUT = PatternFill("solid", fgColor="DCE6F1")   # 手入力が要る列
FILL_CAUTION = PatternFill("solid", fgColor="FFF2CC")  # 空欄＋要確認
THIN = Border(bottom=Side(style="thin", color="000000"))

# 消費者負担送料の消費税率（送料は役務提供のため10%固定。税抜化に用いる）
SHIPPING_TAX_RATE = 10

# 酒類疑いキーワード（税率の自動確定を止め要確認へ回す判定。値は創作しない）
ALCOHOL_HINTS = [
    "酒", "日本酒", "純米", "吟醸", "ワイン", "ビール", "焼酎",
    "リキュール", "梅酒", "スパークリング", "NEIRO", "ml）", "720", "1800",
]

# 出力テンプレートの列定義: (列名, group, 由来)
#   由来: extract=見積転記 / derive=決定論導出 / input=手入力 / fixed=固定
COLUMNS: list[tuple[str, str, str]] = [
    ("仕入先", "meta", "extract"),
    ("商品名_見積準拠", "meta", "extract"),
    ("出典ファイル", "meta", "extract"),
    # ---- Shopify登録フィールド ----
    ("title", "shopify", "input"),
    ("productType", "shopify", "input"),
    ("vendor", "shopify", "fixed"),
    ("status", "shopify", "fixed"),
    ("温度帯", "shopify", "extract"),
    ("tags", "shopify", "derive"),
    ("variant_price_税込", "shopify", "input"),
    ("cost_price_税抜原価", "shopify", "derive"),
    ("tax_rate", "shopify", "derive"),
    ("supplier", "shopify", "extract"),
    ("supplier_product_name", "shopify", "extract"),
    ("商品特徴_説明素材", "shopify", "extract"),
    ("画像URL", "shopify", "input"),
    # ---- 商品仕様（カタログ用・見積記載のみ） ----
    ("内容量", "spec", "extract"),
    ("賞味期限", "spec", "extract"),
    ("保存方法", "spec", "extract"),
    ("発注ロット", "spec", "extract"),
    ("リードタイム", "spec", "extract"),
    ("JANコード", "spec", "extract"),
    ("アレルゲン", "spec", "extract"),
    ("産地", "spec", "extract"),
    ("参考上代_税抜", "spec", "derive"),
    # ---- 採算（基準地域着）。売上＝販売＋消費者送料 / 原価＝仕入＋箱代＋仕入送料 ----
    ("仕入単価_税抜", "econ", "derive"),
    ("箱代等_税抜", "econ", "derive"),
    ("仕入送料_基準地域_税抜", "econ", "derive"),
    ("売上原価_税抜", "econ", "derive"),
    ("想定販売価格_税込", "econ", "input"),
    ("想定販売価格_税抜", "econ", "derive"),
    ("消費者負担送料_基準地域", "econ", "derive"),
    ("売上_税抜", "econ", "derive"),
    ("粗利_税抜", "econ", "derive"),
    ("粗利率", "econ", "derive"),
    ("送料PL_通常時", "econ", "derive"),
    ("粗利_送料無料時", "econ", "derive"),
    ("粗利率_送料無料時", "econ", "derive"),
    # 決済手数料は販管費（粗利の外・参考表示のみ）
    ("決済手数料_販管費", "econ", "derive"),
    ("手数料控除後利益_参考", "econ", "derive"),
    ("採算フラグ", "econ", "derive"),
    # ---- 要確認 ----
    ("要確認メモ", "meta", "derive"),
]

SHOPIFY_IMPORT_COLS: list[str] = [
    "仕入先", "商品名_見積準拠", "title", "productType", "vendor", "status",
    "温度帯", "tags", "variant_price_税込", "cost_price_税抜原価", "tax_rate",
    "supplier", "supplier_product_name", "商品特徴_説明素材", "画像URL", "要確認メモ",
]


def to_float(value: Any) -> Optional[float]:
    """通貨記号・カンマを除去して数値化する. 換算不能・特殊表記は None.

    Args:
        value: 見積由来の生値.

    Returns:
        数値. オープン価格・時価・未受領・空などは None.
    """
    if value is None:
        return None
    s = str(value).strip().replace(",", "").replace("円", "").replace("¥", "").replace("￥", "")
    if s == "" or s in ("未受領", "オープン価格", "時価", "要相談", "不明", "-", "—"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def round6(x: Optional[float]) -> Optional[float | int]:
    """小数6桁で丸める. 整数値は int として返す.

    Args:
        x: 対象値.

    Returns:
        丸めた値. None はそのまま.
    """
    if x is None:
        return None
    r = round(x, 6)
    return int(r) if r == int(r) else r


def temp_to_cool(temp: Optional[str], mapping: dict[str, str]) -> Optional[str]:
    """温度帯（常温/冷蔵/冷凍）を配送区分（常温/クール）へ写す.

    Args:
        temp: 見積記載の温度帯.
        mapping: config の温度帯マッピング.

    Returns:
        '常温' または 'クール'. 対応が無ければ None.
    """
    return mapping.get((temp or "").strip())


def to_excl(value: Optional[float], rate: Optional[int]) -> Optional[float]:
    """税込金額を税抜へ変換する. 税率不明なら変換不可.

    Args:
        value: 税込金額.
        rate: 税率（8 または 10）.

    Returns:
        税抜金額. 入力不足なら None.
    """
    if value is None or rate is None:
        return None
    return value / (1 + rate / 100.0)


def resolve_tax_rate(rec: dict[str, Any]) -> tuple[Optional[int], list[str]]:
    """税率を決定する. 明記>食品既定8（フラグ付）. 酒類疑いは空欄＋要確認.

    Args:
        rec: 商品中間CSVの1行.

    Returns:
        (税率 or None, フラグ一覧).
    """
    flags: list[str] = []
    stated = to_float(rec.get("税率_明記"))
    if stated in (8.0, 10.0):
        return int(stated), flags
    name = (rec.get("商品名_見積準拠") or "") + (rec.get("規格内容") or "")
    if any(h in name for h in ALCOHOL_HINTS):
        flags.append("税率: 酒類の可能性。明記なしのため空欄（要確認: 8%か10%か）")
        return None, flags
    flags.append("税率: 明記なし。食品前提で8%を仮適用（要確認）")
    return 8, flags


def normalize_amount(
    rec: dict[str, Any], val_key: str, kind_key: str, rate: Optional[int]
) -> tuple[Optional[float | int], Optional[str]]:
    """金額を税抜へ正規化する. 区分=税抜は素通し, 税込は変換, 不明は空欄.

    Args:
        rec: 商品中間CSVの1行.
        val_key: 金額列名.
        kind_key: 税区分列名.
        rate: 税率.

    Returns:
        (税抜金額 or None, フラグ or None).
    """
    v = to_float(rec.get(val_key))
    if v is None:
        return None, None
    kind = (rec.get(kind_key) or "").strip()
    if kind == "税抜":
        return round6(v), None
    if kind == "税込":
        excl = to_excl(v, rate)
        if excl is None:
            return None, f"{val_key}: 税込だが税率不明のため税抜換算不可"
        return round6(excl), None
    return None, f"{val_key}: 税区分が不明（税込/税抜の明記なし）。換算せず空欄"


def load_shipping_tables(processed_dir: str) -> dict[str, dict[str, Any]]:
    """processed/*_shipping.csv を仕入先ごとに読み込む.

    Args:
        processed_dir: 中間CSVのディレクトリ.

    Returns:
        {仕入先: {"体系": 送料体系, "rows": [行...]}}.
    """
    tables: dict[str, dict[str, Any]] = {}
    for path in glob.glob(os.path.join(processed_dir, "*_shipping.csv")):
        with open(path, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                sup = (row.get("仕入先") or "").strip()
                if not sup:
                    continue
                table = tables.setdefault(sup, {"体系": None, "rows": []})
                table["体系"] = (row.get("送料体系") or table["体系"] or "").strip()
                table["rows"].append(row)
    return tables


def resolve_inbound_shipping(
    rec: dict[str, Any],
    tables: dict[str, dict[str, Any]],
    ref_region: str,
    temp_map: dict[str, str],
) -> tuple[Optional[float | int], Optional[str]]:
    """仕入送料（基準地域着）を決定する. 不確定は推定せず空欄＋要確認.

    Args:
        rec: 商品中間CSVの1行.
        tables: 仕入先別の送料テーブル.
        ref_region: 基準地域.
        temp_map: 温度帯マッピング.

    Returns:
        (仕入送料 or None, フラグ or None).
    """
    区分 = (rec.get("仕入送料区分") or "").strip()
    sup = (rec.get("仕入先") or "").strip()
    temp = (rec.get("温度帯") or "").strip()
    if 区分 in ("送料込", "込", "0"):
        return 0, None
    if 区分 not in ("産直別途", "別途", "別"):
        return None, "仕入送料: 区分不明（送料込/産直別途の明記なし）。空欄"
    table = tables.get(sup)
    if not table:
        return None, f"仕入送料: {sup}の送料見積が未格納。空欄（要取得）"
    体系 = (table.get("体系") or "").strip()
    if 体系 == "温度帯型":
        cool = temp_to_cool(temp, temp_map)
        for row in table["rows"]:
            if (row.get("キー種別") or "").strip() != "温度帯":
                continue
            if (row.get("地域") or "").strip() != ref_region:
                continue
            kv = (row.get("キー値") or "").strip()
            matched = (
                temp == kv
                or (cool == "クール" and ("冷" in kv or "クール" in kv))
                or (cool == "常温" and "常温" in kv)
            )
            if not matched:
                continue
            base = to_float(row.get("送料_値"))
            add = to_float(row.get("クール付加")) or 0
            if base is None:
                return None, f"仕入送料: テーブルに{ref_region}の値なし。空欄"
            total = base + (add if cool == "クール" else 0)
            return round6(total), None
        return None, f"仕入送料: {ref_region}・{temp}に一致する行なし。空欄"
    if 体系 == "サイズ型":
        return None, "仕入送料: サイズ型テーブル。荷姿サイズが見積に無いため自動引当不可（手動・要確認）"
    return None, "仕入送料: 送料体系不明。空欄"


def build_row(
    rec: dict[str, Any], tables: dict[str, dict[str, Any]], cfg: dict[str, Any]
) -> tuple[dict[str, Any], list[str]]:
    """商品1行を出力テンプレート行へ整形し採算を算出する.

    Args:
        rec: 商品中間CSVの1行.
        tables: 仕入先別の送料テーブル.
        cfg: 送料ポリシー設定.

    Returns:
        (出力行dict, フラグ一覧).
    """
    ref = cfg["_meta"]["基準地域"]
    temp_map = cfg["温度帯マッピング"]
    free_line = cfg["送料無料ライン_税込"]
    fee_rate = cfg["決済手数料率"]
    cool_table = cfg["消費者負担送料_税込"]

    flags: list[str] = []
    out: dict[str, Any] = {c[0]: "" for c in COLUMNS}

    # --- 識別・転記（事実のみ） ---
    out["仕入先"] = rec.get("仕入先", "")
    out["商品名_見積準拠"] = rec.get("商品名_見積準拠", "")
    out["出典ファイル"] = rec.get("出典ファイル", "")
    out["supplier"] = rec.get("仕入先", "")
    out["supplier_product_name"] = rec.get("商品名_見積準拠", "")
    temp = (rec.get("温度帯") or "").strip()
    out["温度帯"] = temp
    if not temp:
        flags.append("温度帯: 見積に明記なし。空欄（推定せず・要確認）")
    for col in ["内容量", "賞味期限", "保存方法", "発注ロット", "リードタイム",
                "JANコード", "アレルゲン", "産地"]:
        out[col] = rec.get(col, "") or ""
    out["商品特徴_説明素材"] = rec.get("商品特徴", "") or ""

    # --- 固定値 ---
    out["vendor"] = "EEZO（エエゾ）"
    out["status"] = "DRAFT"

    # --- 税率 ---
    rate, rflags = resolve_tax_rate(rec)
    flags += rflags
    out["tax_rate"] = rate if rate is not None else ""

    # --- 価格の税抜化 ---
    cost_excl, f_cost = normalize_amount(rec, "仕入単価_値", "仕入単価_税区分", rate)
    if f_cost:
        flags.append(f_cost)
    out["仕入単価_税抜"] = cost_excl if cost_excl is not None else ""
    out["cost_price_税抜原価"] = cost_excl if cost_excl is not None else ""
    if cost_excl is None:
        flags.append("仕入単価: 税抜が確定できず（値・税区分のいずれか不足）")

    box_excl, f_box = normalize_amount(rec, "箱代等_値", "箱代等_税区分", rate)
    if f_box:
        flags.append(f_box)
    out["箱代等_税抜"] = box_excl if box_excl is not None else ""

    jodai_excl, _ = normalize_amount(rec, "参考上代_値", "参考上代_税区分", rate)
    if jodai_excl is not None:
        out["参考上代_税抜"] = jodai_excl

    # --- 仕入送料（基準地域着） ---
    inbound, f_ship = resolve_inbound_shipping(rec, tables, ref, temp_map)
    if f_ship:
        flags.append(f_ship)
    out["仕入送料_基準地域_税抜"] = inbound if inbound is not None else ""

    # --- 売上原価（仕入＋箱代＋仕入送料。必須要素が揃った時のみ） ---
    if cost_excl is not None and inbound is not None:
        売上原価 = round6(cost_excl + (box_excl or 0) + inbound)
        out["売上原価_税抜"] = 売上原価
    else:
        売上原価 = None
        flags.append("売上原価: 仕入単価/仕入送料のいずれか不足で算出不可")

    # --- 想定販売価格（手入力。創作しない） ---
    sale_in = to_float(rec.get("想定販売価格_税込"))
    sale_excl: Optional[float] = None
    if sale_in is not None:
        out["variant_price_税込"] = round6(sale_in)
        out["想定販売価格_税込"] = round6(sale_in)
        sale_excl = to_excl(sale_in, rate)
        out["想定販売価格_税抜"] = round6(sale_excl) if sale_excl is not None else ""
        # 決済手数料は販管費。粗利には含めず参考表示のみ。
        out["決済手数料_販管費"] = round6(sale_in * fee_rate)
    else:
        flags.append("想定販売価格: 未入力。採算は算出されません（手入力で確定）")

    # --- 消費者負担送料（公開ポリシー・基準地域・温度帯別。税抜化して売上へ算入） ---
    cool = temp_to_cool(temp, temp_map)
    cust_ship_excl: Optional[float] = None
    if cool and ref in cool_table.get(cool, {}):
        cust_ship = cool_table[cool][ref]
        out["消費者負担送料_基準地域"] = cust_ship
        cust_ship_excl = cust_ship / (1 + SHIPPING_TAX_RATE / 100.0)
    elif temp:
        flags.append("消費者負担送料: 温度帯→常温/クール対応が取れず空欄")

    # --- 採算（売上＝販売税抜＋消費者送料税抜 / 原価＝売上原価。決済手数料は粗利外） ---
    粗利: Optional[float | int] = None
    if sale_excl is not None and 売上原価 is not None:
        # 送料無料時（消費者送料を受け取らない＝ワーストケース）
        粗利無料 = round6(sale_excl - 売上原価)
        out["粗利_送料無料時"] = 粗利無料
        if sale_excl:
            out["粗利率_送料無料時"] = round6(粗利無料 / sale_excl)
        # 通常時（消費者送料を受領）
        if cust_ship_excl is not None and inbound is not None:
            売上 = round6(sale_excl + cust_ship_excl)
            out["売上_税抜"] = 売上
            粗利 = round6(売上 - 売上原価)
            out["粗利_税抜"] = 粗利
            if 売上:
                out["粗利率"] = round6(粗利 / 売上)
            # 送料は行ってこい確認用（消費者送料 − 仕入送料、ともに税抜）
            out["送料PL_通常時"] = round6(cust_ship_excl - inbound)
            if sale_in is not None and 粗利 is not None:
                out["手数料控除後利益_参考"] = round6(粗利 - sale_in * fee_rate)
        out["採算フラグ"] = _profit_flags(out, 粗利, jodai_excl, sale_excl, sale_in, free_line)

    # --- tags（決定論的に組成。用途タグは手動） ---
    out["tags"] = _build_tags(temp, out["仕入先"], out["産地"])
    out["要確認メモ"] = " / ".join(flags)
    return out, flags


def _profit_flags(
    out: dict[str, Any],
    粗利: Optional[float | int],
    jodai_excl: Optional[float | int],
    sale_excl: Optional[float],
    sale_in: Optional[float],
    free_line: int,
) -> str:
    """採算フラグ文字列を組み立てる."""
    af: list[str] = []
    if 粗利 is not None and 粗利 < 0:
        af.append("通常時粗利が赤字（逆ざや）")
    if out["粗利_送料無料時"] != "" and out["粗利_送料無料時"] < 0:
        af.append("送料無料時粗利が赤字")
    if jodai_excl is not None and sale_excl is not None and sale_excl > jodai_excl:
        af.append("想定販売価格が参考上代を上回る")
    if sale_in is not None and sale_in >= free_line:
        af.append(f"単品で送料無料ライン({free_line:,}円)到達。送料無料時粗利を確認")
    return " / ".join(af)


def _build_tags(temp: str, supplier: str, origin: str) -> str:
    """温度帯・仕入先・産地から tags を決定論的に組成する."""
    parts: list[str] = []
    if temp:
        parts.append("チルド" if temp in ("冷蔵", "チルド") else temp)
    if supplier:
        parts.append(supplier)
    if origin:
        parts.append(origin)
    return ",".join(parts)


def write_excel(rows: list[dict[str, Any]], cfg: dict[str, Any], out_path: str) -> None:
    """出力テンプレ・採算サマリ・要確認・設定の4シートExcelを書き出す.

    Args:
        rows: 出力行一覧.
        cfg: 送料ポリシー設定.
        out_path: 出力xlsxパス.
    """
    wb = openpyxl.Workbook()
    _write_main_sheet(wb.active, rows)
    _write_summary_sheet(wb.create_sheet("採算サマリ"), rows, cfg)
    _write_caution_sheet(wb.create_sheet("要確認一覧"), rows)
    _write_config_sheet(wb.create_sheet("設定"), cfg)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)


def _write_main_sheet(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    """出力テンプレシートを書く."""
    ws.title = "出力テンプレ"
    headers = [c[0] for c in COLUMNS]
    origins = {c[0]: c[2] for c in COLUMNS}
    ws.append(headers)
    for ci, _ in enumerate(headers, 1):
        cell = ws.cell(1, ci)
        cell.font = FONT_BOLD
        cell.fill = FILL_HEADER
        cell.border = THIN
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append([r[h] for h in headers])
    for ri in range(2, ws.max_row + 1):
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(ri, ci)
            cell.font = FONT
            val = cell.value
            if origins[h] == "input" and (val is None or val == ""):
                cell.fill = FILL_INPUT
            if h in ("要確認メモ", "採算フラグ") and val:
                cell.fill = FILL_CAUTION
    ws.freeze_panes = "D2"
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(10, min(28, len(h) + 4))


def _write_summary_sheet(ws: Worksheet, rows: list[dict[str, Any]], cfg: dict[str, Any]) -> None:
    """採算サマリシートを書く."""
    ws["A1"] = (
        f"採算サマリ（基準地域: {cfg['_meta']['基準地域']}着 / "
        f"売上＝販売＋消費者送料・原価＝仕入＋仕入送料 / "
        f"粗利は販管費(決済手数料{cfg['決済手数料率'] * 100:.2f}%)控除前 / "
        f"送料無料ライン: {cfg['送料無料ライン_税込']:,}円）"
    )
    ws["A1"].font = FONT_TITLE
    cols = ["仕入先", "商品名_見積準拠", "温度帯", "想定販売価格_税込",
            "売上_税抜", "売上原価_税抜", "粗利_税抜", "粗利率",
            "粗利_送料無料時", "採算フラグ"]
    ws.append([])
    ws.append(cols)
    for ci in range(1, len(cols) + 1):
        ws.cell(3, ci).font = FONT_BOLD
        ws.cell(3, ci).fill = FILL_HEADER
        ws.cell(3, ci).border = THIN
    for r in rows:
        ws.append([r[c] for c in cols])
    for ri in range(4, ws.max_row + 1):
        for ci, c in enumerate(cols, 1):
            ws.cell(ri, ci).font = FONT
            if c == "採算フラグ" and ws.cell(ri, ci).value:
                ws.cell(ri, ci).fill = FILL_CAUTION
    for ci, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(10, min(26, len(c) + 4))


def _write_caution_sheet(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    """要確認一覧シートを書く."""
    ws.append(["仕入先", "商品名_見積準拠", "区分", "内容"])
    for ci in range(1, 5):
        ws.cell(1, ci).font = FONT_BOLD
        ws.cell(1, ci).fill = FILL_HEADER
        ws.cell(1, ci).border = THIN
    for r in rows:
        for item in [x for x in (r.get("要確認メモ") or "").split(" / ") if x]:
            区分 = "赤旗" if ("赤字" in item or "逆ざや" in item) else "要確認"
            ws.append([r["仕入先"], r["商品名_見積準拠"], 区分, item])
        for item in [x for x in (r.get("採算フラグ") or "").split(" / ") if x]:
            ws.append([r["仕入先"], r["商品名_見積準拠"], "赤旗", item])
    for ri in range(2, ws.max_row + 1):
        is_red = ws.cell(ri, 3).value == "赤旗"
        for ci in range(1, 5):
            ws.cell(ri, ci).font = FONT
            if is_red:
                ws.cell(ri, ci).fill = FILL_CAUTION
    for ci, w in enumerate([16, 28, 8, 70], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def _write_config_sheet(ws: Worksheet, cfg: dict[str, Any]) -> None:
    """設定スナップショットシートを書く."""
    ws.append(["項目", "値"])
    ws.append(["出典URL", cfg["_meta"]["出典"]])
    ws.append(["確認日", cfg["_meta"]["確認日"]])
    ws.append(["基準地域", cfg["_meta"]["基準地域"]])
    ws.append(["送料無料ライン(税込)", cfg["送料無料ライン_税込"]])
    ws.append(["決済手数料率", cfg["決済手数料率"]])
    for ci in range(1, 3):
        ws.cell(1, ci).font = FONT_BOLD
        ws.cell(1, ci).fill = FILL_HEADER
    for ri in range(2, ws.max_row + 1):
        for ci in range(1, 3):
            ws.cell(ri, ci).font = FONT
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 60


def write_shopify_csv(rows: list[dict[str, Any]], out_path: str) -> None:
    """shopify-product-register 連携用CSVを書き出す.

    Args:
        rows: 出力行一覧.
        out_path: 出力CSVパス.
    """
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=SHOPIFY_IMPORT_COLS, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in SHOPIFY_IMPORT_COLS})


def load_products(processed_dir: str) -> list[dict[str, Any]]:
    """processed/*_products.csv を全件読み込む.

    Args:
        processed_dir: 中間CSVのディレクトリ.

    Returns:
        商品レコードのリスト.

    Raises:
        SystemExit: 商品中間CSVが1件も無い場合.
    """
    paths = sorted(glob.glob(os.path.join(processed_dir, "*_products.csv")))
    if not paths:
        raise SystemExit(f"商品中間CSVが見つかりません: {processed_dir}/*_products.csv")
    records: list[dict[str, Any]] = []
    for path in paths:
        with open(path, encoding="utf-8-sig") as f:
            for rec in csv.DictReader(f):
                if (rec.get("商品名_見積準拠") or "").strip():
                    records.append(rec)
    return records


def run(processed_dir: str, config_path: str, out_path: str, csv_path: Optional[str]) -> None:
    """整形パイプラインを実行する.

    Args:
        processed_dir: 中間CSVのディレクトリ.
        config_path: 送料ポリシーJSON.
        out_path: 出力xlsxパス.
        csv_path: Shopify取込CSVパス（None なら out から自動命名）.
    """
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)
    tables = load_shipping_tables(processed_dir)
    records = load_products(processed_dir)

    rows: list[dict[str, Any]] = []
    for rec in records:
        row, _ = build_row(rec, tables, cfg)
        rows.append(row)

    write_excel(rows, cfg, out_path)
    csv_out = csv_path or out_path.replace(".xlsx", "_shopify_import.csv")
    write_shopify_csv(rows, csv_out)

    n_flag = sum(1 for r in rows if r.get("要確認メモ") or r.get("採算フラグ"))
    print(f"出力: {out_path}")
    print(f"Shopify取込CSV: {csv_out}")
    print(f"商品数: {len(rows)} / 要確認あり: {n_flag}")


def main() -> None:
    """CLIエントリポイント."""
    parser = argparse.ArgumentParser(description="EEZO整える ビルドスクリプト")
    parser.add_argument("--processed", default="data/processed")
    parser.add_argument("--config", default="config/eezo_policy.json")
    parser.add_argument("--out", default=f"outputs/{date.today():%Y%m%d}_DATA_EEZO_seibi_output.xlsx")
    parser.add_argument("--csv", default=None, help="Shopify取込CSVの出力先")
    args = parser.parse_args()
    run(args.processed, args.config, args.out, args.csv)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Reconcile incentive-store budgets against local promotion spend exports."""

from __future__ import annotations

import argparse
import csv
import difflib
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


VALID_BUDGETS = {Decimal("300"), Decimal("600"), Decimal("900")}
FALSE_FLAGS = {"否", "无", "无激励", "无流量激励", "不参与", "未参与", "false", "0", "no"}
ALIASES = {
    "name": ["激励名单门店", "门店名称", "门店", "店铺名称", "抖音门店名称", "名称"],
    "budget": ["激励预算（元）", "激励预算", "预算（元）", "预算", "预算金额", "投流预算（元）"],
    "store_id": ["门店id", "店铺id", "抖音门店id", "门店编号", "poi_id"],
    "city": ["城市", "所在城市", "门店城市"],
    "incentive": ["是否有流量激励", "流量激励", "激励状态", "是否参与激励"],
    "match_status": ["匹配状态"],
    "consume": ["消耗（元）", "消耗", "实际消耗（元）", "实际消耗", "推广消耗", "花费（元）", "花费"],
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_name(value: Any) -> str:
    value = clean(value).lower()
    value = re.sub(r"[\s·•・,，。．、()（）【】\[\]{}：:;；_\-—]+", "", value)
    for suffix in ("瑜伽普拉提", "瑜伽", "普拉提"):
        if value.endswith(suffix) and len(value) > len(suffix) + 1:
            # Keep service words unless both sides have them; normalize them away
            # to make common export naming variants comparable.
            value = value[: -len(suffix)]
            break
    return value


def money(value: Any) -> Decimal:
    if value is None or clean(value) == "":
        return Decimal("0.00")
    text = clean(value).replace(",", "").replace("￥", "").replace("¥", "")
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return Decimal("0.00")


def normalize_header(value: Any) -> str:
    return re.sub(r"[\s（）()【】\[\]：:·._-]+", "", clean(value)).lower()


def sheet_rows(path: Path, sheet: str | None = None) -> list[list[Any]]:
    if path.suffix.lower() in {".csv", ".tsv"}:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            dialect = "excel-tab" if path.suffix.lower() == ".tsv" else "excel"
            return list(csv.reader(f, dialect=dialect))
    wb = load_workbook(path, read_only=True, data_only=True)
    if sheet:
        if sheet not in wb.sheetnames:
            raise ValueError(f"{path.name} 中找不到工作表：{sheet}")
        ws = wb[sheet]
    else:
        ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def find_header(rows: list[list[Any]], required: Iterable[str]) -> tuple[int, dict[str, int]]:
    aliases = {k: [normalize_header(v) for v in ALIASES[k]] for k in required}
    best: tuple[int, dict[str, int]] | None = None
    for i, row in enumerate(rows[:30]):
        heads = {normalize_header(v): j for j, v in enumerate(row) if clean(v)}
        found = {k: next((heads[a] for a in options if a in heads), -1) for k, options in aliases.items()}
        score = sum(v >= 0 for v in found.values())
        if best is None or score > sum(v >= 0 for v in best[1].values()):
            best = (i, found)
    minimum = len(tuple(required))
    if best is None or sum(v >= 0 for v in best[1].values()) < minimum:
        keys = ", ".join(required)
        raise ValueError(f"未识别到所需表头（{keys}），请检查导出文件列名。")
    return best


@dataclass
class RosterEntry:
    name: str
    budget: Decimal
    city: str = ""
    store_id: str = ""
    match_status: str = ""
    source: str = ""


@dataclass
class StoreSpend:
    store_id: str
    name: str
    city: str
    spend: Decimal = Decimal("0.00")


def load_roster(paths: list[Path], sheets: list[str | None] | None = None) -> tuple[list[RosterEntry], list[dict[str, Any]]]:
    entries: list[RosterEntry] = []
    excluded: list[dict[str, Any]] = []
    for idx, path in enumerate(paths):
        rows = sheet_rows(path, sheets[idx] if sheets and idx < len(sheets) else None)
        req = ["name", "budget"]
        header_i, cols = find_header(rows, req)
        opt_i, opt_cols = find_header(rows, ["name", "budget"])
        del opt_i
        # Locate optional columns on the selected header row.
        header = {normalize_header(v): j for j, v in enumerate(rows[header_i]) if clean(v)}
        for key in ("store_id", "city", "incentive", "match_status"):
            cols[key] = next((header[normalize_header(v)] for v in ALIASES[key] if normalize_header(v) in header), -1)
        for row in rows[header_i + 1 :]:
            name = clean(row[cols["name"]] if cols["name"] < len(row) else "")
            if not name:
                continue
            raw_budget = row[cols["budget"]] if cols["budget"] < len(row) else None
            budget = money(raw_budget)
            flag = clean(row[cols["incentive"]]).lower() if cols["incentive"] >= 0 and cols["incentive"] < len(row) else ""
            if flag in FALSE_FLAGS or (not budget and flag):
                excluded.append({"门店名称": name, "原因": "无流量激励，已剔除", "来源文件": path.name})
                continue
            if budget not in VALID_BUDGETS:
                excluded.append({"门店名称": name, "原因": f"预算缺失或不在300/600/900元档（原值：{clean(raw_budget) or '空'}）", "来源文件": path.name})
                continue
            def val(key: str) -> str:
                c = cols[key]
                return clean(row[c]) if c >= 0 and c < len(row) else ""
            entries.append(RosterEntry(name, budget, val("city"), val("store_id"), val("match_status"), path.name))
    return entries, excluded


def load_spend(paths: list[Path], sheets: list[str | None] | None = None) -> dict[str, StoreSpend]:
    stores: dict[str, StoreSpend] = {}
    for idx, path in enumerate(paths):
        rows = sheet_rows(path, sheets[idx] if sheets and idx < len(sheets) else None)
        header_i, cols = find_header(rows, ["store_id", "name", "consume"])
        header = {normalize_header(v): j for j, v in enumerate(rows[header_i]) if clean(v)}
        for key in ("city",):
            cols[key] = next((header[normalize_header(v)] for v in ALIASES[key] if normalize_header(v) in header), -1)
        for row in rows[header_i + 1 :]:
            def val(key: str) -> str:
                c = cols[key]
                return clean(row[c]) if c >= 0 and c < len(row) else ""
            store_id = val("store_id")
            name = val("name")
            if not store_id or not name:
                continue
            city = val("city")
            spend = money(row[cols["consume"]] if cols["consume"] < len(row) else None)
            if store_id not in stores:
                stores[store_id] = StoreSpend(store_id, name, city)
            item = stores[store_id]
            if item.city and city and item.city != city:
                item.city = f"{item.city}/{city}"
            elif not item.city:
                item.city = city
            item.spend += spend
    return stores


def load_mapping(path: Path | None) -> dict[str, str]:
    if not path:
        return {}
    rows = sheet_rows(path)
    header_i, cols = find_header(rows, ["name", "store_id"])
    mapping: dict[str, str] = {}
    for row in rows[header_i + 1 :]:
        name = clean(row[cols["name"]] if cols["name"] < len(row) else "")
        store_id = clean(row[cols["store_id"]] if cols["store_id"] < len(row) else "")
        if name and store_id:
            mapping[normalize_name(name)] = store_id
    return mapping


def reconcile(entries: list[RosterEntry], stores: dict[str, StoreSpend], overrides: dict[str, str] | None = None) -> dict[str, list[dict[str, Any]]]:
    overrides = overrides or {}
    by_name: dict[str, list[StoreSpend]] = defaultdict(list)
    for store in stores.values():
        by_name[normalize_name(store.name)].append(store)
    results: list[dict[str, Any]] = []
    pending: list[dict[str, Any]] = []
    used_ids: set[str] = set()
    store_names = {s.store_id: s.name for s in stores.values()}

    for entry in entries:
        target_id = overrides.get(normalize_name(entry.name), "") or entry.store_id
        status = ""
        reason = ""
        matched: StoreSpend | None = None
        if entry.match_status in {"待确认", "需确认", "无法确认"} and not target_id:
            reason = "历史/名单已标记待确认；需人工确认"
        elif target_id:
            if target_id in stores:
                candidate = stores[target_id]
                if entry.city and candidate.city and entry.city not in candidate.city.split("/") and candidate.city not in entry.city:
                    reason = f"门店ID一致但城市不一致（名单：{entry.city}；后台：{candidate.city}）"
                else:
                    matched = candidate
                    status = "规则匹配" if entry.store_id else "人工映射"
            elif entry.store_id:
                # A confirmed ID already present in the roster remains valid even
                # if this export contains no spend rows for that store.
                matched = StoreSpend(target_id, entry.name, entry.city, Decimal("0.00"))
                status = "名单门店ID"
            else:
                reason = f"映射门店ID在本次消耗表中不存在：{target_id}"
        else:
            candidates = by_name.get(normalize_name(entry.name), [])
            if entry.city:
                candidates = [s for s in candidates if not s.city or entry.city in s.city.split("/") or s.city in entry.city]
            if len(candidates) == 1:
                matched = candidates[0]
                status = "名称+城市匹配" if entry.city else "唯一名称匹配"
            elif len(candidates) > 1:
                reason = "同名候选不唯一；需结合门店ID和城市确认"
            else:
                reason = "未找到可确认的同名门店"
        if not matched:
            norm = normalize_name(entry.name)
            similar = difflib.get_close_matches(norm, list(by_name), n=3, cutoff=0.45)
            suggestion = "；".join(f"{store_names[by_name[n][0].store_id]}（{by_name[n][0].city}，ID {by_name[n][0].store_id}）" for n in similar if by_name[n])
            pending.append({"激励名单门店": entry.name, "激励预算（元）": float(entry.budget), "城市": entry.city,
                            "原因": reason, "候选提示（不自动匹配）": suggestion, "来源文件": entry.source})
            continue
        if matched.store_id in used_ids:
            pending.append({"激励名单门店": entry.name, "激励预算（元）": float(entry.budget), "城市": entry.city,
                            "原因": "该门店ID已被其他激励名单行使用；需检查是否重复名单", "候选提示（不自动匹配）": matched.name, "来源文件": entry.source})
            continue
        used_ids.add(matched.store_id)
        spend = matched.spend.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        ratio = spend / entry.budget if entry.budget else Decimal(0)
        if spend > entry.budget:
            state = "已超预算"
        elif ratio >= Decimal("0.80"):
            state = "临近预算"
        else:
            state = "正常"
        results.append({"激励名单门店": entry.name, "激励预算（元）": float(entry.budget), "后台门店名称": matched.name,
                        "城市": matched.city or entry.city, "门店ID": matched.store_id, "实际消耗（元）": float(spend),
                        "剩余/超支（元）": float(entry.budget - spend), "预算使用率": float(ratio), "状态": state,
                        "匹配方式": status, "来源文件": entry.source})
    overspend = [r for r in results if r["状态"] == "已超预算"]
    near = [r for r in results if r["状态"] == "临近预算"]
    return {"results": results, "pending": pending, "overspend": overspend, "near": near}


def write_report(out: Path, report: dict[str, list[dict[str, Any]]], excluded: list[dict[str, Any]]) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "核对摘要"
    matched = report["results"]
    summary_rows = [
        ["门店投流预算核对摘要", ""],
        ["生成时间", __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["已匹配门店", len(matched)], ["待确认门店", len(report["pending"])],
        ["已匹配预算（元）", round(sum(x["激励预算（元）"] for x in matched), 2)],
        ["实际消耗（元）", round(sum(x["实际消耗（元）"] for x in matched), 2)],
        ["超预算门店", len(report["overspend"])],
        ["超支金额（元）", round(sum(max(0, -x["剩余/超支（元）"]) for x in report["overspend"]), 2)],
        ["临近预算门店", len(report["near"])], ["无激励/异常预算已剔除", len(excluded)],
        ["规则", "消耗 > 预算：已超预算；消耗达到预算80%且不超过预算：临近预算。待确认门店不计入汇总。"],
    ]
    for row in summary_rows:
        summary.append(row)
    _write_rows(wb, "核对明细", matched)
    _write_rows(wb, "待确认", report["pending"])
    _write_rows(wb, "超预算", report["overspend"])
    _write_rows(wb, "临近预算", report["near"])
    _write_rows(wb, "已剔除", excluded)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else "A1"
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="305496")
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = min(48, max(12, max((len(clean(c.value)) for c in col), default=10) + 2))
            ws.column_dimensions[letter].width = width
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.value is not None and cell.column == 8 and ws.title == "核对明细":
                    cell.number_format = "0.0%"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def _write_rows(wb: Workbook, title: str, rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(title)
    headers = list(rows[0]) if rows else {
        "核对明细": ["激励名单门店", "激励预算（元）", "后台门店名称", "城市", "门店ID", "实际消耗（元）", "剩余/超支（元）", "预算使用率", "状态", "匹配方式", "来源文件"],
        "待确认": ["激励名单门店", "激励预算（元）", "城市", "原因", "候选提示（不自动匹配）", "来源文件"],
        "超预算": ["激励名单门店", "激励预算（元）", "后台门店名称", "城市", "门店ID", "实际消耗（元）", "剩余/超支（元）", "预算使用率", "状态", "匹配方式", "来源文件"],
        "临近预算": ["激励名单门店", "激励预算（元）", "后台门店名称", "城市", "门店ID", "实际消耗（元）", "剩余/超支（元）", "预算使用率", "状态", "匹配方式", "来源文件"],
        "已剔除": ["门店名称", "原因", "来源文件"],
    }[title]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="导入激励名单和巨量本地推消耗表，生成门店预算核对Excel。")
    parser.add_argument("--roster", nargs="+", type=Path, required=True, help="激励名单Excel/CSV，可传多个文件")
    parser.add_argument("--spend", nargs="+", type=Path, required=True, help="后台消耗Excel/CSV，可传多个导出文件")
    parser.add_argument("--out", type=Path, default=Path("门店投流预算核对.xlsx"), help="结果文件路径")
    parser.add_argument("--roster-sheet", action="append", help="名单工作表名；多文件时按顺序重复传入")
    parser.add_argument("--spend-sheet", action="append", help="消耗工作表名；多文件时按顺序重复传入")
    parser.add_argument("--mapping", type=Path, help="可选人工确认映射CSV/XLSX，需含激励名单门店、门店ID")
    args = parser.parse_args(argv)
    try:
        entries, excluded = load_roster(args.roster, args.roster_sheet)
        stores = load_spend(args.spend, args.spend_sheet)
        report = reconcile(entries, stores, load_mapping(args.mapping))
        write_report(args.out, report, excluded)
    except (OSError, ValueError, KeyError) as exc:
        print(f"核对失败：{exc}", file=sys.stderr)
        return 2
    print(f"结果已保存：{args.out}")
    print(f"已匹配 {len(report['results'])} 家；待确认 {len(report['pending'])} 家；超预算 {len(report['overspend'])} 家；临近预算 {len(report['near'])} 家；已剔除 {len(excluded)} 家。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

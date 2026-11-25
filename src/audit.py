#!/usr/bin/env python3
"""Read-only, portable audit of the November 2026 Grade 4 workbook.

Python 3.9+ and the standard library are sufficient. The input XLSX is never
rewritten. Both the original five-sheet layout and the completed seven-sheet
layout are supported; a layout name is not a claim about file provenance.

Example:
    python src/audit.py --input inputs/2026.11期中考试四年级名单.xlsx --output-dir audit

Outputs: analysis.json, source_cells.json, portable_audit_report.json,
students.csv. Blank input scores remain null/empty, while a recorded numeric
zero remains 0. A reference total is calculated only when the base score is
numeric; a missing bonus then contributes no additional points. This is an
explicit provisional convention, not confirmation that the bonus score is 0.

The script independently recomputes statistics from input cells and checks
selected cached results. It is not a general Excel formula calculation engine.
"""

import argparse
import collections
import csv
import hashlib
import io
import json
import math
import posixpath
import re
import statistics
import sys
import unicodedata
import zipfile
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import xml.etree.ElementTree as ET


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN, "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
BLOCKS = [
    {"class": 1, "columns": "ABCDE", "first": 2, "last": 45},
    {"class": 2, "columns": "FGHIJ", "first": 2, "last": 45},
    {"class": 3, "columns": "KLMNO", "first": 2, "last": 43},
    {"class": 4, "columns": "PQRST", "first": 2, "last": 43},
    {"class": 5, "columns": "UVWXY", "first": 2, "last": 43},
]
# These are historical audit expectations, not values substituted into data.
BASELINE = {
    "names": 214, "base_count": 206, "bonus_count": 199,
    "base_sum": 15948, "total_sum": 17053, "colored_names": 19,
    "missing_base": 8, "missing_bonus_with_base": 7,
    "class_names": [44, 44, 42, 42, 42],
    "class_base_count": [42, 42, 40, 42, 40],
    "base_bands": [28, 9, 36, 73, 60],
}


def number(value):
    return type(value) in (int, float) and math.isfinite(value)


def blank(value):
    return value is None or value == ""


def parse_numeric(text):
    """Parse an XML numeric value, preserving integral scores as Python ints."""
    if text is None:
        return None
    try:
        value = Decimal(text)
        if not value.is_finite():
            return text
        return int(value) if value == value.to_integral_value() else float(value)
    except (InvalidOperation, ValueError, OverflowError):
        return text


def typed_value(cell):
    if cell is None or cell["value"] is None:
        return None
    if cell["type"] in (None, "n"):
        return parse_numeric(cell["value"])
    if cell["type"] == "b":
        return cell["value"] == "1"
    return cell["value"]


def numeric_equal(left, right):
    if number(left) and number(right):
        return math.isclose(left, right, rel_tol=1e-11, abs_tol=1e-10)
    return left == right


def metric_mean(values):
    return statistics.mean(values) if values else None


def metric_median(values):
    return statistics.median(values) if values else None


def metric_sd(values):
    return statistics.pstdev(values) if values else None


def band_counts(values, edges):
    counts = [0] * (len(edges) + 1)
    for value in values:
        slot = next((i for i, edge in enumerate(edges) if value < edge), len(edges))
        counts[slot] += 1
    return counts


def display_mean(total, count):
    if not count:
        return None
    return float((Decimal(str(total)) / Decimal(count)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def text_content(element):
    """Read visible text without accidentally concatenating phonetic runs."""
    if element is None:
        return ""
    parts = []
    for child in element:
        local = child.tag.rsplit("}", 1)[-1]
        if local == "t":
            parts.append(child.text or "")
        elif local == "r":
            parts.extend(node.text or "" for node in child.findall("m:t", NS))
    return "".join(parts)


def col_index(column):
    result = 0
    for char in column:
        result = result * 26 + ord(char) - 64
    return result


def column_name(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def translate_formula(formula, source, target):
    a = re.fullmatch(r"([A-Z]+)(\d+)", source)
    b = re.fullmatch(r"([A-Z]+)(\d+)", target)
    dc = col_index(b.group(1)) - col_index(a.group(1))
    dr = int(b.group(2)) - int(a.group(2))

    def replace(match):
        abs_col, col, abs_row, row = match.groups()
        ci = col_index(col) if abs_col else col_index(col) + dc
        ri = int(row) if abs_row else int(row) + dr
        if ci < 1 or ri < 1:
            return "#REF!"
        return abs_col + column_name(ci) + abs_row + str(ri)

    # Sufficient for this workbook's shared SUM/IF formulas. No formulas are
    # executed, and source_cells keeps the original formula text unchanged.
    return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", replace, formula)


class Workbook:
    def __init__(self, path):
        self.path = path
        self.raw = path.read_bytes()
        self.sha256 = hashlib.sha256(self.raw).hexdigest()
        self.archive = zipfile.ZipFile(io.BytesIO(self.raw))
        strings_path = "xl/sharedStrings.xml"
        self.strings = []
        if strings_path in self.archive.namelist():
            self.strings = [text_content(si) for si in ET.fromstring(self.archive.read(strings_path)).findall("m:si", NS)]
        workbook = ET.fromstring(self.archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(self.archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in relationships}
        self.sheets = []
        self.by_name = {}
        self.shared_anchors = {}
        for sheet in workbook.find("m:sheets", NS):
            target = targets[sheet.get("{" + REL + "}id")]
            path_in_zip = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            root = ET.fromstring(self.archive.read(path_in_zip))
            cells = {}
            anchors = {}
            for node in root.findall("m:sheetData/m:row/m:c", NS):
                coord = node.get("r")
                kind = node.get("t")
                v = node.find("m:v", NS)
                raw = v.text if v is not None else None
                if kind == "s" and raw is not None:
                    raw = self.strings[int(raw)]
                elif kind == "inlineStr":
                    raw = text_content(node.find("m:is", NS))
                f = node.find("m:f", NS)
                formula = f.text if f is not None else None
                attrs = dict(f.attrib) if f is not None else None
                cells[coord] = {"value": raw, "type": kind, "formula": formula,
                                "formula_attrs": attrs, "style": node.get("s")}
                if f is not None and f.get("t") == "shared" and formula:
                    anchors[f.get("si")] = (coord, formula)
            dim = root.find("m:dimension", NS)
            item = {"name": sheet.get("name"), "path": path_in_zip,
                    "state": sheet.get("state", "visible"),
                    "dimension": dim.get("ref") if dim is not None else None,
                    "merges": [m.get("ref") for m in root.findall("m:mergeCells/m:mergeCell", NS)],
                    "cells": cells}
            self.sheets.append(item)
            self.by_name[item["name"]] = item
            self.shared_anchors[item["name"]] = anchors
        styles = ET.fromstring(self.archive.read("xl/styles.xml"))
        self.fonts = styles.find("m:fonts", NS)
        self.fills = styles.find("m:fills", NS)
        self.styles = styles.find("m:cellXfs", NS)

    def cell(self, sheet, coord):
        return self.by_name[sheet]["cells"].get(coord)

    def value(self, sheet, coord):
        return typed_value(self.cell(sheet, coord))

    def formula(self, sheet, coord):
        cell = self.cell(sheet, coord)
        if not cell:
            return None
        if cell["formula"]:
            return cell["formula"]
        attrs = cell["formula_attrs"] or {}
        if attrs.get("t") == "shared":
            anchor = self.shared_anchors[sheet].get(attrs.get("si"))
            if anchor:
                return translate_formula(anchor[1], anchor[0], coord)
        return None

    def colors(self, sheet, coord):
        cell = self.cell(sheet, coord) or {}
        style = self.styles[int(cell.get("style") or "0")]
        font = self.fonts[int(style.get("fontId", "0"))]
        fill = self.fills[int(style.get("fillId", "0"))]
        fc = font.find("m:color", NS)
        pattern = fill.find("m:patternFill", NS)
        fg = pattern.find("m:fgColor", NS) if pattern is not None else None
        font_color = dict(fc.attrib) if fc is not None else None
        fill_color = dict(fg.attrib) if fg is not None else None
        frgb = (font_color or {}).get("rgb", "")[-6:].upper()
        brgb = (fill_color or {}).get("rgb", "")[-6:].upper()
        labels = []
        if frgb in {"FF0000": "红字", "00B050": "绿字", "00B0F0": "蓝字"}:
            labels.append({"FF0000": "红字", "00B050": "绿字", "00B0F0": "蓝字"}[frgb])
        if brgb == "FFFF00":
            labels.append("黄底")
        return {"font_color": font_color, "fill_color": fill_color,
                "fill_pattern": pattern.get("patternType") if pattern is not None else None,
                "mark": "、".join(labels), "is_marked": bool(labels)}


def source_ref(student):
    return {key: student[key] for key in ("class", "row", "name", "cell", "base", "bonus")}


def extract_students(book):
    if "四年级" not in book.by_name:
        raise ValueError("没有找到“四年级”工作表。")
    students = []
    for block in BLOCKS:
        columns = block["columns"]
        for row in range(block["first"], block["last"] + 1):
            coords = [c + str(row) for c in columns]
            values = [book.value("四年级", coord) for coord in coords]
            name, sex, base, bonus, raw_total = values
            if not isinstance(name, str) or not name.strip():
                continue
            formula = book.formula("四年级", coords[4]) or ""
            summary_match = re.fullmatch(r"=?SUM\(\$?" + columns[4] + r"\$?(\d+):\$?" + columns[4] + r"\$?(\d+)\)", formula, re.IGNORECASE)
            is_summary = bool(summary_match and summary_match.group(1) != summary_match.group(2))
            calculated = base + (bonus if number(bonus) else 0) if number(base) else None
            colors = book.colors("四年级", coords[0])
            students.append({"class": block["class"], "row": row, "name": name, "sex": sex,
                "base": base, "bonus": bonus, "original_total": raw_total,
                "raw_total": raw_total, "cell": coords[0], "name_cell": coords[0],
                "source_cells": dict(zip(["name", "sex", "base", "bonus", "total"], coords)),
                "total_is_summary": is_summary,
                "total_formula": formula or None,
                "total_with_blank_bonus_zero": calculated,
                "base_is_numeric": number(base), "bonus_is_numeric": number(bonus),
                "base_missing": blank(base), "bonus_missing": blank(bonus),
                "name_colors": colors})
    return students


def get_old_summary(book):
    completed = "数据核查" in book.by_name
    sheet = "数据核查" if completed else "四年级"
    rows = [6, 7, 8] if completed else [49, 50, 51]
    result = []
    for index, col in enumerate("BCDEF", 1):
        result.append({"class": index, "sum": book.value(sheet, col + str(rows[0])),
            "count": book.value(sheet, col + str(rows[1])), "mean": book.value(sheet, col + str(rows[2])),
            "source_sheet": sheet,
            "source_cells": [col + str(row) for row in rows]})
    return result


def summarize(group, old=None, class_number=None):
    bases = [s["base"] for s in group if number(s["base"])]
    bonuses = [s["bonus"] for s in group if number(s["bonus"])]
    totals = [s["total_with_blank_bonus_zero"] for s in group if number(s["total_with_blank_bonus_zero"])]
    missing = [source_ref(s) for s in group if blank(s["base"])]
    missing_bonus = [source_ref(s) for s in group if number(s["base"]) and blank(s["bonus"])]
    bad_total = []
    for s in group:
        if s["total_is_summary"] or not number(s["base"]):
            continue
        if not numeric_equal(s["original_total"], s["total_with_blank_bonus_zero"]):
            bad_total.append({**source_ref(s), "total_cell": s["source_cells"]["total"],
                "cached": s["original_total"], "recomputed": s["total_with_blank_bonus_zero"]})
    result = {"names": len(group), "base_count": len(bases), "bonus_count": len(bonuses),
        "base_sum": sum(bases), "bonus_sum": sum(bonuses), "sum": sum(totals),
        "total_sum": sum(totals), "mean": metric_mean(totals), "total_mean": metric_mean(totals),
        "base_mean": metric_mean(bases), "base_median": metric_median(bases), "base_sd": metric_sd(bases),
        "bonus_mean_if_blank_zero": metric_mean([s["bonus"] if number(s["bonus"]) else 0 for s in group if number(s["base"])]),
        "bonus_mean_observed": metric_mean(bonuses), "total_median": metric_median(totals),
        "base_bands": band_counts(bases, [60, 70, 80, 90]),
        "total_bands": band_counts(totals, [60, 70, 80, 90, 100]),
        "min": min(totals) if totals else None, "max": max(totals) if totals else None,
        "base_min": min(bases) if bases else None, "base_max": max(bases) if bases else None,
        "missing": missing, "missing_bonus": missing_bonus, "bad_total": bad_total,
        "low": [source_ref(s) for s in group if number(s["base"]) and s["base"] < 60],
        "sex": dict(collections.Counter(s["sex"] for s in group)),
        "complete_score_count": sum(number(s["base"]) and number(s["bonus"]) for s in group),
        "base_ge60_ratio": sum(v >= 60 for v in bases) / len(bases) if bases else None,
        "base_ge80_ratio": sum(v >= 80 for v in bases) / len(bases) if bases else None,
        "base_ge90_ratio": sum(v >= 90 for v in bases) / len(bases) if bases else None}
    if class_number is not None:
        result["class"] = class_number
    if old is not None:
        result.update({"old_sum": old["sum"], "old_count": old["count"], "old_mean": old["mean"],
            "difference": sum(totals) - old["sum"] if number(old["sum"]) else None})
    return result


def check_cached_statistics(book, classes, all_stats):
    checks = []
    if "四年级分析" not in book.by_name:
        return checks
    for col, stats in zip("BCDEFG", classes + [all_stats]):
        expected = {50: stats["names"], 51: stats["base_count"], 52: len(stats["missing"]),
            53: len(stats["missing_bonus"]), 54: stats["base_sum"], 55: stats["base_mean"],
            56: stats["base_median"], 57: stats["base_sd"], 63: stats["base_ge60_ratio"],
            64: stats["base_ge90_ratio"], 65: stats["total_sum"], 66: stats["total_mean"],
            67: stats["base_ge80_ratio"], 68: stats["complete_score_count"],
            69: stats["sex"].get("男", 0), 70: stats["sex"].get("女", 0)}
        expected.update({58 + i: count for i, count in enumerate(stats["base_bands"])})
        for row, value in expected.items():
            coord = col + str(row)
            observed = book.value("四年级", coord)
            checks.append({"sheet": "四年级", "cell": coord, "expected": value,
                "observed": observed, "pass": numeric_equal(value, observed)})
    for row, stats in zip(range(7, 13), classes + [all_stats]):
        for col, total in [("F", stats["base_sum"]), ("L", stats["total_sum"])]:
            coord = col + str(row)
            value = display_mean(total, stats["base_count"])
            observed = book.value("四年级分析", coord)
            checks.append({"sheet": "四年级分析", "cell": coord, "expected": value,
                "observed": observed, "pass": numeric_equal(value, observed)})
    return checks


def chart_cache_checks(book):
    results = []
    for path in book.archive.namelist():
        if "/charts/" not in path or not path.endswith(".xml"):
            continue
        root = ET.fromstring(book.archive.read(path))
        for kind, cache in [("strRef", "strCache"), ("numRef", "numCache")]:
            for ref in root.findall(".//c:" + kind, NS):
                formula = ref.find("c:f", NS)
                if formula is None or not formula.text:
                    continue
                text = formula.text
                match = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", text)
                if not match or match.group(3) != match.group(5):
                    results.append({"chart": path, "reference": text, "pass": None, "note": "引用形式超出本脚本核查范围"})
                    continue
                sheet = (match.group(1) or match.group(2)).replace("''", "'")
                expected = [book.value(sheet, match.group(3) + str(row)) for row in range(int(match.group(4)), int(match.group(6)) + 1)]
                count = ref.find("c:" + cache + "/c:ptCount", NS)
                points = ref.findall("c:" + cache + "/c:pt", NS)
                observed = [p.find("c:v", NS).text for p in sorted(points, key=lambda p: int(p.get("idx", "0")))]
                if kind == "numRef":
                    observed = [parse_numeric(v) for v in observed]
                ok = count is not None and int(count.get("val")) == len(expected) and len(observed) == len(expected)
                ok = ok and all(numeric_equal(a, b) for a, b in zip(observed, expected))
                results.append({"chart": path, "reference": text, "expected": expected, "observed": observed, "pass": ok})
    return results


def any_subset_sum(values, target):
    if not number(target) or target < 0 or int(target) != target:
        return None
    if any(not number(v) or v < 0 or int(v) != v for v in values):
        return None
    target = int(target)
    if target > 1000000:
        return None
    reachable = 1
    mask = (1 << (target + 1)) - 1
    for value in values:
        reachable = (reachable | (reachable << int(value))) & mask
    return bool((reachable >> target) & 1)


def make_report(book, students, classes, all_stats, old_summary):
    marked = [{**source_ref(s), "name_colors": s["name_colors"],
               "reference_total": s["total_with_blank_bonus_zero"]} for s in students if s["name_colors"]["is_marked"]]
    conditioned = []
    for stats, old in zip(classes, old_summary):
        group = [s for s in students if s["class"] == stats["class"]]
        included = [s for s in group if not s["name_colors"]["is_marked"]]
        total = sum(s["total_with_blank_bonus_zero"] for s in included if number(s["total_with_blank_bonus_zero"]))
        conditioned.append({"class": stats["class"], "marked_count": len(group) - len(included),
            "count": len(included), "sum": total, "mean": total / len(included) if included else None,
            "old_sum": old["sum"], "old_count": old["count"],
            "old_minus_conditioned": old["sum"] - total if number(old["sum"]) else None,
            "remaining_missing_bonus": [source_ref(s) for s in included if number(s["base"]) and blank(s["bonus"])]})
    duplicate_groups = collections.defaultdict(list)
    for s in students:
        duplicate_groups[s["name"]].append(source_ref(s))
    within, across = [], []
    for name, group in duplicate_groups.items():
        class_counts = collections.Counter(s["class"] for s in group)
        if any(count > 1 for count in class_counts.values()):
            within.append({"name": name, "records": group})
        if len(class_counts) > 1:
            across.append({"name": name, "records": group})
    odd_names, numeric_text, other_values, zeros = [], [], [], []
    for s in students:
        if any(char.isspace() or unicodedata.category(char).startswith("C") for char in s["name"]):
            odd_names.append(source_ref(s))
        for field in ["base", "bonus"]:
            value = s[field]
            coord = s["source_cells"][field]
            record = {**source_ref(s), "field": field, "score_cell": coord, "raw_value": value}
            if number(value) and value == 0:
                zeros.append(record)
            elif not blank(value) and not number(value):
                raw = book.cell("四年级", coord)
                if raw and raw["type"] in ("s", "inlineStr", "str") and number(parse_numeric(raw["value"])):
                    numeric_text.append(record)
                else:
                    other_values.append(record)
    observed_baseline = {"names": all_stats["names"], "base_count": all_stats["base_count"],
        "bonus_count": all_stats["bonus_count"], "base_sum": all_stats["base_sum"], "total_sum": all_stats["total_sum"],
        "colored_names": len(marked), "missing_base": len(all_stats["missing"]),
        "missing_bonus_with_base": len(all_stats["missing_bonus"]),
        "class_names": [c["names"] for c in classes], "class_base_count": [c["base_count"] for c in classes],
        "base_bands": all_stats["base_bands"]}
    baseline_checks = [{"metric": key, "expected": value, "observed": observed_baseline[key],
                        "pass": value == observed_baseline[key]} for key, value in BASELINE.items()]
    known5 = [s["total_with_blank_bonus_zero"] for s in students if s["class"] == 5 and number(s["total_with_blank_bonus_zero"])]
    subset_possible = any_subset_sum(known5, old_summary[4]["sum"])
    cached_checks = check_cached_statistics(book, classes, all_stats)
    chart_checks = chart_cache_checks(book)
    layout = "completed_seven_sheet_layout" if len(book.sheets) == 7 and "数据核查" in book.by_name else "original_five_sheet_layout" if len(book.sheets) == 5 else "other_layout"
    return {
        "schema_version": "1.0", "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "input": {"file_name": book.path.name, "size_bytes": len(book.raw), "sha256": book.sha256,
            "worksheet_names": [s["name"] for s in book.sheets], "layout": layout,
            "provenance_note": "布局识别不证明文件来源。仅核查本次输入文件；没有与另一份原始上传附件进行字节比对。"},
        "method": {
            "layout_scope": "按既定模板读取：1、2班第2—45行，3、4、5班第2—43行。如增删学生行或扩班，须先更新脚本BLOCKS和工作簿统计区域。",
            "main_population": "基础分已录为数值的学生；保留数值0，不按姓名颜色剔除。",
            "reference_total": "仅在基础分为数值时计算；附加分为空时暂不加分，不能认定已确认得0分。",
            "bands": "基础分按＜60、60—＜70、70—＜80、80—＜90、≥90分组；不是已确认的学校等级标准。",
            "standard_deviation": "总体标准差。",
            "preservation": "source_cells.json保留本次输入的原解码值、类型、公式、公式属性和样式编号；数字仍为字符串。",
            "limitations": "不是通用Excel计算引擎；未核验试卷、学籍、标色原因、缺分原因或正式统计范围。"},
        "actual_verification": {
            "score_input_cells_read": len(students) * 2,
            "student_rows_read": len(students),
            "numeric_base_records_recomputed": all_stats["base_count"],
            "complete_score_records_checked": all_stats["complete_score_count"],
            "student_cached_total_mismatches": all_stats["bad_total"],
            "selected_summary_cache_checks": cached_checks,
            "chart_reference_cache_checks": chart_checks,
            "input_unchanged_after_audit": hashlib.sha256(book.path.read_bytes()).hexdigest() == book.sha256},
        "historical_baseline_checks": baseline_checks,
        "historical_baseline_matches": all(c["pass"] for c in baseline_checks),
        "current_summary": all_stats, "class_summaries": classes,
        "original_manual_summary_as_stored_in_current_input": old_summary,
        "colored_names": marked,
        "color_exclusion_scenario": {
            "status": "仅作条件复算；颜色含义未经确认，不改变主分析名单。", "classes": conditioned},
        "fifth_class_old_sum_feasibility": {
            "current_reference_sum": sum(known5), "old_sum": old_summary[4]["sum"],
            "any_existing_numeric_total_subset_can_match_old_sum": subset_possible,
            "note": "只检查当前数值成绩子集；未录成绩或未同步改分不在本检验中。"},
        "missing_base": all_stats["missing"], "missing_bonus_with_base": all_stats["missing_bonus"],
        "all_missing_bonus": [source_ref(s) for s in students if blank(s["bonus"])],
        "explicit_zero_score_cells": zeros,
        "summary_in_student_total_position": [source_ref(s) | {"total_cell": s["source_cells"]["total"], "cached_value": s["raw_total"]} for s in students if s["total_is_summary"]],
        "duplicate_names_within_class": within, "duplicate_names_across_classes": across,
        "name_whitespace_or_control_characters": odd_names,
        "numeric_scores_stored_as_text": numeric_text, "other_nonnumeric_score_values": other_values,
        "invalid_sex_values": [source_ref(s) | {"sex": s["sex"]} for s in students if s["sex"] not in ("男", "女")],
        "formula_error_cells": [{"sheet": s["name"], "cell": co, "value": c["value"]} for s in book.sheets for co, c in s["cells"].items() if c["type"] == "e"],
    }


def write_outputs(output_dir, book, students, classes, all_stats, report):
    output_dir.mkdir(parents=True, exist_ok=True)
    analysis = {"students": students, "classes": classes, "all": all_stats}
    for filename, data in [("analysis.json", analysis), ("source_cells.json", book.sheets), ("portable_audit_report.json", report)]:
        (output_dir / filename).write_text(json.dumps(data, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    fields = ["class", "row", "name", "sex", "base", "bonus", "original_total", "total_is_summary",
              "total_with_blank_bonus_zero", "cell", "base_cell", "bonus_cell", "total_cell", "name_mark", "base_missing", "bonus_missing"]
    # UTF-8 BOM allows Excel/WPS to recognize Chinese text. Blank scores are
    # empty CSV fields, while actual zero scores are the literal character 0.
    with (output_dir / "students.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for student in students:
            row = {key: student[key] for key in fields if key in student}
            row.update({"base_cell": student["source_cells"]["base"], "bonus_cell": student["source_cells"]["bonus"],
                        "total_cell": student["source_cells"]["total"], "name_mark": student["name_colors"]["mark"]})
            writer.writerow(row)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", type=Path, required=True, help="只读输入XLSX路径")
    parser.add_argument("--output-dir", type=Path, required=True, help="写入JSON和CSV的目录")
    parser.add_argument("--strict-baseline", action="store_true", help="若不匹配历史214人/206份基础分基线，输出结果后以状态1退出")
    args = parser.parse_args(argv)
    try:
        book = Workbook(args.input)
        students = extract_students(book)
        old = get_old_summary(book)
        classes = [summarize([s for s in students if s["class"] == i], old[i - 1], i) for i in range(1, 6)]
        all_stats = summarize(students)
        report = make_report(book, students, classes, all_stats, old)
        write_outputs(args.output_dir, book, students, classes, all_stats, report)
    except (OSError, ValueError, KeyError, IndexError, ET.ParseError, zipfile.BadZipFile) as exc:
        print("审计失败：" + str(exc), file=sys.stderr)
        return 2
    summary_errors = sum(c["pass"] is False for c in report["actual_verification"]["selected_summary_cache_checks"])
    chart_errors = sum(c["pass"] is False for c in report["actual_verification"]["chart_reference_cache_checks"])
    print(json.dumps({"layout": report["input"]["layout"], "students": all_stats["names"],
        "numeric_base_scores": all_stats["base_count"], "base_sum": all_stats["base_sum"],
        "recorded_reference_total": all_stats["total_sum"], "colored_names": len(report["colored_names"]),
        "student_total_mismatches": len(all_stats["bad_total"]), "summary_cache_mismatches": summary_errors,
        "chart_cache_mismatches": chart_errors, "historical_baseline_matches": report["historical_baseline_matches"],
        "input_unchanged": report["actual_verification"]["input_unchanged_after_audit"],
        "output_dir": str(args.output_dir)}, ensure_ascii=False, indent=2))
    return 1 if args.strict_baseline and not report["historical_baseline_matches"] else 0


if __name__ == "__main__":
    sys.exit(main())

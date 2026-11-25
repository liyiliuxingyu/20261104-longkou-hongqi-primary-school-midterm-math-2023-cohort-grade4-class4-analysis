#!/usr/bin/env python3
"""Read-only validator for the restored November 2026 grade-four workbook.

Usage:
    python3 src/validate_workbook.py --input candidate.xlsx \
        --reference inputs/baseline.xlsx --report audit/validation.json

Only the Python standard library is required. The XLSX files are opened read-only;
the only output is the requested JSON report. Exit codes are 0 (passed),
1 (validation failed), and 2 (input/report error).

This is a fixed-snapshot regression validator, not an Excel calculation engine.
It independently recomputes the project statistics, compares important stored
formula results, and checks OOXML chart/metadata preservation. Deliberate future
changes to the pupil data require an explicitly revised snapshot specification.
"""

from __future__ import annotations

import argparse
import colorsys
import hashlib
import json
import math
import posixpath
import re
import statistics
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from urllib.parse import unquote
import xml.etree.ElementTree as ET


NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
EXPECTED_SHEETS = ["一年级", "二年级", "三年级", "四年级", "五年级", "四年级分析", "数据核查"]
OTHER_GRADES = ["一年级", "二年级", "三年级", "五年级"]
BLOCKS = ["ABCDE", "FGHIJ", "KLMNO", "PQRST", "UVWXY"]
ERROR_VALUES = {
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!",
    "#SPILL!", "#CALC!", "#GETTING_DATA", "#CONNECT!", "#BLOCKED!",
    "#UNKNOWN!", "#FIELD!", "#BUSY!",
}
PRINT_TAGS = ["pageMargins", "pageSetup", "printOptions", "headerFooter", "rowBreaks", "colBreaks"]
INDEXED_COLORS = (
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "800000 008000 000080 808000 800080 008080 C0C0C0 808080 "
    "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCCC "
    "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF "
    "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 "
    "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 "
    "003366 339966 003300 333300 993300 993366 333399 333333"
).split()
THEME_KEYS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def is_number(value: Any) -> bool:
    return isinstance(value, Decimal) and value.is_finite()


def is_blank(value: Any) -> bool:
    return value is None or value == ""


def same_value(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    if isinstance(a, (Decimal, int, float)) and not isinstance(a, bool) and isinstance(b, (Decimal, int, float)) and not isinstance(b, bool):
        return math.isclose(float(a), float(b), rel_tol=1e-12, abs_tol=tolerance)
    if is_blank(a) and is_blank(b):
        return True
    return type(a) is type(b) and a == b


def json_default(value: Any) -> Any:
    if isinstance(value, Decimal):
        if not value.is_finite():
            return str(value)
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, tuple):
        return list(value)
    raise TypeError(f"Cannot serialize {type(value).__name__}")


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def col_number(label: str) -> int:
    result = 0
    for char in label.upper():
        result = result * 26 + ord(char) - 64
    return result


def col_label(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def address_key(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", address.upper())
    if not match:
        raise ValueError(f"Invalid cell address: {address}")
    return int(match[2]), col_number(match[1])


def rich_text(element: ET.Element | None) -> str:
    """Read plain/rich text without including phonetic annotation text."""
    if element is None:
        return ""
    pieces = []
    for child in element:
        if local_name(child.tag) == "t":
            pieces.append(child.text or "")
        elif local_name(child.tag) == "r":
            pieces.extend(t.text or "" for t in child.findall("s:t", NS))
    return "".join(pieces)


@dataclass
class Cell:
    address: str
    value: Any = None
    cell_type: str | None = None
    formula: str | None = None
    formula_attributes: dict[str, str] | None = None
    style: int | None = None
    has_value_element: bool = False

    def payload(self) -> dict[str, Any]:
        kind = "number" if is_number(self.value) else "boolean" if isinstance(self.value, bool) else "blank" if is_blank(self.value) else "date" if self.cell_type == "d" else "error" if self.cell_type == "e" else "text"
        return {
            "kind": kind,
            "value": None if kind == "blank" else self.value,
            "formula": self.formula,
            "formula_attributes": self.formula_attributes,
        }


class XlsxReader:
    def __init__(self, path: Path):
        self.path = path
        self.zip = zipfile.ZipFile(path, "r")
        self.parts = set(self.zip.namelist())
        if len(self.parts) != len(self.zip.namelist()):
            raise ValueError("The XLSX archive has duplicate part names.")
        bad_member = self.zip.testzip()
        if bad_member:
            raise ValueError(f"ZIP CRC failure: {bad_member}")
        root_rel = self.relationships("")
        offices = [v for v in root_rel.values() if v["type"].endswith("/officeDocument") and not v["external"]]
        if len(offices) != 1:
            raise ValueError("Expected one internal workbook relationship.")
        self.workbook_part = offices[0]["target"]
        self.workbook = self.xml(self.workbook_part)
        self.workbook_rels = self.relationships(self.workbook_part)
        shared = self.find_relationship_part("/sharedStrings")
        self.strings = [rich_text(item) for item in self.xml(shared).findall("s:si", NS)] if shared else []
        style_part = self.find_relationship_part("/styles")
        self.styles = self.xml(style_part) if style_part else ET.Element("styles")
        fonts = self.styles.find("s:fonts", NS)
        cell_xfs = self.styles.find("s:cellXfs", NS)
        style_xfs = self.styles.find("s:cellStyleXfs", NS)
        self.fonts = list(fonts) if fonts is not None else []
        self.cell_xfs = list(cell_xfs) if cell_xfs is not None else []
        self.style_xfs = list(style_xfs) if style_xfs is not None else []
        self.theme_colors = self.read_theme_colors()
        self.sheets: dict[str, dict[str, Any]] = {}
        self.sheet_names: list[str] = []
        for index, sheet in enumerate(self.workbook.findall("s:sheets/s:sheet", NS)):
            name = sheet.get("name", "")
            rel = self.workbook_rels[sheet.get(f"{{{NS['r']}}}id", "")]
            if rel["external"] or name in self.sheets:
                raise ValueError(f"Invalid worksheet relationship/name: {name}")
            part = rel["target"]
            root = self.xml(part)
            cells = {}
            rows = {}
            for row in root.findall("s:sheetData/s:row", NS):
                rows[int(row.get("r", "0"))] = row.attrib
                for elem in row.findall("s:c", NS):
                    cell = self.parse_cell(elem)
                    if cell.address in cells:
                        raise ValueError(f"Duplicate cell: {name}!{cell.address}")
                    cells[cell.address] = cell
            self.sheet_names.append(name)
            self.sheets[name] = {"part": part, "root": root, "cells": cells, "rows": rows, "index": index}

    def close(self) -> None:
        self.zip.close()

    def xml(self, part: str) -> ET.Element:
        return ET.fromstring(self.zip.read(part))

    def relationships(self, owner: str) -> dict[str, dict[str, Any]]:
        part = posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels") if owner else "_rels/.rels"
        if part not in self.parts:
            return {}
        result = {}
        for node in self.xml(part):
            external = node.get("TargetMode") == "External"
            target = node.get("Target", "")
            if not external:
                target = unquote(target)
                target = posixpath.normpath(target.lstrip("/")) if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
            result[node.get("Id", "")] = {"target": target, "type": node.get("Type", ""), "external": external}
        return result

    def find_relationship_part(self, suffix: str) -> str | None:
        return next((x["target"] for x in self.workbook_rels.values() if x["type"].endswith(suffix) and not x["external"]), None)

    def parse_cell(self, elem: ET.Element) -> Cell:
        address = elem.get("r", "").upper()
        address_key(address)
        cell_type = elem.get("t")
        value_node = elem.find("s:v", NS)
        raw = value_node.text if value_node is not None else None
        formula_node = elem.find("s:f", NS)
        if cell_type == "inlineStr":
            value = rich_text(elem.find("s:is", NS))
        elif cell_type == "s" and raw is not None:
            value = self.strings[int(raw)]
        elif cell_type == "b" and raw is not None:
            if raw not in ("0", "1"):
                raise ValueError(f"Invalid Boolean value at {address}: {raw}")
            value = raw == "1"
        elif cell_type in ("str", "e", "d"):
            value = raw or ""
        elif raw is not None:
            value = Decimal(raw)
            if not value.is_finite():
                raise ValueError(f"Non-finite numeric value at {address}: {raw}")
        else:
            value = None
        return Cell(address, value, cell_type, formula_node.text or "" if formula_node is not None else None,
                    dict(formula_node.attrib) if formula_node is not None else None,
                    int(elem.get("s")) if elem.get("s") is not None else None, value_node is not None)

    def cell(self, sheet: str, address: str) -> Cell:
        return self.sheets[sheet]["cells"].get(address.upper(), Cell(address.upper()))

    def value(self, sheet: str, address: str) -> Any:
        return self.cell(sheet, address).value

    def nonempty(self, sheet: str) -> dict[str, Cell]:
        return {a: c for a, c in self.sheets[sheet]["cells"].items() if not is_blank(c.value) or c.formula is not None}

    def effective_font(self, sheet: str, address: str) -> ET.Element:
        cell = self.cell(sheet, address)
        style = cell.style
        row_number, column_number = address_key(address)
        if style is None:
            row = self.sheets[sheet]["rows"].get(row_number, {})
            if row.get("customFormat") in ("1", "true") and "s" in row:
                style = int(row["s"])
            else:
                for col in self.sheets[sheet]["root"].findall("s:cols/s:col", NS):
                    if int(col.get("min", "0")) <= column_number <= int(col.get("max", "0")) and "style" in col.attrib:
                        style = int(col.get("style", "0"))
                        break
        style = 0 if style is None else style
        if not self.cell_xfs:
            return ET.Element("font")
        xf = self.cell_xfs[style]
        base_index = int(xf.get("xfId", "0"))
        inherited = self.style_xfs[base_index].get("fontId", "0") if base_index < len(self.style_xfs) else "0"
        font_id = int(xf.get("fontId", inherited))
        return self.fonts[font_id] if self.fonts else ET.Element("font")

    def font_metadata(self, sheet: str, address: str) -> dict[str, Any]:
        font = self.effective_font(sheet, address)
        name = font.find("s:name", NS)
        charset = font.find("s:charset", NS)
        return {"name": name.get("val") if name is not None else None,
                "charset": charset.get("val") if charset is not None else None}

    def read_theme_colors(self) -> dict[str, str]:
        part = self.find_relationship_part("/theme")
        if not part:
            return {}
        scheme = self.xml(part).find("a:themeElements/a:clrScheme", NS)
        result = {}
        if scheme is not None:
            for node in scheme:
                if len(node):
                    result[local_name(node.tag)] = node[0].get("lastClr") or node[0].get("val", "")
        return result

    def font_color(self, sheet: str, address: str) -> dict[str, Any]:
        color = self.effective_font(sheet, address).find("s:color", NS)
        raw = dict(color.attrib) if color is not None else {}
        rgb = None
        if "rgb" in raw:
            rgb = raw["rgb"][-6:]
        elif "theme" in raw:
            index = int(raw["theme"])
            rgb = self.theme_colors.get(THEME_KEYS[index]) if 0 <= index < len(THEME_KEYS) else None
        elif "indexed" in raw:
            index = int(raw["indexed"])
            rgb = INDEXED_COLORS[index] if 0 <= index < len(INDEXED_COLORS) else "000000" if index == 64 else "FFFFFF" if index == 65 else None
        elif not raw or raw.get("auto") in ("1", "true"):
            rgb = "000000"
        if rgb and re.fullmatch(r"[0-9a-fA-F]{6}", rgb):
            tint = float(raw.get("tint", "0"))
            if tint:
                red, green, blue = (int(rgb[i:i+2], 16) / 255 for i in (0, 2, 4))
                hue, luminance, saturation = colorsys.rgb_to_hls(red, green, blue)
                luminance = luminance * (1 + tint) if tint < 0 else luminance * (1 - tint) + tint
                rgb = "".join(f"{round(max(0, min(1, x)) * 255):02X}" for x in colorsys.hls_to_rgb(hue, luminance, saturation))
            return {"resolved_rgb": rgb.upper(), "raw": raw}
        return {"resolved_rgb": None, "raw": raw}

    def print_signature(self, sheet: str) -> dict[str, Any]:
        record = self.sheets[sheet]
        rels = self.relationships(record["part"])

        def canonical(node: ET.Element | None) -> Any:
            if node is None:
                return None
            attrs = {}
            for key, value in sorted(node.attrib.items()):
                if key == f"{{{NS['r']}}}id":
                    rel = rels.get(value)
                    attrs["relationship"] = {"type": rel["type"], "external": rel["external"],
                        "target_sha256": hashlib.sha256(self.zip.read(rel["target"])).hexdigest() if not rel["external"] else None} if rel else {"missing": value}
                else:
                    try:
                        attrs[local_name(key)] = str(Decimal(value).normalize())
                    except InvalidOperation:
                        attrs[local_name(key)] = value
            return {"tag": local_name(node.tag), "attributes": attrs,
                    "text": node.text if node.text and node.text.strip() else None,
                    "children": [canonical(child) for child in node]}

        result = {tag: canonical(record["root"].find("s:" + tag, NS)) for tag in PRINT_TAGS}
        result["pageSetUpPr"] = canonical(record["root"].find("s:sheetPr/s:pageSetUpPr", NS))
        names = []
        for node in self.workbook.findall("s:definedNames/s:definedName", NS):
            if node.get("name") not in ("_xlnm.Print_Area", "_xlnm.Print_Titles"):
                continue
            local_id = node.get("localSheetId")
            if local_id is None or int(local_id) == record["index"]:
                names.append({"name": node.get("name"), "scope": "global" if local_id is None else sheet,
                              "value": node.text or "", "attributes": {k: v for k, v in node.attrib.items() if k not in ("name", "localSheetId")}})
        result["defined_print_names"] = sorted(names, key=lambda x: (x["name"], x["scope"]))
        return result

    def chart_parts(self) -> list[str]:
        root = self.xml("[Content_Types].xml")
        return sorted(node.get("PartName", "").lstrip("/") for node in root
                      if node.get("ContentType", "").endswith("drawingml.chart+xml"))


class Checks:
    def __init__(self):
        self.items: list[dict[str, Any]] = []

    def add(self, check_id: str, description: str, condition: bool | None, **details: Any) -> None:
        self.items.append({"id": check_id, "description": description,
                           "status": "skip" if condition is None else "pass" if condition else "fail", **details})

    def differences(self, check_id: str, description: str, differences: list[Any], **details: Any) -> None:
        self.add(check_id, description, not differences, difference_count=len(differences), differences=differences[:50], **details)


def pupils(book: XlsxReader) -> list[dict[str, Any]]:
    records = []
    for class_number, columns in enumerate(BLOCKS, 1):
        for row in range(2, 46):
            values = [book.value("四年级", f"{col}{row}") for col in columns]
            if is_blank(values[0]):
                continue
            records.append({"class": class_number, "row": row, "name_address": f"{columns[0]}{row}",
                            "name": values[0], "sex": values[1], "base": values[2], "bonus": values[3],
                            "total": values[4], "columns": columns})
    return records


def statistics_for(records: list[dict[str, Any]]) -> dict[str, Any]:
    base = [r["base"] for r in records if is_number(r["base"])]
    totals = [r["base"] + (r["bonus"] if is_number(r["bonus"]) else Decimal(0)) for r in records if is_number(r["base"])]
    count = len(base)
    base_sum = sum(base, Decimal(0))
    total_sum = sum(totals, Decimal(0))
    return {
        "names": len(records), "base_count": count,
        "both_blank": sum(is_blank(r["base"]) and is_blank(r["bonus"]) for r in records),
        "base_missing": len(records) - count,
        "bonus_only_blank": sum(is_number(r["base"]) and is_blank(r["bonus"]) for r in records),
        "complete": sum(is_number(r["base"]) and is_number(r["bonus"]) for r in records),
        "base_sum": base_sum, "total_sum": total_sum,
        "base_mean": base_sum / count if count else None,
        "base_median": statistics.median(base) if count else None,
        "base_sd": statistics.pstdev(float(x) for x in base) if count else None,
        "total_mean": total_sum / count if count else None,
        "bands": [sum(x < 60 for x in base), sum(60 <= x < 70 for x in base), sum(70 <= x < 80 for x in base), sum(80 <= x < 90 for x in base), sum(x >= 90 for x in base)],
        "ge60": sum(x >= 60 for x in base) / count if count else None,
        "ge80": sum(x >= 80 for x in base) / count if count else None,
        "ge90": sum(x >= 90 for x in base) / count if count else None,
        "male": sum(r["sex"] == "男" for r in records), "female": sum(r["sex"] == "女" for r in records),
    }


def rounded(value: Any) -> Decimal | None:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if value is not None else None


def check_snapshot(book: XlsxReader, checks: Checks, report: dict[str, Any]) -> None:
    checks.add("sheets.expected", "7个工作表的名称和顺序与当前快照一致", book.sheet_names == EXPECTED_SHEETS,
               expected=EXPECTED_SHEETS, observed=book.sheet_names)
    if "四年级" not in book.sheets:
        return
    records = pupils(book)
    classes = [statistics_for([r for r in records if r["class"] == i]) for i in range(1, 6)]
    overall = statistics_for(records)
    report["statistics"] = {"classes": [{"class": i, **s} for i, s in enumerate(classes, 1)], "year": overall}
    expected = {"names": 214, "base_count": 206, "both_blank": 8, "bonus_only_blank": 7, "complete": 199, "total_sum": 17053, "base_sum": 15948}
    checks.add("pupils.snapshot", "从学生输入独立核算人数、空白状态和总分", all(same_value(overall[k], v) for k, v in expected.items()), expected=expected, observed={k: overall[k] for k in expected})
    expected_class = {"names": [44, 44, 42, 42, 42], "base_count": [42, 42, 40, 42, 40], "both_blank": [2, 2, 2, 0, 2], "bonus_only_blank": [0, 4, 2, 1, 0], "complete": [42, 38, 38, 41, 40], "total_sum": [3729, 3338, 3303, 3534, 3149]}
    observed_class = {key: [c[key] for c in classes] for key in expected_class}
    checks.add("pupils.classes", "五个班分别核算，防止跨班错位", observed_class == expected_class, expected=expected_class, observed=observed_class)
    input_issues, total_issues = [], []
    for rec in records:
        loc = rec["name_address"]
        if not isinstance(rec["name"], str) or rec["name"] != rec["name"].strip() or rec["sex"] not in ("男", "女"):
            input_issues.append({"cell": loc, "issue": "姓名/性别格式异常"})
        for key in ("base", "bonus"):
            if not is_blank(rec[key]) and not is_number(rec[key]):
                input_issues.append({"cell": loc, "field": key, "value": rec[key], "issue": "非空分数不是数字"})
        if not is_number(rec["base"]) and not is_blank(rec["bonus"]):
            input_issues.append({"cell": loc, "issue": "仅附加分已录，超出本快照口径"})
        expected_total = rec["base"] + (rec["bonus"] if is_number(rec["bonus"]) else Decimal(0)) if is_number(rec["base"]) else None
        if not same_value(rec["total"], expected_total):
            total_issues.append({"cell": f"{rec['columns'][4]}{rec['row']}", "expected": expected_total, "observed": rec["total"]})
    checks.differences("pupils.input_types", "姓名/性别完整，分数为数字或真正空白", input_issues, record_count=len(records))
    checks.differences("pupils.total_cache", "214个个人总分缓存与基础分及现录附加分一致，8处缺基础分保持空白", total_issues, record_count=len(records), blank_bonus_policy="仅复核现录参考总分；附加分空白仍待核实，不等同已确认0分")
    keys = [("四年级", "H43", 0), ("四年级", "J43", 0), ("四年级", "M9", 0), ("四年级", "N9", 0), ("四年级", "O9", 0), ("四年级", "J45", None), ("四年级", "J47", 3338), ("四年级分析", "L11", Decimal("78.73")), ("数据核查", "F11", 3132), ("数据核查", "F13", -3)]
    key_details = []
    for sheet, address, expected_value in keys:
        observed = book.value(sheet, address) if sheet in book.sheets else "<missing sheet>"
        key_details.append({"cell": f"{sheet}!{address}", "expected": expected_value, "observed": observed, "matches": same_value(observed, expected_value)})
    checks.add("cells.regressions", "真实0分、空白学生行、移出合计、两位小数及5班差异保持正确", all(x["matches"] for x in key_details), cells=key_details)
    row_mapping = {50: "names", 51: "base_count", 52: "base_missing", 53: "bonus_only_blank", 54: "base_sum", 55: "base_mean", 56: "base_median", 57: "base_sd", 63: "ge60", 64: "ge90", 65: "total_sum", 66: "total_mean", 67: "ge80", 68: "complete", 69: "male", 70: "female"}
    summary_issues = []
    for col, stat in zip("BCDEFG", classes + [overall]):
        for row, key in row_mapping.items():
            cell = book.cell("四年级", f"{col}{row}")
            if not same_value(cell.value, stat[key]) or cell.formula is None:
                summary_issues.append({"cell": f"四年级!{col}{row}", "expected": stat[key], "observed": cell.value, "has_formula": cell.formula is not None})
        for row, expected_value in zip(range(58, 63), stat["bands"]):
            cell = book.cell("四年级", f"{col}{row}")
            if not same_value(cell.value, expected_value) or cell.formula is None:
                summary_issues.append({"cell": f"四年级!{col}{row}", "expected": expected_value, "observed": cell.value, "has_formula": cell.formula is not None})
    checks.differences("summary.independent", "四年级统计区与Python独立计算相符，并保留公式", summary_issues, checked_cells=126)
    if "四年级分析" in book.sheets:
        output_issues = []
        output_mapping = {"B": "names", "C": "base_count", "D": "base_missing", "E": "bonus_only_blank", "F": "base_mean", "G": "base_median", "H": "base_sd", "I": "ge60", "J": "ge90", "K": "total_sum", "L": "total_mean"}
        for row, stat in zip(range(7, 13), classes + [overall]):
            for col, key in output_mapping.items():
                expected_value = rounded(stat[key]) if col in "FGHL" else stat[key]
                cell = book.cell("四年级分析", f"{col}{row}")
                if not same_value(cell.value, expected_value) or cell.formula is None:
                    output_issues.append({"cell": f"四年级分析!{col}{row}", "expected": expected_value, "observed": cell.value, "has_formula": cell.formula is not None})
        checks.differences("analysis.independent", "分析主表与独立计算及十进制四舍五入相符", output_issues, checked_cells=66)


def check_formula_errors(book: XlsxReader, checks: Checks) -> None:
    errors, missing_numeric_caches = [], []
    formula_count = 0
    blank_string_caches = 0
    for sheet in book.sheet_names:
        for address, cell in book.sheets[sheet]["cells"].items():
            if cell.formula is not None:
                formula_count += 1
                if "#REF!" in cell.formula:
                    errors.append({"cell": f"{sheet}!{address}", "formula": cell.formula, "issue": "formula contains #REF!"})
                if is_blank(cell.value):
                    if cell.cell_type in ("str", "inlineStr"):
                        blank_string_caches += 1
                    else:
                        missing_numeric_caches.append(f"{sheet}!{address}")
            if cell.cell_type == "e" or cell.formula is not None and isinstance(cell.value, str) and cell.value in ERROR_VALUES:
                errors.append({"cell": f"{sheet}!{address}", "value": cell.value, "issue": "stored error value"})
    checks.differences("formulas.error_scan", "工作簿中无公式错误缓存或#REF!引用", errors, formula_count=formula_count,
                       scope="检查存储的错误值和引用；不运行Excel公式引擎")
    checks.differences("formulas.cache_presence", "数值公式缓存没有缺失，合法空字符串缓存另行计数", missing_numeric_caches,
                       formula_count=formula_count, blank_string_caches=blank_string_caches)


def parse_range(formula: str) -> tuple[str, list[str]]:
    match = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?", formula.strip().lstrip("="))
    if not match:
        raise ValueError(f"Unsupported chart range reference: {formula}")
    sheet = match[1].replace("''", "'") if match[1] is not None else match[2]
    left, top = col_number(match[3]), int(match[4])
    right, bottom = col_number(match[5] or match[3]), int(match[6] or match[4])
    if right < left or bottom < top or (right != left and bottom != top) or (right - left + 1) * (bottom - top + 1) > 1000:
        raise ValueError(f"Expected a bounded one-dimensional chart range: {formula}")
    return sheet, [f"{col_label(col)}{row}" for row in range(top, bottom + 1) for col in range(left, right + 1)]


def check_charts(book: XlsxReader, checks: Checks, report: dict[str, Any]) -> None:
    parts = book.chart_parts()
    checks.add("charts.count", "存在两个原生图表部件", len(parts) == 2, expected=2, observed=len(parts), parts=parts)
    chart_details = []
    roles = []
    for chart_index, part in enumerate(parts, 1):
        root = book.xml(part)
        title = "".join(n.text or "" for n in root.findall(".//c:title//a:t", NS))
        series = root.findall(".//c:ser", NS)
        series_issues = []
        all_value_refs = []
        detail = {"part": part, "title": title, "series": []}
        for series_index, ser in enumerate(series, 1):
            sd = {"series": series_index, "references": []}
            for role, tag in (("category", "cat"), ("value", "val")):
                parent = ser.find(f"c:{tag}", NS)
                refs = list(parent) if parent is not None else []
                refs = [ref for ref in refs if local_name(ref.tag) in ("strRef", "numRef")]
                if len(refs) != 1:
                    series_issues.append({"series": series_index, "role": role, "issue": "需要一个单元格引用及缓存"})
                    continue
                ref = refs[0]
                formula = ref.findtext("c:f", default="", namespaces=NS)
                cache = ref.find("c:strCache" if local_name(ref.tag) == "strRef" else "c:numCache", NS)
                try:
                    sheet, addresses = parse_range(formula)
                    if sheet not in book.sheets:
                        raise ValueError(f"Missing referenced sheet: {sheet}")
                    points = [] if cache is None else cache.findall("c:pt", NS)
                    count_node = cache.find("c:ptCount", NS) if cache is not None else None
                    count = int(count_node.get("val", "-1")) if count_node is not None else None
                    indices = [int(point.get("idx", "-1")) for point in points]
                    if count != 5 or len(addresses) != 5 or len(points) != 5 or sorted(indices) != list(range(5)):
                        raise ValueError(f"Expected five points with indices 0..4: declared={count}, range={len(addresses)}, points={indices}")
                    cached = {int(p.get("idx", "-1")): p.findtext("c:v", default="", namespaces=NS) for p in points}
                    source_values = [book.value(sheet, address) for address in addresses]
                    parsed = [Decimal(cached[i]) if local_name(ref.tag) == "numRef" else cached[i] for i in range(5)]
                    if any(not same_value(a, b) for a, b in zip(parsed, source_values)):
                        series_issues.append({"series": series_index, "role": role, "formula": formula, "issue": "缓存与引用单元格不同", "cache": parsed, "cells": source_values})
                    sd["references"].append({"role": role, "formula": formula, "point_count": count, "cache": parsed, "cells": source_values})
                    if role == "value":
                        if any(not is_number(v) for v in parsed):
                            raise ValueError("Numeric chart cache contains a non-finite value.")
                        all_value_refs.append((sheet, ":".join([addresses[0], addresses[-1]])))
                except (ValueError, KeyError, InvalidOperation) as exc:
                    series_issues.append({"series": series_index, "role": role, "formula": formula, "issue": str(exc)})
            detail["series"].append(sd)
        checks.differences(f"charts.{chart_index}.references", "各序列分类与数值缓存均有5点且与引用单元格对应", series_issues, part=part, series_count=len(series))
        expected_mean = [("四年级分析", "F7:F11"), ("四年级分析", "G7:G11")]
        expected_distribution = [("四年级分析", "J18:J22")]
        role = "class_comparison" if sorted(all_value_refs) == sorted(expected_mean) else "year_distribution" if all_value_refs == expected_distribution else "unrecognized"
        roles.append(role)
        expected_max = 100 if role == "class_comparison" else 80 if role == "year_distribution" else None
        axes = root.findall(".//c:valAx", NS)
        axis_details = []
        for axis in axes:
            lower = axis.find("c:scaling/c:min", NS)
            upper = axis.find("c:scaling/c:max", NS)
            axis_details.append({"min": Decimal(lower.get("val")) if lower is not None else None,
                                 "max": Decimal(upper.get("val")) if upper is not None else None})
        checks.add(f"charts.{chart_index}.axis", "数值轴从0开始，均值图上限100、分布图上限80", len(axes) == 1 and expected_max is not None and same_value(axis_details[0]["min"], 0) and same_value(axis_details[0]["max"], expected_max), part=part, role=role, expected_max=expected_max, observed=axis_details)
        detail["role"] = role
        detail["axes"] = axis_details
        chart_details.append(detail)
    checks.add("charts.roles", "两个图表分别展示班级均值/中位数和全年级分布", sorted(roles) == ["class_comparison", "year_distribution"], observed=roles)
    anchors, anchor_issues = [], []
    for sheet in book.sheet_names:
        sh = book.sheets[sheet]
        sh_rels = book.relationships(sh["part"])
        for drawing in sh["root"].findall("s:drawing", NS):
            rel = sh_rels.get(drawing.get(f"{{{NS['r']}}}id", ""))
            if rel is None or rel["external"]:
                anchor_issues.append({"sheet": sheet, "issue": "绘图关系缺失或为外部关系"})
                continue
            drawing_part = rel["target"]
            drawing_rels = book.relationships(drawing_part)
            for anchor in book.xml(drawing_part):
                chart_nodes = anchor.findall(".//c:chart", NS)
                if not chart_nodes:
                    continue
                if local_name(anchor.tag) != "twoCellAnchor" or len(chart_nodes) != 1:
                    anchor_issues.append({"sheet": sheet, "issue": "当前快照需要每图一个双单元格锚点", "anchor_type": local_name(anchor.tag)})
                    continue
                def marker(which: str) -> dict[str, int]:
                    elem = anchor.find("xdr:" + which, NS)
                    if elem is None:
                        raise ValueError("Anchor marker missing.")
                    return {key: int(elem.findtext("xdr:" + key, namespaces=NS) or "0") for key in ("col", "colOff", "row", "rowOff")}
                start, end = marker("from"), marker("to")
                chart_rel = drawing_rels.get(chart_nodes[0].get(f"{{{NS['r']}}}id", ""))
                target = chart_rel["target"] if chart_rel and not chart_rel["external"] else None
                if target not in parts or any(v < 0 for v in list(start.values()) + list(end.values())) or (start["col"], start["colOff"]) >= (end["col"], end["colOff"]) or (start["row"], start["rowOff"]) >= (end["row"], end["rowOff"]):
                    anchor_issues.append({"sheet": sheet, "issue": "图表引用或锚点范围无效", "chart": target})
                anchors.append({"sheet": sheet, "drawing": drawing_part, "chart": target, "from": start, "to": end})
    def overlaps(a: dict[str, Any], b: dict[str, Any]) -> bool:
        if a["sheet"] != b["sheet"]:
            return False
        return not ((a["to"]["col"], a["to"]["colOff"]) <= (b["from"]["col"], b["from"]["colOff"]) or
                    (b["to"]["col"], b["to"]["colOff"]) <= (a["from"]["col"], a["from"]["colOff"]) or
                    (a["to"]["row"], a["to"]["rowOff"]) <= (b["from"]["row"], b["from"]["rowOff"]) or
                    (b["to"]["row"], b["to"]["rowOff"]) <= (a["from"]["row"], a["from"]["rowOff"]))
    if len(anchors) != 2 or len({a["chart"] for a in anchors}) != 2 or any(a["sheet"] != "四年级分析" for a in anchors):
        anchor_issues.append({"issue": "需要两个独立图表锚点且均位于四年级分析", "count": len(anchors)})
    if len(anchors) == 2 and overlaps(anchors[0], anchors[1]):
        anchor_issues.append({"issue": "两个图表锚点重叠"})
    checks.differences("charts.anchors", "两个图表锚点均有效、各引用一个图表且互不重叠", anchor_issues, anchors=anchors,
                       scope="按双单元格锚点坐标检查，不替代Excel/WPS视觉检查")
    report["charts"] = chart_details


def compare_reference(book: XlsxReader, reference: XlsxReader | None, checks: Checks) -> None:
    if reference is None:
        checks.add("reference.not_supplied", "未传reference，未验证相对基线的输入、姓名颜色、打印及charset保留情况", None)
        return
    differences, compared = [], 0
    for sheet in OTHER_GRADES:
        if sheet not in book.sheets or sheet not in reference.sheets:
            differences.append({"sheet": sheet, "issue": "候选或基线中缺少工作表"})
            continue
        current, original = book.nonempty(sheet), reference.nonempty(sheet)
        for address in sorted(set(current) | set(original), key=address_key):
            left = current.get(address, Cell(address)).payload()
            right = original.get(address, Cell(address)).payload()
            compared += 1
            if left != right:
                differences.append({"cell": f"{sheet}!{address}", "expected": right, "observed": left})
    checks.differences("reference.other_grades", "其他四个年级的全部非空值及公式与所给基线一致", differences, checked_cells=compared)
    input_diffs, color_diffs, charset_diffs, print_diffs = [], [], [], []
    input_count, name_count, charset_count = 0, 0, 0
    if "四年级" in book.sheets and "四年级" in reference.sheets:
        for rec in pupils(reference):
            row, columns = rec["row"], rec["columns"]
            for col in columns[:4]:
                address = f"{col}{row}"
                left, right = book.cell("四年级", address).payload(), reference.cell("四年级", address).payload()
                input_count += 1
                if left != right:
                    input_diffs.append({"cell": f"四年级!{address}", "expected": right, "observed": left})
            address = rec["name_address"]
            left, right = book.font_color("四年级", address), reference.font_color("四年级", address)
            name_count += 1
            if (left["resolved_rgb"] != right["resolved_rgb"] if left["resolved_rgb"] is not None and right["resolved_rgb"] is not None else left != right):
                color_diffs.append({"cell": f"四年级!{address}", "expected": right, "observed": left})
    else:
        input_diffs.append({"issue": "候选或基线缺少四年级"})
        color_diffs.append({"issue": "候选或基线缺少四年级"})
    for sheet in reference.sheet_names:
        if sheet not in book.sheets:
            print_diffs.append({"sheet": sheet, "issue": "候选缺少工作表"})
            charset_diffs.append({"sheet": sheet, "issue": "候选缺少工作表"})
            continue
        left, right = book.print_signature(sheet), reference.print_signature(sheet)
        if left != right:
            print_diffs.append({"sheet": sheet, "changed_fields": [k for k in right if left[k] != right[k]], "expected": right, "observed": left})
        for address in reference.nonempty(sheet):
            charset_count += 1
            left, right = book.font_metadata(sheet, address), reference.font_metadata(sheet, address)
            if left != right:
                charset_diffs.append({"cell": f"{sheet}!{address}", "expected": right, "observed": left})
    checks.differences("reference.student_inputs", "四年级所有学生的姓名、性别、基础分、附加分与基线一致", input_diffs, checked_cells=input_count)
    checks.differences("reference.name_colors", "四年级姓名原字体颜色保留（比较解析颜色，不依赖样式ID）", color_diffs, checked_names=name_count)
    checks.differences("reference.print_settings", "各基线工作表打印属性及打印区域/标题保留", print_diffs,
                       fields=PRINT_TAGS + ["sheetPr/pageSetUpPr", "_xlnm.Print_Area", "_xlnm.Print_Titles"])
    checks.differences("reference.font_charset", "基线非空单元格的实际字体名称和charset保留（不强行补设未声明的charset）", charset_diffs, checked_cells=charset_count)


def make_report(input_path: Path, reference_path: Path | None) -> dict[str, Any]:
    report: dict[str, Any] = {
        "schema_version": 1,
        "project": "2026.11四年级期中成绩七页恢复快照",
        "validated_at_utc": datetime.now(timezone.utc).isoformat(),
        "status": "error", "ok": False,
        "input": {"path": str(input_path)},
        "reference": {"path": str(reference_path)} if reference_path else None,
        "limitations": [
            "验证器只读XLSX，独立核算本项目统计并检查存储缓存；不执行Excel/WPS公式引擎。",
            "图表缓存和锚点通过不等于已经在Excel/WPS中实际打开或完成视觉验收。",
            "reference检查只证明相对于实际提供基线的保留情况；不能证明与尚未恢复的最初五页原文件相同。",
            "这是固定快照回归校验。后续合法更改学生或成绩后，应明确更新期望值，不能把旧快照通过作为新数据验收。",
            "空基础分不作0分。7处附加分空白、姓名颜色含义和学校正式统计范围仍需人工确认。",
        ],
    }
    checks = Checks()
    book = reference = None
    try:
        report["input"].update(sha256=sha256(input_path), size_bytes=input_path.stat().st_size)
        book = XlsxReader(input_path)
        if reference_path:
            report["reference"].update(sha256=sha256(reference_path), size_bytes=reference_path.stat().st_size)
            reference = XlsxReader(reference_path)
        check_snapshot(book, checks, report)
        check_formula_errors(book, checks)
        check_charts(book, checks, report)
        compare_reference(book, reference, checks)
        report["observations"] = {
            "sheet_names": book.sheet_names,
            "font_definition_charset_counts": dict(Counter((font.find("s:charset", NS).get("val") if font.find("s:charset", NS) is not None else "not_declared") for font in book.fonts)),
            "print_settings": {sheet: book.print_signature(sheet) for sheet in book.sheet_names},
        }
        report["checks"] = checks.items
        report["summary"] = dict(Counter(item["status"] for item in checks.items))
        report["ok"] = not any(item["status"] == "fail" for item in checks.items)
        report["status"] = "passed" if report["ok"] else "failed"
    except Exception as exc:
        report["error"] = {"type": type(exc).__name__, "message": str(exc)}
        report["checks"] = checks.items
        report["summary"] = dict(Counter(item["status"] for item in checks.items))
    finally:
        if book is not None:
            book.close()
        if reference is not None:
            reference.close()
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="只读验证本项目7页成绩工作簿；仅输出JSON报告。")
    parser.add_argument("--input", required=True, type=Path, help="待验证的XLSX文件")
    parser.add_argument("--reference", type=Path, help="可选：用于保留性对比的基线XLSX文件")
    parser.add_argument("--report", required=True, type=Path, help="输出JSON报告路径（必须为.json）")
    args = parser.parse_args()
    input_path = args.input.expanduser().resolve()
    reference_path = args.reference.expanduser().resolve() if args.reference else None
    report_path = args.report.expanduser().resolve()
    if report_path.suffix.lower() != ".json" or report_path in {input_path, reference_path}:
        parser.error("--report必须是独立的.json路径，不能覆盖输入或基线文件。")
    if report_path.exists() and any(p is not None and p.exists() and report_path.samefile(p) for p in (input_path, reference_path)):
        parser.error("--report与输入或基线指向同一文件，拒绝写入。")
    report = make_report(input_path, reference_path)
    try:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=json_default) + "\n", encoding="utf-8")
    except OSError as exc:
        print(json.dumps({"status": "error", "message": f"无法写入验证报告: {exc}"}, ensure_ascii=False), file=sys.stderr)
        return 2
    print(json.dumps({"status": report["status"], "ok": report["ok"], "summary": report.get("summary", {}), "report": str(report_path), **({"error": report["error"]} if "error" in report else {})}, ensure_ascii=False))
    return 0 if report["ok"] else 2 if report["status"] == "error" else 1


if __name__ == "__main__":
    raise SystemExit(main())

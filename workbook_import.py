# -*- coding: utf-8 -*-
"""Parse a test-plan workbook (同 AFE_推荐测试项目与结果记录表1.xlsx 的结构) into a project payload.

The parser is header-driven: it looks for the sheet whose header row best matches
the known column names, so column order may differ and extra columns are ignored.
Workbooks that lack a recognisable test sheet are rejected with a clear message.
"""

from __future__ import annotations

import datetime
import io
import re

TEST_FIELDS = (
    "id", "priority", "module", "title", "purpose", "stimulus", "metrics",
    "conditions", "calculation", "criteria", "instruments", "source_result",
    "source_unit", "source_status", "source_date", "source_notes",
)

MAX_FIELD_CHARS = 4000
HEADER_SCAN_ROWS = 15

TEST_ALIASES = {
    "id": ("序号", "编号", "id", "no", "no.", "序号/编号"),
    "priority": ("优先级", "优先", "priority"),
    "module": ("模块", "分类", "module"),
    "title": ("测试项目", "项目名称", "测试项", "项目", "title"),
    "purpose": ("测试目的", "目的", "purpose"),
    "stimulus": ("激励条件/设置", "测试连接/激励", "激励条件", "激励设置", "激励", "刺激条件"),
    "metrics": ("关注指标", "需记录的指标", "指标", "关注参数"),
    "conditions": ("测试条件设置", "建议测试条件", "测试条件", "条件"),
    "calculation": ("计算/分析方法", "计算/处理方法", "计算方法", "分析方法", "计算"),
    "criteria": ("判定参考/文献依据", "论文参考/建议判据", "判定参考", "判定依据", "判据", "判定标准"),
    "instruments": ("仪器", "仪器设备", "设备"),
    "source_result": ("实测结果", "结果", "测试结果"),
    "source_unit": ("单位", "unit"),
    "source_status": ("pass/fail", "状态", "判定"),
    "source_date": ("测试日期", "日期", "date"),
    "source_notes": ("备注", "notes"),
}

SUMMARY_ALIASES = {
    "模块": ("模块", "分类"),
    "建议优先记录的核心指标": ("建议优先记录的核心指标", "核心指标", "指标", "关注指标"),
    "结果": ("结果", "数值", "实测结果"),
    "单位": ("单位",),
    "目标/判据": ("目标/判据", "目标/研究", "目标", "判据", "目标值"),
    "备注": ("备注",),
}

INSTRUCTION_ALIASES = {
    "项目": ("项目", "事项", "条目", "主题"),
    "说明": ("说明", "内容", "描述"),
}


class WorkbookError(Exception):
    """Raised when the uploaded workbook cannot be understood."""


def _norm(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("\u3000", " ").replace("／", "/")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _label(text: str) -> str:
    return (text or "").strip()


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()[:10] if not isinstance(value, datetime.datetime) \
            else value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    return s[:MAX_FIELD_CHARS]


def _match_headers(row, aliases) -> dict:
    mapping = {}
    for index, cell in enumerate(row):
        key = _norm(cell)
        if not key:
            continue
        for field, names in aliases.items():
            if field in mapping.values():
                continue
            if any(key == _norm(n) for n in names):
                mapping[index] = field
                break
    return mapping


def _best_header(sheets, aliases, required) -> tuple:
    best = (0, None, None, {})
    for name, rows in sheets:
        for row_index, row in enumerate(rows):
            mapping = _match_headers(row, aliases)
            score = len(mapping) + (3 if all(req in mapping.values() for req in required) else 0)
            if score > best[0]:
                best = (score, name, row_index, mapping)
    return best


def _read_rows(ws, limit=None):
    rows = []
    for index, row in enumerate(ws.iter_rows(values_only=True)):
        if limit is not None and index >= limit:
            break
        rows.append(tuple(row))
    return rows


def parse_workbook(data: bytes, filename: str = "") -> dict:
    """Return {'source', 'tests', 'summary', 'instructions', 'sheet'} or raise WorkbookError."""
    try:
        from openpyxl import load_workbook
    except ImportError:  # pragma: no cover - deployment detail
        raise WorkbookError("服务器缺少 openpyxl，无法解析表格") from None

    try:
        book = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl raises many types
        raise WorkbookError(f"无法读取表格：{exc}") from exc

    try:
        names = list(book.sheetnames)
        if not names:
            raise WorkbookError("表格中没有工作表")
        sheets = [(name, _read_rows(book[name], HEADER_SCAN_ROWS)) for name in names]

        score, sheet_name, header_row, mapping = _best_header(sheets, TEST_ALIASES, ("title",))
        if sheet_name is None or score < 3 or "title" not in mapping.values():
            raise WorkbookError(
                "未找到测试项目表：表头需要包含「序号 / 优先级 / 模块 / 测试项目」等列。"
            )

        fields = {field: col for col, field in mapping.items()}
        tests = []
        seen = set()
        ws = book[sheet_name]
        seq = 0
        for row_index, row in enumerate(ws.iter_rows(values_only=True)):
            if row_index <= header_row:
                continue
            record = {}
            for field in TEST_FIELDS:
                col = fields.get(field)
                record[field] = _cell_text(row[col]) if col is not None and col < len(row) else ""
            if not record["id"] and not record["title"]:
                continue
            seq += 1
            if not record["id"]:
                record["id"] = f"P{seq:02d}"
            base = record["id"]
            suffix = 2
            while record["id"] in seen:
                record["id"] = f"{base}-{suffix}"
                suffix += 1
            seen.add(record["id"])
            tests.append(record)

        if not tests:
            raise WorkbookError("测试项目表中没有数据行")

        summary = []
        instructions = []
        for name, rows in sheets:
            if name == sheet_name:
                continue
            s_score, _, s_row, s_map = _best_header([(name, rows)], SUMMARY_ALIASES, ("模块",))
            i_score, _, i_row, i_map = _best_header([(name, rows)], INSTRUCTION_ALIASES, ("项目", "说明"))
            body = _read_rows(book[name])
            if s_score >= 3 and s_row is not None and not summary:
                summary = _extract(body, s_row, s_map, SUMMARY_ALIASES)
            elif i_score >= 2 and i_row is not None and not instructions:
                instructions = _extract(body, i_row, i_map, INSTRUCTION_ALIASES)

        return {
            "source": filename,
            "sheet": sheet_name,
            "tests": tests,
            "summary": summary,
            "instructions": instructions,
        }
    finally:
        book.close()


def _extract(rows, header_row, mapping, aliases) -> list:
    out = []
    for row_index, row in enumerate(rows):
        if row_index <= header_row:
            continue
        item = {}
        empty = True
        for col, field in mapping.items():
            value = _cell_text(row[col]) if col < len(row) else ""
            item[field] = value
            if value:
                empty = False
        if empty:
            continue
        out.append({key: item.get(key, "") for key in aliases})
    return out

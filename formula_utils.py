"""月报公式结构处理工具。"""
from __future__ import annotations

import re
from openpyxl.formula import Tokenizer


SHEET_TOKEN = re.compile(r"^(?:'(?P<quoted>(?:[^']|'')+)'|(?P<bare>[^!]+))!(?P<reference>.+)$")


def replace_sheet_references(formula: str, new_sheet: str) -> str:
    """通过 openpyxl Tokenizer 修改 RANGE 标记，不拆解公式字符串参数。"""
    if not (isinstance(formula, str) and formula.startswith("=")):
        return formula
    tokenizer = Tokenizer(formula)
    changed = False
    escaped = new_sheet.replace("'", "''")
    for token in tokenizer.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        match = SHEET_TOKEN.match(token.value)
        if match:
            token.value = f"'{escaped}'!{match.group('reference')}"
            changed = True
    return "=" + "".join(token.value for token in tokenizer.items) if changed else formula


def has_function(formula: str, function_name: str) -> bool:
    if not (isinstance(formula, str) and formula.startswith("=")):
        return False
    expected = function_name.upper() + "("
    return any(token.type == "FUNC" and token.subtype == "OPEN" and token.value.upper() == expected
               for token in Tokenizer(formula).items)


def normalize_division_columns(sheet, column_indexes, start_row: int, end_row: int) -> int:
    """仅检查标题定位出的指定列，禁止遍历整张工作表。"""
    changed = 0
    for column in dict.fromkeys(column_indexes):
        for row in range(start_row, end_row):
            cell = sheet.cell(row, column)
            formula = cell.value
            if not (isinstance(formula, str) and formula.startswith("=")):
                continue
            tokens = Tokenizer(formula).items
            if not any(token.type == "OPERATOR-INFIX" and token.value == "/" for token in tokens):
                continue
            expression = _unwrap_iferror(tokens) or formula[1:]
            updated = f'=IFERROR(({expression}),"0")'
            if updated != formula:
                cell.value = updated
                changed += 1
    return changed


def normalize_all_division_formulas(sheet) -> int:
    """处理所有实际存在的公式单元格，不扫描未使用的空白矩形区域。"""
    changed = 0
    for cell in tuple(sheet._cells.values()):
        formula = cell.value
        if not (isinstance(formula, str) and formula.startswith("=")):
            continue
        tokens = Tokenizer(formula).items
        if not any(token.type == "OPERATOR-INFIX" and token.value == "/" for token in tokens):
            continue
        expression = _unwrap_iferror(tokens) or formula[1:]
        updated = f'=IFERROR(({expression}),"0")'
        if updated != formula:
            cell.value = updated
            changed += 1
    return changed


def _unwrap_iferror(tokens) -> str | None:
    meaningful = [token for token in tokens if token.type != "WSPACE"]
    if not meaningful or not (
        meaningful[0].type == "FUNC" and meaningful[0].subtype == "OPEN"
        and meaningful[0].value.upper() == "IFERROR("
    ):
        return None
    depth = 1
    result = []
    for token in meaningful[1:]:
        if token.type == "FUNC" and token.subtype == "OPEN":
            depth += 1
        elif token.type == "FUNC" and token.subtype == "CLOSE":
            depth -= 1
            if depth == 0:
                break
        if token.type == "SEP" and token.subtype == "ARG" and depth == 1:
            break
        result.append(token.value)
    return "".join(result) or None

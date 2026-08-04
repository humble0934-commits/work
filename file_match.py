"""品牌名称提取和日报/月报分组。"""
from __future__ import annotations

from datetime import date
from pathlib import Path
import re

from excel_utils import ReportError


REPORT_WORDS = re.compile(r"(?:日报|月报)$")
DATE_PATTERNS = (
    re.compile(r"(?:19|20)\d{6}"),
    re.compile(r"(?:19|20)\d{2}[年._-]\d{1,2}[月._-]\d{1,2}日?"),
    re.compile(r"\d{1,2}[月._-]\d{1,2}日?"),
)


def filtered_brand_text(path: Path | str) -> str:
    """删除报表类型、所有数字及日期标点，只保留用于匹配的品牌文字。"""
    text = Path(path).stem.strip()
    text = re.sub(r"(?:日报|月报)", "", text)
    # 先删除带数字的年月日片段，避免删数字后残留“年月日”。
    text = re.sub(r"[0-9０-９]{1,8}\s*[年月日]", "", text)
    text = re.sub(r"[0-9０-９]", "", text)
    text = re.sub(r"[\s._\-—–/\\]+", "", text)
    text = text.strip("年月日（）()[]【】")
    if not text:
        raise ReportError(f"过滤数字后无法取得品牌关键词：{Path(path).name}")
    return text


def extract_brand(path: Path | str) -> str:
    """兼容接口：返回过滤全部数字后的品牌文字。"""
    return filtered_brand_text(path)


def report_keyword(path: Path | str) -> str:
    """从报表名取得已过滤所有数字的品牌关键词。"""
    return filtered_brand_text(path)


def date_from_filename(path: Path, default_year: int) -> date | None:
    stem = path.stem
    compact = re.search(r"((?:19|20)\d{2})(\d{2})(\d{2})", stem)
    if compact:
        try:
            return date(int(compact.group(1)), int(compact.group(2)), int(compact.group(3)))
        except ValueError:
            return None
    full = re.search(r"((?:19|20)\d{2})[年._-](\d{1,2})[月._-](\d{1,2})日?", stem)
    if full:
        try:
            return date(int(full.group(1)), int(full.group(2)), int(full.group(3)))
        except ValueError:
            return None
    short = re.search(r"(?<!\d)(\d{1,2})[月._-](\d{1,2})日?", stem)
    if short:
        try:
            return date(default_year, int(short.group(1)), int(short.group(2)))
        except ValueError:
            return None
    return None


def match_brand_files(daily_files: list[Path], report_files: list[Path]):
    """按报表关键词与日报文件名双向包含进行模糊匹配。"""
    pairs = []
    for report in report_files:
        keyword = report_keyword(report)
        folded_keyword = keyword.casefold()
        matches = []
        for daily in daily_files:
            daily_brand = filtered_brand_text(daily).casefold()
            if folded_keyword in daily_brand or daily_brand in folded_keyword:
                matches.append(daily)
        if not matches:
            raise ReportError(f"未找到包含关键词：\n{keyword}\n\n的日报文件。")
        pairs.append((sorted(matches, key=lambda item: item.name), report))
    return pairs

#!/usr/bin/env python3
"""Analyze Zhipu billing Excel files by apiKey.

The input is one or more .xlsx files with these columns:
  C: 入账时间
  F: apiKey
  P: 单价单位
  S: 用量

Rows whose 单价单位 is 千token are aggregated by apiKey across all input
files. Usage during the peak window 14:00 <= time < 18:00 on weekdays
(Monday-Friday) is multiplied by 3 for billed usage, while its original
unmultiplied usage is also reported. Weekends (Saturday/Sunday) are never
treated as peak. glm-5.3-flash usage is billed at 0.5x of glm-5.3 at all
times: the discounted usage per key equals the peak-weighted total minus
half of the 5.3-flash peak-weighted usage, ceiled to an integer.

Per-call tool rows (产品名称 K 列为【web-reader】/【search-prime】应用组件,
单价单位=次, 用量=调用次数) are counted at a fixed 2,000,000 tokens per
call and added flat into billed usage (no peak multiplier, no discount).
"""

import argparse
import csv
import datetime as _datetime
import glob
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict
from decimal import ROUND_CEILING, Decimal, InvalidOperation

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

COLUMN_TIME = "C"
COLUMN_API_KEY = "F"
COLUMN_MODEL = "G"
COLUMN_PRODUCT = "K"
COLUMN_UNIT = "P"
COLUMN_USAGE = "S"
TARGET_UNIT = "千token"
FLASH_MODEL = "glm-5.3-flash"
FLASH_BILL_RATE = Decimal("0.5")
PEAK_MULTIPLIER = Decimal("3")
TOOL_PRODUCTS = ("【search-prime】", "【web-reader】")
TOOL_CALL_TOKENS = Decimal("2000000")
EXCEL_EPOCH = _datetime.datetime(1899, 12, 30)


class BillSummary(object):
    def __init__(self, api_key):
        self.api_key = api_key
        self.rows = 0
        self.total_raw_usage = Decimal("0")
        self.off_peak_usage = Decimal("0")
        self.peak_raw_usage = Decimal("0")
        self.peak_weighted_usage = Decimal("0")
        self.billable_usage = Decimal("0")
        self.discounted_usage = Decimal("0")
        self.flash_billable_usage = Decimal("0")
        self.tool_calls = Decimal("0")
        self.tool_usage = Decimal("0")

    def add(self, usage, is_peak, is_flash):
        self.rows += 1
        self.total_raw_usage += usage
        if is_peak:
            self.peak_raw_usage += usage
            self.peak_weighted_usage += usage * PEAK_MULTIPLIER
        else:
            self.off_peak_usage += usage

        weighted = usage * PEAK_MULTIPLIER if is_peak else usage
        self.billable_usage += weighted
        if is_flash:
            self.flash_billable_usage += weighted

    def add_tool(self, calls):
        """【web-reader】/【search-prime】应用组件：每次调用按固定 200万 token 计入总量。"""
        self.rows += 1
        self.tool_calls += calls
        self.tool_usage += calls * TOOL_CALL_TOKENS
        self.billable_usage += calls * TOOL_CALL_TOKENS


def fail(message):
    print("错误: {0}".format(message), file=sys.stderr)
    return 1


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def decimal_from_value(value):
    text = normalize_text(value)
    if not text:
        return None
    text = text.replace(",", "")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def decimal_to_text(value):
    value = value.normalize()
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def display_width(text):
    width = 0
    for char in str(text):
        width += 2 if unicodedata.east_asian_width(char) in ("F", "W") else 1
    return width


def pad_display(text, width, align="left"):
    text = str(text)
    padding = max(0, width - display_width(text))
    if align == "right":
        return " " * padding + text
    return text + " " * padding


def column_name(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref or "")
    return match.group(1) if match else ""


def parse_datetime_value(value):
    """解析入账时间为 datetime（保留日期，用于判断星期几）；无法解析返回 None。"""
    if value is None:
        return None
    if isinstance(value, _datetime.datetime):
        return value
    if isinstance(value, _datetime.date):
        return _datetime.datetime(value.year, value.month, value.day)

    text = normalize_text(value)
    if not text:
        return None

    numeric = decimal_from_value(text)
    if numeric is not None:
        try:
            return EXCEL_EPOCH + _datetime.timedelta(days=float(numeric))
        except (OverflowError, ValueError):
            pass

    normalized = text.replace("T", " ").replace("/", "-")
    if normalized.endswith("Z"):
        normalized = normalized[:-1]
    if "+" in normalized:
        normalized = normalized.split("+", 1)[0]

    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d",
        "%H:%M:%S",
        "%H:%M",
    )
    for fmt in formats:
        try:
            return _datetime.datetime.strptime(normalized, fmt)
        except ValueError:
            continue

    match = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", normalized)
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2))
        second = int(match.group(3) or "0")
        if 0 <= hour <= 23 and 0 <= minute <= 59 and 0 <= second <= 59:
            return _datetime.datetime(1900, 1, 1, hour, minute, second)

    return None


def is_peak_time(parsed_dt, peak_start_hour, peak_end_hour):
    """高峰期判定：周一至周五且时间落在 [start, end) 内；周六周日不算高峰期。"""
    if parsed_dt is None:
        return False
    start = _datetime.time(peak_start_hour, 0, 0)
    end = _datetime.time(peak_end_hour, 0, 0)
    if not (start <= parsed_dt.time() < end):
        return False
    return parsed_dt.weekday() < 5


def read_xml_from_zip(xlsx, path):
    try:
        return xlsx.read(path)
    except KeyError:
        return None


def load_shared_strings(xlsx):
    data = read_xml_from_zip(xlsx, "xl/sharedStrings.xml")
    if data is None:
        return []
    root = ET.fromstring(data)
    values = []
    for item in root.findall(NS_MAIN + "si"):
        parts = []
        direct_text = item.find(NS_MAIN + "t")
        if direct_text is not None and direct_text.text is not None:
            parts.append(direct_text.text)
        for text_node in item.findall(".//" + NS_MAIN + "t"):
            if text_node.text is not None and text_node is not direct_text:
                parts.append(text_node.text)
        values.append("".join(parts))
    return values


def relationship_target(base_dir, target):
    if target.startswith("/"):
        return target.lstrip("/")
    if base_dir:
        return os.path.normpath(os.path.join(base_dir, target)).replace(os.sep, "/")
    return os.path.normpath(target).replace(os.sep, "/")


def load_workbook_sheets(xlsx):
    workbook_data = read_xml_from_zip(xlsx, "xl/workbook.xml")
    rels_data = read_xml_from_zip(xlsx, "xl/_rels/workbook.xml.rels")
    if workbook_data is None or rels_data is None:
        raise ValueError("不是有效的 .xlsx 文件：缺少 workbook 元数据")

    rels_root = ET.fromstring(rels_data)
    rels = {}
    for rel in rels_root.findall(NS_PACKAGE_REL + "Relationship"):
        rels[rel.attrib.get("Id")] = relationship_target("xl", rel.attrib.get("Target", ""))

    workbook_root = ET.fromstring(workbook_data)
    sheets = []
    for sheet in workbook_root.findall(".//" + NS_MAIN + "sheet"):
        rel_id = sheet.attrib.get(NS_REL + "id")
        path = rels.get(rel_id)
        if path:
            sheets.append((sheet.attrib.get("name", ""), path))
    if not sheets:
        raise ValueError("没有在 workbook 中找到工作表")
    return sheets


def get_sheet_path(xlsx, sheet_name):
    sheets = load_workbook_sheets(xlsx)
    if not sheet_name:
        return sheets[0][1]
    for name, path in sheets:
        if name == sheet_name:
            return path
    available = ", ".join(name for name, _ in sheets)
    raise ValueError("找不到工作表 {0}；可用工作表：{1}".format(sheet_name, available))


def cell_text(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    value_node = cell.find(NS_MAIN + "v")

    if cell_type == "inlineStr":
        texts = []
        for text_node in cell.findall(".//" + NS_MAIN + "t"):
            if text_node.text is not None:
                texts.append(text_node.text)
        return "".join(texts)

    if value_node is None or value_node.text is None:
        return ""

    raw = value_node.text
    if cell_type == "s":
        try:
            index = int(raw)
        except ValueError:
            return raw
        if 0 <= index < len(shared_strings):
            return shared_strings[index]
        return raw

    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"

    return raw


def iter_rows(xlsx_path, sheet_name):
    with zipfile.ZipFile(xlsx_path, "r") as xlsx:
        shared_strings = load_shared_strings(xlsx)
        sheet_path = get_sheet_path(xlsx, sheet_name)
        data = read_xml_from_zip(xlsx, sheet_path)
        if data is None:
            raise ValueError("找不到工作表内容：{0}".format(sheet_path))

        root = ET.fromstring(data)
        for row in root.findall(".//" + NS_MAIN + "row"):
            row_values = {}
            for cell in row.findall(NS_MAIN + "c"):
                col = column_name(cell.attrib.get("r", ""))
                if col in (COLUMN_TIME, COLUMN_API_KEY, COLUMN_MODEL, COLUMN_PRODUCT, COLUMN_UNIT, COLUMN_USAGE):
                    row_values[col] = cell_text(cell, shared_strings)
            yield row_values


def key_matches(api_key, key_filter):
    """key 过滤：输入内容与 apiKey 完全一致或为其子串（含前缀）即匹配。"""
    for wanted in key_filter:
        if wanted == api_key or wanted in api_key:
            return True
    return False


def split_keys(values):
    result = []
    for item in values or []:
        for part in item.split(","):
            part = part.strip()
            if part:
                result.append(part)
    return result


def expand_files(paths):
    """展开输入参数：支持逗号分隔多个路径；存在的文件直接使用，否则按通配符模式展开；去重且保持顺序。"""
    files = []
    for item in paths:
        for path in item.split(","):
            path = path.strip()
            if not path:
                continue
            if os.path.isfile(path):
                files.append(path)
                continue
            matched = sorted(glob.glob(path))
            if matched:
                files.extend(matched)
            else:
                files.append(path)
    seen = set()
    result = []
    for path in files:
        if path not in seen:
            seen.add(path)
            result.append(path)
    return result


def new_stats():
    return {
        "total_rows": 0,
        "matched_rows": 0,
        "skipped_header_rows": 0,
        "skipped_unit_rows": 0,
        "skipped_empty_key_rows": 0,
        "skipped_bad_usage_rows": 0,
        "skipped_key_filter_rows": 0,
        "skipped_key_exclude_rows": 0,
        "unparsed_time_rows": 0,
        "keys_seen": set(),
    }


def merge_stats(target, source):
    for key in (
        "total_rows",
        "matched_rows",
        "skipped_header_rows",
        "skipped_unit_rows",
        "skipped_empty_key_rows",
        "skipped_bad_usage_rows",
        "skipped_key_filter_rows",
        "skipped_key_exclude_rows",
        "unparsed_time_rows",
    ):
        target[key] += source[key]
    target["keys_seen"].update(source["keys_seen"])


def merge_summary(target, source):
    target.rows += source.rows
    target.total_raw_usage += source.total_raw_usage
    target.off_peak_usage += source.off_peak_usage
    target.peak_raw_usage += source.peak_raw_usage
    target.peak_weighted_usage += source.peak_weighted_usage
    target.billable_usage += source.billable_usage
    target.discounted_usage += source.discounted_usage
    target.flash_billable_usage += source.flash_billable_usage
    target.tool_calls += source.tool_calls
    target.tool_usage += source.tool_usage


def analyze_file(xlsx_path, sheet_name, peak_start_hour, peak_end_hour, key_filter=None, exclude_filter=None):
    summaries = OrderedDict()
    stats = new_stats()

    for row in iter_rows(xlsx_path, sheet_name):
        stats["total_rows"] += 1
        unit = normalize_text(row.get(COLUMN_UNIT))
        api_key = normalize_text(row.get(COLUMN_API_KEY))
        usage = decimal_from_value(row.get(COLUMN_USAGE))
        time_value = row.get(COLUMN_TIME)

        if unit == "单价单位" or api_key.lower() == "apikey" or normalize_text(time_value) == "入账时间":
            stats["skipped_header_rows"] += 1
            continue
        product = normalize_text(row.get(COLUMN_PRODUCT))
        is_tool = any(name in product for name in TOOL_PRODUCTS)
        if unit != TARGET_UNIT and not is_tool:
            stats["skipped_unit_rows"] += 1
            continue
        if not api_key:
            stats["skipped_empty_key_rows"] += 1
            continue
        stats["keys_seen"].add(api_key)
        if key_filter is not None and not key_matches(api_key, key_filter):
            stats["skipped_key_filter_rows"] += 1
            continue
        if exclude_filter is not None and key_matches(api_key, exclude_filter):
            stats["skipped_key_exclude_rows"] += 1
            continue
        if usage is None:
            stats["skipped_bad_usage_rows"] += 1
            continue

        if is_tool:
            # 用量列(S)为调用次数，每次调用固定按 200万 token 计入总量
            summaries.setdefault(api_key, BillSummary(api_key)).add_tool(usage)
            stats["matched_rows"] += 1
            continue

        parsed_dt = parse_datetime_value(time_value)
        if parsed_dt is None:
            stats["unparsed_time_rows"] += 1
        is_peak = is_peak_time(parsed_dt, peak_start_hour, peak_end_hour)
        is_flash = normalize_text(row.get(COLUMN_MODEL)).lower() == FLASH_MODEL

        if api_key not in summaries:
            summaries[api_key] = BillSummary(api_key)
        summaries[api_key].add(usage, is_peak, is_flash)
        stats["matched_rows"] += 1

    # 打折总量 = 总量 - 5.3-flash总量 × 0.5（向上取整，不带 0.5 小数）
    for summary in summaries.values():
        summary.discounted_usage = (
            summary.billable_usage - summary.flash_billable_usage * FLASH_BILL_RATE
        ).to_integral_value(rounding=ROUND_CEILING)

    return summaries, stats


def write_csv(path, summaries):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "apiukey",
            "总量",
            "打折总量",
            "原始总量",
            "非高峰总量",
            "高峰总量",
            "高峰3倍总量",
            "5.3-flash总量",
            "工具调用次数",
            "工具总量",
        ])
        for summary in summaries.values():
            writer.writerow([
                summary.api_key,
                decimal_to_text(summary.billable_usage),
                decimal_to_text(summary.discounted_usage),
                decimal_to_text(summary.total_raw_usage),
                decimal_to_text(summary.off_peak_usage),
                decimal_to_text(summary.peak_raw_usage),
                decimal_to_text(summary.peak_weighted_usage),
                decimal_to_text(summary.flash_billable_usage),
                decimal_to_text(summary.tool_calls),
                decimal_to_text(summary.tool_usage),
            ])


def print_table(summaries):
    headers = [
        "apiukey",
        "总量",
        "打折总量",
        "原始总量",
        "非高峰总量",
        "高峰总量",
        "高峰3倍总量",
        "5.3-flash总量",
        "工具调用次数",
        "工具总量",
    ]
    rows = []
    for summary in summaries.values():
        rows.append([
            summary.api_key,
            decimal_to_text(summary.billable_usage),
            decimal_to_text(summary.discounted_usage),
            decimal_to_text(summary.total_raw_usage),
            decimal_to_text(summary.off_peak_usage),
            decimal_to_text(summary.peak_raw_usage),
            decimal_to_text(summary.peak_weighted_usage),
            decimal_to_text(summary.flash_billable_usage),
            decimal_to_text(summary.tool_calls),
            decimal_to_text(summary.tool_usage),
        ])

    if not rows:
        print("没有找到 单价单位=千token 的有效记录。")
        return

    widths = []
    for index, header in enumerate(headers):
        width = display_width(header)
        for row in rows:
            width = max(width, display_width(row[index]))
        widths.append(width)

    def format_row(values):
        padded = []
        for index, value in enumerate(values):
            align = "left" if index == 0 else "right"
            padded.append(pad_display(value, widths[index], align))
        return "  ".join(padded)

    print(format_row(headers))
    print(format_row(["-" * width for width in widths]))
    for row in rows:
        print(format_row(row))


def hour_arg(value):
    try:
        hour = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError("小时必须是 0-24 之间的整数")
    if hour < 0 or hour > 24:
        raise argparse.ArgumentTypeError("小时必须是 0-24 之间的整数")
    return hour


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="分析智谱账单 Excel：按 apiKey 汇总千token用量，并对 14:00~18:00 高峰期用量按 3 倍计费。支持传入多个文件（可用通配符）。"
    )
    parser.add_argument(
        "excel",
        nargs="+",
        help="输入 .xlsx 文件路径，可传多个（空格或逗号分隔）；支持通配符模式（如 'zhipu_*.xlsx'）",
    )
    parser.add_argument("--sheet", default="", help="工作表名称；默认读取第一个工作表")
    parser.add_argument("--output", "-o", default="", help="可选：输出 CSV 文件路径")
    parser.add_argument("--peak-start", type=hour_arg, default=14, help="高峰期开始小时，默认 14")
    parser.add_argument("--peak-end", type=hour_arg, default=18, help="高峰期结束小时（不包含），默认 18")
    parser.add_argument(
        "--key",
        action="append",
        default=[],
        dest="keys",
        help="只分析指定的 apiKey；可多次使用 --key，或用逗号分隔多个 key。输入完整 key 或其任意片段（子串匹配）均可",
    )
    parser.add_argument(
        "--exclude-key",
        "-e",
        action="append",
        default=[],
        dest="exclude_keys",
        help="不统计指定的 apiKey（与 --key 规则相同）；可多次使用或用逗号分隔",
    )
    return parser.parse_args(argv[1:])


def main(argv):
    args = parse_args(argv)
    if args.peak_start >= args.peak_end:
        return fail("--peak-start 必须小于 --peak-end")

    files = expand_files(args.excel)
    missing = [path for path in files if not os.path.isfile(path)]
    if missing:
        return fail("输入文件不存在：{0}".format(", ".join(missing)))

    key_filter = split_keys(args.keys)
    if args.keys and not key_filter:
        return fail("--key 参数为空")
    if not key_filter:
        key_filter = None
    exclude_filter = split_keys(args.exclude_keys)
    if args.exclude_keys and not exclude_filter:
        return fail("--exclude-key 参数为空")
    if not exclude_filter:
        exclude_filter = None

    merged = OrderedDict()
    merged_stats = new_stats()
    file_results = []

    try:
        for path in files:
            summaries, stats = analyze_file(
                path, args.sheet, args.peak_start, args.peak_end, key_filter, exclude_filter
            )
            file_results.append((path, summaries, stats))
            for api_key, summary in summaries.items():
                if api_key not in merged:
                    merged[api_key] = BillSummary(api_key)
                merge_summary(merged[api_key], summary)
            merge_stats(merged_stats, stats)
    except zipfile.BadZipFile:
        return fail("输入文件不是有效的 .xlsx：{0}".format(path))
    except (ET.ParseError, ValueError) as exc:
        return fail("{0}：{1}".format(path, exc))

    summaries = merged
    stats = merged_stats

    # 打折总量 = 总量 - 5.3-flash总量 × 0.5（向上取整）；在多文件汇总后按 key 统一取整，避免分文件取整累积偏差
    for summary in summaries.values():
        summary.discounted_usage = (
            summary.billable_usage - summary.flash_billable_usage * FLASH_BILL_RATE
        ).to_integral_value(rounding=ROUND_CEILING)

    if key_filter:
        missing = [wanted for wanted in key_filter if not any(key_matches(k, [wanted]) for k in stats["keys_seen"])]
        if missing:
            print("提示：以下 key 未在账单中找到：{0}".format(", ".join(missing)), file=sys.stderr)
        if not summaries:
            return fail("指定的 key 均未匹配到任何 单价单位=千token 的记录")
    elif not summaries and exclude_filter is not None:
        print("提示：排除后没有任何有效记录。", file=sys.stderr)

    if len(files) > 1:
        print("分文件统计（共 {0} 个文件）：".format(len(files)))
        for path, file_summaries, file_stats in file_results:
            file_billable = sum((s.billable_usage for s in file_summaries.values()), Decimal("0"))
            print(
                "  {0}  有效行 {1}  总量 {2}".format(
                    pad_display(os.path.basename(path), 30),
                    file_stats["matched_rows"],
                    decimal_to_text(file_billable),
                )
            )
        print("")

    print_table(summaries)
    total_raw_usage = sum((summary.total_raw_usage for summary in summaries.values()), Decimal("0"))
    total_off_peak_usage = sum((summary.off_peak_usage for summary in summaries.values()), Decimal("0"))
    total_peak_raw_usage = sum((summary.peak_raw_usage for summary in summaries.values()), Decimal("0"))
    total_peak_weighted_usage = sum((summary.peak_weighted_usage for summary in summaries.values()), Decimal("0"))
    total_billable_usage = sum((summary.billable_usage for summary in summaries.values()), Decimal("0"))
    total_discounted_usage = sum((summary.discounted_usage for summary in summaries.values()), Decimal("0"))
    total_flash_billable_usage = sum((summary.flash_billable_usage for summary in summaries.values()), Decimal("0"))
    total_tool_calls = sum((summary.tool_calls for summary in summaries.values()), Decimal("0"))
    total_tool_usage = sum((summary.tool_usage for summary in summaries.values()), Decimal("0"))

    print("")
    print("总量汇总：")
    print("  总量(非高峰+高峰3倍+工具200万/次，未打折)：{0}".format(decimal_to_text(total_billable_usage)))
    print("  打折总量(总量-5.3-flash总量×0.5)：{0}".format(decimal_to_text(total_discounted_usage)))
    print("  原始总量：{0}".format(decimal_to_text(total_raw_usage)))
    print("  非高峰总量：{0}".format(decimal_to_text(total_off_peak_usage)))
    print("  高峰总量(未加倍)：{0}".format(decimal_to_text(total_peak_raw_usage)))
    print("  高峰3倍总量：{0}".format(decimal_to_text(total_peak_weighted_usage)))
    print("  5.3-flash总量(非高峰+高峰3倍)：{0}".format(decimal_to_text(total_flash_billable_usage)))
    print("  工具调用次数(【web-reader】/【search-prime】)：{0}".format(decimal_to_text(total_tool_calls)))
    print("  工具总量(200万token/次)：{0}".format(decimal_to_text(total_tool_usage)))
    print("")
    print("处理统计：")
    print("  输入文件数：{0}".format(len(files)))
    print("  总行数：{0}".format(stats["total_rows"]))
    print("  有效千token记录：{0}".format(stats["matched_rows"]))
    print("  跳过表头行：{0}".format(stats["skipped_header_rows"]))
    print("  跳过非千token行(工具按次行除外)：{0}".format(stats["skipped_unit_rows"]))
    print("  跳过空 apiKey 行：{0}".format(stats["skipped_empty_key_rows"]))
    print("  跳过无效用量行：{0}".format(stats["skipped_bad_usage_rows"]))
    if key_filter:
        print("  跳过非指定 key 行：{0}".format(stats["skipped_key_filter_rows"]))
    if exclude_filter:
        print("  跳过排除 key 行：{0}".format(stats["skipped_key_exclude_rows"]))
    print("  未能解析时间的有效行：{0}".format(stats["unparsed_time_rows"]))
    if stats["unparsed_time_rows"]:
        print("  提示：未能解析时间的有效行已按非高峰计费，请人工核对入账时间格式。")
    print("  高峰期规则：{0:02d}:00 <= 入账时间 < {1:02d}:00（仅周一至周五，周六周日不算高峰），高峰期用量按 3 倍计入最终计费用量".format(
        args.peak_start, args.peak_end
    ))
    print("  5.3-flash规则：总量=非高峰+高峰3倍(全模型，未打折)；打折总量=总量-5.3-flash总量×0.5（向上取整，仅flash部分打5折）")
    print("  工具规则：【web-reader】/【search-prime】应用组件(单价单位=次，S列=调用次数)每次按200万token固定计入总量，不参与高峰3倍与打折")

    if args.output:
        write_csv(args.output, summaries)
        print("CSV 已写入：{0}".format(args.output))

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

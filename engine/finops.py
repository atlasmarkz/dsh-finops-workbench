#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""经营分析引擎：读取 CSV/XLSX，做数据体检、口径推断、分析与 Excel/Word 交付物。

由 Node 侧工作台服务通过 stdin 传入 JSON 载荷，stdout 返回 JSON 结果。
命令：inspect | analyze | report | sample | selftest
"""
import sys
import os
import io
import re
import csv
import json
import math
import traceback
import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

MAX_ROWS = 200000
DETAIL_SHEET_LIMIT = 20000
TOTAL_WORDS = {"合计", "总计", "小计", "汇总", "总和", "合计：", "总计：", "total", "sum"}


# ---------------------------------------------------------------- 基础工具

def norm_cell(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value).strip()
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    return s


def parse_number(value):
    """把常见金额写法解析为 float；无法解析返回 None。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    s = norm_cell(value)
    if not s:
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    s = s.replace(",", "").replace("，", "").replace(" ", "")
    s = s.replace("￥", "").replace("¥", "").replace("$", "").replace("元", "")
    scale = 1.0
    if s.endswith("万"):
        scale = 10000.0
        s = s[:-1]
    elif s.endswith("亿"):
        scale = 100000000.0
        s = s[:-1]
    percent = False
    if s.endswith("%"):
        percent = True
        s = s[:-1]
    if s.startswith("+"):
        s = s[1:]
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", s):
        return None
    try:
        number = float(s) * scale
    except ValueError:
        return None
    if percent:
        number = number / 100.0
    if negative:
        number = -number
    return number


def parse_date(value):
    """把常见日期/月份写法解析为 date；无法解析返回 None。"""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel 序列日期（1900 体系）
        if 20000 <= float(value) <= 80000:
            try:
                base = datetime.date(1899, 12, 30)
                return base + datetime.timedelta(days=int(value))
            except Exception:
                return None
        if 190001 <= int(value) <= 220012:
            return parse_date(str(int(value)))
        return None
    s = norm_cell(value)
    if not s:
        return None
    s = s.replace("年", "-").replace("月", "-").replace("日", "").replace("/", "-").replace(".", "-")
    s = re.sub(r"\s+", "", s)
    s = s.rstrip("-")
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})(\d{2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2})-(\d{1,2})", s)
    if m:
        year = datetime.date.today().year
        try:
            return datetime.date(year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def is_total_row(row):
    for cell in row[:4]:
        if norm_cell(cell).lower() in {w.lower() for w in TOTAL_WORDS}:
            return True
    return False


def read_csv_rows(path):
    raw = Path(path).read_bytes()
    text = None
    used = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            text = raw.decode(enc)
            used = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("无法识别文件编码，请另存为 UTF-8 或 GBK 的 CSV。")
    sample = text[:8192]
    dialect = csv.excel
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
    except Exception:
        pass
    rows = []
    for row in csv.reader(io.StringIO(text), dialect):
        rows.append([norm_cell(c) for c in row])
    return rows, {"encoding": used, "kind": "csv"}


def read_xlsx_rows(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    names = list(wb.sheetnames)
    target = sheet if sheet and sheet in names else names[0]
    ws = wb[target]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if v is None else v for v in row])
        if len(rows) > MAX_ROWS:
            break
    wb.close()
    return names, target, rows, {"kind": "xlsx"}


def load_table(path, sheet=None):
    suffix = Path(path).suffix.lower()
    if suffix in (".csv", ".txt", ".tsv"):
        rows, meta = read_csv_rows(path)
        names = [Path(path).stem]
        target = names[0]
    elif suffix in (".xlsx", ".xlsm"):
        names, target, rows, meta = read_xlsx_rows(path, sheet)
    else:
        raise ValueError("只支持 .xlsx / .xlsm / .csv 数据文件。")
    rows = [r for r in rows if any(norm_cell(c) for c in r)]
    if not rows:
        raise ValueError("文件里没有可读取的数据。")
    header_index = detect_header(rows)
    header = [norm_cell(c) for c in rows[header_index]]
    width = max(len(r) for r in rows)
    headers = []
    seen = {}
    for i in range(width):
        name = header[i] if i < len(header) else ""
        name = re.sub(r"\s+", " ", name).strip()
        if not name:
            name = "列%d" % (i + 1)
        base = name
        count = seen.get(base, 0)
        seen[base] = count + 1
        if count:
            name = "%s_%d" % (base, count + 1)
        headers.append(name)
    data = []
    for r in rows[header_index + 1:]:
        values = ["" if v is None else v for v in r]
        if len(values) < width:
            values = values + [""] * (width - len(values))
        elif len(values) > width:
            values = values[:width]
        data.append(values)
    return {
        "path": str(Path(path).resolve()),
        "sheet": target,
        "sheetNames": names,
        "headers": headers,
        "rows": data,
        "meta": meta,
        "headerRow": header_index + 1,
    }


def detect_header(rows):
    best = 0
    for i, row in enumerate(rows[:12]):
        vals = [norm_cell(c) for c in row]
        nonempty = [v for v in vals if v]
        if len(nonempty) < 2:
            continue
        numeric = sum(1 for v in nonempty if parse_number(v) is not None)
        dates = sum(1 for v in nonempty if parse_date(v) is not None)
        unique = len(set(nonempty))
        if numeric + dates <= max(1, int(len(nonempty) * 0.3)) and unique >= len(nonempty) * 0.8:
            return i
    return best


def infer_column(values, header):
    total = 0
    numeric = 0
    dates = 0
    missing = 0
    samples = []
    for v in values:
        s = norm_cell(v)
        if not s:
            missing += 1
            continue
        total += 1
        if len(samples) < 4 and s not in samples:
            samples.append(s)
        if parse_number(s) is not None:
            numeric += 1
        elif parse_date(s) is not None:
            dates += 1
    if total == 0:
        kind = "empty"
    elif numeric / total >= 0.8:
        kind = "number"
    elif dates / total >= 0.8:
        kind = "date"
    else:
        kind = "text"
    if re.search(r"日期|时间|月份|date|period|month", header, re.I) and dates > 0:
        kind = "date"
    unit = ""
    if re.search(r"万元|\(万\)|（万）", header):
        unit = "万元"
    elif re.search(r"\(元\)|（元）|金额", header):
        unit = "元"
    return {
        "type": kind,
        "numericRatio": round(numeric / total, 4) if total else 0,
        "dateRatio": round(dates / total, 4) if total else 0,
        "missing": missing,
        "missingRatio": round(missing / len(values), 4) if values else 0,
        "samples": samples,
        "unit": unit,
    }


def suggest_mapping(columns):
    mapping = {"date": None, "revenue": None, "cost": None, "expense": None, "dimension": None}

    def pick(names, patterns, kinds):
        candidates = []
        for i, col in enumerate(names):
            if kinds and col["type"] not in kinds:
                continue
            for order, pat in enumerate(patterns):
                if re.search(pat, col["name"], re.I):
                    candidates.append((order, i))
                    break
        if not candidates:
            return None
        candidates.sort()
        return names[candidates[0][1]]["name"]

    names = [dict(name=h, **infer_column([r[i] for r in CURRENT_ROWS], h)) for i, h in enumerate(CURRENT_HEADERS)]
    mapping["date"] = pick(names, [r"^日期$", r"^时间$", r"日期", r"时间", r"月份", r"date", r"month"], ["date"])
    mapping["revenue"] = pick(names, [r"^营业收入$", r"^营收$", r"^收入$", r"营业收入", r"营收", r"销售收入", r"销售额", r"营业额", r"^收入", r"回款", r"revenue", r"sales"], ["number"])
    mapping["cost"] = pick(names, [r"^营业成本$", r"^成本$", r"营业成本", r"成本", r"采购", r"进价", r"cost", r"cogs"], ["number"])
    mapping["expense"] = pick(names, [r"^费用$", r"^费用合计$", r"^支出$", r"费用", r"支出", r"expense"], ["number"])
    dim = pick(names, [r"渠道", r"平台", r"门店", r"品类", r"分类", r"商品", r"产品", r"客户", r"区域", r"部门", r"channel", r"category", r"product"], ["text"])
    if not dim:
        for col in names:
            if col["type"] == "text" and len(col["samples"]) >= 2:
                dim = col["name"]
                break
    mapping["dimension"] = dim
    return mapping


CURRENT_HEADERS = []
CURRENT_ROWS = []


def build_profile(table):
    global CURRENT_HEADERS, CURRENT_ROWS
    CURRENT_HEADERS = table["headers"]
    CURRENT_ROWS = table["rows"]
    issues = []
    columns = []
    total = len(table["rows"])
    for i, header in enumerate(table["headers"]):
        values = [r[i] for r in table["rows"]]
        info = infer_column(values, header)
        info["index"] = i
        info["name"] = header
        if info["missingRatio"] >= 0.5 and total:
            issues.append({"level": "warn", "code": "missing", "title": "列「%s」缺失较多" % header,
                           "detail": "%.0f%% 的单元格为空，分析时可能被忽略。" % (info["missingRatio"] * 100)})
        columns.append(info)
    if total < 3:
        issues.append({"level": "high", "code": "few-rows", "title": "数据行很少",
                       "detail": "只有 %d 行数据，趋势分析参考意义有限。" % total})
    return {"columns": columns, "issues": issues}


def pick_rows(table, mapping):
    """按口径抽取数据，返回 records 与质量统计。"""
    names = table["headers"]
    index = {name: i for i, name in enumerate(names)}
    date_col = mapping.get("date")
    rev_col = mapping.get("revenue")
    cost_col = mapping.get("cost")
    exp_col = mapping.get("expense")
    dim_col = mapping.get("dimension")
    if not rev_col:
        raise ValueError("还没有指定「营业收入」列，无法分析。")
    records = []
    stats = {"total": 0, "excluded": 0, "missingDate": 0, "missingRevenue": 0,
             "negativeRevenue": 0, "duplicates": 0, "totalRows": 0, "badNumbers": 0}
    seen = set()
    for row in table["rows"]:
        stats["total"] += 1
        if is_total_row(row):
            stats["totalRows"] += 1
            stats["excluded"] += 1
            continue
        def cell(col):
            if not col or col not in index:
                return ""
            return norm_cell(row[index[col]])
        revenue = parse_number(cell(rev_col))
        if revenue is None:
            stats["missingRevenue"] += 1
            stats["excluded"] += 1
            continue
        cost = parse_number(cell(cost_col)) if cost_col else 0.0
        if cost_col and cost is None:
            stats["badNumbers"] += 1
            cost = 0.0
        expense = parse_number(cell(exp_col)) if exp_col else 0.0
        if exp_col and expense is None:
            stats["badNumbers"] += 1
            expense = 0.0
        date_value = parse_date(cell(date_col)) if date_col else None
        if date_col and date_value is None:
            stats["missingDate"] += 1
        if revenue < 0:
            stats["negativeRevenue"] += 1
        key = tuple(norm_cell(c) for c in row)
        if key in seen:
            stats["duplicates"] += 1
        else:
            seen.add(key)
        records.append({
            "period": date_value.strftime("%Y-%m") if date_value else "全部",
            "date": date_value.strftime("%Y-%m-%d") if date_value else "",
            "revenue": revenue,
            "cost": cost,
            "expense": expense,
            "gross": revenue - cost,
            "net": revenue - cost - expense,
            "dimension": cell(dim_col) or "未分类",
            "row": [norm_cell(c) for c in row],
        })
    if not records:
        raise ValueError("按当前口径没有可用数据：请检查「营业收入」列和日期格式。")
    return records, stats


def aggregate(records):
    months = {}
    for r in records:
        m = months.setdefault(r["period"], {"period": r["period"], "revenue": 0.0, "cost": 0.0,
                                             "expense": 0.0, "gross": 0.0, "net": 0.0, "rows": 0})
        for key in ("revenue", "cost", "expense", "gross", "net"):
            m[key] += r[key]
        m["rows"] += 1
    trend = [months[k] for k in sorted(months)]
    for i, row in enumerate(trend):
        row["margin"] = row["gross"] / row["revenue"] if row["revenue"] else None
        row["netMargin"] = row["net"] / row["revenue"] if row["revenue"] else None
        row["mom"] = (row["revenue"] / trend[i - 1]["revenue"] - 1) if i and trend[i - 1]["revenue"] else None
    return trend


def structure(records, trend_revenue):
    groups = {}
    for r in records:
        g = groups.setdefault(r["dimension"], {"name": r["dimension"], "revenue": 0.0, "cost": 0.0,
                                               "expense": 0.0, "gross": 0.0, "net": 0.0, "rows": 0})
        for key in ("revenue", "cost", "expense", "gross", "net"):
            g[key] += r[key]
        g["rows"] += 1
    items = sorted(groups.values(), key=lambda g: g["revenue"], reverse=True)
    total = sum(g["revenue"] for g in items)
    for g in items:
        g["share"] = g["revenue"] / total if total else 0
        g["margin"] = g["gross"] / g["revenue"] if g["revenue"] else None
    top = items[:10]
    rest = items[10:]
    if rest:
        other = {"name": "其他（%d 项）" % len(rest)}
        for key in ("revenue", "cost", "expense", "gross", "net", "rows"):
            other[key] = sum(g[key] for g in rest)
        other["share"] = other["revenue"] / total if total else 0
        other["margin"] = other["gross"] / other["revenue"] if other["revenue"] else None
        top = top + [other]
    return {"items": top, "count": len(items), "totalRevenue": total}


def anomalies(records, trend, structure_data, stats, table):
    result = []

    def add(level, title, detail, value=""):
        result.append({"level": level, "title": title, "detail": detail, "value": value})

    if stats["missingRevenue"]:
        add("high", "有 %d 行缺少营业收入" % stats["missingRevenue"],
            "这些行已排除在分析之外；如果它们本应是有效记录，请检查金额列格式。", str(stats["missingRevenue"]))
    if stats["missingDate"] and table.get("sheet"):
        mapping_date = CURRENT_MAPPING.get("date")
        if mapping_date:
            add("warn", "有 %d 行日期无法识别" % stats["missingDate"],
                "这些行被计入「全部」期间，可能影响月度趋势。", str(stats["missingDate"]))
    if stats["duplicates"]:
        add("warn", "发现 %d 行重复记录" % stats["duplicates"],
            "重复行没有自动删除，可能抬高收入；建议回到源数据核对。", str(stats["duplicates"]))
    if stats["negativeRevenue"]:
        add("warn", "有 %d 行为负收入" % stats["negativeRevenue"],
            "负收入通常代表退款或冲销；当前按原值计入。", str(stats["negativeRevenue"]))
    if stats["badNumbers"]:
        add("info", "有 %d 个成本/费用单元格无法识别为数字" % stats["badNumbers"],
            "这些单元格按 0 处理，可能低估成本。", str(stats["badNumbers"]))
    if stats["totalRows"]:
        add("info", "已跳过 %d 行合计行" % stats["totalRows"], "含「合计/总计/小计」的行不参与分析。", str(stats["totalRows"]))
    revs = [r["revenue"] for r in records]
    if len(revs) >= 4:
        mean = sum(revs) / len(revs)
        variance = sum((x - mean) ** 2 for x in revs) / len(revs)
        sigma = math.sqrt(variance)
        if sigma:
            outliers = [r for r in records if abs(r["revenue"] - mean) > 3 * sigma]
            if outliers:
                add("warn", "发现 %d 条极端金额记录" % len(outliers),
                    "例如：%s（%s，%.0f 元）。建议核对是否录入了错误单位。" % (
                        outliers[0]["dimension"], outliers[0]["period"], outliers[0]["revenue"]), str(len(outliers)))
    for row in trend:
        if row.get("mom") is not None and abs(row["mom"]) >= 0.2:
            add("warn" if row["mom"] < 0 else "info", "%s 收入环比 %+.0f%%" % (row["period"], row["mom"] * 100),
                "上月 %.0f → 本月 %.0f。" % (
                    trend[trend.index(row) - 1]["revenue"] if trend.index(row) else 0, row["revenue"]),
                "%.1f%%" % (row["mom"] * 100))
    loss_months = [r for r in trend if r["net"] < 0]
    if loss_months:
        add("high", "%d 个月净利润为负" % len(loss_months),
            "分别出现在：%s。" % "、".join(str(r["period"]) for r in loss_months), str(len(loss_months)))
    losses = [g for g in structure_data["items"] if g["margin"] is not None and g["margin"] < 0 and g["name"] != "其他"]
    if losses:
        add("warn", "%d 个维度毛利率为负" % len(losses),
            "例如：%s（毛利率 %.0f%%，收入 %.0f）。" % (losses[0]["name"], losses[0]["margin"] * 100, losses[0]["revenue"]),
            str(len(losses)))
    if structure_data["items"] and structure_data["items"][0]["share"] >= 0.4 and structure_data["count"] > 3:
        add("info", "收入集中度较高", "「%s」占收入 %.0f%%，单一渠道波动会明显影响整体。" % (
            structure_data["items"][0]["name"], structure_data["items"][0]["share"] * 100),
            "%.0f%%" % (structure_data["items"][0]["share"] * 100))
    if not result:
        add("info", "未发现明显异常", "数据结构与波动都在常见范围内，可以生成交付物。", "")
    return result


CURRENT_MAPPING = {}


def analyze_payload(payload):
    global CURRENT_MAPPING
    table = load_table(payload["path"], payload.get("sheet"))
    mapping = payload.get("mapping") or suggest_mapping_or_default(table)
    CURRENT_MAPPING = mapping
    records, stats = pick_rows(table, mapping)
    trend = aggregate(records)
    structure_data = structure(records, sum(r["revenue"] for r in records))
    total_revenue = sum(r["revenue"] for r in records)
    total_cost = sum(r["cost"] for r in records)
    total_expense = sum(r["expense"] for r in records)
    gross = total_revenue - total_cost
    net = gross - total_expense
    months = len([r for r in trend if r["period"] != "全部"])
    trend_direction = 0.0
    if len(trend) >= 2 and trend[0]["revenue"]:
        trend_direction = trend[-1]["revenue"] / trend[0]["revenue"] - 1
    kpis = {
        "revenue": total_revenue,
        "cost": total_cost,
        "expense": total_expense,
        "gross": gross,
        "net": net,
        "grossMargin": gross / total_revenue if total_revenue else None,
        "netMargin": net / total_revenue if total_revenue else None,
        "expenseRatio": total_expense / total_revenue if total_revenue else None,
        "avgRevenue": total_revenue / months if months else total_revenue,
        "trendDirection": trend_direction,
        "months": months,
        "records": len(records),
        "periodStart": trend[0]["period"] if trend else "",
        "periodEnd": trend[-1]["period"] if trend else "",
    }
    quality = {
        "stats": stats,
        "issues": build_profile(table)["issues"],
        "columns": build_profile(table)["columns"],
        "dateRange": {"min": min((r["date"] for r in records if r["date"]), default=""),
                       "max": max((r["date"] for r in records if r["date"]), default="")},
        "unit": (next((c for c in build_profile(table)["columns"] if c["name"] == mapping.get("revenue")), {}) or {}).get("unit", ""),
    }
    return {
        "fileName": Path(table["path"]).name,
        "path": table["path"],
        "sheet": table["sheet"],
        "sheetNames": table["sheetNames"],
        "mapping": mapping,
        "headers": table["headers"],
        "kpis": kpis,
        "trend": trend,
        "structure": structure_data,
        "anomalies": anomalies(records, trend, structure_data, stats, table),
        "quality": quality,
    }


def suggest_mapping_or_default(table):
    global CURRENT_HEADERS, CURRENT_ROWS
    CURRENT_HEADERS = table["headers"]
    CURRENT_ROWS = table["rows"]
    return suggest_mapping(table["headers"])


# ---------------------------------------------------------------- 报告生成

FONT = "微软雅黑"
NAVY = "1F3864"
LIGHT = "D9E2F3"
GREEN = "2E7D32"
RED = "C62828"


def style_header_row(ws, row, cols):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border


def set_widths(ws, widths):
    from openpyxl.utils import get_column_letter
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def money_cell(cell):
    cell.number_format = '#,##0.00;[Red]-#,##0.00'
    cell.font = cell.font.copy(name=FONT)


def pct_cell(cell):
    cell.number_format = '0.0%'
    cell.font = cell.font.copy(name=FONT)


def make_xlsx(analysis, out_path, compare=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    kpis = analysis["kpis"]
    trend = analysis["trend"]
    structure_items = analysis["structure"]["items"]
    ws = wb.active
    ws.title = "概览"
    ws["A1"] = "经营分析概览"
    ws["A1"].font = Font(name=FONT, bold=True, size=16, color=NAVY)
    ws.merge_cells("A1:D1")
    meta = [
        ("数据源", analysis["fileName"]),
        ("工作表", analysis["sheet"]),
        ("分析期间", "%s ~ %s" % (kpis["periodStart"], kpis["periodEnd"]) if kpis["months"] else "全部数据"),
        ("记录数", "%d 行" % kpis["records"]),
        ("生成时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    row = 3
    for label, value in meta:
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT, bold=True, color="595959")
        ws.cell(row=row, column=2, value=value).font = Font(name=FONT)
        row += 1
    row += 1
    headers = ["指标", "数值", "说明", "口径"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row, column=i, value=h)
    style_header_row(ws, row, 4)
    trend_sheet = "'趋势'"
    trend_rows = len(trend)
    total_row = 5 + trend_rows  # 表头在第 4 行，数据从第 5 行开始，其后一行是合计
    rows = [
        ("营业收入", "=%s!B%d" % (trend_sheet, total_row), "按期间汇总", "收入列求和"),
        ("营业成本", "=%s!C%d" % (trend_sheet, total_row), "按期间汇总", "成本列求和"),
        ("毛利", "=B%d-B%d" % (row + 1, row + 2), "收入 - 成本", "毛利"),
        ("毛利率", "=IFERROR(B%d/B%d,0)" % (row + 3, row + 1), "毛利 / 收入", "比率"),
        ("期间费用", "=%s!F%d" % (trend_sheet, total_row), "营销、人工等", "费用列求和"),
        ("净利润", "=B%d-B%d" % (row + 3, row + 5), "毛利 - 费用", "净利"),
        ("净利率", "=IFERROR(B%d/B%d,0)" % (row + 6, row + 1), "净利 / 收入", "比率"),
        ("月均收入", "=%s!B%d/%d" % (trend_sheet, total_row, max(1, kpis["months"] or 1)), "总收入 / 月数", "月数见趋势表"),
    ]
    start = row + 1
    for i, (name, formula, note, rule) in enumerate(rows):
        r = start + i
        ws.cell(row=r, column=1, value=name).font = Font(name=FONT, bold=name in ("营业收入", "净利润"))
        value_cell = ws.cell(row=r, column=2, value=formula)
        if name in ("毛利率", "净利率"):
            pct_cell(value_cell)
        else:
            money_cell(value_cell)
        ws.cell(row=r, column=3, value=note).font = Font(name=FONT, color="595959")
        ws.cell(row=r, column=4, value=rule).font = Font(name=FONT, color="595959")
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = border
    ws.cell(row=start + 8, column=1, value="注：本表使用原生 Excel 公式，打开文件后会自动计算。").font = Font(name=FONT, color="808080", size=9)
    set_widths(ws, [16, 20, 24, 20])
    ws.freeze_panes = "A11"

    # 趋势
    ws2 = wb.create_sheet("趋势")
    ws2["A1"] = "月度趋势"
    ws2["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws2.merge_cells("A1:I1")
    headers = ["期间", "营业收入", "营业成本", "毛利", "毛利率", "费用", "净利润", "净利率", "收入环比"]
    for i, h in enumerate(headers, start=1):
        ws2.cell(row=4, column=i, value=h)
    style_header_row(ws2, 4, len(headers))
    first_data = 5
    for i, t in enumerate(trend):
        r = first_data + i
        ws2.cell(row=r, column=1, value=t["period"]).font = Font(name=FONT)
        ws2.cell(row=r, column=2, value=t["revenue"])
        money_cell(ws2.cell(row=r, column=2))
        ws2.cell(row=r, column=3, value=t["cost"])
        money_cell(ws2.cell(row=r, column=3))
        ws2.cell(row=r, column=4, value="=B%d-C%d" % (r, r))
        money_cell(ws2.cell(row=r, column=4))
        ws2.cell(row=r, column=5, value="=IFERROR(D%d/B%d,0)" % (r, r))
        pct_cell(ws2.cell(row=r, column=5))
        ws2.cell(row=r, column=6, value=t["expense"])
        money_cell(ws2.cell(row=r, column=6))
        ws2.cell(row=r, column=7, value="=D%d-F%d" % (r, r))
        money_cell(ws2.cell(row=r, column=7))
        ws2.cell(row=r, column=8, value="=IFERROR(G%d/B%d,0)" % (r, r))
        pct_cell(ws2.cell(row=r, column=8))
        if i == 0:
            ws2.cell(row=r, column=9, value=None)
        else:
            ws2.cell(row=r, column=9, value="=IFERROR(B%d/B%d-1,0)" % (r, r - 1))
            pct_cell(ws2.cell(row=r, column=9))
        for c in range(1, 10):
            ws2.cell(row=r, column=c).border = border
    tr = first_data + len(trend)
    ws2.cell(row=tr, column=1, value="合计").font = Font(name=FONT, bold=True)
    for col in (2, 3, 6, 7):
        cell = ws2.cell(row=tr, column=col, value="=SUM(%s%d:%s%d)" % (
            get_column_letter(col), first_data, get_column_letter(col), tr - 1))
        money_cell(cell)
        cell.font = Font(name=FONT, bold=True)
    ws2.cell(row=tr, column=4, value="=B%d-C%d" % (tr, tr))
    money_cell(ws2.cell(row=tr, column=4))
    ws2.cell(row=tr, column=5, value="=IFERROR(D%d/B%d,0)" % (tr, tr))
    pct_cell(ws2.cell(row=tr, column=5))
    ws2.cell(row=tr, column=8, value="=IFERROR(G%d/B%d,0)" % (tr, tr))
    pct_cell(ws2.cell(row=tr, column=8))
    for c in range(1, 10):
        ws2.cell(row=tr, column=c).border = border
        ws2.cell(row=tr, column=c).fill = PatternFill("solid", fgColor=LIGHT)
    set_widths(ws2, [12, 16, 16, 16, 12, 14, 16, 12, 12])
    ws2.freeze_panes = "A5"

    # 结构
    ws3 = wb.create_sheet("结构")
    ws3["A1"] = "收入结构（按维度）"
    ws3["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws3.merge_cells("A1:F1")
    dim_name = analysis["mapping"].get("dimension") or "维度"
    headers = [dim_name, "营业收入", "收入占比", "毛利", "毛利率", "记录数"]
    for i, h in enumerate(headers, start=1):
        ws3.cell(row=4, column=i, value=h)
    style_header_row(ws3, 4, len(headers))
    first = 5
    for i, item in enumerate(structure_items):
        r = first + i
        ws3.cell(row=r, column=1, value=item["name"]).font = Font(name=FONT)
        ws3.cell(row=r, column=2, value=item["revenue"])
        money_cell(ws3.cell(row=r, column=2))
        ws3.cell(row=r, column=3, value="=IFERROR(B%d/$B$%d,0)" % (r, first + len(structure_items)))
        pct_cell(ws3.cell(row=r, column=3))
        ws3.cell(row=r, column=4, value=item["gross"])
        money_cell(ws3.cell(row=r, column=4))
        ws3.cell(row=r, column=5, value="=IFERROR(D%d/B%d,0)" % (r, r))
        pct_cell(ws3.cell(row=r, column=5))
        ws3.cell(row=r, column=6, value=item["rows"]).font = Font(name=FONT)
        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = border
    tr = first + len(structure_items)
    ws3.cell(row=tr, column=1, value="合计").font = Font(name=FONT, bold=True)
    for col in (2, 4, 6):
        cell = ws3.cell(row=tr, column=col, value="=SUM(%s%d:%s%d)" % (
            get_column_letter(col), first, get_column_letter(col), tr - 1))
        cell.font = Font(name=FONT, bold=True)
        if col in (2, 4):
            money_cell(cell)
    ws3.cell(row=tr, column=3, value="=IFERROR(B%d/$B$%d,0)" % (tr, tr))
    pct_cell(ws3.cell(row=tr, column=3))
    ws3.cell(row=tr, column=5, value="=IFERROR(D%d/B%d,0)" % (tr, tr))
    pct_cell(ws3.cell(row=tr, column=5))
    for c in range(1, 7):
        ws3.cell(row=tr, column=c).border = border
        ws3.cell(row=tr, column=c).fill = PatternFill("solid", fgColor=LIGHT)
    set_widths(ws3, [24, 18, 12, 18, 12, 10])

    # 异常
    ws4 = wb.create_sheet("异常")
    ws4["A1"] = "异常与提示"
    ws4["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws4.merge_cells("A1:C1")
    for i, h in enumerate(["级别", "问题", "说明"], start=1):
        ws4.cell(row=4, column=i, value=h)
    style_header_row(ws4, 4, 3)
    for i, item in enumerate(analysis["anomalies"]):
        r = 5 + i
        level = {"high": "高", "warn": "中", "info": "提示"}.get(item["level"], item["level"])
        ws4.cell(row=r, column=1, value=level).font = Font(name=FONT, bold=True,
                                                          color=RED if item["level"] == "high" else "B45309")
        ws4.cell(row=r, column=2, value=item["title"]).font = Font(name=FONT)
        ws4.cell(row=r, column=3, value=item["detail"]).font = Font(name=FONT)
        ws4.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, 4):
            ws4.cell(row=r, column=c).border = border
    set_widths(ws4, [8, 34, 70])

    # 数据体检
    ws5 = wb.create_sheet("数据体检")
    ws5["A1"] = "数据体检"
    ws5["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws5.merge_cells("A1:D1")
    q = analysis["quality"]
    for i, h in enumerate(["级别", "问题", "说明"], start=1):
        ws5.cell(row=3, column=i, value=h)
    style_header_row(ws5, 3, 3)
    r = 4
    for issue in q["issues"]:
        level = {"high": "高", "warn": "中", "info": "提示"}.get(issue["level"], issue["level"])
        ws5.cell(row=r, column=1, value=level).font = Font(name=FONT)
        ws5.cell(row=r, column=2, value=issue["title"]).font = Font(name=FONT)
        ws5.cell(row=r, column=3, value=issue["detail"]).font = Font(name=FONT)
        ws5.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, 4):
            ws5.cell(row=r, column=c).border = border
        r += 1
    if not q["issues"]:
        ws5.cell(row=r, column=1, value="提示").font = Font(name=FONT)
        ws5.cell(row=r, column=2, value="未发现明显数据质量问题").font = Font(name=FONT)
        r += 1
    r += 1
    ws5.cell(row=r, column=1, value="列画像").font = Font(name=FONT, bold=True, color=NAVY)
    r += 1
    for i, h in enumerate(["列名", "类型", "非空率", "示例", "单位"], start=1):
        ws5.cell(row=r, column=i, value=h)
    style_header_row(ws5, r, 5)
    for col in q["columns"]:
        r += 1
        ws5.cell(row=r, column=1, value=col["name"]).font = Font(name=FONT)
        ws5.cell(row=r, column=2, value={"number": "数值", "date": "日期", "text": "文本", "empty": "空"}.get(col["type"], col["type"]))
        ws5.cell(row=r, column=3, value=1 - col["missingRatio"])
        ws5.cell(row=r, column=3).number_format = "0.0%"
        ws5.cell(row=r, column=4, value="；".join(col["samples"][:3])).font = Font(name=FONT, size=9)
        ws5.cell(row=r, column=5, value=col.get("unit", "")).font = Font(name=FONT, size=9)
    set_widths(ws5, [22, 10, 10, 46, 10])

    # 口径卡
    ws6 = wb.create_sheet("口径卡")
    ws6["A1"] = "分析口径"
    ws6["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws6.merge_cells("A1:C1")
    for i, h in enumerate(["字段", "映射列", "说明"], start=1):
        ws6.cell(row=3, column=i, value=h)
    style_header_row(ws6, 3, 3)
    labels = [("date", "日期"), ("revenue", "营业收入"), ("cost", "营业成本"), ("expense", "期间费用"), ("dimension", "分析维度")]
    notes = {"date": "用于生成月度趋势；无法识别的归入「全部」", "revenue": "必填；按数值原值汇总",
             "cost": "毛利 = 收入 - 成本", "expense": "净利润 = 毛利 - 费用",
             "dimension": "用于收入结构与异常分析"}
    for i, (key, label) in enumerate(labels):
        r = 4 + i
        ws6.cell(row=r, column=1, value=label).font = Font(name=FONT, bold=True)
        ws6.cell(row=r, column=2, value=analysis["mapping"].get(key) or "（未使用）").font = Font(name=FONT)
        ws6.cell(row=r, column=3, value=notes[key]).font = Font(name=FONT)
        ws6.cell(row=r, column=3).alignment = Alignment(wrap_text=True)
        for c in range(1, 4):
            ws6.cell(row=r, column=c).border = border
    r = 10
    unit = analysis["quality"].get("unit") or "数值原值"
    ws6.cell(row=r, column=1, value="金额单位").font = Font(name=FONT, bold=True)
    ws6.cell(row=r, column=2, value=unit).font = Font(name=FONT)
    ws6.cell(row=r, column=3, value="报告中的金额与源表保持同一单位。").font = Font(name=FONT)
    set_widths(ws6, [14, 26, 56])

    if compare:
        ws7 = wb.create_sheet("复跑对比")
        ws7["A1"] = "复跑对比"
        ws7["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
        ws7.merge_cells("A1:D1")
        for i, h in enumerate(["指标", "上次", "本次", "变化"], start=1):
            ws7.cell(row=3, column=i, value=h)
        style_header_row(ws7, 3, 4)
        rows = [("营业收入", compare["previous"]["revenue"], kpis["revenue"], False),
                ("毛利率", compare["previous"]["grossMargin"], kpis["grossMargin"], True),
                ("净利润", compare["previous"]["net"], kpis["net"], False),
                ("净利率", compare["previous"]["netMargin"], kpis["netMargin"], True)]
        for i, (name, before, after, pct) in enumerate(rows):
            r = 4 + i
            ws7.cell(row=r, column=1, value=name).font = Font(name=FONT)
            for col, value in ((2, before), (3, after)):
                cell = ws7.cell(row=r, column=col, value=value)
                if pct:
                    pct_cell(cell)
                else:
                    money_cell(cell)
            change = (after - before) / abs(before) if before else None
            cell = ws7.cell(row=r, column=4, value=change if change is not None else "—")
            if change is not None:
                pct_cell(cell)
            for c in range(1, 5):
                ws7.cell(row=r, column=c).border = border
        set_widths(ws7, [18, 18, 18, 14])

    # 明细
    ws8 = wb.create_sheet("原始明细")
    headers = list(analysis["headers"]) + ["期间（标准化）", "营业收入（标准化）", "营业成本（标准化）", "费用（标准化）"]
    for i, h in enumerate(headers, start=1):
        ws8.cell(row=1, column=i, value=h)
    style_header_row(ws8, 1, len(headers))
    records = analysis.get("_records") or []
    for i, rec in enumerate(records[:DETAIL_SHEET_LIMIT]):
        r = 2 + i
        for c, value in enumerate(rec["row"], start=1):
            cell = ws8.cell(row=r, column=c, value=value)
            cell.font = Font(name=FONT, size=9)
        offset = len(analysis["headers"])
        ws8.cell(row=r, column=offset + 1, value=rec["period"]).font = Font(name=FONT, size=9)
        for c, key in ((2, "revenue"), (3, "cost"), (4, "expense")):
            cell = ws8.cell(row=r, column=offset + c, value=rec[key])
            cell.number_format = '#,##0.00'
            cell.font = Font(name=FONT, size=9)
    set_widths(ws8, [14] * len(headers))

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)


def make_docx(analysis, out_path, compare=None):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    style.paragraph_format.line_spacing = 1.35
    style.paragraph_format.space_after = Pt(4)

    def shade(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def heading(text, size=13, color=NAVY):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(size)
        run.font.name = FONT
        run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        run.font.color.rgb = RGBColor.from_string(color)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

    def table(headers, rows, widths=None, number_cols=None):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(headers):
            cell = t.rows[0].cells[i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(h)
            run.bold = True
            run.font.size = Pt(9)
            run.font.name = FONT
            run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shade(cell, NAVY)
        for raw in rows:
            cells = t.add_row().cells
            for i, value in enumerate(raw):
                cells[i].text = ""
                p = cells[i].paragraphs[0]
                run = p.add_run(value if isinstance(value, str) else str(value))
                run.font.size = Pt(9)
                run.font.name = FONT
                run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
                if isinstance(value, str) and (value.startswith("-") or "下降" in value):
                    run.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
        if widths:
            for row in t.rows:
                for i, width in enumerate(widths):
                    row.cells[i].width = Cm(width)

    kpis = analysis["kpis"]
    period = ("%s 至 %s，共 %d 个月" % (kpis["periodStart"], kpis["periodEnd"], kpis["months"])) if kpis["months"] else "全部数据"

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run("经营分析报告")
    run.bold = True
    run.font.size = Pt(22)
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.color.rgb = RGBColor.from_string(NAVY)

    meta = doc.add_paragraph()
    meta_run = meta.add_run("数据源：%s（%s）｜分析期间：%s｜记录数：%d 行｜生成时间：%s" % (
        analysis["fileName"], analysis["sheet"], period, kpis["records"],
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    meta_run.font.size = Pt(9)
    meta_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    meta_run.font.name = FONT
    meta_run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    def fmt_money(v):
        return "{:,.2f}".format(v)

    def fmt_pct(v):
        return "—" if v is None else "%.1f%%" % (v * 100)

    NUMERALS = ["一", "二", "三", "四", "五", "六", "七"]
    section_index = [1]

    def section(text):
        section_index[0] += 1
        return heading("%s、%s" % (NUMERALS[section_index[0] - 1], text))

    heading("一、结论摘要")
    gross_margin = fmt_pct(kpis["grossMargin"])
    net_margin = fmt_pct(kpis["netMargin"])
    direction = "增长" if kpis.get("trendDirection", 0) >= 0 else "回落"
    summary = (
        "本期营业收入 %s，毛利 %s（毛利率 %s），净利润 %s（净利率 %s）。"
        "收入整体呈%s态势，数据体检发现 %d 条提示信号。"
    ) % (fmt_money(kpis["revenue"]), fmt_money(kpis["gross"]), gross_margin,
         fmt_money(kpis["net"]), net_margin, direction, len(analysis["anomalies"]))
    p = doc.add_paragraph(summary)
    p.paragraph_format.first_line_indent = Pt(21)

    section("关键指标")
    table(["指标", "数值", "说明"], [
        ["营业收入", fmt_money(kpis["revenue"]), "按口径卡汇总"],
        ["营业成本", fmt_money(kpis["cost"]), "按口径卡汇总"],
        ["毛利", fmt_money(kpis["gross"]), "收入 - 成本"],
        ["毛利率", gross_margin, "毛利 / 收入"],
        ["期间费用", fmt_money(kpis["expense"]), "营销、人工等"],
        ["净利润", fmt_money(kpis["net"]), "毛利 - 费用"],
        ["净利率", net_margin, "净利 / 收入"],
    ], widths=[4.0, 4.0, 8.0])

    section("趋势观察")
    trend = analysis["trend"]
    if kpis["months"]:
        rows = []
        for t in trend:
            rows.append([t["period"], fmt_money(t["revenue"]), fmt_pct(t["margin"]),
                         fmt_money(t["net"]), "—" if t.get("mom") is None else "%+.1f%%" % (t["mom"] * 100)])
        table(["期间", "营业收入", "毛利率", "净利润", "收入环比"], rows,
              widths=[3.2, 4.4, 3.2, 4.4, 3.4])
        best = max(trend, key=lambda t: t["revenue"])
        worst = min(trend, key=lambda t: t["revenue"])
        doc.add_paragraph("收入最高的期间为 %s（%s），最低为 %s（%s）。" % (
            best["period"], fmt_money(best["revenue"]), worst["period"], fmt_money(worst["revenue"])))
    else:
        doc.add_paragraph("数据中没有可识别的日期列，因此未生成月度趋势；建议补充日期字段后重新分析。")

    section("收入结构")
    items = analysis["structure"]["items"][:5]
    rows = [[g["name"], fmt_money(g["revenue"]), fmt_pct(g["share"]), fmt_pct(g["margin"])] for g in items]
    table(["维度（Top5）", "营业收入", "收入占比", "毛利率"], rows,
          widths=[6.0, 4.6, 3.6, 3.6])
    if items:
        doc.add_paragraph("「%s」是最大收入来源，占 %.1f%%；建议把它作为下期重点复盘对象。" % (
            items[0]["name"], items[0]["share"] * 100))

    section("异常与建议")
    for item in analysis["anomalies"][:6]:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(item["title"])
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = FONT
        run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        p.add_run("：" + item["detail"]).font.size = Pt(10)
    if compare:
        section("复跑对比")
        prev = compare["previous"]
        rows = [
            ["营业收入", fmt_money(prev["revenue"]), fmt_money(kpis["revenue"]),
             "—" if not prev["revenue"] else "%+.1f%%" % ((kpis["revenue"] / prev["revenue"] - 1) * 100)],
            ["毛利率", fmt_pct(prev["grossMargin"]), fmt_pct(kpis["grossMargin"]),
             "—" if prev["grossMargin"] is None else "%+.1f pp" % ((kpis["grossMargin"] - prev["grossMargin"]) * 100)],
            ["净利润", fmt_money(prev["net"]), fmt_money(kpis["net"]),
             "—" if not prev["net"] else "%+.1f%%" % ((kpis["net"] / prev["net"] - 1) * 100)],
        ]
        table(["指标", "上次", "本次", "变化"], rows, widths=[4.0, 4.0, 4.0, 3.8])
        doc.add_paragraph("对比基准：%s 的分析结果。" % prev.get("generatedAt", "上次运行"))

    section("口径与说明")
    mapping = analysis["mapping"]
    unit = analysis["quality"].get("unit") or "与源表一致"
    items_text = [
        "日期列：%s" % (mapping.get("date") or "未使用"),
        "营业收入列：%s" % (mapping.get("revenue") or "未指定"),
        "营业成本列：%s" % (mapping.get("cost") or "未使用"),
        "期间费用列：%s" % (mapping.get("expense") or "未使用"),
        "分析维度：%s" % (mapping.get("dimension") or "未使用"),
        "金额单位：%s" % unit,
        "分析记录：%d 行有效数据；合计行、缺少金额的行不参与计算。" % kpis["records"],
    ]
    for text in items_text:
        p = doc.add_paragraph(text, style="List Bullet")
        for r in p.runs:
            r.font.size = Pt(9)
            r.font.name = FONT
            r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    doc.save(out_path)


def make_report(payload):
    analysis = analyze_payload(payload)
    analysis["_records"] = pick_rows(load_table(payload["path"], payload.get("sheet")), analysis["mapping"])[0]
    out_dir = Path(payload["outDir"])
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    stem = re.sub(r'[\\/:*?"<>|]+', "_", Path(analysis["fileName"]).stem)
    base = "%s_经营分析_%s" % (stem, stamp)
    xlsx_path = out_dir / (base + ".xlsx")
    docx_path = out_dir / (base + ".docx")
    compare = payload.get("compare")
    make_xlsx(analysis, str(xlsx_path), compare)
    make_docx(analysis, str(docx_path), compare)
    analysis.pop("_records", None)
    return {
        "analysis": analysis,
        "files": [
            {"kind": "xlsx", "path": str(xlsx_path), "bytes": xlsx_path.stat().st_size},
            {"kind": "docx", "path": str(docx_path), "bytes": docx_path.stat().st_size},
        ],
        "generatedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ---------------------------------------------------------------- 示例数据

def make_sample(payload):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    out = Path(payload["out"])
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "经营明细"
    headers = ["月份", "渠道", "营业收入", "营业成本", "营销费用", "订单数"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
    channels = ["天猫", "抖音", "门店", "私域"]
    weights = {"天猫": 0.42, "抖音": 0.3, "门店": 0.18, "私域": 0.1}
    base = 180000.0
    rows = 0
    for m in range(1, 13):
        season = 1.0 + 0.12 * math.sin((m - 2) / 12 * 2 * math.pi)
        growth = 1.0 + 0.025 * (m - 1)
        for ch in channels:
            revenue = base * weights[ch] * growth * season
            if m == 8 and ch == "抖音":
                revenue *= 0.62  # 平台规则调整，收入明显回落
            if m == 11 and ch == "天猫":
                revenue *= 1.35  # 大促
            cost = revenue * (0.55 if ch != "门店" else 0.62)
            expense = base * weights[ch] * 0.14
            orders = int(revenue / (180 if ch == "私域" else 120))
            if m == 4 and ch == "门店":
                expense *= 2.1  # 装修一次性投入
            ws.append([f"2026-{m:02d}", ch, round(revenue, 2), round(cost, 2), round(expense, 2), orders])
            rows += 1
    ws.append(["合计", "", ""] + [""] * 3)
    for width, col in zip([10, 8, 14, 14, 14, 8], "ABCDEF"):
        ws.column_dimensions[col].width = width
    wb.save(out)
    return {"path": str(out), "rows": rows, "sheet": "经营明细"}


def payload_from_stdin():
    raw = sys.stdin.buffer.read().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    return json.loads(raw)


def dispatch():
    if len(sys.argv) < 2:
        raise ValueError("缺少命令")
    command = sys.argv[1]
    payload = payload_from_stdin()
    if True:
        if command == "inspect":
            table = load_table(payload["path"], payload.get("sheet"))
            profile = build_profile(table)
            mapping = suggest_mapping(table["headers"])
            return {
                "path": table["path"],
                "fileName": Path(table["path"]).name,
                "sheet": table["sheet"],
                "sheetNames": table["sheetNames"],
                "rowCount": len(table["rows"]),
                "headerRow": table["headerRow"],
                "columns": profile["columns"],
                "issues": profile["issues"],
                "suggestedMapping": mapping,
                "sampleRows": [[norm_cell(c) for c in r] for r in table["rows"][:8]],
            }
        if command == "analyze":
            analysis = analyze_payload(payload)
            headers = analysis.pop("headers", None)
            result = dict(analysis)
            result["headers"] = headers
            return result
        if command == "report":
            return make_report(payload)
        if command == "sample":
            return make_sample(payload)
        if command == "selftest":
            sample_dir = Path(payload.get("dir") or ".")
            sample = make_sample({"out": str(sample_dir / "示例经营数据.xlsx")})
            analysis = analyze_payload({"path": sample["path"]})
            report = make_report({"path": sample["path"], "outDir": payload["outDir"]})
            return {"sample": sample, "analysis": analysis, "report": report}
        raise ValueError("未知命令：%s" % command)


def main():
    try:
        print(json.dumps(dispatch(), ensure_ascii=False, default=str))
    except Exception as error:
        print(json.dumps({"error": str(error), "trace": traceback.format_exc().splitlines()[-4:]}, ensure_ascii=False))


if __name__ == "__main__":
    main()

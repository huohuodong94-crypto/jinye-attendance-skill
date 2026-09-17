#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
锦业体育考勤表生成脚本 v2

数据源（优先级从高到低）：
  1. 月度考勤汇总 第26-56列 = 逐日考勤结果（权威，含打卡时间/状态/补卡/缺卡/迟到/早退/事假/病假/出差/外出）
  2. 请假表 = 请假具体类型（事由含"调休"→调休假）
  3. 打卡记录表/出差表/外出表 = 辅助（汇总缺数据时兜底）

输出：标准格式考勤表Excel（3个工作表：考勤一览表、陈江办公室、公庄办公室）

用法：
    python3 generate_attendance.py \
        --punch 打卡记录表.xlsx \
        --summary 月度考勤汇总.xlsx \
        --leave 请假.xlsx \
        --travel 出差.xlsx \
        --outside 外出.xlsx \
        --month 7 \
        --output 7月考勤表.xlsx
"""

import argparse
import re
import math
import calendar
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ==================== 常量配置 ====================

# 标准工作时间（按季节切换）
# 5-9月（夏令时）：上午08:30-12:00，下午14:00-18:00
# 10月-次年4月（冬令时）：上午08:30-12:00，下午13:30-17:30
WORK_START_AM = 8 * 60 + 30   # 08:30
WORK_END_AM = 12 * 60          # 12:00

def get_work_hours(month):
    """根据月份返回下午工作时间 (start_min, end_min, std_day_hours)"""
    if 5 <= month <= 9:
        return 14 * 60, 18 * 60, 7.5   # 夏令时
    else:
        return 13 * 60 + 30, 17 * 60 + 30, 7.5  # 冬令时

STD_DAY_HOURS = 7.5            # 每天标准工时

# 带薪假类型
PAID_LEAVES = {"婚假", "产假", "陪产假", "法定节假日"}

# 颜色
COLOR_RED = "FFFF0000"
COLOR_BLUE = "FF0070C0"
COLOR_GREEN = "FF00B050"
COLOR_ORANGE = "FFFFA500"      # 外出（橙色）
COLOR_PURPLE = "FF7030A0"

# 明细打卡块员工（按办公室）
OFFICE_EMPLOYEES = {
    "公庄办公室": ["邱惠浓", "谢铃铃", "钟慧婷", "王文敏", "李健容", "张小萍"],
    "陈江办公室": [
        "罗玉珍", "崔怡", "彭宏", "孙誉婕", "许凡", "翟海浩",
        "李岭恩（离职）", "杨建", "郁魏", "刘致君", "马杨萍", "余唯",
    ],
}

# 计时工（按总工时计薪，不按天考核）：{姓名: 总工时小时数}
HOURS_WORKERS = {"黄志英": 81}

# 非打卡考勤人员（一览表备注"非打卡考勤"）
NON_PUNCH_EMPLOYEES = {"蔡广秀", "张雪", "乐志德", "毛高才"}

# 居家办公人员（一览表备注"居家办公"）
HOME_OFFICE_EMPLOYEES = {"杨俊"}

# 整月出差人员特殊备注：{姓名: 备注}
TRAVEL_MONTHLY_REMARK = {"曾奇峰": "驻福州申请整月出差"}

# 离职人员：{姓名: 离职日期}
LEAVE_COMPANY_DATE = {"李岭恩（离职）": 21}

# 台风等特殊标记：{员工: {日: 加班列标记}}（人工补充，汇总数据无法推断）
SPECIAL_MARKERS = {
    # 7月25日台风特殊标记（仅7月有效，其他月份留空）
    # "罗玉珍": {25: "台风提前下班"},
}


# ==================== 工具函数 ====================

def time_to_minutes(time_str):
    if not time_str or not isinstance(time_str, str):
        return None
    time_str = time_str.strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", time_str)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def get_weekday(year, month, day):
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    return weekdays[datetime(year, month, day).weekday()]


def is_sunday(year, month, day):
    return get_weekday(year, month, day) == "日"


def is_rest_day(year, month, day):
    """判断是否为休息日：7、8月6天制（仅周日休），其他月份双休（周六日休）"""
    wd = get_weekday(year, month, day)
    if month in (7, 8):
        return wd == "日"
    else:
        return wd in ("六", "日")


def get_should_days(year, month):
    """计算当月应出勤天数（扣除休息日）"""
    total = days_in_month(year, month)
    rest = sum(1 for d in range(1, total + 1) if is_rest_day(year, month, d))
    return total - rest


def days_in_month(year, month):
    if month == 12:
        return 31
    return (datetime(year, month + 1, 1) - timedelta(days=1)).day


def format_work_time(total_hours):
    """实际工时 → '27' 或 '26天4小时'"""
    days = int(total_hours // STD_DAY_HOURS)
    remain = round(total_hours % STD_DAY_HOURS, 1)
    if remain == 0:
        return str(days)
    remain_str = str(remain)
    if remain_str.endswith(".0"):
        remain_str = remain_str[:-2]
    return f"{days}天{remain_str}小时"


def format_leave_time(leave_hours, leave_type="事假"):
    """请假时长 → '事假4天3.5小时' / '病假1天' / '调休假3.5小时'"""
    days = int(leave_hours // STD_DAY_HOURS)
    remain = round(leave_hours % STD_DAY_HOURS, 1)
    remain_str = str(remain)
    if remain_str.endswith(".0"):
        remain_str = remain_str[:-2]
    if days > 0 and remain > 0:
        return f"{leave_type}{days}天{remain_str}小时"
    if days > 0:
        return f"{leave_type}{days}天"
    return f"{leave_type}{remain_str}小时"


def safe_write_cell(ws, row, col, value, font=None):
    """安全写入单元格，处理合并单元格"""
    cell = ws.cell(row, col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
    cell.value = value
    if font:
        cell.font = font
    return cell


# ==================== 数据解析 ====================

def parse_summary(filepath, year, month):
    """
    解析月度考勤汇总：
    - 员工基础信息（部门/职位/考勤组/缺卡次数/工作时长）
    - 逐日考勤结果（第26-56列）
    返回 {员工: {info: {...}, daily: {日: (描述, [4个时间或None])}}}
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    result = {}

    for r in range(5, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        name = str(name).strip()
        daily = {}
        for c in range(26, 57):
            day = c - 25
            if day > 31:
                break
            val = ws.cell(r, c).value
            if val is None:
                continue
            daily[day] = parse_result_cell(val)

        result[name] = {
            "info": {
                "group": str(ws.cell(r, 2).value or "").strip(),
                "department": str(ws.cell(r, 3).value or "").strip(),
                "position": str(ws.cell(r, 5).value or "").strip(),
                "miss_in": ws.cell(r, 17).value,
                "miss_out": ws.cell(r, 18).value,
                "work_minutes": ws.cell(r, 9).value,
            },
            "daily": daily,
        }

    wb.close()
    return result


def parse_result_cell(text):
    """
    解析考勤结果单元格
    '正常,补卡申请07-11 14:00到07-11 14:00\n(08:22,12:00,14:00,18:00)'
    → ('正常,补卡申请...', ['08:22','12:00','14:00','18:00'])
    '-' → None（缺卡）
    """
    lines = str(text).split("\n")
    desc = lines[0].strip() if lines else ""
    times_raw = lines[1] if len(lines) > 1 else ""
    times = [None, None, None, None]
    m = re.search(r"\((.*)\)", times_raw)
    if m:
        parts = [p.strip() for p in m.group(1).split(",")]
        for i, p in enumerate(parts[:4]):
            if re.match(r"^\d{1,2}:\d{2}$", p):
                times[i] = p
    return desc, times


def parse_leave_types(filepath, year, month):
    """
    解析请假表，获取员工-日期 → 请假类型（事由含"调休"→"调休假"）
    去重：审批通过的优先，内容完全相同的只保留1条
    返回 {员工: {日: 请假类型}}
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    result = {}
    seen = set()

    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 11).value  # 创建人
        start = ws.cell(r, 4).value
        end = ws.cell(r, 5).value
        reason = ws.cell(r, 7).value
        status = ws.cell(r, 14).value  # 审批状态

        if not name or not start:
            continue
        name = str(name).strip()
        reason = str(reason or "").strip()
        status = str(status or "").strip()

        # 只处理审批通过的记录
        if status and status != "已结束" and "通过" not in status:
            continue

        start_dt = parse_datetime(start)
        end_dt = parse_datetime(end)
        if not start_dt or not end_dt:
            continue

        # 跨月过滤
        if start_dt.month != month and end_dt.month != month:
            continue

        # 去重：员工+事由+开始时间
        key = (name, reason, str(start), str(end))
        if key in seen:
            continue
        seen.add(key)

        # 请假类型
        leave_type = "事假"
        if "调休" in reason:
            leave_type = "调休假"
        elif "病假" in reason or "生病" in reason:
            leave_type = "病假"
        elif "婚" in reason:
            leave_type = "婚假"
        elif "产" in reason:
            leave_type = "产假"

        # 按天展开
        current = start_dt.replace(hour=0, minute=0, second=0)
        end_day = end_dt.replace(hour=0, minute=0, second=0)
        while current <= end_day:
            if current.month == month:
                if name not in result:
                    result[name] = {}
                result[name][current.day] = leave_type
            current += timedelta(days=1)

    wb.close()
    return result


def days_in_month(year, month):
    return calendar.monthrange(year, month)[1]


def is_full_month_travel(desc_marker, month, year):
    """
    判断出差申请是否覆盖整个考勤月（如 '出差07-01 08:30到07-31 18:00 31天'）。
    整月出差申请视为长期挂单，不作为当天出差处理（员工实际按打卡/外出出勤）。
    """
    m = re.search(r"(\d{2})-(\d{2}) \d{1,2}:\d{2}到(\d{2})-(\d{2})", desc_marker)
    if not m:
        return False
    sm, sd = int(m.group(1)), int(m.group(2))
    em, ed = int(m.group(3)), int(m.group(4))
    last_day = days_in_month(year, month)
    return sm == month and sd == 1 and em == month and ed == last_day


def parse_datetime(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in ["%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def extract_cover_times(desc_marker):
    """
    从 '外出07-09 09:00到07-09 18:00 7小时' 提取覆盖起止分钟
    返回 (start_minutes, end_minutes) 或 (None, None)
    """
    m = re.search(r"(\d{1,2}):(\d{2})到.*?(\d{1,2}):(\d{2})", desc_marker)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2)), int(m.group(3)) * 60 + int(m.group(4))
    return None, None


def adjust_cover_for_day(desc_marker, day, month, cs, ce):
    """
    根据当天在跨天申请中的位置调整覆盖起止时间：
      - 起始日：从实际开始时间到当天结束（24:00）
      - 结束日：从当天开始（00:00）到实际结束时间
      - 中间日：全天覆盖（00:00-24:00）
      - 单日申请：保持原始时间
    """
    m = re.search(r"(\d{2})-(\d{2}) (\d{1,2}):(\d{2})到(\d{2})-(\d{2}) (\d{1,2}):(\d{2})", desc_marker)
    if not m:
        return cs, ce
    start = (int(m.group(1)), int(m.group(2)))
    end = (int(m.group(5)), int(m.group(6)))
    start_min = int(m.group(3)) * 60 + int(m.group(4))
    end_min = int(m.group(7)) * 60 + int(m.group(8))
    today = (month, day)
    if start != end:
        if today == start:
            return start_min, 24 * 60
        if today == end:
            return 0, end_min
        return 0, 24 * 60
    return cs, ce


def extract_period(desc_marker, day=None, month=None):
    """
    从 '事假07-04 08:30到07-04 12:00 3.5小时' / '出差07-06 08:30到07-10 18:00 5天'
    提取当天被覆盖的时段：'上午' / '下午' / '全天'
    对跨天申请，结合当天日期判断当天是起始日/结束日/中间日：
      - 起始日：从开始时间起覆盖（开始<12→全天，开始>=13→下午）
      - 结束日：到结束时间止（结束>=14→全天，结束<=12→上午）
      - 中间日：全天
    """
    m = re.search(r"(\d{2})-(\d{2}) (\d{1,2}):(\d{2})到(\d{2})-(\d{2}) (\d{1,2}):(\d{2})", desc_marker)
    if m:
        start_m, start_d = int(m.group(1)), int(m.group(2))
        end_m, end_d = int(m.group(5)), int(m.group(6))
        start_h = int(m.group(3))
        end_h = int(m.group(7))

        if day is not None and month is not None:
            # 单日申请
            if (start_m, start_d) == (end_m, end_d) == (month, day):
                if start_h < 12 and end_h <= 12:
                    return "上午"
                if start_h >= 13 and end_h >= 12:
                    return "下午"
                return "全天"
            # 起始日
            if (start_m, start_d) == (month, day):
                return "全天" if start_h < 12 else "下午"
            # 结束日
            if (end_m, end_d) == (month, day):
                return "全天" if end_h >= 14 else "上午"
            # 中间日
            return "全天"

        # 无当天日期信息：起止日期不同→全天
        if (m.group(1), m.group(2)) != (m.group(5), m.group(6)):
            return "全天"
        if start_h < 12 and end_h <= 12:
            return "上午"
        if start_h >= 13 and end_h >= 12:
            return "下午"
        return "全天"
    else:
        m2 = re.search(r"(\d{1,2}):(\d{2})到.*?(\d{1,2}):(\d{2})", desc_marker)
        if not m2:
            return "全天"
        start_h = int(m2.group(1))
        end_h = int(m2.group(3))
    if start_h < 12 and end_h <= 12:
        return "上午"
    if start_h >= 13 and end_h >= 12:
        return "下午"
    return "全天"


def calc_daily_leave_hours(desc_marker, day, month):
    """
    规则8.8：计算部分请假当天实际请假小时数（与工作时间段的交集）
    从 '事假08-17 11:00到08-19 18:00 19小时' 提取完整起止时间，计算与工作时间的精确交集
    """
    m = re.search(r"(\d{2})-(\d{2}) (\d{1,2}):(\d{2})到(\d{2})-(\d{2}) (\d{1,2}):(\d{2})", desc_marker)
    if not m:
        return 4.0  # 默认半天

    start_m, start_d = int(m.group(1)), int(m.group(2))
    end_m, end_d = int(m.group(5)), int(m.group(6))
    start_h, start_min = int(m.group(3)), int(m.group(4))
    end_h, end_min = int(m.group(7)), int(m.group(8))

    # 获取当天工作时间
    pm_start, pm_end, std_hours = get_work_hours(month)
    am_start = WORK_START_AM  # 08:30
    am_end = WORK_END_AM      # 12:00

    # 判断当天在请假段中的位置
    today = (month, day)
    start_date = (start_m, start_d)
    end_date = (end_m, end_d)

    # 计算当天的请假起止时间
    if start_date == today and end_date == today:
        # 同一天：提取完整起止时间
        leave_start = start_h * 60 + start_min
        leave_end = end_h * 60 + end_min
    elif start_date == today:
        # 起始日：从开始时间到当天工作结束
        leave_start = start_h * 60 + start_min
        leave_end = pm_end  # 到下午下班
    elif end_date == today:
        # 结束日：从当天工作开始到结束时间
        leave_start = am_start  # 从上午上班开始
        leave_end = end_h * 60 + end_min
    else:
        # 中间日（全天）
        return std_hours

    # 计算与上午工作时间的交集
    am_overlap = max(0, min(leave_end, am_end) - max(leave_start, am_start))
    # 计算与下午工作时间的交集
    pm_overlap = max(0, min(leave_end, pm_end) - max(leave_start, pm_start))

    # 转换为小时
    total_minutes = am_overlap + pm_overlap
    return round(total_minutes / 60, 1)


def fmt_minutes(m):
    """分钟数 → 'HH:MM'"""
    return f"{m // 60:02d}:{m % 60:02d}"


def calc_daily_leave_range(desc_marker, day, month):
    """
    按天拆分请假时段：返回当天请假起止时间 'HH:MM-HH:MM'（以请假表申请时间为准）。
    - 同一天：申请原始起止（如 14:00-18:00）
    - 起始日：申请开始时间 → 当天下班
    - 结束日：当天上班 → 申请结束时间
    - 中间日：全天工作时间段
    描述中无起止时间时返回 None。仅用于异常说明括注，不影响工时计算。
    """
    m = re.search(r"(\d{2})-(\d{2}) (\d{1,2}):(\d{2})到(\d{2})-(\d{2}) (\d{1,2}):(\d{2})", desc_marker)
    if not m:
        return None

    start_m, start_d = int(m.group(1)), int(m.group(2))
    end_m, end_d = int(m.group(5)), int(m.group(6))
    start_h, start_min = int(m.group(3)), int(m.group(4))
    end_h, end_min = int(m.group(7)), int(m.group(8))

    pm_start, pm_end, std_hours = get_work_hours(month)
    am_start = WORK_START_AM  # 08:30

    today = (month, day)
    start_date = (start_m, start_d)
    end_date = (end_m, end_d)

    if start_date == today and end_date == today:
        ls, le = start_h * 60 + start_min, end_h * 60 + end_min
    elif start_date == today:
        ls, le = start_h * 60 + start_min, pm_end
    elif end_date == today:
        ls, le = am_start, end_h * 60 + end_min
    else:
        ls, le = am_start, pm_end
    return f"{fmt_minutes(ls)}-{fmt_minutes(le)}"


# ==================== 考勤计算 ====================

def build_daily_entry(desc, times, emp_name, day, leave_map, year, month):
    """
    从考勤结果描述+时间构建每日考勤entry
    """
    entry = {
        "punches": times,       # [4个时间或None]
        "status": "正常",
        "leave_type": None,     # 请假类型（事假/病假/调休假）
        "leave_hours": 0.0,     # 当天请假小时数
        "leave_period": None,   # 半天请假时段
        "out_period": None,     # 半天外出时段
        "travel_period": None,  # 半天出差时段
        "is_late": False,
        "is_early": False,
        "miss_pos": [],         # 缺卡格位置 [0,1,2,3]
        "has_makeup": False,
        "marker": None,         # 加班列标记
    }

    markers = [m.strip() for m in desc.split(",") if m.strip()]

    # 离职（仅"离职"标记，旷工不算离职）
    if any("离职" in m for m in markers):
        entry["status"] = "离职"
        return entry

    # 旷工：无打卡，4格留空，不计工时
    if any("旷工" in m for m in markers):
        entry["status"] = "旷工"
        return entry

    # 统一扫描迟到/早退/缺卡/补卡标记（所有分支共用）
    # 迟到仅"上班1迟到"（"上班2迟到"如李健容8不标）
    entry["official_miss"] = set()
    for m in markers:
        if "迟到" in m and ("上班1迟到" in m or m.startswith("上班迟到")):
            entry["is_late"] = True
        if "早退" in m:
            entry["is_early"] = True
        if "缺卡" in m:
            if "上班1缺卡" in m or ("上班缺卡" in m and "1" in m):
                entry["miss_pos"].append(0)
                entry["official_miss"].add(0)
            if "下班1缺卡" in m:
                entry["miss_pos"].append(1)
                entry["official_miss"].add(1)
            if "上班2缺卡" in m:
                entry["miss_pos"].append(2)
                entry["official_miss"].add(2)
            if "下班2缺卡" in m:
                entry["miss_pos"].append(3)
                entry["official_miss"].add(3)
        if "补卡申请" in m:
            entry["has_makeup"] = True

    # 下午上班迟到：第3次打卡晚于下午上班时间（按季节：5-9月14:00，10-次年4月13:30）
    pm_start, _, _ = get_work_hours(month)
    if len(times) >= 3 and times[2] and times[2] != "-":
        t3 = time_to_minutes(times[2])
        if t3 and t3 > pm_start:
            entry["is_late"] = True

    # 外出优先于出差（同一标记串中外出通常是当天具体活动，如整月出差+当天外出）
    outside_markers = [m for m in markers if "外出" in m]
    if outside_markers:
        # 多条外出记录合并时段
        periods = [extract_period(m, day, month) for m in outside_markers]
        if "全天" in periods or ("上午" in periods and "下午" in periods):
            period = "全天"
        else:
            period = periods[0]
        cs, ce = extract_cover_times(outside_markers[0])
        cs, ce = adjust_cover_for_day(outside_markers[0], day, month, cs, ce)
        entry["cover_start"], entry["cover_end"] = cs, ce
        if period == "全天":
            entry["status"] = "外出"
        else:
            entry["status"] = "部分外出"
            entry["out_period"] = period
        # 计算外出总时长（用于加班列标注）
        outside_hours = 0
        for m in outside_markers:
            hm = re.search(r"([\d.]+)小时", m)
            if hm:
                outside_hours += float(hm.group(1))
        entry["outside_hours"] = outside_hours
        return entry

    # 短期出差（非整月申请）优先于休息（出差期间休息日也算出差）
    travel_markers = [m for m in markers if "出差" in m]
    short_travel = [m for m in travel_markers if not is_full_month_travel(m, month, year)]
    if short_travel:
        # 多条出差记录合并时段（如崔怡30=出差到13:15+出差13:15起→全天）
        periods = [extract_period(m, day, month) for m in short_travel]
        if "全天" in periods or ("上午" in periods and "下午" in periods):
            period = "全天"
        else:
            period = periods[0]
        cs, ce = extract_cover_times(short_travel[0])
        cs, ce = adjust_cover_for_day(short_travel[0], day, month, cs, ce)
        entry["cover_start"], entry["cover_end"] = cs, ce
        if period == "全天":
            entry["status"] = "出差"
        else:
            entry["status"] = "部分出差"
            entry["travel_period"] = period
        return entry

    # 休息（优先于请假判断，休息日显示"休息"）
    # 规则8.7：请假段内的休息日不断开请假段，休息日显示"休息"
    if is_rest_day(year, month, day) and not any("出差" in m for m in markers) and not any("外出" in m for m in markers):
        entry["status"] = "休息"
        return entry

    # 请假（事假/病假/调休假）
    leave_markers = [m for m in markers if "事假" in m or "病假" in m]
    if leave_markers:
        lm = leave_markers[0]
        hm = re.search(r"([\d.]+)小时", lm)
        period = extract_period(lm, day, month)
        hours = float(hm.group(1)) if hm else (7.5 if period == "全天" else 4.0)

        # 规则8.8：部分请假计算当天实际请假小时数（与工作时间段的交集）
        if period != "全天":
            hours = calc_daily_leave_hours(lm, day, month)

        # 请假类型（优先从请假表取，其次从描述判断）
        leave_type = "病假" if "病假" in lm else "事假"
        if emp_name in leave_map and day in leave_map[emp_name]:
            leave_type = leave_map[emp_name][day]

        entry["leave_type"] = leave_type
        entry["leave_hours"] = hours
        entry["leave_range"] = calc_daily_leave_range(lm, day, month)  # 当天申请时段，仅用于异常说明括注
        entry["leave_total_hours"] = float(hm.group(1)) if hm else (7.5 if period == "全天" else 4.0)  # 整段总时长
        entry["leave_desc"] = lm  # 请假标记，用于段首判断（如"事假08-03 13:57到08-05 18:00 19小时"）
        if period == "全天":
            entry["status"] = "请假"
        else:
            entry["status"] = "部分请假"
            entry["leave_period"] = period
            entry["marker"] = f"事假{hours}小时"
        return entry

    # 休息（出差/外出期间的休息日不算）
    if any("休息" in m for m in markers):
        entry["status"] = "休息"
        return entry

    # 整月出差申请（如马杨萍出差07-01到07-31）：不作为出差处理，落到正常日

    # 由括号时间推断缺卡位置（描述中未明确的）
    for i, t in enumerate(times):
        if t is None and i not in entry["miss_pos"]:
            # 根据其他打卡判断：有部分打卡时缺失的位置视为缺卡
            if any(x is not None for x in times):
                entry["miss_pos"].append(i)

    # marker优先级：迟到 > 补卡申请 > 台风（早退且当月25日）
    if entry["is_late"]:
        entry["marker"] = "迟到"
    elif entry["has_makeup"]:
        entry["marker"] = "补卡"
    elif entry["is_early"] and month == 7 and day == 25:
        entry["marker"] = "台风提前下班"

    return entry


def compute_attendance(emp_name, daily, year, month, miss_in_total, miss_out_total):
    """
    计算员工月度出勤汇总
    返回 {应出勤天数, 实际工作小时, 实出勤字符串, 异常说明, 各异常计数}
    """
    should_days = 0
    work_hours = 0.0
    leave_hours = 0.0
    leave_details = []  # [(day, 请假类型, 当天时段)]，用于异常说明括注
    travel_days = 0  # 全天出差
    outside_days = 0  # 全天外出
    partial_outside_info = []  # 部分外出信息列表：[(日期, 小时数)]
    late_count = 0
    early_count = 0
    leave_days = 0
    sick_days = 0
    leave_types = set()
    makeup_count = 0  # 补卡申请次数

    leave_company = LEAVE_COMPANY_DATE.get(emp_name)

    for day in range(1, days_in_month(year, month) + 1):
        # 规则8.9：休息日（周日）有出差/外出必须计入统计
        if is_rest_day(year, month, day):
            entry = daily.get(day)
            if entry and entry["status"] in ("出差", "部分出差", "外出", "部分外出"):
                work_hours += STD_DAY_HOURS
                if "出差" in entry["status"]:
                    travel_days += 1
                if "外出" in entry["status"]:
                    outside_days += 1
            continue
        # 应出勤天数始终按整月工作日计（模板口径：离职员工应出勤仍显示整月）
        should_days += 1
        # 离职后不计实际出勤
        if leave_company and day > leave_company:
            continue

        entry = daily.get(day)
        if entry is None:
            continue

        status = entry["status"]

        if status == "休息":
            continue
        if status == "离职":
            continue
        if status == "旷工":
            # 旷工日不计工时
            continue

        if status == "请假":
            # 带薪假计工时，不带薪不计
            if entry["leave_type"] in PAID_LEAVES:
                work_hours += STD_DAY_HOURS
            else:
                leave_hours += STD_DAY_HOURS
                leave_days += 1
                leave_types.add(entry["leave_type"] or "事假")
                leave_details.append((day, entry["leave_type"] or "事假", entry.get("leave_range")))
                if entry["leave_type"] == "病假":
                    sick_days += 1
        elif status == "部分请假":
            # 半天请假：实际工时 = 满勤 - 请假小时
            work_hours += STD_DAY_HOURS - entry["leave_hours"]
            leave_hours += entry["leave_hours"]
            leave_types.add(entry["leave_type"] or "事假")
            leave_details.append((day, entry["leave_type"] or "事假", entry.get("leave_range")))
        elif status in ("出差", "部分出差", "外出", "部分外出"):
            work_hours += STD_DAY_HOURS
            if "出差" in status:
                travel_days += 1
            if "外出" in status:
                if status == "外出":
                    outside_days += 1  # 只统计全天外出
                elif status == "部分外出":
                    partial_outside_info.append((day, entry.get("outside_hours", 0)))
        else:
            # 正常日：算满勤（迟到早退缺卡不扣）
            work_hours += STD_DAY_HOURS

        if entry.get("is_late"):
            late_count += 1
        if entry.get("is_early"):
            early_count += 1
        # 补卡申请：计数并分配序号（第几次补卡）
        if entry.get("has_makeup"):
            makeup_count += 1
            entry["makeup_index"] = makeup_count

    # 缺卡次数：从明细实际统计，排除外出/出差日的缺卡（与考勤表补卡格保持一致）
    miss_count = 0
    for e in daily.values():
        if e.get("status") in ("出差", "部分出差", "外出", "部分外出"):
            continue
        miss_count += len(e.get("miss_pos", []))

    # 实出勤字符串
    actual_str = format_work_time(work_hours)

    # 规则8.9：异常说明固定顺序
    abnormal_parts = []
    if late_count > 0:
        abnormal_parts.append(f"迟到{late_count}次")
    if early_count > 0:
        abnormal_parts.append(f"早退{early_count}次")
    if miss_count > 0:
        abnormal_parts.append(f"缺卡{miss_count}次")
    if leave_hours > 0:
        # 请假类型：单一类型用该类型，混合用"请假"
        lt = leave_types.pop() if len(leave_types) == 1 else "请假"
        leave_item = format_leave_time(leave_hours, lt)
        # 新规则(2026-09-16客户确认)：仅明细表异常说明括注每天请假明细与时间段，一览表备注保持汇总口径
        paren = ""
        if leave_details:
            leave_details.sort(key=lambda x: x[0])
            single_type = len({t for _, t, _ in leave_details}) == 1
            frags = []
            for d, t, rng in leave_details:
                head = f"{month}.{d}" if single_type else f"{t}{month}.{d}"
                frags.append(f"{head} {rng}" if rng else head)
            paren = "（" + "；".join(frags) + "）"
        leave_item_idx = len(abnormal_parts)
        abnormal_parts.append(leave_item)
        if paren:
            _detail_paren = (leave_item_idx, paren)
        else:
            _detail_paren = None
    if travel_days > 0:
        abnormal_parts.append(f"出差{travel_days}天")
    if outside_days > 0:
        abnormal_parts.append(f"外出{outside_days}天")
    if partial_outside_info:
        for day, hours in partial_outside_info:
            abnormal_parts.append(f"部分外出{hours}小时（{month}.{day}）")
    # 规则8.9：补卡不再出现在异常说明中

    # 明细表版：在请假项后括注每日请假明细（一览表不用）
    try:
        _detail_paren
    except NameError:
        _detail_paren = None
    if _detail_paren:
        idx, paren = _detail_paren
        detail_parts = list(abnormal_parts)
        detail_parts[idx] = detail_parts[idx] + paren
        abnormal_detail_str = "，".join(detail_parts)
    else:
        abnormal_detail_str = "，".join(abnormal_parts) if abnormal_parts else ""

    return {
        "should_days": should_days,
        "actual_str": actual_str,
        "abnormal_summary": "，".join(abnormal_parts) if abnormal_parts else "",
        "abnormal_detail": abnormal_detail_str,
        "work_hours": work_hours,
        "leave_hours": leave_hours,
        "miss_count": miss_count,
        "late_count": late_count,
        "early_count": early_count,
        "travel_days": travel_days,
        "outside_days": outside_days,
        "makeup_count": makeup_count,
        "sick_days": sick_days,
        "absent_days": miss_count,  # 缺卡次数
        "full_attendance": 1 if (late_count == 0 and early_count == 0 and miss_count == 0 and leave_hours == 0) else 0,
    }


# ==================== Excel 生成 ====================

def generate_output(template_path, output_path, employees, year, month, order_list):
    """
    生成考勤表
    employees: {姓名: {info, daily, summary}}
    order_list: 一览表顺序 [(序号, 姓名, 部门, 办公室)]
    """
    wb = openpyxl.load_workbook(template_path)

    # 更新各sheet标题中的月份为实际月份
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for r in range(1, 4):
            for c in range(1, 10):
                cell = ws.cell(r, c)
                if cell.value and isinstance(cell.value, str):
                    val = cell.value
                    # 替换 "2026年7月份" → "2026年{month}月份"
                    import re
                    val = re.sub(r'(\d{4}年)\d{1,2}(月份)', rf'\g<1>{month}\g<2>', val)
                    # 替换 "（7月）" → "（{month}月）"
                    val = re.sub(r'（\d{1,2}月）', f'（{month}月）', val)
                    if val != cell.value:
                        cell.value = val

    # 生成考勤一览表
    generate_summary_sheet(wb["考勤一览表"], employees, order_list, year, month)

    # 生成办公室明细（从order_list动态构建员工列表，排除非打卡/居家/整月出差/计时工）
    exclude_detail = NON_PUNCH_EMPLOYEES | HOME_OFFICE_EMPLOYEES | TRAVEL_MONTHLY_REMARK.keys() | HOURS_WORKERS.keys()
    office_emps = {}
    for seq, name, dept, office in order_list:
        if name in exclude_detail:
            continue
        if name not in employees:
            continue
        if office not in office_emps:
            office_emps[office] = []
        office_emps[office].append(name)

    for office_name, emp_list in office_emps.items():
        if office_name in wb.sheetnames:
            generate_detail_sheet(wb[office_name], employees, emp_list, year, month)

    wb.save(output_path)
    return output_path


def generate_summary_sheet(ws, employees, order_list, year, month):
    """生成考勤一览表"""
    # 清除模板原数据
    for r in range(4, 40):
        for c in range(1, 9):
            cell = ws.cell(r, c)
            if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                cell.value = None

    row = 4
    for seq, name, department, office in order_list:
        emp = employees.get(name, {})
        ws.cell(row, 1).value = seq
        ws.cell(row, 2).value = department
        ws.cell(row, 3).value = office
        ws.cell(row, 4).value = name

        # 特殊类型处理
        if name in HOURS_WORKERS:
            # 计时工：应出勤/实出勤空，备注总工时（从月度汇总工作时长计算）
            ws.cell(row, 5).value = ""
            ws.cell(row, 6).value = ""
            work_min = emp.get("info", {}).get("work_minutes")
            if work_min:
                hours = round(int(work_min) / 60)
                ws.cell(row, 7).value = f"{hours}小时"
            else:
                ws.cell(row, 7).value = f"{HOURS_WORKERS[name]}小时"
        elif name in NON_PUNCH_EMPLOYEES:
            # 非打卡考勤
            ws.cell(row, 5).value = get_should_days(year, month)
            ws.cell(row, 6).value = get_should_days(year, month)
            ws.cell(row, 7).value = "非打卡考勤"
        elif name in HOME_OFFICE_EMPLOYEES:
            ws.cell(row, 5).value = get_should_days(year, month)
            ws.cell(row, 6).value = get_should_days(year, month)
            ws.cell(row, 7).value = "居家办公"
        elif name in TRAVEL_MONTHLY_REMARK:
            # 整月出差
            ws.cell(row, 5).value = get_should_days(year, month)
            ws.cell(row, 6).value = get_should_days(year, month)
            ws.cell(row, 7).value = TRAVEL_MONTHLY_REMARK[name]
        else:
            summary = emp.get("summary", {})
            ws.cell(row, 5).value = summary.get("should_days", 27)
            ws.cell(row, 6).value = summary.get("actual_str", "27")
            ws.cell(row, 7).value = summary.get("abnormal_summary", "")
        ws.cell(row, 8).value = ""  # 签名
        row += 1

    # 制表人
    ws.cell(30, 5).value = "制表：余唯"


def find_employee_blocks(ws):
    """
    从模板扫描所有"姓名："行所在行号，按行号排序。
    模板块高通常为26，但个别块可能因手工调整多出2行（如郁魏），
    因此必须动态定位，不能用固定块高递推。
    返回 [(姓名, 行号)] 按行号排序
    """
    blocks = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 5).value
        if v and "姓名：" in str(v):
            name = str(v).replace("姓名：", "").strip()
            blocks.append((name, r))
    return blocks


def generate_detail_sheet(ws, employees, emp_list, year, month):
    """生成办公室明细，每个员工26行块（按员工顺序分配模板块）"""
    all_blocks = find_employee_blocks(ws)
    # 按员工顺序分配块：第i个员工用第i个模板块
    for i, name in enumerate(emp_list):
        if name not in employees:
            continue
        if i >= len(all_blocks):
            # 员工数超过模板块数，跳过（需复制块，暂不处理）
            continue
        emp = employees[name]
        start_row = all_blocks[i][1]  # 按顺序取块

        # 规则8.11：清除数据区前取消所有合并单元格
        for mr in list(ws.merged_cells.ranges):
            if (mr.min_row >= start_row + 5 and mr.max_row <= start_row + 20 and
                mr.min_col >= 1 and mr.max_col <= 16):
                ws.unmerge_cells(str(mr))

        # 规则8.5：清除打卡数据区并重置字体颜色
        for r in range(start_row + 5, start_row + 21):
            for c in range(1, 17):
                cell = ws.cell(r, c)
                if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                    cell.value = None
                    cell.font = Font(color="FF000000")  # 重置为黑色

        # 规则8.10：加班列设置自动换行
        for r in range(start_row + 5, start_row + 21):
            for c in [7, 15]:  # 左栏G列=7，右栏O列=15
                cell = ws.cell(r, c)
                cell.alignment = Alignment(wrap_text=True, vertical="center")

        # 员工信息行
        safe_write_cell(ws, start_row, 5, f"姓名：{name}")
        safe_write_cell(ws, start_row, 8, f"部门：{emp['info'].get('department', '')}")
        safe_write_cell(ws, start_row, 11, f"日期：{year}.{month:02d}.01-{year}.{month:02d}.{days_in_month(year, month):02d}")
        safe_write_cell(ws, start_row, 15, f"职位：{emp['info'].get('position', '')}")

        # 工作天数/出勤天数/迟到/早退/缺勤/全勤/出差/外出/病假
        summary = emp.get("summary", {})
        work_days_line = (
            f"工作天数：{summary.get('should_days', 27)}  "
            f"出勤天数：{summary.get('actual_str', '27')}  "
            f"迟到次数：{summary.get('late_count', 0)}  "
            f"早退次数：{summary.get('early_count', 0)}  "
            f"缺卡次数：{summary.get('miss_count', 0)}  "
            f"全勤：{summary.get('full_attendance', 0)}  "
            f"出差：{summary.get('travel_days', 0)}  "
            f"外出：{summary.get('outside_days', 0)}  "
            f"病假：{summary.get('sick_days', 0)}"
        )
        safe_write_cell(ws, start_row + 1, 1, work_days_line)

        # 打卡数据（16行，左1-16日，右17-31日）
        # 规则8.7：跟踪请假段，段首标注总时长
        # 注意：休息日不断开请假段
        # 左右栏分别维护状态（避免交替处理导致状态被重置）
        prev_leave_desc_left = None
        prev_is_leave_left = False
        prev_leave_desc_right = None
        prev_is_leave_right = False
        for i in range(16):
            data_row = start_row + 5 + i
            left_day = i + 1
            right_day = i + 17

            if left_day <= days_in_month(year, month):
                entry = emp["daily"].get(left_day)
                # 判断是否为请假段首
                is_leave_start = False
                if entry and entry.get("status") in ("请假", "部分请假"):
                    current_desc = entry.get("leave_desc")
                    if current_desc:
                        import re
                        leave_match = re.search(r'(事假|病假|调休假)\d{2}-\d{2}.*?\d+小时', current_desc)
                        current_leave_marker = leave_match.group(0) if leave_match else current_desc
                        
                        if not prev_is_leave_left or current_leave_marker != prev_leave_desc_left:
                            is_leave_start = True
                        
                        prev_leave_desc_left = current_leave_marker
                        prev_is_leave_left = True
                elif entry and entry.get("status") == "休息":
                    pass
                else:
                    prev_leave_desc_left = None
                    prev_is_leave_left = False
                
                fill_daily_row(ws, data_row, 1, left_day, entry, year, month, is_leave_start)
            
            if right_day <= days_in_month(year, month):
                entry = emp["daily"].get(right_day)
                # 判断是否为请假段首
                is_leave_start = False
                if entry and entry.get("status") in ("请假", "部分请假"):
                    current_desc = entry.get("leave_desc")
                    if current_desc:
                        import re
                        leave_match = re.search(r'(事假|病假|调休假)\d{2}-\d{2}.*?\d+小时', current_desc)
                        current_leave_marker = leave_match.group(0) if leave_match else current_desc
                        
                        if not prev_is_leave_right or current_leave_marker != prev_leave_desc_right:
                            is_leave_start = True
                        
                        prev_leave_desc_right = current_leave_marker
                        prev_is_leave_right = True
                elif entry and entry.get("status") == "休息":
                    pass
                else:
                    prev_leave_desc_right = None
                    prev_is_leave_right = False
                
                fill_daily_row(ws, data_row, 9, right_day, entry, year, month, is_leave_start)

        # 应出勤/实际出勤/员工签字
        safe_write_cell(ws, start_row + 22, 1, f"应出勤天数：{summary.get('should_days', 27)}天")
        safe_write_cell(ws, start_row + 22, 7, f"实际出勤天数：{summary.get('actual_str', '27')}")
        safe_write_cell(ws, start_row + 22, 13, "员工签字：")

        # 考勤异常说明（无论是否有异常，都强制覆盖，防止模板残留数据）
        # 2026-09-16新规则：合并整行(1-16列)+自动换行，行高按内容自适应，确保请假明细完整显示
        ab_row = start_row + 23
        ab_text = f"考勤异常说明：{summary.get('abnormal_detail') or summary['abnormal_summary']}" if summary.get("abnormal_summary") else ""
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row <= ab_row <= mr.max_row:
                ws.unmerge_cells(str(mr))
        if ab_text:
            ws.merge_cells(start_row=ab_row, start_column=1, end_row=ab_row, end_column=16)
        ab_cell = safe_write_cell(ws, ab_row, 1, ab_text)
        ab_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        # 行高：按16列总宽估算所需行数（中文=2单位），只增不减
        width_units = sum((ws.column_dimensions[get_column_letter(c)].width or 9) for c in range(1, 17))
        text_w = sum(2 if ord(ch) > 127 else 1 for ch in ab_text)
        lines = max(1, math.ceil(text_w / max(width_units - 2, 10))) if ab_text else 1
        ab_need = 15 * lines
        if ab_need > (ws.row_dimensions[ab_row].height or 0):
            ws.row_dimensions[ab_row].height = ab_need


def fill_daily_row(ws, row, col_start, day, entry, year, month, is_leave_start=False):
    """填充一行打卡数据（8列）
    is_leave_start: 是否为请假段首（仅段首标注总时长）
    """
    # 日期、星期
    safe_write_cell(ws, row, col_start, f"{month}.{day}")
    safe_write_cell(ws, row, col_start + 1, get_weekday(year, month, day))

    if entry is None:
        return

    status = entry["status"]
    punches = entry.get("punches") or [None, None, None, None]

    # 休息
    if status == "休息":
        for c in range(4):
            safe_write_cell(ws, row, col_start + 2 + c, "休息")
        return

    # 离职
    if status == "离职":
        for c in range(4):
            safe_write_cell(ws, row, col_start + 2 + c, "离职", font=Font(color=COLOR_RED))
        return

    # 旷工：4格留空
    if status == "旷工":
        return

    # 全天请假
    if status == "请假":
        lt = entry["leave_type"] or "事假"
        for c in range(4):
            safe_write_cell(ws, row, col_start + 2 + c, lt, font=Font(color=COLOR_PURPLE))
        # 规则8.7：段首标注整段请假总时长
        if is_leave_start and entry.get("leave_total_hours"):
            total_hours = entry["leave_total_hours"]
            marker = f"事假{total_hours}小时"
            safe_write_cell(ws, row, col_start + 6, marker, font=Font(color=COLOR_RED))
        return

    # 规则8.6：出差/外出日打卡时间必须为黑色
    if status in ("出差", "外出"):
        label = "出差" if status == "出差" else "外出"
        color = COLOR_GREEN if status == "出差" else COLOR_ORANGE
        if status == "出差":
            # 新规则：有打卡全保留时间（黑色），无打卡填"出差"（绿色）
            for c in range(4):
                t = punches[c]
                if t:
                    safe_write_cell(ws, row, col_start + 2 + c, t, font=Font(color="FF000000"))  # 黑色
                else:
                    safe_write_cell(ws, row, col_start + 2 + c, label, font=Font(color=color))
            # 加班列标"出差"（迟到优先）
            if entry.get("is_late"):
                safe_write_cell(ws, row, col_start + 6, "迟到")
            else:
                safe_write_cell(ws, row, col_start + 6, "出差")
        else:
            # 外出：有上午打卡保留上午（黑色），下午"外出"（橙色）；无打卡4格"外出"（橙色）
            has_am = punches[0] or punches[1]
            for c in range(4):
                t = punches[c]
                if has_am and c < 2 and t:
                    safe_write_cell(ws, row, col_start + 2 + c, t, font=Font(color="FF000000"))  # 黑色
                else:
                    safe_write_cell(ws, row, col_start + 2 + c, label, font=Font(color=color))
            # 加班列：外出日不填（迟到例外）
            if entry.get("is_late"):
                safe_write_cell(ws, row, col_start + 6, "迟到")
        return

    # 半天请假
    if status == "部分请假":
        lt = entry["leave_type"] or "事假"
        period = entry["leave_period"]
        if period == "上午":
            # 上午上班格：请假从早上开始，填请假类型
            safe_write_cell(ws, row, col_start + 2, lt, font=Font(color=COLOR_PURPLE))
            # 上午下班格：有打卡填时间，否则请假类型
            if punches[1]:
                safe_write_cell(ws, row, col_start + 3, punches[1])
            else:
                safe_write_cell(ws, row, col_start + 3, lt, font=Font(color=COLOR_PURPLE))
            # 下午两格
            if punches[2]:
                safe_write_cell(ws, row, col_start + 4, punches[2])
            if punches[3]:
                safe_write_cell(ws, row, col_start + 5, punches[3])
        else:  # 下午请假
            # 上午两格
            if punches[0]:
                safe_write_cell(ws, row, col_start + 2, punches[0])
            if punches[1]:
                safe_write_cell(ws, row, col_start + 3, punches[1])
            # 下午上班格：请假类型
            safe_write_cell(ws, row, col_start + 4, lt, font=Font(color=COLOR_PURPLE))
            # 下午下班格：有打卡填时间，否则请假类型
            if punches[3]:
                safe_write_cell(ws, row, col_start + 5, punches[3])
            else:
                safe_write_cell(ws, row, col_start + 5, lt, font=Font(color=COLOR_PURPLE))
        # 规则8.7：段首标注整段请假总时长，否则标注当天实际小时数
        if is_leave_start and entry.get("leave_total_hours"):
            total_hours = entry["leave_total_hours"]
            safe_write_cell(ws, row, col_start + 6, f"事假{total_hours}小时", font=Font(color=COLOR_RED))
        elif entry.get("marker"):
            safe_write_cell(ws, row, col_start + 6, entry["marker"])
        return

    # 半天出差/外出：
    #   被覆盖半天的第1格（C/E）：打卡时间在覆盖开始前才填时间，否则填状态名
    #   被覆盖半天的第2格（D/F）：打卡时间在覆盖结束前填状态名，之后填时间
    #   未覆盖时段的格：有打卡填时间，无打卡填空
    if status in ("部分出差", "部分外出"):
        label = "出差" if status == "部分出差" else "外出"
        color = COLOR_GREEN if status == "部分出差" else COLOR_ORANGE
        period = entry.get("travel_period") or entry.get("out_period")
        cs = entry.get("cover_start")
        ce = entry.get("cover_end")
        covered = [0, 1] if period == "上午" else [2, 3]
        first_covered = covered[0]
        for c in range(4):
            t = punches[c]
            if c in covered:
                if c == first_covered and t:
                    tm = time_to_minutes(t)
                    if cs is None or tm < cs:
                        safe_write_cell(ws, row, col_start + 2 + c, t)
                    else:
                        safe_write_cell(ws, row, col_start + 2 + c, label, font=Font(color=color))
                elif t and ce is not None and time_to_minutes(t) >= ce:
                    # 覆盖结束后打卡（如余唯7.7 12:13>12:00、许凡13 18:00）
                    safe_write_cell(ws, row, col_start + 2 + c, t)
                else:
                    safe_write_cell(ws, row, col_start + 2 + c, label, font=Font(color=color))
            else:
                if t:
                    safe_write_cell(ws, row, col_start + 2 + c, t)
                else:
                    safe_write_cell(ws, row, col_start + 2 + c, "")
        # 加班列：标注"部分外出X小时"（迟到优先）
        if entry.get("is_late"):
            safe_write_cell(ws, row, col_start + 6, "迟到")
        elif status == "部分外出":
            outside_hours = entry.get("outside_hours", 0)
            if outside_hours > 0:
                safe_write_cell(ws, row, col_start + 6, f"部分外出{outside_hours}小时")
        return

    # 正常日：填打卡时间，缺卡格填"补卡"（描述中有缺卡字样）或"缺卡"（推断缺卡）
    miss_pos = entry.get("miss_pos", [])
    official_miss = entry.get("official_miss", set())
    for c in range(4):
        t = punches[c]
        if t:
            font = None
            # 迟到标红（第1格）
            if c == 0 and entry.get("is_late"):
                font = Font(color=COLOR_RED)
            # 早退标红（第4格）
            if c == 3 and entry.get("is_early"):
                font = Font(color=COLOR_RED)
            safe_write_cell(ws, row, col_start + 2 + c, t, font=font)
        elif c in miss_pos:
            # 描述中明确"X缺卡"→补卡（标蓝）；推断缺卡→"缺卡"（标蓝）
            miss_label = "补卡" if c in official_miss else "缺卡"
            safe_write_cell(ws, row, col_start + 2 + c, miss_label, font=Font(color=COLOR_BLUE))
        else:
            safe_write_cell(ws, row, col_start + 2 + c, "")

    # 加班列标记
    marker = entry.get("marker")
    # 补卡标记第几次（如"补卡1"、"补卡2"）
    if entry.get("makeup_index"):
        marker = f"补卡{entry['makeup_index']}"
    if marker:
        safe_write_cell(ws, row, col_start + 6, marker)


# ==================== 特殊标记注入 ====================

def apply_special_markers(employees):
    """将台风等人工补充的特殊标记注入到对应员工的daily"""
    for emp_name, day_markers in SPECIAL_MARKERS.items():
        if emp_name not in employees:
            continue
        for day, marker in day_markers.items():
            entry = employees[emp_name]["daily"].get(day)
            if entry:
                entry["marker"] = marker


# ==================== 主流程 ====================

def main():
    parser = argparse.ArgumentParser(description="锦业体育考勤表生成")
    parser.add_argument("--punch", required=True, help="打卡记录表路径（辅助）")
    parser.add_argument("--summary", required=True, help="月度考勤汇总路径（主数据源）")
    parser.add_argument("--leave", required=True, help="请假表路径")
    parser.add_argument("--travel", required=True, help="出差表路径（辅助）")
    parser.add_argument("--outside", required=True, help="外出表路径（辅助）")
    parser.add_argument("--month", type=int, required=True, help="月份（如7）")
    parser.add_argument("--year", type=int, default=2026, help="年份（默认2026）")
    parser.add_argument("--output", required=True, help="输出文件路径")
    parser.add_argument("--template", default=None, help="模板路径（默认使用skill内置模板）")
    args = parser.parse_args()

    # 模板路径
    if args.template is None:
        script_dir = Path(__file__).parent
        args.template = str(script_dir.parent / "assets" / "attendance-template.xlsx")

    print(f"[1/5] 解析月度考勤汇总（主数据源）...")
    summary_data = parse_summary(args.summary, args.year, args.month)
    print(f"      共 {len(summary_data)} 名员工")

    print(f"[2/5] 解析请假表（调休假类型修正）...")
    leave_map = parse_leave_types(args.leave, args.year, args.month)
    print(f"      涉及 {len(leave_map)} 名员工")

    print(f"[3/5] 构建每日考勤...")
    employees = {}
    for name, data in summary_data.items():
        daily = {}
        for day, (desc, times) in data["daily"].items():
            daily[day] = build_daily_entry(desc, times, name, day, leave_map, args.year, args.month)

        summary = compute_attendance(
            name, daily, args.year, args.month,
            data["info"]["miss_in"], data["info"]["miss_out"]
        )
        employees[name] = {
            "info": data["info"],
            "daily": daily,
            "summary": summary,
        }

    # 注入特殊标记（台风等）
    apply_special_markers(employees)

    print(f"[4/5] 确定一览表顺序（从模板读取）...")
    order_list = build_order_list(args.template, summary_data, employees)

    print(f"[5/5] 生成考勤表...")
    generate_output(args.template, args.output, employees, args.year, args.month, order_list)
    print(f"完成！输出文件：{args.output}")


def build_order_list(template_path, summary_data, employees):
    """
    构建一览表人员顺序：
    1. 优先从月度汇总读取实际在职员工
    2. 用模板的部门/办公室信息补充（模板中有该员工时）
    3. 模板中没有的新员工，从月度汇总读取部门/办公室
    4. 不在月度汇总中的模板员工，仅保留已知特殊人员（非打卡/居家办公/整月出差/计时工），离职人员移除
    5. 排序：公庄办公室在前，陈江办公室在后（按模板顺序），新员工追加
    返回 [(序号, 姓名, 部门, 办公室)]
    """
    # 从模板读取原有顺序和部门/办公室
    wb = openpyxl.load_workbook(template_path, data_only=True)
    ws = wb["考勤一览表"]
    template_order = []  # [(姓名, 部门, 办公室)]
    for r in range(4, 35):
        name = ws.cell(r, 4).value
        if not name:
            continue
        name = str(name).strip()
        if not name:
            continue
        department = str(ws.cell(r, 2).value or "").strip()
        office = str(ws.cell(r, 3).value or "").strip()
        template_order.append((name, department, office))
    wb.close()

    # 特殊人员集合（不在月度汇总中也保留）
    special_keep = NON_PUNCH_EMPLOYEES | HOME_OFFICE_EMPLOYEES | TRAVEL_MONTHLY_REMARK.keys() | HOURS_WORKERS.keys()

    # 月度汇总中的实际员工
    summary_names = set(summary_data.keys())

    # 构建最终顺序：先按模板顺序（保留在月度汇总中或特殊人员），再追加新员工
    final_order = []
    seen = set()

    # 1. 模板顺序中的员工
    for name, department, office in template_order:
        if name in seen:
            continue
        # 离职人员（在LEAVE_COMPANY_DATE中且不在月度汇总）移除
        if name in LEAVE_COMPANY_DATE and name not in summary_names:
            continue
        # 在月度汇总中 或 特殊保留人员
        if name in summary_names or name in special_keep:
            final_order.append((name, department, office))
            seen.add(name)

    # 2. 月度汇总中的新员工（模板中没有的）
    for name in summary_names:
        if name in seen:
            continue
        # 从月度汇总读取部门/办公室
        info = summary_data[name].get("info", {})
        department = info.get("department", "")
        # 办公室从考勤组推断（考勤组含"公庄"→公庄办公室，否则→陈江办公室）
        group = info.get("group", "")
        office = "公庄办公室" if "公庄" in group else "陈江办公室"
        final_order.append((name, department, office))
        seen.add(name)

    # 3. 特殊保留人员（不在月度汇总也不在模板）
    for name in special_keep:
        if name in seen:
            continue
        final_order.append((name, "", ""))
        seen.add(name)

    # 加序号
    result = [(i + 1, name, dept, office) for i, (name, dept, office) in enumerate(final_order)]
    return result


if __name__ == "__main__":
    main()

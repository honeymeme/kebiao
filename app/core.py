# -*- coding: utf-8 -*-
"""课表解析 + 空课表生成 核心逻辑。"""
import os, re, sys, copy

HERE = os.path.dirname(os.path.abspath(__file__))
VENDOR = os.path.join(HERE, "vendor")
if VENDOR not in sys.path:
    sys.path.insert(0, VENDOR)

import pdfplumber
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, Color, Alignment

BLOCKS = [(1, 2), (3, 4), (5, 6), (7, 8)]
BLOCK_TIMES = ["8：30-9：50", "10：00-11：50", "1：40-3：20", "3：40-4：30"]
WEEKDAY_NAMES = ["周一", "周二", "周三", "周四", "周五"]
HEADER_NAMES = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

RE_PERIOD = re.compile(r"[（(]\s*(\d{1,2})\s*(?:[-—~]\s*(\d{1,2}))?\s*节\s*[)）]")
RE_WEEK = re.compile(r"(\d{1,2})\s*(?:[-—~]\s*(\d{1,2}))?\s*周")


class Member:
    def __init__(self, name, klass, phone, dept, is_minister, pdf_path):
        self.name = name
        self.klass = klass
        self.phone = phone
        self.dept = dept
        self.is_minister = is_minister
        self.pdf_path = pdf_path
        self.sid = ""
        self.busy = {}      # (day, block) -> set(weeks)
        self.error = None

    def __repr__(self):
        return "%s/%s/%s%s" % (self.name, self.klass, self.phone, "(部长)" if self.is_minister else "")


def parse_member_filename(path, dept):
    """文件名格式: 部门 姓名 班级 联系方式（顺序可能不同，按内容识别）"""
    base = os.path.splitext(os.path.basename(path))[0].strip()
    parts = [p for p in re.split(r"\s+", base) if p]
    phone = ""
    klass = ""
    name = ""
    rest = []
    for p in parts:
        if re.fullmatch(r"1\d{10}", p):
            if not phone:
                phone = p
            continue
        # 姓名和手机号粘在一起：夏俊杰18251539599
        m = re.fullmatch(r"(.+?)(1\d{10})", p)
        if m and not phone:
            phone = m.group(2)
            p = m.group(1)
        if re.search(r"\d", p) and not klass:
            klass = p
        else:
            rest.append(p)
    if rest and rest[0] == dept:
        rest = rest[1:]
    if rest:
        name = rest[0]
    return name, klass, phone


def collect_members(root_dir, dept, order=None):
    members = []
    for cur, dirs, files in os.walk(root_dir):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            p = os.path.join(cur, f)
            is_min = "部长" in os.path.relpath(p, root_dir)
            name, klass, phone = parse_member_filename(p, dept)
            m = Member(name, klass, phone, dept, is_min, p)
            members.append(m)
    # 排序：部长在前，其余保持来源顺序（zip 内顺序），再按姓名兜底
    seq = {b: i for i, b in enumerate(order or [])}
    members.sort(key=lambda m: (0 if m.is_minister else 1,
                                seq.get(os.path.basename(m.pdf_path), 9999),
                                m.name))
    return members


def parse_cell_periods(text):
    """返回 [(p_start, p_end, weeks_set)]，单/双周只取对应奇偶周。"""
    out = []
    for m in RE_PERIOD.finditer(text or ""):
        p1 = int(m.group(1))
        p2 = int(m.group(2) or p1)
        tail = text[m.end():]
        tail = tail.split("/")[0]
        if "周" not in tail:
            continue
        weeks = set()
        for w in RE_WEEK.finditer(tail):
            a = int(w.group(1)); b = int(w.group(2) or a)
            # 单/双周不作区分，整段按有课处理
            if 1 <= a <= 30:
                weeks.update(range(a, b + 1))
        if weeks:
            out.append((p1, p2, weeks))
    return out


def _clean(c):
    return str(c or "").replace("\n", "").strip()


def _periods_from_spec(s):
    m = re.fullmatch(r"(\d{1,2})(?:\s*[-—~]\s*(\d{1,2}))?", s)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2) or m.group(1))


def _weeks_in_text(text):
    weeks = set()
    for w in RE_WEEK.finditer(text or ""):
        a = int(w.group(1)); b = int(w.group(2) or a)
        if 1 <= a <= 30:
            weeks.update(range(a, b + 1))
    return weeks


def _apply(busy, day, p1, p2, weeks):
    for p in range(p1, p2 + 1):
        blk = (p - 1) // 2
        if 0 <= blk < len(BLOCKS):
            busy.setdefault((day, blk), set()).update(weeks)


def _parse_grid(table, hdr, col_day, busy):
    rows = table[hdr + 1:] if hdr >= 0 else table
    for row in rows:
        for ci, cell in enumerate(row):
            day = col_day.get(ci)
            if not day or not cell:
                continue
            for p1, p2, weeks in parse_cell_periods(str(cell)):
                _apply(busy, day, p1, p2, weeks)


def _parse_day_major(table, busy):
    """行=星期，列=[星期, 节次, 课程…] 的版式"""
    day = None
    periods = None
    for row in table:
        if not row:
            continue
        c0 = _clean(row[0])
        c1 = _clean(row[1]) if len(row) > 1 else ""
        if c0 in HEADER_NAMES:
            day = HEADER_NAMES.index(c0) + 1
            periods = None
        elif c0:
            day = None
        if c1:
            periods = _periods_from_spec(c1)
        if day is None or not periods:
            continue
        text = " ".join(_clean(c) for c in row[2:] if _clean(c))
        if "周" not in text:
            continue
        weeks = _weeks_in_text(text)
        if weeks:
            _apply(busy, day, periods[0], periods[1], weeks)


def parse_pdf(member):
    busy = {}
    try:
        with pdfplumber.open(member.pdf_path) as pdf:
            col_day = None
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    if not table:
                        continue
                    if not member.sid:
                        for row in table[:2]:
                            for c in row:
                                m = re.search(r"学号\s*[:：]\s*([A-Za-z0-9]+)", _clean(c))
                                if m:
                                    member.sid = m.group(1)
                                    break
                            if member.sid:
                                break
                    hdr, best, best_map = None, 0, None
                    for i, row in enumerate(table):
                        cd = {}
                        for ci, c in enumerate(row):
                            s = _clean(c)
                            for di, hn in enumerate(HEADER_NAMES):
                                if hn in s:
                                    cd[ci] = di + 1
                        if len(cd) > best:
                            hdr, best, best_map = i, len(cd), cd
                    if best >= 3:
                        col_day = best_map
                        _parse_grid(table, hdr, col_day, busy)
                    elif best == 1 and 0 in (best_map or {}) and len(table) > 2:
                        _parse_day_major(table, busy)
                    elif col_day is not None and max(len(r) for r in table) >= 9:
                        _parse_grid(table, -1, col_day, busy)
                    else:
                        _parse_day_major(table, busy)
    except Exception as e:
        member.error = "%s: %s" % (type(e).__name__, e)
    member.busy = busy
    return busy


def free_weeks(busy_set, w1, w2):
    return [w for w in range(w1, w2 + 1) if w not in (busy_set or set())]


def fmt_weeks(weeks):
    if not weeks:
        return ""
    spans = []
    for w in sorted(weeks):
        if spans and w == spans[-1][1] + 1:
            spans[-1][1] = w
        else:
            spans.append([w, w])
    return "，".join("%d-%d" % (a, b) if a != b else "%d" % a for a, b in spans)


def build_empty_sheet(template, dept, members, w1, w2, out_path):
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    ws["B2"] = "%s空课表" % dept

    # 计算每格文字
    grid = {}   # (block, day) -> [(name, weeks_str, is_min)]
    for blk in range(len(BLOCKS)):
        for day in range(1, 6):
            items = []
            for mi, m in enumerate(members):
                fw = free_weeks(m.busy.get((day, blk)), w1, w2)
                if fw:
                    items.append((mi, m.name, fmt_weeks(fw), m.is_minister))
            grid[(blk, day)] = items

    max_lines = 1
    for blk in range(len(BLOCKS)):
        for day in range(1, 6):
            max_lines = max(max_lines, len(grid[(blk, day)]))

    size = 14 if max_lines <= 8 else (12 if max_lines <= 12 else 10)
    line_h = {14: 18, 12: 16, 10: 13}[size]

    for blk in range(len(BLOCKS)):
        base_row = 4 + blk * 4
        # 行高
        lines = max(len(grid[(blk, day)]) for day in range(1, 6))
        need = max(39.6, (lines * line_h + 8) / 4.0)
        for r in range(base_row, base_row + 4):
            ws.row_dimensions[r].height = need
        for day in range(1, 6):
            col = 3 + day  # D=4
            cell = ws.cell(row=base_row, column=col)
            items = grid[(blk, day)]
            if not items:
                cell.value = None
                continue
            rt = CellRichText()
            for i, (_mi, name, wk, is_min) in enumerate(items):
                if i:
                    rt.append(TextBlock(InlineFont(rFont="微软雅黑", sz=size), "\n"))
                txt = "%s（%s）" % (name, wk)
                color = "FFFF0000" if is_min else None
                rt.append(TextBlock(InlineFont(rFont="微软雅黑", sz=size, color=color), txt))
            cell.value = rt
            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # 底部信息收集表
    start = 22
    ws["B21"], ws["C21"], ws["D21"] = "姓名", "班级", "联系方式"
    ordered = ([m for m in members if not m.is_minister] +
               [m for m in members if m.is_minister])   # 部长放最下面

    # 清空旧数据
    for r in range(start, max(ws.max_row, start + len(ordered)) + 1):
        for c in (2, 3, 4):
            ws.cell(row=r, column=c).value = None

    last_row = 39
    for i, m in enumerate(ordered):
        r = start + i
        if r > last_row:
            for c in (2, 3, 4):
                src = ws.cell(row=last_row, column=c)
                dst = ws.cell(row=r, column=c)
                dst._style = copy.copy(src._style)
            ws.row_dimensions[r].height = ws.row_dimensions[last_row].height
        for c, val in ((2, m.name), (3, m.klass), (4, m.phone)):
            cell = ws.cell(row=r, column=c)
            cell.value = val
            if m.is_minister:
                f = copy.copy(cell.font)
                f.color = Color(rgb="FFFF0000")
                cell.font = f

    wb.save(out_path)
    return out_path

# -*- coding: utf-8 -*-
"""检查 + 生成 空课表 / 值班表。"""
import os, re, sys, shutil, zipfile, json
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from core import (collect_members, parse_pdf, free_weeks, fmt_weeks, Member,
                  BLOCKS, BLOCK_TIMES, WEEKDAY_NAMES)

import openpyxl
from openpyxl.styles import Font, Alignment, Color

EMPTY_TPL = os.path.join(HERE, "空课表模板.xlsx")
DUTY_TPL = os.path.join(HERE, "值班表模板.xlsx")

CN_DIGITS = "零一二三四五六七八九"


def cn_num(n):
    if n <= 10:
        return "十" if n == 10 else CN_DIGITS[n]
    if n < 20:
        return "十" + CN_DIGITS[n - 10]
    if n % 10 == 0:
        return CN_DIGITS[n // 10] + "十"
    return CN_DIGITS[n // 10] + "十" + CN_DIGITS[n % 10]


def discover_depts(src_dir):
    """返回 {部门: 目录}"""
    depts = {}
    entries = [e for e in os.listdir(src_dir) if not e.startswith("__")]
    dirs = [e for e in entries if os.path.isdir(os.path.join(src_dir, e))]
    files = [e for e in entries if e.lower().endswith(".pdf")]
    for d in dirs:
        name = re.sub(r"课表$", "", d).strip()
        depts[name] = os.path.join(src_dir, d)
    if files:
        # 没有子目录时，按文件名首段分组
        groups = {}
        for f in files:
            token = re.split(r"\s+", f)[0]
            groups.setdefault(token, []).append(f)
        for k, fs in groups.items():
            target = os.path.join(src_dir, "_" + k)
            os.makedirs(target, exist_ok=True)
            for f in fs:
                srcp = os.path.join(src_dir, f)
                dstp = os.path.join(target, f)
                if not os.path.exists(dstp):
                    shutil.copy2(srcp, dstp)
            depts[k] = target
    return depts


def order_from_zip(zip_path):
    with zipfile.ZipFile(zip_path) as z:
        return [os.path.basename(n) for n in z.namelist() if n.lower().endswith(".pdf")]


def load_all(src_dir, order=None):
    depts = discover_depts(src_dir)
    result = {}
    for name, path in depts.items():
        members = collect_members(path, name, order)
        for m in members:
            parse_pdf(m)
        result[name] = members
    return result


def load_any(src_dir, order=None):
    """自动判断是 PDF 通道还是"做好的空课表"通道"""
    has_sheet = any(f.lower().endswith((".xlsx", ".xlsm")) for f in os.listdir(src_dir))
    if has_sheet:
        import sheets
        return "sheet", sheets.load_sheets(src_dir, order)
    return "pdf", load_all(src_dir, order)


def check(src_dir, order=None):
    mode, data = load_any(src_dir, order)
    report = {"depts": [], "problems": [], "total": 0, "mode": mode}
    seen = {}
    for dept, members in data.items():
        info = {"dept": dept, "count": len(members),
                "ministers": [m.name for m in members if m.is_minister],
                "ok": True}
        for m in members:
            fn = os.path.basename(m.pdf_path)
            if not m.name or not m.klass or not m.phone:
                if mode == "pdf":
                    report["problems"].append("%s：文件名「%s」缺少姓名/班级/电话（应为「部门 姓名 班级 电话」）" % (dept, fn))
                else:
                    report["problems"].append("%s：%s 在底部信息表里缺少班级或联系方式" % (dept, m.name or fn))
                info["ok"] = False
            if m.error:
                report["problems"].append("%s：%s 解析失败（%s）" % (dept, m.name or fn, m.error))
                info["ok"] = False
            elif mode == "pdf" and not m.busy:
                report["problems"].append("%s：%s 的 PDF 没有读到任何课程，请检查文件" % (dept, m.name or fn))
                info["ok"] = False
            elif mode == "sheet" and not m.busy:
                report["problems"].append("%s：%s 在空课表里一次空闲都没有出现，无法排班" % (dept, m.name or fn))
                info["ok"] = False
            key = (m.name, m.phone)
            if key in seen:
                report["problems"].append("重复成员：%s（%s 与 %s）" % (m.name, seen[key], dept))
            seen[key] = dept
        report["depts"].append(info)
        report["total"] += len(members)
    for d in DEPT_ORDER:
        pass
    report["depts"].sort(key=lambda x: DEPT_ORDER.index(x["dept"]) if x["dept"] in DEPT_ORDER else 99)
    report["ok"] = not report["problems"]
    return report, data


DEPT_ORDER = ["办公室", "科协", "青协", "生活部", "文体部", "宣传部", "心协", "学习部", "组织部"]


def build_empty(template, dept, members, w1, w2, out_path):
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    ws["B2"] = "%s部门空课表" % dept

    grid = {}
    for blk in range(len(BLOCKS)):
        for day in range(1, 6):
            items = []
            for mi, m in enumerate(members):
                fw = free_weeks(m.busy.get((day, blk)), w1, w2)
                if fw:
                    items.append((mi, m.name, fmt_weeks(fw), m.is_minister))
            grid[(blk, day)] = items

    max_lines = max([len(v) for v in grid.values()] + [1])
    size = 14 if max_lines <= 8 else (12 if max_lines <= 12 else 10)
    line_h = {14: 18, 12: 16, 10: 13}[size]

    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    for blk in range(len(BLOCKS)):
        base_row = 4 + blk * 4
        lines = max(len(grid[(blk, day)]) for day in range(1, 6))
        need = max(39.6, (lines * line_h + 8) / 4.0)
        for r in range(base_row, base_row + 4):
            ws.row_dimensions[r].height = need
        for day in range(1, 6):
            cell = ws.cell(row=base_row, column=3 + day)
            items = grid[(blk, day)]
            if not items:
                cell.value = None
                continue
            rt = CellRichText()
            for i, (_mi, name, wk, is_min) in enumerate(items):
                color = "FFFF0000" if is_min else "FF000000"
                txt = "%s（%s）" % (name, wk)
                if i < len(items) - 1:
                    txt += "\n"      # 换行跟文字放同一段，避免被当成空段丢掉
                rt.append(TextBlock(InlineFont(rFont="微软雅黑", sz=size, color=color), txt))
            cell.value = rt
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    start = 22
    ws["B21"], ws["C21"], ws["D21"] = "姓名", "班级", "联系方式"
    ordered = ([m for m in members if not m.is_minister] +
               [m for m in members if m.is_minister])
    for r in range(start, max(ws.max_row, start + len(ordered)) + 1):
        for c in (2, 3, 4):
            ws.cell(row=r, column=c).value = None

    last_row = 39
    for i, m in enumerate(ordered):
        r = start + i
        if r > last_row:
            for c in (2, 3, 4):
                ws.cell(row=r, column=c)._style = copy.copy(ws.cell(row=last_row, column=c)._style)
            ws.row_dimensions[r].height = ws.row_dimensions[last_row].height
        for c, val in ((2, m.name), (3, m.klass), (4, m.phone)):
            cell = ws.cell(row=r, column=c)
            cell.value = val
            if m.is_minister:
                f = copy.copy(cell.font)
                f.color = Color(rgb="FFFF0000")
                cell.font = f

    wb.save(out_path)
    return grid


def plan_duty(all_members, w1, w2, per_cell=2):
    """按周排班。返回 (plan, notes)"""
    people = []
    for dept, ms in all_members.items():
        for m in ms:
            people.append(m)

    count = {id(p): 0 for p in people}
    last_week = {id(p): -99 for p in people}
    n_normal = max(1, len([p for p in people if not p.is_minister]))
    total_slots = (w2 - w1 + 1) * 20 * per_cell
    cap = max(2, -(-total_slots // n_normal))   # 每人平均应排的次数（向上取整）
    plan = {}
    notes = {}

    for w in range(w1, w2 + 1):
        week_plan = {}
        notes[w] = []
        used_today = {}
        cells = [(blk, day) for blk in range(4) for day in range(1, 6)]
        for blk, day in cells:
            cands = [p for p in people
                     if w not in (p.busy.get((day, blk)) or set())]
            def key(p):
                # 分三档：没排够的干事(0) > 已排够次数的干事(1) > 部长(2)
                # 这样某个时段只有两名干事可用时，也不会把他们钉死在那儿。
                c = count[id(p)]
                if p.is_minister:
                    tier = 2       # 部长排最后
                elif c >= cap + 2:
                    tier = 3       # 干事已经明显超额（比如被钉在某个格子），这种时候让部长顶
                elif c < cap:
                    tier = 0
                else:
                    tier = 1
                penalty = 2 if id(p) in used_today else 0
                dept_idx = DEPT_ORDER.index(p.dept) if p.dept in DEPT_ORDER else 99
                return (tier,
                        c + penalty,
                        c,
                        dept_idx,
                        w - last_week[id(p)],
                        p.name)
            cands.sort(key=key)
            chosen = cands[:per_cell]
            week_plan[(blk, day)] = chosen
            for p in chosen:
                count[id(p)] += 1
                last_week[id(p)] = w
                used_today[id(p)] = used_today.get(id(p), 0) + 1
            if len(chosen) < per_cell:
                notes[w].append("%s %s 只有 %d 人可排" % (WEEKDAY_NAMES[day - 1], BLOCK_TIMES[blk], len(chosen)))
        plan[w] = week_plan

    # ---- 二次平衡：把各人次数拉平（干事优先补位） ----
    def all_cells():
        for w in range(w1, w2 + 1):
            for key, lst in plan[w].items():
                yield w, key, lst

    def count_map():
        c = {id(p): 0 for p in people}
        for _w, _k, lst in all_cells():
            for q in lst:
                c[id(q)] += 1
        return c

    def works_day(w, day, p):
        return any(d == day and p in lst for (_b, d), lst in plan[w].items())

    def try_swap(low, high):
        for w in range(w1, w2 + 1):
            for (blk, day), lst in plan[w].items():
                if high not in lst:
                    continue
                if w in (low.busy.get((day, blk)) or set()):
                    continue
                if works_day(w, day, low):
                    continue
                lst[lst.index(high)] = low
                return True
        return False

    normal = [p for p in people if not p.is_minister]
    for _ in range(3000):
        c = count_map()
        if not normal:
            break
        low = min(normal, key=lambda p: c[id(p)])
        high = max(people, key=lambda p: c[id(p)])
        if c[id(high)] - c[id(low)] <= 2:
            break
        lows = sorted(normal, key=lambda p: c[id(p)])
        highs = sorted(people, key=lambda p: c[id(p)], reverse=True)
        moved = False
        for lo in lows:
            for hi in highs:
                if hi.dept != lo.dept:      # 只在同一个部门内部拉平，不打乱部门优先顺序
                    continue
                if c[id(hi)] - c[id(lo)] <= 2:
                    break
                if try_swap(lo, hi):
                    moved = True
                    break
            if moved:
                break
        if not moved:
            break

    counts_map = count_map()
    stats = {"min": 0, "max": 0, "avg": 0, "zero": []}
    if people:
        counts = sorted(counts_map[id(p)] for p in people)
        stats = {
            "min": counts[0],
            "max": counts[-1],
            "avg": round(sum(counts) / len(counts), 1),
            "zero": [p.name for p in people if counts_map[id(p)] == 0],
        }
    return plan, notes, stats


def build_duty(template, week, week_plan, notes, out_path):
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    ws["A1"] = "第%s周系办值班表" % cn_num(week)
    for blk in range(4):
        row = 3 + blk * 4
        for day in range(1, 6):
            cell = ws.cell(row=row, column=2 + day)
            cell.value = None
            people = week_plan.get((blk, day), [])
            if people:
                cell.value = "\n".join("%s/%s" % (p.name, p.phone) for p in people)
                cell.font = Font(name="微软雅黑", sz=11)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if notes:
        ws["I3"] = "注：" + "；".join(notes) + "，其余位置正常。"
        ws["I3"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    wb.save(out_path)
    return out_path


def build_detail(data, w1, w2, out_path, counts=None):
    """生成"解析明细.xlsx"：每人每格的有课/空闲周次，方便抽查。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    head = ["部门", "姓名", "班级", "联系方式", "部长", "总排班次数"]
    ws.append(head)

    ws2 = wb.create_sheet("时段明细")
    ws2.append(["部门", "姓名", "星期", "时段", "有课周次（1-30）", "范围内空闲周次"])

    for dept, members in data.items():
        for m in members:
            ws.append([dept, m.name, m.klass, m.phone, "是" if m.is_minister else "",
                       (counts or {}).get(m.name, "")])
            for day in range(1, 6):
                for blk in range(4):
                    busy = sorted(m.busy.get((day, blk)) or [])
                    free = free_weeks(m.busy.get((day, blk)), w1, w2)
                    ws2.append([dept, m.name, WEEKDAY_NAMES[day - 1], BLOCK_TIMES[blk],
                                fmt_weeks(busy) or "（无课）", fmt_weeks(free) or "（全满）"])

    for sheet in (ws, ws2):
        widths = {"A": 10, "B": 10, "C": 14, "D": 16, "E": 8, "F": 26}
        for col, w in widths.items():
            sheet.column_dimensions[col].width = w
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(name="微软雅黑", bold=True, sz=11)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = openpyxl.styles.Font(name="微软雅黑", sz=10)
                cell.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")
    ws.freeze_panes = "A2"
    ws2.freeze_panes = "A2"
    wb.save(out_path)
    return out_path

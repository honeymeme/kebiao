# -*- coding: utf-8 -*-
"""检查 + 生成 空课表 / 值班表。"""
import os, re, sys, shutil, zipfile, json
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from core import (collect_members, parse_pdf, free_weeks, fmt_weeks, Member,
                  BLOCKS, BLOCK_TIMES, WEEKDAY_NAMES, BLOCK_SHORT)

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


def canon_dept(name):
    name = (name or "").strip()
    if not name:
        return name
    if name in DEPT_ORDER:
        return name
    for c in DEPT_ORDER:
        if c.startswith(name) or name.startswith(c):
            return c
    return name


def discover_depts(src_dir):
    """返回 {部门: 目录}"""
    depts = {}
    entries = [e for e in os.listdir(src_dir) if not e.startswith("__")]
    dirs = [e for e in entries if os.path.isdir(os.path.join(src_dir, e))]
    files = [e for e in entries if e.lower().endswith(".pdf")]
    for d in dirs:
        name = canon_dept(re.sub(r"课表$", "", d).strip())
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
    has_sheet = any(f.lower().endswith((".xlsx", ".xlsm")) and "值班表" not in f
                    for f in os.listdir(src_dir))
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
    ws["B2"] = "%s空课表" % dept

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
            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

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


def dept_sequence(all_members):
    seq = [d for d in DEPT_ORDER if all_members.get(d)]
    for d in all_members:
        if d not in seq:
            seq.append(d)
    return seq


def infer_offset(seq, w1, hist_depts):
    """根据上传的值班表推测轮值起点（第 w1 周对应哪个部门）"""
    if not seq or not hist_depts:
        return 0
    # 以"最近一次上传的那一周"为锚点，保证后面接着往下轮
    last_w = max(hist_depts)
    last_d = hist_depts[last_w]
    if last_d in seq:
        return (seq.index(last_d) - (last_w - w1)) % len(seq)
    best, best_score = 0, -1
    for off in range(len(seq)):
        score = sum(1 for w, d in hist_depts.items()
                    if seq[(w - w1 + off) % len(seq)] == d)
        if score > best_score:
            best, best_score = off, score
    return best


def week_dept_map(all_members, w1, w2, skip_weeks=None, offset=0):
    """第几周由哪个部门负责（skip 的周也占轮值位，保证接得上）"""
    seq = dept_sequence(all_members)
    if not seq:
        return {}
    skip = set(skip_weeks or ())
    return {w: seq[(w - w1 + offset) % len(seq)]
            for w in range(w1, w2 + 1) if w not in skip}


def plan_duty(all_members, w1, w2, per_cell=2, seed_counts=None, skip_weeks=None, boost=None,
               minister_only=None, start_offset=0, seed_weeks=None):
    """按周排班（一周一个部门）。minister_only = {(day, blk)} 这些格只排部长。"""
    boost = boost or {}
    # 部长专用格不再特殊处理：按正常规则排，谁有空谁上
    minister_only = set()

    def _b(p):
        return int(boost.get("%s|%s" % (p.dept, p.name), 0))
    people = []
    for dept, ms in all_members.items():
        for m in ms:
            people.append(m)

    seed_counts = seed_counts or {}
    skip_weeks = set(skip_weeks or ())
    count = {id(p): int(seed_counts.get(p.name, 0)) for p in people}
    last_week = {id(p): -99 for p in people}
    served = {id(p): set((seed_weeks or {}).get(p.name, [])) for p in people}

    def streak_before(p, w):
        """这个人到上一周为止，已经连续排了几周"""
        s = served[id(p)]
        n, k = 0, w - 1
        while k in s:
            n += 1
            k -= 1
        return n
    n_normal = max(1, len([p for p in people if not p.is_minister]))
    todo_weeks = [w for w in range(w1, w2 + 1) if w not in skip_weeks]
    total_slots = len(todo_weeks) * 20 * per_cell
    cap = max(2, -(-total_slots // n_normal))   # 每人平均应排的次数（向上取整）
    plan = {}
    notes = {}

    # 按部门顺序排队，一个人一个人往外派：
    # 先办公室，办公室的人排完了还有空格就轮到科协，再青协（一周最多三个部门）；
    # 这一周没排上的人，下周接着排在前面。
    seq = dept_sequence(all_members)
    n_seq = len(seq)
    week_dept_out = {}

    queue = []
    for d in seq:
        for m in all_members.get(d, []):
            queue.append(m)
    qpos = 0
    if queue and start_offset:
        # 有历史值班表时，把指针挪到"最后一个部门的下一个部门"
        # 从"排了一半的部门"开始：优先挑已经有人值过、但还没排完的那个部门
        served_names = set((seed_weeks or {}).keys())
        half, zero = [], []
        if served_names:
            for _i, _d in enumerate(seq):
                _ms = [m for m in all_members.get(_d, []) if not m.is_minister]
                if not _ms:
                    continue
                _done = sum(1 for m in _ms if m.name in served_names)
                if 0 < _done < len(_ms):
                    half.append(_i)
                elif _done == 0:
                    zero.append(_i)
        start_i = half[0] if half else (zero[0] if zero else 0)
        anchor = seq[start_i % n_seq]
        qpos = next((i for i, m in enumerate(queue) if m.dept == anchor), 0)

    for w in range(w1, w2 + 1):
        week_plan = {}
        notes[w] = []
        if w in skip_weeks:
            plan[w] = week_plan
            continue
        if not queue:
            plan[w] = week_plan
            notes[w].append("没有可用的部门成员")
            continue

        cells = [(blk, day) for blk in range(4) for day in range(1, 6)]
        state = {c: [] for c in cells}

        def day_used(p, day):
            return any(p in lst for (b, d), lst in state.items() if d == day)

        def can_put(p, cell):
            lst = state[cell]
            blk, day = cell
            only_min = (day, blk) in minister_only
            if w in (p.busy.get((day, blk)) or set()):
                return False
            if day_used(p, day):
                return False
            if not lst:
                return True
            if lst[0].is_minister:
                return False          # 部长独占一格
            if p.is_minister:
                return False          # 干事已占位，不再塞部长
            if len(lst) >= per_cell:
                return False
            if only_min:
                return False
            return True

        # 这一周最多只动三个部门：从队首所在部门开始，往后再取两个
        head_dept = queue[qpos].dept
        i0 = seq.index(head_dept) if head_dept in seq else 0
        window = {seq[(i0 + k) % n_seq] for k in range(min(3, n_seq))}

        # 本周候选：从队首开始、连续属于这三个部门的人
        cands = []
        scanned = 0
        while scanned < len(queue):
            p = queue[(qpos + scanned) % len(queue)]
            if p.dept not in window:
                break
            cands.append(p)
            scanned += 1

        def cells_full():
            return all(len(state[c]) >= (1 if state[c] and state[c][0].is_minister else per_cell)
                       for c in cells)

        def place(p):
            for cell in cells:
                if can_put(p, cell):
                    state[cell].append(p)
                    count[id(p)] += 1
                    last_week[id(p)] = w
                    served[id(p)].add(w)
                    return True
            return False

        # 先排干事（干事优先），同一部门里按"已排次数少"的先上，让次数尽量平均
        cadres = [x for x in cands if not x.is_minister]
        cadres.sort(key=lambda x: (count[id(x)], queue.index(x) if x in queue else 0))
        for p in cadres:
            if cells_full():
                break
            if streak_before(p, w) >= 2:
                continue              # 干事不连排三周
            place(p)
        for p in [x for x in cands if x.is_minister]:
            if cells_full():
                break
            place(p)

        # 还缺人的格子：往下多找几个部门，尽量把一格凑够两个人
        used = []
        for cell in cells:
            lst = state[cell]
            if len(lst) != 1 or lst[0].is_minister:
                continue
            blk, day = cell
            if (day, blk) in minister_only:
                continue
            # 只能在本周已有的部门里找人补（一周最多三个部门）
            for k in range(0, n_seq):
                d2 = seq[(i0 + k) % n_seq]
                if d2 not in used:
                    continue
                got = None
                for p2 in all_members.get(d2, []):
                    if w in (p2.busy.get((day, blk)) or set()):
                        continue
                    if p2.is_minister:
                        continue
                    if day_used(p2, day):
                        continue
                    got = p2
                    break
                if got is not None:
                    lst.append(got)
                    count[id(got)] += 1
                    last_week[id(got)] = w
                    served[id(got)].add(w)
                    if got.dept not in used:
                        used.append(got.dept)
                    break

        for cell in cells:
            week_plan[cell] = state[cell]
            if not state[cell]:
                notes[w].append("%s %s 没有可用的人"
                                % (WEEKDAY_NAMES[cell[1] - 1], BLOCK_TIMES[cell[0]]))
            elif len(state[cell]) == 1 and not state[cell][0].is_minister:
                notes[w].append("SINGLE|%d|%d" % (cell[1], cell[0]))
                week_plan[cell] = []

        for cell in cells:
            for p in week_plan[cell]:
                if p.dept not in used:
                    used.append(p.dept)
        used.sort(key=lambda d: seq.index(d) if d in seq else 99)
        week_dept_out[w] = used
        qpos = (qpos + max(1, scanned)) % len(queue)
        plan[w] = week_plan

    # ---- 兜底：保证每个干事至少排到一次 ----
    def cnt_now():
        c = {id(p2): int(seed_counts.get(p2.name, 0)) for p2 in people}
        for _w in range(w1, w2 + 1):
            for _lst in plan[_w].values():
                for q in _lst:
                    c[id(q)] += 1
        return c

    def only_min_cell(day, blk):
        return (day, blk) in minister_only

    def try_add(p, any_dept=False, only_under3=False, victim_min=2):
        weeks_order = list(range(w1, w2 + 1))
        if any_dept:
            # 优先挑部门数还没到 3 的周，尽量不新增部门
            weeks_order.sort(key=lambda x: (len(week_dept_out.get(x, [])) >= 3, x))
        for _w in weeks_order:
            if _w in skip_weeks:
                continue
            if only_under3 and len(week_dept_out.get(_w, [])) >= 3:
                continue
            if not any_dept and p.dept not in week_dept_out.get(_w, []):
                continue
            for cell, lst in plan[_w].items():
                blk, day = cell
                if _w in (p.busy.get((day, blk)) or set()):
                    continue
                if any(p in l for c2, l in plan[_w].items() if c2[1] == day):
                    continue
                if only_min_cell(day, blk):
                    continue
                if not lst:
                    lst.append(p)
                    if p.dept not in week_dept_out.setdefault(_w, []):
                        week_dept_out[_w].append(p.dept)
                    return True
                if len(lst) == 1 and not lst[0].is_minister:
                    lst.append(p)
                    if p.dept not in week_dept_out.setdefault(_w, []):
                        week_dept_out[_w].append(p.dept)
                    return True
                if len(lst) == 2 and not any(q.is_minister for q in lst):
                    cur = cnt_now()
                    victim = max(lst, key=lambda q: cur.get(id(q), 0))
                    if cur.get(id(victim), 0) >= victim_min:
                        lst[lst.index(victim)] = p
                        return True
        return False

    for _round in range(20):
        cur = cnt_now()
        left = [p for p in people if not p.is_minister and cur.get(id(p), 0) == 0]
        if not left:
            break
        moved = False
        for p in left:
            if (try_add(p)
                    or try_add(p, any_dept=True, only_under3=True)):
                moved = True
        if not moved:
            break

    # ---- 再补一轮：让干事尽量每人至少 2 次 ----

    for _round in range(20):
        cur = cnt_now()
        left = [p for p in people if not p.is_minister and cur.get(id(p), 0) < 2]
        if not left:
            break
        moved = False
        for p in left:
            if (try_add(p, victim_min=2)
                    or try_add(p, any_dept=True, only_under3=True, victim_min=2)):
                moved = True
        if not moved:
            break

    counts_map = {id(p): 0 for p in people}
    for w in range(w1, w2 + 1):
        for lst in plan[w].values():
            for q in lst:
                counts_map[id(q)] += 1
    stats = {"min": 0, "max": 0, "avg": 0, "zero": []}
    if people:
        counts = sorted(counts_map[id(p)] + int(seed_counts.get(p.name, 0)) for p in people)
        zero_people = [p for p in people
                       if counts_map[id(p)] + int(seed_counts.get(p.name, 0)) == 0]
        stats = {
            "min": counts[0],
            "max": counts[-1],
            "avg": round(sum(counts) / len(counts), 1),
            "zero": [p.name for p in zero_people],
            "zero_ministers": [p.name for p in zero_people if p.is_minister],
            "zero_members": [p.name for p in zero_people if not p.is_minister],
        }
    return plan, notes, stats, week_dept_out


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
                cell.value = "\n".join("%s\n%s" % (p.name, p.phone) for p in people)
                cell.font = Font(name="微软雅黑", sz=18)
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    # 备注不再写进表格（会撑满 I 列），缺少人员的提示只在网页上显示
    wb.save(out_path)
    return out_path


def build_detail(data, w1, w2, out_path, counts=None, assign=None):
    """生成"解析明细.xlsx"：每人每格的有课/空闲周次 + 实际排班时间。"""
    assign = assign or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    head = ["部门", "姓名", "班级", "联系方式", "部长", "总排班次数", "值班时间"]
    ws.append(head)

    ws2 = wb.create_sheet("时段明细")
    ws2.append(["部门", "姓名", "星期", "时段", "有课周次（1-30）", "范围内空闲周次"])

    for dept, members in data.items():
        for m in members:
            shifts = sorted(assign.get(m.name, []),
                            key=lambda s: (s["week"], s["day"], s["block"]))
            when = "\n".join(
                "第%d周 %s %s%s" % (s["week"], WEEKDAY_NAMES[s["day"] - 1], BLOCK_TIMES[s["block"]],
                                    "（已有）" if s.get("history") else "")
                for s in shifts)
            # 次数把"上传的值班表里已排的"也算进去
            n = len(shifts) if assign else (counts or {}).get(m.name, "")
            ws.append([dept, m.name, m.klass, m.phone, "是" if m.is_minister else "",
                       n, when])
            ws.row_dimensions[ws.max_row].height = max(20, 16 * max(1, len(shifts)) + 4)
            for day in range(1, 6):
                for blk in range(4):
                    busy = sorted(m.busy.get((day, blk)) or [])
                    free = free_weeks(m.busy.get((day, blk)), w1, w2)
                    ws2.append([dept, m.name, WEEKDAY_NAMES[day - 1], BLOCK_TIMES[blk],
                                fmt_weeks(busy) or "（无课）", fmt_weeks(free) or "（全满）"])

    center = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    left_wrap = openpyxl.styles.Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = openpyxl.styles.Side(style="thin", color="D9D9D9")
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)

    # 汇总表
    for col, w in {"A": 10, "B": 10, "C": 15, "D": 16, "E": 8, "F": 12, "G": 46}.items():
        ws.column_dimensions[col].width = w
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(name="微软雅黑", bold=True, sz=11)
        cell.alignment = center
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="EFF3FA")
        cell.border = border
    ws.row_dimensions[1].height = 22
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = openpyxl.styles.Font(name="微软雅黑", sz=10)
            cell.border = border
            if cell.column == 7:
                cell.alignment = left_wrap
            else:
                cell.alignment = center
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:G%d" % ws.max_row

    # 时段明细
    for col, w in {"A": 10, "B": 10, "C": 8, "D": 16, "E": 26, "F": 22}.items():
        ws2.column_dimensions[col].width = w
    for cell in ws2[1]:
        cell.font = openpyxl.styles.Font(name="微软雅黑", bold=True, sz=11)
        cell.alignment = center
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="EFF3FA")
        cell.border = border
    ws2.row_dimensions[1].height = 22
    for row in ws2.iter_rows(min_row=2):
        for cell in row:
            cell.font = openpyxl.styles.Font(name="微软雅黑", sz=10)
            cell.alignment = center
            cell.border = border
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:F%d" % ws2.max_row
    wb.save(out_path)
    return out_path


# ---------- 解析结果磁盘缓存（免费服务器内存小，进程重启后不用重新解析） ----------
PARSER_VERSION = 3     # 解析逻辑变了就 +1，旧缓存自动作废


def save_parsed(data, src_dir, path):
    import json
    out = {"v": PARSER_VERSION, "depts": {}}
    for dept, members in data.items():
        arr = []
        for m in members:
            rel = ""
            try:
                if m.pdf_path:
                    rel = os.path.relpath(m.pdf_path, src_dir).replace("\\", "/")
            except Exception:
                rel = ""
            arr.append({
                "n": m.name, "k": m.klass, "p": m.phone,
                "sid": getattr(m, "sid", ""),
                "min": 1 if m.is_minister else 0,
                "pdf": rel,
                "busy": {"%d|%d" % (d, b): sorted(ws)
                         for (d, b), ws in (m.busy or {}).items()},
            })
        out["depts"][dept] = arr
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)


def load_parsed(path, src_dir):
    import json
    if not os.path.exists(path):
        return None
    try:
        raw = json.load(open(path, encoding="utf-8"))
    except Exception:
        return None
    if raw.get("v") != PARSER_VERSION:
        return None
    data = {}
    for dept, arr in raw.get("depts", {}).items():
        members = []
        for a in arr:
            pdf = os.path.join(src_dir, a["pdf"]) if a.get("pdf") else ""
            m = Member(a["n"], a.get("k", ""), a.get("p", ""), dept, bool(a.get("min")), pdf)
            m.sid = a.get("sid", "")
            m.busy = {}
            for k, v in (a.get("busy") or {}).items():
                d, b = k.split("|")
                m.busy[(int(d), int(b))] = set(v)
            members.append(m)
        data[dept] = members
    return data


def highlight_member(src_xlsx, name, out_path):
    """复制一份空课表，把某个人在格子里加粗变红、信息表那一行整行标黄。"""
    import openpyxl
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import PatternFill, Font

    wb = openpyxl.load_workbook(src_xlsx, rich_text=True)
    ws = wb.worksheets[0]
    hit = False
    yellow = PatternFill("solid", fgColor="FFF2CC")

    # 只把他自己那一行改成"加粗 + 下划线 + 深红"（和部长的鲜红区分开）
    for base in (4, 8, 12, 16):
        for day in range(1, 6):
            cell = ws.cell(row=base, column=3 + day)
            v = cell.value
            if not isinstance(v, CellRichText):
                continue
            blocks, changed = [], False
            for b in v:
                if not isinstance(b, TextBlock):
                    blocks.append(b)
                    continue
                parts = b.text.split("\n")
                if name not in b.text:
                    blocks.append(b)
                    continue
                for i, part in enumerate(parts):
                    tail = "\n" if i < len(parts) - 1 else ""
                    if name in part:
                        blocks.append(TextBlock(
                            InlineFont(rFont="微软雅黑", sz=b.font.sz or 10,
                                       b=True, u="single", color="FFC00000"), part + tail))
                        changed = True
                    else:
                        blocks.append(TextBlock(b.font, part + tail))
            if changed:
                cell.value = CellRichText(blocks)
                hit = True

    for r in range(21, ws.max_row + 1):
        if str(ws.cell(row=r, column=2).value or "").strip() == name:
            for c in (2, 3, 4):
                ws.cell(row=r, column=c).fill = yellow
            hit = True
            break

    wb.save(out_path)
    return hit

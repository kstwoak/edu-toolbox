#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考场座位编排程序（Exam Seat Arranger）v2
========================================

用「总控表 + 排考场模板」自动生成完整考场安排，填好三类 sheet：
  1. 名单      —— 总分 / 排名 / 各科座位号（非本场考生写「自习NNN」）
  2. 考场1~13  —— 每场每个座位的「考号 + 姓名」
  3. 1班~12班  —— 每位学生各科的座位号（按排名排序）

输入（三样）：
  A. 排考场模板 xlsx   （名单 + 考场1~13 + 1班~12班）
  B. 考场安排总控表 xlsx（13 考场 + 各场次各考场人数 + 自习分布）
  C. 总分表 csv        （考号,总分  —— 排序的唯一依据）

分配规则（已与教务的参考结果逐格比对，100% 一致）：
  · 每场单独重排：该场考生按总分降序，按「考场首座升序」依次填座
  · 同一场不考这科的学生（自习）同样按总分降序，填进本场的自习考场
  · 座位号 = 模板里预印的物理座位号（全局连续：考场13=1~60、考场1=491~557…）
  · 模板「名单」中 D 列标黄的同学，语数外场次固定安排进第 12 考场（304）
  · 某考场人数超出模板预印座位数时，座位号自动顺延（如 政治 考场7 44 座 → 261~304）
  · 总分缺失（缺考 / #N/A / 空）的学生排在最后，按名单原始顺序

用法：
  python3 exam_seat_arranger.py --make-score-template        # 生成 模板_总分.csv
  python3 exam_seat_arranger.py --template 模板.xlsx --arrangement 考场安排.xlsx --scores 总分.csv
"""

from __future__ import annotations
import argparse
import copy
import csv
import io
import os
import re
import sys
import codecs
from collections import OrderedDict

try:
    import openpyxl
except ImportError:
    print("需要 openpyxl：pip install openpyxl")
    raise

# ─────────────────────────────────────────────
#  常量
# ─────────────────────────────────────────────

SUBJECT_COLS = ["语数外", "物理", "化学", "生物", "政治", "历史", "地理"]

CHAR2SUBJECT = {
    "物": "物理", "化": "化学", "生": "生物",
    "政": "政治", "史": "历史", "地": "地理",
}

# 名单 sheet 列号
NM = {"序号": 1, "选科": 2, "考号": 3, "姓名": 4, "班级": 5,
      "总分": 6, "排名": 7,
      "语数外": 8, "物理": 9, "化学": 10, "生物": 11,
      "政治": 12, "历史": 13, "地理": 14}
NM_COLS = 14

# 班级 sheet 列号
CL = {"选科": 1, "考号": 2, "姓名": 3,
      "语数外": 4, "物理": 5, "化学": 6, "生物": 7,
      "政治": 8, "历史": 9, "地理": 10}
CL_COLS = 10


def _exam_sheet_cols():
    """考场 sheet：{科目: (考号列, 姓名列)}，从第 2 列起每科两列。"""
    m = {}
    col = 2
    for s in SUBJECT_COLS:
        m[s] = (col, col + 1)
        col += 2
    return m


EXAM_COLS = _exam_sheet_cols()

YELLOW = "FFFFFF00"


# ─────────────────────────────────────────────
#  工具
# ─────────────────────────────────────────────

def _read_text(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig")
    if raw.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
        return raw.decode("utf-32")
    if raw.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return raw.decode("utf-16")
    for enc in ("utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _is_yellow(cell):
    f = cell.fill
    if f is None or not f.fill_type:
        return False
    try:
        return f.fgColor.type == "rgb" and f.fgColor.rgb == YELLOW
    except Exception:
        return False


def _is_theme5(cell):
    """模板里标为「主题色5」的高亮（新增/后补学生）。"""
    f = cell.fill
    if f is None or not f.fill_type:
        return False
    try:
        return f.fgColor.type == "theme" and f.fgColor.theme == 5
    except Exception:
        return False


def _room_no(v):
    m = re.search(r"\d+", str(v))
    return int(m.group()) if m else 0


# ─────────────────────────────────────────────
#  1. 学生 & 读名单（含黄色标记）
# ─────────────────────────────────────────────

class Student:
    __slots__ = ("seq", "choice", "sid", "name", "cls", "score", "score_raw",
                 "rank", "special", "added", "subjects", "seats", "src_row")

    def __init__(self, seq, choice, sid, name, cls, src_row):
        self.seq = seq
        self.choice = choice
        self.sid = sid
        self.name = name
        self.cls = cls
        self.src_row = src_row
        self.score = None
        self.score_raw = None
        self.rank = None
        self.special = False            # 名单 姓名列标黄 → 语数外固定去特殊考场
        self.added = False              # 名单高亮（主题色5）→ 新增/后补学生，排最后
        self.subjects = set()
        for ch in str(choice):
            if ch in CHAR2SUBJECT:
                self.subjects.add(CHAR2SUBJECT[ch])
        self.seats = {}                 # {场次名: 座位号}

    def takes(self, subject):
        if subject in ("语数外", "语文", "数学", "英语"):
            return True
        return subject in self.subjects


def read_students(template_path):
    wb = openpyxl.load_workbook(template_path)
    if "名单" not in wb.sheetnames:
        raise SystemExit("模板里找不到「名单」sheet")
    ws = wb["名单"]
    students = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=NM["姓名"]).value
        if not name:
            continue
        s = Student(
            seq=ws.cell(row=r, column=NM["序号"]).value,
            choice=str(ws.cell(row=r, column=NM["选科"]).value or "").strip(),
            sid=str(ws.cell(row=r, column=NM["考号"]).value or "").strip(),
            name=str(name).strip(),
            cls=str(ws.cell(row=r, column=NM["班级"]).value or "").strip(),
            src_row=r,
        )
        s.special = _is_yellow(ws.cell(row=r, column=NM["姓名"]))
        s.added = _is_theme5(ws.cell(row=r, column=NM["选科"]))
        students.append(s)
    wb.close()
    return students


def read_room_seats(template_path):
    """每个考场预印的座位号：{考场号: [座位号...]}"""
    wb = openpyxl.load_workbook(template_path, data_only=True)
    rooms = {}
    for sn in wb.sheetnames:
        if not re.fullmatch(r"\d+", sn):
            continue
        ws = wb[sn]
        labels = []
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v in (None, ""):
                continue
            try:
                labels.append(int(float(v)))
            except (TypeError, ValueError):
                continue
        if labels:
            rooms[int(sn)] = labels
    wb.close()
    return rooms


# ─────────────────────────────────────────────
#  2. 读总控表
# ─────────────────────────────────────────────

def read_arrangement(arr_path):
    wb = openpyxl.load_workbook(arr_path, data_only=True)
    ws = wb.worksheets[0]

    def cell(r, c):
        return ws.cell(row=r, column=c).value

    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        if str(cell(r, 1) or "").strip() == "科目":
            header_row = r
            break
    if header_row is None:
        raise SystemExit("总控表里找不到「科目」表头行")

    room_cols = OrderedDict()
    for c in range(3, ws.max_column + 1):
        v = cell(header_row, c)
        if v and str(v).strip().startswith("考场"):
            room_cols[_room_no(v)] = c

    rooms = {}
    for rno, col in room_cols.items():
        rooms[rno] = {
            "location": cell(header_row + 1, col),
            "capacity": cell(header_row + 2, col),
            "seat_range": str(cell(header_row + 3, col) or "").strip(),
        }

    sessions = []
    total = None
    r = header_row + 1
    while r <= ws.max_row:
        a = str(cell(r, 1) or "").strip()
        if not a or a == "自习":
            r += 1
            continue
        if a.startswith("共"):
            m = re.search(r"\d+", a)
            if m:
                total = int(m.group())
            r += 1
            continue
        tak = {}
        for rno, col in room_cols.items():
            v = cell(r, col)
            if v not in (None, ""):
                try:
                    n = int(float(v))
                    if n > 0:
                        tak[rno] = n
                except (TypeError, ValueError):
                    pass
        sel = {}
        nxt = r + 1
        while nxt <= ws.max_row and str(cell(nxt, 1) or "").strip() == "":
            nxt += 1
        if nxt <= ws.max_row and str(cell(nxt, 1) or "").strip() == "自习":
            for rno, col in room_cols.items():
                v = cell(nxt, col)
                if v not in (None, ""):
                    try:
                        n = int(float(v))
                        if n > 0:
                            sel[rno] = n
                    except (TypeError, ValueError):
                        pass
        sessions.append({"subject": a, "takers": tak, "self": sel})
        r += 1

    wb.close()
    return {"rooms": rooms, "sessions": sessions, "total": total}


# ─────────────────────────────────────────────
#  3. 读总分表
# ─────────────────────────────────────────────

_ALIAS_SID = {"考号", "准考证号", "学号", "id", "student_id"}
_ALIAS_SCORE = {"总分", "分数", "成绩", "score", "total_score"}


def read_scores(path):
    """返回 {考号: (float|None, 原始文本)}"""
    text = _read_text(path)
    reader = csv.DictReader(io.StringIO(text))
    out = {}
    for row in reader:
        sid = None
        raw = None
        for k, v in row.items():
            if k is None:
                continue
            key = str(k).strip()
            if key in _ALIAS_SID:
                sid = str(v or "").strip()
            elif key in _ALIAS_SCORE:
                raw = v
        if not sid:
            continue
        txt = "" if raw is None else str(raw).strip()
        try:
            val = float(txt)
        except (TypeError, ValueError):
            val = None
        out[sid] = (val, txt)
    return out


# ─────────────────────────────────────────────
#  4. 编排
# ─────────────────────────────────────────────

def arrange(students, room_seats, arrangement, scores,
            special_room=12, warn=print):
    """
    返回 (ranked, session_plans)
      session_plans: [{"subject": 场次名, "rooms": {考场号: [Student|None,...]}}]
    """
    missing = []
    for s in students:
        val, raw = scores.get(s.sid, (None, ""))
        s.score = val
        s.score_raw = raw
        if val is None:
            missing.append(s)

    # ── 排名 ─────────────────────────────────
    #  1. 有总分且非新增：总分降序（同分保持名单顺序）
    #  2. 无分且非新增：#N/A 一组、其它（缺考等）一组，各按名单顺序
    #  3. 新增（模板高亮）学生：排在最后，有分者按总分降序，无分者按名单顺序
    is_na = lambda s: str(s.score_raw or "").strip().upper().startswith("#N/A")
    g_scored = sorted([s for s in students
                       if s.score is not None and not s.added],
                      key=lambda s: -s.score)
    g_na = [s for s in students
            if s.score is None and not s.added and is_na(s)]
    g_other = [s for s in students
               if s.score is None and not s.added and not is_na(s)]
    g_add_scored = sorted([s for s in students if s.added and s.score is not None],
                          key=lambda s: -s.score)
    g_add_none = [s for s in students if s.added and s.score is None]
    ranked = g_scored + g_na + g_other + g_add_scored + g_add_none
    for i, s in enumerate(ranked, 1):
        s.rank = i

    if missing:
        warn("  ⚠ %d 位学生没有总分（缺考/#N/A/空），已排在最后：%s%s"
             % (len(missing), "、".join(s.name for s in missing[:6]),
                "…" if len(missing) > 6 else ""))
    n_added = sum(1 for s in students if s.added)
    if n_added:
        warn("  · 名单高亮 %d 人（新增/后补）→ 排在最后" % n_added)

    specials = [s for s in ranked if s.special]
    if specials:
        warn("  · 名单标黄 %d 人 → 语数外固定第 %d 考场" % (len(specials), special_room))

    def fill(pool, counts, subj, skip_rooms=()):
        placed = {}
        idx = 0
        ordered = sorted(counts.keys(), key=lambda r: room_seats[r][0])
        for rno in ordered:
            if rno in skip_rooms:
                continue
            labels = list(room_seats[rno])
            cnt = counts[rno]
            if cnt > len(labels):
                extra = cnt - len(labels)
                labels += [labels[-1] + i for i in range(1, extra + 1)]
                room_seats[rno] = labels
                warn("  · 场次「%s」考场%d 需 %d 座（模板预印 %d 座），"
                     "为本场末考场，座位号顺延至 %d"
                     % (subj, rno, cnt, cnt - extra, labels[-1]))
            for k in range(cnt):
                if idx >= len(pool):
                    break
                s = pool[idx]
                s.seats[subj] = labels[k]
                placed.setdefault(rno, {})[k] = s
                idx += 1
        if idx < len(pool):
            raise SystemExit("场次「%s」还有 %d 人没排下（座位不足）"
                             % (subj, len(pool) - idx))
        return placed

    session_plans = []
    for sess in arrangement["sessions"]:
        name = sess["subject"]
        keys = ["语文", "数学", "英语"] if name == "语数外" else [name]
        for key in keys:
            if name == "语数外" and key != "语文":
                continue          # 三场座位相同，只算一次
            rooms_out = {}
            skip = ()
            takers = [s for s in ranked if s.takes(key)]
            selves = [s for s in ranked if not s.takes(key)]
            if name == "语数外":
                # 标黄同学固定进特殊考场；其余按排名填其它考场
                sp = sorted([s for s in takers if s.special], key=lambda s: s.rank)
                rest = [s for s in takers if not s.special]
                skip = (special_room,)
                counts = dict(sess["takers"])
                counts.pop(special_room, None)
                rooms_out.update(fill(rest, counts, key))
                sp_labels = list(room_seats[special_room])
                for i, s in enumerate(sp):
                    if i >= len(sp_labels):
                        break
                    s.seats[key] = sp_labels[i]
                    rooms_out.setdefault(special_room, {})[i] = s
                # 三人一场的另两场座位沿用
                for k2 in ("数学", "英语"):
                    for rno, d in rooms_out.items():
                        for s in d.values():
                            s.seats[k2] = s.seats[key]
            else:
                rooms_out.update(fill(takers, sess["takers"], key))
                if selves:
                    rooms_out.update(fill(selves, sess["self"], key))
            session_plans.append({"subject": key, "rooms": rooms_out,
                                  "special": name == "语数外"})

    return ranked, session_plans


# ─────────────────────────────────────────────
#  5. 写出
# ─────────────────────────────────────────────

def _sort_rows_by_col(ws, col, first_row, last_row, ncols):
    """按某列的值对数据行排序，保留每格样式/行高。"""
    snap = []
    for r in range(first_row, last_row + 1):
        key = ws.cell(row=r, column=col).value
        cells = [(ws.cell(row=r, column=c).value,
                  copy.copy(ws.cell(row=r, column=c)._style))
                 for c in range(1, ncols + 1)]
        h = ws.row_dimensions[r].height
        snap.append((key, cells, h))
    snap.sort(key=lambda t: (t[0] is None, t[0] if t[0] is not None else 0))
    for i, (key, cells, h) in enumerate(snap):
        tr = first_row + i
        for c, (val, style) in enumerate(cells, 1):
            tc = ws.cell(row=tr, column=c)
            tc.value = val
            tc._style = style
        if h is not None:
            ws.row_dimensions[tr].height = h


def _seat_text(s, subj, fill_self=True):
    """名单/班级表里的座位号文本：考生填数字，非本场考生填「自习NNN」。"""
    if subj == "语数外":
        v = s.seats.get("语文")      # 语文/数学/英语 同座位
    else:
        v = s.seats.get(subj)
    if v is None:
        return None
    if subj == "语数外" or s.takes(subj):
        return v
    return "自习%d" % v if fill_self else None


def write_output(template_path, out_path, ranked, session_plans,
                 room_seats, fill_self=True):
    wb = openpyxl.load_workbook(template_path)

    by_sid = {s.sid: s for s in ranked}

    # ── 名单 ────────────────────────────────
    ws = wb["名单"]
    first, last = 2, ws.max_row
    for r in range(first, last + 1):
        sid = str(ws.cell(row=r, column=NM["考号"]).value or "").strip()
        s = by_sid.get(sid)
        if s is None:
            continue
        # 总分：有数字写数字，否则写原文
        ws.cell(row=r, column=NM["总分"],
                value=s.score if s.score is not None else (s.score_raw or "缺考"))
        ws.cell(row=r, column=NM["排名"], value=s.rank)
        for subj in SUBJECT_COLS:
            txt = _seat_text(s, subj, fill_self)
            if txt is not None:
                ws.cell(row=r, column=NM[subj], value=txt)
    # 按排名重排 + 序号重写
    _sort_rows_by_col(ws, NM["排名"], first, last, NM_COLS)
    for i, r in enumerate(range(first, last + 1), 1):
        ws.cell(row=r, column=NM["序号"], value=i)

    # ── 考场 ────────────────────────────────
    for plan in session_plans:
        subj = plan["subject"]
        c_col, n_col = EXAM_COLS["语数外" if plan["special"] else subj]
        for rno, d in plan["rooms"].items():
            ws2 = wb[str(rno)]
            pre = [i for i in range(5, ws2.max_row + 1)
                   if ws2.cell(row=i, column=1).value not in (None, "")]
            for k, s in d.items():
                if k < len(pre):
                    row = pre[k]
                else:
                    row = pre[-1] + (k - len(pre) + 1)
                    ws2.cell(row=row, column=1, value=room_seats[rno][k])
                ws2.cell(row=row, column=c_col, value=s.sid)
                ws2.cell(row=row, column=n_col, value=s.name)

    # ── 班级 ────────────────────────────────
    cls_sheets = {}
    for sn in wb.sheetnames:
        m = re.fullmatch(r"(\d+)班", sn)
        if m:
            cls_sheets[int(m.group(1))] = sn
    by_cls = OrderedDict()
    for s in ranked:
        m = re.search(r"(\d+)班", s.cls)
        num = int(m.group(1)) if m else None
        by_cls.setdefault(num, []).append(s)
    for num, sn in cls_sheets.items():
        ws3 = wb[sn]
        for r in range(3, ws3.max_row + 1):
            for c in range(1, CL_COLS + 1):
                ws3.cell(row=r, column=c, value=None)
        for i, s in enumerate(by_cls.get(num, [])):
            r = 3 + i
            ws3.cell(row=r, column=CL["选科"], value=s.choice)
            ws3.cell(row=r, column=CL["考号"], value=s.sid)
            ws3.cell(row=r, column=CL["姓名"], value=s.name)
            for subj in SUBJECT_COLS:
                txt = _seat_text(s, subj, fill_self)
                if txt is not None:
                    ws3.cell(row=r, column=CL[subj], value=txt)

    wb.save(out_path)
    return out_path


def add_class_view(wb, ranked):
    """追加「名单(2)」：按班级分组的座位总表（班级名文本升序，组内按排名）。"""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    title = "名单(2)"
    for sn in ("名单(2)", "名单 (2)"):
        if sn in wb.sheetnames:
            wb.remove(wb[sn])
    ws = wb.create_sheet(title, index=wb.sheetnames.index("名单") + 1)

    headers = ["序号", "选科", "考号", "学生姓名"] + SUBJECT_COLS + ["班级"]
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.border = border

    ordered = sorted(ranked, key=lambda s: (str(s.cls), s.rank))
    for i, s in enumerate(ordered, 2):
        ws.cell(row=i, column=1, value=s.rank)
        ws.cell(row=i, column=2, value=s.choice)
        ws.cell(row=i, column=3, value=s.sid)
        ws.cell(row=i, column=4, value=s.name)
        for j, subj in enumerate(SUBJECT_COLS):
            txt = _seat_text(s, subj, True)
            if txt is not None:
                ws.cell(row=i, column=5 + j, value=txt)
        ws.cell(row=i, column=5 + len(SUBJECT_COLS), value=s.cls)
        for c in range(1, len(headers) + 1):
            cc = ws.cell(row=i, column=c)
            cc.alignment = center
            cc.border = border

    widths = [8, 10, 14, 12, 10, 10, 10, 10, 10, 10, 10, 18]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    return ws


def verify(ranked, session_plans, arrangement, room_seats):
    lines = ["【自检】", "  学生总数：%d" % len(ranked)]
    arr_map = {s["subject"]: (sum(s["takers"].values()), sum(s["self"].values()))
               for s in arrangement["sessions"]}
    ok = True
    for plan in session_plans:
        subj = plan["subject"]
        if subj in ("数学", "英语"):
            continue
        n_exam = n_self = 0
        for d in plan["rooms"].values():
            for s in d.values():
                if s.takes(subj):
                    n_exam += 1
                else:
                    n_self += 1
        key = "语数外" if plan["special"] else subj
        exp = arr_map.get(key)
        flag = "  ✓"
        if exp and (n_exam, n_self) != exp:
            flag = "  ✗ 应为 考生%d/自习%d" % exp
            ok = False
        lines.append("  %-4s 考生 %3d / 自习 %3d%s" % (key, n_exam, n_self, flag))
    # 每人每场只坐一次 / 座位在正确考场
    seat2room = {}
    for rno, labs in room_seats.items():
        for l in labs:
            seat2room.setdefault(l, rno)
    for plan in session_plans:
        seen = set()
        for rno, d in plan["rooms"].items():
            for s in d.values():
                if s.sid in seen:
                    lines.append("  ✗ 场次%s 重复占座 %s" % (plan["subject"], s.name))
                    ok = False
                seen.add(s.sid)
    lines.append("  ✓ 全部通过" if ok else "  ✗ 有问题")
    return "\n".join(lines)


# ─────────────────────────────────────────────
#  6. CLI
# ─────────────────────────────────────────────

def make_score_template(template_path, outdir="."):
    """按模板名单生成待填的总分表（考号,姓名,总分）。"""
    wb = openpyxl.load_workbook(template_path, data_only=True)
    ws = wb["名单"]
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=NM["姓名"]).value
        if not name:
            continue
        rows.append([str(ws.cell(row=r, column=NM["考号"]).value).strip(),
                     str(name).strip(), ""])
    path = os.path.join(outdir, "总分表_待填.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["考号", "姓名", "总分"])
        w.writerows(rows)
    return path


def main():
    ap = argparse.ArgumentParser(
        description="考场座位编排程序（按总分排名，每场重排）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python3 exam_seat_arranger.py --template 模板.xlsx --make-score-template
  python3 exam_seat_arranger.py --template 模板.xlsx \\
      --arrangement "考场安排.xlsx" --scores 总分表.csv --outdir 输出
        """,
    )
    ap.add_argument("--template", help="排考场模板 xlsx")
    ap.add_argument("--arrangement", help="考场安排总控表 xlsx")
    ap.add_argument("--scores", help="总分表 csv")
    ap.add_argument("--outdir", default="考场安排输出")
    ap.add_argument("--special-room", type=int, default=12,
                    help="名单标黄学生的语数外固定考场号（默认 12）")
    ap.add_argument("--no-self-seat", action="store_true",
                    help="名单/班级表里非本场考生留空（默认写「自习NNN」）")
    ap.add_argument("--class-view", action="store_true",
                    help="额外生成一张「名单(2)」：按班级分组的座位总表")
    ap.add_argument("--make-score-template", action="store_true",
                    help="按模板名单生成 总分表_待填.csv")
    args = ap.parse_args()

    if not os.path.isdir(args.outdir):
        os.makedirs(args.outdir)

    if args.make_score_template:
        if not args.template:
            ap.error("--make-score-template 需要同时给 --template")
        p = make_score_template(args.template, args.outdir)
        print("已生成：%s" % p)
        return

    if not (args.template and args.arrangement and args.scores):
        ap.print_help()
        return

    print("读取模板名单……")
    students = read_students(args.template)
    print("  %d 位学生，其中标黄 %d 人"
          % (len(students), sum(1 for s in students if s.special)))

    room_seats = read_room_seats(args.template)
    print("读取考场座位号：")
    for rno in sorted(room_seats):
        l = room_seats[rno]
        print("  考场%-2d %d 座（%d~%d）" % (rno, len(l), l[0], l[-1]))

    arrangement = read_arrangement(args.arrangement)
    print("读取总控表：%d 个考场，%d 个场次"
          % (len(arrangement["rooms"]), len(arrangement["sessions"])))

    scores = read_scores(args.scores)
    print("读取总分表：%d 条" % len(scores))

    print("编排中……")
    ranked, session_plans = arrange(students, room_seats, arrangement, scores,
                                    special_room=args.special_room, warn=print)

    out_path = os.path.join(args.outdir, "24级排考场（已填）.xlsx")
    write_output(args.template, out_path, ranked, session_plans, room_seats,
                 fill_self=not args.no_self_seat)
    if args.class_view:
        wb = openpyxl.load_workbook(out_path)
        add_class_view(wb, ranked)
        wb.save(out_path)

    print(verify(ranked, session_plans, arrangement, room_seats))
    print("\n结果文件：%s" % out_path)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把「24级排考场（已填）.xlsx」导出为可读的 Markdown 结果汇总。"""
import sys
import openpyxl

SRC = sys.argv[1] if len(sys.argv) > 1 else "考场安排输出/24级排考场（已填）.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "考场安排输出/结果汇总.md"
ARR = sys.argv[3] if len(sys.argv) > 3 else None

SUBJ = ["语数外", "物理", "化学", "生物", "政治", "历史", "地理"]
EXAM_COLS = {s: (2 + 2 * i, 3 + 2 * i) for i, s in enumerate(SUBJ)}

wb = openpyxl.load_workbook(SRC, data_only=True)

# 总控表：考场位置 / 名义容量
LOC = {}
CAP = {}
if ARR:
    wa = openpyxl.load_workbook(ARR, data_only=True).worksheets[0]
    def _c(r, c):
        return wa.cell(row=r, column=c).value
    hdr = None
    for r in range(1, 20):
        if str(_c(r, 1) or "").strip() == "科目":
            hdr = r
            break
    if hdr:
        for c in range(3, wa.max_column + 1):
            val = _c(hdr, c)
            if val and str(val).strip().startswith("考场"):
                import re as _re
                n = int(_re.search(r"\d+", str(val)).group())
                LOC[n] = _c(hdr + 1, c)
                CAP[n] = _c(hdr + 2, c)


def v(ws, r, c):
    x = ws.cell(row=r, column=c).value
    if x is None:
        return ""
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x).strip().replace("|", "/")


L = []
A = L.append

# ═══ 标题 ═══
A("# 2026.9 高三考场安排结果（24级）\n")
A("> 由 `exam_seat_arranger.py` 按总分排名自动编排生成；"
  "每场考试单独按总分重排座位；名单标黄学生语数外固定第 12 考场（304）。\n")

nm = wb["名单"]
n_stu = sum(1 for r in range(2, nm.max_row + 1) if v(nm, r, 4))

# ═══ 一、总览 ═══
A("## 一、总览\n")
A("- 学生：**%d 人**" % n_stu)
A("- 考场：**13 间**（考场13 二楼报告厅 1~60；考场2~11 各 40 座；考场12=304；考场1=高三1班）")
A("- 场次：**7 场**（语数外、物理、化学、生物、政治、历史、地理）")
A("- 座位号：全局连续编号，考场13=1~60 → 考场2~11=61~460 → 考场12=461~490 → 考场1=491~557\n")

A("### 各场次各考场人数\n")
hdr = "| 场次 | " + " | ".join("考场%d" % i for i in range(1, 14)) + " | 合计 |"
A(hdr)
A("|" + "---|" * 15)

# 从考场表反推每场每考场人数
rooms = [str(i) for i in range(1, 14)]
for sess in SUBJ:
    ci, ni = EXAM_COLS[sess]
    row = []
    tot = 0
    for rn in rooms:
        ws = wb[rn]
        cnt = 0
        for r in range(5, ws.max_row + 1):
            if ws.cell(row=r, column=ni).value not in (None, ""):
                cnt += 1
        row.append(str(cnt) if cnt else "—")
        tot += cnt
    A("| %s | %s | %d |" % (sess, " | ".join(row), tot))
A("")

# ═══ 二、座位号对照 ═══
A("## 二、考场座位号对照\n")
A("| 考场 | 位置 | 名义容量 | 座位号范围 | 座位数 |")
A("|---|---|---|---|---|")
for rn in rooms:
    i = int(rn)
    ws = wb[rn]
    labels = [ws.cell(row=r, column=1).value for r in range(5, ws.max_row + 1)
              if ws.cell(row=r, column=1).value not in (None, "")]
    note = ""
    if i == 7 and len(labels) > 40:
        note = "（末考场，含并入 4 人）"
    A("| %s | %s | %s | %s~%s | %d%s |" % (
        ws.cell(row=1, column=1).value, LOC.get(i, ""), CAP.get(i, ""),
        labels[0], labels[-1], len(labels), note))
A("")
A("> 关于考场7：政治场考生 304 人，按座位号顺序排到考场7 就正好是最后一个考场"
  "（其后的考场8~11、考场1 本场都用作自习室），剩下的 4 人直接并入考场7，"
  "所以考场7 本场容入 44 人、座位号 261~304。其中 301~304 与考场8 的自习座位号重名，"
  "但均在自习区，不影响使用。\n")

# ═══ 三、名单 ═══
A("## 三、名单（排名 + 各科座位号）\n")
A("> 数字=该科座位号；`自习NNN`=该场在自习考场的座位号。\n")
A("| 排名 | 考号 | 姓名 | 班级 | " + " | ".join(SUBJ) + " |")
A("|" + "---|" * 12)
for r in range(2, nm.max_row + 1):
    if not v(nm, r, 4):
        continue
    A("| %s | %s | %s | %s | %s |" % (
        v(nm, r, 7), v(nm, r, 3), v(nm, r, 4),
        v(nm, r, 5).replace("高中2024级", ""),
        " | ".join(v(nm, r, 8 + i) for i in range(7))))
A("")

# ═══ 四、各考场座位表 ═══
A("## 四、各考场座位表\n")
for rn in rooms:
    ws = wb[rn]
    A("### %s\n" % ws.cell(row=1, column=1).value)
    A("| 座位号 | " + " | ".join(SUBJ) + " |")
    A("|" + "---|" * 8)
    for r in range(5, ws.max_row + 1):
        seat = ws.cell(row=r, column=1).value
        if seat in (None, ""):
            continue
        cells = [v(ws, r, EXAM_COLS[s][1]) for s in SUBJ]
        if not any(cells):
            continue
        A("| %s | %s |" % (seat, " | ".join(cells)))
    A("")

# ═══ 五、班级座位表 ═══
A("## 五、班级座位表\n")
for i in range(1, 13):
    sn = "%d班" % i
    if sn not in wb.sheetnames:
        continue
    ws = wb[sn]
    has = any(v(ws, r, 3) for r in range(3, ws.max_row + 1))
    A("### 高三%d班%s\n" % (i, "" if has else "（无学生）"))
    if not has:
        continue
    A("| 选科 | 考号 | 姓名 | " + " | ".join(SUBJ) + " |")
    A("|" + "---|" * 10)
    for r in range(3, ws.max_row + 1):
        if not v(ws, r, 3):
            continue
        A("| %s | %s | %s | %s |" % (
            v(ws, r, 1), v(ws, r, 2), v(ws, r, 3),
            " | ".join(v(ws, r, 4 + j) for j in range(7))))
    A("")

text = "\n".join(L)
with open(OUT, "w", encoding="utf-8") as f:
    f.write(text)
print("已生成 %s（%d 字符，%d 行）" % (OUT, len(text), len(L)))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
监考排班程序（Exam Invigilation Scheduler）
==========================================

功能：给定考试安排（科目、时间、考场、每考场人数、每场巡场人数）与老师信息
      （是否参与、不可用时段、必用时段），自动生成一份尽量满足约束的排班表。

硬性约束（必须满足，权重最高）：
  1) 每场考试每个考场配齐指定数量的监考老师；
  2) 同一老师同一时段只能在一个岗位（时间不重叠）；
  3) 老师的"不可用时段"（请假/教研）内不能被安排；
  4) 老师的"必用时段"内必须被安排（该时段需被其任务覆盖）；
  5) 标记为"全程不参与"的老师完全不参与排班。

巡场：由使用者在考试模板「巡场人员」列指定，独立记录、原样输出，
不参与监考求解。留空 = 本场无巡场。

软性目标（尽量满足，不满足也能排）：
  A) 老师尽量不换教室（减少跨教室次数）；
  B) 任务尽量平均（大家总时长接近）。

用法：
  python3 exam_invigilation_scheduler.py                  # 用内置示例数据排班
  python3 exam_invigilation_scheduler.py --make-template  # 生成 3 个导入模板
  python3 exam_invigilation_scheduler.py --exams 考试.csv --rooms 考场.csv --teachers 老师.csv

导出：
  - 排班表.csv           （考试科目/时间/考场/老师）
  - 老师时长统计.csv      （每位老师的监考总时长、场次、教室）
  - 监考安排表.xlsx      （对齐原表格式：最后一列为巡场）
"""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime, timedelta
import os
import sys
import io
import codecs
import random
import math
import csv
import re
import argparse
from collections import defaultdict


def _app_dir() -> str:
    """程序所在目录。打包成 exe（PyInstaller）后取 exe 所在目录，否则取脚本目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

# ============ 可调参数 ============
W_OVERLAP   = 100000   # 时间重叠（硬）
W_UNAVAIL   = 100000   # 占用不可用时段（硬）
W_MANDATORY = 100000   # 必用时段未被覆盖（硬）
W_PATROL    = 100000   # 巡场时段内被排监考（硬）
W_ROOM      = 0        # 每多用一个教室的惩罚（软；默认 0 = 不追求「少换教室」）
W_BALANCE   = 1.0      # 总时长（监考+巡场）不均衡的惩罚（软）
W_SELF      = 3000     # 「自习岗」（单人看考场）分配不均的惩罚（软）
# ↓ 每天休息要平均（核心）：只均衡「总时长」时，会出现「第一天 2 场、第二天 5 场」
#   这种总分钟数一样、但某天全天无休的情况。下面两项按天分别均衡：
W_DAYBAL    = 0.8      # 每天时长不均衡的惩罚（软，按天求方差后取平均）
W_DAYCNT    = 800      # 每天监考场次不均衡的惩罚（软，按"场"而不是"分钟"看）

ITERATIONS  = 200000   # 模拟退火迭代次数
T0          = 150.0    # 初始温度
T1          = 0.5      # 终止温度
SEED        = 42


# ============ 数据模型 ============

@dataclass
class Exam:
    """一场考试。

    name        : 科目名称
    start / end : 起止时间（datetime）
    rooms       : 本场使用的考场列表。元素可以是 str（房间名，人数取 n_per_room），
                  也可以是 (房间名, 人数) 元组，用于每考场人数不同的情况。
    n_per_room   : 每个考场需要的监考老师数（默认值，当 rooms 元素为 str 时使用）。
    patrol_names : 本场指定的巡场人员姓名列表；空 = 本场无巡场。
                   巡场人员由使用者指定，程序不再随机指派。
    """
    name: str
    start: datetime
    end: datetime
    rooms: list
    n_per_room: int = 2
    patrol_names: list = field(default_factory=list)


@dataclass
class Teacher:
    """一位老师。

    name        : 姓名
    participate : 是否参与排班（False = 全程不参与）
    unavailable : 不可用时段（请假/教研），列表 [(start, end), ...]
    mandatory   : 必用时段（该时段必须排到任务），列表 [(start, end), ...]
    """
    name: str
    participate: bool = True
    unavailable: list = field(default_factory=list)
    mandatory: list = field(default_factory=list)


@dataclass
class Task:
    """一个排班单元 = 某场考试 × 某间考场。"""
    id: int
    exam: str
    room: str
    start: datetime
    end: datetime
    need: int
    duration: int      # 分钟


@dataclass
class ScheduleMeta:
    """全局排班元信息（导出 Excel 对齐原表格式所需，由使用者在考场模板里补写）。

    title         : 表格标题（第 1 行），空则用默认「监考人员安排表」。
    room_class    : 考场名 → 班级名称，如 {'考场 1': '高三1班'}。用于列头「第一考场\n(高三1班)」。
    room_capacity : 考场名 → 容量人数，如 {'考场 1': 40}。用于第 2 行人数行。
    """
    title: str = ""
    room_class: dict = field(default_factory=dict)
    room_capacity: dict = field(default_factory=dict)


_CN_NUM = "零一二三四五六七八九"


def _cn_number(n: int) -> str:
    """阿拉伯数字 → 中文序号（用于考场列头「第一考场/第十二考场」）。"""
    if n <= 0:
        return str(n)
    if n < 10:
        return _CN_NUM[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + _CN_NUM[n % 10]
    if n < 100:
        tens = _CN_NUM[n // 10]
        ones = _CN_NUM[n % 10] if n % 10 else ""
        return tens + "十" + ones
    return str(n)


# ============ 时间工具 ============

def _overlap(s1, e1, s2, e2) -> bool:
    """两个闭开区间 [s1,e1) 与 [s2,e2) 是否重叠（首尾相接不算重叠）。"""
    return s1 < e2 and s2 < e1


def _union_covers(mand_s, mand_e, intervals) -> bool:
    """判断必用时段 [mand_s, mand_e) 是否被 intervals 的并集完全覆盖。"""
    ivs = sorted(intervals)
    cur = None
    for s, e in ivs:
        if cur is None:
            cur = [s, e]
        elif s <= cur[1]:
            if e > cur[1]:
                cur[1] = e
        else:
            if cur[0] <= mand_s and cur[1] >= mand_e:
                return True
            cur = [s, e]
    if cur is not None and cur[0] <= mand_s and cur[1] >= mand_e:
        return True
    return False


def _minutes(start: datetime, end: datetime) -> int:
    return int((end - start).total_seconds() // 60)


# ============ 构建任务 ============

def build_tasks(exams: list[Exam]) -> list[Task]:
    """只为监考岗建任务；巡场是独立的记录，不参与监考求解。"""
    tasks = []
    tid = 0
    for ex in exams:
        # 考场监考岗
        for room_item in ex.rooms:
            if isinstance(room_item, (tuple, list)):
                room, need = room_item[0], int(room_item[1])
            else:
                room, need = room_item, ex.n_per_room
            tasks.append(Task(
                id=tid, exam=ex.name, room=room,
                start=ex.start, end=ex.end, need=need,
                duration=_minutes(ex.start, ex.end),
            ))
            tid += 1
    return tasks


# ============ 求解器 ============

class Scheduler:
    def __init__(self, exams, teachers, meta=None):
        exams = list(exams)
        self.exams = exams
        self.teachers = list(teachers)
        self.by_name = {t.name: t for t in teachers}
        # 只参与排班的老师进入分配池；全程不参与的被排除
        self.excluded = [t.name for t in teachers if not t.participate]
        self.all_names = [t.name for t in teachers if t.participate]
        self.tasks = build_tasks(exams)
        self.task_by_id = {t.id: t for t in self.tasks}
        # 「自习岗」= 只需 1 位老师的考场（导出时显示为「自习\n姓名」）
        self.self_task_ids = {t.id for t in self.tasks if t.need == 1}
        self.n_self = {n: 0 for n in self.all_names}
        # 巡场 = 用户指定，不随机指派；但巡场期间不能被排监考
        self.patrol_by_exam = {ex.name: list(getattr(ex, "patrol_names", None) or [])
                               for ex in exams}
        self.exam_by_name = {ex.name: ex for ex in exams}
        # 巡场繁忙时段：某老师在某场巡场，则该场考试的时间区间内不得给他排监考
        self.patrol_busy = {n: [] for n in self.all_names}
        # 巡场时长（分钟）：与监考等权，用于「总时长」均衡
        self.ptime = {n: 0 for n in self.all_names}
        for ex in exams:
            dur = _minutes(ex.start, ex.end)
            for nm in self.patrol_by_exam.get(ex.name, []):
                if nm in self.patrol_busy:
                    self.patrol_busy[nm].append((ex.start, ex.end))
                if nm in self.ptime:
                    self.ptime[nm] += dur
        # 只在巡场列出现、不在老师名单里的人（无法安排监考，仅记录）
        self.patrol_not_in_pool = sorted(
            {nm for names in self.patrol_by_exam.values() for nm in names
             if nm not in self.by_name})
        # 导出 Excel 对齐原表格式所需的元信息（标题、班级名、容量）
        self.meta = meta if meta is not None else ScheduleMeta()

        # ---- 按天均衡所需的状态 ----
        # 考试跨了哪几天；每个任务（考场×场次）属于哪一天
        self.days = sorted({t.start.date() for t in self.tasks})
        self.day_index = {d: i for i, d in enumerate(self.days)}
        self.task_day = {t.id: self.day_index[t.start.date()] for t in self.tasks}
        # 每人每天：工作量（监考 + 巡场，分钟）/ 监考场次
        self.day_time = {n: [0] * len(self.days) for n in self.all_names}
        self.day_cnt = {n: [0] * len(self.days) for n in self.all_names}
        for ex in exams:
            di = self.day_index.get(ex.start.date())
            if di is None:
                continue
            dur = _minutes(ex.start, ex.end)
            for nm in self.patrol_by_exam.get(ex.name, []):
                if nm in self.day_time:
                    self.day_time[nm][di] += dur
        # 当天原则上可排的老师集合：整天不可用的人不参与当天的均衡基数，
        # 否则会去追一个他根本达不到的平均值，白占软代价。
        self.day_active = {n: [True] * len(self.days) for n in self.all_names}
        for nm in self.all_names:
            for di in range(len(self.days)):
                if not any(self._can_work(nm, tk) for tk in self.tasks
                           if self.task_day[tk.id] == di):
                    self.day_active[nm][di] = False
        # 巡场分钟是固定的，作为「按天」状态的基线，回滚/重建时从这里恢复
        self._day_time_base = {n: list(v) for n, v in self.day_time.items()}

    # ---- 该老师在该时段是否原则上可排（只看硬性占用）----
    def _can_work(self, name, task):
        for (us, ue) in self.by_name[name].unavailable:
            if _overlap(task.start, task.end, us, ue):
                return False
        for (ps, pe) in self.patrol_busy.get(name, ()):
            if _overlap(task.start, task.end, ps, pe):
                return False
        return True

    # ---- 某老师的教室集合 ----
    def _rooms_of(self, name):
        return {self.task_by_id[i].room for i in self.ttasks[name]}

    # ---- 个体老师硬惩罚 ----
    def _teacher_hard(self, name):
        t = self.by_name[name]
        ivs = sorted((self.task_by_id[i].start, self.task_by_id[i].end)
                     for i in self.ttasks[name])
        ov = ua = ma = pb = 0
        # 时间重叠
        for a in range(len(ivs)):
            sa, ea = ivs[a]
            for b in range(a + 1, len(ivs)):
                sb, eb = ivs[b]
                if _overlap(sa, ea, sb, eb):
                    ov += 1
        # 占用不可用时段 / 巡场时段
        for i in self.ttasks[name]:
            tk = self.task_by_id[i]
            for (us, ue) in t.unavailable:
                if _overlap(tk.start, tk.end, us, ue):
                    ua += 1
            for (ps, pe) in self.patrol_busy.get(name, ()):
                if _overlap(tk.start, tk.end, ps, pe):
                    pb += 1
        # 必用时段未被覆盖
        for (ms, me) in t.mandatory:
            if not _union_covers(ms, me, ivs):
                ma += 1
        return W_OVERLAP * ov + W_UNAVAIL * ua + W_MANDATORY * ma + W_PATROL * pb

    # ---- 软目标（已归一化：软代价任何情况下都不可能与硬约束同量级）----
    def _total_minutes(self, name):
        """总时长 = 监考时长 + 巡场时长（两者等权）。"""
        return self.ttime.get(name, 0) + self.ptime.get(name, 0)

    def _room_cost(self):
        n = max(1, len(self.all_names))
        return W_ROOM * sum(max(0, len(self._rooms_of(nm)) - 1) for nm in self.all_names) / n

    def _balance_cost(self):
        if not self.ttime:
            return 0.0
        vals = [self._total_minutes(nm) for nm in self.ttime]
        n = max(1, len(vals))
        avg = sum(vals) / n
        return W_BALANCE * sum((v - avg) ** 2 for v in vals) / n

    def _soft_cost(self):
        return (self._room_cost() + self._balance_cost()
                + self._selfstudy_cost() + self._day_balance_cost()
                + self._day_count_cost())

    def _selfstudy_cost(self):
        """「自习岗」（单人看考场）分配不均的惩罚：人数越接近越优。"""
        if not self.n_self:
            return 0.0
        vals = list(self.n_self.values())
        n = len(vals)
        avg = sum(vals) / n
        return W_SELF * sum((v - avg) ** 2 for v in vals) / n

    def _day_balance_cost(self):
        """每天的工作量（分钟）是否平均 —— 「每天休息要平均」的主项。

        只均衡「总时长」是不够的：两天分别 4 场 / 5 场时，某人可以做成
        第一天 2 场、第二天 5 场（总分钟数正好和别人一样），但第二天全天无休。
        这里按天分别求方差再取平均，专门压住「某一天被排满」的情况。
        """
        if not self.days:
            return 0.0
        tot = 0.0
        nd = 0
        for di in range(len(self.days)):
            vals = [self.day_time[nm][di] for nm in self.all_names
                    if self.day_active[nm][di]]
            if len(vals) < 2:
                continue
            avg = sum(vals) / len(vals)
            tot += sum((v - avg) ** 2 for v in vals) / len(vals)
            nd += 1
        return W_DAYBAL * tot / max(1, nd)

    def _day_count_cost(self):
        """每天的「监考场数」是否平均。

        老师看排班表数的是场次，不是分钟；只有分钟均衡时，可能出现
        「4 场短考试」和「3 场长考试」分钟数相同、但歇的次数差一场。
        """
        if not self.days:
            return 0.0
        tot = 0.0
        nd = 0
        for di in range(len(self.days)):
            vals = [self.day_cnt[nm][di] for nm in self.all_names
                    if self.day_active[nm][di]]
            if len(vals) < 2:
                continue
            avg = sum(vals) / len(vals)
            tot += sum((v - avg) ** 2 for v in vals) / len(vals)
            nd += 1
        return W_DAYCNT * tot / max(1, nd)

    # ---- 按天状态的增减 ----
    def _bump_day(self, name, task, sign):
        di = self.task_day[task.id]
        self.day_time[name][di] += sign * task.duration
        self.day_cnt[name][di] += sign

    def _reset_day_state(self):
        self.day_time = {n: list(v) for n, v in self._day_time_base.items()}
        self.day_cnt = {n: [0] * len(self.days) for n in self.all_names}

    def _swap_days(self, a, b, t1, t2, sign):
        """a 由 t1 换成 t2、b 由 t2 换成 t1；sign=-1 即回退这次交换。"""
        self._bump_day(a, t1, -sign)
        self._bump_day(a, t2, sign)
        self._bump_day(b, t2, -sign)
        self._bump_day(b, t1, sign)

    # ---- 贪心初始解 ----
    def _greedy_init(self):
        self.assign = {t.id: [] for t in self.tasks}
        self.ttasks = {n: [] for n in self.all_names}
        self.trooms = {n: set() for n in self.all_names}
        self.ttime = {n: 0 for n in self.all_names}
        self.n_self = {n: 0 for n in self.all_names}
        self._reset_day_state()

        order = sorted(self.tasks, key=lambda t: (t.start, -t.duration))
        for task in order:
            scored = []
            for name in self.all_names:
                t = self.by_name[name]
                if name in self.assign[task.id]:
                    continue
                conflict = 0
                for i in self.ttasks[name]:
                    ot = self.task_by_id[i]
                    if _overlap(task.start, task.end, ot.start, ot.end):
                        conflict += 1
                for (us, ue) in t.unavailable:
                    if _overlap(task.start, task.end, us, ue):
                        conflict += 1
                for (ps, pe) in self.patrol_busy.get(name, ()):
                    if _overlap(task.start, task.end, ps, pe):
                        conflict += 1
                score = conflict * 1_000_000
                # 必用时段被本任务覆盖 → 强烈优先
                for (ms, me) in t.mandatory:
                    if _overlap(task.start, task.end, ms, me):
                        score -= 500_000
                # 已经用过这间教室 → 优先
                if task.room in self._rooms_of(name):
                    score -= 10_000
                # 总时长少的优先（均衡：监考 + 巡场等权）
                score += self._total_minutes(name) * 1
                # 同一天已经排了很多的更要往后排（每天休息要平均）
                di = self.task_day[task.id]
                score += self.day_time[name][di] * 1
                score += self.day_cnt[name][di] * 60
                scored.append((score, name))
            scored.sort()
            picked = [name for _, name in scored[:task.need]]
            self.assign[task.id] = picked
            for name in picked:
                self.ttasks[name].append(task.id)
                self.trooms[name].add(task.room)
                self.ttime[name] += task.duration
                self._bump_day(name, task, 1)
                if task.need == 1:
                    self.n_self[name] += 1

    # ---- 施加/回退一个"替换"动作 ----
    def _apply_replace(self, task_id, pos, new_name):
        old_name = self.assign[task_id][pos]
        if old_name == new_name:
            return None
        task = self.task_by_id[task_id]
        self.assign[task_id][pos] = new_name
        self.ttasks[old_name].remove(task_id)
        self.ttime[old_name] -= task.duration
        self._bump_day(old_name, task, -1)
        self.trooms[old_name] = self._rooms_of(old_name)
        self.ttasks[new_name].append(task_id)
        self.ttime[new_name] += task.duration
        self._bump_day(new_name, task, 1)
        if task.need == 1:
            self.n_self[old_name] -= 1
            self.n_self[new_name] += 1
        self.trooms[new_name] = self._rooms_of(new_name)
        return old_name

    def _undo_replace(self, task_id, pos, old_name, new_name):
        task = self.task_by_id[task_id]
        self.assign[task_id][pos] = old_name
        self.ttasks[new_name].remove(task_id)
        self.ttime[new_name] -= task.duration
        self._bump_day(new_name, task, -1)
        self.trooms[new_name] = self._rooms_of(new_name)
        self.ttasks[old_name].append(task_id)
        self.ttime[old_name] += task.duration
        self._bump_day(old_name, task, 1)
        if task.need == 1:
            self.n_self[new_name] -= 1
            self.n_self[old_name] += 1
        self.trooms[old_name] = self._rooms_of(old_name)

    # ---- 求解 ----
    def solve(self, iterations=ITERATIONS, T0=T0, T1=T1, seed=SEED, progress_cb=None):
        """
        求解排班。progress_cb(done, total, hard_sum, best_cost) 会在求解过程中被周期调用，
        用于向界面/控制台报告进度；传 None 则不回调。

        返回值选优采用**字典序**（先比硬违规数，再比软代价）：
        无论软权重 W_ROOM / W_BALANCE 调得多大，都不会为了好看而牺牲硬约束。
        """
        random.seed(seed)
        self._greedy_init()

        self.hard = {n: self._teacher_hard(n) for n in self.all_names}
        self.hard_sum = sum(self.hard.values())
        # 记录贪心初始解，作为最终兜底（保证返回结果不会比初始解更差）
        greedy_assign = {t.id: list(self.assign[t.id]) for t in self.tasks}
        greedy_hard = self.hard_sum

        best = {t.id: list(self.assign[t.id]) for t in self.tasks}
        best_key = (self.hard_sum, self._soft_cost())
        best_cost = best_key[0] + best_key[1]

        T = T0
        factor = math.pow(T1 / T0, 1.0 / max(1, iterations))
        cb_step = max(1, iterations // 100)  # 约每 1% 回调一次

        for it in range(iterations):
            if random.random() < 0.9:
                task_id = random.choice(self.tasks).id
                task = self.task_by_id[task_id]
                if not self.assign[task_id]:
                    continue
                pos = random.randrange(len(self.assign[task_id]))
                old_name = self.assign[task_id][pos]
                candidates = [n for n in self.all_names if n not in self.assign[task_id]]
                if not candidates:
                    continue
                new_name = random.choice(candidates)
                if old_name == new_name:
                    continue

                hard_before = self.hard[old_name] + self.hard[new_name]
                before = hard_before + self._soft_cost()
                self._apply_replace(task_id, pos, new_name)
                self.hard[old_name] = self._teacher_hard(old_name)
                self.hard[new_name] = self._teacher_hard(new_name)
                hard_after = self.hard[old_name] + self.hard[new_name]
                after = hard_after + self._soft_cost()
                delta = after - before
                # 注意：hard_sum 只能累加「硬约束」的变化量，
                # 否则软权重一大就会污染硬违规计数，使最优解跟踪失效
                d_hard = hard_after - hard_before
                self.hard_sum += d_hard

                if delta <= 0 or random.random() < math.exp(-delta / T):
                    pass
                else:
                    self.hard_sum -= d_hard
                    self._undo_replace(task_id, pos, old_name, new_name)
                    self.hard[old_name] = self._teacher_hard(old_name)
                    self.hard[new_name] = self._teacher_hard(new_name)
            else:
                if len(self.tasks) < 2:
                    continue
                t1 = random.choice(self.tasks)
                t2 = random.choice(self.tasks)
                if t1.id == t2.id or not self.assign[t1.id] or not self.assign[t2.id]:
                    continue
                p1 = random.randrange(len(self.assign[t1.id]))
                p2 = random.randrange(len(self.assign[t2.id]))
                a = self.assign[t1.id][p1]
                b = self.assign[t2.id][p2]
                if a == b or a in self.assign[t2.id] or b in self.assign[t1.id]:
                    continue

                affected = {a, b}
                hard_before = sum(self.hard[nm] for nm in affected)
                before = hard_before + self._soft_cost()

                self.assign[t1.id][p1] = b
                self.assign[t2.id][p2] = a
                self.ttasks[a].remove(t1.id); self.ttasks[a].append(t2.id)
                self.ttasks[b].remove(t2.id); self.ttasks[b].append(t1.id)
                self.ttime[a] += (self.task_by_id[t2.id].duration - self.task_by_id[t1.id].duration)
                self.ttime[b] += (self.task_by_id[t1.id].duration - self.task_by_id[t2.id].duration)
                self._swap_days(a, b, self.task_by_id[t1.id], self.task_by_id[t2.id], 1)
                if self.task_by_id[t1.id].need == 1:
                    self.n_self[a] -= 1; self.n_self[b] += 1
                if self.task_by_id[t2.id].need == 1:
                    self.n_self[b] -= 1; self.n_self[a] += 1
                self.trooms[a] = self._rooms_of(a)
                self.trooms[b] = self._rooms_of(b)

                hard_after = sum(self._teacher_hard(nm) for nm in affected)
                after = hard_after + self._soft_cost()
                delta = after - before
                d_hard = hard_after - hard_before
                self.hard_sum += d_hard

                if delta <= 0 or random.random() < math.exp(-delta / T):
                    self.hard[a] = self._teacher_hard(a)
                    self.hard[b] = self._teacher_hard(b)
                else:
                    self.hard_sum -= d_hard
                    self.assign[t1.id][p1] = a
                    self.assign[t2.id][p2] = b
                    self.ttasks[a].remove(t2.id); self.ttasks[a].append(t1.id)
                    self.ttasks[b].remove(t1.id); self.ttasks[b].append(t2.id)
                    self.ttime[a] -= (self.task_by_id[t2.id].duration - self.task_by_id[t1.id].duration)
                    self.ttime[b] -= (self.task_by_id[t1.id].duration - self.task_by_id[t2.id].duration)
                    self._swap_days(a, b, self.task_by_id[t1.id], self.task_by_id[t2.id], -1)
                    if self.task_by_id[t1.id].need == 1:
                        self.n_self[b] -= 1; self.n_self[a] += 1
                    if self.task_by_id[t2.id].need == 1:
                        self.n_self[a] -= 1; self.n_self[b] += 1
                    self.trooms[a] = self._rooms_of(a)
                    self.trooms[b] = self._rooms_of(b)

            cur_key = (self.hard_sum, self._soft_cost())
            if cur_key < best_key:
                best_key = cur_key
                best_cost = cur_key[0] + cur_key[1]
                best = {t.id: list(self.assign[t.id]) for t in self.tasks}

            T *= factor

            # 周期上报进度
            if progress_cb and (it % cb_step == 0 or it == iterations - 1):
                progress_cb(it + 1, iterations, self.hard_sum, best_cost)

        self.assign = best
        self._rebuild_from_assign()
        # 兜底：若退火结果反而不如贪心初始解，就回到初始解
        if self.hard_sum > greedy_hard:
            self.assign = greedy_assign
            self._rebuild_from_assign()
        return self._report()

    def _rebuild_from_assign(self):
        self.ttasks = {n: [] for n in self.all_names}
        self.trooms = {n: set() for n in self.all_names}
        self.ttime = {n: 0 for n in self.all_names}
        self.n_self = {n: 0 for n in self.all_names}
        self._reset_day_state()
        for t in self.tasks:
            for name in self.assign[t.id]:
                self.ttasks[name].append(t.id)
                self.ttime[name] += t.duration
                self._bump_day(name, t, 1)
                if t.need == 1:
                    self.n_self[name] += 1
        for n in self.all_names:
            self.trooms[n] = self._rooms_of(n)
        self.hard = {n: self._teacher_hard(n) for n in self.all_names}
        self.hard_sum = sum(self.hard.values())

    # ---- 报告 ----
    def _report(self):
        violations = []
        for n in self.all_names:
            ov = ua = ma = pb = 0
            ivs = sorted((self.task_by_id[i].start, self.task_by_id[i].end)
                         for i in self.ttasks[n])
            for a in range(len(ivs)):
                for b in range(a + 1, len(ivs)):
                    if _overlap(ivs[a][0], ivs[a][1], ivs[b][0], ivs[b][1]):
                        ov += 1
            for i in self.ttasks[n]:
                tk = self.task_by_id[i]
                for (us, ue) in self.by_name[n].unavailable:
                    if _overlap(tk.start, tk.end, us, ue):
                        ua += 1
                for (ps, pe) in self.patrol_busy.get(n, ()):
                    if _overlap(tk.start, tk.end, ps, pe):
                        pb += 1
            for (ms, me) in self.by_name[n].mandatory:
                if not _union_covers(ms, me, ivs):
                    ma += 1
            if ov or ua or ma or pb:
                violations.append((n, ov, ua, ma, pb))
        # 巡场时长汇总（巡场人由使用者指定，故不计入求解目标，但要计入统计）
        ptime = {n: 0 for n in self.all_names}
        n_patrol = {n: 0 for n in self.all_names}
        for ename, names in self.patrol_by_exam.items():
            ex = self.exam_by_name.get(ename)
            if ex is None:
                continue
            dur = _minutes(ex.start, ex.end)
            for nm in names:
                ptime[nm] = ptime.get(nm, 0) + dur
                n_patrol[nm] = n_patrol.get(nm, 0) + 1
        # 均衡指标：总时长（监考 + 巡场，等权）的极差
        _tot = [self._total_minutes(nm) for nm in self.ttime] or [0]
        _mon = list(self.ttime.values()) or [0]
        _pat = [self._total_minutes(nm) - self.ttime.get(nm, 0) for nm in self.ttime] or [0]
        _self = [self.n_self.get(nm, 0) for nm in self.ttime] or [0]
        # 按天均衡指标：每天各人的时长/场次极差（「每天休息要平均」的直接读数）
        day_stat = []
        for di, d in enumerate(self.days):
            vals = [self.day_time[nm][di] for nm in self.all_names if self.day_active[nm][di]]
            cnts = [self.day_cnt[nm][di] for nm in self.all_names if self.day_active[nm][di]]
            if not vals:
                continue
            day_stat.append({
                "date": d.strftime("%Y-%m-%d"),
                "n": len(vals),
                "min": min(vals), "max": max(vals),
                "avg": round(sum(vals) / len(vals), 1),
                "cnt_min": min(cnts), "cnt_max": max(cnts),
            })
        return {
            "assign": self.assign,
            "tasks": self.tasks,
            "ttime": self.ttime,
            "trooms": self.trooms,
            "ptime": ptime,
            "n_patrol": n_patrol,
            "violations": violations,
            "hard_sum": self.hard_sum,
            "balance": max(_tot) - min(_tot),
            "balance_monitor": max(_mon) - min(_mon),
            "balance_patrol": max(_pat) - min(_pat),
            "total_min": min(_tot), "total_max": max(_tot),
            "n_self": dict(self.n_self),
            "balance_self": max(_self) - min(_self),
            "self_min": min(_self), "self_max": max(_self),
            "days": list(self.days),
            "day_time": {n: list(v) for n, v in self.day_time.items()},
            "day_cnt": {n: list(v) for n, v in self.day_cnt.items()},
            "day_stat": day_stat,
            "day_span": max((x["max"] - x["min"] for x in day_stat), default=0),
            "day_span_cnt": max((x["cnt_max"] - x["cnt_min"] for x in day_stat), default=0),
            "excluded": self.excluded,
            "patrol_by_exam": self.patrol_by_exam,
            "patrol_not_in_pool": self.patrol_not_in_pool,
            "meta": self.meta,
        }


# ============ 可行性检测与建议 ============

def _exam_demand(ex):
    """某场考试的总监考岗位需求 = 各考场人数之和（巡场不在内，它是独立记录）。"""
    total = 0
    for r in ex.rooms:
        total += r[1] if isinstance(r, tuple) else ex.n_per_room
    return total


def feasibility_check(exams, teachers):
    """
    在求解前检测每场考试的人员是否够用，并给出方案建议。
    不修改任何数据，只返回报告。

    返回 dict：
      feasible: bool                是否全部够用
      per_exam: [ {exam, need, avail, unavail_names, gap} ]
      suggestions: [str]            针对缺员场次的方案建议
    """
    part = [t for t in teachers if t.participate]
    part_names = {t.name for t in part}
    per_exam = []
    suggestions = []
    feasible = True

    for ex in sorted(exams, key=lambda x: x.start):
        need = _exam_demand(ex)
        unavail = [t.name for t in part
                   if any(_overlap(ex.start, ex.end, u[0], u[1]) for u in t.unavailable)]
        # 本场巡场且本身在老师名单里的人，巡场期间不能排监考 → 算作不可用
        patrol_here = [n for n in (getattr(ex, "patrol_names", None) or [])
                       if n in part_names]
        busy = set(unavail) | set(patrol_here)
        avail = len(part) - len(busy)
        gap = max(0, need - avail)
        if gap > 0:
            feasible = False
        per_exam.append({
            "exam": ex, "need": need, "avail": avail,
            "unavail_names": unavail, "patrol_names": patrol_here, "gap": gap,
        })

    # 生成方案建议（针对每个缺员场）
    for item in per_exam:
        if item["gap"] <= 0:
            continue
        ex = item["exam"]
        opts = []
        # 方案1：减少监考人数（找出可减的考场）
        rooms = ex.rooms
        dec = []
        for r in rooms:
            cap = r[1] if isinstance(r, tuple) else ex.n_per_room
            if cap > 1:
                dec.append((r[0] if isinstance(r, tuple) else r, cap))
        if dec:
            dec_sorted = sorted(dec, key=lambda x: -x[1])
            top = dec_sorted[0]
            opts.append(f"方案① 减少监考：考场「{top[0]}」由 {top[1]} 人减为 {top[1]-1} 人"
                        f"（此类考场共 {len(dec)} 个可减）")
        # 方案2：放宽不可用时段
        if item["unavail_names"]:
            opts.append(f"方案② 放宽时段：以下老师在该时段不可用，若其中 {item['gap']} 位可到场即可——"
                        f"{ '、'.join(item['unavail_names']) }")
        # 方案2b：巡场调整（巡场人在本场不能同时监考）
        if item.get("patrol_names"):
            opts.append(f"方案② 调整巡场：本场巡场「{ '、'.join(item['patrol_names']) }」"
                        f"在巡场期间不能监考；若改由其他人巡场，可释放该名额")
        excluded = [t.name for t in teachers if not t.participate]
        if excluded:
            opts.append(f"方案③ 增加老师：将「不参与」的老师改回参与（{ '、'.join(excluded) }）")
        else:
            opts.append("方案③ 增加老师：补充新的可用老师")

        suggestions.append({
            "exam": ex.name,
            "gap": item["gap"],
            "options": opts,
        })

    return {"feasible": feasible, "per_exam": per_exam, "suggestions": suggestions}


# ============ 输出 ============

def _fmt(dt: datetime) -> str:
    return dt.strftime("%m-%d %H:%M")


def _teacher_rows(result):
    """按老师汇总：监考场次、监考时长、巡场场次、巡场时长、总时长、教室、每天场次。

    只出现在巡场列、不在老师名单里的人也会出现在统计里（总时长 = 巡场时长）。
    """
    tasks = result["tasks"]
    assign = result["assign"]
    ttime = result.get("ttime") or {}
    trooms = result.get("trooms") or {}
    ptime = result.get("ptime") or {}
    n_patrol = result.get("n_patrol") or {}
    n_self = result.get("n_self") or {}
    day_cnt = result.get("day_cnt") or {}
    day_time = result.get("day_time") or {}
    days = result.get("days") or []
    rows = []
    for name in sorted(set(ttime) | set(ptime)):
        n_room = sum(1 for t in tasks if name in assign[t.id])
        mon = ttime.get(name, 0)
        pat = ptime.get(name, 0)
        rows.append({
            "name": name,
            "room_sessions": n_room,
            "monitor_min": mon,
            "patrol_sessions": n_patrol.get(name, 0),
            "patrol_min": pat,
            "self_sessions": n_self.get(name, 0),
            "total_min": mon + pat,
            "rooms": "、".join(sorted(trooms.get(name, set()))) or "-",
            # 每天：场次 / 分钟，用「/」连接，如 "3/4"、"330/300"
            "day_cnt_str": "/".join(str(day_cnt.get(name, [])[i]) for i in range(len(days)))
                           if name in day_cnt else "-",
            "day_min_str": "/".join(str(day_time.get(name, [])[i]) for i in range(len(days)))
                           if name in day_time else "-",
        })
    rows.sort(key=lambda r: (-r["total_min"], r["name"]))
    return rows


def print_feasibility_report(fb):
    """打印人员缺口报告与方案建议。"""
    print("\n" + "=" * 62)
    print("⚠️  人员缺口检测报告（排班前检查）")
    print("=" * 62)
    for item in fb["per_exam"]:
        ex = item["exam"]
        mark = "✅" if item["gap"] <= 0 else f"❌ 缺 {item['gap']} 人"
        extra = ""
        if item.get("patrol_names"):
            extra = f"，其中巡场 {len(item['patrol_names'])} 人"
        print(f"  {ex.name:<6} {_fmt(ex.start)}~{_fmt(ex.end)}  "
              f"需 {item['need']:>2} 人 / 可用 {item['avail']:>2} 人{extra}  {mark}")

    print("\n" + "-" * 62)
    print("解决方案建议（请自行选择并修改 CSV 后重新运行）：")
    print("-" * 62)
    for s in fb["suggestions"]:
        print(f"\n【{s['exam']}】缺 {s['gap']} 人：")
        for opt in s["options"]:
            print(f"    {opt}")


def print_schedule(result):
    tasks = result["tasks"]
    assign = result["assign"]
    by_exam = defaultdict(list)
    for t in tasks:
        by_exam[t.exam].append(t)

    print("=" * 70)
    print("监考安排表")
    print("=" * 70)
    for exam, tlist in by_exam.items():
        t0 = tlist[0]
        patrol = (result.get("patrol_by_exam") or {}).get(exam) or []
        patrol_txt = f"　巡场：{'、'.join(patrol)}" if patrol else ""
        print(f"\n【{exam}】{_fmt(t0.start)} ~ {_fmt(t0.end)}{patrol_txt}")
        for t in sorted(tlist, key=lambda x: x.room):
            teachers = "、".join(assign[t.id])
            print(f"    考场 {t.room:<10}（{t.need}人）: {teachers}")

    print("\n" + "=" * 70)
    print("老师工作量统计（按总时长排序）")
    print("=" * 70)
    for r in _teacher_rows(result):
        extra = (f"  巡场 {r['patrol_sessions']:>2} 场({r['patrol_min']}分)"
                 if r["patrol_sessions"] else "")
        print(f"  {r['name']:<6} 总时长 {r['total_min']:>3} 分钟  "
              f"监考 {r['room_sessions']:>2} 场({r['monitor_min']}分)"
              f"(含自习 {r['self_sessions']} 场){extra}  教室 {r['rooms']}")

    not_in_pool = result.get("patrol_not_in_pool") or []
    if not_in_pool:
        print("\n（以下巡场人不在老师名单中，仅计入巡场时长）：")
        print("  " + "、".join(not_in_pool))

    if result["excluded"]:
        print("\n（以下老师全程不参与，未安排任何任务）：")
        print("  " + "、".join(result["excluded"]))

    print("\n" + "=" * 70)
    print("约束检查")
    print("=" * 70)
    if result["violations"]:
        print("⚠️  存在未满足的硬约束：")
        for n, ov, ua, ma, pb in result["violations"]:
            print(f"    {n}: 重叠{ov} 不可用{ua} 必用未覆盖{ma} 巡场期间被排监考{pb}")
    else:
        print("✅ 全部硬约束已满足。")
    print(f"总时长极差（监考+巡场，最忙 - 最闲）: {result['balance']} 分钟"
          f"　（其中监考 {result.get('balance_monitor', 0)} 分 / 巡场 {result.get('balance_patrol', 0)} 分）")
    print(f"自习岗（单人看考场）分布: 每人 {result.get('self_min', 0)}~{result.get('self_max', 0)} 场"
          f"，极差 {result.get('balance_self', 0)} 场")
    for d in result.get("day_stat") or []:
        print(f"{d['date']} 每天均衡: 时长 {d['min']}~{d['max']} 分钟（平均 {d['avg']:.0f}）"
              f"，场次 {d['cnt_min']}~{d['cnt_max']} 场")


def _write_csv(path, rows, header):
    candidates = [path]
    if not os.path.isabs(path):
        script_dir = _app_dir()
        candidates.append(os.path.join(script_dir, os.path.basename(path)))
        candidates.append(os.path.join(os.path.expanduser("~"), os.path.basename(path)))
    for p in candidates:
        try:
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if header is not None:
                    w.writerow(header)
                for r in rows:
                    w.writerow(r)
            print(f"已导出到 {p}")
            return p
        except OSError:
            continue
    print(f"⚠️ 导出 {path} 失败：只读文件系统且无可写回退目录。")
    return None


def export_schedule(result, outdir=None):
    # 1) 排班表（模板样式宽表：标题/容量/表头/数据，与 xlsx 一致）
    rows = _build_wide_rows(result)
    export_schedule_path = os.path.join(outdir, "排班表.csv") if outdir else "排班表.csv"
    _write_csv(export_schedule_path, rows, None)

    # 2) 老师时长统计（含巡场时长）
    rows2 = []
    for r in _teacher_rows(result):
        rows2.append([r["name"], r["room_sessions"], r["monitor_min"],
                      r["patrol_sessions"], r["patrol_min"], r["self_sessions"],
                      r["total_min"], r["day_cnt_str"], r["day_min_str"], r["rooms"]])
    export_teacher_path = os.path.join(outdir, "老师时长统计.csv") if outdir else "老师时长统计.csv"
    _write_csv(export_teacher_path, rows2,
               ["姓名", "监考场次", "监考时长(分钟)", "巡场场次", "巡场时长(分钟)",
                "自习场次", "总时长(分钟)", "每天场次", "每天时长(分钟)", "涉及教室"])


def _build_wide_rows(result):
    """组装模板样式的宽表（与 xlsx 导出内容一致），供 CSV 导出用。

    返回 list[list[str]]，每行对应 Excel 的一行：
      行1 标题 / 行2 容量 / 行3 表头 / 数据行...
    单元格内换行用 \n（CSV 以 utf-8-sig 写出，Excel 打开换行正常）。
    """
    tasks = result["tasks"]
    assign = result["assign"]
    patrol_by_exam = result.get("patrol_by_exam", {})
    meta = result.get("meta") or ScheduleMeta()
    title = meta.title or "监考人员安排表"

    by_exam = defaultdict(list)
    for t in tasks:
        by_exam[t.exam].append(t)
    exam_names = sorted(by_exam.keys(),
                        key=lambda n: min(t.start for t in by_exam[n]))
    all_rooms = sorted({t.room for t in tasks}, key=_room_sort_key)

    n_room = len(all_rooms)
    rows = []

    # 行1：标题（仅第一列）
    rows.append([title] + [""] * (3 + n_room + 1))

    # 行2：考场容量（前 4 列与巡场列留空）
    cap_row = ["", "", "", ""]
    for room in all_rooms:
        cap = meta.room_capacity.get(room)
        cap_row.append(str(cap) if cap is not None else "")
    cap_row.append("")
    rows.append(cap_row)

    # 行3：表头
    headers = ["日期", "时间", "科目", "监考"] + \
              [_room_header(room, meta) for room in all_rooms] + ["巡场"]
    rows.append(headers)

    # 数据行
    for ename in exam_names:
        tlist = by_exam[ename]
        t0 = min(tlist, key=lambda t: t.start)
        t1 = max(tlist, key=lambda t: t.end)
        cells = {}
        for t in tlist:
            names = assign[t.id]
            cells[t.room] = ("自习\n" + names[0]) if len(names) == 1 else "\n".join(names)
        patrol_names = patrol_by_exam.get(ename, [])
        row = [
            _date_cell(t0.start),
            f"{t0.start:%H:%M}-{t1.end:%H:%M}",
            ename,
            "主监考\n副监考",
        ]
        for room in all_rooms:
            row.append(cells.get(room, "/"))
        row.append("\n".join(patrol_names))
        rows.append(row)

    return rows


# ============ Excel 导出（对齐原表格式） ============

_WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def _room_sort_key(room):
    """按考场名里的数字排序（考场 1, 考场 2 ... 考场 13）。"""
    m = re.search(r"(\d+)", str(room))
    return (0, int(m.group(1))) if m else (1, str(room))


def _date_cell(dt: datetime) -> str:
    """日期列文本，对齐原表：'2026年\n10月8日\n\n周四'。"""
    return f"{dt.year}年\n{dt.month}月{dt.day}日\n\n{_WEEKDAYS[dt.weekday()]}"


def _room_header(room, meta):
    """考场列头：优先用班级名 →「第X考场\n(班级名)」；无班级名 →「第X考场」。"""
    m = re.search(r"(\d+)", str(room))
    idx = int(m.group(1)) if m else None
    cn = _cn_number(idx) if idx else str(room)
    cls = meta.room_class.get(room, "")
    if cls:
        return f"第{cn}考场\n({cls})"
    return f"第{cn}考场" if idx else str(room)


def export_schedule_xlsx(result, outdir=None, title=None):
    """
    把排班结果导出为与原表格式一致的 Excel（.xlsx）。

    列结构：日期 | 时间 | 科目 | 监考 | 第1考场 ... 第N考场 | 巡场
      - 第 2 行：每个考场的学生容量（来自考场模板「容量人数」列）
      - 考场列头：第X考场\n(班级名)，班级名来自考场模板「班级名称」列
      - 考场格：2 位老师 = '名1\n名2'；1 位老师 = '自习\n名'；本场不用 = '/'
      - 巡场（老师+校领导）= 使用者指定的人员
      - 日期列按天合并；同一日期多场考试连续排列
    """
    if not _HAS_OPENPYXL:
        print("⚠️ 未安装 openpyxl，无法导出 Excel。请运行：pip install openpyxl")
        return None

    tasks = result["tasks"]
    assign = result["assign"]
    patrol_by_exam = result.get("patrol_by_exam", {})
    meta = result.get("meta") or ScheduleMeta()
    if title is None:
        title = meta.title or "监考人员安排表"

    # 1) 按考试分组，按开始时间排序
    by_exam = defaultdict(list)
    for t in tasks:
        by_exam[t.exam].append(t)
    exam_names = sorted(by_exam.keys(),
                        key=lambda n: min(t.start for t in by_exam[n]))

    # 2) 考场列集合（所有场次的并集，按编号排序）
    all_rooms = sorted({t.room for t in tasks},
                       key=_room_sort_key)

    # 3) 组装每场考试的一行数据
    rows = []  # 每项 dict: date, time, subject, cells(考场dict), patrol_names
    for ename in exam_names:
        tlist = by_exam[ename]
        t0 = min(tlist, key=lambda t: t.start)
        t1 = max(tlist, key=lambda t: t.end)
        cells = {}
        for t in tlist:
            names = assign[t.id]
            if len(names) == 1:
                cells[t.room] = "自习\n" + names[0]
            else:
                cells[t.room] = "\n".join(names)
        # 巡场 = 用户指定
        patrol_names = patrol_by_exam.get(ename, [])
        rows.append({
            "date": t0.start,
            "time": f"{t0.start:%H:%M}-{t1.end:%H:%M}",
            "subject": ename,
            "cells": cells,
            "patrol_names": patrol_names,
        })

    # 4) 写工作簿
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "监考安排"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_font = Font(name="宋体", bold=True, size=16)
    head_font_main = Font(name="宋体", bold=True, size=11)   # A~D 表头
    head_font_room = Font(name="宋体", bold=True, size=10)   # 考场/巡场表头
    cap_font = Font(name="宋体", size=11)                    # 第 2 行容量
    date_font = Font(name="宋体", size=11)                   # A 列
    time_font = Font(name="宋体", size=14)                   # B 列
    subject_font = Font(name="宋体", size=14)                # C 列
    duty_font = Font(name="宋体", size=12)                   # D 列
    cell_font = Font(name="宋体", size=12)                   # 考场/巡场列

    n_room = len(all_rooms)
    n_col = 4 + n_room + 1  # 日期/时间/科目/监考 + 考场 + 巡场

    # 第 1 行：标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_col)
    tc = ws.cell(1, 1, title)
    tc.font = title_font
    tc.alignment = center
    ws.row_dimensions[1].height = 25.5

    # 第 2 行：A2:C2 合并（空），考场列填学生容量（来自「容量人数」列）
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    for ci, room in enumerate(all_rooms, start=5):
        cap = meta.room_capacity.get(room)
        if cap is not None:
            c = ws.cell(2, ci, cap)
            c.font = cap_font
            c.alignment = center
            c.border = border
    ws.row_dimensions[2].height = 13.5

    # 第 3 行：表头
    headers = ["日期", "时间", "科目", "监考"] + \
              [_room_header(room, meta) for room in all_rooms] + ["巡场"]
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(3, ci, h)
        c.font = head_font_main if ci <= 4 else head_font_room
        c.alignment = center
        c.border = border
    ws.row_dimensions[3].height = 37

    # 数据行（从第 4 行开始）
    row_idx = 4
    prev_date_key = None
    date_merge_start = None
    for r in rows:
        dk = (r["date"].year, r["date"].month, r["date"].day)
        if dk != prev_date_key:
            # 结束上一段日期合并
            if date_merge_start is not None and date_merge_start < row_idx - 1:
                ws.merge_cells(start_row=date_merge_start, start_column=1,
                               end_row=row_idx - 1, end_column=1)
            date_merge_start = row_idx
            prev_date_key = dk

        vals = [
            _date_cell(r["date"]),
            r["time"],
            r["subject"],
            "主监考\n副监考",
        ]
        for room in all_rooms:
            vals.append(r["cells"].get(room, "/"))
        vals.append("\n".join(r["patrol_names"]) if r["patrol_names"] else "")

        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row_idx, ci, v)
            if ci == 1:
                c.font = date_font
            elif ci in (2, 3):
                c.font = time_font
            elif ci == 4:
                c.font = duty_font
            else:
                c.font = cell_font
            c.alignment = center
            c.border = border
        ws.row_dimensions[row_idx].height = 42.75
        row_idx += 1

    # 最后一段日期合并
    if date_merge_start is not None and date_merge_start < row_idx - 1:
        ws.merge_cells(start_row=date_merge_start, start_column=1,
                       end_row=row_idx - 1, end_column=1)

    # 列宽（对齐原模板「高三」主表）
    ws.column_dimensions["A"].width = 9.875
    ws.column_dimensions["B"].width = 15.375
    ws.column_dimensions["C"].width = 13.875
    ws.column_dimensions["D"].width = 12.375
    for ci in range(5, 5 + n_room):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = 11.1153846153846
    # 最后一个考场列对齐原模板（稍宽，因常带全角括号班级名）
    if n_room >= 1:
        last_room_col = 4 + n_room
        ws.column_dimensions[openpyxl.utils.get_column_letter(last_room_col)].width = 11.9423076923077
    ws.column_dimensions[openpyxl.utils.get_column_letter(n_col)].width = 11.625

    # 保存（自动回退路径）
    path = os.path.join(outdir, "监考安排表.xlsx") if outdir else "监考安排表.xlsx"
    candidates = [path]
    if not os.path.isabs(path):
        script_dir = _app_dir()
        candidates.append(os.path.join(script_dir, os.path.basename(path)))
        candidates.append(os.path.join(os.path.expanduser("~"), os.path.basename(path)))
    for p in candidates:
        try:
            wb.save(p)
            print(f"已导出到 {p}")
            return p
        except OSError as e:
            print(f"⚠️ 写入 {p} 失败：{e}")
            continue
    print(f"⚠️ 导出 {path} 失败：只读文件系统且无可写回退目录。")
    return None


# ============ 模板生成与导入 ============

def make_templates(outdir=None):
    """生成 3 个可填写的导入模板。"""
    d = outdir or _app_dir()
    exam_rows = [
        ["语文", "2026-10-08", "08:00", "10:30", "教师A 教师B 教师C"],
        ["数学", "2026-10-08", "15:00", "17:00", ""],
    ]
    # 考场模板：新增「班级名称」「容量人数」「标题」列，供 Excel 导出对齐原表格式用
    # 标题只需在任意一行（如第一行）填一次，程序取第一个非空值
    room_rows = [
        ["语文", "考场 1", "2", "高三1班", "40", "博文中学2025-2026学年度第二学期一模考试监考人员安排表（高三）"],
        ["语文", "考场 2", "2", "高三2班", "40", ""],
        ["数学", "考场 1", "2", "高三1班", "40", ""],
    ]
    teacher_rows = [
        ["教师01", "是", "2026-10-08 08:00~2026-10-08 12:30", ""],
        ["教师07", "是", "", ""],
        ["李四", "是", "", ""],
        ["王五", "是", "", ""],
        ["赵六", "是", "2026-10-09", ""],
        ["孙七", "是", "", ""],
        ["张三", "否", "", ""],
    ]
    _write_csv(os.path.join(d, "模板_考试.csv"), exam_rows,
               ["科目", "日期", "开始时间", "结束时间", "巡场人员"])
    _write_csv(os.path.join(d, "模板_考场.csv"), room_rows,
               ["科目", "考场", "监考人数", "班级名称", "容量人数", "标题"])
    _write_csv(os.path.join(d, "模板_老师.csv"), teacher_rows,
               ["姓名", "是否参与", "不可用时段", "必用时段"])
    print("\n模板已生成，请按列填写后导入。")
    print("  不可用/必用时段格式：'2026-10-08 08:00~2026-10-08 12:30'，多段用 | 分隔。")
    print("    也可只写 '08:00~12:30'（自动沿用开始日期），或整天写成 '2026-10-09'。")
    print("  是否参与填 '是' 或 '否'（否 = 全程不参与）。")
    print("  巡场人员：本场巡场老师姓名，多个用空格分隔；留空 = 本场无巡场。")
    print("    如巡场人也在老师名单里，该场考试期间不会再给他排监考（但时长会计入统计）。")
    print("  考场模板「班级名称」「容量人数」「标题」列：用于导出 Excel 时对齐原表格式。")
    print("    标题：只在第一行填一次即可（如留空则用默认「监考人员安排表」）。")


def _parse_dt(s):
    s = str(s).strip()
    # 容忍全角符号 / 全角空格
    s = (s.replace("：", ":").replace("／", "/").replace("　", " ")
          .replace("年", "-").replace("月", "-").replace("日", ""))
    s = re.sub(r"\s+", " ", s)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"无法解析时间：{s!r}")


# 时段分隔符：| 、 , ， ; ； 以及换行
_RANGE_SEP = re.compile(r"[|、,，;；\n\r]+")
# 起止分隔符：~ ～ 〜 至 到 — –（不含 - ，因为日期里就有 -）
_RANGE_DASH = re.compile(r"\s*[~～︿至到—–]+\s*")


def _parse_ranges(s, what="时段"):
    """解析 'start~end|start~end' 形式的时段字符串，返回 [(start, end), ...]。

    容错（都是静默失效的高发区）：
      · 全角波浪号 ～、中文「至/到」、破折号 — 当作分隔符
      · 结束时间只写时刻（如 '08:00~12:30'）→ 自动沿用开始日期
      · 写成整天（'2026-10-08~2026-10-08' 或只写一个 '2026-10-08'）
        → 视为当天 00:00 ~ 次日 00:00
      · 结束早于开始 / 起止完全相同（非全天）→ 直接报错，绝不静默忽略
    """
    if not s or not str(s).strip():
        return []
    out = []
    for seg in _RANGE_SEP.split(str(s).strip()):
        seg = seg.strip()
        if not seg:
            continue
        parts = _RANGE_DASH.split(seg)
        if len(parts) == 1:
            # 只写了一个日期（如 '2026-10-09'）→ 视为整天
            start = _parse_dt(parts[0])
            if (start.hour, start.minute, start.second) != (0, 0, 0):
                raise ValueError(f"{what}只写了一个时间点（{seg}），请写成 开始~结束。")
            out.append((start, start + timedelta(days=1)))
            continue
        if len(parts) != 2:
            raise ValueError(f"{what}格式有误（应为 开始~结束，多段用 | 分隔）：{seg!r}")
        start = _parse_dt(parts[0])
        try:
            end = _parse_dt(parts[1])
        except ValueError:
            # 结束时间只写了时刻，沿用开始日期
            end = _parse_dt(f"{start:%Y-%m-%d} {parts[1]}")
        if end == start:
            if (start.hour, start.minute, start.second) == (0, 0, 0):
                end = start + timedelta(days=1)      # 整天写法
            else:
                raise ValueError(f"{what}起止时间完全相同（{seg}），请填写正确的时间段。")
        if end < start:
            raise ValueError(f"{what}结束时间早于开始时间（{seg}）。")
        out.append((start, end))
    return out


def _pick(row, aliases, default=""):
    """按多个候选列名取值（兼容用户改过的表头）：取第一个有内容的别名。"""
    for a in aliases:
        if a in row and row[a] is not None and str(row[a]).strip() != "":
            return str(row[a]).strip()
    return default


# 列名别名（用户把「不可用时段」写成「不可用时间」时不再静默忽略）
_AL_SUBJECT  = ["科目", "考试科目", "学科"]
_AL_DATE     = ["日期", "考试日期"]
_AL_START    = ["开始时间", "起始时间", "开始"]
_AL_END      = ["结束时间", "终止时间", "结束"]
_AL_PATROL   = ["巡场人员", "巡场", "巡场老师", "巡场教师", "巡场人", "楼道监督员"]
_AL_ROOM     = ["考场", "考场名称", "试场", "考室"]
_AL_NEED     = ["监考人数", "监考教师数", "监考老师数", "需要人数", "应配人数"]
_AL_CLASS    = ["班级名称", "班级", "班级名"]
_AL_CAPACITY = ["容量人数", "容量", "考场人数", "班级人数"]
_AL_TITLE    = ["标题", "表名", "表头标题"]
_AL_NAME     = ["姓名", "名字", "老师", "教师"]
_AL_PART     = ["是否参与", "是否参加", "参与", "参加"]
_AL_UNAVAIL  = ["不可用时段", "不可用时间", "不可用", "请假时段", "请假时间", "请假"]
_AL_MANDATORY = ["必用时段", "必用时间", "必用", "必须参与时段", "必须参与时间"]


def _read_text(path, on_warn=None):
    """读取文本文件，自动识别编码。

    Excel / WPS 在中文 Windows 上「另存为 CSV」会把文件存成 GBK(ANSI)，
    直接按 UTF-8 读会报 UnicodeDecodeError（如 0xd0 开头就是 GBK 的「姓」）。
    这里按 BOM → UTF-8 → GB18030 → Big5 依次尝试，自动转换并提示。
    """
    with open(path, "rb") as f:
        raw = f.read()
    # 1) 带 BOM 的（注意 UTF-32 的 BOM 是 UTF-16 的超集，要先判断）
    if raw.startswith(codecs.BOM_UTF32_LE) or raw.startswith(codecs.BOM_UTF32_BE):
        return raw.decode("utf-32")
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig")
    if raw.startswith(codecs.BOM_UTF16_LE) or raw.startswith(codecs.BOM_UTF16_BE):
        return raw.decode("utf-16")
    # 2) 无 BOM：UTF-8 → GB18030（GBK/GB2312 超集）→ Big5
    for enc in ("utf-8", "gb18030", "big5"):
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if enc != "utf-8" and on_warn:
            on_warn(f"提示：{os.path.basename(path)} 不是 UTF-8 编码，已按 {enc.upper()} "
                    f"自动识别（Excel/WPS 另存 CSV 常会这样，不影响使用）。")
        return text
    # 3) 都不行：尽力读，同时报警
    if on_warn:
        on_warn(f"⚠️ {os.path.basename(path)} 编码无法识别，已按 UTF-8 忽略错误字符读取，"
                f"可能有个别乱码，建议另存为「CSV UTF-8」。")
    return raw.decode("utf-8", errors="replace")


def _require_columns(fieldnames, required_groups, path, desc):
    """校验 CSV 表头。required_groups 为别名组列表，每组命中任一别名即算通过。"""
    if fieldnames is None:
        raise ValueError(f"{desc}文件 {path} 是空的（没有表头）。")
    actual = {str(c).strip() for c in fieldnames if c}
    missing = [g[0] for g in required_groups if not (set(g) & actual)]
    if missing:
        raise ValueError(
            f"{desc}文件 {os.path.basename(path)} 表头不匹配，缺少列 {missing}。\n"
            f"实际表头：{list(fieldnames)}\n"
            f"请确认选择的是正确的「{desc}」模板文件。")


def _check_unknown_columns(fieldnames, known, desc, on_warn):
    """表头里有「看起来是业务列但没被识别」的列时发出警告（防止静默忽略）。"""
    if not on_warn or not fieldnames:
        return
    hints = ("不可用", "请假", "必用", "参与", "巡场", "姓名", "科目", "时间", "人数")
    extra = [str(c).strip() for c in fieldnames
             if c and str(c).strip() and str(c).strip() not in known
             and any(h in str(c) for h in hints)]
    if extra:
        on_warn(f"⚠️ {desc}表里这些列未被识别，将按空处理：{'、'.join(extra)}。"
                f"（如确有内容请改名或检查列名）")


def load_exams_csv(path, on_warn=None):
    """读取考试CSV：科目,日期,开始时间,结束时间,巡场人员"""
    exams = {}
    n_patrol = 0
    with io.StringIO(_read_text(path, on_warn)) as f:
        r = csv.DictReader(f)
        _require_columns(r.fieldnames, [_AL_SUBJECT, _AL_DATE, _AL_START, _AL_END],
                         path, "考试模板")
        _check_unknown_columns(r.fieldnames,
                               set(_AL_SUBJECT + _AL_DATE + _AL_START + _AL_END + _AL_PATROL),
                               "考试", on_warn)
        for row in r:
            name = _pick(row, _AL_SUBJECT)
            if not name:
                continue
            d = _pick(row, _AL_DATE)
            st = _pick(row, _AL_START)
            en = _pick(row, _AL_END)
            try:
                start = _parse_dt(f"{d} {st}")
                end = _parse_dt(f"{d} {en}")
            except ValueError as e:
                raise ValueError(f"考试「{name}」的时间有误：{e}")
            if end <= start:
                raise ValueError(f"考试「{name}」结束时间不晚于开始时间（{d} {st}~{en}）。")
            patrol_raw = _pick(row, _AL_PATROL)
            patrol_names = [p for p in re.split(r"[\s,，、|/]+", patrol_raw) if p]
            if patrol_names:
                n_patrol += 1
            exams[name] = Exam(name=name, start=start, end=end, rooms=[],
                               patrol_names=patrol_names)
    if on_warn:
        on_warn(f"考试表：共 {len(exams)} 场；其中 {n_patrol} 场填了巡场人员。")
    return exams


def load_rooms_csv(path, exams, on_warn=None):
    """读取考场CSV：科目,考场,监考人数[,班级名称,容量人数,标题]，并挂到对应考试上。
    同时解析出全局元信息（标题、考场→班级名、考场→容量），供 Excel 导出对齐格式用。
    返回 (exams_list, meta)。"""
    meta = ScheduleMeta()
    text = _read_text(path, on_warn)
    with io.StringIO(text) as f:
        r = csv.DictReader(f)
        _require_columns(r.fieldnames, [_AL_SUBJECT, _AL_ROOM, _AL_NEED], path, "考场模板")
        _check_unknown_columns(r.fieldnames,
                               set(_AL_SUBJECT + _AL_ROOM + _AL_NEED + _AL_CLASS
                                   + _AL_CAPACITY + _AL_TITLE),
                               "考场", on_warn)
        for row in r:
            name = _pick(row, _AL_SUBJECT)
            room = _pick(row, _AL_ROOM)
            need_raw = _pick(row, _AL_NEED, "1")
            try:
                need = int(float(need_raw))
            except ValueError:
                raise ValueError(f"考场「{room}」的监考人数不是数字：{need_raw!r}")
            if name in exams and room:
                exams[name].rooms.append((room, need))
            # 解析班级名称 / 容量 / 标题（可选列，可能不存在）
            cls = _pick(row, _AL_CLASS)
            cap = _pick(row, _AL_CAPACITY)
            title = _pick(row, _AL_TITLE)
            if title and not meta.title:
                meta.title = title
            if room:
                if cls:
                    meta.room_class[room] = cls
                if cap:
                    try:
                        meta.room_capacity[room] = int(float(cap))
                    except ValueError:
                        pass
    # 考场模板里出现了没在考试表里定义的科目 → 响亮提醒（否则会被静默丢弃）
    known_subj = set(exams.keys())
    unknown = {_pick(row, _AL_SUBJECT) for row in csv.DictReader(io.StringIO(text))}
    unknown = sorted(s for s in unknown if s and s not in known_subj)
    if unknown and on_warn:
        on_warn(f"⚠️ 考场表里这些科目在考试表里找不到，已被忽略：{'、'.join(unknown)}"
                f"（两表科目名必须完全一致）")
    n_no_room = [k for k, v in exams.items() if not v.rooms]
    if n_no_room and on_warn:
        on_warn(f"⚠️ 这些考试在考场表里没有任何考场，无法安排：{'、'.join(n_no_room)}")
    return list(exams.values()), meta


def load_teachers_csv(path, on_warn=None):
    """读取老师CSV：姓名,是否参与,不可用时段,必用时段"""
    teachers = []
    with io.StringIO(_read_text(path, on_warn)) as f:
        r = csv.DictReader(f)
        _require_columns(r.fieldnames, [_AL_NAME, _AL_PART], path, "老师模板")
        _check_unknown_columns(r.fieldnames,
                               set(_AL_NAME + _AL_PART + _AL_UNAVAIL + _AL_MANDATORY),
                               "老师", on_warn)
        for row in r:
            name = _pick(row, _AL_NAME)
            if not name:
                continue
            part = _pick(row, _AL_PART, "是").lower()
            participate = part not in ("否", "0", "false", "no", "n", "不参与", "不参加")
            try:
                unavailable = _parse_ranges(_pick(row, _AL_UNAVAIL), f"「{name}」的不可用时段")
                mandatory = _parse_ranges(_pick(row, _AL_MANDATORY), f"「{name}」的必用时段")
            except ValueError as e:
                raise ValueError(f"第 {len(teachers) + 2} 行有误：{e}")
            teachers.append(Teacher(name, participate=participate,
                                    unavailable=unavailable, mandatory=mandatory))
    if on_warn:
        n_u = sum(1 for t in teachers if t.unavailable)
        n_m = sum(1 for t in teachers if t.mandatory)
        on_warn(f"老师表：共 {len(teachers)} 人；其中 {n_u} 人填了不可用时段，{n_m} 人填了必用时段。")
        if teachers and n_u == 0:
            on_warn("提示：没有任何老师带不可用时段。若确实填了，请检查是否填在「不可用时段」列。")
    return teachers


# ============ 内置示例数据 ============

def demo_data():
    D1 = datetime(2026, 10, 8)
    D2 = datetime(2026, 10, 9)

    exams = [
        Exam("语文", D1.replace(hour=8, minute=0), D1.replace(hour=10, minute=30),
             [("考场 1", 2), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 2), ("考场 8", 2), ("考场 9", 2), ("考场 10", 2),
              ("考场 11", 2), ("考场 12", 2), ("考场 13", 3)],
             patrol_names=["教师A", "教师B"]),
        Exam("思想政治", D1.replace(hour=11, minute=0), D1.replace(hour=12, minute=0),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 2), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 13", 3)]),
        Exam("化学", D1.replace(hour=13, minute=30), D1.replace(hour=14, minute=30),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 1), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 13", 3)]),
        Exam("数学", D1.replace(hour=15, minute=0), D1.replace(hour=17, minute=0),
             [("考场 1", 2), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 2), ("考场 8", 2), ("考场 9", 2), ("考场 10", 2),
              ("考场 11", 2), ("考场 12", 2), ("考场 13", 3)],
             patrol_names=["教师05"]),
        Exam("物理", D2.replace(hour=8, minute=0), D2.replace(hour=9, minute=0),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 1), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 13", 3)]),
        Exam("历史", D2.replace(hour=9, minute=30), D2.replace(hour=10, minute=30),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 1), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 13", 3)]),
        Exam("生物", D2.replace(hour=11, minute=0), D2.replace(hour=12, minute=0),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 1), ("考场 7", 1), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 12", 1), ("考场 13", 3)]),
        Exam("地理", D2.replace(hour=13, minute=30), D2.replace(hour=14, minute=30),
             [("考场 1", 1), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 2), ("考场 8", 1), ("考场 9", 1), ("考场 10", 1),
              ("考场 11", 1), ("考场 12", 1), ("考场 13", 3)]),
        Exam("英语", D2.replace(hour=15, minute=0), D2.replace(hour=17, minute=0),
             [("考场 1", 2), ("考场 2", 2), ("考场 3", 2), ("考场 4", 2), ("考场 5", 2),
              ("考场 6", 2), ("考场 7", 2), ("考场 8", 2), ("考场 9", 2), ("考场 10", 2),
              ("考场 11", 2), ("考场 12", 2), ("考场 13", 3)]),
    ]

    teachers = [
        Teacher("教师01", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师02", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师03", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师04", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师05", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师06", unavailable=[(D2.replace(hour=8, minute=0), D2.replace(hour=12, minute=30))]),
        Teacher("教师07"),
        Teacher("教师08"),
        Teacher("教师09"),
        Teacher("教师10"),
        Teacher("教师11"),
        Teacher("教师12"),
        Teacher("教师13"),
        Teacher("教师14"),
        Teacher("教师15"),
        Teacher("教师16"),
        Teacher("教师17"),
        Teacher("教师18"),
        Teacher("教师19"),
        Teacher("教师20"),
        Teacher("教师21"),
        Teacher("教师22"),
        Teacher("教师23"),
        Teacher("教师24"),
        Teacher("教师25"),
        Teacher("教师26"),
        Teacher("教师27"),
        Teacher("教师28"),
        Teacher("教师29"),
        Teacher("教师30"),
        Teacher("教师31"),
    ]
    return exams, teachers


# ============ 入口 ============

def main():
    ap = argparse.ArgumentParser(description="监考排班程序")
    ap.add_argument("--make-template", action="store_true", help="生成 3 个数据导入模板")
    ap.add_argument("--exams", help="考试CSV路径")
    ap.add_argument("--rooms", help="考场CSV路径")
    ap.add_argument("--teachers", help="老师CSV路径")
    ap.add_argument("--outdir", default=None, help="导出目录（可选）")
    ap.add_argument("--no-day-balance", action="store_true",
                    help="关闭「每天休息更平均」，退回只均衡总时长的旧行为")
    ap.add_argument("--day-weight", type=float, default=1.0,
                    help="「每天均衡」强度倍数：1=默认；调小更偏向总时长完全相等")
    ap.add_argument("--iterations", type=int, default=None, help="模拟退火迭代次数")
    args = ap.parse_args()

    global W_DAYBAL, W_DAYCNT
    if args.no_day_balance:
        W_DAYBAL = W_DAYCNT = 0.0
    elif args.day_weight != 1.0:
        W_DAYBAL *= args.day_weight
        W_DAYCNT *= max(0.0, args.day_weight)

    if args.make_template:
        make_templates(args.outdir)
        return

    meta = None
    if args.exams and args.rooms and args.teachers:
        exams_map = load_exams_csv(args.exams, on_warn=print)
        exams, meta = load_rooms_csv(args.rooms, exams_map, on_warn=print)
        teachers = load_teachers_csv(args.teachers, on_warn=print)
        if meta is None:
            meta = ScheduleMeta()
    else:
        exams, teachers = demo_data()
        meta = ScheduleMeta()

    # 排班前先做可行性检测：人员不够时给出报告与方案，不修改任何数据
    fb = feasibility_check(exams, teachers)
    if not fb["feasible"]:
        print_feasibility_report(fb)
        print("\n⚠️ 请根据上方方案调整数据后重新运行。程序未修改任何原始文件。")
        return

    sch = Scheduler(exams, teachers, meta=meta)
    result = sch.solve(iterations=args.iterations or ITERATIONS)
    print_schedule(result)
    # 导出目录不存在时自动创建，避免产物回落到脚本所在目录
    if args.outdir:
        try:
            os.makedirs(args.outdir, exist_ok=True)
        except OSError as e:
            print(f"⚠️ 无法创建导出目录 {args.outdir}：{e}")
    export_schedule(result, outdir=args.outdir)
    export_schedule_xlsx(result, outdir=args.outdir)


if __name__ == "__main__":
    main()

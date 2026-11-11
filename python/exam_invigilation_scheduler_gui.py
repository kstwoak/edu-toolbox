#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
监考排班系统 · 图形界面版（Exam Invigilation Scheduler GUI）
============================================================

面向教务老师，无需命令行。核心工作流（点按钮即可完成）：

    ① 生成模板  → 自动生成 3 个 Excel 可打开的 CSV 模板
    ② 填写数据  → 老师用 Excel / WPS 填模板
    ③ 导入数据  → 选择填好的 3 个 CSV
    ④ 一键排班  → 自动计算
    ⑤ 查看/导出 → 界面预览排班表 + 老师时长统计，导出 CSV

排班引擎复用 exam_invigilation_scheduler.py（同一目录）。

运行：
    python3 exam_invigilation_scheduler_gui.py
"""

from __future__ import annotations
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue

# 引入排班引擎（同目录）
import exam_invigilation_scheduler as eis


def _open_path(path: str):
    """用系统默认程序打开文件/文件夹（跨平台：Windows / macOS / Linux）。"""
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')


def _base_dir() -> str:
    """获取程序所在目录。打包成 exe 后（PyInstaller）取 exe 所在目录，否则取脚本目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("监考排班系统")
        self.root.geometry("1080x760")

        # 状态
        self.exams_path = None
        self.rooms_path = None
        self.teachers_path = None
        self.exams = None
        self.teachers = None
        self.meta = None
        self.result = None
        self.outdir = _base_dir()
        # 后台线程 → 主线程的通信队列（tkinter 非线程安全，跨线程操作 UI 必须走队列 + after 轮询）
        self._q = queue.Queue()
        self._job_alive = False

        self._build_ui()
        self._log("欢迎使用监考排班系统。\n"
                  "请按 ①→②→③→④→⑤ 顺序操作（也可以直接「一键排班」使用内置示例数据）。")
        self._poll_queue()

    def _poll_queue(self):
        """主线程轮询后台线程发来的消息，安全地更新 UI。"""
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    done, total, hard, best = payload
                    pct = int(done / total * 100) if total else 0
                    self.progress.configure(value=pct)
                    self._log(f"  进度 {pct:>3}%  ({done}/{total} 迭代)  "
                              f"硬约束违规代价 {int(hard)}  当前最优 {int(best)}")
                elif kind == "done":
                    self._job_alive = False
                    self._on_done(payload)
                elif kind == "error":
                    self._job_alive = False
                    self._on_error(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---------- UI ----------
    def _build_ui(self):
        # 顶部工具栏
        bar = ttk.Frame(self.root, padding=8)
        bar.pack(side="top", fill="x")

        self.btn_template = ttk.Button(bar, text="① 生成模板", command=self._make_template)
        self.btn_template.pack(side="left", padx=4)

        self.btn_import = ttk.Button(bar, text="② 导入数据", command=self._import_data)
        self.btn_import.pack(side="left", padx=4)

        self.btn_run = ttk.Button(bar, text="③ 一键排班", command=self._run_schedule)
        self.btn_run.pack(side="left", padx=4)

        self.btn_export = ttk.Button(bar, text="④ 导出结果", command=self._export_result,
                                     state="disabled")
        self.btn_export.pack(side="left", padx=4)

        self.btn_open = ttk.Button(bar, text="⑤ 用 Excel 打开模板/结果", command=self._open_excel)
        self.btn_open.pack(side="left", padx=4)

        # 「每天休息均衡」强度：关闭 / 适中 / 严格
        ttk.Label(bar, text="每天休息均衡：").pack(side="left", padx=(16, 2))
        self.var_day = tk.StringVar(value="适中（推荐）")
        self.cmb_day = ttk.Combobox(bar, textvariable=self.var_day, width=12,
                                    state="readonly", values=["关闭", "适中（推荐）", "严格"])
        self.cmb_day.pack(side="left")

        self.btn_clear = ttk.Button(bar, text="清空", command=self._clear)
        self.btn_clear.pack(side="right", padx=4)

        # 数据摘要
        self.summary = tk.StringVar(value="尚未导入数据（未导入时会用内置示例数据排班）")
        ttk.Label(self.root, textvariable=self.summary, padding=(8, 2),
                  foreground="#555").pack(side="top", fill="x")

        # 进度
        self.progress = ttk.Progressbar(self.root, mode="determinate", maximum=100)
        self.progress.pack(side="top", fill="x", padx=8)

        # 标签页
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(side="top", fill="both", expand=True, padx=8, pady=4)

        # Tab 1: 排班表
        f1 = ttk.Frame(self.nb)
        self.nb.add(f1, text="排班表")
        self.tree_sched = ttk.Treeview(f1, columns=("exam", "time", "room", "teachers"),
                                       show="headings", height=18)
        for col, txt, w in [("exam", "考试科目", 140), ("time", "时间", 170),
                            ("room", "考场", 130),
                            ("teachers", "老师", 260)]:
            self.tree_sched.heading(col, text=txt)
            self.tree_sched.column(col, width=w, anchor="w")
        vs = ttk.Scrollbar(f1, orient="vertical", command=self.tree_sched.yview)
        self.tree_sched.configure(yscrollcommand=vs.set)
        self.tree_sched.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        # Tab 2: 老师时长统计
        f2 = ttk.Frame(self.nb)
        self.nb.add(f2, text="老师时长统计")
        self.tree_teacher = ttk.Treeview(f2, columns=("name", "room", "mon", "self", "pat", "pmin", "min", "day", "rooms"),
                                         show="headings", height=18)
        for col, txt, w in [("name", "姓名", 90), ("room", "监考场次", 75),
                            ("mon", "监考时长", 75), ("self", "自习场次", 75),
                            ("pat", "巡场场次", 75), ("pmin", "巡场时长", 75),
                            ("min", "总时长(分钟)", 95), ("day", "每天场次", 95),
                            ("rooms", "涉及教室", 150)]:
            self.tree_teacher.heading(col, text=txt)
            self.tree_teacher.column(col, width=w, anchor="w")
        vs2 = ttk.Scrollbar(f2, orient="vertical", command=self.tree_teacher.yview)
        self.tree_teacher.configure(yscrollcommand=vs2.set)
        self.tree_teacher.pack(side="left", fill="both", expand=True)
        vs2.pack(side="right", fill="y")

        # Tab 3: 约束检查
        f3 = ttk.Frame(self.nb)
        self.nb.add(f3, text="约束检查")
        self.txt_check = tk.Text(f3, height=18, wrap="word", padx=8, pady=8)
        self.txt_check.pack(side="left", fill="both", expand=True)
        s3 = ttk.Scrollbar(f3, orient="vertical", command=self.txt_check.yview)
        self.txt_check.configure(yscrollcommand=s3.set)
        s3.pack(side="right", fill="y")

        # 底部日志
        self.txt_log = tk.Text(self.root, height=6, wrap="word", state="disabled",
                               padx=8, pady=4)
        self.txt_log.pack(side="bottom", fill="x")

    def _log(self, msg: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ---------- 动作 ----------
    def _make_template(self):
        d = filedialog.askdirectory(title="选择模板保存位置", initialdir=self.outdir)
        if not d:
            return
        try:
            eis.make_templates(d)
        except OSError as e:
            messagebox.showerror("生成失败", f"无法写入模板：\n{e}")
            return
        self.outdir = d
        self._log(f"✅ 已在 {d} 生成 3 个模板文件。\n"
                  "   请用 Excel / WPS 打开并填写：模板_考试.csv、模板_考场.csv、模板_老师.csv")
        messagebox.showinfo("模板已生成",
                            "已生成 3 个模板文件：\n\n"
                            "  模板_考试.csv   （每场考试一行）\n"
                            "  模板_考场.csv   （每个考场一行）\n"
                            "  模板_老师.csv   （每位老师一行）\n\n"
                            "填好后点「② 导入数据」即可。")

    def _import_data(self):
        p1 = filedialog.askopenfilename(title="① 选择 考试模板 CSV",
                                        initialdir=self.outdir,
                                        filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if not p1:
            return
        p2 = filedialog.askopenfilename(title="② 选择 考场模板 CSV",
                                        initialdir=os.path.dirname(p1),
                                        filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if not p2:
            return
        p3 = filedialog.askopenfilename(title="③ 选择 老师模板 CSV",
                                        initialdir=os.path.dirname(p1),
                                        filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if not p3:
            return
        try:
            exams_map = eis.load_exams_csv(p1, on_warn=self._log)
            exams, meta = eis.load_rooms_csv(p2, exams_map, on_warn=self._log)
            teachers = eis.load_teachers_csv(p3, on_warn=self._log)
        except Exception as e:
            messagebox.showerror("导入失败", f"读取 CSV 出错：\n{e}")
            self._log(f"❌ 导入失败：{e}")
            return

        if not exams:
            messagebox.showerror("导入失败", "考试数据为空。\n"
                                    "请确认「考试模板」里至少填了一行考试，"
                                    "且科目名与「考场模板」里的科目名一致。")
            self._log("❌ 导入失败：考试数据为空。")
            return
        if not teachers:
            messagebox.showerror("导入失败", "老师数据为空。\n"
                                    "请确认「老师模板」里至少填了一行老师。")
            self._log("❌ 导入失败：老师数据为空。")
            return

        self.exams_path, self.rooms_path, self.teachers_path = p1, p2, p3
        self.exams, self.teachers = exams, teachers
        self.meta = meta if meta is not None else eis.ScheduleMeta()
        self.outdir = os.path.dirname(p1)

        n_part = sum(1 for t in teachers if t.participate)
        n_excl = sum(1 for t in teachers if not t.participate)
        n_unav = sum(1 for t in teachers if t.unavailable)
        n_patrol = sum(len(getattr(e, "patrol_names", None) or []) for e in exams)
        self.summary.set(f"已导入：{len(exams)} 场考试、{len(teachers)} 位老师"
                         f"（参与 {n_part} 人 / 不参与 {n_excl} 人）"
                         f"；{n_unav} 人填了不可用时段，{n_patrol} 个巡场名额")
        self._log(f"✅ 数据导入成功：{len(exams)} 场考试，{len(teachers)} 位老师"
                  f"（参与 {n_part} 人 / 不参与 {n_excl} 人）；"
                  f"{n_unav} 人填了不可用时段，共 {n_patrol} 个巡场名额。")
        self._clear_result()

    def _run_schedule(self):
        exams = self.exams
        teachers = self.teachers
        meta = getattr(self, "meta", None)
        if exams is None or teachers is None:
            # 用内置示例
            self._log("未导入数据，使用内置示例数据排班……")
            exams, teachers = eis.demo_data()
            meta = eis.ScheduleMeta()
            self.summary.set("（使用内置示例数据）")

        # 排班前先做人员缺口检测：缺人就给方案，不修改任何数据
        fb = eis.feasibility_check(exams, teachers)
        if not fb["feasible"]:
            self._show_feasibility(fb)
            return

        self._set_busy(True)
        self._log("排班中，请稍候……")
        self._job_alive = True
        day_w = self._day_weight()
        t = threading.Thread(target=self._worker, args=(exams, teachers, meta, day_w), daemon=True)
        t.start()

    def _day_weight(self):
        """把「每天休息均衡」下拉选项映射成 (W_DAYBAL, W_DAYCNT)。"""
        v = self.var_day.get()
        if v.startswith("关闭"):
            return (0.0, 0.0)
        if v.startswith("严格"):
            return (3.0, 3000.0)
        return (0.8, 800.0)

    def _show_feasibility(self, fb):
        """弹窗显示人员缺口报告与方案建议，并同步写入「约束检查」页。"""
        lines = []
        for item in fb["per_exam"]:
            ex = item["exam"]
            mark = "✅" if item["gap"] <= 0 else f"❌ 缺 {item['gap']} 人"
            extra = f"，其中巡场 {len(item['patrol_names'])} 人" if item.get("patrol_names") else ""
            lines.append(f"  {ex.name:<6} {eis._fmt(ex.start)}~{eis._fmt(ex.end)}  "
                         f"需 {item['need']:>2} 人 / 可用 {item['avail']:>2} 人{extra}  {mark}")
        lines.append("")
        lines.append("解决方案建议（请自行选择并修改 CSV 后重新导入）：")
        for s in fb["suggestions"]:
            lines.append(f"\n【{s['exam']}】缺 {s['gap']} 人：")
            for opt in s["options"]:
                lines.append(f"    {opt}")
        lines.append("")
        lines.append("⚠️ 程序未修改任何原始文件，请调整数据后重新「导入数据」再排班。")

        self.txt_check.delete("1.0", "end")
        self.txt_check.insert("1.0", "⚠️  人员缺口检测报告\n\n" + "\n".join(lines))
        self.nb.select(2)
        self._log("⚠️ 人员不足，已在「约束检查」页给出方案建议（未修改任何文件）。")
        messagebox.showwarning("人员不足",
                               "检测到人员缺口，无法排满。\n\n"
                               "已在「约束检查」页列出每场缺口与解决方案建议。\n"
                               "请调整数据后重新导入，再排班。")

    def _worker(self, exams, teachers, meta=None, day_w=(0.8, 800.0)):
        """后台线程：求解，并把日志/进度/结果通过队列发给主线程。"""
        try:
            import traceback

            eis.W_DAYBAL, eis.W_DAYCNT = day_w

            def cb(done, total, hard, best):
                self._q.put(("progress", (done, total, hard, best)))

            sch = eis.Scheduler(exams, teachers, meta=meta)
            self._q.put(("log", f"任务数 {len(sch.tasks)}，老师 {len(sch.all_names)} 人，开始求解……"))
            result = sch.solve(progress_cb=cb)
        except Exception:
            self._q.put(("error", traceback.format_exc()))
            return
        self._q.put(("done", result))

    def _on_error(self, tb_text):
        self._set_busy(False)
        self._log("❌ 排班出错，完整报错如下：\n" + tb_text)
        messagebox.showerror("排班出错",
                             "排班过程中出错，完整报错已打印到日志窗口，\n请截图日志内容以便定位问题。")

    def _on_done(self, result):
        self._set_busy(False)
        self.progress.configure(value=100)
        self.result = result
        self.btn_export.configure(state="normal")
        self._fill_trees(result)
        self._fill_check(result)
        if result["violations"]:
            self._log("⚠️ 排班完成，但存在未满足的硬约束（见「约束检查」页）。")
        else:
            self._log(f"✅ 排班完成，全部硬约束满足。总时长（监考+巡场）极差 {result['balance']} 分钟。")
        self.nb.select(0)

    def _set_busy(self, busy: bool):
        for b in (self.btn_template, self.btn_import, self.btn_run, self.btn_export,
                  self.btn_open, self.btn_clear):
            b.configure(state="disabled" if busy else "normal")
        if not self.result:
            self.btn_export.configure(state="disabled")
        if busy:
            self.progress.configure(mode="determinate", value=0)
        else:
            pass

    def _fill_trees(self, result):
        # 清空
        for t in (self.tree_sched, self.tree_teacher):
            for i in t.get_children():
                t.delete(i)
        # 排班表
        tasks = sorted(result["tasks"], key=lambda x: (x.start, x.room))
        for t in tasks:
            self.tree_sched.insert("", "end", values=(
                t.exam, f"{eis._fmt(t.start)} ~ {eis._fmt(t.end)}", t.room,
                "、".join(result["assign"][t.id])))
        # 老师统计（含巡场时长）
        for r in eis._teacher_rows(result):
            self.tree_teacher.insert("", "end", values=(
                r["name"], r["room_sessions"], r["monitor_min"], r["self_sessions"],
                r["patrol_sessions"], r["patrol_min"], r["total_min"],
                r["day_cnt_str"], r["rooms"]))

    def _fill_check(self, result):
        self.txt_check.delete("1.0", "end")
        lines = []
        if result["violations"]:
            lines.append("⚠️ 存在未满足的硬约束：\n")
            for n, ov, ua, ma, pb in result["violations"]:
                lines.append(f"  {n}: 时间重叠 {ov} 次 / 占用不可用时段 {ua} 次"
                             f" / 必用未覆盖 {ma} 次 / 巡场期间被排监考 {pb} 次")
        else:
            lines.append("✅ 全部硬约束已满足。\n")
        lines.append(f"\n总时长极差（监考+巡场，最忙 - 最闲）: {result['balance']} 分钟"
                     f"（其中监考 {result.get('balance_monitor', 0)} 分 / 巡场 {result.get('balance_patrol', 0)} 分）")
        lines.append(f"自习岗（单人看考场）分布: 每人 {result.get('self_min', 0)}~{result.get('self_max', 0)} 场"
                     f"，极差 {result.get('balance_self', 0)} 场")
        for d in result.get("day_stat") or []:
            lines.append(f"{d['date']} 每天均衡: 时长 {d['min']}~{d['max']} 分钟（平均 {d['avg']:.0f}）"
                         f"，场次 {d['cnt_min']}~{d['cnt_max']} 场")
        ptime = {k: v for k, v in (result.get("ptime") or {}).items() if v}
        if ptime:
            lines.append("\n巡场安排（巡场期间不参与监考，但计入时长统计）：")
            for ename, names in (result.get("patrol_by_exam") or {}).items():
                if names:
                    lines.append(f"  {ename}: {'、'.join(names)}")
        not_in_pool = result.get("patrol_not_in_pool") or []
        if not_in_pool:
            lines.append("\n（以下巡场人不在老师名单中，不参与监考排班，仅计入巡场时长）：")
            lines.append("  " + "、".join(not_in_pool))
        if result["excluded"]:
            lines.append(f"\n全程不参与的老师（未安排任务）：")
            lines.append("  " + "、".join(result["excluded"]))
        self.txt_check.insert("1.0", "\n".join(lines))

    def _export_result(self):
        if not self.result:
            return
        d = filedialog.askdirectory(title="选择导出位置", initialdir=self.outdir)
        if not d:
            return
        import traceback
        try:
            eis.export_schedule(self.result, outdir=d)
            xl = eis.export_schedule_xlsx(self.result, outdir=d)
        except Exception as e:
            self._log("❌ 导出出错，完整报错：\n" + traceback.format_exc())
            messagebox.showerror("导出失败",
                                 f"导出过程中出错：\n{e}\n\n"
                                 "完整报错已打印到日志窗口，请截图日志内容以便定位问题。")
            return
        self._log(f"✅ 已导出到 {d}：排班表.csv、老师时长统计.csv" +
                  (f"、监考安排表.xlsx" if xl else ""))
        msg = (f"已导出：\n  {os.path.join(d, '排班表.csv')}\n  {os.path.join(d, '老师时长统计.csv')}")
        if xl:
            msg += f"\n  {os.path.join(d, '监考安排表.xlsx')}"
        messagebox.showinfo("导出完成", msg)

    def _open_excel(self):
        """用系统默认程序（Excel/WPS）打开最近的模板或结果文件。"""
        targets = []
        for p in (self.exams_path, self.rooms_path, self.teachers_path):
            if p and os.path.exists(p):
                targets.append(p)
        if self.result:
            targets.extend([os.path.join(self.outdir, "排班表.csv"),
                            os.path.join(self.outdir, "老师时长统计.csv"),
                            os.path.join(self.outdir, "监考安排表.xlsx")])
        if not targets:
            messagebox.showinfo("提示", "尚未导入或导出任何文件。\n可先点「① 生成模板」或「② 导入数据」。")
            return
        try:
            for p in targets:
                if os.path.exists(p):
                    _open_path(p)
                    self._log(f"已打开：{p}")
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def _clear_result(self):
        """清空上次的排班结果与表格，但保留已导入的数据。"""
        self.result = None
        self.btn_export.configure(state="disabled")
        for t in (self.tree_sched, self.tree_teacher):
            for i in t.get_children():
                t.delete(i)
        self.txt_check.delete("1.0", "end")

    def _clear(self):
        self.exams = self.teachers = self.meta = self.result = None
        self.exams_path = self.rooms_path = self.teachers_path = None
        self.summary.set("尚未导入数据（未导入时会用内置示例数据排班）")
        self.btn_export.configure(state="disabled")
        for t in (self.tree_sched, self.tree_teacher):
            for i in t.get_children():
                t.delete(i)
        self.txt_check.delete("1.0", "end")
        self._log("已清空当前数据。")


def main():
    root = tk.Tk()
    # macOS 上用 ttk 主题，避免 Aqua 背景色问题
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

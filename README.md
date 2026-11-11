# edu-toolbox 教务工具箱

一套面向学校教务场景的考试辅助工具：**监考排班**、**学生考场座位编排**，以及逐步 Go 化的统一工具箱。全部单机离线运行，输入输出都是 Excel/CSV，不依赖任何服务端。

## 工具一览

| 工具 | 语言 | 状态 | 说明 |
|---|---|---|---|
| 监考排班（引擎 + GUI） | Python 3.8 | ✅ 已交付 | 模拟退火求解，Win7 兼容便携版 |
| 学生考场座位编排 | Python 3.8 | ✅ 已交付 | 与教务手工结果逐格一致 |
| 座位编排引擎（Go） | Go 1.20 | ✅ 引擎完成 | 与 Python 版结果 0 差异 |
| 监考排班引擎（Go） | Go 1.20 | ⏳ 规划中 | 待移植 |
| 图形界面（Go） | walk | ⏳ 规划中 | 目标单文件 exe、双击即跑 |

## 1. 监考排班

`python/exam_invigilation_scheduler.py` + `python/exam_invigilation_scheduler_gui.py`

按考试场次、考场容量和教师不可用时段，自动排监考与巡场：

- **硬约束**：同一时段不重排、不可用时段不排、巡场老师不监考、指定监考必须满足
- **软目标**（按权重折算进代价函数）：
  - 监考时长与巡场时长同权，总时长尽量拉齐
  - **每天分别均衡**（时长 + 场次），避免「一天全休、另一天连打 5 场」
  - 自习岗（单人看考场）在教师间平均分配
- 求解：模拟退火 + 贪心初始解，硬约束与软代价分层比较，不会为软目标牺牲硬约束
- 输入：考试表 + 教师表（Excel/CSV，自动识别 GBK/UTF-8 等编码）
- 输出：排班表、老师时长统计（含每天场次/时长）
- 内置 `--demo` 演示数据（教师均为化名）

详细使用说明见 [docs/监考排班系统使用说明.txt](docs/监考排班系统使用说明.txt)，设计说明见 [docs/](docs/) 下的变更记录。

```bash
# 命令行直接排演示数据
python exam_invigilation_scheduler.py --demo --outdir 输出
```

## 2. 学生考场座位编排

`python/exam_seat_arranger.py`

按总分排名把学生编入考场并生成全局连续座位号（考场 1 = 1~30、考场 2 = 31~60…），每场按总分重排：

- 排序依据总分，同分按名单原始顺序；缺考 / `#N/A` 自动置底
- 名单表颜色标记：黄色 = 语数外三场固定在某考场；主题色高亮 = 排名强制置底
- 非本场学生格自动填「自习NNN」
- 可生成按班级分组的「名单(2)」视图
- 附带 Markdown / HTML / PDF 结果汇总脚本（`make_result_md.py` / `make_result_html.py`）

```bash
# 第一次使用先生成总分填写模板
python exam_seat_arranger.py --template 模板.xlsx --make-score-template

# 正式编排
python exam_seat_arranger.py \
  --template 排考场模板.xlsx \
  --arrangement 考场安排总控表.xlsx \
  --scores 总分表.csv \
  --outdir 考场安排输出 --class-view
```

## 3. Go 版工具箱（进行中）

`cmd/edu-toolbox` + `internal/`，目标是把上述工具合并成**单文件原生程序**（go 1.20，兼顾 Win7），不再捆绑 Python 运行时。

```bash
go build -o edu-toolbox ./cmd/edu-toolbox

./edu-toolbox seat \
  --template 排考场模板.xlsx \
  --arrangement 考场安排总控表.xlsx \
  --scores 总分表.csv \
  --outdir 输出 --class-view
```

### 目录

```
cmd/edu-toolbox/       命令行入口（后续也是 GUI 入口）
internal/xlsxutil/     共享 Excel 层：编码识别、取值、填充色、整行搬移
internal/seat/         座位编排引擎 + 写回
python/                交付中的 Python 版工具
tools/compare_xlsx.py  逐格回归比对脚本
```

### 回归基线

Go 版座位编排与 Python 版 / 教务手工结果比对（2026-11-03）：

- 单元格值：**28823 格，0 差异**
- 合并单元格 / 列宽 / 行高 / 打印设置 / 冻结窗格：无实质差异

### 关键实现备注

- **主题色识别**：「后补学生」用主题色 5（tint 0.6）标记，excelize 解析成 RGB `F8CBAD`，按 RGB 匹配比按 theme 索引更宽容
- **数字单元格在 XML 里没有 `t` 属性**，`GetCellType` 返回 Unset，取值需按「能否解析为数字」判断，否则整行搬移会把数字清空
- **excelize v2.8.1 没有 `SetSheetIndex`**，插入 sheet 顺序需直接调整 workbook 的 sheet 列表
- **行高搬移**只处理偏离多数派的行，避免把默认行高整片写成显式值

## 环境要求

- Python 工具：Python 3.8+，`pip install openpyxl`（GUI 用内置 tkinter）
- Go 工具：Go 1.20+，依赖 [excelize v2](https://github.com/xuri/excelize)（`go mod tidy` 自动安装）

## 隐私说明

仓库内所有代码、示例与文档中的师生姓名均为化名或占位符；真实数据文件不入库。

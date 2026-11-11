// Package seat 学生考场座位编排引擎。
//
// 输入三样：排考场模板 xlsx、考场安排总控表 xlsx、总分表 csv。
// 输出填好的工作簿：名单（按排名重排）、考场 1~13、班级 1~12，可选「名单(2)」。
//
// 分配规则（与教务参考结果逐格比对一致）：
//   - 每场按总分降序重排，按「首座座位号升序」依次填座（全局连续座位号）
//   - 不考该科的学生按同样规则填入本场自习考场，座位号写「自习NNN」
//   - 名单姓名列标黄 → 语数外固定进特殊考场（默认第 12 考场 304）
//   - 选科列标为主题色5 → 新增/后补学生，排名置最后
//   - 某考场人数超出模板预印座位数 → 座位号顺延（末考场扩容）
//   - 总分缺失（缺考 / #N/A / 空）排最后，按名单原始顺序
package seat

import (
	"encoding/csv"
	"fmt"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"

	"edu_toolbox/internal/xlsxutil"
)

// 科目列顺序
var SubjectCols = []string{"语数外", "物理", "化学", "生物", "政治", "历史", "地理"}

var char2Subject = map[rune]string{
	'物': "物理", '化': "化学", '生': "生物",
	'政': "政治", '史': "历史", '地': "地理",
}

// 名单 sheet 列号
const (
	nmSeq = iota + 1
	nmChoice
	nmSID
	nmName
	nmClass
	nmScore
	nmRank
	nmSubj1
)

// nmSubjCol 科目在名单表的列号（语数外=8 … 地理=14）
func nmSubjCol(subj string) int {
	for i, s := range SubjectCols {
		if s == subj {
			return nmSubj1 + i
		}
	}
	return 0
}

// 班级 sheet 列号
const (
	clChoice = iota + 1
	clSID
	clName
	clSubj1
)

func clSubjCol(subj string) int {
	for i, s := range SubjectCols {
		if s == subj {
			return clSubj1 + i
		}
	}
	return 0
}

const (
	nmCols = 14
	clCols = 10

	// YellowRGB 名单姓名列的黄色标记（标黄 → 语数外固定特殊考场）
	YellowRGB = "FFFF00"
	// DefaultAddedRGB 模板选科列的“主题色5”（tint 0.6）解析后的 RGB，
	// 表示新增/后补学生，排名置最后。
	DefaultAddedRGB = "F8CBAD"
)

// Student 一名考生。
type Student struct {
	Choice   string
	SID      string
	Name     string
	Class    string
	Score    float64
	HasScore bool
	ScoreRaw string
	Rank     int
	Special  bool // 姓名列标黄
	Added    bool // 选科列主题色5
	Subjects map[string]bool
	Seats    map[string]int
	SrcRow   int
}

// Takes 该生是否参加某科目考试。
func (s *Student) Takes(subject string) bool {
	switch subject {
	case "语数外", "语文", "数学", "英语":
		return true
	}
	return s.Subjects[subject]
}

// Room 总控表里的考场。
type Room struct {
	Location  string
	Capacity  string
	SeatRange string
}

// Session 一个考试场次。
type Session struct {
	Subject string
	Takers  map[int]int // 考场号 → 考生人数
	Self    map[int]int // 考场号 → 自习人数
}

// Arrangement 总控表内容。
type Arrangement struct {
	Rooms    map[int]Room
	Sessions []Session
	Total    int
	HasTotal bool
}

// Options 编排参数。
type Options struct {
	SpecialRoom int    // 标黄学生语数外固定考场号
	AddedRGB    string // 新增/后补学生的填充色
	Warn        func(string)
}

// Plan 一个场次的落座结果：考场号 → 座位序号 → 学生。
type Plan struct {
	Subject string
	Special bool
	Rooms   map[int]map[int]*Student
}

// ReadStudents 读模板名单表。
func ReadStudents(f *excelize.File, addedRGB string) ([]*Student, error) {
	if idx, err := f.GetSheetIndex("名单"); err != nil || idx < 0 {
		return nil, fmt.Errorf("模板里找不到「名单」sheet")
	}
	rows := xlsxutil.SheetRows(f, "名单")
	var out []*Student
	for r := 2; r <= len(rows); r++ {
		name := strings.TrimSpace(xlsxutil.At(rows, r, nmName))
		if name == "" {
			continue
		}
		s := &Student{
			Choice:   strings.TrimSpace(xlsxutil.At(rows, r, nmChoice)),
			SID:      strings.TrimSpace(xlsxutil.At(rows, r, nmSID)),
			Name:     name,
			Class:    strings.TrimSpace(xlsxutil.At(rows, r, nmClass)),
			Subjects: map[string]bool{},
			Seats:    map[string]int{},
			SrcRow:   r,
		}
		for _, ch := range s.Choice {
			if subj, ok := char2Subject[ch]; ok {
				s.Subjects[subj] = true
			}
		}
		s.Special = xlsxutil.FillRGB(f, "名单", xlsxutil.CellName(r, nmName)) == YellowRGB
		s.Added = xlsxutil.FillRGB(f, "名单", xlsxutil.CellName(r, nmChoice)) == addedRGB
		out = append(out, s)
	}
	return out, nil
}

// ReadRoomSeats 每个考场预印的座位号（按模板 A 列自上而下）。
func ReadRoomSeats(f *excelize.File) map[int][]int {
	rooms := map[int][]int{}
	for _, sn := range f.GetSheetList() {
		if !isAllDigits(sn) {
			continue
		}
		rno, _ := strconv.Atoi(sn)
		rows := xlsxutil.SheetRows(f, sn)
		var labels []int
		for r := 5; r <= len(rows); r++ {
			if v, ok := xlsxutil.ParseNum(xlsxutil.At(rows, r, 1)); ok {
				labels = append(labels, v)
			}
		}
		if len(labels) > 0 {
			rooms[rno] = labels
		}
	}
	return rooms
}

func isAllDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, ch := range s {
		if ch < '0' || ch > '9' {
			return false
		}
	}
	return true
}

// firstNum 取字符串里第一段连续数字（"考场12" → 12）。
func firstNum(s string) (int, bool) {
	start, end := -1, -1
	for i, ch := range s {
		if ch >= '0' && ch <= '9' {
			if start < 0 {
				start = i
			}
			end = i + 1
		} else if start >= 0 {
			break
		}
	}
	if start < 0 {
		return 0, false
	}
	n, err := strconv.Atoi(s[start:end])
	if err != nil {
		return 0, false
	}
	return n, true
}

// ReadArrangement 读总控表（第一个工作表）。
func ReadArrangement(f *excelize.File) (*Arrangement, error) {
	list := f.GetSheetList()
	if len(list) == 0 {
		return nil, fmt.Errorf("总控表没有工作表")
	}
	sheet := list[0]
	rows := xlsxutil.SheetRows(f, sheet)

	headerRow := -1
	limit := len(rows)
	if limit > 20 {
		limit = 20
	}
	for r := 1; r <= limit; r++ {
		if xlsxutil.AtTrim(rows, r, 1) == "科目" {
			headerRow = r
			break
		}
	}
	if headerRow < 0 {
		return nil, fmt.Errorf("总控表里找不到「科目」表头行")
	}

	roomCols := map[int]int{}
	var roomOrder []int
	maxCol := 0
	for _, row := range rows {
		if len(row) > maxCol {
			maxCol = len(row)
		}
	}
	for c := 3; c <= maxCol; c++ {
		v := xlsxutil.AtTrim(rows, headerRow, c)
		if strings.HasPrefix(v, "考场") {
			if rno, ok := firstNum(v); ok {
				if _, dup := roomCols[rno]; !dup {
					roomOrder = append(roomOrder, rno)
				}
				roomCols[rno] = c
			}
		}
	}
	sort.Ints(roomOrder)

	arr := &Arrangement{Rooms: map[int]Room{}}
	for rno, c := range roomCols {
		arr.Rooms[rno] = Room{
			Location:  xlsxutil.AtTrim(rows, headerRow+1, c),
			Capacity:  xlsxutil.AtTrim(rows, headerRow+2, c),
			SeatRange: xlsxutil.AtTrim(rows, headerRow+3, c),
		}
	}

	counts := func(r int) map[int]int {
		out := map[int]int{}
		for rno, c := range roomCols {
			if n, ok := xlsxutil.ParseNum(xlsxutil.At(rows, r, c)); ok && n > 0 {
				out[rno] = n
			}
		}
		return out
	}

	for r := headerRow + 1; r <= len(rows); r++ {
		a := xlsxutil.AtTrim(rows, r, 1)
		if a == "" || a == "自习" {
			continue
		}
		if strings.HasPrefix(a, "共") {
			if n, ok := firstNum(a); ok {
				arr.Total, arr.HasTotal = n, true
			}
			continue
		}
		takers := counts(r)
		// 紧跟其后的「自习」行给出本场自习分布（中间可隔空行）
		nxt := r + 1
		for nxt <= len(rows) && xlsxutil.AtTrim(rows, nxt, 1) == "" {
			nxt++
		}
		self := map[int]int{}
		if nxt <= len(rows) && xlsxutil.AtTrim(rows, nxt, 1) == "自习" {
			self = counts(nxt)
		}
		arr.Sessions = append(arr.Sessions, Session{Subject: a, Takers: takers, Self: self})
	}
	return arr, nil
}

var (
	aliasSID   = map[string]bool{"考号": true, "准考证号": true, "学号": true, "id": true, "student_id": true}
	aliasScore = map[string]bool{"总分": true, "分数": true, "成绩": true, "score": true, "total_score": true}
)

// ScoreEntry 总分表里的一条记录。
type ScoreEntry struct {
	Score    float64
	HasScore bool
	Raw      string
}

// ReadScores 读总分表 CSV（自动识别编码）。
func ReadScores(path string) (map[string]ScoreEntry, error) {
	text, err := xlsxutil.ReadText(path)
	if err != nil {
		return nil, err
	}
	r := csv.NewReader(strings.NewReader(text))
	r.FieldsPerRecord = -1
	recs, err := r.ReadAll()
	if err != nil {
		return nil, err
	}
	out := map[string]ScoreEntry{}
	if len(recs) == 0 {
		return out, nil
	}
	header := recs[0]
	sidCol, scoreCol := -1, -1
	for i, h := range header {
		key := strings.TrimSpace(h)
		if aliasSID[key] && sidCol < 0 {
			sidCol = i
		} else if aliasScore[key] && scoreCol < 0 {
			scoreCol = i
		}
	}
	if sidCol < 0 {
		return out, nil
	}
	for _, rec := range recs[1:] {
		sid := ""
		if sidCol < len(rec) {
			sid = strings.TrimSpace(rec[sidCol])
		}
		if sid == "" {
			continue
		}
		raw := ""
		if scoreCol >= 0 && scoreCol < len(rec) {
			raw = strings.TrimSpace(rec[scoreCol])
		}
		e := ScoreEntry{Raw: raw}
		if fl, err := strconv.ParseFloat(raw, 64); err == nil {
			e.Score, e.HasScore = fl, true
		}
		out[sid] = e
	}
	return out, nil
}

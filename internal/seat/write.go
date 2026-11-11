package seat

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"

	"edu_toolbox/internal/xlsxutil"
)

// examCol 考场 sheet 里每科占两列（考号、姓名），从第 2 列起。
func examCol(subj string) (int, int) {
	for i, s := range SubjectCols {
		if s == subj {
			c := 2 + i*2
			return c, c + 1
		}
	}
	return 0, 0
}

// seatText 名单/班级表里的座位号文本：考生写数字，非本场考生写「自习NNN」。
func seatText(s *Student, subj string, fillSelf bool) (interface{}, bool) {
	var v int
	var ok bool
	if subj == "语数外" {
		v, ok = s.Seats["语文"] // 语数英同座位
	} else {
		v, ok = s.Seats[subj]
	}
	if !ok {
		return nil, false
	}
	if subj == "语数外" || s.Takes(subj) {
		return v, true
	}
	if !fillSelf {
		return nil, false
	}
	return fmt.Sprintf("自习%d", v), true
}

// WriteOutput 把编排结果写回模板副本。
func WriteOutput(templatePath, outPath string, res *Arranged, fillSelf bool) error {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return err
	}
	defer f.Close()

	bySID := map[string]*Student{}
	for _, s := range res.Ranked {
		bySID[s.SID] = s
	}

	// ── 名单 ────────────────────────────────
	const sheet = "名单"
	rows := xlsxutil.SheetRows(f, sheet)
	first, last := 2, len(rows)
	for r := first; r <= last; r++ {
		sid := strings.TrimSpace(xlsxutil.At(rows, r, nmSID))
		s := bySID[sid]
		if s == nil {
			continue
		}
		var scoreVal interface{}
		switch {
		case s.HasScore:
			scoreVal = s.Score
		case s.ScoreRaw != "":
			scoreVal = s.ScoreRaw
		default:
			scoreVal = "缺考"
		}
		xlsxutil.SetCell(f, sheet, r, nmScore, scoreVal)
		xlsxutil.SetCell(f, sheet, r, nmRank, s.Rank)
		for _, subj := range SubjectCols {
			if txt, ok := seatText(s, subj, fillSelf); ok {
				xlsxutil.SetCell(f, sheet, r, nmSubjCol(subj), txt)
			}
		}
	}
	// 按排名重排 + 序号重写
	xlsxutil.SortRowsByCol(f, sheet, nmRank, first, last, nmCols)
	for i, r := 0, first; r <= last; i, r = i+1, r+1 {
		xlsxutil.SetCell(f, sheet, r, nmSeq, i+1)
	}

	// ── 考场 ────────────────────────────────
	for _, plan := range res.Plans {
		key := plan.Subject
		if plan.Special {
			key = "语数外"
		}
		cCol, nCol := examCol(key)
		for rno, seats := range plan.Rooms {
			sheetName := strconv.Itoa(rno)
			srows := xlsxutil.SheetRows(f, sheetName)
			var pre []int
			for i := 5; i <= len(srows); i++ {
				if strings.TrimSpace(xlsxutil.At(srows, i, 1)) != "" {
					pre = append(pre, i)
				}
			}
			if len(pre) == 0 {
				continue
			}
			for k, s := range seats {
				var row int
				if k < len(pre) {
					row = pre[k]
				} else {
					row = pre[len(pre)-1] + (k - len(pre) + 1)
					if labels := res.RoomSeat[rno]; k < len(labels) {
						xlsxutil.SetCell(f, sheetName, row, 1, labels[k])
					}
				}
				xlsxutil.SetCell(f, sheetName, row, cCol, s.SID)
				xlsxutil.SetCell(f, sheetName, row, nCol, s.Name)
			}
		}
	}

	// ── 班级 ────────────────────────────────
	clsSheets := map[int]string{}
	reCls := regexp.MustCompile(`^(\d+)班$`)
	for _, sn := range f.GetSheetList() {
		if m := reCls.FindStringSubmatch(sn); m != nil {
			if n, err := strconv.Atoi(m[1]); err == nil {
				clsSheets[n] = sn
			}
		}
	}
	reClsNum := regexp.MustCompile(`(\d+)班`)
	byCls := map[int][]*Student{}
	var clsOrder []int
	for _, s := range res.Ranked {
		num, ok := -1, false
		if m := reClsNum.FindStringSubmatch(s.Class); m != nil {
			if n, err := strconv.Atoi(m[1]); err == nil {
				num, ok = n, true
			}
		}
		if !ok {
			continue
		}
		if _, seen := byCls[num]; !seen {
			clsOrder = append(clsOrder, num)
		}
		byCls[num] = append(byCls[num], s)
	}
	for num, sn := range clsSheets {
		crows := xlsxutil.SheetRows(f, sn)
		for r := 3; r <= len(crows); r++ {
			for c := 1; c <= clCols; c++ {
				xlsxutil.SetCell(f, sn, r, c, nil)
			}
		}
		for i, s := range byCls[num] {
			r := 3 + i
			xlsxutil.SetCell(f, sn, r, clChoice, s.Choice)
			xlsxutil.SetCell(f, sn, r, clSID, s.SID)
			xlsxutil.SetCell(f, sn, r, clName, s.Name)
			for _, subj := range SubjectCols {
				if txt, ok := seatText(s, subj, fillSelf); ok {
					xlsxutil.SetCell(f, sn, r, clSubjCol(subj), txt)
				}
			}
		}
	}

	if err := f.SaveAs(outPath); err != nil {
		return err
	}
	return nil
}

// AddClassView 追加「名单(2)」：按班级分组（组内按排名）的座位总表。
func AddClassView(path string, ranked []*Student) error {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return err
	}
	defer f.Close()

	const title = "名单(2)"
	for _, sn := range []string{"名单(2)", "名单 (2)"} {
		if idx, _ := f.GetSheetIndex(sn); idx >= 0 {
			_ = f.DeleteSheet(sn)
		}
	}
	if _, err := f.NewSheet(title); err != nil {
		return err
	}
	if idx, _ := f.GetSheetIndex("名单"); idx >= 0 {
		moveSheetTo(f, title, idx+1)
	}

	thin := []excelize.Border{
		{Type: "left", Style: 1, Color: "000000"},
		{Type: "right", Style: 1, Color: "000000"},
		{Type: "top", Style: 1, Color: "000000"},
		{Type: "bottom", Style: 1, Color: "000000"},
	}
	center := &excelize.Alignment{Horizontal: "center", Vertical: "center"}
	hdrStyle, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}, Alignment: center, Border: thin})
	if err != nil {
		return err
	}
	bodyStyle, err := f.NewStyle(&excelize.Style{Alignment: center, Border: thin})
	if err != nil {
		return err
	}

	headers := append([]string{"序号", "选科", "考号", "学生姓名"}, SubjectCols...)
	headers = append(headers, "班级")
	for c, h := range headers {
		cell := xlsxutil.CellName(1, c+1)
		_ = f.SetCellValue(title, cell, h)
		_ = f.SetCellStyle(title, cell, cell, hdrStyle)
	}

	ordered := make([]*Student, len(ranked))
	copy(ordered, ranked)
	sort.SliceStable(ordered, func(i, j int) bool {
		if ordered[i].Class != ordered[j].Class {
			return ordered[i].Class < ordered[j].Class
		}
		return ordered[i].Rank < ordered[j].Rank
	})

	for i, s := range ordered {
		r := i + 2
		xlsxutil.SetCell(f, title, r, 1, s.Rank)
		xlsxutil.SetCell(f, title, r, 2, s.Choice)
		xlsxutil.SetCell(f, title, r, 3, s.SID)
		xlsxutil.SetCell(f, title, r, 4, s.Name)
		for j, subj := range SubjectCols {
			if txt, ok := seatText(s, subj, true); ok {
				xlsxutil.SetCell(f, title, r, 5+j, txt)
			}
		}
		xlsxutil.SetCell(f, title, r, 5+len(SubjectCols), s.Class)
		for c := 1; c <= len(headers); c++ {
			cell := xlsxutil.CellName(r, c)
			_ = f.SetCellStyle(title, cell, cell, bodyStyle)
		}
	}

	widths := []float64{8, 10, 14, 12, 10, 10, 10, 10, 10, 10, 10, 18}
	for c, w := range widths {
		col, _ := excelize.ColumnNumberToName(c + 1)
		_ = f.SetColWidth(title, col, col, w)
	}
	_ = f.SetPanes(title, &excelize.Panes{Freeze: true, YSplit: 1, TopLeftCell: "A2", ActivePane: "bottomLeft"})

	return f.Save()
}

// moveSheetTo 把工作表移到指定位置（excelize v2.8.1 没有公开的 SetSheetIndex）。
// 直接操作 workbook 的 sheet 列表——该字段导出，但类型未导出，故只能移动已有元素。
func moveSheetTo(f *excelize.File, name string, pos int) {
	sheets := f.WorkBook.Sheets.Sheet
	idx := -1
	for i, s := range sheets {
		if s.Name == name {
			idx = i
			break
		}
	}
	if idx < 0 || pos < 0 || pos >= len(sheets) || idx == pos {
		return
	}
	sh := sheets[idx]
	sheets = append(sheets[:idx], sheets[idx+1:]...)
	sheets = append(sheets, sh) // 末位占位（值随后被覆盖）
	copy(sheets[pos+1:], sheets[pos:])
	sheets[pos] = sh
	f.WorkBook.Sheets.Sheet = sheets
}

// Verify 自检：各场考生/自习人数是否与总控表吻合、有无重复占座。
func Verify(res *Arranged, arr *Arrangement) (string, bool) {
	var b strings.Builder
	b.WriteString(fmt.Sprintf("【自检】\n  学生总数：%d\n", len(res.Ranked)))
	exp := map[string][2]int{}
	for _, s := range arr.Sessions {
		var t, sf int
		for _, n := range s.Takers {
			t += n
		}
		for _, n := range s.Self {
			sf += n
		}
		exp[s.Subject] = [2]int{t, sf}
	}
	ok := true
	for _, plan := range res.Plans {
		subj := plan.Subject
		if subj == "数学" || subj == "英语" {
			continue
		}
		nExam, nSelf := 0, 0
		for _, d := range plan.Rooms {
			for _, s := range d {
				if s.Takes(subj) {
					nExam++
				} else {
					nSelf++
				}
			}
		}
		key := subj
		if plan.Special {
			key = "语数外"
		}
		flag := "  ✓"
		if e, found := exp[key]; found && [2]int{nExam, nSelf} != e {
			flag = fmt.Sprintf("  ✗ 应为 考生%d/自习%d", e[0], e[1])
			ok = false
		}
		b.WriteString(fmt.Sprintf("  %-4s 考生 %3d / 自习 %3d%s\n", key, nExam, nSelf, flag))
	}
	for _, plan := range res.Plans {
		seen := map[string]bool{}
		for _, d := range plan.Rooms {
			for _, s := range d {
				if seen[s.SID] {
					b.WriteString(fmt.Sprintf("  ✗ 场次%s 重复占座 %s\n", plan.Subject, s.Name))
					ok = false
				}
				seen[s.SID] = true
			}
		}
	}
	if ok {
		b.WriteString("  ✓ 全部通过\n")
	} else {
		b.WriteString("  ✗ 有问题\n")
	}
	return b.String(), ok
}

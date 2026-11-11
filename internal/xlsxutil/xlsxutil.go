// Package xlsxutil 提供 Excel 读写的共享底层工具：
// 编码自动识别、单元格取值、填充色识别、行列快照与搬移。
//
// 这些能力原本由 Python 的 openpyxl 提供，Go 侧用 excelize 实现，
// 并刻意保持与引擎相同的语义（数字 0 基/1 基、空值处理等）。
package xlsxutil

import (
	"bytes"
	"os"
	"strconv"
	"strings"
	"unicode/utf16"

	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/simplifiedchinese"
)

// ReadText 读取文本文件并自动识别编码（BOM / UTF-8 / GB18030）。
// 教务系统导出的 CSV 常常是 GBK，直接按 UTF-8 读会报错。
func ReadText(path string) (string, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	switch {
	case bytes.HasPrefix(raw, []byte{0xEF, 0xBB, 0xBF}):
		return string(raw[3:]), nil
	case bytes.HasPrefix(raw, []byte{0xFF, 0xFE, 0x00, 0x00}):
		return decodeUTF32(raw[4:], true), nil
	case bytes.HasPrefix(raw, []byte{0x00, 0x00, 0xFE, 0xFF}):
		return decodeUTF32(raw[4:], false), nil
	case bytes.HasPrefix(raw, []byte{0xFF, 0xFE}):
		return decodeUTF16(raw[2:], true), nil
	case bytes.HasPrefix(raw, []byte{0xFE, 0xFF}):
		return decodeUTF16(raw[2:], false), nil
	}
	if ok := isValidUTF8(raw); ok {
		return string(raw), nil
	}
	if out, err := simplifiedchinese.GB18030.NewDecoder().Bytes(raw); err == nil {
		return string(out), nil
	}
	return strings.ToValidUTF8(string(raw), ""), nil
}

func isValidUTF8(b []byte) bool {
	for i := 0; i < len(b); {
		c := b[i]
		switch {
		case c < 0x80:
			i++
		case c&0xE0 == 0xC0:
			if i+1 >= len(b) || b[i+1]&0xC0 != 0x80 {
				return false
			}
			i += 2
		case c&0xF0 == 0xE0:
			if i+2 >= len(b) || b[i+1]&0xC0 != 0x80 || b[i+2]&0xC0 != 0x80 {
				return false
			}
			i += 3
		case c&0xF8 == 0xF0:
			if i+3 >= len(b) || b[i+1]&0xC0 != 0x80 || b[i+2]&0xC0 != 0x80 || b[i+3]&0xC0 != 0x80 {
				return false
			}
			i += 4
		default:
			return false
		}
	}
	return true
}

func decodeUTF16(b []byte, little bool) string {
	n := len(b) / 2
	u := make([]uint16, n)
	for i := 0; i < n; i++ {
		if little {
			u[i] = uint16(b[2*i]) | uint16(b[2*i+1])<<8
		} else {
			u[i] = uint16(b[2*i])<<8 | uint16(b[2*i+1])
		}
	}
	return string(utf16.Decode(u))
}

func decodeUTF32(b []byte, little bool) string {
	var sb strings.Builder
	for i := 0; i+3 < len(b); i += 4 {
		var r uint32
		if little {
			r = uint32(b[i]) | uint32(b[i+1])<<8 | uint32(b[i+2])<<16 | uint32(b[i+3])<<24
		} else {
			r = uint32(b[i])<<24 | uint32(b[i+1])<<16 | uint32(b[i+2])<<8 | uint32(b[i+3])
		}
		sb.WriteRune(rune(r))
	}
	return sb.String()
}

// SheetRows 返回整表文本（1 基行/列，越界返回空串）。等价 openpyxl 的 ws.max_row 遍历。
func SheetRows(f *excelize.File, sheet string) [][]string {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil
	}
	return rows
}

// At 取 rows 中第 r 行第 c 列（均 1 基），越界返回 ""。
func At(rows [][]string, r, c int) string {
	if r < 1 || c < 1 || r > len(rows) {
		return ""
	}
	row := rows[r-1]
	if c > len(row) {
		return ""
	}
	return row[c-1]
}

// AtTrim 同 At，但去掉首尾空白。
func AtTrim(rows [][]string, r, c int) string {
	return strings.TrimSpace(At(rows, r, c))
}

// CellName 拼出 "B12" 形式的坐标。
func CellName(row, col int) string {
	name, _ := excelize.CoordinatesToCellName(col, row)
	return name
}

// SetCell 按 1 基行列写值。
func SetCell(f *excelize.File, sheet string, row, col int, v interface{}) {
	_ = f.SetCellValue(sheet, CellName(row, col), v)
}

// ParseNum 把单元格文本转成整数（兼容 "47"、"47.0"、"47人"）。
func ParseNum(s string) (int, bool) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, false
	}
	fl, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0, false
	}
	return int(fl), true
}

// FillRGB 返回单元格填充色（大写 RRGGBB），无实心填充返回 ""。
// excelize 会把主题色自动解析成 RGB，所以主题色也能识别。
func FillRGB(f *excelize.File, sheet, cell string) string {
	id, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return ""
	}
	st, err := f.GetStyle(id)
	if err != nil || st == nil {
		return ""
	}
	if st.Fill.Pattern == 0 || len(st.Fill.Color) == 0 {
		return ""
	}
	c := strings.ToUpper(strings.TrimSpace(st.Fill.Color[0]))
	if len(c) == 8 {
		c = c[2:]
	}
	return c
}

// cellSnap 保存单元格的“值 + 样式”，用于整行搬移时保真。
type cellSnap struct {
	val   interface{}
	style int
}

// SnapRow 快照一行（第 1..ncols 列）。
func snapRow(f *excelize.File, sheet string, row, ncols int) []cellSnap {
	out := make([]cellSnap, ncols)
	for c := 1; c <= ncols; c++ {
		cell := CellName(row, c)
		id, _ := f.GetCellStyle(sheet, cell)
		ct, _ := f.GetCellType(sheet, cell)
		raw, _ := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
		var v interface{}
		switch ct {
		case excelize.CellTypeBool:
			v = strings.EqualFold(strings.TrimSpace(raw), "true")
		case excelize.CellTypeSharedString, excelize.CellTypeInlineString,
			excelize.CellTypeError:
			if raw == "" {
				v = nil
			} else {
				v = raw
			}
		default:
			// 数字单元格在 XML 里没有 t 属性，GetCellType 返回 Unset，
			// 因此这里以“能否解析成数字”为准。
			t := strings.TrimSpace(raw)
			if t == "" {
				v = nil
			} else if fl, err := strconv.ParseFloat(t, 64); err == nil {
				v = fl
			} else {
				v = raw
			}
		}
		out[c-1] = cellSnap{val: v, style: id}
	}
	return out
}

// rowSortItem 一行在排序中的快照。
type rowSortItem struct {
	key    float64
	hasKey bool
	cells  []cellSnap
	height float64
	hasH   bool
}

// SortRowsByCol 按某列数值升序重排行区间，保留每格样式与自定义行高。
// 语义与 Python 版 _sort_rows_by_col 一致：空值排在最后，同值保持原顺序。
func SortRowsByCol(f *excelize.File, sheet string, col, first, last, ncols int) {
	if last < first {
		return
	}
	defH := 0.0
	if props, err := f.GetSheetProps(sheet); err == nil && props.DefaultRowHeight != nil {
		defH = *props.DefaultRowHeight
	}
	// 先取全区间行高，找出“多数派”高度：只有偏离多数派的行才算自定义行高，
	// 避免把默认高度当成自定义值整片写回去。
	heights := make([]float64, last-first+1)
	freq := map[float64]int{}
	for i, r := 0, first; r <= last; i, r = i+1, r+1 {
		h, err := f.GetRowHeight(sheet, r)
		if err != nil {
			h = defH
		}
		heights[i] = h
		freq[h]++
	}
	mode, modeN := defH, 0
	for h, n := range freq {
		if n > modeN || (n == modeN && h == defH) {
			mode, modeN = h, n
		}
	}

	items := make([]rowSortItem, 0, last-first+1)
	for i, r := 0, first; r <= last; i, r = i+1, r+1 {
		ct, _ := f.GetCellType(sheet, CellName(r, col))
		raw, _ := f.GetCellValue(sheet, CellName(r, col), excelize.Options{RawCellValue: true})
		it := rowSortItem{cells: snapRow(f, sheet, r, ncols)}
		if ct == excelize.CellTypeNumber || ct == excelize.CellTypeDate {
			if fl, err := strconv.ParseFloat(strings.TrimSpace(raw), 64); err == nil {
				it.key, it.hasKey = fl, true
			}
		} else if n, ok := ParseNum(raw); ok {
			it.key, it.hasKey = float64(n), true
		}
		if freq[mode] < len(heights) && heights[i] != mode {
			it.height, it.hasH = heights[i], true
		}
		items = append(items, it)
	}
	// 稳定排序：有键在前按升序，无键在后
	sortStable(items)
	for i, it := range items {
		r := first + i
		for c := 1; c <= ncols; c++ {
			cell := CellName(r, c)
			_ = f.SetCellValue(sheet, cell, it.cells[c-1].val)
			_ = f.SetCellStyle(sheet, cell, cell, it.cells[c-1].style)
		}
		if it.hasH {
			_ = f.SetRowHeight(sheet, r, it.height)
		}
	}
}

func sortStable(items []rowSortItem) {
	// 插入排序足够：行数在千级，且保证稳定性、实现简单
	for i := 1; i < len(items); i++ {
		cur := items[i]
		j := i - 1
		for j >= 0 && less(cur, items[j]) {
			items[j+1] = items[j]
			j--
		}
		items[j+1] = cur
	}
}

func less(a, b rowSortItem) bool {
	if a.hasKey != b.hasKey {
		return a.hasKey
	}
	if !a.hasKey {
		return false
	}
	return a.key < b.key
}

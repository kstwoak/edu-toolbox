// edu-toolbox 教务工具箱（命令行版）
//
// 用法：
//
//	edu-toolbox seat --template 模板.xlsx --arrangement 考场安排.xlsx --scores 总分表.csv
//	edu-toolbox seat --template 模板.xlsx --make-score-template
//
// 目前包含「学生考场座位安排」；「监考排班」与「登记表分页拆分」将陆续并入。
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"

	"edu_toolbox/internal/seat"
	"edu_toolbox/internal/xlsxutil"
)

func main() {
	if len(os.Args) < 2 {
		usage()
		return
	}
	switch os.Args[1] {
	case "seat":
		runSeat(os.Args[2:])
	case "-h", "--help", "help":
		usage()
	default:
		fmt.Fprintf(os.Stderr, "未知命令：%s\n", os.Args[1])
		usage()
	}
}

func usage() {
	fmt.Print(`教务工具箱

  edu-toolbox seat      学生考场座位安排
    必填：--template 模板.xlsx  --arrangement 总控表.xlsx  --scores 总分表.csv
    可选：--outdir 输出目录（默认 考场安排输出）
          --special-room N     标黄学生固定考场号（默认 12）
          --no-self-seat       非本场考生留空（默认写「自习NNN」）
          --class-view         额外生成「名单(2)」按班级分组视图
          --added-color RRGGBB 新增/后补学生的标记色（默认 F8CBAD）
          --make-score-template 按模板名单生成 总分表_待填.csv
`)
}

func runSeat(args []string) {
	fs := flag.NewFlagSet("seat", flag.ExitOnError)
	template := fs.String("template", "", "排考场模板 xlsx")
	arrangement := fs.String("arrangement", "", "考场安排总控表 xlsx")
	scores := fs.String("scores", "", "总分表 csv")
	outdir := fs.String("outdir", "考场安排输出", "输出目录")
	specialRoom := fs.Int("special-room", 12, "标黄学生语数外固定考场号")
	noSelfSeat := fs.Bool("no-self-seat", false, "非本场考生留空")
	classView := fs.Bool("class-view", false, "生成名单(2)班级视图")
	makeTpl := fs.Bool("make-score-template", false, "生成总分表填写模板")
	addedColor := fs.String("added-color", seat.DefaultAddedRGB, "新增/后补学生标记色")
	_ = fs.Parse(args)

	if *outdir != "" {
		if err := os.MkdirAll(*outdir, 0o755); err != nil {
			fmt.Fprintf(os.Stderr, "创建输出目录失败：%v\n", err)
			os.Exit(1)
		}
	}

	if *makeTpl {
		if *template == "" {
			fmt.Fprintln(os.Stderr, "--make-score-template 需要同时给 --template")
			os.Exit(2)
		}
		p, err := makeScoreTemplate(*template, *outdir)
		if err != nil {
			fmt.Fprintf(os.Stderr, "生成失败：%v\n", err)
			os.Exit(1)
		}
		fmt.Printf("已生成：%s\n", p)
		return
	}
	if *template == "" || *arrangement == "" || *scores == "" {
		fs.Usage()
		return
	}

	fmt.Println("读取模板名单……")
	f, err := excelize.OpenFile(*template)
	if err != nil {
		fmt.Fprintf(os.Stderr, "打开模板失败：%v\n", err)
		os.Exit(1)
	}
	students, err := seat.ReadStudents(f, strings.ToUpper(*addedColor))
	if err != nil {
		fmt.Fprintf(os.Stderr, "%v\n", err)
		os.Exit(1)
	}
	nSpecial := 0
	for _, s := range students {
		if s.Special {
			nSpecial++
		}
	}
	fmt.Printf("  %d 位学生，其中标黄 %d 人\n", len(students), nSpecial)

	roomSeats := seat.ReadRoomSeats(f)
	_ = f.Close()
	fmt.Println("读取考场座位号：")
	rnos := make([]int, 0, len(roomSeats))
	for rno := range roomSeats {
		rnos = append(rnos, rno)
	}
	sortInts(rnos)
	for _, rno := range rnos {
		l := roomSeats[rno]
		fmt.Printf("  考场%-2d %d 座（%d~%d）\n", rno, len(l), l[0], l[len(l)-1])
	}

	af, err := excelize.OpenFile(*arrangement)
	if err != nil {
		fmt.Fprintf(os.Stderr, "打开总控表失败：%v\n", err)
		os.Exit(1)
	}
	arr, err := seat.ReadArrangement(af)
	_ = af.Close()
	if err != nil {
		fmt.Fprintf(os.Stderr, "%v\n", err)
		os.Exit(1)
	}
	fmt.Printf("读取总控表：%d 个考场，%d 个场次\n", len(arr.Rooms), len(arr.Sessions))

	scoreMap, err := seat.ReadScores(*scores)
	if err != nil {
		fmt.Fprintf(os.Stderr, "读取总分表失败：%v\n", err)
		os.Exit(1)
	}
	fmt.Printf("读取总分表：%d 条\n", len(scoreMap))

	fmt.Println("编排中……")
	res, err := seat.Arrange(students, roomSeats, arr, scoreMap, seat.Options{
		SpecialRoom: *specialRoom,
		AddedRGB:    strings.ToUpper(*addedColor),
		Warn:        func(s string) { fmt.Println(s) },
	})
	if err != nil {
		fmt.Fprintf(os.Stderr, "%v\n", err)
		os.Exit(1)
	}

	outPath := filepath.Join(*outdir, "24级排考场（已填）.xlsx")
	if err := seat.WriteOutput(*template, outPath, res, !*noSelfSeat); err != nil {
		fmt.Fprintf(os.Stderr, "写出失败：%v\n", err)
		os.Exit(1)
	}
	if *classView {
		if err := seat.AddClassView(outPath, res.Ranked); err != nil {
			fmt.Fprintf(os.Stderr, "生成名单(2) 失败：%v\n", err)
			os.Exit(1)
		}
	}

	report, ok := seat.Verify(res, arr)
	fmt.Println(report)
	fmt.Printf("\n结果文件：%s\n", outPath)
	if !ok {
		os.Exit(3)
	}
}

// makeScoreTemplate 按模板名单生成待填的总分表。
func makeScoreTemplate(templatePath, outdir string) (string, error) {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return "", err
	}
	defer f.Close()
	rows := xlsxutil.SheetRows(f, "名单")
	var b strings.Builder
	b.WriteString("\uFEFF考号,姓名,总分\r\n")
	for r := 2; r <= len(rows); r++ {
		name := strings.TrimSpace(xlsxutil.At(rows, r, 4))
		if name == "" {
			continue
		}
		sid := strings.TrimSpace(xlsxutil.At(rows, r, 3))
		b.WriteString(csvField(sid) + "," + csvField(name) + ",\r\n")
	}
	p := filepath.Join(outdir, "总分表_待填.csv")
	if err := os.WriteFile(p, []byte(b.String()), 0o644); err != nil {
		return "", err
	}
	return p, nil
}

func csvField(s string) string {
	if strings.ContainsAny(s, ",\"\r\n") {
		return `"` + strings.ReplaceAll(s, `"`, `""`) + `"`
	}
	return s
}

func sortInts(a []int) {
	for i := 1; i < len(a); i++ {
		cur := a[i]
		j := i - 1
		for j >= 0 && a[j] > cur {
			a[j+1] = a[j]
			j--
		}
		a[j+1] = cur
	}
}

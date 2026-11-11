package seat

import (
	"fmt"
	"sort"
	"strings"
)

// Arranged 编排结果。
type Arranged struct {
	Ranked   []*Student
	Plans    []Plan
	RoomSeat map[int][]int // 可能被末考场扩容改写
}

// Arrange 执行编排。roomSeats 会被就地修改（末考场座位号顺延）。
func Arrange(students []*Student, roomSeats map[int][]int, arr *Arrangement, scores map[string]ScoreEntry, opt Options) (*Arranged, error) {
	warn := opt.Warn
	if warn == nil {
		warn = func(string) {}
	}
	if opt.SpecialRoom == 0 {
		opt.SpecialRoom = 12
	}

	missing := 0
	var missingNames []string
	for _, s := range students {
		e := scores[s.SID]
		s.Score, s.HasScore, s.ScoreRaw = e.Score, e.HasScore, e.Raw
		if !s.HasScore {
			missing++
			missingNames = append(missingNames, s.Name)
		}
	}

	isNA := func(s *Student) bool {
		return strings.HasPrefix(strings.ToUpper(strings.TrimSpace(s.ScoreRaw)), "#N/A")
	}

	// ── 排名 ──────────────────────────────
	// 1. 有分且非新增：总分降序，同分保持名单顺序
	// 2. 无分非新增：#N/A 一组、其它（缺考等）一组，各保持名单顺序
	// 3. 新增（高亮）：排最后，有分者降序，无分者保持名单顺序
	var gScored, gNA, gOther, gAddScored, gAddNone []*Student
	for _, s := range students {
		switch {
		case s.Added && s.HasScore:
			gAddScored = append(gAddScored, s)
		case s.Added:
			gAddNone = append(gAddNone, s)
		case s.HasScore:
			gScored = append(gScored, s)
		case isNA(s):
			gNA = append(gNA, s)
		default:
			gOther = append(gOther, s)
		}
	}
	sort.SliceStable(gScored, func(i, j int) bool { return gScored[i].Score > gScored[j].Score })
	sort.SliceStable(gAddScored, func(i, j int) bool { return gAddScored[i].Score > gAddScored[j].Score })

	ranked := make([]*Student, 0, len(students))
	ranked = append(ranked, gScored...)
	ranked = append(ranked, gNA...)
	ranked = append(ranked, gOther...)
	ranked = append(ranked, gAddScored...)
	ranked = append(ranked, gAddNone...)
	for i, s := range ranked {
		s.Rank = i + 1
	}

	if missing > 0 {
		names := missingNames
		tail := ""
		if len(names) > 6 {
			names, tail = names[:6], "…"
		}
		warn(fmt.Sprintf("  ⚠ %d 位学生没有总分（缺考/#N/A/空），已排在最后：%s%s",
			missing, strings.Join(names, "、"), tail))
	}
	nAdded := 0
	for _, s := range students {
		if s.Added {
			nAdded++
		}
	}
	if nAdded > 0 {
		warn(fmt.Sprintf("  · 名单高亮 %d 人（新增/后补）→ 排在最后", nAdded))
	}
	nSpecial := 0
	for _, s := range ranked {
		if s.Special {
			nSpecial++
		}
	}
	if nSpecial > 0 {
		warn(fmt.Sprintf("  · 名单标黄 %d 人 → 语数外固定第 %d 考场", nSpecial, opt.SpecialRoom))
	}

	// fill 把 pool 依次填入 counts 指定的各考场
	fill := func(pool []*Student, counts map[int]int, subj string, skip ...int) map[int]map[int]*Student {
		placed := map[int]map[int]*Student{}
		idx := 0
		var order []int
		for rno := range counts {
			order = append(order, rno)
		}
		sort.SliceStable(order, func(i, j int) bool {
			return roomSeats[order[i]][0] < roomSeats[order[j]][0]
		})
		skipSet := map[int]bool{}
		for _, s := range skip {
			skipSet[s] = true
		}
		for _, rno := range order {
			if skipSet[rno] {
				continue
			}
			labels := append([]int(nil), roomSeats[rno]...)
			cnt := counts[rno]
			if cnt > len(labels) {
				extra := cnt - len(labels)
				pre := len(labels)
				for i := 1; i <= extra; i++ {
					labels = append(labels, labels[len(labels)-1]+1)
				}
				roomSeats[rno] = labels
				warn(fmt.Sprintf("  · 场次「%s」考场%d 需 %d 座（模板预印 %d 座），为本场末考场，座位号顺延至 %d",
					subj, rno, cnt, pre, labels[len(labels)-1]))
			}
			for k := 0; k < cnt; k++ {
				if idx >= len(pool) {
					break
				}
				s := pool[idx]
				s.Seats[subj] = labels[k]
				if placed[rno] == nil {
					placed[rno] = map[int]*Student{}
				}
				placed[rno][k] = s
				idx++
			}
		}
		if idx < len(pool) {
			warn(fmt.Sprintf("  ✗ 场次「%s」还有 %d 人没排下（座位不足）", subj, len(pool)-idx))
		}
		return placed
	}

	var plans []Plan
	for _, sess := range arr.Sessions {
		name := sess.Subject
		keys := []string{name}
		if name == "语数外" {
			keys = []string{"语文", "数学", "英语"}
		}
		for _, key := range keys {
			if name == "语数外" && key != "语文" {
				continue // 三场座位相同，只算一次
			}
			var takers, selves []*Student
			for _, s := range ranked {
				if s.Takes(key) {
					takers = append(takers, s)
				} else {
					selves = append(selves, s)
				}
			}
			roomsOut := map[int]map[int]*Student{}
			if name == "语数外" {
				var sp, rest []*Student
				for _, s := range takers {
					if s.Special {
						sp = append(sp, s)
					} else {
						rest = append(rest, s)
					}
				}
				sort.SliceStable(sp, func(i, j int) bool { return sp[i].Rank < sp[j].Rank })
				counts := map[int]int{}
				for rno, n := range sess.Takers {
					if rno != opt.SpecialRoom {
						counts[rno] = n
					}
				}
				mergeInto(roomsOut, fill(rest, counts, key, opt.SpecialRoom))

				spLabels := roomSeats[opt.SpecialRoom]
				for i, s := range sp {
					if i >= len(spLabels) {
						break
					}
					s.Seats[key] = spLabels[i]
					if roomsOut[opt.SpecialRoom] == nil {
						roomsOut[opt.SpecialRoom] = map[int]*Student{}
					}
					roomsOut[opt.SpecialRoom][i] = s
				}
				// 数学 / 英语 座位沿用
				for _, d := range roomsOut {
					for _, s := range d {
						s.Seats["数学"] = s.Seats[key]
						s.Seats["英语"] = s.Seats[key]
					}
				}
			} else {
				mergeInto(roomsOut, fill(takers, sess.Takers, key))
				if len(selves) > 0 {
					mergeInto(roomsOut, fill(selves, sess.Self, key))
				}
			}
			plans = append(plans, Plan{Subject: key, Special: name == "语数外", Rooms: roomsOut})
		}
	}
	return &Arranged{Ranked: ranked, Plans: plans, RoomSeat: roomSeats}, nil
}

func mergeInto(dst, src map[int]map[int]*Student) {
	for rno, d := range src {
		if dst[rno] == nil {
			dst[rno] = map[int]*Student{}
		}
		for k, s := range d {
			dst[rno][k] = s
		}
	}
}

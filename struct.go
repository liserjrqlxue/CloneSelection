package main

import (
	"log"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/stringsUtil"
	"github.com/xuri/excelize/v2"
)

type JPs struct {
	List []*JPPanel
	Map  map[string]*JPPanel
}

func (jps *JPs) CreateDetailedList(xlsx *excelize.File, sheet string) {
	simpleUtil.HandleError(xlsx.NewSheet(sheet))

	// 设置格式
	xlsx.SetColWidth(sheet, "B", "B", 16)
	xlsx.SetColWidth(sheet, "D", "F", 16)
	xlsx.SetColWidth(sheet, "G", "G", 70)
	xlsx.SetColWidth(sheet, "H", "H", 40)

	// 设置表头
	index := 0
	cellName := simpleUtil.HandleError(excelize.CoordinatesToCellName(1, index+1))
	simpleUtil.CheckErr(
		xlsx.SetSheetRow(sheet, cellName, &DetailedListTitle),
	)

	for _, jpPanel := range jps.List {
		for _, segmentInfo := range jpPanel.Segments {
			index++
			cellName = simpleUtil.HandleError(excelize.CoordinatesToCellName(1, index+1))
			simpleUtil.CheckErr(
				xlsx.SetSheetRow(
					sheet, cellName,
					&[]any{
						index,
						segmentInfo.ID,
						segmentInfo.Length,
						segmentInfo.SequencePrimer,
						len(segmentInfo.CloneIDs),
						segmentInfo.SequenceCount,
						segmentInfo.ID + "-" + strings.Join(segmentInfo.CloneIDs, "、"),
						segmentInfo.ID + "-" + strings.Join(segmentInfo.SequenceCloneIDs, "、"),
					},
				),
			)
		}
	}
}

type JPPanel struct {
	ID        string
	Index     int
	Date      string
	TY        bool
	Segments  []*Segment
	Gels      [4][25]string
	GelLayout string
}

func (jpPanel *JPPanel) Gels2Segments() {
	var (
		gels       = jpPanel.Gels
		jpCloneMax = MaxJPClone
		// 非ladder克隆序号
		index = 0
	)
	if jpPanel.TY {
		jpCloneMax = MaxJPCloneTY
	}

	// 校验GelLayout
	if gels[0][0] == "/" && gels[1][0] == "/" && gels[2][0] == "/" && gels[3][0] == "/" {
		jpPanel.GelLayout = "first ladder"
		if jpPanel.TY {
			jpPanel.GelLayout += " for TY"
		}
	}
	if gels[0][16] == "/" && gels[1][8] == "/" && gels[2][16] == "/" && gels[3][8] == "/" {
		jpPanel.GelLayout = "partition ladder"
	}
	if jpPanel.GelLayout == "" {
		log.Fatalf("Unknown Gels Layout for [%s]:%+v", jpPanel.ID, gels)
	}

	// 遍历Gels, 更新 Segment
	for j := range MaxGelRow {
		for k := range MaxGelCol {
			gel := gels[j][k]
			if gel != "/" {
				if gel == "Y" {
					segmentIndex := index / jpCloneMax
					segment := jpPanel.Segments[segmentIndex]
					cloneID := strconv.Itoa(index%jpCloneMax + 1)
					segment.CloneIDs = append(segment.CloneIDs, cloneID)
					clone := &Clone{
						ID:    cloneID,
						Index: index%jpCloneMax + 1,
					}
					segment.CloneMap[cloneID] = clone
				}
				index++
			}
		}
	}

	// 更新 segment.SequenceCount
	maxCloneSelect := MaxCloneSelect
	if jpPanel.TY {
		maxCloneSelect = MaxCloneSelectTY
	}
	for j := range jpPanel.Segments {
		segment := jpPanel.Segments[j]
		segment.SequenceCount = min(maxCloneSelect, len(segment.CloneIDs))
		segment.SequenceCloneIDs = segment.CloneIDs[:segment.SequenceCount]
	}
}

func (jpPanel *JPPanel) AddSegment(item map[string]string) *Segment {
	segment := &Segment{
		ID:             item["片段名称"],
		JPID:           jpPanel.ID,
		Length:         stringsUtil.Atoi(item["片段长度"]),
		SequencePrimer: item["测序引物"],
		Note2Product:   item["备注（to生产）"],
		CloneMap:       make(map[string]*Clone),
		FromPanel:      make(map[string]string),
	}

	for primer := range strings.SplitSeq(segment.SequencePrimer, "、") {
		switch primer {
		case "T7":
			segment.T7Primer = true
		case "T7-Term":
			segment.T7TermPrimer = true
		default:
			segment.OtherPrimers = append(segment.OtherPrimers, primer)
		}
	}

	jpPanel.Segments = append(jpPanel.Segments, segment)

	return segment
}

type Segment struct {
	ID             string
	JPID           string
	Length         int
	SequencePrimer string
	T7Primer       bool
	T7TermPrimer   bool
	OtherPrimers   []string
	Note2Product   string
	CloneIDs       []string
	// CloneIDs -> true
	CloneMap map[string]*Clone
	// 送测克隆数
	SequenceCount    int
	SequenceCloneIDs []string
	// Cell Name
	FromPanel map[string]string
}

type Clone struct {
	Index    int
	ID       string
	FromCell string
	ToCell   string
}

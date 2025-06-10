package main

import (
	"fmt"
	"log"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/stringsUtil"
	"github.com/xuri/excelize/v2"
)

type JPs struct {
	List []*JPPanel
	Map  map[string]*JPPanel

	SCs []*Segment
	TYs []*Segment
}

func (jps *JPs) SplitList() {
	for _, jpPanel := range jps.List {
		if jpPanel.TY {
			jps.TYs = append(jps.TYs, jpPanel.Segments...)
		} else {
			jps.SCs = append(jps.SCs, jpPanel.Segments...)
		}
	}
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

func (jps *JPs) CreateFromPanel(xlsx *excelize.File, sheet string, bgStyleMap map[int]int) {
	simpleUtil.HandleError(xlsx.NewSheet(sheet))

	xlsx.SetColWidth(sheet, "A", "B", 12)
	xlsx.SetColWidth(sheet, "C", "N", 18)

	for i, jpPanel := range jps.List {
		jpPanel.AddFromPanel(xlsx, sheet, i, bgStyleMap)
	}
}
func (jps *JPs) CreateToPanel(xlsx *excelize.File, sheet string, bgStyleMap map[int]int) {
	simpleUtil.HandleError(xlsx.NewSheet(sheet))

	xlsx.SetColWidth(sheet, "A", "D", 16)
	xlsx.SetColWidth(sheet, "E", "P", 18)

	var (
		panelSCIndex = 0
		panelTYIndex = 0
	)

	panelSCIndex = jps.AddSCPanel(xlsx, sheet, panelTYIndex, bgStyleMap)
	panelTYIndex = jps.AddTYPanel(xlsx, sheet, panelSCIndex, bgStyleMap)
}

func (jps *JPs) AddSCPanel(xlsx *excelize.File, sheet string, offset int, bgStyleMap map[int]int) (panelSCIndex int) {
	var (
		date = jps.List[0].Date

		panelID   string
		rowOffset int

		cellName string
	)
	panelSCIndex = 0
	for i, segment := range jps.SCs {
		// 板内片段序号, 也是克隆列号
		var segmentIndex = i % MaxSegmentSelectSC

		// 初始化输出板
		if segmentIndex == 0 {
			panelID = fmt.Sprintf("%s-SC%d", date, panelSCIndex+1)

			rowOffset = (panelSCIndex + offset) * TabelRow
			InitToPanel(xlsx, sheet, panelID, rowOffset, bgStyleMap)

			panelSCIndex++
		}

		cellName = CoordinatesToCellName(
			2+segmentIndex/PanelRow,
			2+rowOffset+segmentIndex%PanelRow,
		)
		simpleUtil.CheckErr(xlsx.SetCellStr(sheet, cellName, segment.ID))
		for j, cloneID := range segment.SequenceCloneIDs {
			clone := segment.CloneMap[cloneID]
			toCell := CoordinatesToCellName(
				segmentIndex+5,
				rowOffset+j+2,
			)
			clone.ToPanel = panelID
			clone.ToCell = toCell
			simpleUtil.CheckErr(xlsx.SetCellStr(sheet, toCell, clone.Name))
		}
	}

	return
}

func (jps *JPs) AddTYPanel(xlsx *excelize.File, sheet string, offset int, bgStyleMap map[int]int) (panelTYIndex int) {
	var (
		date = jps.List[0].Date

		panelID   string
		rowOffset int

		cellName string
	)
	panelTYIndex = 0
	for i, segment := range jps.TYs {
		// 板内片段序号, %PanelRow 也是克隆行号
		var segmentIndex = i % MaxSegmentSeclectTY

		// 初始化输出板
		if segmentIndex == 0 {
			panelID = fmt.Sprintf("%s-TY%d", date, panelTYIndex+1)
			rowOffset = (panelTYIndex + offset) * TabelRow
			InitToPanel(xlsx, sheet, panelID, rowOffset, bgStyleMap)

			panelTYIndex++
		}

		var segmentRow = 2 + rowOffset + segmentIndex%PanelRow
		cellName = CoordinatesToCellName(
			2+segmentIndex/PanelRow,
			segmentRow,
		)
		simpleUtil.CheckErr(xlsx.SetCellStr(sheet, cellName, segment.ID))
		for j, cloneID := range segment.SequenceCloneIDs {
			clone := segment.CloneMap[cloneID]
			toCell := CoordinatesToCellName(
				5+j+segmentIndex/PanelRow*PanelRow/2,
				segmentRow,
			)
			clone.ToPanel = panelID
			clone.ToCell = toCell
			simpleUtil.CheckErr(xlsx.SetCellStr(sheet, toCell, clone.Name))
		}
	}
	return
}

func (jps *JPs) CreateYK(xlsx *excelize.File, sheet string, bgStyleMap map[int]int) {
	simpleUtil.HandleError(xlsx.NewSheet(sheet))

	var rIdx = 2
	for _, segment := range jps.SCs {
		var (
			segmentLength = "1-1000"
			primers       []string
			primersName   string
			orientation   = ""
		)

		if segment.Length > 1000 {
			segmentLength = "1001-2000"
			if segment.Length > 2000 {
				log.Fatalf("ID[%s]长度[%d]超标", segment.ID, segment.Length)
			}
		}

		if segment.T7Primer {
			primers = append(primers, "T7")
		}
		if segment.T7TermPrimer {
			primers = append(primers, "T7-Term")
		}
		primersName = strings.Join(primers, "、")

		if segment.T7Primer && segment.T7TermPrimer {
			orientation = "C"
		} else if segment.T7Primer {
			orientation = "A"
		} else if segment.T7TermPrimer {
			orientation = "B"
		}

		for _, cloneID := range segment.SequenceCloneIDs {
			rIdx++
			cellName := CoordinatesToCellName(1, rIdx)

			clone := segment.CloneMap[cloneID]
			simpleUtil.CheckErr(
				xlsx.SetSheetRow(
					sheet, cellName,
					&[]any{
						clone.Name,
						"A", "U1AT",
						segmentLength,
						"A", "A",
						primersName,
						"",
						orientation,
						"样本与表格一一对应",
					},
				),
			)
		}
	}
	InitYK(xlsx, sheet, rIdx, bgStyleMap)
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

var (
	regPanelID = regexp.MustCompile(`(\d+)JP-(\d+)`)
)

func (jpPanel *JPPanel) ParseID() error {
	match := regPanelID.FindStringSubmatch(jpPanel.ID)
	if match == nil {
		return fmt.Errorf("panelID format error![%s]", jpPanel.ID)
	}
	index, err := strconv.Atoi(match[2])
	if err != nil {
		return fmt.Errorf("panelID format error![%s][%w]", jpPanel.ID, err)
	}
	jpPanel.Index = index
	_, err = time.Parse("20060102", match[1])
	jpPanel.Date = match[1]
	if err != nil {
		return fmt.Errorf("panelID format error![%s][%w]", jpPanel.ID, err)
	}
	return nil
}

func (jpPanel *JPPanel) Gels2Segments() {
	var (
		gels       = jpPanel.Gels
		jpCloneMax = MaxJPCloneSC
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
						Name:  fmt.Sprintf("%s-%s", segment.ID, cloneID),
						Index: index%jpCloneMax + 1,
					}
					segment.CloneMap[cloneID] = clone
				}
				index++
			}
		}
	}

	// 更新 segment.SequenceCount
	maxCloneSelect := MaxCloneSelectSC
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

	for primer := range strings.SplitSeq(segment.SequencePrimer, "；") {
		switch primer {
		case "T7":
			segment.T7Primer = true
		case "T7term":
			segment.T7TermPrimer = true
		default:
			segment.OtherPrimers = append(segment.OtherPrimers, primer)
		}
	}

	jpPanel.Segments = append(jpPanel.Segments, segment)

	return segment
}

func (jpPanel *JPPanel) AddFromPanel(xlsx *excelize.File, sheet string, i int, bgStyleMap map[int]int) {
	jpID := jpPanel.ID
	segmentIDs := jpPanel.Segments

	maxSegment := MaxSegmentSC
	maxJPClone := MaxJPCloneSC
	if jpPanel.TY {
		maxSegment = MaxSegmentTY
		maxJPClone = MaxJPCloneTY
	}

	if len(segmentIDs) > maxSegment {
		log.Fatalf("片段超限[%d > %d][%s:%t][%+v]", len(segmentIDs), maxSegment, jpID, jpPanel.TY, segmentIDs)
	}

	cellName := CoordinatesToCellName(1, i*TabelRow+1)
	simpleUtil.CheckErr(
		xlsx.SetSheetRow(sheet, cellName, &[]any{jpID, jpID}),
	)
	cellName = CoordinatesToCellName(3, i*TabelRow+1)
	simpleUtil.CheckErr(
		xlsx.SetSheetRow(sheet, cellName, &PanelColTitle),
	)
	cellName = CoordinatesToCellName(2, i*TabelRow+2)
	simpleUtil.CheckErr(
		xlsx.SetSheetCol(sheet, cellName, &PanelRowTitle),
	)
	// 合并单元格
	simpleUtil.CheckErr(
		xlsx.MergeCell(
			sheet,
			CoordinatesToCellName(1, i*TabelRow+1),
			CoordinatesToCellName(1, i*TabelRow+1+PanelRow),
		),
	)

	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(1, i*TabelRow+1),
			CoordinatesToCellName(14, i*TabelRow+9),
			bgStyleMap[-1],
		),
	)
	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(2, i*TabelRow+1),
			CoordinatesToCellName(2, i*TabelRow+9),
			bgStyleMap[3],
		),
	)
	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(2, i*TabelRow+1),
			CoordinatesToCellName(14, i*TabelRow+1),
			bgStyleMap[3],
		),
	)

	cloneIndex := 0
	for j := range segmentIDs {
		segment := segmentIDs[j]
		// fmt.Printf("%s\t%c\n", jpID, row)
		for k := range maxJPClone {
			row := cloneIndex / PanelCol
			col := 1 + cloneIndex%PanelCol
			fromCel := fmt.Sprintf("%c%d", row+'A', col)
			cloneID := strconv.Itoa(k + 1)
			segment.FromPanel[cloneID] = fromCel
			cellName = simpleUtil.HandleError(excelize.CoordinatesToCellName(2+col, row+2+i*TabelRow))
			ID := fmt.Sprintf("%s-%s", segment.ID, cloneID)
			simpleUtil.CheckErr(xlsx.SetCellStr(sheet, cellName, ID))
			if clone, ok := segment.CloneMap[cloneID]; ok {
				clone.FromCell = fromCel
				clone.FromPanel = jpPanel.ID
				simpleUtil.CheckErr(xlsx.SetCellStyle(sheet, cellName, cellName, bgStyleMap[j%3]))
			}
			cloneIndex++
		}
	}
	// // 合并单元格
	// simpleUtil.CheckErr(
	// 	xlsx.MergeCell(
	// 		sheet,
	// 		CoordinatesToCellName(1, i*TabelRow+1),
	// 		CoordinatesToCellName(1, i*TabelRow+9),
	// 	),
	// )

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
	Index int
	ID    string
	Name  string

	FromPanel string
	FromCell  string

	ToPanel string
	ToCell  string
}

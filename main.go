package main

import (
	"flag"
	"log"
	"regexp"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

// global
var (
	InputSheet = "胶图判定"
	// 每个JP板的片段个数
	MaxSegmentRow = 8
	MaxSegmentSC  = 6
	MaxSegmentTY  = 8
	// 每个片段的克隆个数
	MaxJPCloneSC = 16
	MaxJPCloneTY = 12
	// 每个输出板上的片段个数 Clone*Segment=96
	MaxSegmentSelectSC  = 12
	MaxSegmentSeclectTY = 16
	// 每个输出板上的片段的克隆个数
	MaxCloneSelectSC = 8
	MaxCloneSelectTY = 6

	// 胶图 25x4
	MaxGelCol = 25
	MaxGelRow = 4

	// 96孔板
	PanelCol      = 12
	PanelRow      = 8
	TabelRow      = 10
	PanelColTitle = []int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}
	PanelRowTitle = []string{"A", "B", "C", "D", "E", "F", "G", "H"}

	isTY = regexp.MustCompile(`-TY\d+`)

	ColName12 []string

	// 输出1-清单
	DetailedListTitle = []string{
		"序号",
		"片段名称",
		"片段长度",
		"测序引物",
		"条带正确克隆数",
		"送测克隆数",
		"条带正确克隆",
		"送测克隆",
	}

	// 测序YK
	YKTitle = []string{
		"*样品名称",
		"*样品类型",
		"*载体名称",
		"*片段大小",
		"*抗生素类型",
		"引物信息",
		"",
		"",
		"*测序要求",
		"备注",
	}
	YKPrimerInfoTitle = []string{
		"*引物类别",
		"*引物名称",
		"*自备引物",
	}

	// 测序GWZ
	GWZTitle = []string{
		"样品编号",
		"*样品名称",
		"*片段长度(bp)",
		"*载体名称及抗性",
		"通用引物1",
		"通用引物2",
		`自备引物(如有多个请用分号";"隔开)`,
		"特殊要求",
		"测通",
		"返还样品",
		"样品备注（可备注具体片段长度）",
	}
)

// flag
var (
	input = flag.String(
		"i",
		"",
		"input excel",
	)
	prefix = flag.String(
		"p",
		"",
		"output prefix",
	)
)

func init() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		log.Fatal("-i required")
	}
	if *prefix == "" {
		*prefix = strings.TrimSuffix(*input, ".xlsx")
	}

	for i := range PanelCol {
		ColName12 = append(ColName12, strconv.Itoa(i+1))
	}
}

func main() {
	var jps = LoadInput(*input, InputSheet)
	jps.SplitList()
	// 由Gels更新Segment
	for _, jpPanel := range jps.List {
		jpPanel.Gels2Segments()
	}

	var (
		xlsx       = excelize.NewFile()
		sheet      = ""
		bgStyleMap = CreateStyles(xlsx)
	)

	// 输出1-清单
	sheet = "清单"
	jps.CreateDetailedList(xlsx, sheet)

	// 输出2-选择孔图
	sheet = "选择孔图"
	jps.CreateFromPanel(xlsx, sheet, bgStyleMap)

	// 输出2-输出孔图
	sheet = "输出孔图"
	jps.CreateToPanel(xlsx, sheet, bgStyleMap)

	// 测序YK
	sheet = "测序YK"
	jps.CreateYK(xlsx, sheet, bgStyleMap)

	// 测序GWZ
	sheet = "测序GWZ"
	jps.CreateGWZ(xlsx, sheet, bgStyleMap)

	simpleUtil.CheckErr(xlsx.DeleteSheet("Sheet1"))
	simpleUtil.CheckErr(xlsx.SaveAs(*prefix + ".result.xlsx"))
}

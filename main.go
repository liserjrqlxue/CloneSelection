package main

import (
	"flag"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/PrimerDesigner/v2/pkg/cy0130"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

const (
	// 每个JP板的片段个数, 也即行数
	MaxSegmentRow = 8

	// 胶图 25x4
	MaxGelCol = 25
	MaxGelRow = 4

	// s输出 6孔板
	PanelCol = 12
	PanelRow = 8
	TabelRow = 10
)

// global
var (
	InputSheet = "胶图判定"
	// 每种JP的片段个数
	MaxSegmentSC = 6
	MaxSegmentTY = 8
	// 每个片段的克隆个数
	MaxJPCloneSC = 24
	MaxJPCloneTY = 12
	// 每个输出板上的片段个数 Clone*Segment=96
	MaxSegmentSelectSC  = 8
	MaxSegmentSeclectTY = 16
	// 每个输出板上的片段的克隆个数
	MaxCloneSelectSC = 12
	MaxCloneSelectTY = 6

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

	// Transfer
	TransferTitle = "SourcePlateLable,SourceWellPosition,DesPlateLable,DesWellPosition,Volume,BarCode,ChangeTip,PreAspirateMixNumber,PreAspirateMixVolume,PostDispenseMixNumber,PostDispenseMixVolume,LiquidClass,Pause"
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
	outDir = flag.String(
		"o",
		"",
		"output dir",
	)
)

func init() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		log.Fatal("-i required")
	}
	if *outDir == "" {
		*outDir = strings.TrimSuffix(*input, ".xlsx")
	}
	if *prefix == "" {
		*prefix = filepath.Base(*outDir)
	}

	for i := range PanelCol {
		ColName12 = append(ColName12, strconv.Itoa(i+1))
	}
}

func main() {
	simpleUtil.CheckErr(os.MkdirAll(*outDir, 0755))

	var patterns []string

	var jps, xlsx = LoadInput(*input, InputSheet)
	// 由Gels更新Segment
	for _, jpPanel := range jps.List {
		jpPanel.Gels2Segments()
	}
	// 拆分SC TY
	jps.SplitList()
	jps.WriteSheets(xlsx)

	var result = *prefix + ".result.xlsx"
	patterns = append(patterns, result)
	simpleUtil.CheckErr(xlsx.SaveAs(filepath.Join(*outDir, result)))

	var jpsList = jps.SplitJPs(4)
	for i, jps := range jpsList {
		var (
			xlsx        = excelize.NewFile()
			tagPrefix   = *prefix + "-" + string(rune('A'+i))
			tagResult   = tagPrefix + ".result.xlsx"
			tagTransfer = tagPrefix + ".Transfer.csv"
		)
		patterns = append(patterns, tagResult)
		patterns = append(patterns, tagTransfer)
		tagResult = filepath.Join(*outDir, tagResult)
		tagTransfer = filepath.Join(*outDir, tagTransfer)

		jps.SplitList()

		jps.WriteSheets(xlsx)

		simpleUtil.CheckErr(xlsx.DeleteSheet("Sheet1"))
		simpleUtil.CheckErr(xlsx.SaveAs(tagResult))

		jps.CreateBioHandler()
		jps.WriteTransfer(tagTransfer)
	}

	// Compress-Archive
	// simpleUtil.CheckErr(
	// 	sge.Run(
	// 		"powershell",
	// 		"Compress-Archive",
	// 		"-DestinationPath", *outDir+".zip",
	// 		"-Path", *outDir+"/*",
	// 		"-Force",
	// 	),
	// )
	absOutDir := simpleUtil.HandleError(filepath.Abs(*outDir))
	simpleUtil.CheckErr(cy0130.CompressArchive(absOutDir, absOutDir+".zip", patterns))
}

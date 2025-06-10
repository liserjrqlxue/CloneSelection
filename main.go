package main

import (
	"flag"
	"fmt"
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
	simpleUtil.CheckErr(xlsx.DeleteSheet("Sheet1"))
	simpleUtil.CheckErr(xlsx.SaveAs(*prefix + ".result.xlsx"))

	fmt.Println("==END==")

	// 测序YK
	index := 0
	fmt.Println("==测序YK==")
	sheet = "测序YK"
	for _, jpPanel := range jps.List {
		segmentIDs := jpPanel.Segments
		for j := range segmentIDs {
			segmentInfo := segmentIDs[j]
			for k := range segmentInfo.CloneIDs {
				cloneID := segmentInfo.CloneIDs[k]
				ID := fmt.Sprintf("%s-%s", segmentInfo.ID, cloneID)
				segmentLength := "1-1000"
				if segmentInfo.Length > 1000 {
					segmentLength = "1001-2000"
					if segmentInfo.Length > 2000 {
						log.Fatalf("ID[%s]长度[%d]超标", ID, segmentInfo.Length)
					}
				}
				var primers []string
				if segmentInfo.T7Primer {
					primers = append(primers, "T7")
				}
				if segmentInfo.T7TermPrimer {
					primers = append(primers, "T7-Term")
				}
				var orientation = ""
				if segmentInfo.T7Primer && segmentInfo.T7TermPrimer {
					orientation = "C"
				} else if segmentInfo.T7Primer {
					orientation = "A"
				} else if segmentInfo.T7TermPrimer {
					orientation = "B"
				}

				fmt.Printf(
					"%s\n",
					strings.Join(
						[]string{
							ID,
							"A",
							"U1AT",
							segmentLength,
							"A",
							"A",
							strings.Join(primers, "、"),
							"",
							orientation,
							"样本与表格一一对应",
						},
						"\t",
					))
			}
			index++
		}
	}

}

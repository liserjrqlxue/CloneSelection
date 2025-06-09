package main

import (
	"flag"
	"fmt"
	"log"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

// global
var (
	InputSheet = "胶图判定"
	// 每个JP板的片段个数
	MaxSegmentRow = 8
	MaxSegment    = 6
	MaxSegmentTY  = 8
	// 每个片段的克隆个数
	MaxJPClone       = 16
	MaxJPCloneTY     = 12
	MaxCloneSelect   = 8
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

	simpleUtil.CheckErr(xlsx.DeleteSheet("Sheet1"))
	simpleUtil.CheckErr(xlsx.SaveAs(*prefix + ".result.xlsx"))

	// 输出2-输出孔图
	fmt.Println("==输出2-输出孔图==")
	index := 0
	panelID := ""
	for i, jpPanel := range jps.List {
		if i%2 == 0 {
			// new panel
			date := time.Now().Format("20050102")
			panelID = fmt.Sprintf("[%s]-SC%d", date, (i+1)/2)
			fmt.Printf("%s\t片段名称1\t片段名称2\t%s\t%s", panelID, panelID, strings.Join(ColName12, "\t"))
			index = 0
		}
		segmentIDs := jpPanel.Segments
		for j := range segmentIDs {
			segmentInfo := segmentIDs[j]
			fmt.Printf("%s\t%s\t/\t%c", panelID, segmentInfo.ID, 'A'+index)
			for k := range segmentInfo.CloneIDs {
				cloneID := segmentInfo.CloneIDs[k]
				fmt.Printf("\t%s-%s", segmentInfo.ID, cloneID)
			}
			index++
		}
	}
	fmt.Println("==END==")

	// 测序YK
	fmt.Println("==测序YK==")
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

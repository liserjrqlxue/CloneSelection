package main

import (
	"flag"
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"
)

// global
var (
	InputSheet = "胶图判定"
	// 每个JP板的片段个数
	MaxSegment = 6
	// 每个片段的克隆个数
	MaxClone       = 16
	MaxCloneSelect = 12

	// 胶图 25x4
	MaxGelCol = 25
	MaxGelRow = 4

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

	for i := range MaxCloneSelect {
		ColName12 = append(ColName12, strconv.Itoa(i+1))
	}
}

func main() {
	var jpPanelMap, jpPanelList = LoadInput(*input, InputSheet)
	// 由Gels更新Segment
	for i := range jpPanelList {
		jpID := jpPanelList[i]
		jpPanel := jpPanelMap[jpID]
		jpPanel.Gels2Segments()
	}

	// 输出1-清单
	fmt.Println("==输出1-清单==")
	fmt.Println(strings.Join(DetailedListTitle, "\t"))
	i := 0
	for _, jpID := range jpPanelList {
		jpPanel := jpPanelMap[jpID]
		for _, segmentInfo := range jpPanel.Segments {
			i++
			fmt.Printf(
				"%d\t%s\t%d\t%s\t%d\t%d\t%s\t%s\n",
				i,
				segmentInfo.ID,
				segmentInfo.Length,
				segmentInfo.SequencePrimer,
				len(segmentInfo.CloneIDs),
				min(MaxCloneSelect, len(segmentInfo.CloneIDs)),
				segmentInfo.ID+"-"+strings.Join(segmentInfo.CloneIDs, "、"),
				segmentInfo.ID+"-"+strings.Join(segmentInfo.CloneIDs[:min(MaxCloneSelect, len(segmentInfo.CloneIDs))], "、"),
			)
		}
	}
	fmt.Println("==END==")

	// 输出2-选择孔图
	fmt.Println("==输出2-选择孔图==")
	for _, jpID := range jpPanelList {
		jpPanel := jpPanelMap[jpID]
		segmentIDs := jpPanel.Segments
		if len(segmentIDs) > 6 {
			log.Fatalf("片段超限[%s][%+v]", jpID, segmentIDs)
		}

		fmt.Printf("%s\t%s\t%s", jpID, jpID, strings.Join(ColName12, "\t"))
		fmt.Println()

		for j := range 6 {
			if j >= len(segmentIDs) {
				fmt.Printf("%s\t%c\n", jpID, 'A'+j*2)
				fmt.Printf("%s\t%c\n", jpID, 'A'+j*2+1)
				continue
			}
			segmentInfo := segmentIDs[j]
			// for k:=range 24{
			// 	row:='A'+j*2+k/12
			// }
			row := 'A' + j*2
			fmt.Printf("%s\t%c\n", jpID, row)
			for k := range MaxCloneSelect {
				cloneID := strconv.Itoa(k + 1)
				segmentInfo.FromPanel[cloneID] = fmt.Sprintf("%c%d", row, k+1)
				fmt.Printf("\t%s:%s-%s:%t", segmentInfo.FromPanel[cloneID], segmentInfo.ID, cloneID, segmentInfo.CloneStatus[cloneID])
			}
			fmt.Printf("%s\t%c", jpID, 'A'+j*2+1)
			for k := range MaxCloneSelect {
				cloneID := strconv.Itoa(k + 1 + MaxCloneSelect)
				segmentInfo.FromPanel[cloneID] = fmt.Sprintf("%c%d", row, k+1)
				fmt.Printf("\t%s:%s-%s:%t", segmentInfo.FromPanel[cloneID], segmentInfo.ID, cloneID, segmentInfo.CloneStatus[cloneID])
			}
		}
	}
	fmt.Println("==END==")

	// 输出2-输出孔图
	fmt.Println("==输出2-输出孔图==")
	index := 0
	panelID := ""
	for i, jpID := range jpPanelList {
		if i%2 == 0 {
			// new panel
			date := time.Now().Format("20050102")
			panelID = fmt.Sprintf("[%s]-SC%d", date, (i+1)/2)
			fmt.Printf("%s\t片段名称1\t片段名称2\t%s\t%s", panelID, panelID, strings.Join(ColName12, "\t"))
			index = 0
		}
		jpPanel := jpPanelMap[jpID]
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
	for i := range jpPanelList {
		jpID := jpPanelList[i]
		jpPanel := jpPanelMap[jpID]
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

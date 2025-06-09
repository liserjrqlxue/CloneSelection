package main

import (
	"flag"
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/stringsUtil"

	"github.com/xuri/excelize/v2"
)

// global
var (
	MaxClone  = 12
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

	for i := range MaxClone {
		ColName12 = append(ColName12, strconv.Itoa(i+1))
	}
}

func main() {
	var (
		xlsx  = simpleUtil.HandleError(excelize.OpenFile(*input))
		rows  = simpleUtil.HandleError(xlsx.GetRows("胶图判定(人工输入)"))
		title []string

		JPID   string
		JPInfo = make(map[string]*SegmentInfo)
		// JPID -> []string
		SegmentList = make(map[string][]string)
		JPlist      []string
	)
	for i := range rows {
		if i == 0 {
			title = rows[0]
			continue
		}
		item := make(map[string]string)
		for j := range rows[i] {
			if j < len(title) {
				item[title[j]] = rows[i][j]
			}
		}
		jpID, ok := item["JP-日期"]
		if i%4 == 1 {
			if !ok {
				log.Fatalf("JP-日期:A%d empty", i+1)
			} else {
				JPID = jpID
			}
		} else {
			if JPID == "" {
				log.Fatalf("JP-日期:before A%d empty", i+1)
			}
			if !ok {
				item["JP-日期"] = JPID
			} else {
				log.Fatalf("JP-日期:A%d not empty", i+1)
			}
		}
		segmentInfo := NewSegmentInfo(item)
		if _, ok := JPInfo[segmentInfo.ID]; ok {
			log.Fatalf("片段重复:[%d:%s]", i+1, segmentInfo.ID)
		}
		JPInfo[segmentInfo.ID] = segmentInfo
		SegmentList[JPID] = append(SegmentList[JPID], segmentInfo.ID)
		JPlist = append(JPlist, JPID)
	}

	// 输出1-清单
	fmt.Println("==输出1-清单==")
	fmt.Println(strings.Join(DetailedListTitle, "\t"))
	i := 0
	for _, jpID := range JPlist {
		for _, segmentID := range SegmentList[jpID] {
			segmentInfo := JPInfo[segmentID]
			i++
			fmt.Printf(
				"%d\t%s\t%d\t%s\t%d\t%d\t%s\t%s\n",
				i,
				segmentInfo.ID,
				segmentInfo.Length,
				segmentInfo.SequencePrimer,
				len(segmentInfo.CloneIDs),
				min(MaxClone, len(segmentInfo.CloneIDs)),
				segmentInfo.ID+"-"+strings.Join(segmentInfo.CloneIDs, "、"),
				segmentInfo.ID+"-"+strings.Join(segmentInfo.CloneIDs[:min(MaxClone, len(segmentInfo.CloneIDs))], "、"),
			)
		}
	}
	fmt.Println("==END==")

	// 输出2-选择孔图
	fmt.Println("==输出2-选择孔图==")
	for _, jpID := range JPlist {
		segmentIDs := SegmentList[jpID]
		if len(segmentIDs) > 4 {
			log.Fatalf("片段超限[%s][%+v]", jpID, segmentIDs)
		}

		fmt.Printf("%s\t%s\t%s", jpID, jpID, strings.Join(ColName12, "\t"))
		fmt.Println()

		for j := range 4 {
			if j >= len(segmentIDs) {
				fmt.Printf("%s\t%c\n", jpID, 'A'+j*2)
				fmt.Printf("%s\t%c\n", jpID, 'A'+j*2+1)
				continue
			}
			segmentID := segmentIDs[j]
			segmentInfo := JPInfo[segmentID]
			// for k:=range 24{
			// 	row:='A'+j*2+k/12
			// }
			row := 'A' + j*2
			fmt.Printf("%s\t%c\n", jpID, row)
			for k := range MaxClone {
				cloneID := strconv.Itoa(k + 1)
				segmentInfo.FromPanel[cloneID] = fmt.Sprintf("%c%d", row, k+1)
				fmt.Printf("\t%s:%s-%s:%t", segmentInfo.FromPanel[cloneID], segmentID, cloneID, segmentInfo.CloneStatus[cloneID])
			}
			fmt.Printf("%s\t%c", jpID, 'A'+j*2+1)
			for k := range MaxClone {
				cloneID := strconv.Itoa(k + 1 + MaxClone)
				segmentInfo.FromPanel[cloneID] = fmt.Sprintf("%c%d", row, k+1)
				fmt.Printf("\t%s:%s-%s:%t", segmentInfo.FromPanel[cloneID], segmentID, cloneID, segmentInfo.CloneStatus[cloneID])
			}
		}
	}
	fmt.Println("==END==")

	// 输出2-输出孔图
	fmt.Println("==输出2-输出孔图==")
	index := 0
	panelID := ""
	for i, jpID := range JPlist {
		if i%2 == 0 {
			// new panel
			date := time.Now().Format("20050102")
			panelID = fmt.Sprintf("[%s]-SC%d", date, (i+1)/2)
			fmt.Printf("%s\t片段名称1\t片段名称2\t%s\t%s", panelID, panelID, strings.Join(ColName12, "\t"))
			index = 0
		}
		segmentIDs := SegmentList[jpID]
		for j := range segmentIDs {
			segmentID := segmentIDs[j]
			fmt.Printf("%s\t%s\t/\t%c", panelID, segmentID, 'A'+index)
			segmentInfo := JPInfo[segmentID]
			for k := range segmentInfo.CloneIDs {
				cloneID := segmentInfo.CloneIDs[k]
				fmt.Printf("\t%s-%s", segmentID, cloneID)
			}
			index++
		}
	}
	fmt.Println("==END==")

	// 测序YK
	fmt.Println("==测序YK==")
	for i := range JPlist {
		jpID := JPlist[i]
		segmentIDs := SegmentList[jpID]
		for j := range segmentIDs {
			segmentID := segmentIDs[j]
			segmentInfo := JPInfo[segmentID]
			for k := range segmentInfo.CloneIDs {
				cloneID := segmentInfo.CloneIDs[k]
				ID := fmt.Sprintf("%s-%s", segmentID, cloneID)
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

type SegmentInfo struct {
	ID             string
	JPID           string
	Length         int
	SequencePrimer string
	T7Primer       bool
	T7TermPrimer   bool
	OtherPrimers   []string
	Note2Product   string
	Laddar         string
	CloneIDs       []string
	// CloneIDs -> true
	CloneStatus map[string]bool
	// 送测克隆数
	SequenceCount int
	// Cell Name
	FromPanel map[string]string
}

func NewSegmentInfo(item map[string]string) *SegmentInfo {
	segmentInfo := &SegmentInfo{
		ID:             item["片段名称"],
		JPID:           item["JP-日期"],
		Length:         stringsUtil.Atoi(item["片段长度"]),
		SequencePrimer: item["测序引物"],
		Note2Product:   item["备注（to生产）"],
		Laddar:         item["Laddar"],
		CloneStatus:    make(map[string]bool),
		FromPanel:      make(map[string]string),
	}

	for primer := range strings.SplitSeq(segmentInfo.SequencePrimer, "、") {
		switch primer {
		case "T7":
			segmentInfo.T7Primer = true
		case "T7-Term":
			segmentInfo.T7TermPrimer = true
		default:
			segmentInfo.OtherPrimers = append(segmentInfo.OtherPrimers, primer)
		}

	}

	// 0..23
	for i := range 24 {
		cloneID := strconv.Itoa(i + 1)
		if item[cloneID] == "Y" {
			segmentInfo.CloneIDs = append(segmentInfo.CloneIDs, cloneID)
			segmentInfo.CloneStatus[cloneID] = true
		}
	}
	segmentInfo.SequenceCount = min(MaxClone, len(segmentInfo.CloneIDs))

	return segmentInfo
}

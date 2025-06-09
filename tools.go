package main

import (
	"log"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

func LoadInput(input string) (JPInfo map[string]*SegmentInfo, JPlist []string, SegmentList map[string][]string) {
	var (
		xlsx  = simpleUtil.HandleError(excelize.OpenFile(input))
		rows  = simpleUtil.HandleError(xlsx.GetRows("胶图判定(人工输入)"))
		title []string

		JPID string
		// JPID -> []string
	)
	JPInfo = make(map[string]*SegmentInfo)
	SegmentList = make(map[string][]string)
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
	return
}

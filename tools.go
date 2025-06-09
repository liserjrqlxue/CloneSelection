package main

import (
	"log"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

func LoadInput(excel, sheet string) (jpPanelMap map[string]*JPPanel, jpPanelList []string) {
	var (
		xlsx  = simpleUtil.HandleError(excelize.OpenFile(excel))
		rows  = simpleUtil.HandleError(xlsx.GetRows(sheet))
		title []string

		jpPanel *JPPanel
		// 检查 segmentID 是否重复
		segmentInfoMap = make(map[string]bool)
	)
	jpPanelMap = make(map[string]*JPPanel)
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
		if i%6 == 1 {
			if !ok {
				log.Fatalf("JP-日期:A%d empty", i+1)
			} else {
				jpPanel = &JPPanel{
					ID: jpID,
				}
				jpPanelMap[jpPanel.ID] = jpPanel
				jpPanelList = append(jpPanelList, jpPanel.ID)
			}
		} else {
			if jpPanel == nil || jpPanel.ID == "" {
				log.Fatalf("JP-日期:before A%d empty", i+1)
			}
			if !ok || jpID == "" {
				item["JP-日期"] = jpPanel.ID
			} else {
				log.Fatalf("JP-日期:A%d not empty[%+v]", i+1, item)
			}
		}
		segmentInfo := NewSegmentInfo(item)
		if segmentInfoMap[segmentInfo.ID] {
			log.Fatalf("片段重复:[%d:%s]", i+1, segmentInfo.ID)
		}
		segmentInfoMap[segmentInfo.ID] = true
		jpPanel.Segments = append(jpPanel.Segments, segmentInfo)
	}
	return
}

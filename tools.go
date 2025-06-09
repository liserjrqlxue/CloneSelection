package main

import (
	"log"
	"strconv"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

func LoadInput(excel, sheet string) (jpPanelMap map[string]*JPPanel, jpPanelList []string) {
	var (
		xlsx  = simpleUtil.HandleError(excelize.OpenFile(excel))
		rows  = simpleUtil.HandleError(xlsx.GetRows(sheet))
		title []string

		// 存储当前JPPanel
		jpPanelCurrent *JPPanel
		// 检查 segmentID 是否重复
		segmentInfoMap = make(map[string]bool)
	)
	jpPanelMap = make(map[string]*JPPanel)
	for i := range rows {
		// 读取 row[i] -> item
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
		// JP-日期 是 合并单元格，仅在第一行有值
		if i%MaxSegment == 1 {
			if !ok {
				log.Fatalf("JP-日期:A%d empty", i+1)
			} else {
				jpPanelCurrent = &JPPanel{
					ID: jpID,
					TY: isTY.MatchString(item["片段名称"]),
				}
				jpPanelMap[jpPanelCurrent.ID] = jpPanelCurrent
				jpPanelList = append(jpPanelList, jpPanelCurrent.ID)
			}
		} else {
			if jpPanelCurrent == nil || jpPanelCurrent.ID == "" {
				log.Fatalf("JP-日期:before A%d empty", i+1)
			}
			if jpID != "" {
				log.Fatalf("JP-日期:A%d not empty[%+v]", i+1, item)
			}
			if jpPanelCurrent.TY != isTY.MatchString(item["片段名称"]) {
				log.Fatalf("TY冲突[%s:%s]", jpPanelCurrent.ID, item["片段名称"])
			}
		}

		// 更新Gels
		if j := (i - 1) % MaxSegment; j < 4 {
			for k := range 25 {
				col := strconv.Itoa(k + 1)
				jpPanelCurrent.Gels[j][k] = item[col]
			}
		}

		segmentInfo := jpPanelCurrent.AddSegment(item)
		if segmentInfoMap[segmentInfo.ID] {
			log.Fatalf("片段重复:[%d:%s]", i+1, segmentInfo.ID)
		}
		segmentInfoMap[segmentInfo.ID] = true

	}
	return
}

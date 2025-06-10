package main

import (
	"log"
	"strconv"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

func LoadInput(excel, sheet string) (jps *JPs) {
	var (
		xlsx  = simpleUtil.HandleError(excelize.OpenFile(excel))
		rows  = simpleUtil.HandleError(xlsx.GetRows(sheet))
		title []string

		// 存储当前JPPanel
		current *JPPanel
		// 检查 segmentID 是否重复
		segmentInfoMap = make(map[string]bool)
	)
	jps = &JPs{
		Map: make(map[string]*JPPanel),
	}
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
		if i%MaxSegmentRow == 1 {
			if !ok {
				log.Fatalf("JP-日期:A%d empty", i+1)
			} else {
				current = &JPPanel{
					ID: jpID,
					TY: isTY.MatchString(item["片段名称"]),
				}
				simpleUtil.CheckErr(current.ParseID())
				jps.Map[current.ID] = current
				jps.List = append(jps.List, current)
			}
		} else {
			if current == nil || current.ID == "" {
				log.Fatalf("JP-日期:before A%d empty", i+1)
			}
			if jpID != "" {
				log.Fatalf("JP-日期:A%d not empty[%+v]", i+1, item)
			}
			if item["片段名称"] != "" && current.TY != isTY.MatchString(item["片段名称"]) {
				log.Fatalf("TY冲突[%s:%s]", current.ID, item["片段名称"])
			}
		}

		// 更新Gels
		if j := (i - 1) % MaxSegmentRow; j < 4 {
			for k := range 25 {
				col := strconv.Itoa(k + 1)
				current.Gels[j][k] = item[col]
			}
		}

		// skip
		if item["片段名称"] == "" {
			continue
		}

		segmentInfo := current.AddSegment(item)
		if segmentInfoMap[segmentInfo.ID] {
			log.Fatalf("片段重复:[%d:%s]", i+1, segmentInfo.ID)
		}
		segmentInfoMap[segmentInfo.ID] = true

		// date 相同
		date := jps.List[0].Date
		for _, jpPanel := range jps.List {
			if jpPanel.Date != date {
				log.Fatalf("Date不一致:[%s]vs[%s]", jps.List[0].ID, jpPanel.ID)
			}
		}

	}
	return
}

func InitToPanel(xlsx *excelize.File, sheet, panelID string, rowOffset int, bgStyleMap map[int]int) {
	var (
		cellName string

		startRow = 1 + rowOffset
		endRow   = startRow + PanelRow
	)

	// Table格式
	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(1, startRow),
			CoordinatesToCellName(PanelCol+4, endRow),
			bgStyleMap[-1],
		),
	)

	// 行标题
	cellName = CoordinatesToCellName(1, startRow)
	simpleUtil.CheckErr(
		xlsx.SetSheetRow(
			sheet, cellName,
			&[]any{panelID, "片段名称1", "片段名称2", panelID},
		),
	)
	cellName = CoordinatesToCellName(5, startRow)
	simpleUtil.CheckErr(
		xlsx.SetSheetRow(sheet, cellName, &PanelColTitle),
	)
	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(4, startRow),
			CoordinatesToCellName(16, startRow),
			bgStyleMap[3],
		),
	)

	// 列标题
	cellName = CoordinatesToCellName(4, startRow+1)
	simpleUtil.CheckErr(
		xlsx.SetSheetCol(sheet, cellName, &PanelRowTitle),
	)
	simpleUtil.CheckErr(
		xlsx.SetCellStyle(
			sheet,
			CoordinatesToCellName(4, startRow),
			CoordinatesToCellName(4, endRow),
			bgStyleMap[3],
		),
	)

	// 合并单元格
	simpleUtil.CheckErr(
		xlsx.MergeCell(
			sheet,
			CoordinatesToCellName(1, startRow),
			CoordinatesToCellName(1, endRow),
		),
	)

}

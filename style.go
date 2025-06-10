package main

import (
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

// style
var (
	bgColor1   = "#DDEBF7"
	bgColor2   = "#FCE4D6"
	bgColor3   = "#E2EFDA"
	bgGreen60  = "#C6E0B4"
	bgWhite15  = "#D9D9D9"
	tableStyle = excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			WrapText:   true,
			Horizontal: "center",
			Vertical:   "center",
		},
	}
	bgStyle1 = excelize.Style{
		Border:    tableStyle.Border,
		Alignment: tableStyle.Alignment,
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{bgColor1},
			Pattern: 1,
		},
	}
	bgStyle2 = excelize.Style{
		Border:    tableStyle.Border,
		Alignment: tableStyle.Alignment,
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{bgColor2},
			Pattern: 1,
		},
	}
	bgStyle3 = excelize.Style{
		Border:    tableStyle.Border,
		Alignment: tableStyle.Alignment,
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{bgColor3},
			Pattern: 1,
		},
	}
	bgStyleGreen60 = excelize.Style{
		Border:    tableStyle.Border,
		Alignment: tableStyle.Alignment,
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{bgGreen60},
			Pattern: 1,
		},
		Font: &excelize.Font{
			Bold: true,
		},
	}
	bgStyleWhite15 = excelize.Style{
		Border:    tableStyle.Border,
		Alignment: tableStyle.Alignment,
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{bgWhite15},
			Pattern: 1,
		},
		Font: &excelize.Font{
			Bold: true,
		},
	}
)

func NewStyle(xlsx *excelize.File, style *excelize.Style) (styleID int) {
	styleID = simpleUtil.HandleError(xlsx.NewStyle(style))
	return
}

func CreateStyles(xlsx *excelize.File) map[int]int {
	// styleID
	var (
		tableStyleID     = NewStyle(xlsx, &tableStyle)
		bgStyle1ID       = NewStyle(xlsx, &bgStyle1)
		bgStyle2ID       = NewStyle(xlsx, &bgStyle2)
		bgStyle3ID       = NewStyle(xlsx, &bgStyle3)
		bgStyleGreen60ID = NewStyle(xlsx, &bgStyleGreen60)
		bgStyleWhite15ID = NewStyle(xlsx, &bgStyleWhite15)
		bgStyleMap       = map[int]int{
			-1: tableStyleID,
			0:  bgStyle1ID,
			1:  bgStyle2ID,
			2:  bgStyle3ID,
			3:  bgStyleGreen60ID,
			4:  bgStyleWhite15ID,
		}
	)
	return bgStyleMap
}

func CoordinatesToCellName(col, row int) string {
	return simpleUtil.HandleError(
		excelize.CoordinatesToCellName(col, row),
	)
}
func SetCellStyle(xlsx *excelize.File, sheet string, hCol, hRow, vCol, vRow, styleID int) {
	var hCell = CoordinatesToCellName(hCol, hRow)
	var vCell = CoordinatesToCellName(vCol, vRow)
	simpleUtil.CheckErr(xlsx.SetCellStyle(sheet, hCell, vCell, styleID))
}

func MergeCell(xlsx *excelize.File, sheet string, topLeftCol, topLeftRow, bottomRightCol, bottomRightRow int) {
	simpleUtil.CheckErr(
		xlsx.MergeCell(
			sheet,
			CoordinatesToCellName(topLeftCol, topLeftRow),
			CoordinatesToCellName(bottomRightCol, bottomRightRow),
		),
	)
}

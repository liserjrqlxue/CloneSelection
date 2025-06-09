package main

import (
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/stringsUtil"
)

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

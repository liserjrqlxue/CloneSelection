package main

import (
	"log"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/stringsUtil"
)

type JPPanel struct {
	ID        string
	Index     int
	Date      string
	TY        bool
	Segments  []*Segment
	Gels      [4][25]string
	GelLayout string
}

func (jpPanel *JPPanel) Gels2Segments() {
	var (
		gels       = jpPanel.Gels
		jpCloneMax = MaxJPClone
		// 非ladder克隆序号
		index = 0
	)
	if jpPanel.TY {
		jpCloneMax = MaxJPCloneTY
	}

	// 校验GelLayout
	if gels[0][0] == "/" && gels[1][0] == "/" && gels[2][0] == "/" && gels[3][0] == "/" {
		jpPanel.GelLayout = "first ladder"
		if jpPanel.TY {
			jpPanel.GelLayout += " for TY"
		}
	}
	if gels[0][16] == "/" && gels[1][8] == "/" && gels[2][16] == "/" && gels[3][8] == "/" {
		jpPanel.GelLayout = "partition ladder"
	}
	if jpPanel.GelLayout == "" {
		log.Fatalf("Unknown Gels Layout for [%s]:%+v", jpPanel.ID, gels)
	}

	// 遍历Gels, 更新 Segment
	for j := range MaxGelRow {
		for k := range MaxGelCol {
			gel := gels[j][k]
			if gel != "/" {
				if gel == "Y" {
					segmentIndex := index / jpCloneMax
					segment := jpPanel.Segments[segmentIndex]
					cloneID := strconv.Itoa(index%jpCloneMax + 1)
					segment.CloneIDs = append(segment.CloneIDs, cloneID)
					clone := &Clone{
						ID:    cloneID,
						Index: index%jpCloneMax + 1,
					}
					segment.CloneStatus[cloneID] = clone
				}
				index++
			}
		}
	}

	// 更新 segment.SequenceCount
	for j := range jpPanel.Segments {
		segment := jpPanel.Segments[j]
		segment.SequenceCount = min(MaxCloneSelect, len(segment.CloneIDs))
	}
}

func (jpPanel *JPPanel) AddSegment(item map[string]string) *Segment {
	segment := &Segment{
		ID:             item["片段名称"],
		JPID:           jpPanel.ID,
		Length:         stringsUtil.Atoi(item["片段长度"]),
		SequencePrimer: item["测序引物"],
		Note2Product:   item["备注（to生产）"],
		CloneStatus:    make(map[string]*Clone),
		FromPanel:      make(map[string]string),
	}

	for primer := range strings.SplitSeq(segment.SequencePrimer, "、") {
		switch primer {
		case "T7":
			segment.T7Primer = true
		case "T7-Term":
			segment.T7TermPrimer = true
		default:
			segment.OtherPrimers = append(segment.OtherPrimers, primer)
		}
	}

	jpPanel.Segments = append(jpPanel.Segments, segment)

	return segment
}

type Segment struct {
	ID             string
	JPID           string
	Length         int
	SequencePrimer string
	T7Primer       bool
	T7TermPrimer   bool
	OtherPrimers   []string
	Note2Product   string
	CloneIDs       []string
	// CloneIDs -> true
	CloneStatus map[string]*Clone
	// 送测克隆数
	SequenceCount int
	// Cell Name
	FromPanel map[string]string
}

type Clone struct {
	Index    int
	ID       string
	FromCell string
	ToCell   string
}

package main

import (
	"flag"
	"fmt"
	"log"
	"log/slog"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/fmtUtil"
	"github.com/liserjrqlxue/goUtil/osUtil"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

// flag
var (
	input = flag.String(
		"i",
		"",
		"input excel",
	)
	output = flag.String(
		"o",
		"",
		"output prefix",
	)
)

var (
	GeneInfoTitle = []string{
		"序号",
		"基因内名",
		"基因名称",
		"序列",
		"酶切位点A",
		"酶切位点A.1",
		"酶切位点B",
		"酶切位点B.1",
		"载体",
	}

	CarrierListTitle = []string{
		"载体名称",
		"载体序列",
	}

	CreateSheet      = "构建完成"
	CreateSheetTitle = []string{
		"质粒名称",
		"序列",
		"备注",
	}
)

func main() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		log.Fatal("-i required")
	}
	if *output == "" {
		*output = strings.TrimSuffix(*input, "xlsx") + "构建完成"
	}

	// load Input
	xlsx := simpleUtil.HandleError(excelize.OpenFile(*input))
	carrierInfo := LoadCarrierList(xlsx, "载体清单", CarrierListTitle)
	geneInfos := LoadGeneInfo(xlsx, "Sheet1", GeneInfoTitle)

	// create Output
	var FA = osUtil.Create(*output + ".fasta")
	defer simpleUtil.DeferClose(FA)
	xlsx.NewSheet(CreateSheet)
	xlsx.SetSheetRow(CreateSheet, "A1", &CreateSheetTitle)
	var plasmids []*Plasmid

	// 循环处理
	for i := range geneInfos {
		data := geneInfos[i]
		plasmid := &Plasmid{
			Name: fmt.Sprintf("%s-%s-%s", data["基因内名"], data["基因名称"], data["载体"]),
		}
		plasmids = append(plasmids, plasmid)

		carrierSeq, ok := carrierInfo[data["载体"]]
		if !ok {
			plasmid.Name = "载体不存在"
			slog.Error("载体不存在", "序号", data["序号"], "载体", data["载体"])
			continue
		}
		plasmid.Update(data, carrierSeq)
	}

	// Write Output
	for i, p := range plasmids {
		if p.Note == "" {
			fmtUtil.Fprintf(FA, ">%s\n%s\n", p.Name, p.Sequence)
		}
		xlsx.SetSheetRow(CreateSheet, "A"+strconv.Itoa(i+2), &[]string{p.Name, p.Sequence, p.Note})
	}
	simpleUtil.CheckErr(xlsx.SaveAs(*output + ".xlsx"))
}

// 读取载体清单 name -> seq
func LoadCarrierList(xlsx *excelize.File, sheet string, title []string) map[string]string {
	carrierInfo := make(map[string]string)
	slice := simpleUtil.HandleError(xlsx.GetRows(sheet))
	for i := range slice {
		row := slice[i]
		if i == 0 { // 校验表头
			if len(row) < len(title) {
				log.Fatalf("title leak column:[%+v]", row)
			}
			for i := range title {
				if row[i] != title[i] {
					log.Fatalf("title1 error:%d[%s]vs[%s]", i+1, row[i], title[i])
				}
			}
			continue
		}
		carrierInfo[row[0]] = strings.ToUpper(row[1])
	}
	return carrierInfo
}

// 读取基因信息
func LoadGeneInfo(xlsx *excelize.File, sheet string, title []string) (dataArray []map[string]string) {
	slice := simpleUtil.HandleError(xlsx.GetRows(sheet))
	for i := range slice {
		row := slice[i]
		if i == 0 { // 校验表头
			if len(row) < len(title) {
				log.Fatalf("title leak column:[%+v]", row)
			}
			for i := range title {
				if row[i] != title[i] {
					log.Fatalf("title1 error:%d[%s]vs[%s]", i+1, row[i], title[i])
				}
			}
			continue
		}
		data := make(map[string]string)
		for j := range title {
			data[title[j]] = row[j]
		}
		dataArray = append(dataArray, data)
	}
	return
}

type Plasmid struct {
	Name     string
	Sequence string
	Note     string
}

func (plasmid *Plasmid) Update(data map[string]string, carrierSeq string) {
	var e1start, e1end, e2start, e2end int

	e1seq := strings.ToUpper(data["酶切位点A.1"])
	e1start = FindIndex(carrierSeq, e1seq)
	e1end = e1start + len(e1seq)

	e2seq := strings.ToUpper(data["酶切位点B.1"])
	e2start = FindIndex(carrierSeq, e2seq)
	e2end = e2start + len(e2seq)

	if e1start > 0 && e2start > 0 {
		if e2start > e1end {
			data["E1起点"] = strconv.Itoa(e1start)
			data["E1终点"] = strconv.Itoa(e1end)
			data["E2起点"] = strconv.Itoa(e2start)
			data["E2终点"] = strconv.Itoa(e2end)
			data["图谱"] = carrierSeq[:e1end] + data["序列"] + carrierSeq[e2start:]
			plasmid.Sequence = data["图谱"]
		} else {
			slog.Info("酶切位置错误", "序号", data["序号"], "e1start", e1start, "e1end", e1end, "e2start", e2start, "e2end", e2end)
			plasmid.Note = "酶切位置错误"
		}
	} else {
		slog.Info("酶切位置找不到", "序号", data["序号"], "e1start", e1start, "e1end", e1end, "e2start", e2start, "e2end", e2end)
		plasmid.Note = "酶切位置找不到"
	}
}

// -1 for not found
func FindIndex(query, target string) int {
	l := len(target)
	for i := range len(query) - l {
		str := query[i : i+l]
		if str == target {
			return i
		}
	}
	return -1
}

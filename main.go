package main

import (
	"flag"
	"log"
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
	title1 = []string{
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

	title2 = []string{
		"载体名称",
		"载体序列",
	}
)

func main() {
	flag.Parse()
	if *input == "" || *output == "" {
		flag.Usage()
		log.Fatal("-i required")
	}

	xlsx := simpleUtil.HandleError(excelize.OpenFile(*input))
	// name -> seq
	carrierInfo := make(map[string]string)
	slice := simpleUtil.HandleError(xlsx.GetRows("载体清单"))
	for i := range slice {
		row := slice[i]
		if i == 0 {
			if len(row) < len(title1) {
				log.Fatalf("title leak column:[%+v]", row)
			}
			for i := range title2 {
				if row[i] != title2[i] {
					log.Fatalf("title1 error:%d[%s]vs[%s]", i+1, row[i], title2[i])
				}
			}
			continue
		}
		data := make(map[string]string)
		for j := range title1 {
			data[title1[j]] = row[j]
		}
		carrierInfo[data["载体名称"]] = data["载体序列"]
	}

	out := osUtil.Create(*output + ".fasta")
	defer simpleUtil.DeferClose(out)
	slice = simpleUtil.HandleError(xlsx.GetRows("Sheet1"))
	for i := range slice {
		row := slice[i]
		if i == 0 {
			if len(row) < len(title1) {
				log.Fatalf("title leak column:[%+v]", row)
			}
			for i := range title1 {
				if row[i] != title1[i] {
					log.Fatalf("title1 error:%d[%s]vs[%s]", i+1, row[i], title1[i])
				}
			}
			continue
		}

		data := make(map[string]string)
		for j := range title1 {
			data[title1[j]] = row[j]
		}
		carrierSeq, ok := carrierInfo[data["载体"]]
		if !ok {
			log.Fatalf("载体不存在:[%s]", data["载体"])
		}
		var e1start, e1end, e2start, e2end int
		e1seq := strings.ToUpper(data["酶切位点A.1"])
		e1len := len(e1seq)
		for j := range len(carrierSeq) - e1len {
			tseq := carrierSeq[j : j+e1len]
			if tseq == e1seq {
				e1start = j
				e1end = j + e1len
				break
			}
		}
		e2seq := strings.ToUpper(data["酶切位点B.1"])
		e2len := len(e2seq)
		for j := range len(carrierSeq) - e2len {
			tseq := carrierSeq[j : j+e2len]
			if tseq == e2seq {
				e2start = j
				e2end = j + e2len
				break
			}
		}
		if e1len > 0 && e2len > 0 {
			if e2start > e1len {
				data["E1起点"] = strconv.Itoa(e1start)
				data["E1终点"] = strconv.Itoa(e1end)
				data["E2起点"] = strconv.Itoa(e2start)
				data["E2终点"] = strconv.Itoa(e2end)
				data["图谱"] = carrierSeq[:e1end] + data["序列"] + carrierSeq[e2start:]
			} else {
				log.Fatalf("酶切位置错误:[%d,%d],[%d,%d]", e1start, e1end, e2start, e2end)
			}
		} else {
			log.Fatalf("酶切位置找不到:[%d,%d],[%d,%d]", e1start, e1end, e2start, e2end)
		}

		fmtUtil.Fprintf(out, ">%s-%s-%s\n%s\n", data["基因内名"], data["基因名称"], data["载体"], data["图谱"])
	}
}

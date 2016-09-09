package main

import (
	"flag"
	"log"
	"regexp"

	"github.com/tealeg/xlsx"
	"github.com/yargevad/xlsxutil"
)

var inFlag = flag.String("in", "", "input file")
var outFlag = flag.String("out", "", "output file")
var colnumFlag = flag.Int("col", 0, "0-based column index")

func main() {
	flag.Parse()

	if *inFlag == "" {
		log.Fatal("--in is required")
	} else if *outFlag == "" {
		log.Fatal("--out is required")
	}

	xin, err := xlsx.OpenFile(*inFlag)
	if err != nil {
		log.Fatal(err)
	}

	opts := &xlsxutil.SheetSplitOpts{
		InFile:   xin,
		ColIndex: *colnumFlag,
		IgnorePatterns: []*regexp.Regexp{
			regexp.MustCompile(`^(?i:table)\s+\d+`),
			regexp.MustCompile(`^(?i:county)\b`),
		},
	}
	xout, err := xlsxutil.SheetSplit(opts)
	if err != nil {
		log.Fatal(err)
	}

	err = xout.Save(*outFlag)
	if err != nil {
		log.Fatal(err)
	}
}

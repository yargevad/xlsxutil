package xlsxutil

import (
	"regexp"
	"strings"

	"github.com/pkg/errors"
	"github.com/tealeg/xlsx"
)

type SheetSplitOpts struct {
	InFile         *xlsx.File
	ColIndex       int
	ShortRowErr    bool
	TrimSheetNames bool
	IgnorePatterns []*regexp.Regexp
}

// SheetSplit reads the provided XLSX document and returns a new document.
// It creates one sheet for each unique value in the specified column.
func SheetSplit(opts *SheetSplitOpts) (*xlsx.File, error) {
	// sanity check input
	if opts.InFile == nil {
		return nil, errors.New("input file is required")
	} else if len(opts.InFile.Sheets) > 1 {
		return nil, errors.New("input files may not have more than one sheet")
	}

	// the output file we're writing to
	outFile := xlsx.NewFile()
	// sheet lookup by name
	sheets := map[string]*xlsx.Sheet{}

	// iterate over input sheets and rows
	for _, sheet := range opts.InFile.Sheets {
	NEXTROW:
		for rowIdx, row := range sheet.Rows {
			// error/ignore short rows based on config
			if row.Cells == nil || len(row.Cells) < (opts.ColIndex+1) {
				if opts.ShortRowErr {
					return nil, errors.Errorf("not enough cells in row %d", rowIdx)
				}
				continue
			}

			// get the cell value
			cellStr, err := row.Cells[opts.ColIndex].String()
			if err != nil {
				return nil, errors.Wrap(err, "coercing cell to string")
			}

			// suppress rows based on config
			if len(opts.IgnorePatterns) > 0 {
				for _, pat := range opts.IgnorePatterns {
					if pat.Find([]byte(cellStr)) != nil {
						continue NEXTROW
					}
				}
			}

			// and maybe trim it, based on config
			if opts.TrimSheetNames {
				cellStr = strings.TrimSpace(cellStr)
			}

			// create/return sheet where the current row belongs
			var outSheet *xlsx.Sheet
			var ok bool
			if outSheet, ok = sheets[cellStr]; !ok {
				sheets[cellStr], err = outFile.AddSheet(cellStr)
				if err != nil {
					return nil, errors.Wrap(err, "adding sheet to output file")
				}
				outSheet = sheets[cellStr]
			}

			// add a row to the selected sheet
			newRow := outSheet.AddRow()
			// write the input row to the row we just added
			stringRow, err := StringifyRow(row)
			if err != nil {
				return nil, err
			}
			rc := newRow.WriteSlice(&stringRow, -1)
			if rc < 0 {
				return nil, errors.Errorf("bad cell write (%d)", rc)
			}
		}
	}
	return outFile, nil
}

func StringifyRow(r *xlsx.Row) ([]string, error) {
	ncells := len(r.Cells)
	row := make([]string, ncells, ncells)
	for i, cell := range r.Cells {
		str, err := cell.String()
		if err != nil {
			return nil, errors.Wrap(err, "coercing cell to string")
		}
		row[i] = str
	}
	return row, nil
}

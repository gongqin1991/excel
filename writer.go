package excel

import (
	"fmt"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
	"strings"
)

var (
	colAlias = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
		"L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
)

type writer struct {
	dirty    bool
	filename string
	sheet    string
	w        *excelize.File
	err      error

	sheetIdx int //工作表索引
}

func NewWriter(path, sheet string) *writer {
	fs := excelize.NewFile()
	var err error
	if oldName := fs.GetSheetName(0); oldName != sheet {
		err = fs.SetSheetName(oldName, sheet)
	}
	wr := &writer{
		filename: path,
		w:        fs,
		sheet:    sheet,
		err:      err,
	}
	return wr
}

func OpenWriter(w *writer, sheet string) *writer {
	fs := w.w
	var (
		err        error
		sheetIndex int
	)
	sheetIndex, err = fs.GetSheetIndex(sheet)
	if err == nil && sheetIndex == -1 {
		sheetIndex, err = fs.NewSheet(sheet)
	}
	wr := &writer{
		filename: w.filename,
		w:        fs,
		sheet:    sheet,
		err:      err,
		sheetIdx: sheetIndex,
	}
	return wr
}

func (w *writer) WriteColumns(columns []string, rows int) {
	checkRow(w, rows)
	for i, n := 0, len(columns); i < n && w.err == nil; i++ {
		if !w.dirty {
			w.dirty = true
		}
		cell := fmt.Sprintf("%s%d", w.indexToCol(i), rows)
		w.err = w.w.SetCellStr(w.sheet, cell, columns[i])
	}
}

func checkRow(w *writer, row int) {
	if w.err != nil {
		return
	}
	if row <= 0 {
		w.err = errors.New("row index start at 1")
	}
}
func checkCol(w *writer, col int) {
	if w.err != nil {
		return
	}
	if col <= 0 {
		w.err = errors.New("column index start at 1")
	}
}

func (w *writer) WriteHeader(columns []string) {
	w.WriteColumns(columns, 1)
}

func (w *writer) MergeRows(cols, start, size int) {
	checkCol(w, cols)
	checkRow(w, start)
	if w.err != nil {
		return
	}
	colstr := w.indexToCol(cols - 1)
	s := fmt.Sprintf("%s%d", colstr, start)
	t := fmt.Sprintf("%s%d", colstr, start+size-1)
	w.err = w.w.MergeCell(w.sheet, s, t)
}

func (w *writer) indexToCol(cols int) string {
	C := cols
	cba := make([]string, 0)
	for N := len(colAlias); C >= 0; {
		a := C % N
		cba = append(cba, colAlias[a])
		C = C/N - 1
	}
	for i, j := 0, len(cba)-1; i < j; {
		cba[i], cba[j] = cba[j], cba[i]
		i++
		j--
	}
	return strings.Join(cba, "")
}

func (w *writer) SaveTo() {
	if w.err != nil || !w.dirty {
		return
	}
	w.err = w.w.SaveAs(w.filename)
}

func (w *writer) Err() error {
	return w.err
}

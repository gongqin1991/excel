package excel

import (
	"fmt"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
	"strings"
)

const (
	INVALID = iota - 1
	HEAD
	ROW
)

var (
	CA = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
		"L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
)

type writer struct {
	w          *excelize.File
	dirty      bool
	filename   string
	sheetName  string
	sheetIndex int
	hierarchy  int
	err        error
}

func NewWriter(path, sheet string) *writer {
	fs := excelize.NewFile()
	var (
		err    error
		offset int
	)
	if oldName := fs.GetSheetName(offset); sheet != "" && oldName != sheet {
		err = fs.SetSheetName(oldName, sheet)
	} else if sheet == "" {
		offset = INVALID
	}
	wr := &writer{
		filename:   path,
		w:          fs,
		sheetIndex: offset,
		sheetName:  sheet,
		err:        err,
	}
	return wr
}

func NewWriter2(path string) *writer {
	return NewWriter(path, "")
}

func checkSheetName(name string) {
	if name == "" || strings.TrimSpace(name) == "" {
		panic("invalid sheet name")
	}
}

func OpenWriter(w *writer, sheet string) *writer {
	fs := w.w
	var (
		err    error
		offset int
	)
	checkSheetName(sheet)
	if w.hierarchy == 0 && w.sheetIndex == INVALID {
		offset = HEAD
		if oldName := fs.GetSheetName(offset); oldName != sheet {
			err = fs.SetSheetName(oldName, sheet)
		}

		w.err = err
		w.sheetIndex = offset
		w.sheetName = sheet
		return w
	}

	offset, err = fs.GetSheetIndex(sheet)
	if err == nil && offset == INVALID {
		offset, err = fs.NewSheet(sheet)
	}

	wr := &writer{
		filename:   w.filename,
		dirty:      w.dirty,
		w:          fs,
		sheetIndex: offset,
		sheetName:  sheet,
		hierarchy:  w.hierarchy + 1,
		err:        err,
	}
	return wr
}

// Deprecated:被 WriteRow 替换
func (w *writer) WriteColumns(columns []string, rows int) {
	checkRow(w, rows)
	for i, n := 0, len(columns); i < n && w.err == nil; i++ {
		if !w.dirty {
			w.dirty = true
		}
		cell := fmt.Sprintf("%s%d", w.indexToCol(i), rows)
		w.err = w.w.SetCellStr(w.sheetName, cell, columns[i])
	}
}

func (w *writer) WriteRow(columns []string, rows int) {
	w.WriteColumns(columns, rows)
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
	w.WriteColumns(columns, ROW)
}

func (w *writer) MergeRows(cols, start, size int) {
	checkCol(w, cols)
	checkRow(w, start)
	if w.err != nil {
		return
	}
	colStr := w.indexToCol(cols - 1)
	s := fmt.Sprintf("%s%d", colStr, start)
	t := fmt.Sprintf("%s%d", colStr, start+size-1)
	w.err = w.w.MergeCell(w.sheetName, s, t)
}

func (w *writer) indexToCol(cols int) string {
	C := cols
	cba := make([]string, 0)
	for N := len(CA); C >= 0; {
		a := C % N
		cba = append(cba, CA[a])
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
	err := w.err
	if err != nil || !w.dirty {
		return
	}
	wr := w.w
	err = wr.SaveAs(w.filename)
	if err == nil {
		err = wr.Close()
	}
	w.err = err
}

func (w *writer) Err() error {
	return w.err
}

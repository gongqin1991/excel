package excel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
)

type reader struct {
	io.ReadCloser
	real  *excelize.File
	sheet int
	err   error
}

type cols struct {
	rd *reader
	*excelize.Rows
}

func (c cols) HasNext() bool {
	if c.rd.err != nil {
		return false
	}
	return c.Rows.Next()
}

func (c cols) Get() []string {
	if c.rd.err != nil {
		return nil
	}
	values, err := c.Columns()
	if err != nil {
		c.rd.err = err
	}
	return values
}

func OpenReader(path, sheet string) *reader {
	fs, err := os.Open(path)
	return NewReader(fs, err, sheet)
}

func NewReader(fd io.ReadCloser, err error, sheet string) (rd *reader) {
	rd = &reader{}
	defer func() {
		rd.err = err
	}()
	if err != nil {
		return
	}
	rd.ReadCloser = fd
	f, err := excelize.OpenReader(fd)
	if err != nil {
		return
	}
	rd.real = f
	index, err := f.GetSheetIndex(sheet)
	if err != nil {
		return
	}
	if index == -1 {
		err = errors.New("no sheet found")
		return
	}
	rd.sheet = index
	return
}

func (r *reader) GetRows() (col cols) {
	//col = &cols{}
	col.rd = r
	if r.err != nil {
		return
	}
	sheetName := r.real.GetSheetName(r.sheet)
	rows, err := r.real.Rows(sheetName)
	if err != nil {
		r.err = err
		return
	}
	col.Rows = rows
	return
}

func (r *reader) Err() error {
	return r.err
}

package excel

import (
	"bytes"
	"io"
)

type cacheBuffer struct {
	io.ReadCloser
	buff *bytes.Buffer
}

type buffer struct {
	io.Reader
}

func NewBufferReader(reader io.ReadCloser) *cacheBuffer {
	buf := cacheBuffer{
		ReadCloser: reader,
		buff:       bytes.NewBuffer(nil),
	}
	return &buf
}

func (r *cacheBuffer) Read(p []byte) (n int, err error) {
	n, err = r.ReadCloser.Read(p)
	if n > 0 {
		r.buff.Write(p[:n])
	}
	return n, err
}

func (r *cacheBuffer) NewReader() io.ReadCloser {
	return &buffer{Reader: r.buff}
}

func (r *buffer) Close() error {
	return nil
}

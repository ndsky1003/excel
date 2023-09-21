package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func head(index uint) (ret string) {
	for index != 0 {
		index -= 1
		shang := index / 26
		yu := index % 26
		ret = string([]byte{65 + byte(yu)}) + ret
		index = shang
	}
	return
}

type page struct {
	f     *excelize.File
	rowid int
	name  string
}

func NewPage(name ...string) (*page, error) {
	tmpname := ""
	if len(name) != 0 {
		tmpname = name[0]
	} else {
		tmpname = "Sheet1"
	}
	f := excelize.NewFile()
	_, err := f.NewSheet(tmpname)
	if err != nil {
		return nil, err
	}
	return &page{
		f:    f,
		name: tmpname,
	}, nil
}

func (this *page) Close() error {
	if this.f != nil {
		return this.f.Close()
	}
	return nil
}

func (this *page) Save(path string) error {
	if !strings.HasSuffix(path, "xlsx") {
		path = fmt.Sprintf("%s.xlsx", path)
	}
	return this.f.SaveAs(path)
}

func (this *page) SetTitle(titles ...any) {
	this.PushRow(titles...)
}

func (this *page) PushRow(datas ...any) {
	this.rowid++
	for i, v := range datas {
		_ = this.f.SetCellValue(this.name, head(uint(i)+1)+fmt.Sprintf("%d", this.rowid), v)
	}
}

package main

import (
	"flag"
	"io/ioutil"
	"fmt"
	"path/filepath"

	"github.com/tealeg/xlsx"
)

type xlsxFile struct {
	Filename string
	Workbook *xlsx.File
}

func getXlsxFiles(dir string) []xlsxFile {
	files, err := ioutil.ReadDir(dir)
	if err != nil {
		return nil
	}

	var slice []xlsxFile
	for _, file := range files {
		if  filepath.Ext(file.Name()) != ".xlsx"{
			continue
		}
		wb, err :=	xlsx.OpenFile(filepath.Join(dir, file.Name()))
		if err != nil {
			fmt.Println(err)
			continue
		}

		slice = append(slice, xlsxFile{Filename: file.Name(), Workbook:wb })
	}
	return slice
}

func main() {
	// var files string
	var out string
	var dir string
	flag.StringVar(&dir, "d", "", "指定需要合并的工作薄目录")
	// flag.StringVar(&files, "i", "", "指定需要合并的工作薄，格式：[wb1.xlsx, wb2.xlsx]")
	flag.StringVar(&out, "o", "out.xlsx", "指定合并后文件的命名，格式：out.xlsx")
	flag.Parse()

	
	wb:= xlsx.NewFile()

	xfiles := getXlsxFiles(dir)
	for _, xf := range xfiles {
		fmt.Println("发现xlsx文件", xf.Filename)
		for _, sheet := range xf.Workbook.Sheets {
			wb.AppendSheet(*sheet, sheet.Name)
			fmt.Println("增加sheet", sheet.Name)
		}
	}

	if len(wb.Sheets) == 0 {
		return
	}

	wb.Save(out)
	fmt.Println("合并出文件：", out)
}
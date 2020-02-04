package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"path/filepath"
	"strconv"

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
		if filepath.Ext(file.Name()) != ".xlsx" {
			continue
		}
		wb, err := xlsx.OpenFile(filepath.Join(dir, file.Name()))
		if err != nil {
			fmt.Println(err)
			continue
		}

		slice = append(slice, xlsxFile{Filename: file.Name(), Workbook: wb})
	}
	return slice
}

func main() {
	var out string
	var dir string
	var single bool
	flag.StringVar(&dir, "d", "", "指定需要合并的工作薄目录")
	flag.StringVar(&out, "o", "out.xlsx", "指定合并后文件的命名，格式：out.xlsx")
	flag.BoolVar(&single, "s", false, "使用文件名，命名单sheet工作薄")
	flag.Parse()

	wb := xlsx.NewFile()

	unique := make(map[string]int)
	xfiles := getXlsxFiles(dir)
	for _, xf := range xfiles {
		fmt.Println("发现xlsx文件", xf.Filename)
		var count int
		var sheetName string
		for _, sheet := range xf.Workbook.Sheets {
			if single && len(xf.Workbook.Sheets) == 1 {
				_, f := filepath.Split(xf.Filename)
				sheetName = f[:len(f)-len(filepath.Ext(xf.Filename))]
			} else {
				count = unique[sheet.Name]
				sheetName = sheet.Name
				if count > 0 {
					sheetName += " (" + strconv.Itoa(count) + ")"
				}
			}

			wb.AppendSheet(*sheet, sheetName)
			unique[sheet.Name]++
			if count > 0 {
				fmt.Println("增加sheet", sheet.Name, " => ", sheetName)
			} else {
				fmt.Println("增加sheet", sheet.Name)
			}
		}
	}

	if len(wb.Sheets) == 0 {
		return
	}

	wb.Save(out)
	fmt.Println("合并出文件：", out)
}

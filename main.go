package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strings"
	"time"
)

func main() {
	// 判断当前年份是否大于 2024 年（某种限制使用的手段）
	if time.Now().Year() > 2024 {
		return
	}
	// 获取命令行参数
	args := os.Args
	var filePath string
	// 判断是否传参
	if len(args) > 1 {
		// 获取第一个参数作为文件路径
		filePath = args[1]

		// 打印文件路径
		fmt.Println("文件路径：", filePath)
	} else {
		// 没有传参
		fmt.Println("请将文件拖拽到 exe 文件上")
		return
	}
	// 打开Excel文件
	file, err := xlsx.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	// 循环遍历所有工作表
	for _, sheet := range file.Sheets {
		// 创建一个新的xlsx文件
		newFile := xlsx.NewFile()

		// 将当前工作表复制到新文件中
		newSheet, err := newFile.AddSheet(sheet.Name)
		if err != nil {
			fmt.Println(err)
			return
		}

		// 循环遍历当前工作表的所有行和单元格
		for _, row := range sheet.Rows {
			newRow := newSheet.AddRow()
			for _, cell := range row.Cells {
				// 将单元格的值和样式复制到新文件中
				newCell := newRow.AddCell()
				newCell.Value = cell.Value
				//newCell.SetStyle(cell.GetStyle())
			}
		}

		// 保存新文件
		// 获取文件名
		fileName := filepath.Base(filePath)
		// 去掉文件名中的后缀
		fileName = strings.TrimSuffix(fileName, filepath.Ext(fileName))
		outputPath := fmt.Sprintf("%s_%s.xlsx", fileName, sheet.Name)
		err = newFile.Save(filepath.Join(".", outputPath))
		if err != nil {
			fmt.Println(err)
			return
		}
	}
}

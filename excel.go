package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
	"strings"
)

var FOUND = "匹配"

// 需要比较的字段
var mainColumns []string
var targetColumns []string

// 比较结果输出列
var mainResult string
var targetResult string

func main() {
	var mainPath string
	var tartPath string
	var mainCStr string
	var targetCStr string

	flag.StringVar(&mainPath, "m", "/Users/song/Downloads/银行日记账-测试-主表.xlsx", "主表目录")
	flag.StringVar(&tartPath, "t", "/Users/song/Downloads/银行日记账-U8-副表.xlsx", "副表目录")
	flag.StringVar(&mainCStr, "mc", "C,E", "主表需要比较的列")
	flag.StringVar(&targetCStr, "tc", "H,J", "副表需要比较的列")
	flag.StringVar(&mainResult, "mr", "O", "主表结果列")
	flag.StringVar(&targetResult, "tr", "O", "副表表结果列")
	flag.Parse()

	mainFile, err := excelize.OpenFile(mainPath)
	targetFile, tae := excelize.OpenFile(tartPath)
	if err != nil {
		fmt.Println(mainPath)
		return
	}
	if tae != nil {
		panic("未找到副表")
		return
	}

	mainColumns = strings.Split(mainCStr, ",")
	targetColumns = strings.Split(targetCStr, ",")

	// 获取 Sheet1 上所有单元格
	mainSheet := mainFile.GetSheetName(0)
	rows, err := mainFile.GetRows(mainSheet)
	for rowNum := 3; rowNum < len(rows)+1; rowNum++ {
		compareVal := findCellValues(mainFile, mainSheet, rowNum, mainColumns)
		compare := findPayerCashOut(targetFile, compareVal)
		if compare {
			markFund(mainFile, mainSheet, mainResult, rowNum)
		}
		if rowNum%10 == 0 {
			fmt.Println("已完成行数：", rowNum)
		}
	}

	// 保存处理结果
	me := mainFile.Save()
	te := targetFile.Save()
	if me != nil || te != nil {
		panic("处理结果更新失败")
	}
	fmt.Println("比较成功")
}

func findPayerCashOut(targetFile *excelize.File, compareVal []string) bool {
	targetSheet := targetFile.GetSheetName(0)
	rows, _ := targetFile.GetRows(targetSheet)
	for num := 2; num < len(rows)+1; num++ {
		flag := targetResult + strconv.Itoa(num)
		flagV, _ := targetFile.GetCellValue(targetSheet, flag)
		if FOUND == flagV {
			continue
		}

		targetVal := findCellValues(targetFile, targetSheet, num, targetColumns)
		if reflect.DeepEqual(compareVal, targetVal) {
			markFund(targetFile, targetSheet, targetResult, num)
			return true
		}
	}
	return false
}

func findCellValues(file *excelize.File, sheet string, rowNum int, columns []string) []string {
	compareVal := make([]string, 0)
	for _, column := range columns {
		cellVal, _ := file.GetCellValue(sheet, column+strconv.Itoa(rowNum))
		compareVal = append(compareVal, cellVal)
	}
	return compareVal
}

func markFund(file *excelize.File, sheet string, column string, rowNum int) {
	cell := column + strconv.Itoa(rowNum)
	e := file.SetCellValue(sheet, cell, FOUND)
	if e != nil {
		panic(e)
	}
}

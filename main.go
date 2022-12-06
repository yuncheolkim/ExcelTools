package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"io/fs"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"unicode"
)

const tint = "int"
const tLong = "long"
const tFloat = "float"
const tJson = "json"

type ExcelData struct {
	Name   string
	Sheets []ExcelSheetData
}

type ExcelSheetData struct {
	Name string
	Rows []ExcelRowData
}

type ExcelRowData struct {
	DataMap map[string]interface{}
}

type FieldType struct {
	Field string
	Type  string
}

var JsonPretty = true

func ReadXlsx(filename string) ExcelData {
	xlFile, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}
	baseName := strings.TrimSuffix(filepath.Base(filename), filepath.Ext(filename))
	excelData := ExcelData{Name: baseName}

	for _, sheet := range xlFile.Sheets {
		excelData.Sheets = append(excelData.Sheets, makeSheet(sheet))
	}

	return excelData
}

func makeSheet(sheet *xlsx.Sheet) ExcelSheetData {
	s := ExcelSheetData{}
	s.Name = sheet.Name

	field, e := sheet.Row(1)
	if e != nil {
		panic(errors.New("title error"))
	}

	fieldType, e := sheet.Row(2)
	if e != nil {
		panic(errors.New("field type error"))
	}

	m := make(map[int]FieldType)

	maxCol := sheet.MaxCol

	for i := 0; i < maxCol; i++ {
		ft := fieldType.GetCell(i)
		c := field.GetCell(i)
		v := FieldType{
			Field: c.Value,
			Type:  ft.Value,
		}
		m[i] = v
	}

	// rows
	for i := 3; i < sheet.MaxRow; i++ {
		r, e := sheet.Row(i)
		row := makeRow(i, r, e, m)
		if len(row.DataMap) == 0 {
			continue
		}
		s.Rows = append(s.Rows, row)
	}

	return s
}

func makeRow(rowNum int, row *xlsx.Row, e error, m map[int]FieldType) ExcelRowData {
	if e != nil {
		panic("row parse error")
	}

	rd := ExcelRowData{
		DataMap: make(map[string]interface{}),
	}
	colCount := len(m)
	if row.GetCell(0).String() == "#" {
		fmt.Println("Skip row:", row.GetCell(0))
		return ExcelRowData{nil}
	}

	for i := 0; i < colCount; i++ {
		var cell interface{}
		var e error
		getCell := row.GetCell(i)
		if getCell.String() == "" {
			continue
		}
		switch m[i].Type {
		case tint:
			cell, e = getCell.Int()
		case tLong:
			cell, e = getCell.Int64()
		case tFloat:
			cell, e = getCell.Float()
		case tJson:
			e = json.Unmarshal([]byte(getCell.String()), &cell)
		default:
			cell = getCell.String()
		}

		if e != nil {
			println(row.Sheet.Name, "行", rowNum, "列", i)
			panic(e)
		}
		rd.DataMap[m[i].Field] = cell
	}
	delete(rd.DataMap, "")

	return rd
}

func ParseToFile(data interface{}, path string) {
	var bytes []byte
	var e error
	if JsonPretty {
		bytes, e = json.MarshalIndent(data, "", "    ")
	} else {
		bytes, e = json.Marshal(data)

	}
	if e != nil {
		panic(e)
	}

	f, e := os.OpenFile(path, os.O_WRONLY|os.O_CREATE, fs.ModePerm)
	if e != nil {
		panic(e)
	}
	//e = ioutil.WriteFile(path, bytes, fs.ModePerm)

	f.Write(bytes)
	f.Close()

}

var pattern2 = regexp.MustCompile("[A-Za-z]+")

func main() {

	inDir := "/t1"
	outDir := "/t2"
	t := 1
	println(outDir)

	var pretty bool

	flag.BoolVar(&pretty, "format", true, "是否格式化json")
	flag.StringVar(&inDir, "inDir", "", "excel所在的目录")
	flag.StringVar(&outDir, "outDir", "", "json输出目录")
	flag.IntVar(&t, "t", 1, "工具类型，1:excel导出，2：中文字符整理")
	flag.Parse()
	JsonPretty = pretty

	if t == 1 {

		file, _ := ioutil.ReadDir(inDir)
		for _, v := range file {
			if v.IsDir() {
				continue
			}

			name := v.Name()
			if filepath.Ext(name) != ".xlsx" || strings.HasPrefix(name, "~$") || strings.HasPrefix(name, "test") || strings.HasPrefix(name, "temp") {
				continue
			}

			println("读取数据:", name)

			result := ReadXlsx(filepath.Clean(filepath.Join(inDir, name)))
			for _, v := range result.Sheets {
				if strings.HasPrefix(v.Name, "test") {
					continue
				}
				name := result.Name + "_" + v.Name

				m := make(map[int]map[string]interface{})
				println("parse", name)

				for key, row := range v.Rows {
					println("key:", key)
					if row.DataMap["id"] == nil {
						break
					}
					switch row.DataMap["id"].(type) {
					case int:
						i := row.DataMap["id"].(int)
						m[i] = row.DataMap
					case string:
						fmt.Println(row.DataMap)
					}
				}
				ParseToFile(m, filepath.Clean(filepath.Join(outDir, name+".json")))
			}
		}
	} else if t == 2 { // 整理字符
		bytes, e := ioutil.ReadFile(inDir)

		if e != nil {
			fmt.Println(e)
			return
		}

		context := string(bytes)

		cn := make(map[string]bool)
		var count = 0
		var allChar = ""

		for _, v := range context {
			vstr := string(v)
			r := v
			if unicode.IsSpace(r) {
				continue
			}

			if unicode.IsNumber(r) {

				fmt.Println("发现：IsNumber", vstr)
				continue
			}
			if unicode.IsPunct(r) {
				fmt.Println("发现：IsPunct", vstr)
				continue
			}

			if pattern2.MatchString(string(v)) {
				fmt.Println("发现英文", vstr)
				continue
			}

			if unicode.Is(unicode.Han, r) {
				// 中文
				if _, ok := cn[vstr]; !ok {
					count++
					cn[vstr] = true
					allChar += vstr
				}
			} else {
				fmt.Println("其他", v)
			}

		}
		ioutil.WriteFile(outDir, []byte(allChar), 0777)

		fmt.Println("count:", count)

	}
}

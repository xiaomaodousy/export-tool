package xlsx2

import (
	"bytes"
	"errors"
	"export-tool/utils"
	"fmt"
	"github.com/astaxie/beego/logs"
	"github.com/tealeg/xlsx/v3"
	"io/ioutil"
	"os"
	"path/filepath"
	"runtime/debug"
	"strconv"
	"strings"
	"time"
)

type JsonTool struct {
	BaseTool
}

func NewJsonTool(option BaseTool) *JsonTool {
	return &JsonTool{BaseTool{
		InputDir:           option.InputDir,
		OutputDir:          option.OutputDir,
		Force:              option.Force,
		ErrLog:             option.ErrLog,
		outputForServerDir: filepath.Join(option.OutputDir, "server"),
		outputForClientDir: filepath.Join(option.OutputDir, "client"),
		checkCSVNames:      map[string]*checkCSVData{},
		Filenames:          option.Filenames,
		AllFile:            option.AllFile,
	}}
}

func (j *JsonTool) Exec() (err error) {
	if j.ErrLog != "" {
		j.errLogFile, err = os.OpenFile(j.ErrLog, os.O_CREATE|os.O_TRUNC, 0666)
		if err != nil {
			return
		}
	}
	defer func() {
		e := recover()
		if e != nil {
			_ = os.RemoveAll(j.outputForServerDir)
			debug.PrintStack()
			switch er := e.(type) {
			case string:
				err = errors.New(er)
			case error:
				err = er
			default:
				panic(e)
			}
		}
		_ = j.errLogFile.Close()
	}()
	files, err := ioutil.ReadDir(j.InputDir)
	if err != nil {
		logs.Error(err, 111)
		return
	}
	err = j.checkOrCreateSubDir(j.outputForServerDir)
	if err != nil {
		logs.Error(err, j.outputForServerDir, 111)

		return
	}
	for _, file := range files {
		if strings.HasPrefix(file.Name(), "~") || strings.HasPrefix(file.Name(), ".") {
			continue
		}
		if _, ok := j.Filenames[file.Name()]; ok || j.AllFile {
			j.execXlsx(filepath.Join(j.InputDir, file.Name()))
		}
	}
	total := len(j.checkCSVNames) + len(files) + len(j.afterCheckSaveFuncs)
	counter := len(files)
	if j.CallProgress != nil {
		j.CallProgress(counter, total)
	}
	return
}

func (j *JsonTool) execXlsx(filename string) {
	start := time.Now()
	defer func() {
		fmt.Println("execXlsx usage for:", filename, time.Now().Sub(start))
	}()
	var err error
	xlsxFileLd, err = xlsx.OpenFile(filename)
	if err != nil {
		j.writeLineLog(fmt.Sprintf("file: %s open err:%s", filename, err.Error()))
		return
	}
	basename := strings.Split(filepath.Base(filename), ".")[0]
	for _, sheet := range xlsxFileLd.Sheets {
		j.execSheet2Json(basename, sheet)
	}
}

func (j *JsonTool) execSheet2Json(name string, sheet *xlsx.Sheet) {
	start := time.Now()
	defer func() {
		fmt.Println("execSheet2Csv usage for:", name, sheet.Name, time.Now().Sub(start))
	}()
	// 解析标题头和输出类型
	var forServer, forClient bool
	var lineLimit int
	r, err := sheet.Row(0)
	if err != nil {
		j.writeSheetError(name, sheet.Name, err)
		return
	}
	meta := r.GetCell(0).Value
	metas := strings.Split(meta, "#")
	csvName := metas[0]
	if len(metas) == 1 {
		forServer = true
		forClient = true
	}
	if len(metas) > 1 && metas[1] == "1" {
		forServer = true
	}
	if len(metas) > 2 && metas[2] == "1" {
		forClient = true
	}
	if len(metas) > 3 {
		lineLimit = utils.StrToInt[int](metas[3])
	}

	var buildCheckError = func(msg any) string {
		return fmt.Sprintf("CHECK SERVER CSV FILE: <<%s>> - %s ERROR: %v ", name, sheet.Name, msg)
	}

	if lineLimit > 0 && lineLimit < 3 {
		panic(buildCheckError("行数限制不能小于3, 前3行是必须的头; 不限制请设置为0"))
	}

	//stripQuot := csvName == "Language" || csvName == "BadWords"
	//tsExport := csvName == "Language"
	serverFile := filepath.Join(j.outputForServerDir, csvName+".json")
	clientFile := filepath.Join(j.outputForClientDir, csvName+".json")
	if forServer && !j.Force && j.pathExist(serverFile) {
		forServer = false
		j.writeLineLog(fmt.Sprintf("file: %s exist skip", serverFile), FlagWarn)
	}
	if forClient && !j.Force && j.pathExist(clientFile) {
		forClient = false
		j.writeLineLog(fmt.Sprintf("file: %s exist skip", clientFile), FlagWarn)
	}
	// 如果全部都不处理 则直接过滤掉了
	if !forServer && !forClient {
		return
	}
	var datasForServer bytes.Buffer
	var datasForClient bytes.Buffer
	defer func() {
		datasForServer.Reset()
		datasForClient.Reset()
	}()
	var index int
	// 不导出的列
	var unexportCellIndex = map[int]struct{}{}
	var titleMapping = map[int]string{}
	var typesForClient = map[int]string{}

	var allData []map[string]any
	//var tsText []string
	var _ = sheet.ForEachRow(func(r *xlsx.Row) error {
		defer func() {
			index++
		}()
		// 行数限制
		if lineLimit > 0 && index+1 > lineLimit {
			return nil
		}
		var contents = map[string]any{}
		var cellIndex int

		_ = r.ForEachCell(func(c *xlsx.Cell) (err error) {
			defer func() {
				cellIndex++
			}()
			value := c.Value
			if index == 1 {
				titleMapping[cellIndex] = c.Value
			}
			if index == 0 {
				value = strings.ReplaceAll(value, ",", " ")
				value = strings.ReplaceAll(value, "\r\n", " ")
				value = strings.ReplaceAll(value, "\n", " ")
			} else {
				value = strings.ReplaceAll(value, ",", "，")
				value = strings.ReplaceAll(value, "\r\n", "\n")
				value = strings.ReplaceAll(value, "\n", "\\n")
			}

			// 排除列 index == 0 是字段名称列
			if index == 0 && (strings.HasPrefix(value, "UNEXPORT_") || strings.TrimSpace(c.Value) == "") {
				unexportCellIndex[cellIndex] = struct{}{}
				return
			}

			if _, ok := unexportCellIndex[cellIndex]; ok {
				return
			}

			if index == 1 && strings.TrimSpace(c.Value) == "" {
				// 如果title是空 则不导出
				unexportCellIndex[cellIndex] = struct{}{}
				return
			}

			if index == 2 { // index==2 是类型列
				if strings.Contains(value, "#") {
					values := strings.Split(value, "#")
					typesForClient[cellIndex] = values[0]
				} else {
					typesForClient[cellIndex] = value

				}
			}

			if index > 2 {
				switch typesForClient[cellIndex] {
				case "int":
					contents[titleMapping[cellIndex]] = ValueFormat[int](c.Value)
				case "string":
					contents[titleMapping[cellIndex]] = ValueFormat[string](c.Value)
				default:
					contents[titleMapping[cellIndex]] = c.Value
				}
			}

			return
		})
		if index < 3 {
			return nil
		}

		allData = append(allData, contents)

		return nil
	})

	vs, err := JSONEncodeWithNoEscape(allData)

	err = ioutil.WriteFile(serverFile, vs, 0666)
	if err != nil {
		j.writeSheetError(name, sheet.Name, err)
		return
	}
}

func ValueFormat[T any](value string) T {
	var result any

	switch any(*new(T)).(type) {
	case int:
		result, _ = strconv.Atoi(value)
	case string:
		result = value
	default:
		return *new(T)
	}
	return result.(T)
}

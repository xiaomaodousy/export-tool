package xlsx2

import (
	"bytes"
	"encoding/json"
	"errors"
	"export-tool/utils"
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"io/ioutil"
	"os"
	"path/filepath"
	"runtime/debug"
	"strings"
	"time"
)

const (
	FlagWarn  = "warn"
	FlagError = "error"
)

type Xlsx2CsvTool struct {
	BaseTool
}

type checkCSVData struct {
	csvName string
	fname   string
	sname   string
}

type CallProgressFunc = func(counter int, total int)

func NewXlsx2CsvTool(option BaseTool) *Xlsx2CsvTool {
	return &Xlsx2CsvTool{
		BaseTool{
			InputDir:           option.InputDir,
			OutputDir:          option.OutputDir,
			Force:              option.Force,
			ErrLog:             option.ErrLog,
			outputForServerDir: filepath.Join(option.OutputDir, "server"),
			outputForClientDir: filepath.Join(option.OutputDir, "client"),
			checkCSVNames:      map[string]*checkCSVData{},
			Filenames:          option.Filenames,
			AllFile:            option.AllFile,
		},
	}
}

func (x *Xlsx2CsvTool) SetCallProgressFunc(f CallProgressFunc) {
	x.CallProgress = f
}

type CheckFunc = func()

func (x *Xlsx2CsvTool) Exec() (err error) {
	// 崩溃的情况才返回error  否则log打印即可
	if x.ErrLog != "" {
		x.errLogFile, err = os.OpenFile(x.ErrLog, os.O_CREATE|os.O_TRUNC, 0666)
		if err != nil {
			return
		}
	}
	defer func() {
		e := recover()
		if e != nil {
			_ = os.RemoveAll(x.outputForServerDir)
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
		_ = x.errLogFile.Close()
	}()
	files, err := ioutil.ReadDir(x.InputDir)
	if err != nil {
		return
	}
	err = x.checkOrCreateSubDir(x.outputForServerDir)
	if err != nil {
		return
	}
	err = x.checkOrCreateSubDir(x.outputForClientDir)
	if err != nil {
		return
	}
	for _, file := range files {
		if strings.HasPrefix(file.Name(), "~") || strings.HasPrefix(file.Name(), ".") {
			continue
		}
		if _, ok := x.Filenames[file.Name()]; ok || x.AllFile {
			x.execXlsx(filepath.Join(x.InputDir, file.Name()))
		}
	}
	total := len(x.checkCSVNames) + len(files) + len(x.afterCheckSaveFuncs)
	counter := len(files)
	if x.CallProgress != nil {
		x.CallProgress(counter, total)
	}

	for _, data := range x.checkCSVNames {
		func() {
			defer func() {
				err := recover()
				if err != nil {
					panic(fmt.Sprintf("CHECK SERVER CSV FILE: <<%s>> - %s ERROR: %v ", data.fname, data.sname, err))
				}
			}()
		}()
		counter++
		if x.CallProgress != nil {
			x.CallProgress(counter, total)
		}
	}

	for _, f := range x.afterCheckSaveFuncs {
		f()
		counter++
		if x.CallProgress != nil {
			x.CallProgress(counter, total)
		}
	}
	return
}

var xlsxFileLd *xlsx.File

// 处理一个xlsx
func (x *Xlsx2CsvTool) execXlsx(filename string) {
	start := time.Now()
	defer func() {
		fmt.Println("execXlsx usage for:", filename, time.Now().Sub(start))
	}()
	var err error
	xlsxFileLd, err = xlsx.OpenFile(filename)
	if err != nil {
		x.writeLineLog(fmt.Sprintf("file: %s open err:%s", filename, err.Error()))
		return
	}
	basename := strings.Split(filepath.Base(filename), ".")[0]
	for _, sheet := range xlsxFileLd.Sheets {
		x.execSheet2Csv(basename, sheet)
	}
}

// 将一个sheet转为csv
func (x *Xlsx2CsvTool) execSheet2Csv(name string, sheet *xlsx.Sheet) {
	start := time.Now()
	defer func() {
		fmt.Println("execSheet2Csv usage for:", name, sheet.Name, time.Now().Sub(start))
	}()
	// 解析标题头和输出类型
	var forServer, forClient bool
	var lineLimit int
	r, err := sheet.Row(0)
	if err != nil {
		x.writeSheetError(name, sheet.Name, err)
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

	stripQuot := csvName == "Language" || csvName == "BadWords"
	tsExport := csvName == "Language"
	serverCsvFile := filepath.Join(x.outputForServerDir, csvName+".csv")
	clientCsvFile := filepath.Join(x.outputForClientDir, csvName+".csv")
	if forServer && !x.Force && x.pathExist(serverCsvFile) {
		forServer = false
		x.writeLineLog(fmt.Sprintf("file: %s exist skip", serverCsvFile), FlagWarn)
	}
	if forClient && !x.Force && x.pathExist(clientCsvFile) {
		forClient = false
		x.writeLineLog(fmt.Sprintf("file: %s exist skip", clientCsvFile), FlagWarn)
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
	var titleMapping map[int]string

	var tsText []string
	var _ = sheet.ForEachRow(func(r *xlsx.Row) error {
		defer func() {
			index++
		}()
		// 行数限制
		if lineLimit > 0 && index+1 > lineLimit {
			return nil
		}
		var typesForClient []string
		var contents []string
		var cellIndex int
		if index == 1 {
			titleMapping = map[int]string{}
		}
		var dataMapping map[string]string
		var rowKey string
		if tsExport {
			dataMapping = map[string]string{}
		}
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

			if index > 2 && cellIndex == 0 {
				rowKey = c.Value
			}
			if index > 2 && tsExport && dataMapping != nil && cellIndex > 0 {
				dataMapping[titleMapping[cellIndex]] = c.Value
			}
			// 导出的列才进行导出
			// 客户端和服务器类型分离
			if index == 2 { // index==2 是类型列
				if strings.Contains(value, "#") {
					values := strings.Split(value, "#")
					typesForClient = append(typesForClient, values[1])
					// 通用数据使用#前的数据
					value = values[0]
				} else {
					typesForClient = append(typesForClient, value)
				}
			}
			if csvName == "BadWords" {
				value = strings.ReplaceAll(value, `"`, ``)
			}
			if stripQuot && (strings.Contains(value, "\\n") || strings.Contains(value, ",") || strings.Contains(value, "，")) {
				// 如果带有换行符
				// 将"做成双"号 否则前后加"会有问题
				value = strings.ReplaceAll(value, `"`, `""`)
				value = `"` + value + `"`
			}

			contents = append(contents, value)
			return
		})
		datasForServer.WriteString(strings.Join(contents, ",") + "\n")
		if index > 0 {
			if len(typesForClient) > 0 {
				// 如果使用client type
				contents = typesForClient
			}
			datasForClient.WriteString(strings.Join(contents, ",") + "\n")
		}
		if tsExport && index > 2 {
			vs, err := JSONEncodeWithNoEscape(dataMapping)
			if err != nil {
				return err
			}
			tsText = append(tsText, fmt.Sprintf(`	"%s":%s,`, rowKey, strings.Trim(strings.ReplaceAll(string(vs), "\\\\n", "\\n"), "\n")))
		}
		return nil
	})

	var writeToFile = func(fname string, content []byte) {
		err = ioutil.WriteFile(fname, content, 0666)
		if err != nil {
			x.writeSheetError(name, sheet.Name, err)
			return
		}
	}
	if forServer {
		writeToFile(serverCsvFile, datasForServer.Bytes())
		f, err := os.OpenFile(serverCsvFile, os.O_RDONLY, 0666)
		if err != nil {
			fmt.Println("OPEN serverCsvFile error: ", serverCsvFile)
		} else {
			defer func() {
				_ = f.Close()
			}()
			// 先加载所有的表
			x.checkCSVNames[csvName] = &checkCSVData{
				csvName: csvName,
				fname:   name,
				sname:   sheet.Name,
			}
		}

	}
	if forClient {
		var content = datasForClient.Bytes()
		//copy(datasForClient.Bytes(), content)
		x.afterCheckSaveFuncs = append(x.afterCheckSaveFuncs, func() {
			writeToFile(clientCsvFile, content)
		})
	}

	//if tsExport && len(tsText) > 0 {
	//	content := fmt.Sprintf(tsTemplate, strings.Join(tsText, "\n"))
	//	fname := filepath.Join(x.OutputDir, "CXTranslationText.ts")
	//	x.afterCheckSaveFuncs = append(x.afterCheckSaveFuncs, func() {
	//		err = ioutil.WriteFile(fname, []byte(content), 0666)
	//		if err != nil {
	//			fmt.Println("导出CXTranslationText.ts失败")
	//			return
	//		}
	//		fmt.Println("CXTranslationText.ts文件已经导出: ", fname)
	//	})
	//}
}

func JSONEncodeWithNoEscape(data any) ([]byte, error) {
	body := bytes.NewBuffer([]byte{})
	encoder := json.NewEncoder(body)
	encoder.SetEscapeHTML(false) // 设置不转义HTML标记
	encoder.SetIndent("", "")
	// 将Person结构体转换成JSON字符串并输出
	err := encoder.Encode(data)
	if err != nil {
		return nil, err
	}
	return body.Bytes(), nil
}

package xlsx2

import (
	"fmt"
	"os"
)

const (
	OutputTypeCsv  = "csv"
	OutputTypeJson = "json"
)

type Tool interface {
	Exec() (err error)
	SetCallProgressFunc(f CallProgressFunc)
}

type BaseTool struct {
	InputDir            string // 导入的xlsx目录
	OutputDir           string // 导出的csv目录 // 将自动创建 server和client的子目录作为服务器和客户端使用
	OutputType          string // 导出的数据类型
	Force               bool   // 强制覆盖  如果文档存在
	ErrLog              string // 错误文档输出文件 不设置默认为当前目录
	errLogFile          *os.File
	outputForServerDir  string
	outputForClientDir  string
	afterCheckSaveFuncs []func() // 检查之后的其他文件保存逻辑 延后处理
	checkCSVNames       map[string]*checkCSVData
	CallProgress        CallProgressFunc
	Filenames           map[string]struct{}
	AllFile             bool
}

func (b *BaseTool) Exec() (err error) {
	//TODO implement me
	panic("implement me")
}

func (b *BaseTool) SetCallProgressFunc(f CallProgressFunc) {
	b.CallProgress = f
}

func NewTool(toolOption BaseTool) Tool {
	fmt.Println(toolOption.OutputType)
	switch toolOption.OutputType {
	case OutputTypeCsv:
		return NewXlsx2CsvTool(toolOption)
	case OutputTypeJson:
		return NewJsonTool(toolOption)
	default:

	}
	return nil
}

func (b *BaseTool) checkOrCreateSubDir(dirName string) error {
	if !b.pathExist(dirName) {
		return os.Mkdir(dirName, 0755)
	}
	return nil
}

func (b *BaseTool) pathExist(path string) bool {
	if _, err := os.Stat(path); err != nil {
		return !os.IsNotExist(err)
	}
	return true
}

func (b *BaseTool) writeLineLog(content string, flags ...string) {
	var flag = FlagError
	if len(flags) > 0 {
		flag = flags[0]
	}
	_, err := b.errLogFile.WriteString(fmt.Sprintf("[%s] %s\n", flag, content))
	if err != nil {
		panic(err)
	}
}

func (b *BaseTool) writeErrLog(err error) {
	b.writeLineLog(err.Error())
}

func (b *BaseTool) writeSheetError(filename, sheetName string, err error) {
	b.writeLineLog(fmt.Sprintf("file: %s, sheet: %s, err: %s", filename, sheetName, err.Error()))
}

package main

import (
	"encoding/json"
	"errors"
	"export-tool/xlsx2"
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/flopp/go-findfont"
	"io"
	"io/fs"
	"io/ioutil"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

var application fyne.App
var window fyne.Window

type ToolData struct {
	initFontError    error
	setting          ToolDataSetting
	serverRegion     string
	exportFileAll    bool
	exportFileChoose map[string]struct{}
}

type ToolDataSetting struct {
	InputDir   string
	OutputDir  string
	ClientDir  string
	ServerDir  string
	OutPutType string
}

const (
	AppConfigJson = "./luckluTool.json"
)

const (
	TextInputExcelDir  = "choose excel dir: "
	TextOutputDir      = "choose output dir: "
	TextClientDir      = "choose client dir(xd-client): "
	TextServerDir      = "choose server dir(xd): "
	TextOutPutType     = "choose output type(default csv): "
	TextOutPutTypeCsv  = "csv"
	TextOutPutTypeJson = "json"

	TextBtnSave          = "Save Config"
	TextLabelSetting     = "Setting"
	TextLabelSettingTips = "(excel dir must be like /配置表/数据表 which has excel files, ouput dir should be 导表工具, client dir must be xd-hclient)"
	TextLabelExport      = "export"
	TextLabelStart       = "Start"
	TextLabelSync        = "Sync CSV Server"
	TextLabelSyncClient  = "Sync Client CSV&ts"

	TextErrNotDir         = "not a dir"
	TextErrNoClientDir    = "client dir must be xd-client"
	TextErrNotSetSetting  = "not set setting!!"
	TextErrTsFileNotExist = "CXTranslationText.ts is not exist"
)

var chineseTextMapping = map[string]string{
	TextInputExcelDir: "请选择数据表目录: ",
	TextOutputDir:     "请选择输出目录: ",
	TextOutPutType:    "请选择输出方式(默认csv):",
	TextClientDir:     "请选择客户端目录(xd-client): ",
	TextServerDir:     "请选择服务端目录(xd): ",

	TextErrNotDir:         "当前的路径不是一个目录",
	TextErrNoClientDir:    "客户端目录必须是xd-client",
	TextBtnSave:           "保存配置",
	TextLabelSetting:      "配置",
	TextLabelExport:       "导表工具",
	TextLabelStart:        "开始",
	TextLabelSync:         "服务器配表同步",
	TextErrNotSetSetting:  "未设置目录, 不能运行!!",
	TextLabelSyncClient:   "同步客户端CSV和语言ts",
	TextErrTsFileNotExist: "导出目录下不存在CXTranslationText.ts, 请重新导出",
	TextLabelSettingTips:  "(数据表请选择到 客户端目录/配置表/数据表 有excel的目录, 输出目录请设置到导表工具目录即可, 客户端目录必须到xd-hclient的目录)",
}

var data = &ToolData{exportFileChoose: map[string]struct{}{}}

func getLabelText(key string) string {
	if text, ok := chineseTextMapping[key]; ok && data.initFontError == nil {
		return text
	}
	return key
}

func InitFont() error {
	for _, path := range findfont.List() {
		if strings.Contains(path, "msyh.ttf") || strings.Contains(path, "simhei.ttf") || strings.Contains(path, "simsun.ttc") || strings.Contains(path, "simkai.ttf") {
			return os.Setenv("FYNE_FONT", path)
		}
	}
	return errors.New("Not Found chinese Font!")
}

var (
	UIUpdateLabel = widget.NewLabel("check update...")
	UIUpdateBtn   = widget.NewButton("Update", func() {
		updateFilePath, _ := filepath.Abs("./update.exe")
		fmt.Println("执行: ", updateFilePath)
		cmd := exec.Command(updateFilePath)
		err := cmd.Start()
		if err != nil {
			ShowErrorMessage(err)
			return
		}
		// 关闭当前程序
		//os.Exit(0)
	})
)

func main() {
	data.initFontError = InitFont()
	application = app.New()
	window = application.NewWindow("工具箱LuckLu Tool Box")

	checkLoadConfig()
	window.Resize(fyne.NewSize(960, 600))
	window.SetContent(mainView())
	window.ShowAndRun()
}

func checkLoadConfig() {
	content, err := ioutil.ReadFile(AppConfigJson)
	if err != nil {
		//ShowErrorMessage(err)
		return
	}
	err = json.Unmarshal(content, &data.setting)
	if err != nil {
		//ShowErrorMessage(err)
		return
	}
}

func createExportView() *fyne.Container {
	var ui []fyne.CanvasObject
	if data.initFontError != nil {
		warning := widget.NewLabel("can't init chinese font, please try using admin to run this app!")
		warning.TextStyle.Bold = true
		ui = append(ui, warning)
	}
	ui = append(ui, createSettingUI(), createExportUI())
	return container.NewVBox(ui...)
}

//func createPBToolView() *fyne.Container {
//	label := widget.NewLabel("PB tools")
//	exportBtn := widget.NewButton(getLabelText(TextLabelStart), func() {
//		if data.setting.ServerDir == "" {
//			ShowErrorMessage(errors.New("server dir is empty"))
//			return
//		}
//		parser := pbtool.NewParse(filepath.Join(data.setting.ServerDir, "scripts/protocol/CProto.proto"))
//		err := parser.DoParse()
//		if err != nil {
//			ShowErrorMessage(err)
//			return
//		}
//		err = parser.ToServer(filepath.Join(data.setting.ServerDir, "internal/pkg/protocol"))
//		if err != nil {
//			ShowErrorMessage(err)
//			return
//		}
//		err = parser.ToClient(filepath.Join(data.setting.ClientDir, "assets/scripts/framework/net/protobuf"))
//		if err != nil {
//			ShowErrorMessage(err)
//			return
//		}
//		showOKMsg("OK")
//	})
//
//	return container.NewVBox(label, exportBtn)
//}

func mainView() *fyne.Container {
	tabs := container.NewAppTabs(container.NewTabItem("导表工具", createExportView()))
	return container.New(layout.NewBorderLayout(nil, nil, tabs, nil), tabs)
}

func checkDirPath(path string) error {
	if f, err := os.Stat(path); err != nil {
		return err
	} else if !f.IsDir() {
		return errors.New(getLabelText(TextErrNotDir))
	}
	return nil
}

func checkFileExist(path string) bool {
	if _, err := os.Stat(path); err != nil {
		return false
	}
	return true
}

func createUpdateUI() fyne.CanvasObject {
	UIUpdateBtn.Disable()
	return container.NewVBox(UIUpdateLabel, UIUpdateBtn)
}

func createSettingUI() fyne.CanvasObject {
	var setOK = func(label *widget.Label, path string) {
		label.SetText("√ " + path)
	}

	label := widget.NewLabel(getLabelText(TextLabelSetting))

	labelInfo := widget.NewLabel(getLabelText(TextLabelSettingTips))

	c1Info := widget.NewLabel("")
	if data.setting.InputDir != "" {
		setOK(c1Info, data.setting.InputDir)
	}
	c1 := container.NewHBox(
		widget.NewLabel(getLabelText(TextInputExcelDir)),
		widget.NewButton("Select Folder", func() {
			showFolderSelector(func(path string) {
				if err := checkDirPath(path); err != nil {
					c1Info.SetText("× " + err.Error())
				} else {
					data.setting.InputDir = path
					setOK(c1Info, data.setting.InputDir)
				}
			})
		}),
		c1Info,
	)

	c2Info := widget.NewLabel("")
	if data.setting.OutputDir != "" {
		setOK(c2Info, data.setting.OutputDir)
	}

	c2 := container.NewHBox(
		widget.NewLabel(getLabelText(TextOutputDir)),
		widget.NewButton("Select Folder", func() {
			showFolderSelector(func(path string) {
				if err := checkDirPath(path); err != nil {
					c2Info.SetText("× " + err.Error())
				} else {
					data.setting.OutputDir = path
					setOK(c2Info, data.setting.OutputDir)
				}
			})
		}),
		c2Info,
	)

	outPutTypeRadio := widget.NewRadioGroup([]string{TextOutPutTypeCsv, TextOutPutTypeJson}, func(s string) {
		data.setting.OutPutType = s
	})
	outPutTypeRadio.Horizontal = true // 横向排列
	outPutTypeRadio.Required = true
	if data.setting.OutPutType != "" {
		outPutTypeRadio.SetSelected(data.setting.OutPutType)
	} else {
		outPutTypeRadio.SetSelected(TextOutPutTypeCsv)
	}

	c3 := container.NewHBox(widget.NewLabel(getLabelText(TextOutPutType)), outPutTypeRadio)

	saveBtn := widget.NewButton(getLabelText(TextBtnSave), func() {
		content, err := json.Marshal(data.setting)
		if err != nil {
			ShowErrorMessage(err)
			return
		}
		err = ioutil.WriteFile(AppConfigJson, content, 0666)
		if err != nil {
			ShowErrorMessage(err)
			return
		}
		showOKMsg("OK")
	})
	return container.NewVBox(label, labelInfo, c1, c2, c3, saveBtn)
}

func checkSetting() bool {
	if data.setting.InputDir == "" || data.setting.OutputDir == "" {
		ShowErrorMessage(errors.New(getLabelText(TextErrNotSetSetting)))
		return false
	}
	return true
}

func createExportUI() fyne.CanvasObject {
	label := widget.NewLabel(getLabelText(TextLabelExport))
	chooseFileLabel := widget.NewLabel("选择到处文件")
	chooseAllFileCheck := widget.NewCheck("全选", func(b bool) {
		data.exportFileAll = b
	})
	var excelFileList []string
	chooseFileChecksBox := container.New(layout.NewGridLayout(10))

	readFileNameBtn := widget.NewButton("读取", func() {
		chooseFileChecksBox.RemoveAll()
		excelFileList = readUploadExcelFileNames()
		for _, s := range excelFileList {
			name := s
			chooseFileChecksBox.Add(widget.NewCheck(s, func(b bool) {
				if b {
					data.exportFileChoose[name] = struct{}{}
				} else {
					delete(data.exportFileChoose, name)
				}
			}))
		}
	})

	chooseFileContainer := container.New(layout.NewGridWrapLayout(fyne.NewSize(200, 30)), chooseFileLabel, chooseAllFileCheck, readFileNameBtn)

	startBtn := widget.NewButton(getLabelText(TextLabelStart), nil)
	var progress float64
	progressData := binding.BindFloat(&progress)
	progressBar := widget.NewProgressBarWithData(progressData)

	startBtn.OnTapped = func() {
		if !checkSetting() {
			return
		}
		startBtn.SetText("Running...")
		fmt.Println(data.exportFileAll, data.exportFileChoose)
		startBtn.Disable()
		progressBar.SetValue(0)

		options := xlsx2.BaseTool{
			InputDir:   data.setting.InputDir,
			OutputDir:  data.setting.OutputDir,
			OutputType: data.setting.OutPutType,
			Force:      true,
			ErrLog:     "",
			Filenames:  data.exportFileChoose,
			AllFile:    data.exportFileAll,
		}
		tool := xlsx2.NewTool(options)
		tool.SetCallProgressFunc(func(counter int, total int) {
			if total > 0 {
				progressBar.SetValue(float64(counter) / float64(total))
				startBtn.SetText(fmt.Sprintf("Progress:  %d / %d", counter, total))
			}
		})
		if err := tool.Exec(); err != nil {
			ShowErrorMessage(err)
		} else {
			showOKMsg("OK")
		}
		startBtn.SetText(getLabelText(TextLabelStart))
		startBtn.Enable()
	}
	return container.NewVBox(label, chooseFileContainer, container.New(layout.NewGridWrapLayout(fyne.NewSize(1500, 200)), chooseFileChecksBox), startBtn, progressBar)

}

func copyDirFileWithNoSubDir(src, dest string, filter func(info fs.FileInfo) bool) error {
	files, err := ioutil.ReadDir(src)
	if err != nil {
		return err
	}

	for _, file := range files {
		if !file.IsDir() && !filter(file) {
			// 构建源文件和目标文件路径
			sourceFile := filepath.Join(src, file.Name())
			destFile := filepath.Join(dest, file.Name())

			// 拷贝文件
			err := copyFile(sourceFile, destFile)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func copyFile(src, dest string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return err
	}
	defer sourceFile.Close()

	destFile, err := os.Create(dest)
	if err != nil {
		return err
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	if err != nil {
		return err
	}

	return nil
}

func showFolderSelector(setCallback func(path string)) {
	dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
		if err == nil && uri != nil {
			if err != nil {
				ShowErrorMessage(err)
				return
			}
			setCallback(uri.Path())
		}
	}, window)
}

func ShowErrorMessage(err error) {
	dialog.ShowError(err, window)
}

func showOKMsg(msg string) {
	dialog.ShowInformation("OK", "=================^ "+msg+" ^=================", window)
}

func readUploadExcelFileNames() []string {
	var names []string
	fs, _ := ioutil.ReadDir(data.setting.InputDir)
	for _, f := range fs {
		if f.IsDir() {
			continue
		}
		if strings.HasSuffix(f.Name(), ".xlsx") &&
			(!strings.HasPrefix(f.Name(), ".") && !strings.HasPrefix(f.Name(), "~")) {
			names = append(names, f.Name())
		}
	}
	return names
}

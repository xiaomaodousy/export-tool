package main

import (
	"bufio"
	"fmt"
	"github.com/ttacon/chalk"
	"os"
	"xd/cmd/tools/gui/core"
)

var _ = core.VERSION

func printErr(err error) {
	fmt.Println(chalk.Red, err.Error(), chalk.Reset)
}

func main() {
	core.InitVersion()
	fmt.Println(chalk.Yellow, "START...")
	fmt.Println(chalk.Yellow, "正在检查更新版本......", chalk.Reset)

	has, version, err := core.CheckNewUpdate()
	if err != nil {
		printErr(err)
		wait()
		return
	}
	if !has {
		fmt.Println(chalk.Yellow, "目前版本已经是最新了.", chalk.Reset)
		return
	}
	fmt.Println(chalk.Yellow, fmt.Sprintf("版本准备更新: %s => %s, 在主程序退出后, 点击回车进行更新...", core.VERSION, version), chalk.Reset)
	wait()
	fmt.Println(chalk.Yellow, "开始更新.... main.exe ...", chalk.Reset)
	err = core.DoUpdate()
	if err != nil {
		printErr(err)
		wait()
		return
	} else {
		fmt.Println(chalk.Green, "更新完成! 点击回车关闭窗口...", chalk.Reset)
	}
	wait()
}

func wait() {
	reader := bufio.NewReader(os.Stdin)
	_, _ = reader.ReadString('\n')
}

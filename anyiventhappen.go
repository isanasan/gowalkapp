package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/lxn/walk"
	"github.com/tealeg/xlsx"
)

// ComboBox 変更時イベント処理
func (mw *MyMainWindow) comboChanged() {
	var list []string

	excel, err1 := xlsx.FileToSlice("Table.xlsx")

	if err1 != nil {
		fmt.Println("エラー : Table.xlsxを開けませんでした。")
		fmt.Printf(err1.Error())
	}

	mw.database = excel[6]

	switch mw.combo1.Text() {
	case "hoge":
		mw.datatype = "hoge"
		mw.excelsheet = excel[2]

	case "hage":
		mw.datatype = "hage"
		mw.excelsheet = excel[4]

	}

	for _, item := range mw.excelsheet {
		if item[0] != "" && !strings.Contains(item[0], "フォルダ名") {
			list = append(list, item[0])
		}
	}

	_ = mw.combo2.SetModel(list)
}

func (mw *MyMainWindow) comboChanged2() {

	var list []string

	mw.row = 0
	for _, item := range mw.excelsheet[0] {
		if strings.Contains(item, mw.combo2.Text()) {
			break
		}
		mw.row++
	}

	for _, item := range mw.excelsheet[2:] {
		if item[mw.row] == "" {
			break
		} else if item[mw.row+1] != "廃版" {
			list = append(list, item[mw.row])
		}
	}

	mw.list = list
	_ = mw.combo3.SetModel(list)

}

func (mw *MyMainWindow) comboChanged3() {
	var list []string

	for _, item := range mw.excelsheet[1:] {
		if item[mw.row+2] == "hoge" {
			break
		} else {
			list = append(list, item[mw.row+2])
		}
	}

	_ = mw.combo4.SetModel(list)
}

func (mw *MyMainWindow) comboChanged4() {

	var line string
	var list []string
	var outputlist []string

	mw.dirpath = "$\\" + mw.datatype + "\\" + mw.combo2.Text() + "\\" + mw.combo4.Text() + "\\03_装置構成書(" + mw.combo3.Text() + ")"
	mw.workfield = mw.root + "\\" + mw.datatype + "\\" + mw.combo2.Text() + "\\" + mw.combo4.Text() + "\\03_装置構成書(" + mw.combo3.Text() + ")"

	os.Setenv("ssdir", "\\\\ssdata\\"+mw.userdatabase)
	stdout := runCmdStr("ss", "dir", mw.dirpath, "-r")

	out := string(stdout) //[]byte => string

	//stringから一行ずつ取り出して[]stringへ変換
	for _, s := range out {
		if s == '\n' || s == '\r' {
			list = append(list, line)
			line = ""
		} else if s != '\n' || s == '\r' {
			line = line + string(s)
		}
	}

	for _, text := range list {
		if strings.Contains(text, "構成書") && !strings.Contains(text, "hoge") {
			outputlist = append(outputlist, text)
		}
	}

	mw.modellist = outputlist
	mw.iofoldername = strings.Replace(mw.excelsheet[mw.combo4.CurrentIndex()+1][mw.row+3], "¥n", ",", -1)
	fmt.Println(mw.iofoldername)

	_ = mw.lb.SetModel(mw.modellist) //リストボックスに一覧をセット
}

func (mw *MyMainWindow) pbClicked() {

	prevDir, _ := filepath.Abs(".")
	//ここから作業
	//作業後元の場所に戻る
	defer os.Chdir(prevDir)

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////

	icon1, err := walk.NewIconFromFile(prevDir + "/img/check.ico")
	if err != nil {
		fmt.Println("アイコン1エラー")
	}

	i := mw.lb.CurrentIndex()

	mw.checkoutpath = mw.dirpath + "\\" + mw.modellist[i] //
	iraihyopath := mw.dirpath + "\\" + strings.Replace(mw.modellist[i], "構成書", "依頼票", 1)

	os.Setenv("ssdir", "\\\\ssdata\\"+mw.userdatabase)

	//作業フォルダの存在確認
	//あればnil:無ければディレクトリを作成してnil:エラーの時err
	err = searchandmkdir(mw.workfield)
	if err != nil {
		fmt.Printf("エラー:%s", err)
	}

	os.Chdir(mw.workfield)
	out := runCmdStr("ss", "get", mw.checkoutpath)
	if out == nil {
		mw.aboutAction_Triggered("hogehogehoge")
		return
	}
	out = runCmdStr("ss", "get", iraihyopath)
	if out == nil {
		mw.aboutAction_Triggered("hogehogehoge")
		return
	}

	//////////////////////////////////////////取得処理完了////////////////////////////////////////////////////////

	mw.sbi.SetText("取得完了")
	mw.sbi.SetIcon(icon1)
}

//チェックアウトしていたものをチェックインする
func (mw *MyMainWindow) pbClicked2() {

	iraihyopath := strings.Replace(mw.checkoutpath, "構成書_", "依頼票_", 1)

	os.Setenv("ssdir", "\\\\ssdata\\"+mw.userdatabase)

	prevDir, _ := filepath.Abs(".")
	os.Chdir(mw.workfield)
	//ここから作業

	runCmdStr("ss", "Checkin", iraihyopath)
	runCmdStr("ss", "Checkin", mw.checkoutpath)

	//作業後元の場所に戻る
	os.Chdir(prevDir)

	mw.sbi.SetText("チェックイン完了")
}

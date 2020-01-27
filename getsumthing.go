package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"os/exec"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

func searchandmkdir(dirpath string) error {

	if f, err := os.Stat(dirpath); os.IsNotExist(err) || !f.IsDir() {
		fmt.Printf("%sは存在しません¥n", dirpath)
		if err := os.MkdirAll(dirpath, 0777); err != nil {
			return err
		}
	}

	return nil
}

func (mw *MyMainWindow) databasesearch2() {

	//以下100行は公開できない情報が含まれているため省略

	return
}

func (mw *MyMainWindow) getkisofile() {
	prevDir, _ := filepath.Abs(".")
	//ここから作業
	//作業後元の場所に戻る
	defer os.Chdir(prevDir)

	/////////////////////////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////ファイルを取得//////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////

	os.Setenv("ssdir", "\\\\hoge\\"+mw.kisodatabase) //データベースを指定

	basestring := mw.datatype + "\\" + mw.combo2.Text() + "\\hoge\\" + mw.combo3.Text()

	a := mw.root + "\\" + basestring + "\\hoge"
	b := mw.root + "\\" + basestring + "\\hoge"

	//作業フォルダの存在確認
	//あればnil:無ければディレクトリを作成してnil:エラーの時err
	err := searchandmkdir(a)
	if err != nil {
		fmt.Printf("エラー:%s", err)
		return
	}
	err = searchandmkdir(b)
	if err != nil {
		fmt.Printf("エラー:%s", err)
		return
	}
	/////////////////////////////////////////////////////////

	kisofilepath := "$/" + basestring + "\\hoge"
	formatpath := "$/" + basestring + "\\hoge"

	os.Chdir(a)
	runCmdStr("ss", "get", kisofilepath, "-w", "-r", "-i-y")

	os.Chdir(b)
	runCmdStr("ss", "get", formatpath, "-w", "-r", "-i-y")

	return
}

func (mw *MyMainWindow) getIO() {
	prevDir, _ := filepath.Abs(".")
	//ここから作業
	//作業後元の場所に戻る
	defer os.Chdir(prevDir)

	/////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////表を取得///////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////

	i := mw.lb.CurrentIndex()
	// "装置構成書_601488_1491.xlsm" -> []string{"601488_1491","xlsm"}
	temp := strings.Split(strings.Replace(mw.modellist[i], "構成書_", "", 1), ".")
	if strings.Contains(temp[0], "_") {
		temp = strings.Split(temp[0], "_") // "601488_1491" -> {"601488","1491"}
	}
	mw.seiban = temp[0]

	// fmt.Printf("フォルダは%sです¥n", mw.iofoldername)
	// fmt.Println(mw.seiban)

	os.Setenv("ssdir", "\\\\ssdata\\io6\\")

	ioprojectlist := []string{"$/hoge", "$/hoge", "$/hoge", "$/hoge"}

	var getiopathbyte []byte
	//var tempgetiopath string
	var getiopath string
	temparray := strings.Split(mw.iofoldername, ",")
	for _, item := range ioprojectlist {
		fmt.Println(item)
		for _, m := range temparray {
			getiopathbyte = runCmdStr("ss", "dir", item+"/"+m+"/_"+mw.seiban, "-F-") //[]byte型をうまく引き回しできないので
			if getiopathbyte != nil {                                                //戻り値はここの判定にしか使わない
				fmt.Println("ミッケ")
				//tempgetiopath = string(getiopathbyte)
				getiopath = item + "/" + m + "/_" + mw.seiban
				goto L
			}
		}
	}

	if getiopath == "" {
		mw.aboutAction_Triggered("見つかりませんでした。データベースを確認してください")
		return
	}

L:

	localfolder := strings.Replace(getiopath, "$", mw.root, 1)

	//作業フォルダの存在確認
	//あればnil:無ければディレクトリを作成してnil:エラーの時err
	err := searchandmkdir(localfolder)
	if err != nil {
		fmt.Printf("エラー:%s", err)
	}

	os.Chdir(localfolder)
	runCmdStr("ss", "get", getiopath, "-w", "-r", "-i-")

	return
}

func getmacroinfo(root string) {

	prevDir, _ := filepath.Abs(".")

	//作業後もとのディレクトリに帰る
	defer os.Chdir(prevDir)

	os.Setenv("ssdir", "\\\\hoge\\hoge")

	//作業フォルダの存在確認
	//あればnil:無ければディレクトリを作成してnil:エラーの時err
	err := searchandmkdir(root + "\\hoge")
	if err != nil {
		fmt.Printf("エラー:%s", err)
		return
	}

	err = searchandmkdir(root + "hoge")
	if err != nil {
		fmt.Printf("エラー:%s", err)
		return
	}

	err = searchandmkdir(root + "hoge")
	if err != nil {
		fmt.Printf("エラー:%s", err)
		return
	}

	os.Chdir(root)

	///////////////////////////////////////////////////////////////////////////////////////////
	///////////////////////////////////////DataMacroの更新/////////////////////////////////////
	///////////////////////////////////////////////////////////////////////////////////////////

	//os.Chdir(root + "\\00_FC3000DataMacro")

	runCmdStr("ss", "get", "$/hoge", "-w", "-r", "-i-")

	///////////////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////InfoSheetの更新///////////////////////////////////////
	///////////////////////////////////////////////////////////////////////////////////////////
	os.Chdir(root + "\\hoge")
	runCmdStr("ss", "get", "$/hoge", "-w", "-r", "-i-y-GCD")
	os.Chdir(root + "\\hoge")
	runCmdStr("ss", "get", "$/hoge", "-w", "-r", "-i-y-GCD")

	////////////////////////////////////////////////////////////////////////////////////
	///////////////////////////////referrencetable.xlsmの取得////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////

	//ここの処理はリファレンステーブルの拡張子がxlsだとgoで扱えないからxlsxを作るためにやっている
	//windows COM という技術です
	ole.CoInitialize(0)

	unknown, _ := oleutil.CreateObject("Excel.Application")
	excel, _ := unknown.QueryInterface(ole.IID_IDispatch)

	oleutil.PutProperty(excel, "DisplayAlerts", false)
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook := oleutil.MustCallMethod(workbooks, "Open", root+"\\hoge\\Table.xls").ToIDispatch()

	//xlsxで上書き保存
	oleutil.MustCallMethod(workbook, "SaveAs", prevDir+"\\Table.xlsx", 51, nil, nil).ToIDispatch()
	oleutil.CallMethod(workbook, "Close", false)
	oleutil.CallMethod(excel, "Quit")

	excel.Release() //リソースを破棄

	ole.CoUninitialize()

	//ここまで

	return
}

//func runCmdStr(cmdstr []string) []byte {
func runCmdStr(str ...string) []byte {

	cmdstr := []string{"cmd", "/c"}

	cmdstr = append(cmdstr, str...)

	cmd := exec.Command(cmdstr[0], cmdstr[1:]...)

	stdoutpipe, err := cmd.StdoutPipe()
	if err != nil {
		fmt.Printf("StdoutPipe Error: %v¥n", err)
		return nil
	}
	defer stdoutpipe.Close()

	err = cmd.Start()
	if err != nil {
		fmt.Printf("Command Start Error: %v¥n", err)
		return nil
	}

	stdout, err := ioutil.ReadAll(
		transform.NewReader(stdoutpipe, japanese.ShiftJIS.NewDecoder()))
	if err != nil {
		fmt.Printf("Command Error: %v¥n", err)
		return nil
	}

	err = cmd.Wait()
	if err != nil {
		fmt.Printf("Command Wait Error: %v¥n", err)
		return nil
	}

	fmt.Printf("%s¥n", stdout)

	return stdout
}

package main

import (
        "fmt"
        "os"

        "github.com/lxn/walk"
        . "github.com/lxn/walk/declarative"
        "gopkg.in/ini.v1"
)

type ErrorWindow struct {
        *walk.MainWindow
}

type MyMainWindow struct {
        *walk.MainWindow
        combo1 *walk.ComboBox
        combo2 *walk.ComboBox
        combo3 *walk.ComboBox
        combo4 *walk.ComboBox
        sbi    *walk.StatusBarItem
        lb     *walk.ListBox

        //model *EnvModel
        modellist []string

        root      string //作業フォルダのルートパス
        path      string //VSSのインストールパス
        dirpath      string
        username     string
        password     string
        checkoutpath string
        kisodatabase string
        userdatabase string
        workfield    string
        seiban       string
        iofoldername string
        datatype     string

        list []string

        //xlfile *file
        excelsheet [][]string
        database   [][]string
        row        int
}

func main() {
        
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////configの読み込み///////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////

        configdata, err := ini.Load("config.ini")

        if err != nil {
                er := &ErrorWindow{}
                walk.MsgBox(er,
                        "エラー",
                        "config.iniがありません",
                        walk.MsgBoxOK|walk.MsgBoxIconInformation)
                os.Exit(1)
        }

        mw := &MyMainWindow{

                root:  configdata.Section("hoge").Key("root").String(),
                path:  configdata.Section("hoge").Key("path").String(),
                username: configdata.Section("hoge").Key("hoge").String(),
                password: configdata.Section("hoge").Key("hoge").String(),
                //model:    NewEnvModel(),
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////環境変数の設定////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        str := os.Getenv("PATH")

        str = str + ";" + mw.path
        os.Setenv("PATH", str)

        os.Setenv("hoge", mw.username)
        os.Setenv("hoge", mw.password)

        ///////////////////////////////////////////////////////////////////////////////////////////
        ///////////////DataMacro,InfoSheet,referrencetable.xlsmの更新//////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////

        getmacroinfo(mw.root)

        ////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////

        var machintype = []string{ // Combo1項目リスト
                "hoge",
                "hage",
        }

        MW := MainWindow{ //GUI画面の表示を宣言
                AssignTo: &mw.MainWindow,
                Title:    "ファイル取得",
                MinSize:  Size{300, 200},
                Size:     Size{350, 450},
                // MenuItems: []MenuItem{
                //      Menu{
                //              Text: "&Option",
                //              Items: []MenuItem{
                //                      Action{
                //                              Text:        "Setting",
                //                              OnTriggered: func() { mw.aboutActionTriggered() },
                //                      },
                //              },
                //      },
                // },
                Layout: VBox{},
                Children: []Widget{

                        Composite{
                                Layout: HBox{},
                                Children: []Widget{
                                        Label{
                                                Text: "type:",
                                        },
                                        ComboBox{
                                                AssignTo:              &mw.combo1,
                                                Model:                 machintype,
                                                OnCurrentIndexChanged: mw.comboChanged,
                                        },
                                },
                        },

                        Composite{
                                Layout: HBox{},
                                Children: []Widget{
                                        Label{
                                                Text: "フォルダ:",
                                        },
                                        ComboBox{
                                                AssignTo: &mw.combo2,
                                                //Model:                 mw.BaseFolder,
                                                OnCurrentIndexChanged: mw.comboChanged2,
                                        },
                                },
                        },

                        Composite{
                                Layout: HBox{},
                                Children: []Widget{
                                        Label{
                                                Text: "version:",
                                        },
                                        ComboBox{
                                                AssignTo: &mw.combo3,
                                                //Model:    mw.version,
                                                OnCurrentIndexChanged: mw.comboChanged3,
                                        },
                                },
                        },
                        Composite{
                                Layout: HBox{},
                                Children: []Widget{
                                        Label{
                                                Text: "user:",
                                        },
                                        ComboBox{
                                                AssignTo: &mw.combo4,
                                                //Model:    mw.version,
                                                OnCurrentIndexChanged: func() {
                                                        mw.databasesearch2()
                                                        mw.comboChanged4()
                                                },
                                        },
                                },
                        },
                        Composite{
                                Layout: VBox{},
                                Children: []Widget{
                                        Label{
                                                Text: "一覧:",
                                        },
                                        ListBox{
                                                AssignTo: &mw.lb,
                                                //Model:                 mw.model,
                                                //OnCurrentIndexChanged: mw.lb_CurrentIndexChanged,
                                                //OnItemActivated: mw.lb_ItemActivated,
                                        },
                                },
                        },

                        PushButton{
                                Text: "取得",
                                //OnClicked: mw.pbClicked,
                                OnClicked: func() {
                                        switch {
                                        case mw.combo1.Text() == "":
                                                mw.sbi.SetText("combo1選択してください")
                                                return
                                        case mw.combo2.Text() == "":
                                                mw.sbi.SetText("combo2を選択してください")
                                                return
                                        case mw.combo3.Text() == "":
                                                mw.sbi.SetText("combo3を選択してください")
                                                return
                                        case mw.combo4.Text() == "" || mw.lb.CurrentIndex() == -1:
                                                mw.sbi.SetText("combo4を選択してください")
                                                return
                                        }
                                        mw.sbi.SetText("取得中です")
                                        mw.getkisofile() 
                                        mw.getIO()       
                                        mw.pbClicked()   
                                },
                        },
                        // PushButton{
                        //      Text:      "作業終了",
                        //      OnClicked: mw.pbClicked2,
                        // },
                },
                StatusBarItems: []StatusBarItem{

                        StatusBarItem{
                                AssignTo: &mw.sbi,
                                Text:     "スタンバイ : 最新の状態です",
                                Width:    120,
                                //ToolTipText: "no tooltip for me",
                        },
                },
        }

        if _, err := MW.Run(); err != nil {
                fmt.Fprintln(os.Stderr, err)
                os.Exit(1)
        }

}

func (mw *MyMainWindow) aboutAction_Triggered(msg string) {
        walk.MsgBox(mw,
                "エラー",
                msg,
                walk.MsgBoxOK|walk.MsgBoxIconInformation)
}

// func (mw *MyMainWindow) aboutActionTriggered() {
//      var dlg *walk.Dialog
//      var db *walk.DataBinder
//      var acceptPB, cancelPB *walk.PushButton

//      return Dialog{
//              AssignTo:      &dlg,
//              Title:         "Setting",
//              DefaultButton: &acceptPB,
//              CancelButton:  &cancelPB,
//              DataBinder: DataBinder{
//                      AssignTo: &db,
//                      Name:     "animal",
//                      //DataSource:     animal,
//                      ErrorPresenter: ToolTipErrorPresenter{},
//              },
//              MinSize: Size{300, 300},
//              Layout:  VBox{},
//              Children: []Widget{
//                      Composite{
//                              Layout: Grid{Columns: 2},
//                              Children: []Widget{
//                                      Label{
//                                              Text: "Name:",
//                                      },
//                                      LineEdit{
//                                              Text: Bind("Name"),
//                                      },

//                                      Label{
//                                              Text: "Arrival Date:",
//                                      },
//                                      DateEdit{
//                                              Date: Bind("ArrivalDate"),
//                                      },

//                                      Label{
//                                              Text: "Species:",
//                                      },
//                                      ComboBox{
//                                              Value:         Bind("SpeciesId", SelRequired{}),
//                                              BindingMember: "Id",
//                                              DisplayMember: "Name",
//                                              //Model:         KnownSpecies(),
//                                      },

//                                      Label{
//                                              Text: "Speed:",
//                                      },
//                                      Slider{
//                                              Value: Bind("Speed"),
//                                      },

//                                      Label{
//                                              Text: "Weight:",
//                                      },
//                                      NumberEdit{
//                                              Value:    Bind("Weight", Range{0.01, 9999.99}),
//                                              Suffix:   " kg",
//                                              Decimals: 2,
//                                      },

//                                      Label{
//                                              Text: "Domesticated:",
//                                      },
//                                      CheckBox{
//                                              Checked: Bind("Domesticated"),
//                                      },

//                                      VSpacer{
//                                              ColumnSpan: 2,
//                                              Size:       8,
//                                      },

//                                      Label{
//                                              ColumnSpan: 2,
//                                              Text:       "Remarks:",
//                                      },
//                                      TextEdit{
//                                              ColumnSpan: 2,
//                                              MinSize:    Size{100, 50},
//                                              Text:       Bind("Remarks"),
//                                      },

//                                      Label{
//                                              ColumnSpan: 2,
//                                              Text:       "Patience:",
//                                      },
//                                      LineEdit{
//                                              ColumnSpan: 2,
//                                              Text:       Bind("PatienceField"),
//                                      },
//                              },
//                      },
//                      Composite{
//                              Layout: HBox{},
//                              Children: []Widget{
//                                      HSpacer{},
//                                      PushButton{
//                                              AssignTo: &acceptPB,
//                                              Text:     "OK",
//                                              // OnClicked: func() {
//                                              //      if err := db.Submit(); err != nil {
//                                              //              log.Print(err)
//                                              //              return
//                                              //      }

//                                              //      dlg.Accept()
//                                              // },
//                                      },
//                                      PushButton{
//                                              AssignTo: &cancelPB,
//                                              Text:     "Cancel",
//                                              //OnClicked: return,
//                                      },
//                              },
//                      },
//              },
//      }return

// }





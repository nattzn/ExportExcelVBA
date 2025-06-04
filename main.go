package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"runtime"

	"github.com/gdamore/tcell/v2"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/rivo/tview"
)

type FileEntry struct {
	Name     string
	FullPath string
}

func main() {
	// フラグで出力先ディレクトリを指定
	outDir := flag.String("d", "exported", "出力先ディレクトリ")
	flag.Parse()

	// .xlsm ファイル一覧取得
	dir, _ := os.Getwd()
	entries, _ := os.ReadDir(dir)

	var filesList []FileEntry
	for _, entry := range entries {
		if !entry.IsDir() && filepath.Ext(entry.Name()) == ".xlsm" {
			filesList = append(filesList, FileEntry{
				Name:     entry.Name(),
				FullPath: filepath.Join(dir, entry.Name()),
			})
		}
	}

	if len(filesList) == 0 {
		log.Fatal("xlsmファイルが見つかりません")
	}

	// TUIで選択 → index取得（UIは別スレッドでもOK）
	selectedIdx := runTUI(filesList)
	if selectedIdx >= 0 {

		// mainスレッドでOLE処理（安全）
		selected := filesList[selectedIdx]
		fmt.Println("選択ファイル:", selected.FullPath)

		// 出力先ディレクトリ
		out_path, err := filepath.Abs(*outDir)
		if err != nil {
			log.Fatalf("出力ディレクトリの絶対パス取得に失敗: %v", err)
		}
		os.MkdirAll(out_path, 0755)
		fmt.Println("出力フォルダ:", out_path)

		exportVBA(selected.FullPath, out_path)
	}
}

func runTUI(filesList []FileEntry) int {
	app := tview.NewApplication()
	textView := tview.NewTextView().SetDynamicColors(true).SetWrap(false)
	cursor := 0

	update := func() {
		textView.Clear()
		fmt.Fprintf(textView, "スクリプトを取り出すファイルを選択してください\n")
		for i, f := range filesList {
			prefix := "   "
			if i == cursor {
				prefix = "▶ "
			}
			fmt.Fprintf(textView, "%s%s\n", prefix, f.Name)
		}
	}

	update()

	textView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		switch event.Key() {
		case tcell.KeyUp:
			if cursor > 0 {
				cursor--
			}
		case tcell.KeyDown:
			if cursor < len(filesList)-1 {
				cursor++
			}
		case tcell.KeyEnter:
			app.Stop()
		case tcell.KeyEscape:
			cursor = -1
			app.Stop()
		}
		update()
		return nil
	})

	if err := app.SetRoot(textView, true).Run(); err != nil {
		log.Fatal(err)
	}
	return cursor
}

func exportVBA(path string, out_path string) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := ole.CoInitialize(0); err != nil {
		log.Fatal("OLE初期化失敗（exportVBA）:", err)
	}
	defer ole.CoUninitialize()

	excelObj, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		log.Fatal("Excel.Application 作成失敗:", err)
	}
	defer excelObj.Release()

	excelDisp, err := excelObj.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal("QueryInterface 失敗:", err)
	}
	defer excelDisp.Release()

	oleutil.PutProperty(excelDisp, "Visible", false)
	oleutil.PutProperty(excelDisp, "EnableEvents", false)
	oleutil.PutProperty(excelDisp, "DisplayAlerts", false)
	oleutil.PutProperty(excelDisp, "AutomationSecurity", 3) // msoAutomationSecurityForceDisable

	workbooksRaw, err := oleutil.GetProperty(excelDisp, "Workbooks")
	if err != nil || workbooksRaw == nil {
		log.Fatal("Workbooks取得失敗:", err)
	}
	workbooks := workbooksRaw.ToIDispatch()

	wbRaw, err := oleutil.CallMethod(workbooks, "Open", path)
	if err != nil || wbRaw == nil {
		log.Fatal("ファイルオープン失敗:", err)
	}
	wb := wbRaw.ToIDispatch()
	defer wb.Release()

	vbprojRaw, err := oleutil.GetProperty(wb, "VBProject")
	if err != nil || vbprojRaw == nil {
		log.Fatal("VBProject取得失敗:", err)
	}
	vbproj := vbprojRaw.ToIDispatch()

	componentsRaw, err := oleutil.GetProperty(vbproj, "VBComponents")
	if err != nil || componentsRaw == nil {
		log.Fatal("VBComponents取得失敗:", err)
	}
	components := componentsRaw.ToIDispatch()

	countRaw, err := oleutil.GetProperty(components, "Count")
	if err != nil {
		log.Fatal("Count取得失敗:", err)
	}
	count := int(countRaw.Val)

	for i := 1; i <= count; i++ {
		compRaw, err := oleutil.CallMethod(components, "Item", i)
		if err != nil || compRaw == nil {
			log.Printf("Item(%d)取得失敗: %v", i, err)
			continue
		}
		comp := compRaw.ToIDispatch()

		nameRaw, err := oleutil.GetProperty(comp, "Name")
		if err != nil {
			log.Println("モジュール名取得失敗:", err)
			comp.Release()
			continue
		}
		name := nameRaw.ToString()

		// モジュール種別の取得（Type）
		typeRaw, err := oleutil.GetProperty(comp, "Type")
		if err != nil {
			comp.Release()
			continue
		}
		typeVal := int(typeRaw.Val)
		// ファイル拡張子の決定
		ext := ".vbs" // fallback シートモジュールなど
		switch typeVal {
		case 1:
			ext = ".bas.vbs" // 標準モジュール
		case 2:
			ext = ".cls.vbs" // クラスモジュール
		case 3:
			ext = ".frm.vbs" // フォーム
		default:
			// 続行
		}

		codeModRaw, err := oleutil.GetProperty(comp, "CodeModule")
		if err != nil || codeModRaw == nil {
			log.Printf("CodeModule取得失敗: %v", err)
			comp.Release()
			continue
		}
		codeMod := codeModRaw.ToIDispatch()

		linesRaw, err := oleutil.GetProperty(codeMod, "CountOfLines")
		if err != nil {
			log.Printf("CountOfLines取得失敗: %v", err)
			codeMod.Release()
			comp.Release()
			continue
		}
		lines := int(linesRaw.Val)
		if lines == 0 {
			codeMod.Release()
			comp.Release()
			continue
		}

		codeRaw, err := oleutil.GetProperty(codeMod, "Lines", 1, lines)
		if err != nil || codeRaw == nil {
			log.Printf("Lines取得失敗 (%s): %v", name, err)
			codeMod.Release()
			comp.Release()
			continue
		}
		code := codeRaw.ToString()

		// 書き込み
		filename := filepath.Join(out_path, name+ext)
		err = os.WriteFile(filename, []byte(code), 0644)
		if err == nil {
			fmt.Println("出力:", filename)
		} else {
			log.Printf("書き込み失敗: %v", err)
			codeMod.Release()
			comp.Release()
			continue
		}

		codeMod.Release()
		comp.Release()
	}

	wb.CallMethod("Close", false)
	excelDisp.CallMethod("Quit")
}

package main

import (
	"errors"
	"fmt"
	"os"
	"path"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/tealeg/xlsx"
)

const (
	mainWindowHight = 600
	mainWindowWidth = 600

	mainTitleText   = "Merge Excels"
	mainContentText = "select the folder where the Excel files you want to merge are located"
)

func main() {
	myApp := app.New()

	win := initMainWindow(myApp)

	showSelectedFolder, changeSelectedFolder := useLabel("Show Selected Folder", "[empty]")

	selectBtn := widget.NewButton("SELECT", func() {
		dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, win)
				return
			}

			if uri == nil {
				return
			}

			fileInfo, err := os.Stat(uri.Path())
			if err != nil {
				dialog.ShowError(err, win)
				return
			}

			if !fileInfo.IsDir() {
				dialog.ShowError(errors.New("select dir rather than file"), win)
				return
			}

			changeSelectedFolder(uri.Path())
		}, win)
	})

	mergeBtn := widget.NewButton("MERGE", func() {
		if err := mergeExcels(showSelectedFolder.Text); err != nil {
			dialog.ShowCustom("ERROR", "OK", widget.NewLabel(err.Error()), win)
		} else {
			dialog.ShowCustom("NOTIFY", "OK", widget.NewLabel("call merge action"), win)
		}
	})

	win.SetContent(container.NewVBox(
		initContent(),
		showSelectedFolder,
		selectBtn,
		mergeBtn,
	))

	win.ShowAndRun()
}

func initMainWindow(myApp fyne.App) fyne.Window {
	win := myApp.NewWindow(mainTitleText)
	win.Resize(fyne.NewSize(mainWindowHight, mainWindowWidth))
	return win
}

func initContent() *widget.Label {
	return widget.NewLabel(mainContentText)
}

func useLabel(label, content string) (*widget.Label, func(string)) {
	_ = label
	win := widget.NewLabel(content)
	return win, win.SetText
}

func mergeExcels(dir string) error {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return fmt.Errorf("read dir err: %v", err)
	}

	newXlFile := xlsx.NewFile()
	newSheet, err := newXlFile.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("add sheet err: %v", err)
	}

	for _, entry := range entries {
		wholePath := path.Join(dir, entry.Name())

		if ext := path.Ext(wholePath); ext != ".xlsx" {
			continue
		}

		xlFile, err := xlsx.OpenFile(wholePath)
		if err != nil {
			return fmt.Errorf("xl open err: %v", err)
		}

		if len(xlFile.Sheets) < 1 {
			continue
		}

		for _, row := range xlFile.Sheets[0].Rows {
			newRow := newSheet.AddRow()
			// 遍历每一列
			for _, cell := range row.Cells {
				newCell := newRow.AddCell()
				newCell.Value = cell.Value
			}
		}
	}

	if err := newXlFile.Save("merged.xlsx"); err != nil {
		return fmt.Errorf("merged err: %v", err)
	}

	return nil
}

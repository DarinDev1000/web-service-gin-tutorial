package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type Album struct {
	ID     string  `json:"id"`
	Title  string  `json:"title"`
	Artist string  `json:"artist"`
	Price  float64 `json:"price"`
}

func CreateExcelFile() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	_, err := f.NewSheet("Albums")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Set Headers
	f.SetSheetRow("Albums", "A1", &[]interface{}{"ID", "Title", "Artist", "Price"})
	if err := f.DeleteSheet("Sheet1"); err != nil {
		fmt.Println(err)
		return
	}
	// // Set active sheet of the workbook.
	// f.SetActiveSheet(index)

	saveFile(f)
}

func AddAlbum(album Album) {
	f, err := openFile()
	if err != nil {
		fmt.Println(err)
		return
	}
	// Add Values
	f.SetSheetRow("Albums")
	var a Album = Album
}

func openFile() (*excelize.File, error) {
	f, err := excelize.OpenFile("Albums.xlsx")
	if err != nil {
		fmt.Println(err)
		return &excelize.File{}, err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	return f, nil
}

func saveFile(f *excelize.File) error {
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Albums.xlsx"); err != nil {
		fmt.Println(err)
		return err
	}
	return nil
}

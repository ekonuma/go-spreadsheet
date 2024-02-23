package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	/*f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	index, err := f.NewSheet("Sheet2")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(index)
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)
	// f.SetActiveSheet(index)
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
	*/
	file, err := excelize.OpenFile("Book1.xlsx")
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
			return
		}
	}()
	if err != nil {
		fmt.Println(err)
	}
	file.SetCellValue("Sheet1", "A1", "First edit")
	file.Save()
}

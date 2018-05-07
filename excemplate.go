package main

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type inOut int

const (
	in = iota
	out
)

type datacell struct {
	List int
	Cell int
	Row  int
	Data string
}

//RootFolder - return current folder of running file
func RootFolder() string {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	rootPath := filepath.Dir(ex)
	return rootPath
}

//CreateFolder - create folder requied
func CreateFolder(folder string) {
	os.MkdirAll(folder, os.ModePerm)
}

//DeleteFolder - create folder requied
func DeleteFolder(folder string) {
	os.Remove(folder)
}

//ClearFolder - remove all contents of a directory
func ClearFolder(folder string) {
	dtemplate, err := os.Open(folder)
	if err != nil {
		panic(err)
	}
	defer dtemplate.Close()

	names, err := dtemplate.Readdirnames(-1)
	if err != nil {
		panic(err)
	}
	for _, name := range names {
		err = os.RemoveAll(filepath.Join(folder, name))
		if err != nil {
			panic(err)
		}
	}
}

func main() {
	xlsx, err := excelize.OpenFile(".\\TEMPLATE\\IN.xlsx")
	if err != nil {
		os.Exit(1)
	}
	x := foundData(xlsx)
	xlsx2, err := excelize.OpenFile(".\\TEMPLATE\\OUT.xlsx")
	if err != nil {
		os.Exit(1)
	}
	y := foundData(xlsx2)
	fmt.Println(x, y)
	fmt.Println(isCorrected(&x, &y))
}

func myminitest() {
	folders := []string{".\\TEMPLATE", ".\\OUTPUT", ".\\INPUT"}
	for _, f := range folders {
		k := fileInFolder(f, 0)
		fmt.Println(k)
	}
}

//Корректность шаблонов вход/выход
//TODO:: add found all close?
//TODO:: space #LName and #FName
func isCorrected(in, out *[]datacell) (ok bool, message []string) {
	var inData, outData []string
	for _, dataIn := range *in {
		inData = append(inData, dataIn.Data[:len(dataIn.Data)-2])
	}
	for _, dataOut := range *out {
		outData = append(outData, dataOut.Data[:len(dataOut.Data)-2])
	}
	var diff []string

	// Loop two times, first to find slice1 strings not in slice2,
	// second loop to find slice2 strings not in slice1
	for i := 0; i < 2; i++ {
		for _, s1 := range inData {
			ok = false
			for _, s2 := range outData {
				if s1 == s2 {
					ok = true
					break
				}
			}
			// String not found. We add it to return slice
			if !ok {
				if i == 0 {
					diff = append(diff, "В входящем нет: "+s1)
				} else {
					diff = append(diff, "В исходящем нет: "+s1)
				}
			}
		}
		// Swap the slices, only if it was the first loop
		if i == 0 {
			inData, outData = outData, inData
		}
	}
	return ok, diff
}

func readAndWriteData(in, out *[]datacell, fileData, fileTemple *excelize.File) error {

	return nil
}

//listOfFile take all filenames
func listOfFile(root string) (files []string) {
	err := filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if filepath.Ext(path) == ".xlsx" {
			files = append(files, path)
		}
		return nil
	})
	if err != nil {
		panic(err)
	}
	return files
}

//Find file with type
func fileInFolder(root string, inout inOut) string {
	var file string
	err := filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if inout == 0 {
			if info.Name() == "IN.xlsx" {
				file += path
			}
		} else if inout == 1 {
			if info.Name() == "OUT.xlsx" {
				file += path
			}
		}
		return nil
	})
	if err != nil {
		panic(err)
	}
	return file
}

func newcell(su, ru, cu int, data string) datacell {
	return datacell{
		List: su,
		Cell: ru,
		Row:  cu,
		Data: data,
	}
}

//Create list of templates
func foundData(file *excelize.File) (data []datacell) {
	for s, n := range findAllSheet(file) {
		rows := file.GetRows(n) //take row from all sheest
		for r, row := range rows {
			for c, cell := range row {
				if strings.HasPrefix(cell, "#") {
					for _, i := range strings.Split(cell, ";") {
						temp := newcell(s, r, c, i)
						data = append(data, temp)
					}
				}
			}
		}
	}
	return
}

func findAllSheet(file *excelize.File) (s []string) {
	for i := 0; i < file.SheetCount+1; i++ { //coz initilazre with 0
		s = append(s, file.GetSheetName(i))
	}
	return
}

// Если не равны вывести из первой что нет во второй
func cheakandfinddifirend(in, out *[]datacell) (different bool, s []string) {
	different = reflect.DeepEqual(in, out)
	if !different {
		mb := map[string]bool{}
		for _, x := range *out {
			mb[x.Data] = true
		}
		for _, x := range *in {
			if _, ok := mb[x.Data]; !ok {
				s = append(s, x.Data)
			}
		}
		return
	}
	return
}

//Found correct folder for selected template
//	Сравнимаем 2 файла загрузки и выгрузки и находим расхождения, если найдены сообщаем, значет необходимо добавит в ввыходной шаблон?
//		обходим входные файлы и записываем данные

//Create list of files and store it
//Create new folder for output correct template
//Collect data from file
//Write data

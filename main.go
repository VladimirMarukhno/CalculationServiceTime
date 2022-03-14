package main

import (
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
	"math"
	"strconv"
)

func exel(input string) {
	f, err := excelize.OpenFile(input) //Открываем exel файл с данными сотрудников
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	conversionSum(rows, input)
}

func conversionSum(rows [][]string, name string) {
	var output []float32
	for i := 0; i < len(rows); i++ {
		var sumMinute, sumHour int
		var minut, sum float32
		for a := 0; a < len(rows[i]); a++ { // Преобразуем значение стринг во float и разделяем часы и минуты
			if n, err := strconv.ParseFloat(rows[i][a], 32); err == nil {
				n = n * 100
				sumMinute += int(math.Round(n)) % 100
				sumHour += int(math.Round(n/100))
			} else {
				continue
			}
		}
		if sumMinute > 60 { // Вычисляем часы из минут
			sumHour += sumMinute / 60
			sumMinute = sumMinute % 60
			minut = float32(sumMinute) / 100
		} else {
			minut = float32(sumMinute) / 100
		}
		sum = float32(sumHour) + minut
		output = append(output, sum)
	}
	NewFile(output, name, rows)
}

func NewFile(sumTime []float32, name string, rows [][]string) {
	f, err := excelize.OpenFile(name) //Открываем exel файл с данными сотрудников
	if err != nil {
		fmt.Println(err)
		return
	}
	index := f.NewSheet("Sheet2")
	for i := 0; i < len(sumTime); i++ { // заполняем таблицу
		s := strconv.Itoa(i + 1)
		f.SetCellValue("Sheet2", "A"+s, rows[i][0])
		f.SetCellValue("Sheet2", "B"+s, sumTime[i])
		if err := f.SaveAs(name); err != nil {
			fmt.Println(err)
		}
	}
	f.SetActiveSheet(index)
}

func OpenFileManager() fyne.URI {
	var name fyne.URI
	a := app.New()
	w := a.NewWindow("Укажите путь к файлу.") // создаём новое окно
	w.Resize(fyne.NewSize(800, 800)) // указываем размер окна
	btn := widget.NewButton("Укажите .xlsx файл", func() {
		file_Dialog := dialog.NewFileOpen(
			func(r fyne.URIReadCloser, _ error) {
				name = r.URI() // присваеваем значение с информацией о файле : путь, название .....
				w.Close()
			},w)
		file_Dialog.Resize(fyne.NewSize(1920,1080)) // размер окна с поиском файла.
		file_Dialog.SetFilter(
			storage.NewExtensionFileFilter([]string{".xlsx"})) // указываем тип расщирений файлов которые будут отображаться.
		file_Dialog.Show()
	})

	w.SetContent(container.NewCenter(
		btn,
	))
	w.ShowAndRun()
	return name
}

func main() {
	name := OpenFileManager() // открытие файлового менеджера
	exel(name.Path())
}

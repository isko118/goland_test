package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type Person struct {
	Name   string
	Age    int
	Height float64
	Weight float64
}

func main() {
	f, err := excelize.OpenFile("Sheet1.xlsx")
	if err != nil {
		log.Fatalf("Ошибка при открытии файла: %v", err)
	}
	defer f.Close()
	rows, err := f.GetRows("Лист1")
	if err != nil {
		log.Fatalf("Ошибка при чтении листа: %v", err)
	}
	var filteredPeople []Person
	for i, row := range rows {
		if i == 0 || len(row) < 4 {
			continue
		}
		name := strings.TrimSpace(row[0])
		if name == "" {
			continue
		}
		age, err := strconv.Atoi(strings.TrimSpace(row[1]))
		if err != nil || age < 18 || age > 55 {
			continue
		}
		weightStr := strings.Replace(strings.TrimSpace(row[2]), ",", ".", -1)
		weight, err := strconv.ParseFloat(weightStr, 64)
		if err != nil || weight <= 0 || weight > 500 {
			continue
		}
		heightStr := strings.Replace(strings.TrimSpace(row[3]), ",", ".", -1)
		height, err := strconv.ParseFloat(heightStr, 64)
		if err != nil || height <= 0 || height > 300 {
			continue
		}
		filteredPeople = append(filteredPeople, Person{
			Name:   name,
			Age:    age,
			Weight: weight,
			Height: height,
		})
	}
	newFile := excelize.NewFile()
	newFile.SetSheetName("Sheet1", "ИМТ")
	headers := []string{"Имя", "Возраст", "Рост (см)", "Вес (кг)", "ИМТ"}
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		newFile.SetCellValue("ИМТ", cell, header)
	}
	for row, person := range filteredPeople {
		cellA, _ := excelize.CoordinatesToCellName(1, row+2)
		cellB, _ := excelize.CoordinatesToCellName(2, row+2)
		cellC, _ := excelize.CoordinatesToCellName(3, row+2)
		cellD, _ := excelize.CoordinatesToCellName(4, row+2)
		cellE, _ := excelize.CoordinatesToCellName(5, row+2)
		newFile.SetCellValue("ИМТ", cellA, person.Name)
		newFile.SetCellValue("ИМТ", cellB, person.Age)
		newFile.SetCellValue("ИМТ", cellC, person.Height)
		newFile.SetCellValue("ИМТ", cellD, person.Weight)
		newFile.SetCellFormula("ИМТ", cellE, fmt.Sprintf("=%s/(%s/100)^2", cellD, cellC))
	}
	filename := time.Now().Format("06-01-02 15-04") + ".xlsx"
	if err := newFile.SaveAs(filename); err != nil {
		log.Fatalf("Ошибка при сохранении файла: %v", err)
	}
	fmt.Printf("Создан файл: %s\nОбработано записей: %d\n", filename, len(filteredPeople))
}

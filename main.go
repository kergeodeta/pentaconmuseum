package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/pkg/errors"
	"html/template"
	"log"
	"os"
	"strconv"
	"strings"
)

const htmlPath = "./generated"
const IdIndex = 0
const NameIndex = 1

var fromColumn = 'A'
var toColumn = 'U'
var spreadsheetPath *string

var tpl *template.Template

func init() {
	var err error

	tpl, err = template.ParseGlob("./templates/*.gohtml")
	if err != nil {
		log.Fatalf("HTML sablonok betöltése sikertelen! %s\n", err.Error())
	}

	if _, err = os.Stat(htmlPath); os.IsNotExist(err) {
		if err := os.Mkdir(htmlPath, 0755); err != nil {
			log.Fatalf("Generált HTML fájlok mappájának létrehozása sikertelen! %s\n", err.Error())
		}
	}

	spreadsheetPath = flag.String("in", "", "Bemeneti adatokat tartalmazó XLSX fájl elérési útvonala")
	fromColParam := flag.String("fromColumn", "A", "Melyik oszloptól kezdődjön az adatok beolvasása")
	fromColumn = []rune(*fromColParam)[0]
	toColParam := flag.String("toColumn", "U", "Melyik oszlopig kerüljenek felolvasásra az értékek")
	toColumn = []rune(*toColParam)[0]

	flag.Parse()
}

func get(s []string, i int) string {
	return s[i]
}

func main() {
	xlsx, err := excelize.OpenFile(*spreadsheetPath)
	if err != nil {
		log.Fatalf("A megadott XLSX fájl megnyitása (%s) sikertelen! %s\n", *spreadsheetPath, err.Error())
	}

	sheets := xlsx.GetSheetMap()
	workingSheet := sheets[1]

	header, err := readHeader(xlsx, workingSheet)
	if err != nil {
		log.Fatalf("A fejléc felolvasása sikertelen! %s\n", err.Error())
	}

	rowIndex := 2
	contents := make(map[int]string)
	for {
		row, err := readRow(xlsx, workingSheet, rowIndex)
		if err != nil {
			log.Println(err.Error())
			continue
		}

		id, err := strconv.Atoi(row[IdIndex])
		if err != nil {
			log.Printf("A %d. sorban megadott azonosító (%v) nem megfelelő formátumú. Csak pozitív egész számok! %s\n", rowIndex, row[IdIndex], err.Error())
			continue
		}

		contents[id] = row[NameIndex]

		var position string
		if rowIndex == 2 {
			position = "first"
		}

		if !hasContent(xlsx, workingSheet, rowIndex+1) {
			position = "last"
		}

		values, err := readRow(xlsx, workingSheet, rowIndex)
		if err != nil {
			log.Println(err.Error())
			continue
		}

		if err := generateHtml(header, values, position); err != nil {
			log.Println(err.Error())
		}

		if !hasContent(xlsx, workingSheet, rowIndex+1) {
			break
		}

		rowIndex += 1
	}

	if err := generateIndex(contents); err != nil {
		log.Printf("Tartalomjegyzék generálása sikertelen! %s\n", err.Error())
	}
}

func generateIndex(items map[int]string) error {
	f, err := os.Create(fmt.Sprintf("%s/index.html", htmlPath))
	if err != nil {
		return errors.Wrap(err, "A tartalomjegyzék generálása sikertelen!")
	}

	if err = tpl.ExecuteTemplate(f, "index.gohtml", items); err != nil {
		return errors.Wrap(err, "A tartalomjegyzék generálása sikertelen!")
	}

	return nil
}

type TableRow struct {
	Label string
	Value string
}

func generateHtml(names, values []string, position string) error {
	id, err := strconv.Atoi(values[IdIndex])
	if err != nil {
		return errors.Wrapf(err, "A '%s' azonosítójú fájl generálása sikertelen. Az azonosítónak egész számnak kell lennie!", values[IdIndex])
	}

	f, err := os.Create(fmt.Sprintf("%s/%d.html", htmlPath, id))
	if err != nil {
		return errors.Wrapf(err, "A(z) '%d' azonosítójú sor alapján a HTML fájl generálása sikertelen!", id)
	}

	tableRows := []TableRow{}
	var imgUrl, imgAlt string

	for i, v := range names {
		htmlValue := values[i]
		if strings.HasPrefix(values[i], "pic|") {
			img := strings.Split(htmlValue, "|")
			imgAlt = img[1]
			imgUrl = img[2]
			htmlValue = "pic"
		}

		tableRows = append(tableRows, TableRow{v, htmlValue})
	}

	payload := struct {
		Rows     []TableRow
		Position string
		ImgUrl   string
		ImgAlt   string
	}{
		tableRows,
		position,
		imgUrl, imgAlt,
	}

	if err = tpl.ExecuteTemplate(f, "item.gohtml", payload); err != nil {
		return errors.Wrapf(err, "A(z) '%d' azonosítójú sor alapján a HTML fájl generálása sikertelen!", id)
	}

	return nil
}

func hasContent(xlsx *excelize.File, sheet string, row int) bool {
	if val, _ := getCellValue(xlsx, sheet, fmt.Sprintf("A%d", row)); val == "" {
		return false
	}

	return true
}

// A táblázat első sorát felolvasó eljárás. Konvenció szerint a táblázatnak az első sorának kell tartalmaznia a
// mezőneveket
func readHeader(xlsx *excelize.File, sheet string) ([]string, error) {
	return readRow(xlsx, sheet, 1)
}

// Felolvassa az adott táblázat, adott munkalapjának meghatározott sorát
func readRow(xlsx *excelize.File, sheet string, row int) ([]string, error) {
	var values []string

	for col := fromColumn; col <= toColumn; col++ {
		axis := fmt.Sprintf("%c%d", col, row)
		value, err := getCellValue(xlsx, sheet, axis)
		if err != nil {
			return nil, errors.Wrapf(err, "A '%s' munkalap %d. sorának felolvasása sikertelen", sheet, row)
		}

		values = append(values, value)
	}

	return values, nil
}

func getCellValue(xlsx *excelize.File, sheet, axis string) (string, error) {
	val, err := xlsx.GetCellValue(sheet, axis)
	if err != nil {
		return "", errors.Wrapf(err, "A '%s' munkalap '%s' cellájának felolvasása sikertelen!", sheet, axis)
	}

	return val, nil
}

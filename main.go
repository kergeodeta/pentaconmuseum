package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/pkg/errors"
	"html/template"
	"io/ioutil"
	"log"
	"os"
	"strconv"
)

const htmlPath = "./generated"

type Item struct {
	Id                  int    // A - Objektívek
	Name                string // B - Típus és fejléc és kép felugró szövege a HTML-ben
	Category            string // C - Kategória
	Manufactured        string // D - Gyártási időszak
	Focus               string // E - Fókusztávolság
	Aperture            string // F - Rekesz tartomány
	NearPoint           string // G - Közelpont
	LensCount           string // H - Lencsetagok
	LensGroup           string // I - Lencse csoportok
	ApertureLamells     string // J - Apertúra lamellák
	DiagonalViewAngle   string // K - Átlós látószög
	HorizontalViewAngle string // L - 36 x 24 mm látószög
	FilterDiameter      string // M - Szűrőátmérő
	Length              string // N - Hossza
	Diameter            string // O - Átmérője
	Weight              string // P - Tömege
	Socket              string // Q - Foglalat
	Manufacturer        string // R - Gyártó
	SerialNumber        string // S - Gyári szám
	Comment             string // T - Megjegyzés
	PicturePath         string // U - Szerkezeti ábra elérési útvonala
	Position            string
}

func (i Item) generateHtml() error {
	f, err := os.Create(fmt.Sprintf("%s/%d.html", htmlPath, i.Id))
	if err != nil {
		return errors.Wrapf(err, "A(z) '%d' azonosítójú sor alapján a HTML fájl generálása sikertelen!", i.Id)
	}

	if err = tpl.Execute(f, i); err != nil {
		return errors.Wrapf(err, "A(z) '%d' azonosítójú sor alapján a HTML fájl generálása sikertelen!", i.Id)
	}

	return nil
}

var spreadsheetPath *string
var skipRows *int
var tpl *template.Template

func init() {
	tplData, err := ioutil.ReadFile("item.gohtml")
	if err != nil {
		log.Fatalf("HTML sablon betöltése sikertelen! %s\n", err.Error())
	}

	tpl, err = template.New("museum").Parse(string(tplData))
	if err != nil {
		log.Fatalf("HTML sablon betöltése sikertelen! %s\n", err.Error())
	}

	if _, err = os.Stat(htmlPath); os.IsNotExist(err) {
		if err := os.Mkdir(htmlPath, 0755); err != nil {
			log.Fatalf("Generált HTML fájlok mappájának létrehozása sikertelen! %s\n", err.Error())
		}
	}

	spreadsheetPath = flag.String("in", "", "Bemeneti adatokat tartalmazó XLSX fájl elérési útvonala")
	skipRows = flag.Int("skipRows", 1, "Megadja, hogy az első hány sort ne olvassa fel")
	flag.Parse()
}

func main() {
	xlsx, err := excelize.OpenFile(*spreadsheetPath)
	if err != nil {
		log.Fatalf("A megadott XLSX fájl megnyitása (%s) sikertelen! %s\n", *spreadsheetPath, err.Error())
	}

	sheets := xlsx.GetSheetMap()
	workingSheet := sheets[1]
	row := *skipRows + 1
	index := make(map[int]string)
	for {
		item := getItemFrom(xlsx, workingSheet, row)
		if row == *skipRows+1 {
			item.Position = "first"
		}

		if isLast(xlsx, workingSheet, row+1) {
			item.Position = "last"
		}

		if err := item.generateHtml(); err != nil {
			log.Println(err.Error())
		}

		index[item.Id] = item.Name

		if isLast(xlsx, workingSheet, row+1) {
			break
		}
		row += 1
	}

	if err := generateIndex(index); err != nil {
		log.Printf("Tartalomjegyzék generálása sikertelen! %s\n", err.Error())
	}
}

func getItemFrom(xlsx *excelize.File, sheet string, row int) *Item {
	id, err := strconv.Atoi(getCellValue(xlsx, sheet, fmt.Sprintf("A%d", row)))
	if err != nil {
		log.Println("A %d sor értelmezése sikertelen! Helytelen ID formátum!")
		return &Item{}
	}

	return &Item{
		Id:                  id,
		Name:                getCellValue(xlsx, sheet, fmt.Sprintf("B%d", row)),
		Category:            getCellValue(xlsx, sheet, fmt.Sprintf("C%d", row)),
		Manufactured:        getCellValue(xlsx, sheet, fmt.Sprintf("D%d", row)),
		Focus:               getCellValue(xlsx, sheet, fmt.Sprintf("E%d", row)),
		Aperture:            getCellValue(xlsx, sheet, fmt.Sprintf("F%d", row)),
		NearPoint:           getCellValue(xlsx, sheet, fmt.Sprintf("G%d", row)),
		LensCount:           getCellValue(xlsx, sheet, fmt.Sprintf("H%d", row)),
		LensGroup:           getCellValue(xlsx, sheet, fmt.Sprintf("I%d", row)),
		ApertureLamells:     getCellValue(xlsx, sheet, fmt.Sprintf("J%d", row)),
		DiagonalViewAngle:   getCellValue(xlsx, sheet, fmt.Sprintf("K%d", row)),
		HorizontalViewAngle: getCellValue(xlsx, sheet, fmt.Sprintf("L%d", row)),
		FilterDiameter:      getCellValue(xlsx, sheet, fmt.Sprintf("M%d", row)),
		Length:              getCellValue(xlsx, sheet, fmt.Sprintf("N%d", row)),
		Diameter:            getCellValue(xlsx, sheet, fmt.Sprintf("O%d", row)),
		Weight:              getCellValue(xlsx, sheet, fmt.Sprintf("P%d", row)),
		Socket:              getCellValue(xlsx, sheet, fmt.Sprintf("Q%d", row)),
		Manufacturer:        getCellValue(xlsx, sheet, fmt.Sprintf("R%d", row)),
		SerialNumber:        getCellValue(xlsx, sheet, fmt.Sprintf("S%d", row)),
		Comment:             getCellValue(xlsx, sheet, fmt.Sprintf("T%d", row)),
		PicturePath:         getCellValue(xlsx, sheet, fmt.Sprintf("U%d", row)),
	}
}

func isLast(xlsx *excelize.File, sheet string, row int) bool {
	if getCellValue(xlsx, sheet, fmt.Sprintf("A%d", row)) == "" {
		return true
	}

	return false
}

func getCellValue(xlsx *excelize.File, sheet, axis string) string {
	val, err := xlsx.GetCellValue(sheet, axis)
	if err != nil {
		log.Printf("A %s cella értékének kiolvasása sikertelen! %s\n", axis, err.Error())
	}

	return val
}

func generateIndex(items map[int]string) error {
	f, err := os.Create(fmt.Sprintf("%s/index.html", htmlPath))
	if err != nil {
		return errors.Wrap(err, "A tartalomjegyzék generálása sikertelen!")
	}

	indexTpl := template.Must(template.ParseFiles("index.gohtml"))
	if err = indexTpl.Execute(f, items); err != nil {
		return errors.Wrap(err, "A tartalomjegyzék generálása sikertelen!")
	}

	return nil
}
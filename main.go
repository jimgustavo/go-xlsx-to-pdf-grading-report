package main

import (
	"fmt"
	"log"
	"time"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

// Count the filled rows in the students list
func countFilledRows(f *excelize.File) (int, error) {
	// Find the last non-empty row in column F
	var numRows int
	for i := 2; i <= 50; i++ { // Starting from F2 to the last possible row, to maximum 50
		cellValue, err := f.GetCellValue("DATA", fmt.Sprintf("F%d", i))
		if err != nil {
			return 0, err
		}
		if cellValue == "" {
			break
		}
		numRows++
	}

	return numRows, nil
}

func main() {
	// Open the Excel file
	f, err := excelize.OpenFile("consolidado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	//Get the number of filled rows
	numRows, err := countFilledRows(f)
	if err != nil {
		log.Fatal(err)
	}
	//fmt.Println("Number of filled rows:", numRows)   //Print the number of filled rows in column F to the console

	// Create a new PDF
	pdf := gofpdf.New("L", "mm", "A4", "")

	// Get data from specific cells in the DATA sheet
	dataSheet := "DATA"
	var institution, class, tutor_teacher, school_year, period, workday, city string

	institution, err = f.GetCellValue(dataSheet, "B2")
	if err != nil {
		log.Fatal(err)
	}
	class, err = f.GetCellValue(dataSheet, "B3")
	if err != nil {
		log.Fatal(err)
	}
	tutor_teacher, err = f.GetCellValue(dataSheet, "B5")
	if err != nil {
		log.Fatal(err)
	}
	school_year, err = f.GetCellValue(dataSheet, "B6")
	if err != nil {
		log.Fatal(err)
	}
	period, err = f.GetCellValue(dataSheet, "B7")
	if err != nil {
		log.Fatal(err)
	}
	workday, err = f.GetCellValue(dataSheet, "B8")
	if err != nil {
		log.Fatal(err)
	}
	city, err = f.GetCellValue(dataSheet, "B9")
	if err != nil {
		log.Fatal(err)
	}

	// Loop through each student
	rows, err := f.GetRows("DATA")
	if err != nil {
		log.Fatal(err)
	}

	// Get the current date and time
	currentTime := time.Now()
	// Truncate the time to seconds
	truncatedTime := currentTime.Truncate(time.Second)

	for i, row := range rows[1 : numRows+1] { // Assuming student list starts from row F2 to the last filled row.
		// Extract student name
		studentName := row[5] //// Assuming student list starts in F2 row or 5 row

		// Add a new page for each student
		pdf.AddPage()

		// Add logo image
		logoPath := "ue12f_logo.jpeg"
		pdf.Image(logoPath, 5, 5, 20, 0, false, "", 0, "ue12f_logo")

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, institution, "0", 0, "C", false, 0, "")
		pdf.Ln(5)

		pdf.SetFont("Arial", "", 9)
		// Add title
		pdf.CellFormat(280, 10, city, "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, "Reporte de Calificaciones del "+period, "0", 0, "C", false, 0, "")
		pdf.Ln(20)

		pdf.SetFont("Arial", "", 10)
		// Write specific data from the DATA sheet
		pdf.Cell(40, 10, "Curso: "+class)
		pdf.Cell(60, 10, "Modalidad: "+workday)
		pdf.Cell(60, 10, "Docente tutor: "+tutor_teacher)
		pdf.Cell(60, 10, "Periodo: "+school_year)
		pdf.Ln(10)

		// Write student name to PDF
		pdf.Cell(40, 10, "Estudiante: "+studentName)
		pdf.Ln(10)

		// Extract math grades
		mathGrades, err := f.GetRows("Matematicas")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		languageGrades, err := f.GetRows("Lenguaje")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		scienceGrades, err := f.GetRows("CCNN")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		social_studiesGrades, err := f.GetRows("EESS")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		englishGrades, err := f.GetRows("Ingles")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		physical_cultureGrades, err := f.GetRows("EEFF")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		art_cultureGrades, err := f.GetRows("ECA")
		if err != nil {
			log.Fatal(err)
		}

		// Write the labels
		pdf.SetFont("Helvetica", "", 9) // Set font and size for the table
		pdf.SetFillColor(211, 211, 211) // Set fill color to light gray (RGB: 211, 211, 211)
		pdf.CellFormat(50, 5, "", "", 0, "L", false, 0, "")
		pdf.CellFormat(82, 5, "Primer Bimestre", "0", 0, "C", true, 0, "")
		pdf.CellFormat(2, 5, "", "", 0, "L", false, 0, "")
		pdf.CellFormat(82, 5, "Segundo Bimestre", "0", 0, "C", true, 0, "")
		pdf.Ln(5)
		pdf.Cell(50, 10, "")
		pdf.Cell(12, 10, "P1")
		pdf.Cell(12, 10, "P2")
		pdf.Cell(12, 10, "Pro")
		pdf.Cell(12, 10, "Pro-80%")
		pdf.Cell(12, 10, "Ex")
		pdf.Cell(12, 10, "Ex-20%")
		pdf.Cell(12, 10, "1merB")

		pdf.Cell(12, 10, "P1")
		pdf.Cell(12, 10, "P2")
		pdf.Cell(12, 10, "Pro")
		pdf.Cell(12, 10, "Pro-80%")
		pdf.Cell(12, 10, "Ex")
		pdf.Cell(12, 10, "Ex-20%")
		pdf.Cell(12, 10, "2doB")

		pdf.Cell(12, 10, "Anual")
		pdf.Cell(12, 10, "Supl")

		pdf.Cell(12, 10, "Falt")
		pdf.Cell(12, 10, "Comp")
		pdf.Ln(10)

		// Write math grades
		pdf.Cell(50, 10, "MATEMATICAS:")
		pdf.Cell(12, 10, mathGrades[i+6][2])
		pdf.Cell(12, 10, mathGrades[i+6][3])
		pdf.Cell(12, 10, mathGrades[i+6][4])
		pdf.Cell(12, 10, mathGrades[i+6][5])
		pdf.Cell(12, 10, mathGrades[i+6][6])
		pdf.Cell(12, 10, mathGrades[i+6][7])
		pdf.Cell(12, 10, mathGrades[i+6][8])
		pdf.Cell(12, 10, mathGrades[i+6][9])
		pdf.Cell(12, 10, mathGrades[i+6][10])
		pdf.Cell(12, 10, mathGrades[i+6][11])
		pdf.Cell(12, 10, mathGrades[i+6][12])
		pdf.Cell(12, 10, mathGrades[i+6][13])
		pdf.Cell(12, 10, mathGrades[i+6][14])
		pdf.Cell(12, 10, mathGrades[i+6][15])
		pdf.Cell(12, 10, mathGrades[i+6][16])
		pdf.Cell(12, 10, mathGrades[i+6][17])
		pdf.Cell(12, 10, mathGrades[i+6][18])
		pdf.Cell(12, 10, mathGrades[i+6][19])
		pdf.Ln(5)

		// Write language grades
		pdf.Cell(50, 10, "LENGUAJE:")
		pdf.Cell(12, 10, languageGrades[i+6][2])
		pdf.Cell(12, 10, languageGrades[i+6][3])
		pdf.Cell(12, 10, languageGrades[i+6][4])
		pdf.Cell(12, 10, languageGrades[i+6][5])
		pdf.Cell(12, 10, languageGrades[i+6][6])
		pdf.Cell(12, 10, languageGrades[i+6][7])
		pdf.Cell(12, 10, languageGrades[i+6][8])
		pdf.Cell(12, 10, languageGrades[i+6][9])
		pdf.Cell(12, 10, languageGrades[i+6][10])
		pdf.Cell(12, 10, languageGrades[i+6][11])
		pdf.Cell(12, 10, languageGrades[i+6][12])
		pdf.Cell(12, 10, languageGrades[i+6][13])
		pdf.Cell(12, 10, languageGrades[i+6][14])
		pdf.Cell(12, 10, languageGrades[i+6][15])
		pdf.Cell(12, 10, languageGrades[i+6][16])
		pdf.Cell(12, 10, languageGrades[i+6][17])
		pdf.Cell(12, 10, languageGrades[i+6][18])
		pdf.Cell(12, 10, languageGrades[i+6][19])
		pdf.Ln(5)

		// Write science grades
		pdf.Cell(50, 10, "CIENCIAS NATURALES:")
		pdf.Cell(12, 10, scienceGrades[i+6][2])
		pdf.Cell(12, 10, scienceGrades[i+6][3])
		pdf.Cell(12, 10, scienceGrades[i+6][4])
		pdf.Cell(12, 10, scienceGrades[i+6][5])
		pdf.Cell(12, 10, scienceGrades[i+6][6])
		pdf.Cell(12, 10, scienceGrades[i+6][7])
		pdf.Cell(12, 10, scienceGrades[i+6][8])
		pdf.Cell(12, 10, scienceGrades[i+6][9])
		pdf.Cell(12, 10, scienceGrades[i+6][10])
		pdf.Cell(12, 10, scienceGrades[i+6][11])
		pdf.Cell(12, 10, scienceGrades[i+6][12])
		pdf.Cell(12, 10, scienceGrades[i+6][13])
		pdf.Cell(12, 10, scienceGrades[i+6][14])
		pdf.Cell(12, 10, scienceGrades[i+6][15])
		pdf.Cell(12, 10, scienceGrades[i+6][16])
		pdf.Cell(12, 10, scienceGrades[i+6][17])
		pdf.Cell(12, 10, scienceGrades[i+6][18])
		pdf.Cell(12, 10, scienceGrades[i+6][19])
		pdf.Ln(5)

		// Write social studies grades
		pdf.Cell(50, 10, "ESTUDIOS SOCIALES:")
		pdf.Cell(12, 10, social_studiesGrades[i+6][2])
		pdf.Cell(12, 10, social_studiesGrades[i+6][3])
		pdf.Cell(12, 10, social_studiesGrades[i+6][4])
		pdf.Cell(12, 10, social_studiesGrades[i+6][5])
		pdf.Cell(12, 10, social_studiesGrades[i+6][6])
		pdf.Cell(12, 10, social_studiesGrades[i+6][7])
		pdf.Cell(12, 10, social_studiesGrades[i+6][8])
		pdf.Cell(12, 10, social_studiesGrades[i+6][9])
		pdf.Cell(12, 10, social_studiesGrades[i+6][10])
		pdf.Cell(12, 10, social_studiesGrades[i+6][11])
		pdf.Cell(12, 10, social_studiesGrades[i+6][12])
		pdf.Cell(12, 10, social_studiesGrades[i+6][13])
		pdf.Cell(12, 10, social_studiesGrades[i+6][14])
		pdf.Cell(12, 10, social_studiesGrades[i+6][15])
		pdf.Cell(12, 10, social_studiesGrades[i+6][16])
		pdf.Cell(12, 10, social_studiesGrades[i+6][17])
		pdf.Cell(12, 10, social_studiesGrades[i+6][18])
		pdf.Cell(12, 10, social_studiesGrades[i+6][19])
		pdf.Ln(5)

		// Write english grades
		pdf.Cell(50, 10, "INGLES:")
		pdf.Cell(12, 10, englishGrades[i+6][2])
		pdf.Cell(12, 10, englishGrades[i+6][3])
		pdf.Cell(12, 10, englishGrades[i+6][4])
		pdf.Cell(12, 10, englishGrades[i+6][5])
		pdf.Cell(12, 10, englishGrades[i+6][6])
		pdf.Cell(12, 10, englishGrades[i+6][7])
		pdf.Cell(12, 10, englishGrades[i+6][8])
		pdf.Cell(12, 10, englishGrades[i+6][9])
		pdf.Cell(12, 10, englishGrades[i+6][10])
		pdf.Cell(12, 10, englishGrades[i+6][11])
		pdf.Cell(12, 10, englishGrades[i+6][12])
		pdf.Cell(12, 10, englishGrades[i+6][13])
		pdf.Cell(12, 10, englishGrades[i+6][14])
		pdf.Cell(12, 10, englishGrades[i+6][15])
		pdf.Cell(12, 10, englishGrades[i+6][16])
		pdf.Cell(12, 10, englishGrades[i+6][17])
		pdf.Cell(12, 10, englishGrades[i+6][18])
		pdf.Cell(12, 10, englishGrades[i+6][19])
		pdf.Ln(5)

		// Write physical culture grades
		pdf.Cell(50, 10, "CULTURA FISICA:")
		pdf.Cell(12, 10, physical_cultureGrades[i+6][2])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][3])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][4])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][5])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][6])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][7])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][8])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][9])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][10])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][11])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][12])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][13])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][14])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][15])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][16])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][17])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][18])
		pdf.Cell(12, 10, physical_cultureGrades[i+6][19])
		pdf.Ln(5)

		// Write art culture grades
		pdf.Cell(50, 10, "ECA:")
		pdf.Cell(12, 10, art_cultureGrades[i+6][2])
		pdf.Cell(12, 10, art_cultureGrades[i+6][3])
		pdf.Cell(12, 10, art_cultureGrades[i+6][4])
		pdf.Cell(12, 10, art_cultureGrades[i+6][5])
		pdf.Cell(12, 10, art_cultureGrades[i+6][6])
		pdf.Cell(12, 10, art_cultureGrades[i+6][7])
		pdf.Cell(12, 10, art_cultureGrades[i+6][8])
		pdf.Cell(12, 10, art_cultureGrades[i+6][9])
		pdf.Cell(12, 10, art_cultureGrades[i+6][10])
		pdf.Cell(12, 10, art_cultureGrades[i+6][11])
		pdf.Cell(12, 10, art_cultureGrades[i+6][12])
		pdf.Cell(12, 10, art_cultureGrades[i+6][13])
		pdf.Cell(12, 10, art_cultureGrades[i+6][14])
		pdf.Cell(12, 10, art_cultureGrades[i+6][15])
		pdf.Cell(12, 10, art_cultureGrades[i+6][16])
		pdf.Cell(12, 10, art_cultureGrades[i+6][17])
		pdf.Cell(12, 10, art_cultureGrades[i+6][18])
		pdf.Cell(12, 10, art_cultureGrades[i+6][19])
		pdf.Ln(5)

		// Set font and size for the report closing section
		pdf.SetFont("Arial", "", 10)

		// Add Date and Time
		pdf.Cell(40, 45, "Fecha: "+truncatedTime.Local().String())
		pdf.Ln(10)

		// Add teacher signature
		pdf.CellFormat(40, 45, "________________", "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "________________", "0", 0, "C", false, 0, "")  //In case it'needed authority signature
		pdf.Ln(5)
		pdf.CellFormat(40, 45, tutor_teacher, "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "Authority Name", "0", 0, "C", false, 0, "")    //In case it'needed authority signature
		pdf.Ln(5)
		pdf.CellFormat(40, 45, "Docente Tutor", "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "Authority", "0", 0, "C", false, 0, "")         //In case it'needed authority signature
	}

	// Save PDF to files
	err = pdf.OutputFileAndClose("grading_report.pdf")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Report generated successfully.")
}

//------------For Linux-------------------//
//GOOS=windows GOARCH=amd64 go build -o reportes-individuales    //Replace linux with windows or darwin depending on the target platform. Replace amd64 with other architectures if needed.

//------------For Windows----------------//
//set GOOS=windows
//set GOARCH=amd64
//go build -o reportes-individuales.exe

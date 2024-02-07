package main

import (
	"fmt"
	"log"
	"time"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Open the Excel file
	f, err := excelize.OpenFile("consolidado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Create a new PDF
	pdf := gofpdf.New("L", "mm", "A4", "")

	// Get data from specific cells in the DATA sheet
	dataSheet := "DATA"
	var institution, class, subject, teacher, school_year, period, workday string

	institution, err = f.GetCellValue(dataSheet, "B2")
	if err != nil {
		log.Fatal(err)
	}
	class, err = f.GetCellValue(dataSheet, "B3")
	if err != nil {
		log.Fatal(err)
	}
	subject, err = f.GetCellValue(dataSheet, "B4")
	if err != nil {
		log.Fatal(err)
	}
	teacher, err = f.GetCellValue(dataSheet, "B5")
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

	// Loop through each student
	rows, err := f.GetRows("DATA")
	if err != nil {
		log.Fatal(err)
	}

	// Get the current date and time
	currentTime := time.Now()
	// Truncate the time to seconds
	truncatedTime := currentTime.Truncate(time.Second)

	for i, row := range rows[10:30] { // Assuming student list starts from row 11
		// Extract student name
		studentName := row[1]

		// Add a new page for each student
		pdf.AddPage()
		pdf.SetFont("Arial", "", 10)

		// Add logo image
		logoPath := "ue12f_logo.jpeg"
		pdf.Image(logoPath, 5, 5, 20, 0, false, "", 0, "ue12f_logo")

		// Add title
		pdf.CellFormat(280, 10, institution, "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		// Add title
		pdf.CellFormat(280, 10, "Grades Report -"+period, "0", 0, "C", false, 0, "")
		pdf.Ln(20)

		// Write specific data from the DATA sheet
		pdf.Cell(40, 10, "Class: "+class)
		pdf.Cell(40, 10, "Subject: "+subject)
		pdf.Cell(60, 10, "Teacher: "+teacher)
		pdf.Cell(60, 10, "School Year: "+school_year)
		pdf.Cell(60, 10, "Workday-Modality: "+workday)
		pdf.Ln(10)

		// Write student name to PDF
		pdf.Cell(40, 10, "Student: "+studentName)
		pdf.Ln(10)

		// Extract math grades
		mathGrades, err := f.GetRows("math")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		scienceGrades, err := f.GetRows("science")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		social_studiesGrades, err := f.GetRows("social_studies")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		languageGrades, err := f.GetRows("language")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		englishGrades, err := f.GetRows("english")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		physical_cultureGrades, err := f.GetRows("physical_culture")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		art_cultureGrades, err := f.GetRows("art_culture")
		if err != nil {
			log.Fatal(err)
		}

		// Write the labels
		pdf.SetFont("Arial", "", 9)     // Set font and size for the table
		pdf.SetFillColor(211, 211, 211) // Set fill color to light gray (RGB: 211, 211, 211)
		pdf.CellFormat(50, 5, "", "", 0, "L", false, 0, "")
		pdf.CellFormat(82, 5, "First Bimester", "0", 0, "C", true, 0, "")
		pdf.CellFormat(2, 5, "", "", 0, "L", false, 0, "")
		pdf.CellFormat(82, 5, "Second Bimester", "0", 0, "C", true, 0, "")
		pdf.Ln(5)
		pdf.Cell(50, 10, "")
		pdf.Cell(12, 10, "Term1")
		pdf.Cell(12, 10, "Term2")
		pdf.Cell(12, 10, "Aver1")
		pdf.Cell(12, 10, "Av-80%")
		pdf.Cell(12, 10, "Ex")
		pdf.Cell(12, 10, "Ex-20%")
		pdf.Cell(12, 10, "1merB")

		pdf.Cell(12, 10, "Term1")
		pdf.Cell(12, 10, "Term2")
		pdf.Cell(12, 10, "Aver2")
		pdf.Cell(12, 10, "Av-80%")
		pdf.Cell(12, 10, "Ex")
		pdf.Cell(12, 10, "Ex-20%")
		pdf.Cell(12, 10, "2doB")

		pdf.Cell(12, 10, "T. Year")

		pdf.Cell(12, 10, "Abse")
		pdf.Cell(12, 10, "Beha")
		pdf.Ln(10)

		// Write math grades
		pdf.Cell(50, 10, "Math Grades:")
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
		pdf.Ln(5)

		// Write science grades
		pdf.Cell(50, 10, "Science Grades:")
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
		pdf.Ln(5)

		// Write social studies grades
		pdf.Cell(50, 10, "Social studies Grades:")
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
		pdf.Ln(5)

		// Write language grades
		pdf.Cell(50, 10, "Language Grades:")
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
		pdf.Ln(5)

		// Write english grades
		pdf.Cell(50, 10, "English Grades:")
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
		pdf.Ln(5)

		// Write physical culture grades
		pdf.Cell(50, 10, "Physical culture Grades:")
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
		pdf.Ln(5)

		// Write art culture grades
		pdf.Cell(50, 10, "Art culture Grades:")
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
		pdf.Ln(5)

		// Set font and size for the report closing section
		pdf.SetFont("Arial", "", 10)

		// Add Date and Time
		pdf.Cell(40, 50, "Date: "+truncatedTime.Local().String())
		pdf.Ln(10)

		// Add teacher signature
		pdf.CellFormat(40, 50, "________________", "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "________________", "0", 0, "C", false, 0, "")  //In case it'needed authority signature
		pdf.Ln(5)
		pdf.CellFormat(40, 50, teacher, "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "Authority Name", "0", 0, "C", false, 0, "")    //In case it'needed authority signature
		pdf.Ln(5)
		pdf.CellFormat(40, 50, "Tutor Teacher", "0", 0, "C", false, 0, "")
		//pdf.CellFormat(40, 50, "Authority", "0", 0, "C", false, 0, "")         //In case it'needed authority signature
	}

	// Save PDF to files
	err = pdf.OutputFileAndClose("grading_report.pdf")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Report generated successfully.")
}

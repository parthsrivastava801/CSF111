package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	EmplID         string
	CampusID       string
	QuizScore      float64
	MidSemScore    float64
	LabTestScore   float64
	WeeklyLabScore float64
	PreCompreScore float64
	CompreScore    float64
	TotalScore     float64
}

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Error: No file path provided. Usage: go run main.go <path_to_xlsx>")
	}
	filePath := os.Args[1]

	// to open the provided Excel file
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open the file: %s. Ensure it exists and is in a valid format. Error: %v", filePath, err)
	}

	sheetName := file.GetSheetName(0) // for getting the first sheet
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Error reading the sheet. Please check if the file is properly formatted. Error: %v", err)
	}
    var students []Student
	var totalSum, quizSum, midSemSum, labTestSum, weeklyLabSum, preCompreSum, compreSum float64
	branchWiseScores := make(map[string][]float64) // to store total scores for each branch
	headerSkipped := false                         // flag to skip header row

	for _, row := range rows {
		if !headerSkipped {
			headerSkipped = true
			continue // Skip the header row
		}

		// to ensure there are enough columns in the row
		if len(row) < 11 {
			continue // skipping incomplete rows
		}

		// parsing marks from the relevant columns
		quizScore, _ := strconv.ParseFloat(row[4], 64)
		midSemScore, _ := strconv.ParseFloat(row[5], 64)
		labTestScore, _ := strconv.ParseFloat(row[6], 64)
		weeklyLabScore, _ := strconv.ParseFloat(row[7], 64)
		preCompreScore, _ := strconv.ParseFloat(row[8], 64)
		compreScore, _ := strconv.ParseFloat(row[9], 64)
		totalScore, _ := strconv.ParseFloat(row[10], 64)
        // Validate computed total against provided total
		computedTotal := quizScore + midSemScore + labTestScore + weeklyLabScore + compreScore
		if totalScore != computedTotal {
			log.Printf("Mismatch detected for EmplID %s. Expected: %.2f, Found: %.2f.", row[2], computedTotal, totalScore)
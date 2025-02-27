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
        // validating computed total against provided total
		computedTotal := quizScore + midSemScore + labTestScore + weeklyLabScore + compreScore
		if totalScore != computedTotal {
			log.Printf("Mismatch detected for EmplID %s. Expected: %.2f, Found: %.2f.", row[2], computedTotal, totalScore)
		}

		// storing student details
		students = append(students, Student{
			EmplID:         row[2],
			CampusID:       row[3],
			QuizScore:      quizScore,
			MidSemScore:    midSemScore,
			LabTestScore:   labTestScore,
			WeeklyLabScore: weeklyLabScore,
			PreCompreScore: preCompreScore,
			CompreScore:    compreScore,
			TotalScore:     totalScore,
		})

		// accumulating total marks for general averages
		totalSum += totalScore
		quizSum += quizScore
		midSemSum += midSemScore
		labTestSum += labTestScore
		weeklyLabSum += weeklyLabScore
		preCompreSum += preCompreScore
		compreSum += compreScore

		// extracting branchwise data for 2024 batch
		if len(row[3]) >= 6 && row[3][:4] == "2024" {
			branchCode := row[3][4:6] // extracting branch code
			branchWiseScores[branchCode] = append(branchWiseScores[branchCode], totalScore)
		}
	}
	totalStudents := float64(len(students))


	// to print general average scores
	fmt.Println("General Averages:")
	fmt.Printf("Quiz: %.2f\n", quizSum/totalStudents)
	fmt.Printf("Mid-Sem: %.2f\n", midSemSum/totalStudents)
	fmt.Printf("Lab Test: %.2f\n", labTestSum/totalStudents)
	fmt.Printf("Weekly Labs: %.2f\n", weeklyLabSum/totalStudents)
	fmt.Printf("Pre-Compre: %.2f\n", preCompreSum/totalStudents)
	fmt.Printf("Compre: %.2f\n", compreSum/totalStudents)
	fmt.Printf("Overall Total: %.2f\n", totalSum/totalStudents)

	// to print branch-wise average scores
	fmt.Println("\nBranch-wise Averages (2024 Batch):")
	for branch, scores := range branchWiseScores {
		branchTotal := 0.0
		for _, score := range scores {
			branchTotal += score
		}
		fmt.Printf("Branch %s: %.2f\n", branch, branchTotal/float64(len(scores)))
	}

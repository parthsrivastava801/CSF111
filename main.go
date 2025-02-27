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
	EmplID         string  // Unique identifier for the student
	CampusID       string  // College-issued campus ID (contains batch & branch info)
	QuizScore      float64 // Marks obtained in Quiz
	MidSemScore    float64 // Marks obtained in Mid-Semester Exam
	LabTestScore   float64 // Marks obtained in Lab Test
	WeeklyLabScore float64 // Marks obtained in Weekly Labs
	PreCompreScore float64 // Marks obtained in Pre-Comprehensive Exam
	CompreScore    float64 // Marks obtained in Comprehensive Exam
	TotalScore     float64 // Total marks as given in the sheet
}

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Error: No file path provided. Usage: go run main.go <path_to_xlsx>")
	}
	filePath := os.Args[1]

	// Open the provided Excel file
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open the file: %s. Ensure the file exists and is in a valid format. Error: %v", filePath, err)
	}

	sheetName := file.GetSheetName(0) // Fetch the first sheet
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Error reading the sheet. Please check if the file is properly formatted. Error: %v", err)
	}

	var students []Student
	var totalSum, quizSum, midSemSum, labTestSum, weeklyLabSum, preCompreSum, compreSum float64
	branchWiseScores := make(map[string][]float64) // Stores total scores for each branch
	headerSkipped := false                         // Flag to skip header row

	for _, row := range rows {
		if !headerSkipped {
			headerSkipped = true
			continue // Skip the header row
		}

		// Ensure there are enough columns in the row
		if len(row) < 11 {
			continue // Skip incomplete rows
		}

		// Parse marks from the relevant columns
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
		}

		// Store student details
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

		// Accumulate total marks for general averages
		totalSum += totalScore
		quizSum += quizScore
		midSemSum += midSemScore
		labTestSum += labTestScore
		weeklyLabSum += weeklyLabScore
		preCompreSum += preCompreScore
		compreSum += compreScore

		// Extract branch-wise data for 2024 batch
		if len(row[3]) >= 6 && row[3][:4] == "2024" {
			branchCode := row[3][4:6] // Extracting branch code
			branchWiseScores[branchCode] = append(branchWiseScores[branchCode], totalScore)
		}
	}

	totalStudents := float64(len(students))

	// Print general average scores
	fmt.Println("General Averages:")
	fmt.Printf("Quiz: %.2f\n", quizSum/totalStudents)
	fmt.Printf("Mid-Sem: %.2f\n", midSemSum/totalStudents)
	fmt.Printf("Lab Test: %.2f\n", labTestSum/totalStudents)
	fmt.Printf("Weekly Labs: %.2f\n", weeklyLabSum/totalStudents)
	fmt.Printf("Pre-Compre: %.2f\n", preCompreSum/totalStudents)
	fmt.Printf("Compre: %.2f\n", compreSum/totalStudents)
	fmt.Printf("Overall Total: %.2f\n", totalSum/totalStudents)

	// Print branch-wise average scores
	fmt.Println("\nBranch-wise Averages (2024 Batch):")
	for branch, scores := range branchWiseScores {
		branchTotal := 0.0
		for _, score := range scores {
			branchTotal += score
		}
		fmt.Printf("Branch %s: %.2f\n", branch, branchTotal/float64(len(scores)))
	}

	// Identify and display top 3 students for each category
	fmt.Println("\nTop 3 Students for Each Category:")
	categories := map[string]func(Student) float64{
		"Quiz":        func(s Student) float64 { return s.QuizScore },
		"Mid-Sem":     func(s Student) float64 { return s.MidSemScore },
		"Lab Test":    func(s Student) float64 { return s.LabTestScore },
		"Weekly Labs": func(s Student) float64 { return s.WeeklyLabScore },
		"Pre-Compre":  func(s Student) float64 { return s.PreCompreScore },
		"Compre":      func(s Student) float64 { return s.CompreScore },
		"Total":       func(s Student) float64 { return s.TotalScore },
	}

	for category, scoreExtractor := range categories {
		fmt.Printf("\n%s:\n", category)
		sort.Slice(students, func(i, j int) bool {
			return scoreExtractor(students[i]) > scoreExtractor(students[j])
		})
		for i := 0; i < 3 && i < len(students); i++ {
			fmt.Printf("%d. EmplID: %s, Marks: %.2f\n", i+1, students[i].EmplID, scoreExtractor(students[i]))
		}
	}
}

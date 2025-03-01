package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"

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

// JSON export structures
type SummaryReport struct {
	GeneralAverages CategoryAverages            `json:"generalAverages"`
	BranchAverages  map[string]float64          `json:"branchAverages"`
	TopStudents     map[string][]StudentRanking `json:"topStudents"`
	Discrepancies   []Discrepancy               `json:"discrepancies"`
}

type CategoryAverages struct {
	Quiz       float64 `json:"quiz"`
	MidSem     float64 `json:"midSem"`
	LabTest    float64 `json:"labTest"`
	WeeklyLabs float64 `json:"weeklyLabs"`
	PreCompre  float64 `json:"preCompre"`
	Compre     float64 `json:"compre"`
	Total      float64 `json:"total"`
}

type StudentRanking struct {
	EmplID string  `json:"emplID"`
	Marks  float64 `json:"marks"`
	Rank   int     `json:"rank"`
}

type Discrepancy struct {
	EmplID        string  `json:"emplID"`
	ExpectedTotal float64 `json:"expectedTotal"`
	ActualTotal   float64 `json:"actualTotal"`
}

func main() {
	// to add command-line flag for JSON export
	exportFlag := flag.String("export", "", "Export format (json)")
	flag.Parse()
	args := flag.Args()

	if len(args) < 1 {
		log.Fatal("Error: No file path provided. Usage: go run main.go [--export=json] <path_to_xlsx>")
	}
	filePath := args[0]

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

	// for JSON export
	var discrepancies []Discrepancy

	for _, row := range rows {
		if !headerSkipped {
			headerSkipped = true
			continue // skipping the header row
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
			// to store discrepancy for JSON export
			discrepancies = append(discrepancies, Discrepancy{
				EmplID:        row[2],
				ExpectedTotal: computedTotal,
				ActualTotal:   totalScore,
			})
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

	// calculating branch-wise averages for JSON export
	branchAverages := make(map[string]float64)
	for branch, scores := range branchWiseScores {
		branchTotal := 0.0
		for _, score := range scores {
			branchTotal += score
		}
		branchAverages[branch] = branchTotal / float64(len(scores))
	}

	// Categories for Top students
	categories := map[string]func(Student) float64{
		"Quiz":        func(s Student) float64 { return s.QuizScore },
		"Mid-Sem":     func(s Student) float64 { return s.MidSemScore },
		"Lab Test":    func(s Student) float64 { return s.LabTestScore },
		"Weekly Labs": func(s Student) float64 { return s.WeeklyLabScore },
		"Pre-Compre":  func(s Student) float64 { return s.PreCompreScore },
		"Compre":      func(s Student) float64 { return s.CompreScore },
		"Total":       func(s Student) float64 { return s.TotalScore },
	}

	// storing top students for JSON export
	topStudents := make(map[string][]StudentRanking)

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
	// identifying and display top 3 students for each category
	fmt.Println("\nTop 3 Students for Each Category:")

	for category, scoreExtractor := range categories {
		fmt.Printf("\n%s:\n", category)
		sort.Slice(students, func(i, j int) bool {
			return scoreExtractor(students[i]) > scoreExtractor(students[j])
		})

		categoryTopStudents := []StudentRanking{}
		for i := 0; i < 3 && i < len(students); i++ {
			fmt.Printf("%d. EmplID: %s, Marks: %.2f\n", i+1, students[i].EmplID, scoreExtractor(students[i]))
			categoryTopStudents = append(categoryTopStudents, StudentRanking{
				EmplID: students[i].EmplID,
				Marks:  scoreExtractor(students[i]),
				Rank:   i + 1,
			})
		}
		topStudents[category] = categoryTopStudents
	}

	// to handle JSON export if flag is set
	if *exportFlag == "json" {
		// creating summary report for export
		report := SummaryReport{
			GeneralAverages: CategoryAverages{
				Quiz:       quizSum / totalStudents,
				MidSem:     midSemSum / totalStudents,
				LabTest:    labTestSum / totalStudents,
				WeeklyLabs: weeklyLabSum / totalStudents,
				PreCompre:  preCompreSum / totalStudents,
				Compre:     compreSum / totalStudents,
				Total:      totalSum / totalStudents,
			},
			BranchAverages: branchAverages,
			TopStudents:    topStudents,
			Discrepancies:  discrepancies,
		}

		// to generate output filename based on input file
		baseName := strings.TrimSuffix(filePath, ".xlsx")
		outputFile := baseName + "_report.json"

		// Marshal to JSON and save to file
		jsonData, err := json.MarshalIndent(report, "", "  ")
		if err != nil {
			log.Fatalf("Failed to generate JSON: %v", err)
		}

		err = os.WriteFile(outputFile, jsonData, 0644)
		if err != nil {
			log.Fatalf("Failed to write JSON file: %v", err)
		}

		fmt.Printf("\nReport exported to %s\n", outputFile)
	}
}

package main

// Tasks remaining to do:
// - Create subtotal for each donor
// - write into Word documents

import (
	"cmp"
	"os"
	"fmt"
	"strconv"
	"strings"
	"slices"
	"regexp"
	"errors"
	"archive/zip"
    "github.com/xuri/excelize/v2"
    "github.com/lukasjarosch/go-docx"
)

// Define a struct
type transaction struct {
	date string
	checkNumber string
	checkType string
	amount string
}

func fileExists(path string) bool {
    _, err := os.Stat(path)
    if err == nil {
        return true
    }
    if errors.Is(err, os.ErrNotExist) {
        return false
    }
    // File may exist but is inaccessible (e.g., permission denied)
    return false 
}

func isStringAnInt(s string) bool {
	// Atoi is a shortcut for ParseInt(s, 10, 0)
	if _, err := strconv.Atoi(s); err == nil {
		return true
	}
	return false
}

func main() {

	if len(os.Args) < 3 {
		fmt.Println()
		fmt.Println("Usage: generate-tithe-receipts <zipfile_path> <template_path")
		fmt.Println()
		fmt.Println("       <zipfile_path> is the path to a zip file containing giving sheets")
		fmt.Println("       <template_path> is the path to a Word document template")
		fmt.Println()
		return
	}

	filePath := os.Args[1]
	templatePath := os.Args[2]

	transactionLog := make(map[string][]transaction)
	yearCount := make(map[string]int)

	r, err := zip.OpenReader(filePath)
	if err != nil {
		fmt.Println(err)
	}
	defer r.Close()

	// sort by filename so that dates will be in order
	slices.SortFunc(r.File, func(a, b *zip.File) int {
		return cmp.Compare(a.Name, b.Name)
	})

	// loop through files in zip archive
	for _, file := range r.File {
		//fmt.Printf("File Name: %s\n", file.Name)

		// 3. Open the file within the archive.
		rc, err := file.Open()
        if err != nil {
        	panic(err)
        }
        defer rc.Close()

		f, err := excelize.OpenReader(rc)
		if err != nil {
			fmt.Println(err)
			return
		}
		// Get value from cell by given worksheet name and axis.
		a1cell, err := f.GetCellValue("Sheet1", "A1")
		if err != nil {
			fmt.Println(err)
			return
		}
		a2cell, err := f.GetCellValue("Sheet1", "A2")
		if err != nil {
			fmt.Println(err)
			return
		}
		if len(a1cell) != 38 ||
		   len(a2cell) != 16 {
			fmt.Println("error - this appears to not be a giving sheet")
			return
		}
		fileDate, err := f.GetCellValue("Sheet1", "B3")
		if err != nil {
			fmt.Println(err)
			return
		}

		// use regex to increment value of yearCount
		// so we can guess the tax year of the donations
		//fmt.Println(fileDate)
		yearRegex := `(\d{2})$`
		re := regexp.MustCompile(yearRegex)
		transactionYear := re.FindStringSubmatch(fileDate)
		yearCount["20" + transactionYear[0]]++

		// Get all the rows in the Sheet1.
		rows, err := f.GetRows("Sheet1")
		if err != nil {
			fmt.Println(err)
			return
		}
		checkType := "Other"
		rowCounter := 0
		for _, row := range rows {
			rowCounter++
			if len(row) == 0 {
				continue
			}
			if row[0] == "General Offering Checks:" {
				checkType = "General Offering"
			} else if row[0] == "Deacon Offering Checks" {
				checkType = "Deacons Fund"
			} else if row[0] == "Other Designated Checks (Blg, book, Splits for Deacon Fund, etc.)" {
				checkType = "Other"
			} else if isStringAnInt(row[0]) && len(row) >= 4 {
				//if we are here this line must be a check
				if row[1] == "" {
					continue
				}
				tempResetCheckType := false
				if checkType == "Other" && len(row) >= 5 {
					tempResetCheckType = true
					if strings.Contains(row[4], "Building") {
						checkType = "Building Fund"
					} else if strings.Contains(row[4], "Deacon") {
						checkType = "Deacons Fund" 
					} else if row[4] == "Thank Offering" {
						checkType = "Thank Offering"
					}
				}
				personName := row[1]
				checkNumber := row[2]
				amount := strings.TrimSpace(row[3])
				//fmt.Printf("%10s\t%10s\t%40s\t%18s\t%10s\n", fileDate, checkNumber, personName, checkType, amount)
				transactionLog[personName] = append(transactionLog[personName], transaction{date: fileDate, checkNumber: checkNumber, amount: amount, checkType: checkType})
				if tempResetCheckType {
					checkType = "Other"
				}
			}
			// check for designated cash
			if len(row) >= 6 {
				if strings.ContainsRune(row[4], ',') {
					personName := row[4]
					checkNumber := "cash"
					amount := strings.TrimSpace(row[6])
					//fmt.Printf("%10s\t%10s\t%40s\t%18s\t%10s\n", fileDate, checkNumber, personName, checkType, amount)
					transactionLog[personName] = append(transactionLog[personName], transaction{date: fileDate, checkNumber: checkNumber, amount: amount, checkType: checkType})
				}
			}
		}
	}

	// guess the year based on which year has the highest yearCount
    taxYear := ""
    maxCount := 0
    for year, count := range yearCount {
        if count > maxCount {
            maxCount = count
            taxYear = year
        }
    }
	//fmt.Print("The year is ", taxYear)

	// loop through people and print out their receipts
	for personName, titheslice := range transactionLog {
		donationTable := ""
		//fmt.Printf("%d tithes for %s\n", len(titheslice), personName)
		for _, t := range titheslice {
			donationTable += fmt.Sprintf("- %10s\t%10s\t%20s\t%18s\t%10s\n", t.date, t.checkNumber, personName, t.checkType, t.amount)
		}
		replaceMap := docx.PlaceholderMap{
			"year":				taxYear,
			"name":				personName,
			"donationTable":	donationTable,
		}
		doc, err := docx.Open(templatePath)
		if err != nil {
			panic(err)
		}
		err = doc.ReplaceAll(replaceMap)
		if err != nil {
			panic(err)
		}
		re := regexp.MustCompile(`,* +`)
		newPersonName := re.ReplaceAllString(personName, "-")
		re = regexp.MustCompile(`\(+`)
		newPersonName = re.ReplaceAllString(newPersonName, "")
		re = regexp.MustCompile(`\)+`)
		newPersonName = re.ReplaceAllString(newPersonName, "")
		re = regexp.MustCompile(`\&+`)
		newPersonName = re.ReplaceAllString(newPersonName, "and")
		outputPath := "receipts-" + taxYear + "-" + newPersonName + ".docx"
		if fileExists(outputPath) {
			//panic("duplicate file name for " + outputPath)
			fmt.Println("Warning: Overwriting file due to duplicate file name for " + outputPath)
		}
		err = doc.WriteToFile(outputPath)
		if err != nil {
			panic(err)
		}
	}
}

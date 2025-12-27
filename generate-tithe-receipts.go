package main

import (
	"os"
	"fmt"
	"strconv"
	"strings"
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

func isStringAnInt(s string) bool {
	// Atoi is a shortcut for ParseInt(s, 10, 0)
	if _, err := strconv.Atoi(s); err == nil {
		return true
	}
	return false
}

func main() {

	if len(os.Args) < 2 {
		fmt.Println()
		fmt.Println("Usage: generate-tithe-receipts <file_path>")
		fmt.Println()
		fmt.Println("       <file_path> is the path to a zip file containing giving sheets")
		fmt.Println()
		return
	}

	filePath := os.Args[1]

	transactionLog := make(map[string][]transaction)

	r, err := zip.OpenReader(filePath)
	if err != nil {
		fmt.Println(err)
	}
	defer r.Close()

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
	// loop through people and print out their receipts
	for personName, titheslice := range transactionLog {
		//fmt.Printf("%d tithes for %s\n", len(titheslice), personName)
		for _, t := range titheslice {
			fmt.Printf("%10s\t%10s\t%40s\t%18s\t%10s\n", t.date, t.checkNumber, personName, t.checkType, t.amount)
		}
	}
}

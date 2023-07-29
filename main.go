package main

import (
	"fmt"
	"math/rand"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// Function: generateRandomPhoneNumber
// Description: Function for generating random phone number
// Returns: It will return a string of phone number that starts with 628 and followed by 8 random digits
func generateRandomPhoneNumber() string {
	rand.Seed(time.Now().UnixNano())
	phonePrefix := "628"
	randomNumber := rand.Intn(90000000) + 10000000
	return phonePrefix + strconv.Itoa(randomNumber)
}

// Function: generateRandomAmount
// Description: Function for generating random amount
// Returns: It will return integer of random amount between 10000 and 50000
func generateRandomAmount() int {
	rand.Seed(time.Now().UnixNano())
	return rand.Intn(40001) + 10000
}

func main() {
	var numColumns int
	fmt.Println("Enter the number of columns you want:")
	fmt.Scan(&numColumns)

	if numColumns <= 0 {
		fmt.Println("Invalid number of columns.")
		return
	}

	// Create a new Excel file
	file := excelize.NewFile()
	sheetName := "Sheet1"

	// We're set the column names
	columnNames := []string{
		"Beneficiary Bank Code",
		"Beneficiary Account",
		"Mobile Number",
		"Amount",
	}

	// Set the cell style for center alignment
	centerStyle, _ := file.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})

	// Write the column names to the Excel file
	for colIndex, columnName := range columnNames {
		cell := fmt.Sprintf("%c%d", 'A'+colIndex, 1)
		file.SetCellValue(sheetName, cell, columnName)
		file.SetCellStyle(sheetName, cell, cell, centerStyle)
	}

	totalAmount := 0
	for i := 1; i <= numColumns; i++ {
		// Set the values for each row
		data := []interface{}{
			// If you want take from request we can change this, same with BeneficiaryAccount
			"GNESIDJA",
			510654320,
			generateRandomPhoneNumber(),
		}

		// Im doing this because I need to to calculate total of amount
		amount := generateRandomAmount()
		data = append(data, amount)

		// Calculate the total amount
		totalAmount += amount

		// Write the data to the Excel file
		rowIndex := strconv.Itoa(i + 1)
		for colIndex, value := range data {
			cell := fmt.Sprintf("%c%s", 'A'+colIndex, rowIndex)
			file.SetCellValue(sheetName, cell, value)
			file.SetCellStyle(sheetName, cell, cell, centerStyle)
		}
	}

	// Save the Excel file
	fileName := ""
	fmt.Println("Enter the name of file you want: ")
	fmt.Scan(&fileName)

	excelFile := fileName + ".csv"

	err := file.SaveAs(excelFile)
	if err != nil {
		fmt.Println("Error saving the file:", err)
		return
	}

	fmt.Println("Total Amount:", totalAmount)
	fmt.Println("Excel file generated and saved as " + excelFile + ".")
}

package convert

import (
	"archive/zip"
	"errors"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"unicode"
)

type SharedString struct {
	totalCount         int
	uniqueCount        int
	sharedStrings      string
	sharedStringBuffer map[string]int
}

type ExelFile struct {
	sharedStringFile     *os.File
	workBookFile         *os.File
	stylesFile           *os.File
	worksheetFile        *os.File
	relationFile         *os.File
	contentTypeFile      *os.File
	workbookRelationFile *os.File
}

func printLogMessage(message string) {
	fmt.Println(message)
}

func readCsvFile(inputFileName string) ([]byte, error) {
	if inputFileName == "" {
		return nil, errors.New("input file not found")
	}

	filename := strings.Split(inputFileName, ".")

	if len(filename) != 2 {
		return nil, errors.New("unknown  file  found")
	}

	if filename[1] != "csv" {
		return nil, errors.New("unknown file  format" + filename[1])
	}

	data, err := os.ReadFile(inputFileName)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func getDataType(data string) string {
	runeData := []rune(data)
	for _, char := range runeData {
		if unicode.IsDigit(char) {
			continue
		} else if string(char) == "." {
			continue
		} else {
			return "s"
		}
	}

	return "n"
}

func createExcelFiles() error {

	if err := os.MkdirAll("_excel", 0755); err != nil {
		return err
	}

	if _, err := os.Create(filepath.Join("_excel", "[Content_Types].xml")); err != nil {
		return err
	}

	// _rels directory
	if err := os.MkdirAll(filepath.Join("_excel", "_rels"), 0755); err != nil {
		return err
	}

	if _, err := os.Create(filepath.Join("_excel", "_rels", ".rels")); err != nil {
		return err
	}

	// xl directories
	if err := os.MkdirAll(filepath.Join("_excel", "xl", "worksheets"), 0755); err != nil {
		return err
	}

	if err := os.MkdirAll(filepath.Join("_excel", "xl", "_rels"), 0755); err != nil {
		return err
	}

	// xl files
	if _, err := os.Create(filepath.Join("_excel", "xl", "workbook.xml")); err != nil {
		return err
	}

	if _, err := os.Create(filepath.Join("_excel", "xl", "sharedStrings.xml")); err != nil {
		return err
	}

	if _, err := os.Create(filepath.Join("_excel", "xl", "styles.xml")); err != nil {
		return err
	}

	// worksheet
	if _, err := os.Create(filepath.Join("_excel", "xl", "worksheets", "sheet1.xml")); err != nil {
		return err
	}

	// workbook rels
	if _, err := os.Create(filepath.Join("_excel", "xl", "_rels", "workbook.xml.rels")); err != nil {
		return err
	}

	log.Println("Successfully Created File Structure")
	return nil
}

// utility function to write to file
func writeToFile(file *os.File, content string) error {
	_, err := file.WriteString(content)
	return err
}

// function that helps to create shared strings table and index
func addSharedString(
	sharedString *SharedString,
	key string) int {

	if value, ok := sharedString.sharedStringBuffer[key]; ok {
		sharedString.totalCount++
		return value
	}

	currentIndex := len(sharedString.sharedStringBuffer)
	(sharedString.sharedStringBuffer)[key] = currentIndex + 1
	sharedString.totalCount++
	sharedString.uniqueCount++
	sharedString.sharedStrings += formatedSharedString(key)
	return currentIndex + 1
}

// retives the shared string id from the string buffer
func getSharedStringIndex(
	sharedString *SharedString,
	key string) int {
	idx, _ := sharedString.sharedStringBuffer[key]
	return idx
}

func formatedSharedString(value string) string {
	return "<si><t>" + value + "</t></si>"
}

func sharedString(file *os.File, sharedString SharedString) error {
	_, err := file.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	if err != nil {
		return err
	}

	metaData := fmt.Sprintf(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">`, sharedString.totalCount, sharedString.uniqueCount)

	_, err = file.WriteString(metaData)

	if err != nil {
		return err
	}
	_, err = file.WriteString(sharedString.sharedStrings)

	if err != nil {
		return err
	}

	_, err = file.WriteString("</sst>")

	if err != nil {
		return err
	}
	return nil
}

func openAllExcelFiles() (*ExelFile, error) {

	contentTypeFile, err := os.OpenFile(filepath.Join("_excel", "[Content_Types].xml"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the relation file for maintaining relation of excel
	relationFile, err := os.OpenFile(filepath.Join("_excel", "_rels", ".rels"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the workbook file for maintaining various work books of excelkka
	workBookFile, err := os.OpenFile(filepath.Join("_excel", "xl", "workbook.xml"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the sharedstring  file to store string and index them later
	sharedStringFile, err := os.OpenFile(filepath.Join("_excel", "xl", "sharedStrings.xml"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the stlyle file to store files
	stylesFile, err := os.OpenFile(filepath.Join("_excel", "xl", "styles.xml"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the worksheet file for maintaining various sheets
	workSheetFile, err := os.OpenFile(filepath.Join("_excel", "xl", "worksheets", "sheet1.xml"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}

	// opeaning the workbookRelation file for maintaining  relation between workboks.
	workbookRelationFile, err := os.OpenFile(filepath.Join("_excel", "xl", "_rels", "workbook.xml.rels"), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return nil, err
	}
	return &ExelFile{
		workBookFile:         workBookFile,
		worksheetFile:        workSheetFile,
		relationFile:         relationFile,
		workbookRelationFile: workbookRelationFile,
		contentTypeFile:      contentTypeFile,
		stylesFile:           stylesFile,
		sharedStringFile:     sharedStringFile,
	}, nil
}

func writeContent(file *os.File) error {
	content := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
	return writeToFile(file, content)
}

func writeRelation(file *os.File) error {
	content := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
	return writeToFile(file, content)

}

func writeWorkbook(file *os.File) error {
	content := `<?xml version="1.0" encoding="UTF-8"?>
	<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<sheets>
			<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
		</sheets>
	</workbook>`
	return writeToFile(file, content)
}

func writeWorkbookRelations(file *os.File) error {
	content := `<?xml version="1.0" encoding="UTF-8"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship
			Id="rId1"
			Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
			Target="worksheets/sheet1.xml"/>

		<Relationship
			Id="rId2"
			Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
			Target="sharedStrings.xml"/>

		<Relationship 
			Id="rId3" 
			Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" 
			Target="styles.xml"/>
	</Relationships>

	`
	return writeToFile(file, content)
}

func writeStyle(file *os.File) error {
	content := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`
	return writeToFile(file, content)
}

func getExcelColumns(colIdx int) string {
	col := ""

	for colIdx > 0 {
		colIdx-- // adjust for 1-based indexing
		col = string(rune('A'+(colIdx%26))) + col
		colIdx /= 26
	}

	return col
}

// writes the remaining data to worksheet/sheets1.xml
func writeToWorkSheet(file *os.File, data string) error {
	_, err := file.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`)
	if err != nil {
		return err
	}
	_, err = file.WriteString(`<sheetData>`)

	if err != nil {
		return err
	}
	_, err = file.WriteString(data)

	if err != nil {
		return err
	}
	_, err = file.WriteString(`</sheetData>`)

	if err != nil {
		return err
	}

	_, err = file.WriteString(`</worksheet>`)

	if err != nil {
		return err
	}
	return nil
}

// Created worksheet data from xml string
func getXmlData(csvLineData []string, delimeter string, sharedStringBuffer *SharedString) string {
	xmlData := ""
	for row, rowData := range csvLineData {
		cellData := strings.Split(rowData, delimeter)
		xmlData += fmt.Sprintf("<row r = \"%d\" >", row+1)

		for col, colData := range cellData {
			alphCol := fmt.Sprintf("%s%d", getExcelColumns(col+1), row+1)
			dataType := getDataType(colData)
			//shared string data storing for writing later
			if dataType == "s" {
				addSharedString(sharedStringBuffer, colData)

				xmlData += fmt.Sprintf(`<c r="%s" t="%s"><v>%d</v></c>`, alphCol, dataType, getSharedStringIndex(sharedStringBuffer, colData))

			} else {

				xmlData += fmt.Sprintf(`<c r="%s"><v>%s</v></c>`, alphCol, colData)

			}
		}
		xmlData += "</row>"
	}
	return xmlData
}

// converts the xml data into xml and creates all the necessary files. Eg: sharedStrings, worksheet/sheet1
func parseCsvToXml(data, delimeter string) error {
	csvLineData := strings.Split(data, "\n")

	printLogMessage("Creating Necessary Files And Directories")
	// opens necessary files for excel
	files, err := openAllExcelFiles()
	if err != nil {
		return err
	}

	printLogMessage("Writing to content file")
	// write content metadata to the content type file
	err = writeContent(files.contentTypeFile)
	if err != nil {
		return err
	}

	printLogMessage("Writing relations")
	// write relation to the relations file
	err = writeRelation(files.relationFile)
	if err != nil {
		return err
	}

	defer files.workBookFile.Close()
	defer files.contentTypeFile.Close()
	defer files.stylesFile.Close()
	defer files.workbookRelationFile.Close()
	defer files.sharedStringFile.Close()
	defer files.relationFile.Close()
	defer files.worksheetFile.Close()
	// writing types to the content type files

	//sharedStringBuffer
	sharedStringBuffer := &SharedString{
		totalCount:         0,
		uniqueCount:        0,
		sharedStrings:      "",
		sharedStringBuffer: map[string]int{},
	}

	xmlData := getXmlData(csvLineData, delimeter, sharedStringBuffer)
	printLogMessage("Writing XML data to worksheet")
	err = writeToWorkSheet(files.worksheetFile, xmlData)
	if err != nil {
		return err
	}

	err = sharedString(files.sharedStringFile, *sharedStringBuffer)
	if err != nil {
		return err
	}

	err = writeWorkbook(files.workBookFile)
	if err != nil {
		return err
	}

	err = writeWorkbookRelations(files.workbookRelationFile)
	if err != nil {
		return err
	}

	err = writeStyle(files.stylesFile)
	if err != nil {
		return err
	}
	return nil
}

func createExcelStructure() error {
	//Create file structure for the exc1234567891011el conversion
	return createExcelFiles()

}

// Compress all the xml file into single xlxs file
func zipExcel(sourceDir string, output string) error {

	zipFile, err := os.Create(output)
	if err != nil {
		return err
	}
	defer zipFile.Close()

	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	return filepath.Walk(sourceDir, func(path string, info os.FileInfo, err error) error {

		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		relPath, err := filepath.Rel(sourceDir, path)
		if err != nil {
			return err
		}

		file, err := os.Open(path)
		if err != nil {
			return err
		}
		defer file.Close()

		writer, err := zipWriter.Create(relPath)
		if err != nil {
			return err
		}

		_, err = io.Copy(writer, file)
		return err
	})
}

func Convert(inputFileName, outputFileName, delimeter string) error {

	data, csvReadingError := readCsvFile(inputFileName)

	if csvReadingError != nil {
		return csvReadingError
	}

	err := createExcelStructure()
	if err != nil {
		return err
	}

	err = parseCsvToXml(string(data), delimeter)
	if err != nil {
		return err
	}

	printLogMessage("Compressing Xml to Xlsx")
	err = zipExcel("_excel", outputFileName)
	if err != nil {
		return err
	}
	printLogMessage("Removing Uncessary Files")
	return os.RemoveAll("_excel")
}

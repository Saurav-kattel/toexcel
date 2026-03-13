package main

import (
	"flag"
	"fmt"
	"sauravkattel/efmt/internal/convert"
)

func showAvailableCmd() {
	fmt.Println("emft	-h	<shows all the commands and subcommands>	")
	fmt.Println("emft	-c	[ input file ]	[ output file ]	<converts the space included file into excel file>	")
	fmt.Println("emft	-d	<delimiter that seperated values. Default is comma>	")
}

func main() {
	helpFlag := flag.Bool("h", false, "help")
	convertFlag := flag.Bool("c", false, "convert mode")
	delimFlag := flag.String("d", ",", "delimieter flang")

	flag.Parse()

	// help flags prints all the available commands
	if helpFlag != nil && *helpFlag == true {
		showAvailableCmd()
		return
	}

	if convertFlag != nil && *convertFlag == true {
		convertFileArgs := flag.Args()
		if len(convertFileArgs) != 2 {
			fmt.Println("efmt -c [ input file ] [ output file ]")
			return
		}
		inputFileName := convertFileArgs[0]
		outputFileName := convertFileArgs[1]

		convertError := convert.Convert(inputFileName, outputFileName, *delimFlag)
		if convertError != nil {
			fmt.Println("Error: " + convertError.Error())
		} else {
			fmt.Println("Conversion successful \nGenerated " + outputFileName)
		}

	}

}

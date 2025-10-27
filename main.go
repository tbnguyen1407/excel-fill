package main

import (
	"errors"
	"flag"
	"fmt"
	"log/slog"
	"os"
	"regexp"
	"strconv"

	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v3"
)

var (
	version       = "unset"                     // set during build with -ldflags
	mFilterRegexp = map[string]*regexp.Regexp{} // map[opidx_filteridx]filtervalue
)

func main() {
	// define flags
	var inFilePath string
	var configFilePath string
	var outFilePath string
	var sheetIndex int
	var versionFlag bool

	// parse flags
	flag.StringVar(&inFilePath, "in", "REQUIRED", "Path to input XLSX workbook")
	flag.StringVar(&outFilePath, "out", "out.xlsx", "Path to output XLSX workbook")
	flag.StringVar(&configFilePath, "config", "REQUIRED", "Path to config file")
	flag.IntVar(&sheetIndex, "sheet", 0, "Worksheet number")
	flag.BoolVar(&versionFlag, "version", false, "Print app version")
	flag.Parse()

	// print version
	if versionFlag {
		fmt.Println(version)
		return
	}

	// load book
	bookFile, e := excelize.OpenFile(inFilePath)
	exitOnError(e)
	defer bookFile.Close()

	sheetName := bookFile.GetSheetName(sheetIndex)
	if sheetName == "" {
		exitOnError(errors.New("sheet not found"))
	}

	// load config
	configBytes, e := os.ReadFile(configFilePath)
	exitOnError(e)

	var config Config
	e = yaml.Unmarshal(configBytes, &config)
	exitOnError(e)

	// validate config
	e = validate(&config)
	exitOnError(e)

	// process
	e = process(bookFile, sheetName, &config)
	exitOnError(e)

	// write
	e = bookFile.SaveAs(outFilePath)
	exitOnError(e)
}

func validate(config *Config) error {
	if config == nil {
		return errors.New("input is nil")
	}

	for opIdx, op := range config.Operations {
		for filterIdx, filter := range op.Filters {
			// validate
			if filter.Column == "" {
				return fmt.Errorf("operations[%d].when[%d].column is empty", opIdx, filterIdx)
			}

			// cache
			cacheKey := strconv.Itoa(opIdx) + "_" + strconv.Itoa(filterIdx)
			cacheVal := regexp.MustCompile(filter.Value)
			mFilterRegexp[cacheKey] = cacheVal
		}

		for actionIdx, action := range op.Actions {
			// validate
			if action.Column == "" {
				return fmt.Errorf("operations[%d].do[%d].column is empty", opIdx, actionIdx)
			}
		}
	}
	return nil
}

func process(bookFile *excelize.File, sheetName string, config *Config) error {
	rows, e := bookFile.GetRows(sheetName)
	if e != nil {
		return e
	}

	for row := 2; row <= len(rows); row++ {
		for opIdx, op := range config.Operations {
			// check if all filters match
			matched := true
			for filterIdx, filter := range op.Filters {
				// get cell value
				cellAddr := filter.Column + strconv.Itoa(row)
				cellValue, e := bookFile.GetCellValue(sheetName, cellAddr)
				if e != nil {
					return e
				}

				// retrieve cached filter regexp
				cacheKey := strconv.Itoa(opIdx) + "_" + strconv.Itoa(filterIdx)
				filterRegexp := mFilterRegexp[cacheKey]

				// compare cell value
				if !filterRegexp.MatchString(cellValue) {
					matched = false
					break
				}
			}

			// execute action if matched
			if matched {
				for _, action := range op.Actions {
					actionCellAddr := action.Column + strconv.Itoa(row)
					e := bookFile.SetCellValue(sheetName, actionCellAddr, action.Value)
					if e != nil {
						return e
					}
					slog.Info("match", "r", row, "op", op.Name, "action", actionCellAddr+" -> "+action.Value)
				}
			}
		}
	}

	return nil
}

func exitOnError(e error) {
	if e != nil {
		slog.Error("terminating", "error", e)
		os.Exit(1)
	}
}

type Config struct {
	Operations []Operation `yanl:"operations"`
}

type Operation struct {
	Name    string   `yaml:"name"`
	Filters []Filter `yaml:"filters"`
	Actions []Action `yaml:"actions"`
}

type Filter struct {
	Column string `yaml:"column"`
	Value  string `yaml:"value"`
}

type Action struct {
	Column string `yaml:"column"`
	Value  string `yaml:"value"`
}

[![codecov](https://codecov.io/gh/gofika/excel/branch/main/graph/badge.svg)](https://codecov.io/gh/gofika/excel)
[![Build Status](https://github.com/gofika/excel/workflows/build/badge.svg)](https://github.com/gofika/excel)
[![go.dev](https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white)](https://pkg.go.dev/github.com/gofika/excel)
[![Go Report Card](https://goreportcard.com/badge/github.com/gofika/excel)](https://goreportcard.com/report/github.com/gofika/excel)
[![Licenses](https://img.shields.io/github/license/gofika/excel)](LICENSE)
[![donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](about::blank)

# Excel

Microsoft Excel .xlsx read/write for golang

## Basic Usage

### Installation

To get the package, execute:

```bash
go get github.com/gofika/excel
```

To import this package, add the following line to your code:

```js
import "github.com/gofika/excel"
```

### Create spreadsheet

Here is example usage that will create xlsx file.

```go
package main

import (
	"fmt"
	"github.com/gofika/excel"
	"time"
)

func main() {
	f := excel.NewFile()

	sheet := f.NewSheet("Sheet2")
	sheet.SetCellValue(excel.ColumnNumber("A"), 1, "Name")
	sheet.SetCellValue(excel.ColumnNumber("A"), 2, "Jason")
	sheet.SetCellValue(excel.ColumnNumber("B"), 1, "Score")
	sheet.SetCellValue(excel.ColumnNumber("B"), 2, 100)
	// date value
	sheet.SetCellValue(3, 1, "Date")
	sheet.Cell(3, 2).SetDateValue(time.Date(1980, 9, 8, 0, 0, 0, 0, time.Local))
	// time value
	sheet.AxisCell("D1").SetStringValue("LastTime")
	sheet.AxisCell("D2").
		SetTimeValue(time.Now()).
		SetNumberFormat("yyyy-mm-dd hh:mm:ss")

	if err := f.SaveFile("Document1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
```

### Reading spreadsheet

The following constitutes the bare to read a spreadsheet document.

```go
package main

import (
	"fmt"
	"github.com/gofika/excel"
)

func main() {
	f, err := excel.OpenFile("Document1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	sheet := f.OpenSheet("Sheet2")
	A1 := sheet.GetCellString(1, 1)
	fmt.Println(A1)

	cell := sheet.AxisCell("B2")
	fmt.Println(cell.GetIntValue())
}
```

## TODO:

- [x] Basic File Format
- [x] File: NewFile, OpenFile, SaveFile, Write, Sheets
- [ ] Sheet:
    - [x] NewSheet, OpenSheet
    - [x] SetCellValue, GetCellString, GetCellInt, Cell, AxisCell
    - [ ] ...
- [ ] Cell:
    - [x] Row, Col
    - [x] SetValue, SetIntValue, SetFloatValue, SetFloatValuePrec, SetStringValue, SetBoolValue, SetDefaultValue,
      SetTimeValue, SetDateValue, SetDurationValue
    - [x] GetIntValue, GetStringValue
    - [x] SetNumberFormat
    - [ ] ...

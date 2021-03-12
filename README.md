[![codecov](https://codecov.io/gh/leaker/excel/branch/main/graph/badge.svg?token=7OMXdv8pb2)](https://codecov.io/gh/leaker/excel)
[![Build Status](https://github.com/leaker/excel/workflows/build/badge.svg)](https://github.com/leaker/excel)
[![go.dev](https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white)](https://pkg.go.dev/github.com/leaker/excel)
[![Go Report Card](https://goreportcard.com/badge/github.com/leaker/excel)](https://goreportcard.com/report/github.com/leaker/excel)
[![Licenses](https://img.shields.io/badge/license-bsd-orange.svg)](https://opensource.org/licenses/BSD-2-Clause)
[![donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](about::blank)

# XLSX
Microsoft Excel .xlsx read/write for golang


## Basic Usage

### Installation

To get the package, execute:
```bash
go get github.com/leaker/excel
```
To import this package, add the following line to your code:
```go
import "github.com/leaker/excel"
```


### Create spreadsheet

Here is example usage that will create xlsx file.

```go
package main

import (
    "fmt"

    "github.com/leaker/excel"
)

func main() {
    f := xlsx.NewFile()
    
    sheet := f.NewSheet("Sheet2")
	sheet.SetCellValue(xlsx.ColumnNumber("A"), 1, "Name")
	sheet.SetCellValue(xlsx.ColumnNumber("B"), 1, "Score")
	sheet.SetCellValue(xlsx.ColumnNumber("A"), 2, "Jason")
	sheet.SetCellValue(xlsx.ColumnNumber("B"), 2, 100)
    
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

    "github.com/leaker/excel"
)

func main() {
    f, err := xlsx.OpenFile("Document1.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }
    
    sheet := f.OpenSheet("Sheet2")
    A1 := sheet.GetCellString(1, 1)
    fmt.Println(A1)
    
    cell := sheet.Cell(xlsx.ColumnNumber("B"), 2)
    fmt.Println(cell.GetIntValue())
}
```

## TODO:
- [x] Basic File Format
- [x] File: NewFile, OpenFile, SaveFile, Write
- [ ] Sheet: NewSheet, OpenSheet ...
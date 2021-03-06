package excel

import "io"

// File define for operation xlsx file
type File interface {
	// SaveFile save xlsx file
	SaveFile(name string) error

	// Write save to steam
	Write(w io.Writer) error

	// OpenSheet open a exist Sheet by name
	//
	// Example:
	//
	//     sheet := file.OpenSheet("Sheet1")
	//
	// return nil if sheet not exist
	OpenSheet(name string) Sheet

	// NewSheet create a new Sheet with sheet name
	// Example:
	//
	//     sheet := file.NewSheet("Sheet2")
	NewSheet(name string) Sheet

	// Sheets return all sheet for operator
	Sheets() []Sheet
}

// NewFile create a default xlsx File with default template
func NewFile() File {
	return newFile()
}

// OpenFile open a xlsx for operator
func OpenFile(name string) (File, error) {
	return openFile(name)
}

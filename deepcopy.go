package excel

import (
	"bytes"
	"encoding/gob"
)

// DeepCopy deepcopy object. can without same type
//
// Example:
//     type A struct {
//         Name string
//         Value int
//     }
//
//     type B struct {
//         Name string
//         Value int
//     }
//
//     a := &A { Name: "Jason", 100}
//     var b B
//     xlsx.DeepCopy(&b, a)
//     fmt.Printf("%+v\n", b)
func DeepCopy(dst, src interface{}) error {
	var buffer bytes.Buffer
	if err := gob.NewEncoder(&buffer).Encode(src); err != nil {
		return err
	}
	return gob.NewDecoder(&buffer).Decode(dst)
}

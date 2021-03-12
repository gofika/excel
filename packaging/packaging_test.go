package packaging

import (
	"encoding/xml"
	"os"
	"path"
	"strconv"
	"testing"

	. "gopkg.in/check.v1"
)

func Test(t *testing.T) { TestingT(t) }

type PackagingSuite struct{}

var _ = Suite(&PackagingSuite{})

func XMLMarshalAppendHead(v interface{}) (ret string, err error) {
	output, err := xml.Marshal(v)
	if err != nil {
		return
	}
	ret = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + string(output)
	return
}

func XMLMarshalAppendHeadIndent(v interface{}) (ret string, err error) {
	output, err := xml.MarshalIndent(v, "", "    ")
	if err != nil {
		return
	}
	ret = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + string(output)
	return
}

var templatePath = path.Join("../template")

const (
	isAssertDefaultTemplate = false
)

var (
	needWriteTestFile, _ = strconv.ParseBool(os.Getenv("NEED_WRITE_TEST_FILE"))
)

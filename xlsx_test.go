package excel

import (
	"testing"

	. "gopkg.in/check.v1"
)

func Test(t *testing.T) { TestingT(t) }

type XlsxSuite struct{}

var _ = Suite(&XlsxSuite{})

package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
※ 미구현
IXHwpSummaryInfo: 읽기전용??
IXHwpDocument 객체가 가진 XHwpSummaryInfo
ParameterSetObject.pdf 파일 19페이지 HSummaryInfo:문서 정보 속성 개체의 값을 읽어오는 객체가 되는듯
*/
type IXHwpSummaryInfo struct {
	variant *ole.VARIANT
}

/*
지은이
*/
func (p *IXHwpSummaryInfo) Author() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Author")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

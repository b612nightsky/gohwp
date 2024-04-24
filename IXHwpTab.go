package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpTab:

	IXHwpTab: 탭 오브젝트
	(Tab 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpTab struct {
	variant *ole.VARIANT
}

/*
탭을 닫는다.

	isDirty :
	true: 문서 내용이 변경된 경우 닫지 않는다.
	false: 문서 내용이 변경된 것과 상관없이 강제로 닫는다.
*/
func (p *IXHwpTab) Close(isDirty bool) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
}

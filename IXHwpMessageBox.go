package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpMessageBox:

	OLE Automation standard object
	(메세지박스 MessageBox 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpMessageBox struct {
	variant *ole.VARIANT
}

/*
최상위 오브젝트를 얻어옴 (IHwpObject - 읽기 전용)
*/
func (p *IXHwpMessageBox) Application() *IHwpObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Application")
	if err != nil {
		panic(err)
	}
	var ho IHwpObject
	ho.variant = res
	ho.dispatch = res.Value().(*ole.IDispatch)
	return &ho
}

/*
메시지 박스에 넣을 문자열
*/
func (p *IXHwpMessageBox) String(text string) {
	_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "String", text)
	if err != nil {
		panic(err)
	}
}

/*
메시지 박스에 사용할 Flag

	0: 확인
	1: 학인, 취소
	2: 중단, 재시도, 무시
	3: 예, 아니오, 취소
	4: 예, 아니오
	5: 재시도, 취소
	6: 예, 아니오, 모두, 취소
*/
func (p *IXHwpMessageBox) Flag(iflag int) {
	_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Flag", uint16(iflag))
	if err != nil {
		panic(err)
	}
}

/*
메시지 박스의 리턴값
*/
func (p *IXHwpMessageBox) Result() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Result")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint16))
}

/*
메시지 박스 보이기
*/
func (p *IXHwpMessageBox) DoModal() {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "DoModal")
	if err != nil {
		panic(err)
	}
}

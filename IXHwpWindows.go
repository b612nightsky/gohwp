package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpWindows:

	IXHwpWindows: IXHwpWindow오브젝트를 관리하는 오브젝트
	(Window를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpWindows struct {
	variant *ole.VARIANT
}

/*
최상위 오브젝트를 얻어옴 (IHwpObject - 읽기 전용)
*/
func (p *IXHwpWindows) Application() *IHwpObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Application")
	if err != nil {
		panic(err)
	}
	var ho IHwpObject
	ho.dispatch = res.Value().(*ole.IDispatch)
	return &ho
}

/*
현재 활성화 상태인 윈도우 Object를 얻어온다.(IXHwpWindow)
*/
func (p *IXHwpWindows) Active_XHwpWindow() *IXHwpWindow {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Active_XHwpWindow")
	if err != nil {
		panic(err)
	}
	var hw IXHwpWindow
	hw.variant = res
	return &hw
}

/*
지정한 원소의 윈도우 오브젝트를 얻어온다.
*/
func (p *IXHwpWindows) Item(index int) *IXHwpWindow {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Item", index)
	if err != nil {
		panic(err)
	}
	var hw IXHwpWindow
	hw.variant = res
	return &hw
}

/*
원소의 총개수
*/
func (p *IXHwpWindows) Count() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
윈도우를 하나 추가한다.(새창으로 열기 기능과 동일)
*/
func (p *IXHwpWindows) Add() *IXHwpWindow {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Add")
	if err != nil {
		panic(err)
	}
	var hw IXHwpWindow
	hw.variant = res
	return &hw
}

/*
윈도우를 모두 닫는다.

	isDirty :
	true: 문서 내용이 변경된 경우 닫지 않는다.
	false: 문서 내용이 변경된 것과 상관없이 강제로 닫는다.
*/
func (p *IXHwpWindows) Close(isDirty bool) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

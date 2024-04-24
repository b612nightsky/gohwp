package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpTabs:

	IXHwpTabs: IXHwpTab오브젝트를 관리하는 오브젝트
	(Tab를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpTabs struct {
	variant *ole.VARIANT
}

/*
윈도우에 열려있는 탭의 개수(읽기 전용)
*/
func (p *IXHwpTabs) Count() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
지정한 원소의 탭 오브젝트를 얻어온다.
*/
func (p *IXHwpTabs) Item(index int) *IXHwpTab {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Item", index)
	if err != nil {
		panic(err)
	}
	var ht IXHwpTab
	ht.variant = res
	return &ht
}

/*
지정한 원소의 탭 오브젝트를 추가한다.(문서를 새 탭으로 열기)
*/
func (p *IXHwpTabs) Add() *IXHwpTab {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Add")
	if err != nil {
		panic(err)
	}
	var ht IXHwpTab
	ht.variant = res
	return &ht
}

/*
탭을 모두 닫는다.

	isDirty :
	true: 문서 내용이 변경된 경우 닫지 않는다.
	false: 문서 내용이 변경된 것과 상관없이 강제로 닫는다.
*/
func (p *IXHwpTabs) Close(isDirty bool) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

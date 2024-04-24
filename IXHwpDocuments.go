package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpDocuments:

	IXHwpDocument(도큐먼트) 오브젝트를 관리하는 오브젝트
	(Document를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpDocuments struct {
	variant *ole.VARIANT
}

/*
최상위 오브젝트를 얻어옴 (IHwpObject)
*/
func (p *IXHwpDocuments) Application() *IHwpObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Application")
	if err != nil {
		panic(err)
	}
	var ho IHwpObject
	ho.variant = res
	return &ho
}

/*
지정한 원소의 도큐먼트 오브젝트를 얻어온다.(IXHwpDocument)

	index : 원소의 인덱스
*/
func (p *IXHwpDocuments) Item(index int) *IXHwpDocument {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Item", index)
	if err != nil {
		panic(err)
	}
	var hd IXHwpDocument
	hd.variant = res
	return &hd
}

/*
원소의 총개수
*/
func (p *IXHwpDocuments) Count() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
현재 활성화 상태인 도큐먼트 Object를 얻어온다.(IXHwpDocument)
*/
func (p *IXHwpDocuments) Active_XHwpDocument() *IXHwpDocument {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Active_XHwpDocument")
	if err != nil {
		panic(err)
	}
	var hd IXHwpDocument
	hd.variant = res
	return &hd
}

/*
도큐먼트 오브젝트를 추가한다.

	isTab
	true: 새탭으로 열리는 도큐먼트
	false: 새창으로 열리는 도큐먼트

	return IXHwpDocument
*/
func (p *IXHwpDocuments) Add(isTab bool) *IXHwpDocument {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Add", isTab)
	if err != nil {
		panic(err)
	}
	var hd IXHwpDocument
	hd.variant = res
	return &hd
}

/*
관리하고 있는 도큐먼트 오브젝트를 삭제한다.

	isDirty
	true: 변경된 문서는 닫지 않는다
	false: 변경된 문서도 닫는다
*/
func (p *IXHwpDocuments) Close(isDirty bool) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
}

/*
도큐먼트 아이디로 지정된 도큐먼트 오브젝트를 얻는다.

	docid: 도큐먼트의 고유 ID
	return : 도큐먼트 ID에 해당하는 도큐먼트 오브젝트(IXHwpDocument)
*/
func (p *IXHwpDocuments) FindItem(docid int) *IXHwpDocument {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "FindItem", docid)
	if err != nil {
		panic(err)
	}
	var hd IXHwpDocument
	hd.variant = res
	return &hd
}

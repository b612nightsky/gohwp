package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IXHwpWindow:

	IXHwpWindow: 윈도우 오브젝트
	(Window 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpWindow struct {
	variant *ole.VARIANT
}

/*
최상위 오브젝트를 얻어옴 (IHwpObject - 읽기 전용)
*/
func (p *IXHwpWindow) Application() *IHwpObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Application")
	if err != nil {
		panic(err)
	}
	var ho IHwpObject
	ho.dispatch = res.Value().(*ole.IDispatch)
	return &ho
}

/*
도큐먼트 관리 오브젝트를 얻어옴(IXHwpDocuments - 읽기 전용)
*/
func (p *IXHwpWindow) XHwpDocuments() *IXHwpDocuments {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "XHwpDocuments")
	if err != nil {
		panic(err)
	}
	var hd IXHwpDocuments
	hd.variant = res
	return &hd
}

/*
탭 관리 오브젝트를 얻어옴(IXHwpTabs - 읽기 전용)
*/
func (p *IXHwpWindow) XHwpTabs() *IXHwpTabs {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "XHwpTabs")
	if err != nil {
		panic(err)
	}
	var hts IXHwpTabs
	hts.variant = res
	return &hts
}

/*
윈도우의 좌측 위치 좌표를 설정/얻음

	Left() 좌표만 얻음
	Left(int) 좌표 설정 및 값 리턴
*/
func (p *IXHwpWindow) Left(pos ...int) int {
	var position int
	lenPos := len(pos)
	switch lenPos {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Left", int32(pos[0]))
		if err != nil {
			panic(err)
		}
		position = pos[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Left")
		if err != nil {
			panic(err)
		}
		position = int(res.Value().(int32))
	default:
		panic("Left(...int) 인자는 0개 또는 1개이여야 함")
	}
	return position
}

/*
윈도우의 맨위 위치 좌표를 설정/얻음

	Top() 좌표만 얻음
	Top(int) 좌표 설정 및 값 리턴
*/
func (p *IXHwpWindow) Top(pos ...int) int {
	var position int
	lenPos := len(pos)
	switch lenPos {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Top", int32(pos[0]))
		if err != nil {
			panic(err)
		}
		position = pos[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Top")
		if err != nil {
			panic(err)
		}
		position = int(res.Value().(int32))
	default:
		panic("Top(...int) 인자는 0개 또는 1개이여야 함")
	}
	return position
}

/*
윈도우의 너비를 설정/얻음

	Width() 너비만 얻음
	Width(int) 너비 설정 및 값 리턴
*/
func (p *IXHwpWindow) Width(width ...int) int {
	var w int
	lenWidth := len(width)
	switch lenWidth {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Width", int32(width[0]))
		if err != nil {
			panic(err)
		}
		w = width[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Width")
		if err != nil {
			panic(err)
		}
		w = int(res.Value().(int32))
	default:
		panic("Width(...int) 인자는 0개 또는 1개이여야 함")
	}
	return w
}

/*
윈도우의 높이를 설정/얻음

	Height() 높이만 얻음
	Height(int) 높이 설정 및 값 리턴
*/
func (p *IXHwpWindow) Height(height ...int) int {
	var h int
	lenHeight := len(height)
	switch lenHeight {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Height", int32(height[0]))
		if err != nil {
			panic(err)
		}
		h = height[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Height")
		if err != nil {
			panic(err)
		}
		h = int(res.Value().(int32))
	default:
		panic("Height(...int) 인자는 0개 또는 1개이여야 함")
	}
	return h
}

/*
윈도우 보이기/보이지 않기 설정/얻음

	Visible() 보임 여부 true/false 값만 얻음
	Visible(bool) 보임 여부 true/false 값 지정 및 값 리턴
*/
func (p *IXHwpWindow) Visible(visible ...bool) bool {
	var vb bool
	lenBool := len(visible)
	switch lenBool {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Visible", visible[0])
		if err != nil {
			panic(err)
		}
		vb = visible[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Visible")
		if err != nil {
			panic(err)
		}
		vb = res.Value().(bool)
	default:
		panic("Visible(...bool) 인자는 0개 또는 1개이여야 함")
	}
	return vb
}

/*
윈도우를 모두 닫는다.

	isDirty :
	true: 문서 내용이 변경된 경우 닫지 않는다.
	false: 문서 내용이 변경된 것과 상관없이 강제로 닫는다.
*/
func (p *IXHwpWindow) Close(isDirty bool) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
}

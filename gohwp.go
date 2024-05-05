package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type IHwpObject struct {
	variant  *ole.VARIANT
	dispatch *ole.IDispatch // HWP COM 객체의 IDispatch
}

/*
초기화

OLE 시스템 초기화, HWP COM 객체 및 IDispatch 인터페이스 객체 생성
*/
func Initialize() *IHwpObject {

	var hwp IHwpObject

	// OLE 시스템 초기화
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		panic(err)
	}
	//defer ole.CoUninitialize()

	// HWP COM 객체 생성
	com, err := oleutil.CreateObject("HWPFrame.HwpObject")
	if err != nil {
		panic(err)
	}
	//defer hwp.Release()

	// IDispatch 인터페이스 요청 및 반환
	dispatch, err := com.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		panic(err)
	}
	defer dispatch.Release()

	hwp.dispatch = dispatch

	return &hwp
}

/*
소멸자

마지막 OLE 시스템 없애기, Initialize()로 초기화 이후 defer UnInitialize()로 따라오기
*/
func (hwp *IHwpObject) UnInitialize() {
	hwp.dispatch.Release()
	//hwp.hwpCom.Release()
	ole.CoUninitialize()
}

/*
Action Object 모음
*/
func (hwp *IHwpObject) ActionObject() *ActionObject {
	var ab ActionObject
	ab.hwp = hwp
	return &ab
}

package gohwp

import (
	"fmt"
	"reflect"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IDHwpAction:

	IDHwpAction: 한/글에서 특정 기능을 수행하기 위한 액션 오브젝트
	(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpAction과 동일)
*/
type IDHwpAction struct {
	variant *ole.VARIANT
}

/*
액션 ID를 나타낸다. 읽기 전용.
*/
func (p *IDHwpAction) ActID() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ActID")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
액션이 사용하는 parameter set ID를 나타낸다. 읽기 전용.
*/
func (p *IDHwpAction) SetID() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "SetID")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
현재 상태에 따라 액션 실행에 필요한 인수를 구한다.

	param : 인수를 저장할 Parameter Set
*/
func (p *IDHwpAction) GetDefault(param *IDHwpParameterSet) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "GetDefault", param.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
}

/*
액션과 대응하는 Parameter Set을 생성한다.
*/
func (p *IDHwpAction) CreateSet() *IDHwpParameterSet {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "CreateSet")
	if err != nil {
		panic(err)
	}
	var dhps IDHwpParameterSet
	dhps.variant = res
	return &dhps
}

/*
지정한 인수로 액션을 실행한다.

	parma : 액션의 실행을 제어할 인수, ParameterSet의 종류와 아이템의 의미는 액션이 정의한 바에 따라 다르다.(IDHwpParameterSet)
*/
func (p *IDHwpAction) Execute(param *IDHwpParameterSet) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Execute", param.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
}

/*
액션의 대화상자를 띄운다.

	parma : 여기에 지정된 아이템의 값에 따라 대화상자의 컨트롤의 초기값이 결정되고, 대화상자가 닫힌 후에는 사용자가 지정한 값들이 담겨 돌아온다.
	return : 액션이 정의하기에 따라 다르지만, 일반적으로 다음과 같은 modal dialog result를 리턴한다.
		1: 다이얼로그 박스의 확인버튼을 눌렀을 경우 리턴 되는 값
		2: 다이얼로그 박스의 취소버튼을 눌렀을 경우 리턴 되는 값
		-1:  실행시 에러가 발생 하였을 경우 리턴 되는 값 (에러의 경우는 확인해보지 못했음)
*/
func (p *IDHwpAction) PopupDialog(param *IDHwpParameterSet) {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "PopupDialog", param.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
	fmt.Println(res.Value())
	fmt.Println(reflect.TypeOf(res.Value()))
}

/*
액션을 실행한다.

	CreateSet, GetDefault, PopupDialog, Execute를 차례대로 부른 것과 같다.
*/
func (p *IDHwpAction) Run() {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Run")
	if err != nil {
		panic(err)
	}
}

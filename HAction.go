package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HAction:

	HAction : 한/글에서 특정 기능을 수행하기 위한 액션 오브젝트

	예) 아래와 같은 형태로 사용하는 것이 HAction을 사용하는 올바른 사용법이다.
	act := hwp.HAction() // 액션 생성
	para := hwp.HParameterSet().HPrint() // 파라미터 생성
	act.GetDefault("Print",para.HSet()) // 액션 & 파라미터 초기화
	para.NumCopy(3) // 인쇄 매수 3장 지정
	act.Execute("Print",para.HSet()) // 액션 수행
*/
type HAction struct {
	variant *ole.VARIANT
}

/*
지정된 액션이름과 HSet을 받아 액션을 초기화 한다.

	actname : 액션 이름
	object : HSet Object
*/
func (p *HAction) GetDefault(actname string, object *HSet) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "GetDefault", actname, object.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정된 액션이름과 HSet을 받아 액션을 수행 한다.

	actname : 액션 이름
	object : HSet Object
*/
func (p *HAction) Execute(actname string, object *HSet) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Execute", actname, object.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정된 액션이름과 HSet을 받아 액션의 다이얼로그를생성한다.

	actname : 액션 이름
	object : HSet Object
*/
func (p *HAction) PopupDialog(actname string, object *HSet) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "PopupDialog", actname, object.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
GetDefault, PopupDialog, Execute를 동시에 수행하도록 한다.
*/
func (p *HAction) Run(actname string) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Run", actname)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

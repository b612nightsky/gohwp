package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HSet:

	HSet : ParameterSet Item 데이터들의 집합
*/
type HSet struct {
	variant *ole.VARIANT
}

/*
지정된 아이템이름을 가진 데이터에 VARIANT값을 대입한다.
*/
func (p *HSet) SetItem(bstr string, value interface{}) {
	switch v := value.(type) {
	case int:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case int8:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case int16:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case int32:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case uint:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case uint8:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case uint16:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case uint32:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, i)
		if err != nil {
			panic(err)
		}
	case string:
		s := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, s)
		if err != nil {
			panic(err)
		}
	case bool:
		b := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", bstr, b)
		if err != nil {
			panic(err)
		}
	default:
		panic("SetItem(bstr string, value interface{}) 함수의 bstr: \"" + bstr + "\"의 value 타입 확인 필요")
	}
}

package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IDHwpParameterArray:

	IDHwpParameterArray: Parameter Set의 아이템으로 배열을 표현하는데 사용된다. 일반적인 Method의 독립적인 인수로 사용되는 일은 없고, Parameter Set의 아이템으로만 사용된다.
	(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpParameterArray와 동일)
*/
type IDHwpParameterArray struct {
	variant *ole.VARIANT
}

/*
배열의 크기를 나타낸다.
*/
func (p *IDHwpParameterArray) Count() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
Parameter Set인지 연부를 나타낸다. (읽기 전용)
*/
func (p *IDHwpParameterArray) IsSet() bool {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "IsSet")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
동일한 크기와 데이터를 갖는 ParameterArray 개체를 복사하여 돌려준다
*/
func (p *IDHwpParameterArray) Clone() *IDHwpParameterArray {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Clone")
	if err != nil {
		panic(err)
	}
	var clone IDHwpParameterArray
	clone.variant = res
	return &clone
}

/*
배열을 복사한다.
*/
func (p *IDHwpParameterArray) Copy(srcarray *IDHwpParameterArray) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Copy", srcarray.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
}

/*
지정한 원소의 값을 리턴한다.
*/
func (p *IDHwpParameterArray) Item(index int) interface{} {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", index)
	if err != nil {
		panic(err)
	}
	return res.Value()
}

/*
지정한 원소의 값을 설정한다.

	index : 원소의 인덱스, 0부터 시작
	value : 원소의 값
*/
func (p *IDHwpParameterArray) SetItem(index int, value interface{}) {
	switch v := value.(type) {
	case int:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case int8:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case int16:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case int32:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case uint:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case uint8:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case uint16:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case uint32:
		i := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, i)
		if err != nil {
			panic(err)
		}
	case string:
		s := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, s)
		if err != nil {
			panic(err)
		}
	case bool:
		b := v
		_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", index, b)
		if err != nil {
			panic(err)
		}
	default:
		panic("SetItem(index int, value interface{}) 함수의 value 타입 확인 필요")
	}
}

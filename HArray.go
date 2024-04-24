package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HArray

	HArray : ParameterArray의 데이터 집합
*/
type HArray struct {
	variant *ole.VARIANT
}

/*
Array에 크기를 지정하거나 얻을 수 있다.
*/
func (p *HArray) Count(v ...int32) int {
	if len(v) > 0 {
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Count", v[0])
		if err != nil {
			panic(err)
		}
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
Array에 지정된 index에 해당하는 값을 VARIANT로 지정하거나 얻을 수 있다.
*/
func (p *HArray) Item(idx int32, value ...interface{}) interface{} {
	if len(value) > 0 {
		switch v := value[0].(type) {
		case int:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case int8:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case int16:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case int32:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case uint:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case uint8:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case uint16:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case uint32:
			i := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, i)
			if err != nil {
				panic(err)
			}
		case string:
			s := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, s)
			if err != nil {
				panic(err)
			}
		case bool:
			b := v
			_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Item", idx, b)
			if err != nil {
				panic(err)
			}
		default:
			panic("Item(idx int32, value ...interface{}) 함수의 idx: \"" + string(idx) + "\"의 value 타입 확인 필요")
		}
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Item", idx)
	if err != nil {
		panic(err)
	}
	return res.Value()
}

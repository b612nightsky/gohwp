package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IDHwpParameterSet:

	IDHwpParameterSet: 오브젝트 또는 액션의 실행에 필요한 정보를 주고 받을 수 있도록 하기 위한 오브젝트
	(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpParameterSet과 동일)
*/
type IDHwpParameterSet struct {
	variant *ole.VARIANT
}

/*
현재 존재하는 아이템의 개수를 나타낸다. (읽기 전용)
*/
func (p *IDHwpParameterSet) Count() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Count")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
Parameter Set인지 연부를 나타낸다. (읽기 전용)
*/
func (p *IDHwpParameterSet) IsSet() bool {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "IsSet")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
Parameter Set의 ID를 나타낸다. (읽기 전용)
*/
func (p *IDHwpParameterSet) SetID() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "SetID")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
동일한 데이터를 가진 Parameter Set을 복사하여 리턴한다.
*/
func (p *IDHwpParameterSet) Clone() *IDHwpParameterSet {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Clone")
	if err != nil {
		panic(err)
	}
	var clone IDHwpParameterSet
	clone.variant = res
	return &clone
}

/*
아이템으로 Parameter Array 타입의 배열을 생성한다.

	itemid : 아이템 ID
	count : 생성할 배열의 초기 크기
	return : 생성된 parameter array 오브젝트 (IDHwpParameterArray)
*/
func (p *IDHwpParameterSet) CreateItemArray(itemid string, count int) *IDHwpParameterArray {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "CreateItemArray", itemid, count)
	if err != nil {
		panic(err)
	}
	var dhpa IDHwpParameterArray
	dhpa.variant = res
	return &dhpa
}

/*
아이템으로 Parameter Array 타입의 배열을 생성한다.

	itemid : 아이템 ID
	setid : 생성할 Parameter Set ID
	return : 생성된 서브 parameter Set 오브젝트 (IDHwpParameterSet)
	ParameterSet 내부에 아이템으로 또 다른 Parameter Set을 가지는 서브셋의 개념이다.
*/
func (p *IDHwpParameterSet) CreateItemSet(itemid, setid string) *IDHwpParameterSet {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "CreateItemSet", itemid, setid)
	if err != nil {
		panic(err)
	}
	var dhps IDHwpParameterSet
	dhps.variant = res
	return &dhps
}

/*
두 Parameter Set에 공통적으로 존재하고, 값도 동일한 아이템만으로 구성된 intersection Set을 구한다.

	this와 srcset의 intersection이 this에 저장된다.
*/
func (p *IDHwpParameterSet) GetIntersection(srcset *IDHwpParameterSet) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "GetIntersection", srcset.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
}

/*
두 Parameter Set의 내용이 동일한 값을 가지고 있는지 검사한다.

	this와 srcset의 비교한 결과를 리턴한다.
	return : 동일하면 true, 다르면 false
*/
func (p *IDHwpParameterSet) IsEquivalent(srcset *IDHwpParameterSet) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "IsEquivalent", srcset.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정한 아이템의 값을 리턴한다.

	itemid : 아이템 ID
	return : 아이템의 값 interface{}로 반환, 형변환 필요
*/
func (p *IDHwpParameterSet) Item(itemid string) interface{} {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", itemid)
	if err != nil {
		panic(err)
	}
	return res.Value()
}

/*
지정한 아이템이 존재하는지 검사한다.

	itemid : 아이템 ID
	return : 존재하면 TRUE, 존재하지 않으면 FALSE
*/
func (p *IDHwpParameterSet) ItemExist(itemid string) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "ItemExist", itemid)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
두 Parameter Set의 내용을 병합한다.

	this와 srcset이 병합되어 this에 저장된다.
	결과는 "this의 모든 아이템 + srcset에만 존재하는 아이템"이다.
*/
func (p *IDHwpParameterSet) Merge(srcset *IDHwpParameterSet) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Merge", srcset.variant.ToIDispatch())
	if err != nil {
		panic(err)
	}
}

/*
Parameter Set을 초기화 한다.

	setid : 새로 적용할 Set ID
	이미 존재하는 Parameter Set 오브젝트를 이용해 새로운 타입의 Parameter Set으로 초기화하여 재사용하는 목적에 사용된다.
*/
func (p *IDHwpParameterSet) RemoveAll(setid string) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "RemoveAll", setid)
	if err != nil {
		panic(err)
	}
}

/*
지정한 아이템을 삭제한다.

	itemid : 아이템 ID
*/
func (p *IDHwpParameterSet) RemoveItem(itemid string) {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "RemoveItem", itemid)
	if err != nil {
		panic(err)
	}
}

/*
지정한 아이템의 값을 설정한다.
*/
func (p *IDHwpParameterSet) SetItem(bstr string, value interface{}) {
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

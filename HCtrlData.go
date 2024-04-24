package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HCtrlData:

	HCtrlData :
*/
type HCtrlData struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCtrlData) HSet() *HSet {
	if p.hset == nil {
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSet")
		if err != nil {
			panic(err)
		}
		var hs HSet
		hs.variant = res
		p.hset = &hs
	}
	return p.hset
}

/*
사용자가 지정한 컨트롤의 이름.
*/
func (p *HCtrlData) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

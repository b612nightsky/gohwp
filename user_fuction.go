package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HWP 창 보이기

	활성화 된 HWP 작업창 보이기(true), 백그라운드 실행(false)
*/
func (hwp *IHwpObject) ShowWindow(tf bool) {
	hwp.XHwpWindows().Active_XHwpWindow().Visible(tf)
}

/*
액션 실행: 액션 이름과 파라미터 세트를 동시에 함수 인자로 넘겨 실행

ActionWithParameters(액션이름, "옵션명1", 값1, "옵션명2", "값2", ...)
*/
func (hwp *IHwpObject) ActionWithParameters(actName string, paras ...interface{}) {
	act := hwp.CreateAction(actName)
	para := act.CreateSet()
	act.GetDefault(para)
	chk := len(paras) % 2 // 파라미터 갯수가 짝수인지 확인
	switch chk {
	case 1:
		txt := "ActionWithParameters(actName string, paras ...interface{}) 함수 인자 중\n"
		txt += "paras는 \"옵션명\", 옵션값, ... pair로 총 갯수는 짝수가 되어야 함"
		panic(txt)
	default:
		for i := 0; i < len(paras)/2; i++ {
			switch v := paras[i*2].(type) {
			case string:
				s := v
				para.SetItem(s, paras[i*2+1])
			default:
				txt := "ActionWithParameters(actName string, paras ...interface{}) 함수 인자 중\n"
				txt += "paras는 \"옵션명\", 옵션값, ... 옵션명은 string 값을 가져야 함"
				panic(txt)
			}
		}
	}
	act.Execute(para)
}

/*
글꼴(Font) 설정

	fontname: 글꼴 이름
	시스템 글꼴만 가능: (예) 바탕(체), 돋움(체), 굴림(체), 궁서(체), 맑은 고딕
*/
func (hwp *IHwpObject) SetFont(fontname string) {
	hwp.ActionWithParameters("CharShape",
		"FaceNameHangul", fontname,
		"FaceNameLatin", fontname,
		"FaceNameHanja", fontname,
		"FaceNameJapanese", fontname,
		"FaceNameOther", fontname,
		"FaceNameSymbol", fontname,
		"FaceNameUser", fontname)
}

/*
텍스트 입력
*/
func (hwp *IHwpObject) InsertText(text string) {
	hwp.ActionWithParameters("InsertText", "Text", text)
}

/*
문서 중에서 텍스트를 얻는다.

	GetText(*string) int 를 GetTextU() string, int 로 변형
*/
func (hwp *IHwpObject) GetTextU() (string, int) {
	text := ""
	res := hwp.GetText(&text)
	return text, res
}

/*
현재 캐럿의 위치 정보를 얻어온다.

	GetPos(list, para, pos *int)를 GetPosU() (list, para, pos int)로 변형
	list: 리스트 ID, para: 문단 ID, pos: 문단 내 글자 위치
*/
func (hwp *IHwpObject) GetPosU() (list int, para int, pos int) {
	var ilist, ipara, ipos int
	_, err := oleutil.CallMethod(hwp.dispatch, "GetPos", &ilist, &ipara, &ipos)
	if err != nil {
		panic(err)
	}
	return ilist, ipara, ipos
}

/*
현재 설정된 블록의 위치정보를 얻어온다.

	slist: 설정된 블록의 시작 리스트 아이디.
	spara: 설정된 블록의 시작 문단 아이디.
	spos: 설정된 블록의 문단 내 시작 글자 단위 위치.
	elist: 설정된 블록의 끝 리스트 아이디.
	epara: 설정된 블록의 끝 문단 아이디.
	epos: 설정된 블록의 문단 내 끝 글자 단위 위치.
*/
func (hwp *IHwpObject) GetSelectedPosU() (slist, spara, spos, elist, epara, epos int) {
	var slisti, sparai, sposi, elisti, eparai, eposi int
	_, err := oleutil.CallMethod(hwp.dispatch, "GetSelectedPos", &slisti, &sparai, &sposi, &elisti, &eparai, &eposi)
	if err != nil {
		panic(err)
	}
	return slisti, sparai, sposi, elisti, eparai, eposi
}

/*
상태바에 나타날 정보를 알아낸다.

	seccnt: 총 구역, secno: 현재 구역, prnpageno: 쪽, colno: 단
	line: 줄, pos: 칸, over: (0: 삽입, 1: 수정), ctrlname
	KeyIndicator(seccnt, secno, prnpageno, colno, line, pos *int, over *int16, ctrlname *string) 함수 변형
*/
func (hwp *IHwpObject) KeyIndicatorU() (seccnt, secno, prnpageno, colno, line, pos, over int, ctrlname string) {
	var iseccnt, isecno, iprnpageno, icolno, iline, ipos int
	var iover int16
	var sctrlname string
	_, err := oleutil.CallMethod(hwp.dispatch, "KeyIndicator",
		&iseccnt, &isecno, &iprnpageno, &icolno, &iline, &ipos, &iover, &sctrlname)
	if err != nil {
		panic(err)
	}
	return iseccnt, isecno, iprnpageno, icolno, iline, ipos, int(iover), sctrlname
}

/*
HParameterSet의 아이템 H파라미터(HCharshape, HSecDef 등등)의 string 타입의 값을 읽어오거나 지정

	※ 내부적으로 사용될 뿐 직접 사용할 일은 없음.
*/
func funcParaSetString(v *ole.VARIANT, item string, value []string) string {
	if len(value) == 1 {
		oleutil.PutProperty(v.ToIDispatch(), item, value[0])
	}
	res, err := oleutil.GetProperty(v.ToIDispatch(), item)
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
HParameterSet의 아이템 H파라미터(HCharshape, HSecDef 등등)의 int 타입의 값을 읽어오거나 지정

	※ 내부적으로 사용될 뿐 직접 사용할 일은 없음.
*/
func funcParaSetInt[T int8 | int16 | int32 | int64 | uint8 | uint16 | uint32 | uint64 | int](v *ole.VARIANT, item string, value []T) int {
	if len(value) == 1 {
		oleutil.PutProperty(v.ToIDispatch(), item, value[0])
	}
	res, err := oleutil.GetProperty(v.ToIDispatch(), item)
	if err != nil {
		panic(err)
	}
	return int(res.Value().(T))
}

/*
HArray 받아오기

	※ 내부적으로 사용될 뿐 직접 사용할 일은 없음.
*/
func GetHArray(p *ole.VARIANT, itemID string) *HArray {
	res, err := oleutil.GetProperty(p.ToIDispatch(), itemID)
	if err != nil {
		panic(err)
	}
	var ha HArray
	ha.variant = res
	return &ha
}

/*
PIT_ARRAY 생성

	PIT_ARRAY를 위해 필요
	※ 내부적으로 사용될 뿐 직접 사용할 일은 없음.
*/
func CreateItemArray(p *ole.VARIANT, itemID string, num int32) {
	_, err := oleutil.CallMethod(p.ToIDispatch(), "CreateItemArray", itemID, num)
	if err != nil {
		panic(err)
	}
}

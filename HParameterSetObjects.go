package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HSecDef:

	HSecDef :  구역의 속성
*/
type HSecDef struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSecDef) HSet() *HSet {
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
글자 방향
*/
func (p *HSecDef) TextDirection(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextDirection", v)
}

/*
구역 나눔으로 새 페이지가 생길 때의 페이지 번호 적용 옵션

	0 = 이어서, 1 = 홀수, 2 = 짝수, 3 = 임의 값
*/
func (p *HSecDef) StartsOn(v ...uint16) int {
	return funcParaSetInt(p.variant, "StartsOn", v)
}

/*
세로로 줄맞춤을 할지 여부.

	0 = off, 1 - n = 간격을 HWPUNIT 단위로 지정.
*/
func (p *HSecDef) LineGrid(v ...int32) int {
	return funcParaSetInt(p.variant, "LineGrid", v)
}

/*
가로로 칸맞춤을 할지 여부

	0 = off, 1 - n = 간격을 HWPUNIT 단위로 지정.
*/
func (p *HSecDef) CharGrid(v ...int32) int {
	return funcParaSetInt(p.variant, "CharGrid", v)
}

/*
용지 설정 정보 서브셋(HPageDef)
*/
func (p *HSecDef) PageDef(v ...*HPageDef) *HPageDef {
	var para HPageDef
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "PageDef", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "PageDef")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
빈 줄 감춤 on/off
*/
func (p *HSecDef) HideEmptyLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideEmptyLine", v)
}

/*
동일한 페이지에서 서로 다른 단 사이의 간격(HWPUNIT)
*/
func (p *HSecDef) SpaceBetweenColumns(v ...int32) int {
	return funcParaSetInt(p.variant, "SpaceBetweenColumns", v)
}

/*
기본 탭 간격(URC)
*/
func (p *HSecDef) TabStop(v ...int32) int {
	return funcParaSetInt(p.variant, "TabStop", v)
}

/*
각주 모양 서브셋(HFootnoteShape)
*/
func (p *HSecDef) FootnoteShape(v ...*HFootnoteShape) *HFootnoteShape {
	var para HFootnoteShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FootnoteShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FootnoteShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
미주 모양 서브셋(HFootnoteShape)
*/
func (p *HSecDef) EndnoteShape(v ...*HFootnoteShape) *HFootnoteShape {
	var para HFootnoteShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "EndnoteShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "EndnoteShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
구역의 첫 쪽에만 머리말 감추기 옵션 on/off
*/
func (p *HSecDef) HideHeader(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideHeader", v)
}

/*
구역의 첫 쪽에만 꼬리말 감추기 옵션 on/off
*/
func (p *HSecDef) HideFooter(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideFooter", v)
}

/*
구역의 첫 쪽에만 바탕쪽 감추기 옵션 on/off
*/
func (p *HSecDef) HideMasterPage(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideMasterPage", v)
}

/*
구역의 첫 쪽에만 테두리 감추기 옵션 on/off
*/
func (p *HSecDef) HideBorder(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideBorder", v)
}

/*
구역의 첫 쪽에만 배경 감추기 옵션 on/off
*/
func (p *HSecDef) HideFill(v ...uint16) int {
	return funcParaSetInt(p.variant, "HideFill", v)
}

/*
구역의 첫 쪽에만 쪽 번호 감추기 옵션 on/off
*/
func (p *HSecDef) HidePageNumPos(v ...uint16) int {
	return funcParaSetInt(p.variant, "HidePageNumPos", v)
}

/*
구역의 첫 쪽에만 테두리 표시 옵션 on/off
*/
func (p *HSecDef) FirstBorder(v ...uint16) int {
	return funcParaSetInt(p.variant, "FirstBorder", v)
}

/*
구역의 첫 쪽에만 배경 표시 옵션 on/off
*/
func (p *HSecDef) FirstFill(v ...uint16) int {
	return funcParaSetInt(p.variant, "FirstFill", v)
}

/*
개요 번호 서브셋(HNumberingShape)
*/
func (p *HSecDef) OutlineShape(v ...*HNumberingShape) *HNumberingShape {
	var para HNumberingShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "OutlineShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "OutlineShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
쪽 테두리/배경 서브셋 (양쪽) (HPageBorderFill)
*/
func (p *HSecDef) PageBorderFillBoth(v ...*HPageBorderFill) *HPageBorderFill {
	var para HPageBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "PageBorderFillBoth", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "PageBorderFillBoth")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
쪽 테두리/배경 서브셋 (짝수) (HPageBorderFill)
*/
func (p *HSecDef) PageBorderFillEven(v ...*HPageBorderFill) *HPageBorderFill {
	var para HPageBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "PageBorderFillEven", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "PageBorderFillEven")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
쪽 테두리/배경 서브셋 (홀수) (HPageBorderFill)
*/
func (p *HSecDef) PageBorderFillOdd(v ...*HPageBorderFill) *HPageBorderFill {
	var para HPageBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "PageBorderFillOdd", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "PageBorderFillOdd")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
쪽 시작 번호

	0 = 앞 구역에 이어, n = 새 번호로 시작
*/
func (p *HSecDef) PageNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "PageNumber", v)
}

/*
그림 시작 번호

	0 = 앞 구역에 이어, n = 새 번호로 시작
*/
func (p *HSecDef) FigureNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "FigureNumber", v)
}

/*
표 시작 번호

	0 = 앞 구역에 이어, n = 새 번호로 시작
*/
func (p *HSecDef) TableNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "TableNumber", v)
}

/*
수식 시작 번호

	0 = 앞 구역에 이어, n = 새 번호로 시작
*/
func (p *HSecDef) EquationNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "EquationNumber", v)
}

/*
원고지 방식의 포매팅. CHAR_GRID가 지정되어야 함.
*/
func (p *HSecDef) WongojiFormat(v ...uint16) int {
	return funcParaSetInt(p.variant, "WongojiFormat", v)
}

/*
메모 모양 서브셋(HMemoShape)
*/
func (p *HSecDef) MemoShape(v ...*HMemoShape) *HMemoShape {
	var para HMemoShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "MemoShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "MemoShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HPageDef:

	HPageDef :  구역 내의 용지 설정 속성
*/
type HPageDef struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPageDef) HSet() *HSet {
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
용지 가로 크기 (HWPUNIT)
*/
func (p *HPageDef) PaperWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "PaperWidth", v)
}

/*
용지 세로 크기
*/
func (p *HPageDef) PaperHeight(v ...int32) int {
	return funcParaSetInt(p.variant, "PaperHeight", v)
}

/*
용지 방향

	0 = 좁게, 1 = 넓게
*/
func (p *HPageDef) Landscape(v ...uint16) int {
	return funcParaSetInt(p.variant, "Landscape", v)
}

/*
왼쪽 마진 (HWPUNIT)
*/
func (p *HPageDef) LeftMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "LeftMargin", v)
}

/*
오른쪽 마진 (HWPUNIT)
*/
func (p *HPageDef) RightMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "RightMargin", v)
}

/*
위 마진 (HWPUNIT)
*/
func (p *HPageDef) TopMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "TopMargin", v)
}

/*
아래 마진 (HWPUNIT)
*/
func (p *HPageDef) BottomMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "BottomMargin", v)
}

/*
머리말 길이 (HWPUNIT)
*/
func (p *HPageDef) HeaderLen(v ...int32) int {
	return funcParaSetInt(p.variant, "HeaderLen", v)
}

/*
꼬리말 길이 (HWPUNIT)
*/
func (p *HPageDef) FooterLen(v ...int32) int {
	return funcParaSetInt(p.variant, "FooterLen", v)
}

/*
제본 여백 (HWPUNIT)
*/
func (p *HPageDef) GutterLen(v ...int32) int {
	return funcParaSetInt(p.variant, "GutterLen", v)
}

/*
편집 방법

	0 = 한쪽 편집, 1 = 맞쪽 편집, 2 = 위로 넘기기
*/
func (p *HPageDef) GutterType(v ...uint16) int {
	return funcParaSetInt(p.variant, "GutterType", v)
}

/*
HColDef:

	HColDef :  단 정의 속성
*/
type HColDef struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HColDef) HSet() *HSet {
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
단 종류

	0 = 보통 다단, 1 = 배분 다단, 2 = 평행 다단
*/
func (p *HColDef) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
단 개수. 1-255까지
*/
func (p *HColDef) Count(v ...uint16) int {
	return funcParaSetInt(p.variant, "Count", v)
}

/*
0 = 단 너비 각자 지정, 1 = 단 너비 동일
*/
func (p *HColDef) SameSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "SameSize", v)
}

/*
단 사이 간격 (HWPUNIT)

	SAME_SIZE가 1일 때만 사용된다.
*/
func (p *HColDef) SameGap(v ...int32) int {
	return funcParaSetInt(p.variant, "SameGap", v)
}

/*
각 단의 너비와 간격

	※ CreateItemArray("WidthGap", Count*2-1) 먼저 실행해야함
	col*2 = 단의 폭,
	col*2 + 1 = 단 사이 간격 영역
	전체의 폭을 Column ratio base (=32768)로 보았을 때의 비율로 환산한다.
	SameSize가 0일 때만 사용된다.
	배열의 아이템의 개수는 Count*2-1과 같아야 한다.
*/
func (p *HColDef) WidthGap() *HArray {
	return GetHArray(p.variant, "WidthGap")
}

/*
PIT_ARRAY 생성

	WidthGap을 위해 필요
*/
func (p *HColDef) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
단 방향 지정

	0 = 왼쪽부터, 1 = 오른쪽부터, 2 = 맞쪽
*/
func (p *HColDef) Layout(v ...uint16) int {
	return funcParaSetInt(p.variant, "Layout", v)
}

/*
선 종류
*/
func (p *HColDef) LineType(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineType", v)
}

/*
선 굵기
*/
func (p *HColDef) LineWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineWidth", v)
}

/*
선 색깔(COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HColDef) LineColor(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineColor", v)
}

/*
HCharShape:

	HCharShape :  글자 모양
*/
type HCharShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCharShape) HSet() *HSet {
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
글꼴 이름 (한글)
*/
func (p *HCharShape) FaceNameHangul(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameHangul", v)
}

/*
글꼴 이름 (영문)
*/
func (p *HCharShape) FaceNameLatin(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameLatin", v)
}

/*
글꼴 이름 (한자)
*/
func (p *HCharShape) FaceNameHanja(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameHanja", v)
}

/*
글꼴 이름 (일본어)
*/
func (p *HCharShape) FaceNameJapanese(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameJapanese", v)
}

/*
글꼴 이름 (기타)
*/
func (p *HCharShape) FaceNameOther(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameOther", v)
}

/*
글꼴 이름 (기호)
*/
func (p *HCharShape) FaceNameSymbol(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameSymbol", v)
}

/*
글꼴 이름 (사용자)
*/
func (p *HCharShape) FaceNameUser(v ...string) string {
	return funcParaSetString(p.variant, "FaceNameUser", v)
}

/*
폰트 종류 (한글)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeHangul(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeHangul", v)
}

/*
폰트 종류 (영문)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeLatin(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeLatin", v)
}

/*
폰트 종류 (한자)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeHanja(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeHanja", v)
}

/*
폰트 종류 (일본어)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeJapanese(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeJapanese", v)
}

/*
폰트 종류 (기타)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeOther(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeOther", v)
}

/*
폰트 종류 (기호)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeSymbol(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeSymbol", v)
}

/*
폰트 종류 (사용자)

	0 = don't care, 1 = TTF, 2 = HFT
*/
func (p *HCharShape) FontTypeUser(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontTypeUser", v)
}

/*
각 언어별 크기 비율. (한글) 10% - 250%
*/
func (p *HCharShape) SizeHangul(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeHangul", v)
}

/*
각 언어별 크기 비율. (영문) 10% - 250%
*/
func (p *HCharShape) SizeLatin(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeLatin", v)
}

/*
각 언어별 크기 비율. (한자) 10% - 250%
*/
func (p *HCharShape) SizeHanja(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeHanja", v)
}

/*
각 언어별 크기 비율. (일본어) 10% - 250%
*/
func (p *HCharShape) SizeJapanese(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeJapanese", v)
}

/*
각 언어별 크기 비율. (기타) 10% - 250%
*/
func (p *HCharShape) SizeOther(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeOther", v)
}

/*
각 언어별 크기 비율. (기호) 10% - 250%
*/
func (p *HCharShape) SizeSymbol(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeSymbol", v)
}

/*
각 언어별 크기 비율. (사용자) 10% - 250%
*/
func (p *HCharShape) SizeUser(v ...uint16) int {
	return funcParaSetInt(p.variant, "SizeUser", v)
}

/*
각 언어별 장평 비율. (한글) 10% - 200%
*/
func (p *HCharShape) RatioHangul(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioHangul", v)
}

/*
각 언어별 장평 비율. (영문) 10% - 200%
*/
func (p *HCharShape) RatioLatin(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioLatin", v)
}

/*
각 언어별 장평 비율. (한자) 10% - 200%
*/
func (p *HCharShape) RatioHanja(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioHanja", v)
}

/*
각 언어별 장평 비율. (일본어) 10% - 200%
*/
func (p *HCharShape) RatioJapanese(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioJapanese", v)
}

/*
각 언어별 장평 비율. (기타) 10% - 200%
*/
func (p *HCharShape) RatioOther(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioOther", v)
}

/*
각 언어별 장평 비율. (기호) 10% - 200%
*/
func (p *HCharShape) RatioSymbol(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioSymbol", v)
}

/*
각 언어별 장평 비율. (사용자) 10% - 200%
*/
func (p *HCharShape) RatioUser(v ...uint16) int {
	return funcParaSetInt(p.variant, "RatioUser", v)
}

/*
각 언어별 자간. (한글) -50% - 50%
*/
func (p *HCharShape) SpacingHangul(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingHangul", v)
}

/*
각 언어별 자간. (영문) -50% - 50%
*/
func (p *HCharShape) SpacingLatin(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingLatin", v)
}

/*
각 언어별 자간. (한자) -50% - 50%
*/
func (p *HCharShape) SpacingHanja(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingHanja", v)
}

/*
각 언어별 자간. (일본어) -50% - 50%
*/
func (p *HCharShape) SpacingJapanese(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingJapanese", v)
}

/*
각 언어별 자간. (기타) -50% - 50%
*/
func (p *HCharShape) SpacingOther(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingOther", v)
}

/*
각 언어별 자간. (기호) -50% - 50%
*/
func (p *HCharShape) SpacingSymbol(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingSymbol", v)
}

/*
각 언어별 자간. (사용자) -50% - 50%
*/
func (p *HCharShape) SpacingUser(v ...int16) int {
	return funcParaSetInt(p.variant, "SpacingUser", v)
}

/*
각 언어별 오프셋. (한글) -100% - 100%
*/
func (p *HCharShape) OffsetHangul(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetHangul", v)
}

/*
각 언어별 오프셋. (영문) -100% - 100%
*/
func (p *HCharShape) OffsetLatin(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetLatin", v)
}

/*
각 언어별 오프셋. (한자) -100% - 100%
*/
func (p *HCharShape) OffsetHanja(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetHanja", v)
}

/*
각 언어별 오프셋. (일본어) -100% - 100%
*/
func (p *HCharShape) OffsetJapanese(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetJapanese", v)
}

/*
각 언어별 오프셋. (기타) -100% - 100%
*/
func (p *HCharShape) OffsetOther(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetOther", v)
}

/*
각 언어별 오프셋. (기호) -100% - 100%
*/
func (p *HCharShape) OffsetSymbol(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetSymbol", v)
}

/*
각 언어별 오프셋. (사용자) -100% - 100%
*/
func (p *HCharShape) OffsetUser(v ...int16) int {
	return funcParaSetInt(p.variant, "OffsetUser", v)
}

/*
Bold on/off
*/
func (p *HCharShape) Bold(v ...uint16) int {
	return funcParaSetInt(p.variant, "Bold", v)
}

/*
Italic on/off
*/
func (p *HCharShape) Italic(v ...uint16) int {
	return funcParaSetInt(p.variant, "Italic", v)
}

/*
Small Caps on/off
*/
func (p *HCharShape) SmallCaps(v ...uint16) int {
	return funcParaSetInt(p.variant, "SmallCaps", v)
}

/*
Emboss on/off
*/
func (p *HCharShape) Emboss(v ...uint16) int {
	return funcParaSetInt(p.variant, "Emboss", v)
}

/*
Engrave on/off
*/
func (p *HCharShape) Engrave(v ...uint16) int {
	return funcParaSetInt(p.variant, "Engrave", v)
}

/*
Superscript on/off
*/
func (p *HCharShape) SuperScript(v ...uint16) int {
	return funcParaSetInt(p.variant, "SuperScript", v)
}

/*
Subscript on/off
*/
func (p *HCharShape) SubScript(v ...uint16) int {
	return funcParaSetInt(p.variant, "SubScript", v)
}

/*
밑줄 종류

	0 = none, 1 = bottom, 2 = center, 3 = top
*/
func (p *HCharShape) UnderlineType(v ...uint16) int {
	return funcParaSetInt(p.variant, "UnderlineType", v)
}

/*
외곽선 종류

	0 = none 1 = solid 2 = dot 3 = thick 4 = dash 5 = dashdot 6 = dashdotdot
*/
func (p *HCharShape) OutLineType(v ...uint16) int {
	return funcParaSetInt(p.variant, "OutLineType", v)
}

/*
그림자 종류

	0 = none, 1 = drop, 2 = continuous
*/
func (p *HCharShape) ShadowType(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShadowType", v)
}

/*
글자색 (COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCharShape) TextColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "TextColor", v)
}

/*
음영색 (COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCharShape) ShadeColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "ShadeColor", v)
}

/*
밑줄 모양 선 종류 - 1 과 동일.
*/
func (p *HCharShape) UnderlineShape(v ...uint16) int {
	return funcParaSetInt(p.variant, "UnderlineShape", v)
}

/*
밑줄색 (COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCharShape) UnderlineColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "UnderlineColor", v)
}

/*
그림자 간격 (X 방향) -100% - 100%
*/
func (p *HCharShape) ShadowOffsetX(v ...int16) int {
	return funcParaSetInt(p.variant, "ShadowOffsetX", v)
}

/*
그림자 간격 (Y 방향) -100% - 100%
*/
func (p *HCharShape) ShadowOffsetY(v ...int16) int {
	return funcParaSetInt(p.variant, "ShadowOffsetY", v)
}

/*
그림자 색 (COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCharShape) ShadowColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "ShadowColor", v)
}

/*
취소선 종류

	0 = none 1 = red single 2 = red double 3 = text single 4 = text double
*/
func (p *HCharShape) StrikeOutType(v ...uint16) int {
	return funcParaSetInt(p.variant, "StrikeOutType", v)
}

/*
강조점 종류

	0 = none, 1 = 검정 동그라미, 2 = 속 빈 동그라미
*/
func (p *HCharShape) DiacSymMark(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiacSymMark", v)
}

/*
글꼴에 어울리는 빈칸 on/off
*/
func (p *HCharShape) UseFontSpace(v ...uint16) int {
	return funcParaSetInt(p.variant, "UseFontSpace", v)
}

/*
커닝 on/off
*/
func (p *HCharShape) UseKerning(v ...uint16) int {
	return funcParaSetInt(p.variant, "UseKerning", v)
}

/*
글자 크기 (HWPUNIT)
*/
func (p *HCharShape) Height(v ...int32) int {
	return funcParaSetInt(p.variant, "Height", v)
}

/*
HViewProperties:

	HViewProperties :  뷰의 속성 개체
*/
type HViewProperties struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HViewProperties) HSet() *HSet {
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
뷰 옵션 플랙. 여러개를 OR연산하여 지정할 수 있음

	1 = off : 쪽윤곽, on : 기본 보기 2 = 공백과 폭이 없는 컨트 롤을 기호로 4 = 문단 마크 기호로 8 = 안내선 16 = 그리기 격자 32 = 그림 감춤
*/
func (p *HViewProperties) OptionFlag(v ...uint32) int {
	return funcParaSetInt(p.variant, "OptionFlag", v)
}

/*
화면 확대 종류

	0 = 사용자 정의 1 = 쪽 맞춤 2 = 폭 맞춤 3 = 여러 쪽
*/
func (p *HViewProperties) ZoomType(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomType", v)
}

/*
화면 확대 종류가 “사용자 정의”인 경우 화면 확대 비율, 10% ~ 500%
*/
func (p *HViewProperties) ZoomRatio(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomRatio", v)
}

/*
화면 확대 종류가 “여러 쪽”인 경우 가로 페이지 수. 1 ~ 8
*/
func (p *HViewProperties) ZoomCntX(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomCntX", v)
}

/*
화면 확대 종류가 “여러 쪽”인 경우 세로 페이지 수. 1 ~ 8
*/
func (p *HViewProperties) ZoomCntY(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomCntY", v)
}

/*
화면 확대 종류가 " 맞 쪽" 인 경우
*/
func (p *HViewProperties) ZoomMirror(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomMirror", v)
}

/*
HAppState:

	HAppState :  어플리케이션 글로벌 상태 정보 개체
*/
type HAppState struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HAppState) HSet() *HSet {
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
마우스 개체 핸들러 타입

	0 = Normal 1= 개체 생성  2 = 개체 선택 3 = 표 4 = 갈무리 5 = 양식개체 생성 6 = 형광펜 7 = 스크립트 매크로 기록
*/
func (p *HAppState) HandlerType(v ...uint32) int {
	return funcParaSetInt(p.variant, "HandlerType", v)
}

/*
개체 생성 모드

	0 = 핸들러 취소 1 = 하나 생성 2 = 연속 생성
*/
func (p *HAppState) CreationMode(v ...uint16) int {
	return funcParaSetInt(p.variant, "CreationMode", v)
}

/*
그리기 개체 생성 ID

	0 = default 1 = 선 2 = 사각형 3 = 원 4 = 호 5 = 다각형 6 = 곡선 7 = 글상자 8 = 펜으로 그리기 9 = 그림 10 = 그리기 마당 11 = 표 그리기 (1 X 1 셀)  12 = 표 지우기 13 = 표 만들기 (n X n 셀)
*/
func (p *HAppState) CreationID(v ...uint16) int {
	return funcParaSetInt(p.variant, "CreationID", v)
}

/*
키매크로 상태

	0x00 = 매크로 기록 중지  0x01 = 매크로 기록 일시 중지 0x02 = 매크로 실행 0x04 = 매크로 기록
*/
func (p *HAppState) KeyMacro(v ...uint32) int {
	return funcParaSetInt(p.variant, "KeyMacro", v)
}

/*
입력 상태

	0 = 삽입 1 = 수정
*/
func (p *HAppState) Overwrite(v ...uint16) int {
	return funcParaSetInt(p.variant, "Overwrite", v)
}

/*
프레젠테이션 상태

	0 = 보통 1 = 실
*/
func (p *HAppState) Presentation(v ...uint16) int {
	return funcParaSetInt(p.variant, "Presentation", v)
}

/*
스크립트 매크로 상태

	0x00 = 매크로 기록 중지 0x01 = 매크로 기록 일시 중지 0x02 = 매크로 실행 0x04 = 매크로 기록
*/
func (p *HAppState) ScrMacro(v ...uint32) int {
	return funcParaSetInt(p.variant, "ScrMacro", v)
}

/*
HShapeObject:

	HShapeObject :  그리기 개체 속성 개체
*/
type HShapeObject struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HShapeObject) HSet() *HSet {
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
글자처럼 취급 on / off
*/
func (p *HShapeObject) TreatAsChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "TreatAsChar", v)
}

/*
줄 간격에 영향을 줄지 여부 on / off
*/
func (p *HShapeObject) AffectsLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "AffectsLine", v)
}

/*
세로 위치의 기준
*/
func (p *HShapeObject) VertRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertRelTo", v)
}

/*
VERT_REL_TO에 대한 상대적인 배열 방식
*/
func (p *HShapeObject) VertAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
VERT_REL_TO와 VERT_ALIGN을 기준점으로 한 상대적인 오프셋 값 HWPUNIT 단위
*/
func (p *HShapeObject) VertOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "VertOffset", v)
}

/*
가로 위치의 기준
*/
func (p *HShapeObject) HorzRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzRelTo", v)
}

/*
HORZ_REL_TO에 대한 상대적인 배열 방식
*/
func (p *HShapeObject) HorzAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzAlign", v)
}

/*
HORZ_REL_TO와 HORZ_ALIGN을 기준점으로 한 상대적인 오프셋 값 HWPUNIT 단위
*/
func (p *HShapeObject) HorzOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "HorzOffset", v)
}

/*
오브젝트의 세로 위치를 본문 영역으로 제한할지 여부 on / off
*/
func (p *HShapeObject) FlowWithText(v ...uint16) int {
	return funcParaSetInt(p.variant, "FlowWithText", v)
}

/*
다른 오브젝트와 겹치는 것을 허용할지 여부 on / off
*/
func (p *HShapeObject) AllowOverlap(v ...uint16) int {
	return funcParaSetInt(p.variant, "AllowOverlap", v)
}

/*
오브젝트 폭의 기준
*/
func (p *HShapeObject) WidthRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthRelTo", v)
}

/*
오브젝트 폭의 값
*/
func (p *HShapeObject) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
오브젝트 높이의 기준
*/
func (p *HShapeObject) HeightRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeightRelTo", v)
}

/*
오브젝트의 높이 값
*/
func (p *HShapeObject) Height(v ...int32) int {
	return funcParaSetInt(p.variant, "Height", v)
}

/*
보호 on / off
*/
func (p *HShapeObject) ProtectSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "ProtectSize", v)
}

/*
오브젝트 주위를 텍스트가 어떻게 흘러갈지 지정하는 옵션 * "TreatAsChar" = FALSE일 때만 사용됨

	0 = bound rect를 따라 1 = 좌, 우에는 텍스트를 배치하지 않음 2 = 글과 겹치게 하여 글 뒤로 3 = 글과 겹치게 하여 글 앞으로 4 = 오브젝트의 outline을 따라 5 = 오브젝트 내부의 빈 공간까지
*/
func (p *HShapeObject) TextWrap(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextWrap", v)
}

/*
오브젝트의 좌/우 어느쪽에 글을 배치할지 지정하는 옵션 * "TextWrap"가 0, 4, 5일 때만 사용된다.

	0 = 양쪽 1 = 왼쪽 2 = 오른쪽 3 = 큰쪽
*/
func (p *HShapeObject) TextFlow(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextFlow", v)
}

/*
오브젝트의 바깥 여백. (왼쪽) HWPUNIT 단위
*/
func (p *HShapeObject) OutsideMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginLeft", v)
}

/*
오브젝트의 바깥 여백. (오른쪽) HWPUNIT 단위
*/
func (p *HShapeObject) OutsideMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginRight", v)
}

/*
오브젝트의 바깥 여백. (위) HWPUNIT 단위
*/
func (p *HShapeObject) OutsideMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginTop", v)
}

/*
오브젝트의 바깥 여백. (아래) HWPUNIT 단위
*/
func (p *HShapeObject) OutsideMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginBottom", v)
}

/*
이 개체가 속하는 번호 범주

	0 = 없음, 1 = 그림, 2 = 표, 3 = 수식
*/
func (p *HShapeObject) NumberingType(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumberingType", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 폭의 값
*/
func (p *HShapeObject) LayoutWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutWidth", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 높이 값
*/
func (p *HShapeObject) LayoutHeight(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutHeight", v)
}

/*
오브젝트 선택 가능 on / off
*/
func (p *HShapeObject) Lock(v ...uint16) int {
	return funcParaSetInt(p.variant, "Lock", v)
}

/*
쪽 나눔 방지 II on/off
*/
func (p *HShapeObject) HoldAnchorObj(v ...uint16) int {
	return funcParaSetInt(p.variant, "HoldAnchorObj", v)
}

/*
개체가 존재 하는 페이지
*/
func (p *HShapeObject) PageNumber(v ...uint32) int {
	return funcParaSetInt(p.variant, "PageNumber", v)
}

/*
개체 Selection 상태 TRUE/FASLE
*/
func (p *HShapeObject) AdjustSelection(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustSelection", v)
}

/*
글상자로 TRUE/FASLE
*/
func (p *HShapeObject) AdjustTextbox(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustTextbox", v)
}

/*
앞개체 속성 따라가기 TRUE/FASLE
*/
func (p *HShapeObject) AdjustPrevObjAttr(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustPrevObjAttr", v)
}

/*
그리기 개체 레이아웃(HDrawLayOut)
*/
func (p *HShapeObject) ShapeDrawLayOut(v ...*HDrawLayOut) *HDrawLayOut {
	var para HDrawLayOut
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawLayOut", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawLayOut")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 선 속성(HDrawLineAttr)
*/
func (p *HShapeObject) ShapeDrawLineAttr(v ...*HDrawLineAttr) *HDrawLineAttr {
	var para HDrawLineAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawLineAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawLineAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 채우기 속성(HDrawFillAttr)
*/
func (p *HShapeObject) ShapeDrawFillAttr(v ...*HDrawFillAttr) *HDrawFillAttr {
	var para HDrawFillAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawFillAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawFillAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 그림 속성(HDrawImageAttr)
*/
func (p *HShapeObject) ShapeDrawImageAttr(v ...*HDrawImageAttr) *HDrawImageAttr {
	var para HDrawImageAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawImageAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawImageAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 사각형 속성(HDrawRectAttr)
*/
func (p *HShapeObject) ShapeDrawRectType(v ...*HDrawRectType) *HDrawRectType {
	var para HDrawRectType
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawRectType", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawRectType")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 호 속성(HDrawArcAttr)
*/
func (p *HShapeObject) ShapeDrawArcType(v ...*HDrawArcType) *HDrawArcType {
	var para HDrawArcType
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawArcType", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawArcType")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 늘이기 (HDrawResize)
*/
func (p *HShapeObject) ShapeDrawResize(v ...*HDrawResize) *HDrawResize {
	var para HDrawResize
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawResize", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawResize")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 회전 (HDrawRotate)
*/
func (p *HShapeObject) ShapeDrawRotate(v ...*HDrawRotate) *HDrawRotate {
	var para HDrawRotate
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawRotate", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawRotate")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 점 편집 (HDrawEditDetail)
*/
func (p *HShapeObject) ShapeDrawEditDetail(v ...*HDrawEditDetail) *HDrawEditDetail {
	var para HDrawEditDetail
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawEditDetail", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawEditDetail")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 그림 자르기 (HDrawImageScissoring)
*/
func (p *HShapeObject) ShapeDrawImageScissoring(v ...*HDrawImageScissoring) *HDrawImageScissoring {
	var para HDrawImageScissoring
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawImageScissoring", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawImageScissoring")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 뒤집기(HDrawScAction)
*/
func (p *HShapeObject) ShapeDrawScAction(v ...*HDrawScAction) *HDrawScAction {
	var para HDrawScAction
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawScAction", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawScAction")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 하이퍼 링크(HDrawCtrlHyperlink)
*/
func (p *HShapeObject) ShapeDrawCtrlHyperlink(v ...*HDrawCtrlHyperlink) *HDrawCtrlHyperlink {
	var para HDrawCtrlHyperlink
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawCtrlHyperlink", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawCtrlHyperlink")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 좌표 정보(HDrawCoordInfo)
*/
func (p *HShapeObject) ShapeDrawCoordInfo(v ...*HDrawCoordInfo) *HDrawCoordInfo {
	var para HDrawCoordInfo
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawCoordInfo", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawCoordInfo")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
쉬운 개체 선택, 마우스 끌기로 만들기 정보(HDrawSoMouseOption)
*/
func (p *HShapeObject) ShapeDrawSoMouseOption(v ...*HDrawSoMouseOption) *HDrawSoMouseOption {
	var para HDrawSoMouseOption
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawSoMouseOption", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawSoMouseOption")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
수식 개체 정보(HDrawSoEquationOption)
*/
func (p *HShapeObject) ShapeDrawSoEquationOption(v ...*HDrawSoEquationOption) *HDrawSoEquationOption {
	var para HDrawSoEquationOption
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawSoEquationOption", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawSoEquationOption")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
그리기 개체 기울기(HDrawShear)
*/
func (p *HShapeObject) ShapeDrawShear(v ...*HDrawShear) *HDrawShear {
	var para HDrawShear
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawShear", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawShear")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
글맵시(HDrawTextart)
*/
func (p *HShapeObject) ShapeDrawTextart(v ...*HDrawTextart) *HDrawTextart {
	var para HDrawTextart
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeDrawTextart", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeDrawTextart")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
셀 속성 개체(HCell)
*/
func (p *HShapeObject) ShapeTableCell(v ...*HCell) *HCell {
	var para HCell
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeTableCell", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeTableCell")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
서브 리스트의 속성(HListProperties)
*/
func (p *HShapeObject) ShapeListProperties(v ...*HListProperties) *HListProperties {
	var para HListProperties
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeListProperties", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeListProperties")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
캡션 속성(HCaption)
*/
func (p *HShapeObject) ShapeCaption(v ...*HCaption) *HCaption {
	var para HCaption
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ShapeCaption", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ShapeCaption")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
페이지 경계에서 나누는 방식.
*/
func (p *HShapeObject) PageBreak(v ...uint16) int {
	return funcParaSetInt(p.variant, "PageBreak", v)
}

/*
제목 행을 반복할지 여부. (on/off)
*/
func (p *HShapeObject) RepeatHeader(v ...uint16) int {
	return funcParaSetInt(p.variant, "RepeatHeader", v)
}

/*
셀 간격(HTML의 셀 간격과 동일 의미). HWPUNIT.
*/
func (p *HShapeObject) CellSpacing(v ...uint32) int {
	return funcParaSetInt(p.variant, "CellSpacing", v)
}

/*
각 방향 기본 셀 마진. 왼쪽
*/
func (p *HShapeObject) CellMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginLeft", v)
}

/*
각 방향 기본 셀 마진. 오른쪽
*/
func (p *HShapeObject) CellMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginRight", v)
}

/*
각 방향 기본 셀 마진. 위
*/
func (p *HShapeObject) CellMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginTop", v)
}

/*
각 방향 기본 셀 마진. 아래
*/
func (p *HShapeObject) CellMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginBottom", v)
}

/*
표에 링크된 차트의 정보 (HTableChartInfo)
*/
func (p *HShapeObject) TableCharInfo(v ...*HTableChartInfo) *HTableChartInfo {
	var para HTableChartInfo
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableCharInfo", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableCharInfo")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
표 테두리/배경 (HBorderFill)
*/
func (p *HShapeObject) TableBorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableBorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableBorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
수식 스크립트
*/
func (p *HShapeObject) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
HWPUNIT인 수식의 Base 크기이다. (기본값은 POINT 10)
*/
func (p *HShapeObject) BaseUnit(v ...int32) int {
	return funcParaSetInt(p.variant, "BaseUnit", v)
}

/*
HWPUNIT인 수식의 Base 크기이다. ( 기본값은 WINDOWTEXT 색)
*/
func (p *HShapeObject) Color(v ...int32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
줄 단위를 사용할지의 여부
*/
func (p *HShapeObject) LineMode(v ...int32) int {
	return funcParaSetInt(p.variant, "LineMode", v)
}

/*
수식 스크립트 버전
*/
func (p *HShapeObject) Version(v ...string) string {
	return funcParaSetString(p.variant, "Version", v)
}

/*
HPrint:

	HPrint :  프린트 속성 개체
*/
type HPrint struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPrint) HSet() *HSet {
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
인쇄 범위
*/
func (p *HPrint) Range(v ...uint16) int {
	return funcParaSetInt(p.variant, "Range", v)
}

/*
사용자가 직접 입력한 인쇄 범위
*/
func (p *HPrint) RangeCustom(v ...string) string {
	return funcParaSetString(p.variant, "RangeCustom", v)
}

/*
인쇄 매수
*/
func (p *HPrint) NumCopy(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumCopy", v)
}

/*
한 부씩 찍기
*/
func (p *HPrint) Collate(v ...uint16) int {
	return funcParaSetInt(p.variant, "Collate", v)
}

/*
인쇄 방법
*/
func (p *HPrint) PrintMethod(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintMethod", v)
}

/*
공급용지 종류(DEVMODE.dmPaperSize)
*/
func (p *HPrint) PrinterPaperSize(v ...int32) int {
	return funcParaSetInt(p.variant, "PrinterPaperSize", v)
}

/*
공급용지 종류(DEVMODE.dmPaperWidth)
*/
func (p *HPrint) PrinterPaperWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "PrinterPaperWidth", v)
}

/*
공급용지 종류(DEVMODE.dmPaperLength)
*/
func (p *HPrint) PrinterPaperLength(v ...int32) int {
	return funcParaSetInt(p.variant, "PrinterPaperLength", v)
}

/*
꼬리말 자동 인쇄
*/
func (p *HPrint) PrintAutoFootNote(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintAutoFootNote", v)
}

/*
머리말 자동 인쇄
*/
func (p *HPrint) PrintAutoHeadNote(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintAutoHeadNote", v)
}

/*
프린터
*/
func (p *HPrint) PrinterName(v ...string) string {
	return funcParaSetString(p.variant, "PrinterName", v)
}

/*
인쇄 결과를 파일로 저장
*/
func (p *HPrint) PrintToFile(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintToFile", v)
}

/*
인쇄 결과를 저장할 파일 이름
*/
func (p *HPrint) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
역순 인쇄
*/
func (p *HPrint) ReverseOrder(v ...uint16) int {
	return funcParaSetInt(p.variant, "ReverseOrder", v)
}

/*
끊어 찍기 매수
*/
func (p *HPrint) Pause(v ...uint16) int {
	return funcParaSetInt(p.variant, "Pause", v)
}

/*
그림 개체
*/
func (p *HPrint) PrintImage(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintImage", v)
}

/*
그리기 개체
*/
func (p *HPrint) PrintDrawObj(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintDrawObj", v)
}

/*
누름틀
*/
func (p *HPrint) PrintClickHere(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintClickHere", v)
}

/*
편집 용지 표시
*/
func (p *HPrint) PrintCropMark(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintCropMark", v)
}

/*
배경 그림
*/
func (p *HPrint) IdcPrintWallPaper(v ...uint16) int {
	return funcParaSetInt(p.variant, "IdcPrintWallPaper", v)
}

/*
빈 마지막 쪽
*/
func (p *HPrint) LastBlankPage(v ...uint16) int {
	return funcParaSetInt(p.variant, "LastBlankPage", v)
}

/*
바인더 구멍
*/
func (p *HPrint) BinderHoleType(v ...uint16) int {
	return funcParaSetInt(p.variant, "BinderHoleType", v)
}

/*
홀짝 인쇄
*/
func (p *HPrint) EvenOddPageType(v ...uint16) int {
	return funcParaSetInt(p.variant, "EvenOddPageType", v)
}

/*
가로 확대
*/
func (p *HPrint) ZoomX(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomX", v)
}

/*
세로 확대
*/
func (p *HPrint) ZoomY(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZoomY", v)
}

/*
자동 머리말의 왼쪽 String
*/
func (p *HPrint) PrintAutoHeadnoteLtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoHeadnoteLtext", v)
}

/*
자동 머리말의 가운데 String
*/
func (p *HPrint) PrintAutoHeadnoteCtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoHeadnoteCtext", v)
}

/*
자동 머리말의 오른쪽 String
*/
func (p *HPrint) PrintAutoHeadnoteRtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoHeadnoteRtext", v)
}

/*
자동 꼬리말의 왼쪽 String
*/
func (p *HPrint) PrintAutoFootnoteLtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoFootnoteLtext", v)
}

/*
자동 꼬리말의 가운데 String
*/
func (p *HPrint) PrintAutoFootnoteCtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoFootnoteCtext", v)
}

/*
자동 꼬리말의 오른쪽 String
*/
func (p *HPrint) PrintAutoFootnoteRtext(v ...string) string {
	return funcParaSetString(p.variant, "PrintAutoFootnoteRtext", v)
}

/*
문제 해결을 위한 고급 선택 사항
*/
func (p *HPrint) Flags(v ...uint32) int {
	return funcParaSetInt(p.variant, "Flags", v)
}

/*
인쇄 방향(장치) 0 : 프린터 1: 팩스 2: 그림으로 저장 3: PDF 파일로 저장 4: 미리보기 PrintFormObj PIT_UI1 양식 개체
*/
func (p *HPrint) Device(v ...uint16) int {
	return funcParaSetInt(p.variant, "Device", v)
}

/*
형광펜
*/
func (p *HPrint) PrintMarkPen(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintMarkPen", v)
}

/*
메모
*/
func (p *HPrint) PrintMemo(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintMemo", v)
}

/*
메모 내용
*/
func (p *HPrint) PrintMemoContents(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrintMemoContents", v)
}

/*
HTabDef:

	HTabDef :  탭 정의 속성 개체
*/
type HTabDef struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTabDef) HSet() *HSet {
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
문단 왼쪽 끝 탭 (on/off)
*/
func (p *HTabDef) AutoTabLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoTabLeft", v)
}

/*
문단 오른쪽 끝 탭 (on/off)
*/
func (p *HTabDef) AutoTabRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoTabRight", v)
}

/*
각각의 탭 정의.

	하나의 탭 아이템은 세 개의 인수로 표현되어 있음.
	(n * 3 + 0) - PIT_I : 탭 위치 (URC)
	(n * 3 + 1) - PIT_I : 채울 모양 (아래참조)
	(n * 3 + 2) - PIT_I : 탭 종류 (아래참조)

	채울 모양 : 선 종류 탭 종류 : 0 = 왼쪽 1 = 오른쪽 2 = 가운데 3 = 소수점
*/
func (p *HTabDef) TabItem() *HArray {
	return GetHArray(p.variant, "TabItem")
}

/*
PIT_ARRAY 생성

	PIT_ARRAY를 위해 필요
*/
func (p *HTabDef) CreateItemArray(itemID string, num int32) {
	CreateItemArray(p.variant, itemID, num)
}

/*
HParaShape:

	HParaShape :  문단 모양 속성 개체
*/
type HParaShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HParaShape) HSet() *HSet {
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
왼쪽 여백 (URC)
*/
func (p *HParaShape) LeftMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "LeftMargin", v)
}

/*
오른쪽 여백 (URC)
*/
func (p *HParaShape) RightMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "RightMargin", v)
}

/*
들여쓰기/내어쓰기 (URC)
*/
func (p *HParaShape) Indentation(v ...int32) int {
	return funcParaSetInt(p.variant, "Indentation", v)
}

/*
문단 간격 위 (URC)
*/
func (p *HParaShape) PrevSpacing(v ...int32) int {
	return funcParaSetInt(p.variant, "PrevSpacing", v)
}

/*
문단 간격 아래 (URC)
*/
func (p *HParaShape) NextSpacing(v ...int32) int {
	return funcParaSetInt(p.variant, "NextSpacing", v)
}

/*
줄 간격 종류 (HWPUNIT) 0 = 글자에 따라, 1 = 고정 값, 2 = 여백만 지정
*/
func (p *HParaShape) LineSpacingType(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineSpacingType", v)
}

/*
줄 간격 값 줄 간견 종류(LineSpacingType)에 따라 : -“글자에 따라”일 경우(0 - 500%) -“고정값”일 경우(URC) -“여백만 지정”일 경우(URC)
*/
func (p *HParaShape) LineSpacing(v ...int32) int {
	return funcParaSetInt(p.variant, "LineSpacing", v)
}

/*
정렬 방식 0 = 양쪽 정렬 1 = 왼쪽 정렬 2 = 오른쪽 정렬 3 = 가운데 정렬 4 = 배분 정렬 5 = 나눔 정렬 (공백에만 배분)
*/
func (p *HParaShape) AlignType(v ...uint16) int {
	return funcParaSetInt(p.variant, "AlignType", v)
}

/*
줄 나눔 단위 (라틴 문자) 0 = 단어, 1 = 하이픈, 2 = 글자
*/
func (p *HParaShape) BreakLatinWord(v ...uint16) int {
	return funcParaSetInt(p.variant, "BreakLatinWord", v)
}

/*
단위 (비라틴 문자) TRUE = 글자, FALSE = 어절
*/
func (p *HParaShape) BreakNonLatinWord(v ...uint16) int {
	return funcParaSetInt(p.variant, "BreakNonLatinWord", v)
}

/*
편집 용지의 줄 격자 사용 (on/off)
*/
func (p *HParaShape) SnapToGrid(v ...uint16) int {
	return funcParaSetInt(p.variant, "SnapToGrid", v)
}

/*
공백 최솟값 (0 - 75%)
*/
func (p *HParaShape) Condense(v ...uint16) int {
	return funcParaSetInt(p.variant, "Condense", v)
}

/*
외톨이줄 보호 (on/off)
*/
func (p *HParaShape) WidowOrphan(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidowOrphan", v)
}

/*
다음 문단과 함께 (on/off)
*/
func (p *HParaShape) KeepWithNext(v ...uint16) int {
	return funcParaSetInt(p.variant, "KeepWithNext", v)
}

/*
문단 보호 (on/off)
*/
func (p *HParaShape) KeepLinesTogether(v ...uint16) int {
	return funcParaSetInt(p.variant, "KeepLinesTogether", v)
}

/*
문단 앞에서 항상 쪽나눔 (on/off)
*/
func (p *HParaShape) PagebreakBefore(v ...uint16) int {
	return funcParaSetInt(p.variant, "PagebreakBefore", v)
}

/*
세로 정렬 0 = 글꼴기준, 1 = 위, 2 = 가운데, 3 = 아래
*/
func (p *HParaShape) TextAlignment(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextAlignment", v)
}

/*
글꼴에 어울리는 줄높이 (on/off)
*/
func (p *HParaShape) FontLineHeight(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontLineHeight", v)
}

/*
문단 머리 모양 0 = 없음, 1 = 개요, 2 = 번호, 3 = 불릿
*/
func (p *HParaShape) HeadingType(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeadingType", v)
}

/*
단계 (0 - 6)
*/
func (p *HParaShape) Level(v ...uint16) int {
	return funcParaSetInt(p.variant, "Level", v)
}

/*
문단 테두리/배경 - 테두리 연결 (on/off)
*/
func (p *HParaShape) BorderConnect(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderConnect", v)
}

/*
문단 테두리/배경 - 여백 무시 (0 = 단, 1 = 텍스트)
*/
func (p *HParaShape) BorderText(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderText", v)
}

/*
문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 왼쪽
*/
func (p *HParaShape) BorderOffsetLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "BorderOffsetLeft", v)
}

/*
문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 오른쪽
*/
func (p *HParaShape) BorderOffsetRight(v ...int32) int {
	return funcParaSetInt(p.variant, "BorderOffsetRight", v)
}

/*
문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 위
*/
func (p *HParaShape) BorderOffsetTop(v ...int32) int {
	return funcParaSetInt(p.variant, "BorderOffsetTop", v)
}

/*
문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 아래
*/
func (p *HParaShape) BorderOffsetBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "BorderOffsetBottom", v)
}

/*
탭 정의 서브셋 (HTabDef)
*/
func (p *HParaShape) TabDef(v ...*HTabDef) *HTabDef {
	var para HTabDef
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TabDef", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TabDef")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
문단 번호 서브셋 (HNumberingShape) 문단 머리 모양 HeadingType이 ‘개요’, ‘번호’일 때 사용
*/
func (p *HParaShape) Numbering(v ...*HNumberingShape) *HNumberingShape {
	var para HNumberingShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "Numbering", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Numbering")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
불릿 모양 서브셋 (HBulletShape) 문단 머리 모양 HeadingType이 ‘불릿’ 일 때 사용
*/
func (p *HParaShape) Bullet(v ...*HBulletShape) *HBulletShape {
	var para HBulletShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "Bullet", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Bullet")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
테두리/배경 서브셋 (HBorderFill)
*/
func (p *HParaShape) BorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "BorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "BorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HFootnoteShape:

	HFootnoteShape :  구역 내 각주/미주 속성 개체
*/
type HFootnoteShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFootnoteShape) HSet() *HSet {
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
번호 모양
*/
func (p *HFootnoteShape) NumberFormat(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumberFormat", v)
}

/*
사용자 기호 (WCHAR)
*/
func (p *HFootnoteShape) UserChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "UserChar", v)
}

/*
앞 장식 문자 (WCHAR)
*/
func (p *HFootnoteShape) PrefixChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrefixChar", v)
}

/*
뒤 장식 문자 (WCHAR)
*/
func (p *HFootnoteShape) Suffix(v ...uint16) int {
	return funcParaSetInt(p.variant, "Suffix", v)
}

/*
위치 - 각주용 옵션 (한 페이지 내에서 각주를 다단에 어떻게 위치시킬지) 0 = 각 단마다 따로 배열 1 = 통단으로 배열 2 = 가장 오른쪽 단에 배열 - 미주용 옵션 (문서 내에서 미주를 어디에 위치시킬지) 0 = 문서의 마지막 1 = 구역의 마지막
*/
func (p *HFootnoteShape) PlaceAt(v ...uint16) int {
	return funcParaSetInt(p.variant, "PlaceAt", v)
}

/*
번호 매기기 0 = 앞 구역에 이어서 1 = 현재 구역부터 새로 시작 2 = 쪽마다 새로 시작 (각주 전용)
*/
func (p *HFootnoteShape) Restart(v ...uint16) int {
	return funcParaSetInt(p.variant, "Restart", v)
}

/*
시작 번호 (1 .. n) 번호 매기기 값이 ‘쪽마다 새로 시작’ 일 때만 사용된다.
*/
func (p *HFootnoteShape) NewNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "NewNumber", v)
}

/*
구분선 길이 (HWPUNIT)
*/
func (p *HFootnoteShape) LineLength(v ...int32) int {
	return funcParaSetInt(p.variant, "LineLength", v)
}

/*
선 종류
*/
func (p *HFootnoteShape) LineType(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineType", v)
}

/*
선 굵기
*/
func (p *HFootnoteShape) LineWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineWidth", v)
}

/*
구분선 위 여백 (HWPUNIT)
*/
func (p *HFootnoteShape) SpaceAboveLine(v ...int32) int {
	return funcParaSetInt(p.variant, "SpaceAboveLine", v)
}

/*
구분선 아래 여백 (HWPUNIT)
*/
func (p *HFootnoteShape) SpaceBelowLine(v ...int32) int {
	return funcParaSetInt(p.variant, "SpaceBelowLine", v)
}

/*
주석 사이 여백 (HWPUNIT)
*/
func (p *HFootnoteShape) SpaceBetweenNotes(v ...int32) int {
	return funcParaSetInt(p.variant, "SpaceBetweenNotes", v)
}

/*
각주 내용 중 번호 코드의 모양을 윗 첨자 형식으로 할지
*/
func (p *HFootnoteShape) SuperScript(v ...uint16) int {
	return funcParaSetInt(p.variant, "SuperScript", v)
}

/*
텍스트에 이어 바로 출력할지 여부 (on/off)
*/
func (p *HFootnoteShape) BeneathText(v ...uint16) int {
	return funcParaSetInt(p.variant, "BeneathText", v)
}

/*
선 색 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HFootnoteShape) LineColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "LineColor", v)
}

/*
HAutoNum:

	HAutoNum :  번호 넣기/새 번호로 속성 개체
*/
type HAutoNum struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HAutoNum) HSet() *HSet {
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
번호 종류 0 = 쪽 번호 1 = 각주 번호 2 = 미주 번호 3 = 그림 번호 4 = 표 번호 5 = 수식 번호
*/
func (p *HAutoNum) NumType(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumType", v)
}

/*
새 시작 번호 (1 .. n)
*/
func (p *HAutoNum) NewNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "NewNumber", v)
}

/*
HPageNumCtrl:

	HPageNumCtrl :  페이지 번호 제어 (97의 홀수 쪽에서 시작) 속성 개체
*/
type HPageNumCtrl struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPageNumCtrl) HSet() *HSet {
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
페이지 번호 적용 옵션 0 = 양쪽, 1 = 짝수쪽, 2 = 홀수쪽
*/
func (p *HPageNumCtrl) PageStartsOn(v ...uint16) int {
	return funcParaSetInt(p.variant, "PageStartsOn", v)
}

/*
HPageHiding:

	HPageHiding :  감추기 속성 개체
*/
type HPageHiding struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPageHiding) HSet() *HSet {
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
감출 대상 비트 필드 0x01 = 머리말 0x02 = 꼬리말 0x04 = 바탕쪽 0x08 = 테두리 0x10 = 배경 0x20 = 쪽 번호 위치
*/
func (p *HPageHiding) Fields(v ...uint32) int {
	return funcParaSetInt(p.variant, "Fields", v)
}

/*
HPageNumPos:

	HPageNumPos :  쪽번호 위치 속성 개체
*/
type HPageNumPos struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPageNumPos) HSet() *HSet {
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
번호모양
*/
func (p *HPageNumPos) NumberFormat(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumberFormat", v)
}

/*
사용자 기호 (WCHAR)
*/
func (p *HPageNumPos) UserChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "UserChar", v)
}

/*
앞 장식 문자 (WCHAR)
*/
func (p *HPageNumPos) PrefixChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "PrefixChar", v)
}

/*
뒤 장식 문자 (WCHAR)
*/
func (p *HPageNumPos) SuffixChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "SuffixChar", v)
}

/*
양쪽 옆 장식 문자 (WCHAR)
*/
func (p *HPageNumPos) SideChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "SideChar", v)
}

/*
번호 위치 0 = NONE 1 = TOP_LEFT 2 = TOP_CENTER 3 = TOP_RIGHT 4 = BOTTOM_LEFT 5 = BOTTOM_CENTER 6 = BOTTOM_RIGHT 7 = OUTSIDE_TOP 8 = OUTSIDE_BOTTOM 9 = INSIDE_TOP 10 = INSIDE_BOTTOM
*/
func (p *HPageNumPos) DrawPos(v ...uint16) int {
	return funcParaSetInt(p.variant, "DrawPos", v)
}

/*
HHeaderFooter:

	HHeaderFooter :  머리말/꼬리말 속성 개체
*/
type HHeaderFooter struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HHeaderFooter) HSet() *HSet {
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
0 = 양쪽 1 = 짝수쪽 2 = 홀수쪽
*/
func (p *HHeaderFooter) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HMailMergeGenerate:

	HMailMergeGenerate :  메일 머지 속성 개체
*/
type HMailMergeGenerate struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMailMergeGenerate) HSet() *HSet {
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
자료 종류
*/
func (p *HMailMergeGenerate) Input(v ...uint16) int {
	return funcParaSetInt(p.variant, "Input", v)
}

/*
Hwp 문서 경로. 비어 있으면 "HwpId"를 사용
*/
func (p *HMailMergeGenerate) HwpPath(v ...string) string {
	return funcParaSetString(p.variant, "HwpPath", v)
}

/*
Hwp 문서 ID
*/
func (p *HMailMergeGenerate) HwpId(v ...uint32) int {
	return funcParaSetInt(p.variant, "HwpId", v)
}

/*
dbf file path
*/
func (p *HMailMergeGenerate) DbfPath(v ...string) string {
	return funcParaSetString(p.variant, "DbfPath", v)
}

/*
dbf file codepage
*/
func (p *HMailMergeGenerate) DbfCode(v ...uint16) int {
	return funcParaSetInt(p.variant, "DbfCode", v)
}

/*
출력 방향
*/
func (p *HMailMergeGenerate) Output(v ...uint16) int {
	return funcParaSetInt(p.variant, "Output", v)
}

/*
파일 이름
*/
func (p *HMailMergeGenerate) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
쪽 번호 잇기
*/
func (p *HMailMergeGenerate) Continue(v ...uint16) int {
	return funcParaSetInt(p.variant, "Continue", v)
}

/*
인쇄 선택 사항

	미구현
*/
func (p *HMailMergeGenerate) PrintSet() {
}

/*
메일 제목
*/
func (p *HMailMergeGenerate) Subjec(v ...string) string {
	return funcParaSetString(p.variant, "Subjec", v)
}

/*
메일 종류 0 = 본문 1 = 첨부 파일 2 = 사용자 직접 메일에 첨부 파일
*/
func (p *HMailMergeGenerate) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
메일 주소 필드
*/
func (p *HMailMergeGenerate) Field(v ...string) string {
	return funcParaSetString(p.variant, "Field", v)
}

/*
필드단위 업데이트
*/
func (p *HMailMergeGenerate) FieldUpdate(v ...uint16) int {
	return funcParaSetInt(p.variant, "FieldUpdate", v)
}

/*
HListProperties:

	HListProperties :  서브 리스트 속성 개체
*/
type HListProperties struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HListProperties) HSet() *HSet {
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
글자 방향
*/
func (p *HListProperties) TextDirection(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextDirection", v)
}

/*
경계에서 줄 나눔 방식 0 = 일반적인 줄 바꿈 1 = 줄을 바꾸지 않음 2 = 자간을 조정하여 한 줄을 유지
*/
func (p *HListProperties) LineWrap(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineWrap", v)
}

/*
세로 정렬 0 = ALIGN_TOP 1 = ALIGN_CENTER 2 = ALIGN_BOTTOM
*/
func (p *HListProperties) VertAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
각 방향 마진 : 왼쪽
*/
func (p *HListProperties) MarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginLeft", v)
}

/*
각 방향 마진 : 오른쪽
*/
func (p *HListProperties) MarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginRight", v)
}

/*
각 방향 마진 : 위
*/
func (p *HListProperties) MarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginTop", v)
}

/*
각 방향 마진 : 아래
*/
func (p *HListProperties) MarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginBottom", v)
}

/*
HTable:

	HTable :  표 속성 개체
*/
type HTable struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTable) HSet() *HSet {
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
글자처럼 취급 on / off
*/
func (p *HTable) TreatAsChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "TreatAsChar", v)
}

/*
줄 간격에 영향을 줄지 여부 on / off
*/
func (p *HTable) AffectsLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "AffectsLine", v)
}

/*
세로 위치의 기준
*/
func (p *HTable) VertRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertRelTo", v)
}

/*
"VertRelTo"에 대한 상대적인 배열 방식
*/
func (p *HTable) VertAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
"VertRelTo"와 "VertAlign"을 기준점으로 한 상대적인 오프셋 값 HWPUNIT 단위
*/
func (p *HTable) VertOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "VertOffset", v)
}

/*
가로 위치의 기준
*/
func (p *HTable) HorzRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzRelTo", v)
}

/*
"HorzRelTo"에 대한 상대적인 배열 방식
*/
func (p *HTable) HorzAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzAlign", v)
}

/*
"HorzRelTo"와 "HorzAlign"을 기준점으로 한 상대적인 오프셋 값 HWPUNIT 단위
*/
func (p *HTable) HorzOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "HorzOffset", v)
}

/*
오브젝트의 세로 위치를 본문 영역으로 제한할지 여부 on / off
*/
func (p *HTable) FlowWithText(v ...uint16) int {
	return funcParaSetInt(p.variant, "FlowWithText", v)
}

/*
다른 오브젝트와 겹치는 것을 허용할지 여부 on / off
*/
func (p *HTable) AllowOverlap(v ...uint16) int {
	return funcParaSetInt(p.variant, "AllowOverlap", v)
}

/*
오브젝트 폭의 기준
*/
func (p *HTable) WidthRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthRelTo", v)
}

/*
오브젝트 폭의 값
*/
func (p *HTable) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
오브젝트 높이의 기준
*/
func (p *HTable) HeightRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeightRelTo", v)
}

/*
오브젝트의 높이 값
*/
func (p *HTable) Height(v ...int32) int {
	return funcParaSetInt(p.variant, "Height", v)
}

/*
크기 보호 on / off
*/
func (p *HTable) ProtectSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "ProtectSize", v)
}

/*
오브젝트 주위를 텍스트가 어떻게 흘러갈지 지정하는 옵션. * "TreatAsChar" = FALSE일 때만 사용됨 0 = bound rect를 따라 1 = 좌, 우에는 텍스트를 배치하지 않음 2 = 글과 겹치게 하여 글 뒤로 3 = 글과 겹치게 하여 글 앞으로 4 = 오브젝트의 outline을 따라 5 = 오브젝트 내부의 빈 공간까지
*/
func (p *HTable) TextWrap(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextWrap", v)
}

/*
오브젝트의 좌/우 어느쪽에 글을 배치할지 지정하는 옵션. * "TextWrap"가 0, 4, 5일 때만 사용된다. 0 = 양쪽 1 = 왼쪽 2 = 오른쪽 3 = 큰쪽
*/
func (p *HTable) TextFlow(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextFlow", v)
}

/*
오브젝트의 바깥 여백. (왼쪽) HWPUNIT 단위
*/
func (p *HTable) OutsideMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginLeft", v)
}

/*
오브젝트의 바깥 여백. (오른쪽) HWPUNIT 단위
*/
func (p *HTable) OutsideMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginRight", v)
}

/*
오브젝트의 바깥 여백. (위) HWPUNIT 단위
*/
func (p *HTable) OutsideMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginTop", v)
}

/*
오브젝트의 바깥 여백. (아래) HWPUNIT 단위
*/
func (p *HTable) OutsideMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginBottom", v)
}

/*
이 개체가 속하는 번호 범주 0 = 없음, 1 = 그림, 2 = 표, 3 = 수식
*/
func (p *HTable) NumberingType(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumberingType", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 폭의 값 글상자 등이 늘어나면 늘어난 값을 계산해서 가진다.
*/
func (p *HTable) LayoutWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutWidth", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 높이 값 글상자 등이 늘어나면 늘어난 값을 계산해서 가진다.
*/
func (p *HTable) LayoutHeight(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutHeight", v)
}

/*
오브젝트 선택 가능 on / off
*/
func (p *HTable) Lock(v ...uint16) int {
	return funcParaSetInt(p.variant, "Lock", v)
}

/*
쪽나눔방지II on/off
*/
func (p *HTable) HoldAnchorObj(v ...uint16) int {
	return funcParaSetInt(p.variant, "HoldAnchorObj", v)
}

/*
개체가 존재 하는 페이지
*/
func (p *HTable) PageNumber(v ...uint32) int {
	return funcParaSetInt(p.variant, "PageNumber", v)
}

/*
개체 Selection 상태 TRUE/FASLE
*/
func (p *HTable) AdjustSelection(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustSelection", v)
}

/*
글상자로 TRUE/FASLE
*/
func (p *HTable) AdjustTextbox(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustTextbox", v)
}

/*
앞개체 속성 따라가기 TRUE/FASLE
*/
func (p *HTable) AdjustPrevObjAttr(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustPrevObjAttr", v)
}

/*
페이지 경계에서 나누는 방식.
*/
func (p *HTable) PageBreak(v ...uint16) int {
	return funcParaSetInt(p.variant, "PageBreak", v)
}

/*
제목 행을 반복할지 여부. (on/off)
*/
func (p *HTable) RepeatHeader(v ...uint16) int {
	return funcParaSetInt(p.variant, "RepeatHeader", v)
}

/*
셀 간격(HTML의 셀간격과 동일 의미). HWPUNIT.
*/
func (p *HTable) CellSpacing(v ...uint32) int {
	return funcParaSetInt(p.variant, "CellSpacing", v)
}

/*
각 방향 기본 셀 마진 : 왼쪽
*/
func (p *HTable) CellMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginLeft", v)
}

/*
각 방향 기본 셀 마진 : 오른쪽
*/
func (p *HTable) CellMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginRight", v)
}

/*
각 방향 기본 셀 마진 : 위
*/
func (p *HTable) CellMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginTop", v)
}

/*
각 방향 기본 셀 마진 : 아래
*/
func (p *HTable) CellMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "CellMarginBottom", v)
}

/*
표에 적용되는 테두리/배경 (HBorderFill)
*/
func (p *HTable) BorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "BorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "BorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
표에 링크된 차트의 정보 (HTableChartInfo)
*/
func (p *HTable) TableCharInfo(v ...*HTableChartInfo) *HTableChartInfo {
	var para HTableChartInfo
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableCharInfo", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableCharInfo")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
표 테두리/배경 (HBorderFill)
*/
func (p *HTable) TableBorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableBorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableBorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HCell:

	HCell :  셀 속성 개체
*/
type HCell struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCell) HSet() *HSet {
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
글자 방향
*/
func (p *HCell) TextDirection(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextDirection", v)
}

/*
경계에서 줄 나눔 방식 0 = 일반적인 줄 바꿈 1 = 줄을 바꾸지 않음 2 = 자간을 조정하여 한 줄을 유지
*/
func (p *HCell) LineWrap(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineWrap", v)
}

/*
세로 정렬 0 = ALIGN_TOP 1 = ALIGN_CENTER 2 = ALIGN_BOTTOM
*/
func (p *HCell) VertAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
각 방향 마진 : 왼쪽
*/
func (p *HCell) MarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginLeft", v)
}

/*
각 방향 마진 : 오른쪽
*/
func (p *HCell) MarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginRight", v)
}

/*
각 방향 마진 : 위
*/
func (p *HCell) MarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginTop", v)
}

/*
각 방향 마진 : 아래
*/
func (p *HCell) MarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "MarginBottom", v)
}

/*
테이블의 기본 셀 마진 대신 자체 셀 마진을 적용할지 여부. (on/off)
*/
func (p *HCell) HasMargin(v ...uint16) int {
	return funcParaSetInt(p.variant, "HasMargin", v)
}

/*
사용자 편집을 막을지 여부. (on/off)
*/
func (p *HCell) Protected(v ...uint16) int {
	return funcParaSetInt(p.variant, "Protected", v)
}

/*
제목 셀인지 여부. (on/off)
*/
func (p *HCell) Header(v ...uint16) int {
	return funcParaSetInt(p.variant, "Header", v)
}

/*
셀의 폭 (HWPUNIT)
*/
func (p *HCell) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
셀의 높이 (HWPUNIT)
*/
func (p *HCell) Height(v ...int32) int {
	return funcParaSetInt(p.variant, "Height", v)
}

/*
양식모드에서 편집 가능 여부(on/off)
*/
func (p *HCell) Editable(v ...uint16) int {
	return funcParaSetInt(p.variant, "Editable", v)
}

/*
HCtrlData (HCtrlData)
*/
func (p *HCell) CellCtrlData(v ...*HCtrlData) *HCtrlData {
	var para HCtrlData
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "CellCtrlData", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CellCtrlData")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HTableCreation:

	HTableCreation :  표 생성 속성 개체
*/
type HTableCreation struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableCreation) HSet() *HSet {
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
행 수 (생략하면 5)
*/
func (p *HTableCreation) Rows(v ...uint16) int {
	if len(v) == 0 {
		return funcParaSetInt(p.variant, "Rows", []uint16{5})
	}
	return funcParaSetInt(p.variant, "Rows", v)
}

/*
컬럼 수 (생략하면 5)
*/
func (p *HTableCreation) Cols(v ...uint16) int {
	if len(v) == 0 {
		return funcParaSetInt(p.variant, "Cols", []uint16{5})
	}
	return funcParaSetInt(p.variant, "Cols", v)
}

/*
행의 디폴트 높이 (PIT_I4)

	PIT_ARRAY (PIT_I4)
	각 index가 행의 높이를 HWPUNIT 단위로 나타낸다.
	행의 개수보다 배열 원소의 수가 부족할 때는 이후 행은 마지막 행의 폭을 따라간다.
*/
func (p *HTableCreation) RowHeight() *HArray {
	return GetHArray(p.variant, "RowHeight")
}

/*
컬럼의 디폴트 폭

	PIT_ARRAY (PIT_I4)
*/
func (p *HTableCreation) ColWidth() *HArray {
	return GetHArray(p.variant, "ColWidth")
}

/*
정보가 없는 셀은 디폴트값을 따라가므로 모든 셀에 대해 정보를 줄 필요는 없다.
*/
func (p *HTableCreation) CellInfo() *HArray {
	return GetHArray(p.variant, "CellInfo")
}

/*
PIT_ARRAY 생성

	PIT_ARRAY를 위해 필요
	※ 내부적으로 사용될 뿐 직접 사용할 일은 없음.
*/
func (p *HTableCreation) CreateItemArray(itemID string, num int32) {
	CreateItemArray(p.variant, itemID, num)
}

/*
너비

	0:단에맞춤
	1:문단에맞춤
	2:임의값
*/
func (p *HTableCreation) WidthType(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthType", v)
}

/*
높이

	0:자동
	1:임의값
*/
func (p *HTableCreation) HeightType(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeightType", v)
}

/*
너비 값
*/
func (p *HTableCreation) WidthValue(v ...int32) int {
	return funcParaSetInt(p.variant, "WidthValue", v)
}

/*
높이 값
*/
func (p *HTableCreation) HeightValue(v ...int32) int {
	return funcParaSetInt(p.variant, "HeightValue", v)
}

/*
표 마당 적용 여부
*/
func (p *HTableCreation) TableTemplateValue(v ...uint16) int {
	return funcParaSetInt(p.variant, "TableTemplateValue", v)
}

/*
초기 표 속성 (HTable)
*/
func (p *HTableCreation) TableProperties(v ...*HTable) *HTable {
	var para HTable
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableProperties", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableProperties")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
표 마당 적용 속성
*/
func (p *HTableCreation) TableTemplate(v ...*HTableTemplate) *HTableTemplate {
	var para HTableTemplate
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableTemplate", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableTemplate")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
마우스로 선을 그릴 때 속성

	미구현
	HTableDrawpen인가 싶긴 하지만 에러남.
*/
func (p *HTableCreation) TableDrawProperties() {
}

/*
HSummaryInfo:

	HSummaryInfo :  문서 정보 속성 개체
*/
type HSummaryInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSummaryInfo) HSet() *HSet {
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
제목
*/
func (p *HSummaryInfo) GetTitle(v ...string) string {
	return funcParaSetString(p.variant, "GetTitle", v)
}

/*
주제
*/
func (p *HSummaryInfo) Subject(v ...string) string {
	return funcParaSetString(p.variant, "Subject", v)
}

/*
지은이
*/
func (p *HSummaryInfo) Author(v ...string) string {
	return funcParaSetString(p.variant, "Author", v)
}

/*
날짜
*/
func (p *HSummaryInfo) Date(v ...string) string {
	return funcParaSetString(p.variant, "Date", v)
}

/*
키워드
*/
func (p *HSummaryInfo) Keywords(v ...string) string {
	return funcParaSetString(p.variant, "Keywords", v)
}

/*
기타
*/
func (p *HSummaryInfo) Comments(v ...string) string {
	return funcParaSetString(p.variant, "Comments", v)
}

/*
작성한 날짜 (low)
*/
func (p *HSummaryInfo) CreationTimeLow(v ...uint32) int {
	return funcParaSetInt(p.variant, "CreationTimeLow", v)
}

/*
작성한 날짜 (high)
*/
func (p *HSummaryInfo) CreationTimeHigh(v ...uint32) int {
	return funcParaSetInt(p.variant, "CreationTimeHigh", v)
}

/*
마지막 수정한 날짜 (low)
*/
func (p *HSummaryInfo) ModifiedTimeLow(v ...uint32) int {
	return funcParaSetInt(p.variant, "ModifiedTimeLow", v)
}

/*
마지막 수정한 날짜 (high)
*/
func (p *HSummaryInfo) ModifiedTimeHigh(v ...uint32) int {
	return funcParaSetInt(p.variant, "ModifiedTimeHigh", v)
}

/*
마지막 인쇄한 날짜 (low)
*/
func (p *HSummaryInfo) PrintedTimeLow(v ...uint32) int {
	return funcParaSetInt(p.variant, "PrintedTimeLow", v)
}

/*
마지막 인쇄한 날짜 (high)
*/
func (p *HSummaryInfo) PrintedTimeHigh(v ...uint32) int {
	return funcParaSetInt(p.variant, "PrintedTimeHigh", v)
}

/*
마지막 저장한 사람
*/
func (p *HSummaryInfo) LastSavedBy(v ...string) string {
	return funcParaSetString(p.variant, "LastSavedBy", v)
}

/*
문서분량 (글자)
*/
func (p *HSummaryInfo) Characters(v ...int32) int {
	return funcParaSetInt(p.variant, "Characters", v)
}

/*
문서분량 (낱말)
*/
func (p *HSummaryInfo) Words(v ...int32) int {
	return funcParaSetInt(p.variant, "Words", v)
}

/*
문서분량 (줄)
*/
func (p *HSummaryInfo) Lines(v ...int32) int {
	return funcParaSetInt(p.variant, "Lines", v)
}

/*
문서분량 (문단)
*/
func (p *HSummaryInfo) Paragraphs(v ...int32) int {
	return funcParaSetInt(p.variant, "Paragraphs", v)
}

/*
문서분량 (쪽)
*/
func (p *HSummaryInfo) Pages(v ...int32) int {
	return funcParaSetInt(p.variant, "Pages", v)
}

/*
문서분량 (원고지)
*/
func (p *HSummaryInfo) CopyPapers(v ...int32) int {
	return funcParaSetInt(p.variant, "CopyPapers", v)
}

/*
문서분량 (표, 그림 등)
*/
func (p *HSummaryInfo) Etcetera(v ...int32) int {
	return funcParaSetInt(p.variant, "Etcetera", v)
}

/*
문서 파일 버전
*/
func (p *HSummaryInfo) DocVersion(v ...string) string {
	return funcParaSetString(p.variant, "DocVersion", v)
}

/*
문서가 만들어진 Hwp 버전
*/
func (p *HSummaryInfo) HwpVersion(v ...string) string {
	return funcParaSetString(p.variant, "HwpVersion", v)
}

/*
문서분량 (한자 수)
*/
func (p *HSummaryInfo) HanjaChar(v ...int16) int {
	return funcParaSetInt(p.variant, "HanjaChar", v)
}

/*
HStyleItem:

	HStyleItem :  스타일 속성 개체
*/
type HStyleItem struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HStyleItem) HSet() *HSet {
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
스타일 종류 0 = 문단 스타일 1 = 글자 스타일
*/
func (p *HStyleItem) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
스타일 이름 (로컬과 영문)
*/
func (p *HStyleItem) NameLocal(v ...string) string {
	return funcParaSetString(p.variant, "NameLocal", v)
}

/*
스타일 이름 (로컬과 영문)
*/
func (p *HStyleItem) NameEng(v ...string) string {
	return funcParaSetString(p.variant, "NameEng", v)
}

/*
다음 스타일 인덱스
*/
func (p *HStyleItem) Next(v ...int32) int {
	return funcParaSetInt(p.variant, "Next", v)
}

/*
양식 스타일 보호
*/
func (p *HStyleItem) LockForm(v ...uint16) int {
	return funcParaSetInt(p.variant, "LockForm", v)
}

/*
글자 모양 (HCharShape)
*/
func (p *HStyleItem) CharShape(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "CharShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CharShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
문단 모양 (HParaShape) 스타일 종류가 ‘글자 스타일’일 때는 사용되지 않음
*/
func (p *HStyleItem) ParaShape(v ...*HParaShape) *HParaShape {
	var para HParaShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ParaShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ParaShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HGridInfo:

	HGridInfo :  줄 간격 속성 개체
*/
type HGridInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HGridInfo) HSet() *HSet {
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
격자 방식

	0: 격자에 상관없이
	1: 격자 자석 효과
	2: 격자에만 붙이기
*/
func (p *HGridInfo) Method(v ...uint16) int {
	return funcParaSetInt(p.variant, "Method", v)
}

/*
격자 기준(쪽/종이)
*/
func (p *HGridInfo) Align(v ...uint16) int {
	return funcParaSetInt(p.variant, "Align", v)
}

/*
격자 기준 가로 offset
*/
func (p *HGridInfo) HorzAlign(v ...int32) int {
	return funcParaSetInt(p.variant, "HorzAlign", v)
}

/*
격자 기준 세로 offset
*/
func (p *HGridInfo) VertAlign(v ...int32) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
격자 모양

	0: 점
	1: 선
*/
func (p *HGridInfo) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
가로 간격
*/
func (p *HGridInfo) HorzSpan(v ...uint32) int {
	return funcParaSetInt(p.variant, "HorzSpan", v)
}

/*
세로 간격
*/
func (p *HGridInfo) VertSpan(v ...uint32) int {
	return funcParaSetInt(p.variant, "VertSpan", v)
}

/*
가로 자석 범위
*/
func (p *HGridInfo) HorzRange(v ...uint32) int {
	return funcParaSetInt(p.variant, "HorzRange", v)
}

/*
세로 자석 범위
*/
func (p *HGridInfo) VertRange(v ...uint32) int {
	return funcParaSetInt(p.variant, "VertRange", v)
}

/*
격자 보이기
*/
func (p *HGridInfo) Show(v ...uint16) int {
	return funcParaSetInt(p.variant, "Show", v)
}

/*
격자 위치(글 위/글 아래)
*/
func (p *HGridInfo) ZOrder(v ...uint16) int {
	return funcParaSetInt(p.variant, "ZOrder", v)
}

/*
선 격자 보이기 종류

	0: 가로/세로
	1: 가로선
	2: 세로선
*/
func (p *HGridInfo) ViewLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "ViewLine", v)
}

/*
HNumberingShape:

	HNumberingShape :  문단 번호 모양 속성 개체
*/
type HNumberingShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HNumberingShape) HSet() *HSet {
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
7개 수준별 자체적인 글자 모양을 사용할지 여부 (0 = 스타일을 따라감, 1 = 자체 모양을 가짐)
*/
func (p *HNumberingShape) HasCharShapeLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "HasCharShapeLevel0", v)
}

/*
7개 수준별 NS_HAS_CHAR_SHAPE이 TRUE일 때만 사용되는 글자 모양 정의 (HCharShape)
*/
func (p *HNumberingShape) CharShapeLevel0(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "CharShapeLevel0", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CharShapeLevel0")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
7개 수준별 번호 너비 보정 값 (HWPUNIT)
*/
func (p *HNumberingShape) WidthAdjustLevel0(v ...int32) int {
	return funcParaSetInt(p.variant, "WidthAdjustLevel0", v)
}

/*
7개 수준별 본문과의 거리 (percent or HWPUNIT)
*/
func (p *HNumberingShape) TextOffsetLevel0(v ...int32) int {
	return funcParaSetInt(p.variant, "TextOffsetLevel0", v)
}

/*
7개 수준별 번호 정렬
*/
func (p *HNumberingShape) AlignmentLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "AlignmentLevel0", v)
}

/*
7개 수준별 번호 너비를 실제 인스턴스 문자열의 너비에 따를지 여부
*/
func (p *HNumberingShape) UseInstWidthLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "UseInstWidthLevel0", v)
}

/*
7개 번호 너비 자동 들여쓰기 여부
*/
func (p *HNumberingShape) AutoIndentLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoIndentLevel0", v)
}

/*
7개 수준별 본문과의 거리 종류 (0 = 퍼센트, 1 = HWPUNIT)
*/
func (p *HNumberingShape) TextOffsetTypeLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextOffsetTypeLevel0", v)
}

/*
7개 수준별 포맷 문자열
*/
func (p *HNumberingShape) StrFormatLevel0(v ...string) string {
	return funcParaSetString(p.variant, "StrFormatLevel0", v)
}

/*
7개 수준별 번호 포맷
*/
func (p *HNumberingShape) NumFormatLevel0(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumFormatLevel0", v)
}

/*
시작 번호 (0 = 앞 구역에 이어, n = 지정한 번호로 시작) * 0은 구역의 개요 정의에만 사용된다.
*/
func (p *HNumberingShape) StartNumber(v ...uint16) int {
	return funcParaSetInt(p.variant, "StartNumber", v)
}

/*
새로운 번호 목록을 시작할지 여부
*/
func (p *HNumberingShape) NewList(v ...uint16) int {
	return funcParaSetInt(p.variant, "NewList", v)
}

/*
HBulletShape:

	HBulletShape :  불릿 모양 속성 개체
*/
type HBulletShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HBulletShape) HSet() *HSet {
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
자체적인 글자 모양을 사용할지 여부 0 = 스타일을 따라감, 1 = 자체 모양을 가짐
*/
func (p *HBulletShape) HasCharShape(v ...uint16) int {
	return funcParaSetInt(p.variant, "HasCharShape", v)
}

/*
NS_HAS_CHAR_SHAPE이 TRUE일 때만 사용되는 글자 모양 정의 (HCharShape)
*/
func (p *HBulletShape) CharShape(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "CharShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CharShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
번호 너비 보정값 (HWPUNIT)
*/
func (p *HBulletShape) WidthAdjust(v ...int32) int {
	return funcParaSetInt(p.variant, "WidthAdjust", v)
}

/*
본문과의 거리 (percent or HWPUNIT)
*/
func (p *HBulletShape) TextOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "TextOffset", v)
}

/*
번호 정렬
*/
func (p *HBulletShape) Alignment(v ...uint16) int {
	return funcParaSetInt(p.variant, "Alignment", v)
}

/*
번호 너비를 실제 인스턴스 문자열의 너비에 따를지 여부
*/
func (p *HBulletShape) UseInstWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "UseInstWidth", v)
}

/*
번호 너비 자동 들여쓰기 여부
*/
func (p *HBulletShape) AutoIndent(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoIndent", v)
}

/*
본문과의 거리 종류 (0 = 퍼센트, 1 = HWPUNIT)
*/
func (p *HBulletShape) TextOffsetType(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextOffsetType", v)
}

/*
불릿 문자 코드
*/
func (p *HBulletShape) BulletChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "BulletChar", v)
}

/*
HIndexMark:

	HIndexMark :  찾아보기 표식 속성 개체
*/
type HIndexMark struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HIndexMark) HSet() *HSet {
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
첫 번째 키
*/
func (p *HIndexMark) First(v ...string) string {
	return funcParaSetString(p.variant, "First", v)
}

/*
두 번째 키
*/
func (p *HIndexMark) Second(v ...string) string {
	return funcParaSetString(p.variant, "Second", v)
}

/*
HChCompose:

	HChCompose :  글자 겹침 속성 개체
*/
type HChCompose struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HChCompose) HSet() *HSet {
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
겹쳐질 글자들
*/
func (p *HChCompose) Chars(v ...string) string {
	return funcParaSetString(p.variant, "Chars", v)
}

/*
HTableSplitCell:

	HTableSplitCell :  셀 나누기 속성 개체
*/
type HTableSplitCell struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableSplitCell) HSet() *HSet {
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
칸 수
*/
func (p *HTableSplitCell) Cols(v ...uint16) int {
	return funcParaSetInt(p.variant, "Cols", v)
}

/*
줄 수
*/
func (p *HTableSplitCell) Rows(v ...uint16) int {
	return funcParaSetInt(p.variant, "Rows", v)
}

/*
줄 높이를 같게
*/
func (p *HTableSplitCell) DistributeHeight(v ...uint16) int {
	return funcParaSetInt(p.variant, "DistributeHeight", v)
}

/*
나누기 전에 합치기
*/
func (p *HTableSplitCell) Merge(v ...uint16) int {
	return funcParaSetInt(p.variant, "Merge", v)
}

/*
HBorderFill:

	HBorderFill :  테두리/배경의 일반 속성 개체
*/
type HBorderFill struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HBorderFill) HSet() *HSet {
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
4방향 테두리 종류 왼쪽
*/
func (p *HBorderFill) BorderTypeLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeLeft", v)
}

/*
4방향 테두리 종류 오른쪽
*/
func (p *HBorderFill) BorderTypeRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeRight", v)
}

/*
4방향 테두리 종류 위
*/
func (p *HBorderFill) BorderTypeTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeTop", v)
}

/*
4방향 테두리 종류 아래
*/
func (p *HBorderFill) BorderTypeBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeBottom", v)
}

/*
4방향 테두리 두께 왼쪽
*/
func (p *HBorderFill) BorderWidthLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthLeft", v)
}

/*
4방향 테두리 두께 오른쪽
*/
func (p *HBorderFill) BorderWidthRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthRight", v)
}

/*
4방향 테두리 두께 위
*/
func (p *HBorderFill) BorderWidthTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthTop", v)
}

/*
4방향 테두리 두께 아래
*/
func (p *HBorderFill) BorderWidthBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthBottom", v)
}

/*
4방향 테두리 색깔 왼쪽
*/
func (p *HBorderFill) BorderCorlorLeft(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderCorlorLeft", v)
}

/*
4방향 테두리 색깔 오른쪽
*/
func (p *HBorderFill) BorderColorRight(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorRight", v)
}

/*
4방향 테두리 색깔 위
*/
func (p *HBorderFill) BorderColorTop(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorTop", v)
}

/*
4방향 테두리 색깔 아래
*/
func (p *HBorderFill) BorderColorBottom(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorBottom", v)
}

/*
슬래쉬 대각선 플랙 (bit 0-2가 시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HBorderFill) SlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "SlashFlag", v)
}

/*
백슬래쉬 대각선 플랙 (bit 0-2가 반시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HBorderFill) BackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "BackSlashFlag", v)
}

/*
대각선 종류
*/
func (p *HBorderFill) DiagonalType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiagonalType", v)
}

/*
대각선 두께
*/
func (p *HBorderFill) DiagonalWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiagonalWidth", v)
}

/*
대각선 색깔
*/
func (p *HBorderFill) DiagonalColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "DiagonalColor", v)
}

/*
3차원 효과 on/off
*/
func (p *HBorderFill) BorderFill3D(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderFill3D", v)
}

/*
그림자 효과 on/off
*/
func (p *HBorderFill) Shadow(v ...uint16) int {
	return funcParaSetInt(p.variant, "Shadow", v)
}

/*
배경 채우기 속성 (HDrawFillAttr)
*/
func (p *HBorderFill) FillAttr(v ...*HDrawFillAttr) *HDrawFillAttr {
	var para HDrawFillAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FillAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FillAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
꺽어진 대각선 플랙 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺾어진 대각선임을 나타낸다.)
*/
func (p *HBorderFill) CrookedSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CrookedSlashFlag", v)
}

/*
자동으로 나뉜 표의 경계선 설정 TRUE/FALSE
*/
func (p *HBorderFill) BreakCellSeparateLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "BreakCellSeparateLine", v)
}

/*
음영 비율 증가/감소 음영 비율 증가/감소 액션에서 전달 됨. 구현용으로만 쓰임. 이 아이템이 없으면 음영 비율 증가/감소는 없는 것이고 있다면 값이 TRUE이면 증가, FALSE이면 감소이다.
*/
func (p *HBorderFill) ShadeFaceColorIncDec(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShadeFaceColorIncDec", v)
}

/*
슬래쉬 대각선의 역방향 플랙(우상향 대각선)
*/
func (p *HBorderFill) CounterSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterSlashFlag", v)
}

/*
역슬래쉬 대각선의 역방향 플랙 (좌상향 대각선)
*/
func (p *HBorderFill) CounterBackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterBackSlashFlag", v)
}

/*
HBorderFillExt:

	HBorderFillExt :  테두리/배경의 일반 속성 확장 개체 (UI 관련)
*/
type HBorderFillExt struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HBorderFillExt) HSet() *HSet {
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
중앙선 종류 가로
*/
func (p *HBorderFillExt) TypeHorz(v ...uint16) int {
	return funcParaSetInt(p.variant, "TypeHorz", v)
}

/*
중앙선 종류 세로
*/
func (p *HBorderFillExt) TypeVert(v ...uint16) int {
	return funcParaSetInt(p.variant, "TypeVert", v)
}

/*
중앙선 두께 세로
*/
func (p *HBorderFillExt) WidthHorz(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthHorz", v)
}

/*
중앙선 두께 가로
*/
func (p *HBorderFillExt) WidthVert(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthVert", v)
}

/*
중앙선 색깔 가로 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HBorderFillExt) ColorHorz(v ...uint32) int {
	return funcParaSetInt(p.variant, "ColorHorz", v)
}

/*
중앙선 색깔 세로 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HBorderFillExt) ColorVert(v ...uint32) int {
	return funcParaSetInt(p.variant, "ColorVert", v)
}

/*
HPageBorderFill:

	HPageBorderFill :  구역의 쪽 테두리/배경 개체
*/
type HPageBorderFill struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPageBorderFill) HSet() *HSet {
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
4방향 테두리 종류 : 왼쪽 [선 종류]
*/
func (p *HPageBorderFill) BorderTypeLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeLeft", v)
}

/*
4방향 테두리 종류 : 오른쪽 [선 종류]
*/
func (p *HPageBorderFill) BorderTypeRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeRight", v)
}

/*
4방향 테두리 종류 : 위 [선 종류]
*/
func (p *HPageBorderFill) BorderTypeTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeTop", v)
}

/*
4방향 테두리 종류 : 아래 [선 종류]
*/
func (p *HPageBorderFill) BorderTypeBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeBottom", v)
}

/*
4방향 테두리 두께 : 왼쪽
*/
func (p *HPageBorderFill) BorderWidthLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthLeft", v)
}

/*
4방향 테두리 두께 : 오른쪽
*/
func (p *HPageBorderFill) BorderWidthRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthRight", v)
}

/*
4방향 테두리 두께 : 위
*/
func (p *HPageBorderFill) BorderWidthTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthTop", v)
}

/*
4방향 테두리 두께 : 아래
*/
func (p *HPageBorderFill) BorderWidthBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthBottom", v)
}

/*
4방향 테두리 색깔 : 왼쪽 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HPageBorderFill) BorderCorlorLeft(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderCorlorLeft", v)
}

/*
4방향 테두리 색깔 : 오른쪽 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HPageBorderFill) BorderColorRight(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorRight", v)
}

/*
4방향 테두리 색깔 : 위 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HPageBorderFill) BorderColorTop(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorTop", v)
}

/*
4방향 테두리 색깔 : 아래 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HPageBorderFill) BorderColorBottom(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorBottom", v)
}

/*
슬래쉬 대각선 플랙 (bit 0-2가 시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HPageBorderFill) SlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "SlashFlag", v)
}

/*
백슬래쉬 대각선 플랙 (bit 0-2가 반시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HPageBorderFill) BackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "BackSlashFlag", v)
}

/*
대각선 종류
*/
func (p *HPageBorderFill) DiagonalType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiagonalType", v)
}

/*
대각선 두께 [선 굵기]
*/
func (p *HPageBorderFill) DiagonalWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "tDiagonalWidth", v)
}

/*
대각선 색깔 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HPageBorderFill) DiagonalColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "DiagonalColor", v)
}

/*
3차원 효과 on/off
*/
func (p *HPageBorderFill) BorderFill3D(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderFill3D", v)
}

/*
그림자 효과 on/off
*/
func (p *HPageBorderFill) Shadow(v ...uint16) int {
	return funcParaSetInt(p.variant, "Shadow", v)
}

/*
배경 채우기 속성 (HDrawFillAttr)
*/
func (p *HPageBorderFill) FillAttr(v ...*HDrawFillAttr) *HDrawFillAttr {
	var para HDrawFillAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FillAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FillAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
꺾어진 대각선 플랙 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺾어진 대각선임을 나타낸다.)
*/
func (p *HPageBorderFill) CrookedSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CrookedSlashFlag", v)
}

/*
자동으로 나뉜 표의 경계선 설정 TRUE/FALSE
*/
func (p *HPageBorderFill) BreakCellSeparateLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "BreakCellSeparateLine", v)
}

/*
음영 비율 증가/감소음영 비율 증가/감소 액션에서 전달 됨. 구현용으로만 쓰임.
*/
func (p *HPageBorderFill) ShadeFaceColorIncDec(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShadeFaceColorIncDec", v)
}

/*
슬래쉬 대각선의 역방향 플랙(우상향 대각선)
*/
func (p *HPageBorderFill) CounterSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterSlashFlag", v)
}

/*
역슬래쉬 대각선의 역방향 플랙(좌상향 대각선)
*/
func (p *HPageBorderFill) CounterBackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterBackSlashFlag", v)
}

/*
TRUE = 본문 기준, FALSE = 종이 기준
*/
func (p *HPageBorderFill) TextBorder(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextBorder", v)
}

/*
머리말 포함 on/off
*/
func (p *HPageBorderFill) HeaderInside(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeaderInside", v)
}

/*
꼬리말 포함 on/off
*/
func (p *HPageBorderFill) FooterInside(v ...uint16) int {
	return funcParaSetInt(p.variant, "FooterInside", v)
}

/*
채울 영역 0 = 종이, 1 = 쪽, 2 = 테두리
*/
func (p *HPageBorderFill) FillArea(v ...uint16) int {
	return funcParaSetInt(p.variant, "FillArea", v)
}

/*
4방향 간격 왼쪽
*/
func (p *HPageBorderFill) OffsetLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "OffsetLeft", v)
}

/*
4방향 간격 오른쪽
*/
func (p *HPageBorderFill) OffsetRight(v ...int32) int {
	return funcParaSetInt(p.variant, "OffsetRight", v)
}

/*
4방향 간격 위
*/
func (p *HPageBorderFill) OffsetTop(v ...int32) int {
	return funcParaSetInt(p.variant, "OffsetTop", v)
}

/*
4방향 간격 아래
*/
func (p *HPageBorderFill) OffsetBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "OffsetBottom", v)
}

/*
HPassword:

	HPassword :  문서 암호 개체
*/
type HPassword struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPassword) HSet() *HSet {
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
암호 문자열
*/
func (p *HPassword) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
TRUE = 유니코드 모든 문자를 사용, FALSE = 영문자만 사용
*/
func (p *HPassword) FullRange(v ...uint16) int {
	return funcParaSetInt(p.variant, "FullRange", v)
}

/*
TRUE = 문서 암호를 확인, FALSE = 문서 암호를 설정
*/
func (p *HPassword) Ask(v ...uint16) int {
	return funcParaSetInt(p.variant, "Ask", v)
}

/*
HFieldCtrl:

	HFieldCtrl :  필드 컨트롤의 공통 데이터 개체
*/
type HFieldCtrl struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFieldCtrl) HSet() *HSet {
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
필드의 명령 문자열
*/
func (p *HFieldCtrl) Command(v ...string) string {
	return funcParaSetString(p.variant, "Command", v)
}

/*
일부분 readonly mode에서 편집 가능한 필드인지 여부
*/
func (p *HFieldCtrl) Editable(v ...uint16) int {
	return funcParaSetInt(p.variant, "Editable", v)
}

/*
필드가 초기화 상태인지 수정되어 있는 상태인지 여부
*/
func (p *HFieldCtrl) FieldDirty(v ...uint16) int {
	return funcParaSetInt(p.variant, "FieldDirty", v)
}

/*
HDutmal:

	HDutmal :  덧말 개체
*/
type HDutmal struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDutmal) HSet() *HSet {
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
본말
*/
func (p *HDutmal) ResultText(v ...string) string {
	return funcParaSetString(p.variant, "ResultText", v)
}

/*
덧말
*/
func (p *HDutmal) SubText(v ...string) string {
	return funcParaSetString(p.variant, "SubText", v)
}

/*
덧말 위치 { SUBT_BOTTOM, SUBT_TOP }
*/
func (p *HDutmal) PosType(v ...uint16) int {
	return funcParaSetInt(p.variant, "PosType", v)
}

/*
덧말 크기 Percent { 0=이면 기본 50% }
*/
func (p *HDutmal) FsizeRatio(v ...uint16) int {
	return funcParaSetInt(p.variant, "FsizeRatio", v)
}

/*
옵션
*/
func (p *HDutmal) Option(v ...uint16) int {
	return funcParaSetInt(p.variant, "Option", v)
}

/*
스타일 번호 (옵션이 켜있을 때)
*/
func (p *HDutmal) StyleNo(v ...uint16) int {
	return funcParaSetInt(p.variant, "StyleNo", v)
}

/*
덧말의 좌우 Align 0 = 양쪽 정렬 1 = 왼쪽 정렬 2 = 오른쪽 정렬 3 = 가운데 정렬 4 = 배분 정렬 5 = 나눔 정렬 기본은 가운데 정렬이며 옵션이 켜있을 때만 유효
*/
func (p *HDutmal) Align(v ...uint16) int {
	return funcParaSetInt(p.variant, "Align", v)
}

/*
덧말 지움
*/
func (p *HDutmal) Delete(v ...uint16) int {
	return funcParaSetInt(p.variant, "Delete", v)
}

/*
덧말 편집 모드
*/
func (p *HDutmal) Modify(v ...uint16) int {
	return funcParaSetInt(p.variant, "Modify", v)
}

/*
HCellBorderFill:

	HCellBorderFill :  셀 테두리/배경 개체
*/
type HCellBorderFill struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCellBorderFill) HSet() *HSet {
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
4방향 테두리 종류 왼쪽
*/
func (p *HCellBorderFill) BorderTypeLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeLeft", v)
}

/*
4방향 테두리 종류 오른쪽
*/
func (p *HCellBorderFill) BorderTypeRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeRight", v)
}

/*
4방향 테두리 종류 위
*/
func (p *HCellBorderFill) BorderTypeTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeTop", v)
}

/*
4방향 테두리 종류 아래
*/
func (p *HCellBorderFill) BorderTypeBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderTypeBottom", v)
}

/*
테두리 종류 내부 수직
*/
func (p *HCellBorderFill) TypeVert(v ...uint16) int {
	return funcParaSetInt(p.variant, "TypeVert", v)
}

/*
테두리 종류 내부 수평
*/
func (p *HCellBorderFill) TypeHorz(v ...uint16) int {
	return funcParaSetInt(p.variant, "TypeHorz", v)
}

/*
4방향 테두리 두께 왼쪽
*/
func (p *HCellBorderFill) BorderWidthLeft(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthLeft", v)
}

/*
4방향 테두리 두께 오른쪽
*/
func (p *HCellBorderFill) BorderWidthRight(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthRight", v)
}

/*
4방향 테두리 두께 위
*/
func (p *HCellBorderFill) BorderWidthTop(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthTop", v)
}

/*
4방향 테두리 두께 아래
*/
func (p *HCellBorderFill) BorderWidthBottom(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderWidthBottom", v)
}

/*
테두리 두께 내부 수직
*/
func (p *HCellBorderFill) WidthVert(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthVert", v)
}

/*
테두리 두께 내부 수평
*/
func (p *HCellBorderFill) WidthHorz(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthHorz", v)
}

/*
4방향 테두리 색깔 : 왼쪽 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) BorderCorlorLeft(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderCorlorLeft", v)
}

/*
4방향 테두리 색깔 : 오른쪽 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) BorderColorRight(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorRight", v)
}

/*
4방향 테두리 색깔 : 위 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) BorderColorTop(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorTop", v)
}

/*
4방향 테두리 색깔 : 아래 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) BorderColorBottom(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderColorBottom", v)
}

/*
테두리 색깔 : 내부 수직 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) ColorVert(v ...uint32) int {
	return funcParaSetInt(p.variant, "ColorVert", v)
}

/*
테두리 색깔 : 내부 수평 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) ColorHorz(v ...uint32) int {
	return funcParaSetInt(p.variant, "ColorHorz", v)
}

/*
슬래쉬 대각선 플랙 (bit 0-2가 시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HCellBorderFill) SlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "SlashFlag", v)
}

/*
백슬래쉬 대각선 플랙 (bit 0-2가 반시계 방향으로 각각의 대각선 유무를 나타냄)
*/
func (p *HCellBorderFill) BackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "BackSlashFlag", v)
}

/*
대각선 종류 [선 종류]
*/
func (p *HCellBorderFill) DiagonalType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiagonalType", v)
}

/*
대각선 두께
*/
func (p *HCellBorderFill) DiagonalWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "DiagonalWidth", v)
}

/*
대각선 색깔 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HCellBorderFill) DiagonalColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "DiagonalColor", v)
}

/*
3차원 효과 on/off
*/
func (p *HCellBorderFill) BorderFill3D(v ...uint16) int {
	return funcParaSetInt(p.variant, "BorderFill3D", v)
}

/*
그림자 효과 on/off
*/
func (p *HCellBorderFill) Shadow(v ...uint16) int {
	return funcParaSetInt(p.variant, "Shadow", v)
}

/*
배경 채우기 속성 (HDrawFillAttr)
*/
func (p *HCellBorderFill) FillAttr(v ...*HDrawFillAttr) *HDrawFillAttr {
	var para HDrawFillAttr
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FillAttr", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FillAttr")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
꺽어진 대각선 플랙 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺾어진 대각선임을 나타낸다.)
*/
func (p *HCellBorderFill) CrookedSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CrookedSlashFlag", v)
}

/*
자동으로 나뉜 표의 경계선 설정 TRUE/FALSE
*/
func (p *HCellBorderFill) BreakCellSeparateLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "BreakCellSeparateLine", v)
}

/*
음영 비율 증가/감소 음영 비율 증가/감소 액션에서 전달 됨. 구현용으로만 쓰임. 이 아이템이 없으면(음영 비율 증가/감소는 없는 것이고 있다면 값이 TRUE이면 증가, FALSE이면 감소이다.)
*/
func (p *HCellBorderFill) ShadeFaceColorIncDec(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShadeFaceColorIncDec", v)
}

/*
슬래쉬 대각선의 역방향 플랙(우상향 대각선)
*/
func (p *HCellBorderFill) CounterSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterSlashFlag", v)
}

/*
역슬래쉬 대각선의 역방향 플랙(좌상향 대각선)
*/
func (p *HCellBorderFill) CounterBackSlashFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CounterBackSlashFlag", v)
}

/*
표 테두리/배경 (HBorderFill)
*/
func (p *HCellBorderFill) TableBorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "TableBorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "TableBorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
적용 대상 0 : 각 셀 1, 3: 전체 셀 2 : 선택된 셀
*/
func (p *HCellBorderFill) ApplyTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "ApplyTo", v)
}

/*
주변 셀에 선 모양을 적용하지 않을지(존재하고 값이 1일 때)
*/
func (p *HCellBorderFill) NoNeighborCell(v ...uint16) int {
	return funcParaSetInt(p.variant, "NoNeighborCell", v)
}

/*
선택된 셀의 테두리/배경 (HBorderFill)
*/
func (p *HCellBorderFill) SelCellsBorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "SelCellsBorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "SelCellsBorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
전체 셀의 테두리/배경 (HBorderFill)
*/
func (p *HCellBorderFill) AllCellsBorderFill(v ...*HBorderFill) *HBorderFill {
	var para HBorderFill
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "AllCellsBorderFill", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "AllCellsBorderFill")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HTableDeleteLine:

	HTableDeleteLine :  표 줄/칸 삭제 개체
*/
type HTableDeleteLine struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableDeleteLine) HSet() *HSet {
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
줄/칸
*/
func (p *HTableDeleteLine) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HTableInsertLine:

	HTableInsertLine :  표 줄/칸 삽입 개체
*/
type HTableInsertLine struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableInsertLine) HSet() *HSet {
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
방향
*/
func (p *HTableInsertLine) Side(v ...uint16) int {
	return funcParaSetInt(p.variant, "Side", v)
}

/*
개수
*/
func (p *HTableInsertLine) Count(v ...uint16) int {
	return funcParaSetInt(p.variant, "Count", v)
}

/*
HCaption:

	HCaption :  캡션 속성 개체
*/
type HCaption struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCaption) HSet() *HSet {
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
방향

	0 = 왼쪽
	1 = 오른쪽
	2 = 위
	3 = 아래
*/
func (p *HCaption) Side(v ...uint16) int {
	return funcParaSetInt(p.variant, "Side", v)
}

/*
캡션 폭 (가로 방향일 때만 사용됨)
*/
func (p *HCaption) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
캡션과 틀 사이 간격
*/
func (p *HCaption) Gap(v ...int32) int {
	return funcParaSetInt(p.variant, "Gap", v)
}

/*
캡션 폭에 마진을 포함할지 여부 (세로 방향일 때만 사용됨)
*/
func (p *HCaption) CapFullSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "CapFullSize", v)
}

/*
HHyperlinkJump:

	HHyperlinkJump :  하이퍼 링크 점프 속성 개체
*/
type HHyperlinkJump struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HHyperlinkJump) HSet() *HSet {
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
Source에 대한 Object Path
*/
func (p *HHyperlinkJump) Source(v ...string) string {
	return funcParaSetString(p.variant, "Source", v)
}

/*
이동할 Target에 대한 Object Path
*/
func (p *HHyperlinkJump) Target(v ...string) string {
	return funcParaSetString(p.variant, "Target", v)
}

/*
HTableDrawPen:

	HTableDrawPen :  마우스로 테이블을 그릴 때 쓰이는 펜
*/
type HTableDrawPen struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableDrawPen) HSet() *HSet {
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
Table을 그리는 연필(펜)의 선 종류
*/
func (p *HTableDrawPen) Style(v ...uint16) int {
	return funcParaSetInt(p.variant, "Style", v)
}

/*
Table을 그리는 연필(펜)의 선 굵기
*/
func (p *HTableDrawPen) Width(v ...uint16) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
Table을 그리는 연필(펜)의 선 색깔 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HTableDrawPen) Color(v ...uint32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
HEqEdit:

	HEqEdit :  수식 속성 개체
*/
type HEqEdit struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HEqEdit) HSet() *HSet {
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
글자처럼 취급 on / off
*/
func (p *HEqEdit) TreatAsChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "TreatAsChar", v)
}

/*
줄 간격에 영향을 줄지 여부 on / off * "TreatAsChar" = TRUE일 때만 사용됨
*/
func (p *HEqEdit) AffectsLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "AffectsLine", v)
}

/*
세로 위치의 기준 * "TreatAsChar" = FALSE일 때만 사용됨
*/
func (p *HEqEdit) VertRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertRelTo", v)
}

/*
"VertRelTo"에 대한 상대적인 배열 방식.
*/
func (p *HEqEdit) VertAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "VertAlign", v)
}

/*
"VertRelTo"와 "VertAlign"을 기준점으로 한 상대적인 오프셋 값 HWPUNIT 단위.
*/
func (p *HEqEdit) VertOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "VertOffset", v)
}

/*
가로 위치의 기준 * "TreatAsChar" = FALSE일 때만 사용됨
*/
func (p *HEqEdit) HorzRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzRelTo", v)
}

/*
"HorzRelTo"에 대한 상대적인 배열 방식
*/
func (p *HEqEdit) HorzAlign(v ...uint16) int {
	return funcParaSetInt(p.variant, "HorzAlign", v)
}

/*
"HorzRelTo"와 "HorzAlign"을 기준점으로 한 상대적인 오프셋 값. HWPUNIT 단위.
*/
func (p *HEqEdit) HorzOffset(v ...int32) int {
	return funcParaSetInt(p.variant, "HorzOffset", v)
}

/*
오브젝트의 세로 위치를 본문 영역으로 제한할지 여부 on / off
*/
func (p *HEqEdit) FlowWithText(v ...uint16) int {
	return funcParaSetInt(p.variant, "FlowWithText", v)
}

/*
다른 오브젝트와 겹치는 것을 허용할지 여부 on / off  * "TreatAsChar" = FALSE일 때만 사용됨 * "FlowWithText" = TRUE이면 언제나 FALSE로 간주한다.
*/
func (p *HEqEdit) AllowOverlap(v ...uint16) int {
	return funcParaSetInt(p.variant, "AllowOverlap", v)
}

/*
오브젝트 폭의 기준
*/
func (p *HEqEdit) WidthRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "WidthRelTo", v)
}

/*
오브젝트 폭의 값
*/
func (p *HEqEdit) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
오브젝트 높이의 기준
*/
func (p *HEqEdit) HeightRelTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "HeightRelTo", v)
}

/*
오브젝트의 높이 값
*/
func (p *HEqEdit) Height(v ...int32) int {
	return funcParaSetInt(p.variant, "Height", v)
}

/*
크기 보호 on / off
*/
func (p *HEqEdit) ProtectSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "ProtectSize", v)
}

/*
오브젝트 주위를 텍스트가 어떻게 흘러갈지 지정하는 옵션. * "TreatAsChar" = FALSE일 때만 사용됨 0 = bound rect를 따라 1 = 좌, 우에는 텍스트를 배치하지 않음 2 = 글과 겹치게 하여 글 뒤로 3 = 글과 겹치게 하여 글 앞으로 4 = 오브젝트의 outline을 따라 5 = 오브젝트 내부의 빈 공간까지
*/
func (p *HEqEdit) TextWrap(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextWrap", v)
}

/*
오브젝트의 좌/우 어느쪽에 글을 배치할지 지정하는 옵션 * "TextWrap"가 0, 4, 5일 때만 사용된다. 0 = 양쪽 1 = 왼쪽 2 = 오른쪽 3 = 큰쪽
*/
func (p *HEqEdit) TextFlow(v ...uint16) int {
	return funcParaSetInt(p.variant, "TextFlow", v)
}

/*
오브젝트의 바깥 여백, (왼쪽) HWPUNIT 단위
*/
func (p *HEqEdit) OutsideMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginLeft", v)
}

/*
오브젝트의 바깥 여백, (오른쪽) HWPUNIT 단위
*/
func (p *HEqEdit) OutsideMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginRight", v)
}

/*
오브젝트의 바깥 여백, (위) HWPUNIT 단위
*/
func (p *HEqEdit) OutsideMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginTop", v)
}

/*
오브젝트의 바깥 여백, (아래) HWPUNIT 단위
*/
func (p *HEqEdit) OutsideMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "OutsideMarginBottom", v)
}

/*
이 개체가 속하는 번호 범주
*/
func (p *HEqEdit) NumberingType(v ...uint16) int {
	return funcParaSetInt(p.variant, "NumberingType", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 폭의 값 글상자등이 늘어나면 늘어난 값을 계산해서 가진다. 단위는 "Width"의 단위와 같다.
*/
func (p *HEqEdit) LayoutWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutWidth", v)
}

/*
오브젝트가 페이지에 배열될 때 계산되는 높이 값 글상자등이 늘어나면 늘어난 값을 계산해서 가진다. 단위는 "Height"의 단위와 같다.
*/
func (p *HEqEdit) LayoutHeight(v ...int32) int {
	return funcParaSetInt(p.variant, "LayoutHeight", v)
}

/*
오브젝트 선택 가능 on / off
*/
func (p *HEqEdit) Lock(v ...uint16) int {
	return funcParaSetInt(p.variant, "Lock", v)
}

/*
쪽나눔방지II on/off
*/
func (p *HEqEdit) HoldAnchorObj(v ...uint16) int {
	return funcParaSetInt(p.variant, "HoldAnchorObj", v)
}

/*
개체가 존재 하는 페이지
*/
func (p *HEqEdit) PageNumber(v ...uint32) int {
	return funcParaSetInt(p.variant, "PageNumber", v)
}

/*
개체 Selection 상태 TRUE/FASLE
*/
func (p *HEqEdit) AdjustSelection(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustSelection", v)
}

/*
글상자로 TRUE/FASLE
*/
func (p *HEqEdit) AdjustTextbox(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustTextbox", v)
}

/*
앞 개체 속성 따라가기 TRUE/FASLE
*/
func (p *HEqEdit) AdjustPrevObjAttr(v ...uint16) int {
	return funcParaSetInt(p.variant, "AdjustPrevObjAttr", v)
}

/*
수식 스크립트이다
*/
func (p *HEqEdit) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
HWPUNIT인 수식의 Base 크기이다. (기본값은 POINT 10 )
*/
func (p *HEqEdit) BaseUnit(v ...int32) int {
	return funcParaSetInt(p.variant, "BaseUnit", v)
}

/*
HWPUNIT인 수식의 Base 크기이다. (기본값은 WINDOWTEXT 색)
*/
func (p *HEqEdit) Color(v ...int32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
줄 단위를 사용할지의 여부

	0: 글자 단위 영역
	1: 줄 단위 영역
*/
func (p *HEqEdit) LineMode(v ...int32) int {
	return funcParaSetInt(p.variant, "LineMode", v)
}

/*
수식 스크립트 버전
*/
func (p *HEqEdit) Version(v ...string) string {
	return funcParaSetString(p.variant, "Version", v)
}

/*
HCaptureEnd:

	HCaptureEnd :  갈무리의 시작/끝점 좌표 속성 개체
*/
type HCaptureEnd struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCaptureEnd) HSet() *HSet {
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
시작점과 끝점의 좌표 (페이지 X좌표)
*/
func (p *HCaptureEnd) BeginX(v ...int32) int {
	return funcParaSetInt(p.variant, "BeginX", v)
}

/*
시작점과 끝점의 좌표 (페이지 y좌표)
*/
func (p *HCaptureEnd) BeginY(v ...int32) int {
	return funcParaSetInt(p.variant, "BeginY", v)
}

/*
끝점의 X 좌표(페이지 X좌표)
*/
func (p *HCaptureEnd) EndX(v ...int32) int {
	return funcParaSetInt(p.variant, "EndX", v)
}

/*
끝점의 X 좌표(페이지 y좌표)
*/
func (p *HCaptureEnd) EndY(v ...int32) int {
	return funcParaSetInt(p.variant, "EndY", v)
}

/*
페이지 번호
*/
func (p *HCaptureEnd) PageNum(v ...uint32) int {
	return funcParaSetInt(p.variant, "PageNum", v)
}

/*
HPrintToImage:

	HPrintToImage :  문서 내용을 이미지로 변환하기 개체
*/
type HPrintToImage struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPrintToImage) HSet() *HSet {
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
그림 형식 0 : none 1 : BMP 2 : GIF 3 : PNG 4 : JPG 5 : WMF
*/
func (p *HPrintToImage) Format(v ...uint16) int {
	return funcParaSetInt(p.variant, "Format", v)
}

/*
그림 경로
*/
func (p *HPrintToImage) FileName(v ...string) string {
	return funcParaSetString(p.variant, "FileName", v)
}

/*
색상 수 (bits: 8, 16...)
*/
func (p *HPrintToImage) Colordepth(v ...uint16) int {
	return funcParaSetInt(p.variant, "Colordepth", v)
}

/*
해상도
*/
func (p *HPrintToImage) Resolution(v ...uint16) int {
	return funcParaSetInt(p.variant, "Resolution", v)
}

/*
HFileSaveBlock:

	HFileSaveBlock :  블럭 저장 개체
*/
type HFileSaveBlock struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileSaveBlock) HSet() *HSet {
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
파일 이름
*/
func (p *HFileSaveBlock) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
파일 포맷
*/
func (p *HFileSaveBlock) Format(v ...string) string {
	return funcParaSetString(p.variant, "Format", v)
}

/*
Argument
*/
func (p *HFileSaveBlock) Argument(v ...string) string {
	return funcParaSetString(p.variant, "Argument", v)
}

/*
HFileSendMail:

	HFileSendMail :  메일 보내기 개체
*/
type HFileSendMail struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileSendMail) HSet() *HSet {
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
받는 사람
*/
func (p *HFileSendMail) To(v ...string) string {
	return funcParaSetString(p.variant, "To", v)
}

/*
메일 종류 0 = 본문 1 = 첨부 파일 2 = 사용자 직접 메일에 첨부 파일
*/
func (p *HFileSendMail) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
제목
*/
func (p *HFileSendMail) Subject(v ...string) string {
	return funcParaSetString(p.variant, "Subject", v)
}

/*
HHncMessageBox:

	HHncMessageBox :  메세지 박스 개체
*/
type HHncMessageBox struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HHncMessageBox) HSet() *HSet {
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
메시지 박스에 넣을 문자열
*/
func (p *HHncMessageBox) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
메시지 박스에서 사용할 Flag
*/
func (p *HHncMessageBox) Flag(v ...uint16) int {
	return funcParaSetInt(p.variant, "Flag", v)
}

/*
메시지 박스의 리턴값
*/
func (p *HHncMessageBox) Result(v ...uint16) int {
	return funcParaSetInt(p.variant, "Result", v)
}

/*
HTableChartInfo:

	HTableChartInfo :  표에 링크된 차트의 정보 개체
*/
type HTableChartInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableChartInfo) HSet() *HSet {
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
표에 링크된 차드의 컨트롤 ID
*/
func (p *HTableChartInfo) CtrlID(v ...uint32) int {
	return funcParaSetInt(p.variant, "CtrlID", v)
}

/*
셀 블록의 시작 ROW
*/
func (p *HTableChartInfo) StartRow(v ...uint32) int {
	return funcParaSetInt(p.variant, "StartRow", v)
}

/*
셀 블록의 시작 COL
*/
func (p *HTableChartInfo) StartCo(v ...uint32) int {
	return funcParaSetInt(p.variant, "StartCo", v)
}

/*
셀 블록의 끝 ROW
*/
func (p *HTableChartInfo) EndRow(v ...uint32) int {
	return funcParaSetInt(p.variant, "EndRow", v)
}

/*
셀 블록의 끝 COL
*/
func (p *HTableChartInfo) EndCo(v ...uint32) int {
	return funcParaSetInt(p.variant, "EndCo", v)
}

/*
표에 링크되어 있는 NEXT 차트
*/
func (p *HTableChartInfo) NextCtrlID(v ...uint32) int {
	return funcParaSetInt(p.variant, "NextCtrlID", v)
}

/*
HTableTblToStr:

	HTableTblToStr :  표를 문자열로 바꾸기 속성 개체
*/
type HTableTblToStr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableTblToStr) HSet() *HSet {
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
분리 문자(탭, 쉼표, 공백)
*/
func (p *HTableTblToStr) DelimiterType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DelimiterType", v)
}

/*
사용자 정의 필드 구분 기호
*/
func (p *HTableTblToStr) UserDefine(v ...string) string {
	return funcParaSetString(p.variant, "UserDefine", v)
}

/*
HTableStrToTbl:

	HTableStrToTbl :  문자열을 표로
*/
type HTableStrToTbl struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableStrToTbl) HSet() *HSet {
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
분리 문자(탭, 쉼표, 공백)
*/
func (p *HTableStrToTbl) DelimiterType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DelimiterType", v)
}

/*
사용자 정의 필드 구분 기호
*/
func (p *HTableStrToTbl) UserDefine(v ...string) string {
	return funcParaSetString(p.variant, "UserDefine", v)
}

/*
자동으로 할 것인지 분리 문자를 지정 할 것인지를 결정
*/
func (p *HTableStrToTbl) AutoOrDefine(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoOrDefine", v)
}

/*
선택 사항 (구분자 유지)
*/
func (p *HTableStrToTbl) KeepSeperator(v ...uint16) int {
	return funcParaSetInt(p.variant, "KeepSeperator", v)
}

/*
기타 문자 필드 구분 기호
*/
func (p *HTableStrToTbl) DelimiterEtc(v ...string) string {
	return funcParaSetString(p.variant, "DelimiterEtc", v)
}

/*
HFlashProperties:

	HFlashProperties :  플래시 삽입
*/
type HFlashProperties struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFlashProperties) HSet() *HSet {
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
경로의 BASE
*/
func (p *HFlashProperties) Base(v ...string) string {
	return funcParaSetString(p.variant, "Base", v)
}

/*
Play Qulaity
*/
func (p *HFlashProperties) Qulaity(v ...string) string {
	return funcParaSetString(p.variant, "Qulaity", v)
}

/*
Scale
*/
func (p *HFlashProperties) Scale(v ...string) string {
	return funcParaSetString(p.variant, "Scale", v)
}

/*
윈도우 모드
*/
func (p *HFlashProperties) WMode(v ...string) string {
	return funcParaSetString(p.variant, "WMode", v)
}

/*
자동 재생 여부
*/
func (p *HFlashProperties) AutoPlay(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoPlay", v)
}

/*
반복 재생 어부
*/
func (p *HFlashProperties) LoopPlay(v ...uint16) int {
	return funcParaSetInt(p.variant, "LoopPlay", v)
}

/*
메뉴 보이기
*/
func (p *HFlashProperties) ShowMenu(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowMenu", v)
}

/*
배경색 (COLORREF)
*/
func (p *HFlashProperties) BgColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "BgColor", v)
}

/*
HMovieProperties:

	HMovieProperties :  동영상 파일 삽입시 필요한 옵션
*/
type HMovieProperties struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMovieProperties) HSet() *HSet {
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
경로의 BASE
*/
func (p *HMovieProperties) GetBase(v ...string) string {
	return funcParaSetString(p.variant, "GetBase", v)
}

/*
캡션 파일의 경로
*/
func (p *HMovieProperties) Caption(v ...string) string {
	return funcParaSetString(p.variant, "Caption", v)
}

/*
자동 재생 여부
*/
func (p *HMovieProperties) AutoPlay(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoPlay", v)
}

/*
리와인드 여부
*/
func (p *HMovieProperties) AutoRewind(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoRewind", v)
}

/*
메뉴 보이기
*/
func (p *HMovieProperties) ShowMenu(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowMenu", v)
}

/*
컨트롤 패널 보이기
*/
func (p *HMovieProperties) ShowCtrlPanel(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowCtrlPanel", v)
}

/*
위치 컨트롤 보이기
*/
func (p *HMovieProperties) ShowPosCtrl(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowPosCtrl", v)
}

/*
위치 컨트롤 변경 가능
*/
func (p *HMovieProperties) EnablePos(v ...uint16) int {
	return funcParaSetInt(p.variant, "EnablePos", v)
}

/*
트랙바 보기
*/
func (p *HMovieProperties) ShowTrackBar(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowTrackBar", v)
}

/*
트랙바 변경 가능
*/
func (p *HMovieProperties) EnableTrack(v ...uint16) int {
	return funcParaSetInt(p.variant, "EnableTrack", v)
}

/*
캡션 보이기
*/
func (p *HMovieProperties) ShowChaption(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowChaption", v)
}

/*
오디오 설정 보이기
*/
func (p *HMovieProperties) ShowAudio(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowAudio", v)
}

/*
상태 바 보기 (진행 시간 등을 표시)
*/
func (p *HMovieProperties) ShowStatus(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowStatus", v)
}

/*
HFindReplace:

	HFindReplace :  찾아 바꾸기 옵션
*/
type HFindReplace struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFindReplace) HSet() *HSet {
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
찾을 문자열
*/
func (p *HFindReplace) FindString(v ...string) string {
	return funcParaSetString(p.variant, "FindString", v)
}

/*
바꿀 문자열
*/
func (p *HFindReplace) ReplaceString(v ...string) string {
	return funcParaSetString(p.variant, "ReplaceString", v)
}

/*
찾을 방향 0 = 앞으로 1 = 뒤로 2 = 문서 전체
*/
func (p *HFindReplace) Direction(v ...uint16) int {
	return funcParaSetInt(p.variant, "Direction", v)
}

/*
대소문자 구별 (on/off)
*/
func (p *HFindReplace) MatchCase(v ...uint16) int {
	return funcParaSetInt(p.variant, "MatchCase", v)
}

/*
문자열 결합 (on/off)
*/
func (p *HFindReplace) AllWordForms(v ...uint16) int {
	return funcParaSetInt(p.variant, "AllWordForms", v)
}

/*
여러 단어 찾기 (on/off)
*/
func (p *HFindReplace) SeveralWords(v ...uint16) int {
	return funcParaSetInt(p.variant, "SeveralWords", v)
}

/*
아무개 문자 (on/off)
*/
func (p *HFindReplace) UseWildCards(v ...uint16) int {
	return funcParaSetInt(p.variant, "UseWildCards", v)
}

/*
온전한 낱말 (on/off)
*/
func (p *HFindReplace) WholeWordOnly(v ...uint16) int {
	return funcParaSetInt(p.variant, "WholeWordOnly", v)
}

/*
토씨 자동 교정 (on/off)
*/
func (p *HFindReplace) AutoSpell(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoSpell", v)
}

/*
찾아 바꾸기 모드 (on/off)
*/
func (p *HFindReplace) ReplaceMode(v ...uint16) int {
	return funcParaSetInt(p.variant, "ReplaceMode", v)
}

/*
찾을 문자열 무시 (on/off)
*/
func (p *HFindReplace) IgnoreFindString(v ...uint16) int {
	return funcParaSetInt(p.variant, "IgnoreFindString", v)
}

/*
바꿀문자열 무시 (on/off)
*/
func (p *HFindReplace) IgnoreReplaceString(v ...uint16) int {
	return funcParaSetInt(p.variant, "IgnoreReplaceString", v)
}

/*
찾을 글자 모양 (HCharShape)
*/
func (p *HFindReplace) FindCharShape(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FindCharShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FindCharShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
찾을 문단 모양 (HParaShape)
*/
func (p *HFindReplace) FindParaShape(v ...*HParaShape) *HParaShape {
	var para HParaShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FindParaShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FindParaShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
바꿀 글자 모양 (HCharShape)
*/
func (p *HFindReplace) ReplaceCharShape(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ReplaceCharShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ReplaceCharShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
바꿀 문단 모양 (HParaShape)
*/
func (p *HFindReplace) ReplaceParaShape(v ...*HParaShape) *HParaShape {
	var para HParaShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "ReplaceParaShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "ReplaceParaShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
찾을 스타일
*/
func (p *HFindReplace) FindStyle(v ...string) string {
	return funcParaSetString(p.variant, "FindStyle", v)
}

/*
바꿀 스타일
*/
func (p *HFindReplace) ReplaceStyle(v ...string) string {
	return funcParaSetString(p.variant, "ReplaceStyle", v)
}

/*
메시지 박스 표시 안함. (on/off)
*/
func (p *HFindReplace) IgnoreMessage(v ...uint16) int {
	return funcParaSetInt(p.variant, "IgnoreMessage", v)
}

/*
한글 음으로 한자 찾기. (on/off)
*/
func (p *HFindReplace) HanjaFromHangul(v ...uint16) int {
	return funcParaSetInt(p.variant, "HanjaFromHangul", v)
}

/*
자소 단위로 찾기. (on/off)
*/
func (p *HFindReplace) FindJaso(v ...uint16) int {
	return funcParaSetInt(p.variant, "FindJaso", v)
}

/*
정규식(조건식)으로 찾기. (on/off)
*/
func (p *HFindReplace) FindRegExp(v ...uint16) int {
	return funcParaSetInt(p.variant, "FindRegExp", v)
}

/*
FindType ???
*/
func (p *HFindReplace) FindType(v ...uint16) int {
	return funcParaSetInt(p.variant, "FindType", v)
}

/*
HTableTemplate:

	HTableTemplate :  표 마당
*/
type HTableTemplate struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableTemplate) HSet() *HSet {
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
적용할 서식

	0x01 = 테두리
	0x02 = 글자 모양과 문단 모양
	0x04 = 셀 배경
	0x08 = 그레이 스케일
*/
func (p *HTableTemplate) Format(v ...uint32) int {
	return funcParaSetInt(p.variant, "Format", v)
}

/*
적용 대상

	0x01 = 제목 줄
	0x02 = 마지막 줄
	0x04 = 첫째 칸
	0x08 = 마지막 칸
*/
func (p *HTableTemplate) ApplyTarget(v ...uint32) int {
	return funcParaSetInt(p.variant, "ApplyTarget", v)
}

/*
표 마당 파일 이름
*/
func (p *HTableTemplate) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
표 만들기 모드 (표 만들기에서 제목 줄에 제목 속성 넣기 위해)
*/
func (p *HTableTemplate) CreateMode(v ...uint16) int {
	return funcParaSetInt(p.variant, "CreateMode", v)
}

/*
HFileSecurity:

	HFileSecurity :  문서 보호
*/
type HFileSecurity struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileSecurity) HSet() *HSet {
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
TRUE = 유니코드 모든 문자를 사용, FALSE = 영문자만 사용
*/
func (p *HFileSecurity) PasswordFullRange(v ...uint16) int {
	return funcParaSetInt(p.variant, "PasswordFullRange", v)
}

/*
문서 보호를 설정하기 위한 암호
*/
func (p *HFileSecurity) PasswordString(v ...string) string {
	return funcParaSetString(p.variant, "PasswordString", v)
}

/*
TRUE = 문서 인쇄 불가능 FALSE = 문서 인쇄 가능. on/off
*/
func (p *HFileSecurity) NoPrint(v ...uint16) int {
	return funcParaSetInt(p.variant, "NoPrint", v)
}

/*
TRUE = 문서 내용 복사 및 추출 불가능 FALSE = 문서 내용 복사 및 추출 가능. on/off
*/
func (p *HFileSecurity) NoCopy(v ...uint16) int {
	return funcParaSetInt(p.variant, "NoCopy", v)
}

/*
TRUE = 문서 암호를 확인, FALSE = 문서 암호를 설정
*/
func (p *HFileSecurity) PasswordAsk(v ...uint16) int {
	return funcParaSetInt(p.variant, "PasswordAsk", v)
}

/*
HLabel:

	HLabel :  라벨
*/
type HLabel struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HLabel) HSet() *HSet {
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
용지 여백 : 위쪽
*/
func (p *HLabel) TopMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "TopMargin", v)
}

/*
용지 여백 : 왼쪽
*/
func (p *HLabel) LeftMargin(v ...int32) int {
	return funcParaSetInt(p.variant, "LeftMargin", v)
}

/*
이름표 크기 : 폭
*/
func (p *HLabel) BoxWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "BoxWidth", v)
}

/*
이름표 크기 : 길이
*/
func (p *HLabel) BoxLength(v ...int32) int {
	return funcParaSetInt(p.variant, "BoxLength", v)
}

/*
이름표 크기 : 좌우
*/
func (p *HLabel) BoxMarginHor(v ...int32) int {
	return funcParaSetInt(p.variant, "BoxMarginHor", v)
}

/*
이름표 크기 : 상하
*/
func (p *HLabel) BoxMarginVer(v ...int32) int {
	return funcParaSetInt(p.variant, "BoxMarginVer", v)
}

/*
이름표 개수 : 줄 수(or 세로)
*/
func (p *HLabel) LabelCols(v ...int32) int {
	return funcParaSetInt(p.variant, "LabelCols", v)
}

/*
이름표 개수 : 칸 수(or 가로)
*/
func (p *HLabel) LabelRows(v ...int32) int {
	return funcParaSetInt(p.variant, "LabelRows", v)
}

/*
용지 방향
*/
func (p *HLabel) Landscape(v ...int16) int {
	return funcParaSetInt(p.variant, "Landscape", v)
}

/*
문서의 폭
*/
func (p *HLabel) PageWidth(v ...int32) int {
	return funcParaSetInt(p.variant, "PageWidth", v)
}

/*
문서의 길이
*/
func (p *HLabel) PageLen(v ...int32) int {
	return funcParaSetInt(p.variant, "PageLen", v)
}

/*
문서의 쪽수
*/
func (p *HLabel) PageCnt(v ...int32) int {
	return funcParaSetInt(p.variant, "PageCnt", v)
}

/*
HMarkpenShape:

	HMarkpenShape :  형광펜
*/
type HMarkpenShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMarkpenShape) HSet() *HSet {
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
형광펜색 (COLORREF)
*/
func (p *HMarkpenShape) Color(v ...uint32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
HPasteHtml:

	HPasteHtml :  HTML 문서 붙이기
*/
type HPasteHtml struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPasteHtml) HSet() *HSet {
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
기본 값으로 지정
*/
func (p *HPasteHtml) Default(v ...uint32) int {
	return funcParaSetInt(p.variant, "Default", v)
}

/*
데이터 형식(원본 형식 유지/텍스트 형식으로 붙이기)
*/
func (p *HPasteHtml) Format(v ...uint32) int {
	return funcParaSetInt(p.variant, "Format", v)
}

/*
HOleCreation:

	HOleCreation :  OLE 개체
*/
type HOleCreation struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HOleCreation) HSet() *HSet {
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
생성 방식

	0 = 새로 개체 생성
	1 = 파일로부터 개체 생성
	2 = 파일로 링크된 개체 생성
*/
func (p *HOleCreation) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
클래스 ID ("Type" = 0 일때 사용)
*/
func (p *HOleCreation) Clsid(v ...string) string {
	return funcParaSetString(p.variant, "Clsid", v)
}

/*
파일 경로 ("Type" = 1 또는 "Type" = 2 일 때 사용)
*/
func (p *HOleCreation) Path(v ...string) string {
	return funcParaSetString(p.variant, "Path", v)
}

/*
draw aspect
*/
func (p *HOleCreation) Aspect(v ...int32) int {
	return funcParaSetInt(p.variant, "Aspect", v)
}

/*
DVASPECT_ICON일 때 적용할 아이콘 데이터 (TYPE_EMBEDDING) CopyMetaFile()로 저장된 형식이다.

	미구현
*/
func (p *HOleCreation) IconMetafile() {
}

/*
DVASPECT_ICON일 때 METAFILEPICT의 데이터들
*/
func (p *HOleCreation) IconMM(v ...int32) int {
	return funcParaSetInt(p.variant, "IconMM", v)
}

/*
DVASPECT_ICON일 때 METAFILEPICT의 데이터들
*/
func (p *HOleCreation) IconXext(v ...int32) int {
	return funcParaSetInt(p.variant, "IconXext", v)
}

/*
DVASPECT_ICON일 때 METAFILEPICT의 데이터들
*/
func (p *HOleCreation) IconYext(v ...int32) int {
	return funcParaSetInt(p.variant, "IconYext", v)
}

/*
한글 내부에서 생성한 OCX 인지에 대한 표시
*/
func (p *HOleCreation) InnerOCX(v ...int32) int {
	return funcParaSetInt(p.variant, "InnerOCX", v)
}

/*
초기 shape object 속성
*/
func (p *HOleCreation) SoProperties(v ...*HShapeObject) *HShapeObject {
	var para HShapeObject
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "SoProperties", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "SoProperties")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
플래시 파일 삽입 시 필요한 옵션 셋
*/
func (p *HOleCreation) FlashProperties(v ...*HFlashProperties) *HFlashProperties {
	var para HFlashProperties
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "FlashProperties", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FlashProperties")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
동영상 파일 삽입 시 필요한 옵션 셋
*/
func (p *HOleCreation) MovieProperties(v ...*HMovieProperties) *HMovieProperties {
	var para HMovieProperties
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "MovieProperties", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "MovieProperties")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HTableSwap:

	HTableSwap :  표 뒤집기
*/
type HTableSwap struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HTableSwap) HSet() *HSet {
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
표 뒤집기 형식
*/
func (p *HTableSwap) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
여백 뒤집기 지원
*/
func (p *HTableSwap) SwapMargin(v ...uint16) int {
	return funcParaSetInt(p.variant, "SwapMargin", v)
}

/*
HVersionInfo:

	HVersionInfo :  버전 정보
*/
type HVersionInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HVersionInfo) HSet() *HSet {
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
버전 비교용 소스 패스
*/
func (p *HVersionInfo) SourcePath(v ...string) string {
	return funcParaSetString(p.variant, "SourcePath", v)
}

/*
버전 비교용 타깃 패스
*/
func (p *HVersionInfo) TargetPath(v ...string) string {
	return funcParaSetString(p.variant, "TargetPath", v)
}

/*
작성자 정보
*/
func (p *HVersionInfo) ItemInfoWriter(v ...string) string {
	return funcParaSetString(p.variant, "ItemInfoWriter", v)
}

/*
해당 버전에 대한 설명
*/
func (p *HVersionInfo) ItemInfoDescription(v ...string) string {
	return funcParaSetString(p.variant, "ItemInfoDescription", v)
}

/*
날짜 정보, window API FILETIME의 HIWORD
*/
func (p *HVersionInfo) ItemInfoTimeHi(v ...uint32) int {
	return funcParaSetInt(p.variant, "ItemInfoTimeHi", v)
}

/*
날짜 정보, window API FILETIME의 LOWORD
*/
func (p *HVersionInfo) ItemInfoTimeLo(v ...uint32) int {
	return funcParaSetInt(p.variant, "ItemInfoTimeLo", v)
}

/*
히스토리 정보 수정 플랙
*/
func (p *HVersionInfo) ItemInfoLock(v ...uint16) int {
	return funcParaSetInt(p.variant, "ItemInfoLock", v)
}

/*
HMemoShape:

	HMemoShape :  메모 모양
*/
type HMemoShape struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMemoShape) HSet() *HSet {
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
너비 (HWPUNIT)
*/
func (p *HMemoShape) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
선 종류
*/
func (p *HMemoShape) LineType(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineType", v)
}

/*
선 굵기
*/
func (p *HMemoShape) LineWidth(v ...uint16) int {
	return funcParaSetInt(p.variant, "LineWidth", v)
}

/*
선 색깔 (COLORREF)
*/
func (p *HMemoShape) LineColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "LineColor", v)
}

/*
채우기 색깔 (COLORREF)
*/
func (p *HMemoShape) FillColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "FillColor", v)
}

/*
활성화 된 채우기 색깔 (COLORREF)
*/
func (p *HMemoShape) ActiveFillColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "ActiveFillColor", v)
}

/*
HSort:

	HSort :  소트
*/
type HSort struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSort) HSet() *HSet {
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
PIT_ARRAY 생성

	KeyOption을 위해 필요
*/
func (p *HSort) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
키 콤보에서 선택된 키를 저장함
*/
func (p *HSort) KeyOption() *HArray {
	return GetHArray(p.variant, "KeyOption")
}

/*
자소 단위 비교 Flag - 종,중,초
*/
func (p *HSort) CheckJasoReverse(v ...uint16) int {
	return funcParaSetInt(p.variant, "CheckJasoReverse", v)
}

/*
필드 구분 기호 형식
*/
func (p *HSort) DelimiterType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DelimiterType", v)
}

/*
필드 구분 기호들
*/
func (p *HSort) DelimiterChars(v ...string) string {
	return funcParaSetString(p.variant, "DelimiterChars", v)
}

/*
연속되는 구분기호 무시 Flag
*/
func (p *HSort) IgnoreMultiDelimiter(v ...uint16) int {
	return funcParaSetInt(p.variant, "IgnoreMultiDelimiter", v)
}

/*
단어 뒤에서 부터 비교 Flag
*/
func (p *HSort) CheckFromRear(v ...uint16) int {
	return funcParaSetInt(p.variant, "CheckFromRear", v)
}

/*
두 자리 년도 확장 check Flag
*/
func (p *HSort) CheckExtendYear(v ...uint16) int {
	return funcParaSetInt(p.variant, "CheckExtendYear", v)
}

/*
두 자리 년도 시작 년도
*/
func (p *HSort) YearBase(v ...uint32) int {
	return funcParaSetInt(p.variant, "YearBase", v)
}

/*
사전 언어 순서 값
*/
func (p *HSort) LangOrderType(v ...uint16) int {
	return funcParaSetInt(p.variant, "LangOrderType", v)
}

/*
자소 단위 비교 Flag - 초,중,종
*/
func (p *HSort) CheckJaso(v ...uint16) int {
	return funcParaSetInt(p.variant, "CheckJaso", v)
}

/*
HDrawLayOut:

	HDrawLayOut :  그리기 개체 레이아웃
*/
type HDrawLayOut struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawLayOut) HSet() *HSet {
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
선택된 개체가 선택되었을 때 나타나는 점의 생성
*/
func (p *HDrawLayOut) CreateNumPt(v ...uint32) int {
	return funcParaSetInt(p.variant, "CreateNumPt", v)
}

/*
PIT_ARRAY 생성

	CreatePt, CurveSegmentInfo을 위해 필요
*/
func (p *HDrawLayOut) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
점을 생성
*/
func (p *HDrawLayOut) CreatePt() *HArray {
	return GetHArray(p.variant, "CreatePt")
}

/*
곡선 세그먼트 정보
*/
func (p *HDrawLayOut) CurveSegmentInfo() *HArray {
	return GetHArray(p.variant, "CurveSegmentInfo")
}

/*
HDrawLineAttr:

	HDrawLineAttr :  그리기 개체 선
*/
type HDrawLineAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawLineAttr) HSet() *HSet {
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
선 색깔 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HDrawLineAttr) Color(v ...uint32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
선의 style
*/
func (p *HDrawLineAttr) Style(v ...int32) int {
	return funcParaSetInt(p.variant, "Style", v)
}

/*
선의 width
*/
func (p *HDrawLineAttr) Width(v ...int32) int {
	return funcParaSetInt(p.variant, "Width", v)
}

/*
선의 endcap
*/
func (p *HDrawLineAttr) EndCap(v ...int32) int {
	return funcParaSetInt(p.variant, "EndCap", v)
}

/*
선의 시작 화살표 모양
*/
func (p *HDrawLineAttr) HeadStyle(v ...int32) int {
	return funcParaSetInt(p.variant, "HeadStyle", v)
}

/*
선의 끝 화살표 모양
*/
func (p *HDrawLineAttr) TailStyle(v ...int32) int {
	return funcParaSetInt(p.variant, "TailStyle", v)
}

/*
선의 시작 화살표 크기
*/
func (p *HDrawLineAttr) HeadSize(v ...int32) int {
	return funcParaSetInt(p.variant, "HeadSize", v)
}

/*
선의 끝 화살표 크기
*/
func (p *HDrawLineAttr) TailSize(v ...int32) int {
	return funcParaSetInt(p.variant, "TailSize", v)
}

/*
선의 시작 화살표 채움 유무
*/
func (p *HDrawLineAttr) HeadFill(v ...int32) int {
	return funcParaSetInt(p.variant, "HeadFill", v)
}

/*
선의 끝 화살표 채움 유무
*/
func (p *HDrawLineAttr) TailFill(v ...int32) int {
	return funcParaSetInt(p.variant, "HeadFill", v)
}

/*
선두께 외부/내부/일반 적용
*/
func (p *HDrawLineAttr) OutLineStyle(v ...uint32) int {
	return funcParaSetInt(p.variant, "OutLineStyle", v)
}

/*
HDrawFillAttr:

	HDrawFillAttr :  그리기 개체 채우기 속성
*/
type HDrawFillAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawFillAttr) HSet() *HSet {
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
브러쉬 타입
*/
func (p *HDrawFillAttr) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
그러데이션 타입
*/
func (p *HDrawFillAttr) GradationType(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationType", v)
}

/*
그러데이션 기울임
*/
func (p *HDrawFillAttr) GradationAngle(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationAngle", v)
}

/*
그러데이션 가로 중 X 좌표
*/
func (p *HDrawFillAttr) GradationCenterX(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationCenterX", v)
}

/*
그러데이션 가로 중 Y 좌표
*/
func (p *HDrawFillAttr) GradationCenterY(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationCenterY", v)
}

/*
그러데이션 번짐 정도
*/
func (p *HDrawFillAttr) GradationStep(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationStep", v)
}

/*
그러데이션 색 수
*/
func (p *HDrawFillAttr) GradationColorNum(v ...int32) int {
	return funcParaSetInt(p.variant, "GradationColorNum", v)
}

/*
PIT_ARRAY 생성

	GradationColor, GradationIndexPos을 위해 필요
*/
func (p *HDrawFillAttr) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
그러데이션 색깔(PIT_I)
*/
func (p *HDrawFillAttr) GradationColor() *HArray {
	return GetHArray(p.variant, "GradationColor")
}

/*
그러데이션 다음 색깔과의 거리(PIT_I)
*/
func (p *HDrawFillAttr) GradationIndexPos() *HArray {
	return GetHArray(p.variant, "GradationIndexPos")
}

/*
그러데이션 번짐 정도의 중심
*/
func (p *HDrawFillAttr) GradationStepCenter(v ...uint32) int {
	return funcParaSetInt(p.variant, "GradationStepCenter", v)
}

/*
브러쉬 색깔
*/
func (p *HDrawFillAttr) WinBrushFaceColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "WinBrushFaceColor", v)
}

/*
브러쉬 선 긋기 색깔
*/
func (p *HDrawFillAttr) WinBrushHatchColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "WinBrushHatchColor", v)
}

/*
브러쉬 선 긋기 스타일
*/
func (p *HDrawFillAttr) WinBrushFaceStyle(v ...int16) int {
	return funcParaSetInt(p.variant, "WinBrushFaceStyle", v)
}

/*
이미지 파일 경로
*/
func (p *HDrawFillAttr) FileName(v ...string) string {
	return funcParaSetString(p.variant, "FileName", v)
}

/*
문서에 삽입(TRUE) / 파일로 연결(FALSE)
*/
func (p *HDrawFillAttr) Embedded(v ...uint16) int {
	return funcParaSetInt(p.variant, "Embedded", v)
}

/*
그림 효과
*/
func (p *HDrawFillAttr) PicEffect(v ...uint16) int {
	return funcParaSetInt(p.variant, "PicEffect", v)
}

/*
명도 (-100 ~ 100)
*/
func (p *HDrawFillAttr) Brightness(v ...int16) int {
	return funcParaSetInt(p.variant, "Brightness", v)
}

/*
밝기 (-100 ~ 100)
*/
func (p *HDrawFillAttr) Contrast(v ...int16) int {
	return funcParaSetInt(p.variant, "Contrast", v)
}

/*
반전 유무
*/
func (p *HDrawFillAttr) Reverse(v ...uint16) int {
	return funcParaSetInt(p.variant, "Reverse", v)
}

/*
브러쉬로 쓰일 때 채우기 스타일 0 = 바둑판 식으로 1 = 가로/위만 바둑판식으로 배열 2 = 가로/아래만 바둑판식으로 배열 3 = 세로/왼쪽만 바둑판식으로 배열 4 = 세로/오른쪽만 바둑판식으로 배열 5 = 크기에 맞추어 6 = 가운데로 7 = 가운데 위로 8 = 가운데 아래로 9 = 왼쪽 가운데로 10 = 왼쪽 위로 11 = 왼쪽 아래로 12 = 오른쪽 가운데로 13 = 오른쪽 위로 14 = 오른쪽 아래로
*/
func (p *HDrawFillAttr) DrawFillImageType(v ...int32) int {
	return funcParaSetInt(p.variant, "DrawFillImageType", v)
}

/*
자르기 왼쪽
*/
func (p *HDrawFillAttr) SkipLeft(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipLeft", v)
}

/*
자르기 오른쪽
*/
func (p *HDrawFillAttr) SkipRight(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipRight", v)
}

/*
자르기 윗쪽
*/
func (p *HDrawFillAttr) SkipTop(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipTop", v)
}

/*
자르기 아랫쪽
*/
func (p *HDrawFillAttr) SkipBottom(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipBottom", v)
}

/*
이미지 원본 크기 X 사이즈
*/
func (p *HDrawFillAttr) OriginalSizeX(v ...uint32) int {
	return funcParaSetInt(p.variant, "OriginalSizeX", v)
}

/*
이미지 원본 크기 Y 사이즈
*/
func (p *HDrawFillAttr) OriginalSizeY(v ...uint32) int {
	return funcParaSetInt(p.variant, "OriginalSizeY", v)
}

/*
이미지 안쪽 여백. (오른쪽)
*/
func (p *HDrawFillAttr) InsideMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginRight", v)
}

/*
이미지 안쪽 여백. (위)
*/
func (p *HDrawFillAttr) InsideMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginTop", v)
}

/*
이미지 안쪽 여백. (아래)
*/
func (p *HDrawFillAttr) InsideMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginBottom", v)
}

/*
현재 선택된 brush의 type을 나타냄 (TRUE/FALSE)
*/
func (p *HDrawFillAttr) WindowsBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "WindowsBrush", v)
}

/*
현재 선택된 brush의 type이 그러데이션 브러시인가를 나타냄 (TRUE/FALSE)
*/
func (p *HDrawFillAttr) GradationBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "GradationBrush", v)
}

/*
현재 선택된 brush의 type이 이미지 브러시인가를 나타냄 (TRUE/FALSE)
*/
func (p *HDrawFillAttr) ImageBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "ImageBrush", v)
}

/*
HDrawImageAttr:

	HDrawImageAttr :  그리기 개체 그림
*/
type HDrawImageAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawImageAttr) HSet() *HSet {
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
이미지의 파일 경로
*/
func (p *HDrawImageAttr) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
문서에 삽입(TRUE) / 파일로 연결(FALSE)
*/
func (p *HDrawImageAttr) Embedded(v ...uint16) int {
	return funcParaSetInt(p.variant, "Embedded", v)
}

/*
그림 효과
*/
func (p *HDrawImageAttr) PicEffect(v ...uint16) int {
	return funcParaSetInt(p.variant, "PicEffect", v)
}

/*
명도 (-100 ~ 100)
*/
func (p *HDrawImageAttr) Brightness(v ...int16) int {
	return funcParaSetInt(p.variant, "Brightness", v)
}

/*
밝기 (-100 ~ 100)
*/
func (p *HDrawImageAttr) Contrast(v ...int16) int {
	return funcParaSetInt(p.variant, "Contrast", v)
}

/*
반전 유무
*/
func (p *HDrawImageAttr) Reverse(v ...uint16) int {
	return funcParaSetInt(p.variant, "Reverse", v)
}

/*
브러쉬로 쓰일 때 채우기 스타일 0 = 바둑판 식으로 1 = 가로/위만 바둑판식으로 배열 2 = 가로/아래만 바둑판식으로 배열 3 = 세로/왼쪽만 바둑판식으로 배열 4 = 세로/오른쪽만 바둑판식으로 배열 5 = 크기에 맞추어 6 = 가운데로 7 = 가운데 위로 8 = 가운데 아래로 9 = 왼쪽 가운데로 10 = 왼쪽 위로 11 = 왼쪽 아래로 12 = 오른쪽 가운데로 13 = 오른쪽 위로 14 = 오른쪽 아래로
*/
func (p *HDrawImageAttr) DrawFillImageType(v ...int32) int {
	return funcParaSetInt(p.variant, "DrawFillImageType", v)
}

/*
자르기 왼쪽
*/
func (p *HDrawImageAttr) SkipLeft(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipLeft", v)
}

/*
자르기 오른쪽
*/
func (p *HDrawImageAttr) SkipRight(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipRight", v)
}

/*
자르기 윗쪽
*/
func (p *HDrawImageAttr) SkipTop(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipTop", v)
}

/*
자르기 아랫쪽
*/
func (p *HDrawImageAttr) SkipBottom(v ...uint32) int {
	return funcParaSetInt(p.variant, "SkipBottom", v)
}

/*
이미지 원본 크기 X 사이즈
*/
func (p *HDrawImageAttr) OriginalSizeX(v ...uint32) int {
	return funcParaSetInt(p.variant, "OriginalSizeX", v)
}

/*
이미지 원본 크기 Y 사이즈
*/
func (p *HDrawImageAttr) OriginalSizeY(v ...uint32) int {
	return funcParaSetInt(p.variant, "OriginalSizeY", v)
}

/*
이미지 안쪽 여백. (왼쪽)
*/
func (p *HDrawImageAttr) InsideMarginLeft(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginLeft", v)
}

/*
이미지 안쪽 여백. (오른쪽)
*/
func (p *HDrawImageAttr) InsideMarginRight(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginRight", v)
}

/*
이미지 안쪽 여백. (위)
*/
func (p *HDrawImageAttr) InsideMarginTop(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginTop", v)
}

/*
이미지 안쪽 여백. (아래)
*/
func (p *HDrawImageAttr) InsideMarginBottom(v ...int32) int {
	return funcParaSetInt(p.variant, "InsideMarginBottom", v)
}

/*
현재 선택된 brush의 type을 나타냄 (TRUE/FALSE)
*/
func (p *HDrawImageAttr) WindowsBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "WindowsBrush", v)
}

/*
현재 선택된 brush의 type이 그러데이션 브러시인가를 나타냄 (TRUE/FALSE)
*/
func (p *HDrawImageAttr) GradationBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "GradationBrush", v)
}

/*
현재 선택된 brush의 type이 이미지 브러시인가를 나타냄 (TRUE/FALSE)
*/
func (p *HDrawImageAttr) ImageBrush(v ...uint16) int {
	return funcParaSetInt(p.variant, "ImageBrush", v)
}

/*
HDrawRectType:

	HDrawRectType :  그리기 개체 사각형 속성
*/
type HDrawRectType struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawRectType) HSet() *HSet {
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
모서리 모양 (둥근모양/각진모양)
*/
func (p *HDrawRectType) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HDrawArcType:

	HDrawArcType :  그리기 개체 호
*/
type HDrawArcType struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawArcType) HSet() *HSet {
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
모서리 모양(둥근 모양/각진 모양)
*/
func (p *HDrawArcType) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
PIT_ARRAY 생성

	Interval을 위해 필요
*/
func (p *HDrawArcType) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
확장타원(호)에서 호의 시작점과 끝점을 나타내게 한다.
*/
func (p *HDrawArcType) Interval() *HArray {
	return GetHArray(p.variant, "Interval")
}

/*
HDrawResize:

	HDrawResize :  그리기 개체 늘이기
*/
type HDrawResize struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawResize) HSet() *HSet {
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
그리기 개체 X 좌표 위치
*/
func (p *HDrawResize) Xoffset(v ...int32) int {
	return funcParaSetInt(p.variant, "Xoffset", v)
}

/*
그리기 개체 Y 좌표 위치
*/
func (p *HDrawResize) Yoffset(v ...int32) int {
	return funcParaSetInt(p.variant, "Yoffset", v)
}

/*
Reserved
*/
func (p *HDrawResize) HandleIndex(v ...uint32) int {
	return funcParaSetInt(p.variant, "HandleIndex", v)
}

/*
Reserved
*/
func (p *HDrawResize) Mode(v ...uint32) int {
	return funcParaSetInt(p.variant, "Mode", v)
}

/*
HDrawRotate:

	HDrawRotate :  그리기 개체의 회전
*/
type HDrawRotate struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawRotate) HSet() *HSet {
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
그리기 개체 회전 타입
*/
func (p *HDrawRotate) Command(v ...uint32) int {
	return funcParaSetInt(p.variant, "Command", v)
}

/*
회전 중심의 X 좌표
*/
func (p *HDrawRotate) CenterX(v ...int32) int {
	return funcParaSetInt(p.variant, "CenterX", v)
}

/*
회전 중심의 y 좌표
*/
func (p *HDrawRotate) CenterY(v ...int32) int {
	return funcParaSetInt(p.variant, "CenterY", v)
}

/*
그리기 개체의 중심 X 좌표
*/
func (p *HDrawRotate) ObjectCenterX(v ...int32) int {
	return funcParaSetInt(p.variant, "ObjectCenterX", v)
}

/*
그리기 개체의 중심 y 좌표
*/
func (p *HDrawRotate) ObjectCenterY(v ...int32) int {
	return funcParaSetInt(p.variant, "ObjectCenterY", v)
}

/*
회전 각
*/
func (p *HDrawRotate) Angle(v ...int32) int {
	return funcParaSetInt(p.variant, "Angle", v)
}

/*
HDrawEditDetail:

	HDrawEditDetail :  그리기 개체 점 편집
*/
type HDrawEditDetail struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawEditDetail) HSet() *HSet {
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
Reserved
*/
func (p *HDrawEditDetail) Command(v ...uint32) int {
	return funcParaSetInt(p.variant, "Command", v)
}

/*
고칠 점의 인덱스
*/
func (p *HDrawEditDetail) Index(v ...uint32) int {
	return funcParaSetInt(p.variant, "Index", v)
}

/*
점의 X 좌표
*/
func (p *HDrawEditDetail) PointX(v ...int32) int {
	return funcParaSetInt(p.variant, "PointX", v)
}

/*
점의 y 좌표
*/
func (p *HDrawEditDetail) PointY(v ...int32) int {
	return funcParaSetInt(p.variant, "PointY", v)
}

/*
HDrawImageScissoring:

	HDrawImageScissoring :  그리기 개체 그림 자르기
*/
type HDrawImageScissoring struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawImageScissoring) HSet() *HSet {
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
그리기 개체 X좌표 위치
*/
func (p *HDrawImageScissoring) Xoffset(v ...int32) int {
	return funcParaSetInt(p.variant, "Xoffset", v)
}

/*
그리기 개체 Y좌표 위치
*/
func (p *HDrawImageScissoring) Yoffset(v ...int32) int {
	return funcParaSetInt(p.variant, "Yoffset", v)
}

/*
Reserved
*/
func (p *HDrawImageScissoring) HandleIndex(v ...uint32) int {
	return funcParaSetInt(p.variant, "HandleIndex", v)
}

/*
HDrawScAction:

	HDrawScAction :  그리기 개체 뒤집기
*/
type HDrawScAction struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawScAction) HSet() *HSet {
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
회전 중심 X 좌표
*/
func (p *HDrawScAction) RotateCenterX(v ...int32) int {
	return funcParaSetInt(p.variant, "RotateCenterX", v)
}

/*
회전 중심 Y 좌표
*/
func (p *HDrawScAction) RotateCenterY(v ...int32) int {
	return funcParaSetInt(p.variant, "RotateCenterY", v)
}

/*
회전각
*/
func (p *HDrawScAction) RotateAngel(v ...int32) int {
	return funcParaSetInt(p.variant, "RotateAngel", v)
}

/*
수평 flip
*/
func (p *HDrawScAction) HorzFlip(v ...uint32) int {
	return funcParaSetInt(p.variant, "HorzFlip", v)
}

/*
수직 flip
*/
func (p *HDrawScAction) VertFlip(v ...uint32) int {
	return funcParaSetInt(p.variant, "VertFlip", v)
}

/*
HDrawCtrlHyperlink:

	HDrawCtrlHyperlink :  그리기 개체 하이퍼링크
*/
type HDrawCtrlHyperlink struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawCtrlHyperlink) HSet() *HSet {
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
하이퍼링크 경로
*/
func (p *HDrawCtrlHyperlink) Command(v ...string) string {
	return funcParaSetString(p.variant, "Command", v)
}

/*
HDrawCoordInfo:

	HDrawCoordInfo :  그리기 개체 좌표 정보
*/
type HDrawCoordInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawCoordInfo) HSet() *HSet {
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
점의 개수
*/
func (p *HDrawCoordInfo) Count(v ...uint32) int {
	return funcParaSetInt(p.variant, "Count", v)
}

/*
PIT_ARRAY 생성

	Point, Line을 위해 필요
*/
func (p *HDrawCoordInfo) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
좌표 array (x1, y1), (x2, y2),... (xn, yn) 각각의 좌표는 PIT_I4 type 이다.
*/
func (p *HDrawCoordInfo) Point() *HArray {
	return GetHArray(p.variant, "Point")
}

/*
선 속성 array(곡선에서만 쓰임) 각각의 선 속성은 PIT_UI1 type 이다.
*/
func (p *HDrawCoordInfo) Line() *HArray {
	return GetHArray(p.variant, "Line")
}

/*
HDrawSoMouseOption:

	HDrawSoMouseOption :  쉬운 개체 선택, 마우스 끌기로 만들기 정보
*/
type HDrawSoMouseOption struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawSoMouseOption) HSet() *HSet {
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
쉬운 개체 선택(부분 선택이 되어도 개체가 선택되는 모드) TRUE/FALSE
*/
func (p *HDrawSoMouseOption) EasySelect(v ...uint16) int {
	return funcParaSetInt(p.variant, "EasySelect", v)
}

/*
마우스 끌기로 틀 만들기 TRUE/FALSE
*/
func (p *HDrawSoMouseOption) CreateOnDrag(v ...uint16) int {
	return funcParaSetInt(p.variant, "CreateOnDrag", v)
}

/*
HDrawSoEquationOption:

	HDrawSoEquationOption :  환경설정의 틀 속성
*/
type HDrawSoEquationOption struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawSoEquationOption) HSet() *HSet {
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
한글 97식의 수식 만들기를 기본으로 지원할지를 나타냄 TRUE/FALSE
*/
func (p *HDrawSoEquationOption) ScriptDefault(v ...uint16) int {
	return funcParaSetInt(p.variant, "ScriptDefault", v)
}

/*
문서를 읽을 때 자동으로 이전 판 수식을 워디안 수식으로 바꿀지를 나타냄 TRUE/FALSE
*/
func (p *HDrawSoEquationOption) AutoConvert(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoConvert", v)
}

/*
HDrawShear:

	HDrawShear :  그리기 개체 기울기
*/
type HDrawShear struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawShear) HSet() *HSet {
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
XFactor
*/
func (p *HDrawShear) XFactor(v ...int32) int {
	return funcParaSetInt(p.variant, "XFactor", v)
}

/*
YFactor
*/
func (p *HDrawShear) YFactor(v ...int32) int {
	return funcParaSetInt(p.variant, "YFactor", v)
}

/*
HDrawTextart:

	HDrawTextart :  글맵시
*/
type HDrawTextart struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDrawTextart) HSet() *HSet {
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
글맵시에 넣을 문자열 내용
*/
func (p *HDrawTextart) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
글꼴
*/
func (p *HDrawTextart) FontName(v ...string) string {
	return funcParaSetString(p.variant, "FontName", v)
}

/*
글꼴 스타일
*/
func (p *HDrawTextart) FontStyle(v ...string) string {
	return funcParaSetString(p.variant, "FontStyle", v)
}

/*
글꼴 타입
*/
func (p *HDrawTextart) FontType(v ...uint16) int {
	return funcParaSetInt(p.variant, "FontType", v)
}

/*
줄간격 (50 ~ 500) : 20
*/
func (p *HDrawTextart) LineSpacing(v ...int32) int {
	return funcParaSetInt(p.variant, "LineSpacing", v)
}

/*
글자간격 (50 ~ 500) : 100
*/
func (p *HDrawTextart) CharSpacing(v ...int32) int {
	return funcParaSetInt(p.variant, "CharSpacing", v)
}

/*
정렬 방식
*/
func (p *HDrawTextart) AlignType(v ...uint16) int {
	return funcParaSetInt(p.variant, "AlignType", v)
}

/*
형태 (0 ~ 54) : 0
*/
func (p *HDrawTextart) Shape(v ...uint16) int {
	return funcParaSetInt(p.variant, "Shape", v)
}

/*
그림자 속성
*/
func (p *HDrawTextart) ShadowType(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShadowType", v)
}

/*
그림자 간격 X 방향 (-48% ~ 48%) 10
*/
func (p *HDrawTextart) ShadowOffsetX(v ...int16) int {
	return funcParaSetInt(p.variant, "ShadowOffsetX", v)
}

/*
그림자 간격 y 방향 (-48% ~ 48%) 10
*/
func (p *HDrawTextart) ShadowOffsetY(v ...int16) int {
	return funcParaSetInt(p.variant, "ShadowOffsetY", v)
}

/*
음영색 (COLORREF) RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HDrawTextart) ShadowColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "ShadowColor", v)
}

/*
HFormGeneral:

	HFormGeneral :  양식 개체 타입 정보
*/
type HFormGeneral struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormGeneral) HSet() *HSet {
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
양식 개체 타입
*/
func (p *HFormGeneral) FormObjType(v ...uint32) int {
	return funcParaSetInt(p.variant, "FormObjType", v)
}

/*
HFormCommonAttr:

	HFormCommonAttr :  양식 개체 공통 속성
*/
type HFormCommonAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormCommonAttr) HSet() *HSet {
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
이름 (ID)
*/
func (p *HFormCommonAttr) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

/*
전경색
*/
func (p *HFormCommonAttr) Color(v ...uint32) int {
	return funcParaSetInt(p.variant, "Color", v)
}

/*
배경색
*/
func (p *HFormCommonAttr) BackColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "BackColor", v)
}

/*
그룹 이름
*/
func (p *HFormCommonAttr) GroupName(v ...string) string {
	return funcParaSetString(p.variant, "GroupName", v)
}

/*
Tab Stop
*/
func (p *HFormCommonAttr) TabStop(v ...uint16) int {
	return funcParaSetInt(p.variant, "TabStop", v)
}

/*
탭 순서
*/
func (p *HFormCommonAttr) TabOrder(v ...int32) int {
	return funcParaSetInt(p.variant, "TabOrder", v)
}

/*
활성화/비활성화 옵션
*/
func (p *HFormCommonAttr) Enabled(v ...uint16) int {
	return funcParaSetInt(p.variant, "Enabled", v)
}

/*
테두리 타입
*/
func (p *HFormCommonAttr) BorderType(v ...uint32) int {
	return funcParaSetInt(p.variant, "BorderType", v)
}

/*
틀을 그릴지 여부
*/
func (p *HFormCommonAttr) DrawFrame(v ...uint16) int {
	return funcParaSetInt(p.variant, "DrawFrame", v)
}

/*
Command String
*/
func (p *HFormCommonAttr) Command(v ...string) string {
	return funcParaSetString(p.variant, "Command", v)
}

/*
편집 가능 여부
*/
func (p *HFormCommonAttr) Editable(v ...uint16) int {
	return funcParaSetInt(p.variant, "Editable", v)
}

/*
양식 개체 인쇄 가능 여부
*/
func (p *HFormCommonAttr) Printable(v ...uint16) int {
	return funcParaSetInt(p.variant, "Printable", v)
}

/*
HFormCharshapeattr:

	HFormCharshapeattr :  양식 개체 글자모양 속성
*/
type HFormCharshapeattr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormCharshapeattr) HSet() *HSet {
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
글자 속성 속성 (HCharShape)
*/
func (p *HFormCharshapeattr) CharShape(v ...*HCharShape) *HCharShape {
	var para HCharShape
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "CharShape", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CharShape")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
주위의 글자 속성을 따를지 여부
*/
func (p *HFormCharshapeattr) FollowContext(v ...uint16) int {
	return funcParaSetInt(p.variant, "FollowContext", v)
}

/*
글자 크기에 맞게 개체 크기가 바뀜
*/
func (p *HFormCharshapeattr) AutoSize(v ...uint16) int {
	return funcParaSetInt(p.variant, "AutoSize", v)
}

/*
자동 줄 바꿈
*/
func (p *HFormCharshapeattr) WordWarp(v ...uint16) int {
	return funcParaSetInt(p.variant, "WordWarp", v)
}

/*
HFormButtonAttr:

	HFormButtonAttr :  양식 개체 버튼
*/
type HFormButtonAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormButtonAttr) HSet() *HSet {
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
캡션
*/
func (p *HFormButtonAttr) Caption(v ...string) string {
	return funcParaSetString(p.variant, "Caption", v)
}

/*
체크 상태 값
*/
func (p *HFormButtonAttr) Value(v ...uint16) int {
	return funcParaSetInt(p.variant, "Value", v)
}

/*
Radio 버튼 그룹 이름
*/
func (p *HFormButtonAttr) RadioGroupName(v ...string) string {
	return funcParaSetString(p.variant, "RadioGroupName", v)
}

/*
체크 상태 옵션
*/
func (p *HFormButtonAttr) TriState(v ...uint16) int {
	return funcParaSetInt(p.variant, "TriState", v)
}

/*
배경 투명도
*/
func (p *HFormButtonAttr) BackStyle(v ...uint16) int {
	return funcParaSetInt(p.variant, "BackStyle", v)
}

/*
HFormComboboxAttr:

	HFormComboboxAttr :  양식 개체 콤보 박스
*/
type HFormComboboxAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormComboboxAttr) HSet() *HSet {
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
DropDown ListBox 가 펼쳐졌을 때 최대 줄 수
*/
func (p *HFormComboboxAttr) ListBoxRows(v ...uint32) int {
	return funcParaSetInt(p.variant, "ListBoxRows", v)
}

/*
DropDown ListBox 가 펼쳐졌을 때 폭
*/
func (p *HFormComboboxAttr) ListBoxWidth(v ...uint32) int {
	return funcParaSetInt(p.variant, "ListBoxWidth", v)
}

/*
콤보의 에디트 텍스트
*/
func (p *HFormComboboxAttr) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
콤보의 에디트 입력 가능 여부
*/
func (p *HFormComboboxAttr) EditEnable(v ...uint16) int {
	return funcParaSetInt(p.variant, "EditEnable", v)
}

/*
HFormEditAttr:

	HFormEditAttr :  양식 개체 에디트 상자
*/
type HFormEditAttr struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormEditAttr) HSet() *HSet {
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
에디트 텍스트
*/
func (p *HFormEditAttr) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
MultiLine
*/
func (p *HFormEditAttr) MultiLine(v ...uint16) int {
	return funcParaSetInt(p.variant, "MultiLine", v)
}

/*
패스워드 표시에 사용할 글자
*/
func (p *HFormEditAttr) PasswordChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "PasswordChar", v)
}

/*
입력 가능한 최대 글자 수
*/
func (p *HFormEditAttr) MaxLength(v ...uint32) int {
	return funcParaSetInt(p.variant, "MaxLength", v)
}

/*
스크롤바 표시
*/
func (p *HFormEditAttr) ScrollBars(v ...uint16) int {
	return funcParaSetInt(p.variant, "ScrollBars", v)
}

/*
탭 키를 눌렀을 때 반응
*/
func (p *HFormEditAttr) TabKeyBehavior(v ...uint16) int {
	return funcParaSetInt(p.variant, "TabKeyBehavior", v)
}

/*
숫자만 입력 가능
*/
func (p *HFormEditAttr) Number(v ...uint16) int {
	return funcParaSetInt(p.variant, "Number", v)
}

/*
읽기만 가능
*/
func (p *HFormEditAttr) ReadOnly(v ...uint16) int {
	return funcParaSetInt(p.variant, "ReadOnly", v)
}

/*
HStyle:

	HStyle :  스타일
*/
type HStyle struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HStyle) HSet() *HSet {
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
적용할 스타일 인덱스
*/
func (p *HStyle) Apply(v ...int32) int {
	return funcParaSetInt(p.variant, "Apply", v)
}

/*
HFileOpen:

	HFileOpen :  불러오기
*/
type HFileOpen struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileOpen) HSet() *HSet {
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
파일 이름
*/
func (p *HFileOpen) OpenFileName(v ...string) string {
	return funcParaSetString(p.variant, "OpenFileName", v)
}

/*
파일 이름, Open과 같은 값을 쓴다.
*/
func (p *HFileOpen) SaveFileName(v ...string) string {
	return funcParaSetString(p.variant, "SaveFileName", v)
}

/*
format
*/
func (p *HFileOpen) OpenFormat(v ...string) string {
	return funcParaSetString(p.variant, "OpenFormat", v)
}

/*
format, Open과 같은 값을 쓴다.
*/
func (p *HFileOpen) SaveFormat(v ...string) string {
	return funcParaSetString(p.variant, "SaveFormat", v)
}

/*
읽기 전용
*/
func (p *HFileOpen) OpenReadOnly(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenReadOnly", v)
}

/*
덮어 쓰기
*/
func (p *HFileOpen) SaveOverWrite(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveOverWrite", v)
}

/*
문서를 열때 추가로 필요한 옵션
*/
func (p *HFileOpen) OpenFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenFlag", v)
}

/*
문서 내용 변경 상태
*/
func (p *HFileOpen) ModifiedFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "ModifiedFlag", v)
}

/*
해당 format 파일을 Open 및 Save할 경우 사용되는 옵션 (Open API 참고)
*/
func (p *HFileOpen) Argument(v ...string) string {
	return funcParaSetString(p.variant, "Argument", v)
}

/*
97 호환 저장
*/
func (p *HFileOpen) SaveCMFDoc30(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveCMFDoc30", v)
}

/*
97 문서로 저장
*/
func (p *HFileOpen) SaveHwp97(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveHwp97", v)
}

/*
배포용 문서로 저장
*/
func (p *HFileOpen) SaveDistribute(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveDistribute", v)
}

/*
HFileSaveAs:

	HFileSaveAs :  새 이름으로 저장하기
*/
type HFileSaveAs struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileSaveAs) HSet() *HSet {
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
파일 이름
*/
func (p *HFileSaveAs) OpenFileName(v ...string) string {
	return funcParaSetString(p.variant, "OpenFileName", v)
}

/*
파일 이름, Open과 같은 값을 쓴다.
*/
func (p *HFileSaveAs) SaveFileName(v ...string) string {
	return funcParaSetString(p.variant, "SaveFileName", v)
}

/*
format
*/
func (p *HFileSaveAs) OpenFormat(v ...string) string {
	return funcParaSetString(p.variant, "OpenFormat", v)
}

/*
format, Open과 같은 값을 쓴다.
*/
func (p *HFileSaveAs) SaveFormat(v ...string) string {
	return funcParaSetString(p.variant, "SaveFormat", v)
}

/*
읽기 전용
*/
func (p *HFileSaveAs) OpenReadOnly(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenReadOnly", v)
}

/*
덮어 쓰기
*/
func (p *HFileSaveAs) SaveOverWrite(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveOverWrite", v)
}

/*
문서를 저장할때 추가로 필요한 옵션
*/
func (p *HFileSaveAs) OpenFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenFlag", v)
}

/*
문서 내용 변경 상태
*/
func (p *HFileSaveAs) ModifiedFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "ModifiedFlag", v)
}

/*
해당 format 파일을 Open 및 Save할 경우 사용되는 옵션 (Open API 참고)
*/
func (p *HFileSaveAs) Argument(v ...string) string {
	return funcParaSetString(p.variant, "Argument", v)
}

/*
97 호환 저장
*/
func (p *HFileSaveAs) SaveCMFDoc30(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveCMFDoc30", v)
}

/*
97 문서로 저장
*/
func (p *HFileSaveAs) SaveHwp97(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveHwp97", v)
}

/*
배포용 문서로 저장
*/
func (p *HFileSaveAs) SaveDistribute(v ...uint16) int {
	return funcParaSetInt(p.variant, "SaveDistribute", v)
}

/*
HDropCap:

	HDropCap :  문단 첫 글자 장식
*/
type HDropCap struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDropCap) HSet() *HSet {
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
스타일 0=없음 1=2줄 차지 2=3줄 차지 4=여백
*/
func (p *HDropCap) Style(v ...uint32) int {
	return funcParaSetInt(p.variant, "Style", v)
}

/*
글꼴
*/
func (p *HDropCap) FaceName(v ...string) string {
	return funcParaSetString(p.variant, "FaceName", v)
}

/*
선 스타일
*/
func (p *HDropCap) LineStyle(v ...int32) int {
	return funcParaSetInt(p.variant, "LineStyle", v)
}

/*
선 굵기
*/
func (p *HDropCap) LineWeight(v ...uint32) int {
	return funcParaSetInt(p.variant, "LineWeight", v)
}

/*
선 색 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HDropCap) LineColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "LineColor", v)
}

/*
글꼴 색 RGB color를 나타내기 위한 32비트 값 (0x00bbggrr)
*/
func (p *HDropCap) FaceColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "FaceColor", v)
}

/*
본문과의 간격
*/
func (p *HDropCap) Spacing(v ...int32) int {
	return funcParaSetInt(p.variant, "Spacing", v)
}

/*
HKeyMacro:

	HKeyMacro :  키 매크로
*/
type HKeyMacro struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HKeyMacro) HSet() *HSet {
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
정의(or 실행)할 매크로의 인덱스
*/
func (p *HKeyMacro) Index(v ...int32) int {
	return funcParaSetInt(p.variant, "Index", v)
}

/*
실행 반복 횟수
*/
func (p *HKeyMacro) RepeatCount(v ...int32) int {
	return funcParaSetInt(p.variant, "RepeatCount", v)
}

/*
매크로 이름
*/
func (p *HKeyMacro) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

/*
HShapeCopyPaste:

	HShapeCopyPaste :  모양 복사
*/
type HShapeCopyPaste struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HShapeCopyPaste) HSet() *HSet {
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
모양 복사 종류 0 = 글자 모양 1 = 문단 모양
*/
func (p *HShapeCopyPaste) Type(v ...int32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
셀 속성 복사
*/
func (p *HShapeCopyPaste) CellAttr(v ...uint16) int {
	return funcParaSetInt(p.variant, "CellAttr", v)
}

/*
셀 테두리 복사
*/
func (p *HShapeCopyPaste) CellBorder(v ...uint16) int {
	return funcParaSetInt(p.variant, "CellBorder", v)
}

/*
셀 배경 복사
*/
func (p *HShapeCopyPaste) CellFill(v ...uint16) int {
	return funcParaSetInt(p.variant, "CellFill", v)
}

/*
본문과 셀 모양 둘 다 복사/셀 모양만 복사
*/
func (p *HShapeCopyPaste) TypeBodyAndCellOnly(v ...int32) int {
	return funcParaSetInt(p.variant, "TypeBodyAndCellOnly", v)
}

/*
HStyleTemplate:

	HStyleTemplate :  스타일마당
*/
type HStyleTemplate struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HStyleTemplate) HSet() *HSet {
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
파일 이름
*/
func (p *HStyleTemplate) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
스타일 이름
*/
func (p *HStyleTemplate) NameLocals(v ...string) string {
	return funcParaSetString(p.variant, "NameLocals", v)
}

/*
스타일 이름(영문)
*/
func (p *HStyleTemplate) NameEngs(v ...string) string {
	return funcParaSetString(p.variant, "NameEngs", v)
}

/*
HMasterPage:

	HMasterPage :  바탕쪽
*/
type HMasterPage struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMasterPage) HSet() *HSet {
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
바탕쪽 종류
*/
func (p *HMasterPage) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
기존 바탕쪽과 겹침 (On/Off)
*/
func (p *HMasterPage) Duplicate(v ...uint16) int {
	return funcParaSetInt(p.variant, "Duplicate", v)
}

/*
바탕쪽과 앞으로 보내기 (On/Off)
*/
func (p *HMasterPage) Front(v ...uint16) int {
	return funcParaSetInt(p.variant, "Front", v)
}

/*
적용 대상
*/
func (p *HMasterPage) ApplyTo(v ...uint16) int {
	return funcParaSetInt(p.variant, "ApplyTo", v)
}

/*
HGotoE:

	HGotoE :  찾아가기
*/
type HGotoE struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HGotoE) HSet() *HSet {
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
현재 선택되어 있는 라디오 값
*/
func (p *HGotoE) SetSelectionIndex(v ...uint16) int {
	return funcParaSetInt(p.variant, "SetSelectionIndex", v)
}

/*
HInsertText:

	HInsertText :  텍스트 삽입
*/
type HInsertText struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HInsertText) HSet() *HSet {
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
삽입할 텍스트
*/
func (p *HInsertText) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
HBookMark:

	HBookMark :  책갈피
*/
type HBookMark struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HBookMark) HSet() *HSet {
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
책갈피 이름
*/
func (p *HBookMark) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

/*
책갈피 종류
*/
func (p *HBookMark) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
책갈피 명령의 종류
*/
func (p *HBookMark) Command(v ...uint32) int {
	return funcParaSetInt(p.variant, "Command", v)
}

/*
HHyperLink:

	HHyperLink :  하이퍼링크 삽입 / 고치기
*/
type HHyperLink struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HHyperLink) HSet() *HSet {
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
하이퍼링크가 표시되는 문자열
*/
func (p *HHyperLink) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
하이퍼링크 경로
*/
func (p *HHyperLink) Command(v ...string) string {
	return funcParaSetString(p.variant, "Command", v)
}

/*
"연결 안함" 여부
*/
func (p *HHyperLink) NoLInk(v ...int32) int {
	return funcParaSetInt(p.variant, "NoLInk", v)
}

/*
그림 및 그리기 객체가 selection되어 있는지 여부
*/
func (p *HHyperLink) ShapeObject(v ...int32) int {
	return funcParaSetInt(p.variant, "ShapeObject", v)
}

/*
HInsertFieldTemplate:

	HInsertFieldTemplate :  문서 마당 정보
*/
type HInsertFieldTemplate struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HInsertFieldTemplate) HSet() *HSet {
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
문서 마당 정보 PropertySheet 대화상자에서 원하는 페이지(탭)만 보이기
*/
func (p *HInsertFieldTemplate) ShowSingle(v ...uint16) int {
	return funcParaSetInt(p.variant, "ShowSingle", v)
}

/*
필드 컨트롤의 안내문/지시문
*/
func (p *HInsertFieldTemplate) TemplateDirection(v ...string) string {
	return funcParaSetString(p.variant, "TemplateDirection", v)
}

/*
필드 컨트롤의 도움말
*/
func (p *HInsertFieldTemplate) TemplateHelp(v ...string) string {
	return funcParaSetString(p.variant, "TemplateHelp", v)
}

/*
필드 이름 (name)
*/
func (p *HInsertFieldTemplate) TemplateName(v ...string) string {
	return funcParaSetString(p.variant, "TemplateName", v)
}

/*
필드의 종류

	0 = 누름틀
	1 = 개인 정보
	2 = 문서 요약
	3 = 문서를 만든 날짜
	4 = 파일 이름/경로
*/
func (p *HInsertFieldTemplate) TemplateType(v ...uint16) int {
	return funcParaSetInt(p.variant, "TemplateType", v)
}

/*
부분편집 모드에서 편집속성
*/
func (p *HInsertFieldTemplate) Editable(v ...uint16) int {
	return funcParaSetInt(p.variant, "Editable", v)
}

/*
HIdiom:

	HIdiom :  상용구
*/
type HIdiom struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HIdiom) HSet() *HSet {
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
삽입될 스트링/끼워 넣을 파일
*/
func (p *HIdiom) InputText(v ...string) string {
	return funcParaSetString(p.variant, "InputText", v)
}

/*
입력기 상용구/한글 상용구
*/
func (p *HIdiom) InputType(v ...uint16) int {
	return funcParaSetInt(p.variant, "InputType", v)
}

/*
HCodeTable:

	HCodeTable :  문자표
*/
type HCodeTable struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HCodeTable) HSet() *HSet {
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
삽입될 스트링
*/
func (p *HCodeTable) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
HConvertCase:

	HConvertCase :  대소문자 변환
*/
type HConvertCase struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HConvertCase) HSet() *HSet {
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
변환 종류
*/
func (p *HConvertCase) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HConvertToHangul:

	HConvertToHangul :  한자,일어,구결을 한글로
*/
type HConvertToHangul struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HConvertToHangul) HSet() *HSet {
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
변환 종류
*/
func (p *HConvertToHangul) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
한자
*/
func (p *HConvertToHangul) Hanja(v ...uint32) int {
	return funcParaSetInt(p.variant, "Hanja", v)
}

/*
히라카나
*/
func (p *HConvertToHangul) Hira(v ...uint32) int {
	return funcParaSetInt(p.variant, "Hira", v)
}

/*
가타가나
*/
func (p *HConvertToHangul) Gata(v ...uint32) int {
	return funcParaSetInt(p.variant, "Gata", v)
}

/*
구결
*/
func (p *HConvertToHangul) Gu(v ...uint32) int {
	return funcParaSetInt(p.variant, "Gu", v)
}

/*
HConvertHiraToGata:

	HConvertHiraToGata :  히라가나/가타가나 변환
*/
type HConvertHiraToGata struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HConvertHiraToGata) HSet() *HSet {
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
변환 종류
*/
func (p *HConvertHiraToGata) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HConvertFullHalf:

	HConvertFullHalf :  전/반각 변환
*/
type HConvertFullHalf struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HConvertFullHalf) HSet() *HSet {
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
변환 종류
*/
func (p *HConvertFullHalf) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
숫자
*/
func (p *HConvertFullHalf) Number(v ...uint32) int {
	return funcParaSetInt(p.variant, "Number", v)
}

/*
알파벳
*/
func (p *HConvertFullHalf) Alpha(v ...uint32) int {
	return funcParaSetInt(p.variant, "Alpha", v)
}

/*
심볼
*/
func (p *HConvertFullHalf) Symbol(v ...uint32) int {
	return funcParaSetInt(p.variant, "Symbol", v)
}

/*
가타가나
*/
func (p *HConvertFullHalf) Gata(v ...uint32) int {
	return funcParaSetInt(p.variant, "Gata", v)
}

/*
한글 자모
*/
func (p *HConvertFullHalf) HGJamo(v ...uint32) int {
	return funcParaSetInt(p.variant, "HGJamo", v)
}

/*
HInsertFile:

	HInsertFile :  끼워 넣기(파일 삽입)
*/
type HInsertFile struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HInsertFile) HSet() *HSet {
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
삽입할 파일 이름
*/
func (p *HInsertFile) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
삽입할 파일의 확장자
*/
func (p *HInsertFile) FileFormat(v ...string) string {
	return funcParaSetString(p.variant, "FileFormat", v)
}

/*
삽입할 파일의 Argument
*/
func (p *HInsertFile) FileArg(v ...string) string {
	return funcParaSetString(p.variant, "FileArg", v)
}

/*
HDocFilters:

	HDocFilters :  도큐먼트 필터 리스트
*/
type HDocFilters struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDocFilters) HSet() *HSet {
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
'|'문자로 분리된 필터 리스트 스트링(MFC 와 같은 방법)을 가져옴
*/
func (p *HDocFilters) DocFilters(v ...string) string {
	return funcParaSetString(p.variant, "DocFilters", v)
}

/*
PIT_ARRAY 생성

	Format을 위해 필요
*/
func (p *HDocFilters) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
필터 리스트를 string array 형태로 가져옴 (Export시에만)
*/
func (p *HDocFilters) Format() *HArray {
	return GetHArray(p.variant, "Format")
}

/*
Import 리스트를 가져올 것인지 Export 리스트를 가져올 것인지의 관한 타입.

	Import = TRUE, Export = FALSE (on input)
*/
func (p *HDocFilters) Type(v ...uint16) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HQCorrect:

	HQCorrect :  빠른 교정
*/
type HQCorrect struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HQCorrect) HSet() *HSet {
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
빠른 교정을 실행한 키 정보
*/
func (p *HQCorrect) LauncherKey(v ...uint16) int {
	return funcParaSetInt(p.variant, "LauncherKey", v)
}

/*
URL 또는 email 하이퍼링크 작성 키 정보
*/
func (p *HQCorrect) HyperLinkRunKey(v ...uint16) int {
	return funcParaSetInt(p.variant, "HyperLinkRunKey", v)
}

/*
HMakeContents:

	HMakeContents :  차례 만들기
*/
type HMakeContents struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMakeContents) HSet() *HSet {
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
만들 차례
*/
func (p *HMakeContents) Make(v ...uint32) int {
	return funcParaSetInt(p.variant, "Make", v)
}

/*
개요 수준
*/
func (p *HMakeContents) Level(v ...int32) int {
	return funcParaSetInt(p.variant, "Level", v)
}

/*
문단 오른쪽 끝 자동 탭
*/
func (p *HMakeContents) AutoTabRight(v ...int16) int {
	return funcParaSetInt(p.variant, "AutoTabRight", v)
}

/*
탭 채울 모양
*/
func (p *HMakeContents) Leader(v ...uint32) int {
	return funcParaSetInt(p.variant, "Leader", v)
}

/*
개요 문단 모으기
*/
func (p *HMakeContents) OutlineNumber(v ...int16) int {
	return funcParaSetInt(p.variant, "OutlineNumber", v)
}

/*
PIT_ARRAY 생성

	Styles을 위해 필요
*/
func (p *HMakeContents) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
모을 스타일 목록
*/
func (p *HMakeContents) Styles() *HArray {
	return GetHArray(p.variant, "Styles")
}

/*
모을 스타일 이름들 예:"스타일1;스타일2;...;"
*/
func (p *HMakeContents) StyleName(v ...string) string {
	return funcParaSetString(p.variant, "StyleName", v)
}

/*
만들 파일. 없으면 새탭에...
*/
func (p *HMakeContents) OutFileName(v ...string) string {
	return funcParaSetString(p.variant, "OutFileName", v)
}

/*
HPreference:

	HPreference :  환경 설정
*/
type HPreference struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HPreference) HSet() *HSet {
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
환경 설정 PropertySheet에 표시할 PropertyPage 번호 (하나만 선택)
*/
func (p *HPreference) ShowSinglePage(v ...int32) int {
	return funcParaSetInt(p.variant, "ShowSinglePage", v)
}

/*
하이퍼링크 글자 속성 문서 전체에 적용하기 여부 (on/off)
*/
func (p *HPreference) ApplyLinkAttr(v ...uint16) int {
	return funcParaSetInt(p.variant, "ApplyLinkAttr", v)
}

/*
HInternet:

	HInternet :  인터넷 문서 URL로 불러오기
*/
type HInternet struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HInternet) HSet() *HSet {
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
웹 브라우저로 불러오거나 한글 문서로 불러오기
*/
func (p *HInternet) OpenUrlWhere(v ...uint32) int {
	return funcParaSetInt(p.variant, "OpenUrlWhere", v)
}

/*
불러올 문서가 존재하는 URL
*/
func (p *HInternet) OpenUrlString(v ...string) string {
	return funcParaSetString(p.variant, "OpenUrlString", v)
}

/*
HFtpUpload:

	HFtpUpload :  웹 서버로 올리기
*/
type HFtpUpload struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFtpUpload) HSet() *HSet {
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
Ftp 서버 네임
*/
func (p *HFtpUpload) Server(v ...string) string {
	return funcParaSetString(p.variant, "Server", v)
}

/*
사용자 이름
*/
func (p *HFtpUpload) UserName_(v ...string) string {
	return funcParaSetString(p.variant, "UserName_", v)
}

/*
사용자 패스워드
*/
func (p *HFtpUpload) Password(v ...string) string {
	return funcParaSetString(p.variant, "Password", v)
}

/*
디렉토리
*/
func (p *HFtpUpload) Directory(v ...string) string {
	return funcParaSetString(p.variant, "Directory", v)
}

/*
파일 명
*/
func (p *HFtpUpload) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
PIT_ARRAY 생성

	SiteName을 위해 필요
*/
func (p *HFtpUpload) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
사이트 이름
*/
func (p *HFtpUpload) SiteName() *HArray {
	return GetHArray(p.variant, "SiteName")
}

/*
저장할 포맷. 0 = HTML 1 = HWP
*/
func (p *HFtpUpload) SaveType(v ...uint32) int {
	return funcParaSetInt(p.variant, "SaveType", v)
}

/*
HFileConvert:

	HFileConvert :  파일 변환(여러 파일을 동시에 특정 포맷으로 변환하여 저장)
*/
type HFileConvert struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileConvert) HSet() *HSet {
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
변환 포맷
*/
func (p *HFileConvert) Format(v ...string) string {
	return funcParaSetString(p.variant, "Format", v)
}

/*
PIT_ARRAY 생성

	SrcFileList, DestFileList을 위해 필요
*/
func (p *HFileConvert) CreateItemArray(itemname string, num int32) {
	CreateItemArray(p.variant, itemname, num)
}

/*
Source 파일 리스트
*/
func (p *HFileConvert) SrcFileList() *HArray {
	return GetHArray(p.variant, "SrcFileList")
}

/*
Destination 파일 리스트
*/
func (p *HFileConvert) DestFileList() *HArray {
	return GetHArray(p.variant, "DestFileList")
}

/*
HSelectionOpt:

	HSelectionOpt :  클립보드 Cut/Delete/Paset 옵션 (셀 블록 상태 일때))
*/
type HSelectionOpt struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSelectionOpt) HSet() *HSet {
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
셀 블록 상태 일 때 클립보드 Cut/Delete/Paste 처리 옵션

	Cut/Delete Selection 일때 :
		0 = 셀 모양을 남겨두고 셀 내용만 지운다.
		1 = 셀 모양을 오려낸다.
	Paste Selection 일때 :
		0 = 왼쪽에 끼워넣기.
		1 = 오른쪽에 끼워넣기.
		2 = 위쪽에 끼워넣기.
		3 = 아래쪽에 끼워넣기.
		4 = 셀 안에 표로 넣기.
		5 = 덮어 쓰기.
*/
func (p *HSelectionOpt) Option(v ...int16) int {
	return funcParaSetInt(p.variant, "Option", v)
}

/*
HDocumentInfo:

	HDocumentInfo :  문서 정보
*/
type HDocumentInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDocumentInfo) HSet() *HSet {
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
현재 위치한 문단
*/
func (p *HDocumentInfo) CurPara(v ...uint32) int {
	return funcParaSetInt(p.variant, "CurPara", v)
}

/*
현재 위치한 오프셋
*/
func (p *HDocumentInfo) CurPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "CurPos", v)
}

/*
현재 위치한 문단의 길이
*/
func (p *HDocumentInfo) CurParaLen(v ...uint32) int {
	return funcParaSetInt(p.variant, "CurParaLen", v)
}

/*
현재 리스트의 종류
*/
func (p *HDocumentInfo) CurCtrl(v ...uint32) int {
	return funcParaSetInt(p.variant, "CurCtrl", v)
}

/*
현재 리스트의 문단 수
*/
func (p *HDocumentInfo) CurParaCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "CurParaCount", v)
}

/*
루트 리스트의 현재 문단
*/
func (p *HDocumentInfo) RootPara(v ...uint32) int {
	return funcParaSetInt(p.variant, "RootPara", v)
}

/*
루트 리스트의 현재 오프셋
*/
func (p *HDocumentInfo) RootPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "RootPos", v)
}

/*
루트 리스트의 문단 수
*/
func (p *HDocumentInfo) RootParaCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "RootParaCount", v)
}

/*
자세한 정보까지 구할지 여부 on/off
*/
func (p *HDocumentInfo) DetailInfo(v ...int16) int {
	return funcParaSetInt(p.variant, "DetailInfo", v)
}

/*
문서에 포함된 글자 수
*/
func (p *HDocumentInfo) DetailCharCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailCharCount", v)
}

/*
문서에 포함된 어절 수
*/
func (p *HDocumentInfo) DetailWordCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailWordCount", v)
}

/*
문서에 포함된 줄 수
*/
func (p *HDocumentInfo) DetailLineCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailLineCount", v)
}

/*
문서에 포함된 쪽 수
*/
func (p *HDocumentInfo) DetailPageCount(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailPageCount", v)
}

/*
현재 쪽 번호
*/
func (p *HDocumentInfo) DetailCurPage(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailCurPage", v)
}

/*
현재 쪽 번호 (인쇄 번호)
*/
func (p *HDocumentInfo) DetailCurPrtPage(v ...uint32) int {
	return funcParaSetInt(p.variant, "DetailCurPrtPage", v)
}

/*
구역의 속성까지 구할지 여부 on/off
*/
func (p *HDocumentInfo) SectionInfo(v ...int16) int {
	return funcParaSetInt(p.variant, "SectionInfo", v)
}

/*
구역의 속성
*/
func (p *HDocumentInfo) SecDef(v ...*HSecDef) *HSecDef {
	var para HSecDef
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "SecDef", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "SecDef")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HExchangeFootnoteEndNote:

	HExchangeFootnoteEndNote :  각주/미주 변환
*/
type HExchangeFootnoteEndNote struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HExchangeFootnoteEndNote) HSet() *HSet {
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
옵션

	0 : 모든 각주를 미주로 바꾸기
	1 : 모든 미주를 각주로 바꾸기
	2 : 각주와 미주를 서로 바꾸기
*/
func (p *HExchangeFootnoteEndNote) Flag(v ...uint16) int {
	return funcParaSetInt(p.variant, "Flag", v)
}

/*
HSaveFootnote:

	HSaveFootnote :  주석 저장
*/
type HSaveFootnote struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSaveFootnote) HSet() *HSet {
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
파일 이름
*/
func (p *HSaveFootnote) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
옵션

	1 : 각주 저장
	2 : 미주 저장
	3 : 각주/미주 저장
*/
func (p *HSaveFootnote) Flag(v ...uint16) int {
	return funcParaSetInt(p.variant, "Flag", v)
}

/*
HChangeRome:

	HChangeRome :  로마자 변환
*/
type HChangeRome struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HChangeRome) HSet() *HSet {
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
변환 옵션
*/
func (p *HChangeRome) Option(v ...uint16) int {
	return funcParaSetInt(p.variant, "Option", v)
}

/*
변환할 한글 문자열
*/
func (p *HChangeRome) HanString(v ...string) string {
	return funcParaSetString(p.variant, "HanString", v)
}

/*
HSearchAddress:

	HSearchAddress :  주소 검색
*/
type HSearchAddress struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSearchAddress) HSet() *HSet {
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
검색할 단어
*/
func (p *HSearchAddress) SearchWord(v ...string) string {
	return funcParaSetString(p.variant, "SearchWord", v)
}

/*
주소
*/
func (p *HSearchAddress) Address(v ...string) string {
	return funcParaSetString(p.variant, "Address", v)
}

/*
우편
*/
func (p *HSearchAddress) Post(v ...string) string {
	return funcParaSetString(p.variant, "Post", v)
}

/*
기타 주소
*/
func (p *HSearchAddress) EtcAddress(v ...string) string {
	return funcParaSetString(p.variant, "EtcAddress", v)
}

/*
주소 검색 옵션

	미구현
*/
func (p *HSearchAddress) Option(v ...uint32) int {
	return funcParaSetInt(p.variant, "Option", v)
}

/*
로마자 변환 유

	미구현
*/
func (p *HSearchAddress) IsChangeRome(v ...uint32) int {
	return funcParaSetInt(p.variant, "IsChangeRome", v)
}

/*
HAddHanjaWord:

	HAddHanjaWord :  한자 단어 추가
*/
type HAddHanjaWord struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HAddHanjaWord) HSet() *HSet {
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
한글음
*/
func (p *HAddHanjaWord) Hangul(v ...string) string {
	return funcParaSetString(p.variant, "Hangul", v)
}

/*
한자
*/
func (p *HAddHanjaWord) Hanja(v ...string) string {
	return funcParaSetString(p.variant, "Hanja", v)
}

/*
HDocFindInfo:

	HDocFindInfo :  문서 찾기 결과
*/
type HDocFindInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDocFindInfo) HSet() *HSet {
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
List ID
*/
func (p *HDocFindInfo) ListID(v ...uint32) int {
	return funcParaSetInt(p.variant, "ListID", v)
}

/*
Para ID
*/
func (p *HDocFindInfo) ParaID(v ...uint32) int {
	return funcParaSetInt(p.variant, "ParaID", v)
}

/*
pos
*/
func (p *HDocFindInfo) Pos(v ...uint32) int {
	return funcParaSetInt(p.variant, "Pos", v)
}

/*
HSum:

	HSum :  블록 계산
*/
type HSum struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSum) HSet() *HSet {
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
합
*/
func (p *HSum) Sum(v ...string) string {
	return funcParaSetString(p.variant, "Sum", v)
}

/*
평균
*/
func (p *HSum) Average(v ...string) string {
	return funcParaSetString(p.variant, "Average", v)
}

/*
줄수
*/
func (p *HSum) LineCount(v ...string) string {
	return funcParaSetString(p.variant, "LineCount", v)
}

/*
세 자리마다 쉼표로 자리 구분 (On/Off)
*/
func (p *HSum) Comma(v ...uint16) int {
	return funcParaSetInt(p.variant, "Comma", v)
}

/*
형식 옵션
*/
func (p *HSum) Option(v ...int32) int {
	return funcParaSetInt(p.variant, "Option", v)
}

/*
HListParaPos:

	HListParaPos :  현재 캐럿의 위치 (GetPos, SetPos API에서 사용)
*/
type HListParaPos struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HListParaPos) HSet() *HSet {
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
List ID
*/
func (p *HListParaPos) List(v ...uint32) int {
	return funcParaSetInt(p.variant, "List", v)
}

/*
Para ID
*/
func (p *HListParaPos) Para(v ...uint32) int {
	return funcParaSetInt(p.variant, "Para", v)
}

/*
Pos
*/
func (p *HListParaPos) Pos(v ...uint32) int {
	return funcParaSetInt(p.variant, "Pos", v)
}

/*
HGetText:

	HGetText :  본문의 내용 얻어오기 (GetText API에서 사용)
*/
type HGetText struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HGetText) HSet() *HSet {
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
본문 내용
*/
func (p *HGetText) Text(v ...string) string {
	return funcParaSetString(p.variant, "Text", v)
}

/*
HConvertJianFan:

	HConvertJianFan :  간/번체 변환
*/
type HConvertJianFan struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HConvertJianFan) HSet() *HSet {
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
변환 타입
*/
func (p *HConvertJianFan) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
HMousePos: 	마우스 위치 알아오기 (GetMousePos API에서 사용)
*/
type HMousePos struct {
	variant *ole.VARIANT
}

/*
상대적 기준 X

	0: 종이 기준
	1: 쪽 기준
*/
func (p *HMousePos) XRelTo() int {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", "XRelTo")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint32))
}

/*
상대적 기준 Y

	0: 종이 기준
	1: 쪽 기준
*/
func (p *HMousePos) YRelTo() int {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", "YRelTo")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint32))
}

/*
페이지 번호 (0 based)
*/
func (p *HMousePos) Page() int {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", "Page")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint32))
}

/*
X 위치 HWPUNIT
*/
func (p *HMousePos) X() int {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", "X")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint32))
}

/*
Y 위치 HWPUNIT
*/
func (p *HMousePos) Y() int {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Item", "Y")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(uint32))
}

/*
HFileInfo:

	HFileInfo :  파일 정보 알아오기 (GetFileInfo API에서 사용)
*/
type HFileInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileInfo) HSet() *HSet {
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
파일 형식
*/
func (p *HFileInfo) Format(v ...string) string {
	return funcParaSetString(p.variant, "Format", v)
}

/*
버전
*/
func (p *HFileInfo) VersionStr(v ...string) string {
	return funcParaSetString(p.variant, "VersionStr", v)
}

/*
버전
*/
func (p *HFileInfo) VersionNum(v ...uint32) int {
	return funcParaSetInt(p.variant, "VersionNum", v)
}

/*
암호화 여부
*/
func (p *HFileInfo) Encrypted(v ...int32) int {
	return funcParaSetInt(p.variant, "Encrypted", v)
}

/*
HActionCrossRef:

	HActionCrossRef :  상호 참조
*/
type HActionCrossRef struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HActionCrossRef) HSet() *HSet {
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
경로
*/
func (p *HActionCrossRef) Command(v ...string) string {
	return funcParaSetString(p.variant, "Command", v)
}

/*
HSaveUserInfoFile:

	HSaveUserInfoFile :  사용자 추가 정보 저장
*/
type HSaveUserInfoFile struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HSaveUserInfoFile) HSet() *HSet {
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
바탕 문서 정보저장 여부
*/
func (p *HSaveUserInfoFile) SaveBaseTemplDoc(v ...int32) int {
	return funcParaSetInt(p.variant, "SaveBaseTemplDoc", v)
}

/*
매크로와 상용구 파일 경로의 존재 여부
*/
func (p *HSaveUserInfoFile) ExistMacroIdiomPath(v ...int32) int {
	return funcParaSetInt(p.variant, "ExistMacroIdiomPath", v)
}

/*
매크로와 상용구 파일 경로
*/
func (p *HSaveUserInfoFile) MacroIdiomPath(v ...string) string {
	return funcParaSetString(p.variant, "MacroIdiomPath", v)
}

/*
바탕문서 파일 경로의 존재 여부
*/
func (p *HSaveUserInfoFile) ExistTemplDocPath(v ...int32) int {
	return funcParaSetInt(p.variant, "ExistTemplDocPath", v)
}

/*
바탕문서 파일 경로
*/
func (p *HSaveUserInfoFile) TemplDocPath(v ...string) string {
	return funcParaSetString(p.variant, "TemplDocPath", v)
}

/*
HLoadUserInfoFile:

	HLoadUserInfoFile :  사용자 추가 정보 가져오기
*/
type HLoadUserInfoFile struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HLoadUserInfoFile) HSet() *HSet {
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
현재 정보에 추가
*/
func (p *HLoadUserInfoFile) MergeCurUserInfo(v ...int32) int {
	return funcParaSetInt(p.variant, "MergeCurUserInfo", v)
}

/*
현재 정보 지움
*/
func (p *HLoadUserInfoFile) DeleteCurUserInfo(v ...int32) int {
	return funcParaSetInt(p.variant, "DeleteCurUserInfo", v)
}

/*
바탕 문서 가져오기
*/
func (p *HLoadUserInfoFile) LoadBaseTemplDoc(v ...int32) int {
	return funcParaSetInt(p.variant, "LoadBaseTemplDoc", v)
}

/*
현재 사용자 정보 저장
*/
func (p *HLoadUserInfoFile) SaveCurUserInfo(v ...int32) int {
	return funcParaSetInt(p.variant, "SaveCurUserInfo", v)
}

/*
매크로와 상용구 파일 경로의 존재 여부
*/
func (p *HLoadUserInfoFile) ExistMacroIdiomPath(v ...int32) int {
	return funcParaSetInt(p.variant, "ExistMacroIdiomPath", v)
}

/*
매크로와 상용구 파일 경로
*/
func (p *HLoadUserInfoFile) MacroIdiomPath(v ...string) string {
	return funcParaSetString(p.variant, "MacroIdiomPath", v)
}

/*
바탕문서 파일 경로의 존재 여부
*/
func (p *HLoadUserInfoFile) ExistTemplDocPath(v ...int32) int {
	return funcParaSetInt(p.variant, "ExistTemplDocPath", v)
}

/*
바탕문서 파일 경로
*/
func (p *HLoadUserInfoFile) TemplDocPath(v ...string) string {
	return funcParaSetString(p.variant, "TemplDocPath", v)
}

/*
HUserQCommandFile:

	HUserQCommandFile :  사용자 자동 명령 파일 저장/로드
*/
type HUserQCommandFile struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HUserQCommandFile) HSet() *HSet {
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
저장 (TRUE = Save, FALSE = Open)
*/
func (p *HUserQCommandFile) Save(v ...int32) int {
	return funcParaSetInt(p.variant, "Save", v)
}

/*
파일명
*/
func (p *HUserQCommandFile) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
로드 방법 (TRUE = Overwrite, FALSE = Merge)
*/
func (p *HUserQCommandFile) LoadType(v ...int32) int {
	return funcParaSetInt(p.variant, "LoadType", v)
}

/*
HInputDateStyle:

	HInputDateStyle :  날짜/시간 표시 형식
*/
type HInputDateStyle struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HInputDateStyle) HSet() *HSet {
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
문자열로 넣기/코드로 넣기
*/
func (p *HInputDateStyle) DateStyleType(v ...uint16) int {
	return funcParaSetInt(p.variant, "DateStyleType", v)
}

/*
필드 컨트롤의 안내문/지시문
*/
func (p *HInputDateStyle) DateStyleDataForm(v ...string) string {
	return funcParaSetString(p.variant, "DateStyleDataForm", v)
}

/*
HFileSetSecurity:

	HFileSetSecurity :  배포용 문서 암호
*/
type HFileSetSecurity struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileSetSecurity) HSet() *HSet {
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
배포용 문서 암호
*/
func (p *HFileSetSecurity) Password(v ...string) string {
	return funcParaSetString(p.variant, "Password", v)
}

/*
TRUE = 문서 인쇄 불가능, FALSE = 문서 인쇄 가능. on/off
*/
func (p *HFileSetSecurity) NoPrint(v ...uint16) int {
	return funcParaSetInt(p.variant, "NoPrint", v)
}

/*
TRUE = 문서 내용 복사 및 추출 불가능, FALSE = 문서 내용 복사 및 추출 가능. on/off
*/
func (p *HFileSetSecurity) NoCopy(v ...uint16) int {
	return funcParaSetInt(p.variant, "NoCopy", v)
}

/*
HLinkDocument:

	HLinkDocument :  문서 연결
*/
type HLinkDocument struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HLinkDocument) HSet() *HSet {
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
연결 문서 FILE NAME
*/
func (p *HLinkDocument) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

/*
TRUE = 쪽 번호 잇기, FALSE = 쪽 번호 잇지 않기
*/
func (p *HLinkDocument) PageInherit(v ...uint16) int {
	return funcParaSetInt(p.variant, "PageInherit", v)
}

/*
TRUE = 각주 번호 잇기, FALSE = 각주 번호 잇지 않기
*/
func (p *HLinkDocument) FootnoteInherit(v ...uint16) int {
	return funcParaSetInt(p.variant, "FootnoteInherit", v)
}

/*
HDocumentFilterDialog:

	HDocumentFilterDialog :  파일 필터에서 사용하는 다이얼로그
*/
type HDocumentFilterDialog struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HDocumentFilterDialog) HSet() *HSet {
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
문자 코드 선택
*/
func (p *HDocumentFilterDialog) TextCodeType(v ...int32) int {
	return funcParaSetInt(p.variant, "TextCodeType", v)
}

/*
기본 값으로 지정
*/
func (p *HDocumentFilterDialog) TextDefault(v ...uint32) int {
	return funcParaSetInt(p.variant, "TextDefault", v)
}

/*
문서 경로
*/
func (p *HDocumentFilterDialog) TextFilePath(v ...string) string {
	return funcParaSetString(p.variant, "TextFilePath", v)
}

/*
선택 사항 : 0 = 줄 단위, 1 = 문단 단위
*/
func (p *HDocumentFilterDialog) TextReadType(v ...uint32) int {
	return funcParaSetInt(p.variant, "TextReadType", v)
}

/*
글꼴 설정
*/
func (p *HDocumentFilterDialog) TextFontName(v ...string) string {
	return funcParaSetString(p.variant, "TextFontName", v)
}

/*
정렬 유지
*/
func (p *HDocumentFilterDialog) TextKeepAlign(v ...uint32) int {
	return funcParaSetInt(p.variant, "TextKeepAlign", v)
}

/*
문자 코드 선택
*/
func (p *HDocumentFilterDialog) UnicodeCodeType(v ...int32) int {
	return funcParaSetInt(p.variant, "UnicodeCodeType", v)
}

/*
문서 경로
*/
func (p *HDocumentFilterDialog) UnicodeFilePath(v ...string) string {
	return funcParaSetString(p.variant, "UnicodeFilePath", v)
}

/*
정렬 유지
*/
func (p *HDocumentFilterDialog) UnicodeKeepAlign(v ...uint32) int {
	return funcParaSetInt(p.variant, "UnicodeKeepAlign", v)
}

/*
문자 코드 선택
*/
func (p *HDocumentFilterDialog) HtmlCodeType(v ...int32) int {
	return funcParaSetInt(p.variant, "HtmlCodeType", v)
}

/*
기본 값으로 지정
*/
func (p *HDocumentFilterDialog) HtmlDefault(v ...uint32) int {
	return funcParaSetInt(p.variant, "HtmlDefault", v)
}

/*
문서 경로
*/
func (p *HDocumentFilterDialog) HtmlFilePath(v ...string) string {
	return funcParaSetString(p.variant, "HtmlFilePath", v)
}

/*
보내기의 대상 0x0040 = 편지로 보내기 0x0080 = 웹 브라우저로 보내기 FlashPageStart PIT_UI 시작 쪽
*/
func (p *HDocumentFilterDialog) HtmlSend(v ...uint32) int {
	return funcParaSetInt(p.variant, "HtmlSend", v)
}

/*
끝 쪽
*/
func (p *HDocumentFilterDialog) FlashPageEnd(v ...uint32) int {
	return funcParaSetInt(p.variant, "FlashPageEnd", v)
}

/*
배경 색
*/
func (p *HDocumentFilterDialog) FlashBackgroundColor(v ...uint32) int {
	return funcParaSetInt(p.variant, "FlashBackgroundColor", v)
}

/*
쪽 이동 단추 넣기
*/
func (p *HDocumentFilterDialog) FlashNextButton(v ...uint32) int {
	return funcParaSetInt(p.variant, "FlashNextButton", v)
}

/*
선택되지 않은 이미지 경로
*/
func (p *HDocumentFilterDialog) FlashNonselectImagePath(v ...string) string {
	return funcParaSetString(p.variant, "FlashNonselectImagePath", v)
}

/*
선택된 이미지 경로
*/
func (p *HDocumentFilterDialog) FlashSelectImagePath(v ...string) string {
	return funcParaSetString(p.variant, "FlashSelectImagePath", v)
}

/*
저장 옵션 : TRUE = 그림으로 저장, FALSE = 인터넷 문서로 저장
*/
func (p *HDocumentFilterDialog) DocImgSaveType(v ...uint32) int {
	return funcParaSetInt(p.variant, "DocImgSaveType", v)
}

/*
문자 코드 선택
*/
func (p *HDocumentFilterDialog) DbfCodeType(v ...int32) int {
	return funcParaSetInt(p.variant, "DbfCodeType", v)
}

/*
읽기 방식 : TRUE = 표로 읽기, FALSE = 텍스트 형식으로 읽기
*/
func (p *HDocumentFilterDialog) DbfReadType(v ...uint32) int {
	return funcParaSetInt(p.variant, "DbfReadType", v)
}

/*
레코드 시작 범위
*/
func (p *HDocumentFilterDialog) DbfRecordStart(v ...uint32) int {
	return funcParaSetInt(p.variant, "DbfRecordStart", v)
}

/*
레코드 끝 범위
*/
func (p *HDocumentFilterDialog) DbfRecordEnd(v ...uint32) int {
	return funcParaSetInt(p.variant, "DbfRecordEnd", v)
}

/*
레코드 번호 붙이기
*/
func (p *HDocumentFilterDialog) DbfRecordNumbering(v ...uint32) int {
	return funcParaSetInt(p.variant, "DbfRecordNumbering", v)
}

/*
필드 오른쪽 공백 없애기
*/
func (p *HDocumentFilterDialog) DbfDelSpace(v ...uint32) int {
	return funcParaSetInt(p.variant, "DbfDelSpace", v)
}

/*
HShapeObjSaveAsPicture:

	HShapeObjSaveAsPicture :  그리기개체를 그림으로 저장하기
*/
type HShapeObjSaveAsPicture struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HShapeObjSaveAsPicture) HSet() *HSet {
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
선택되지 않은 이미지 PATH (파일명까지 포함한다)
*/
func (p *HShapeObjSaveAsPicture) Path(v ...string) string {
	return funcParaSetString(p.variant, "Path", v)
}

/*
선택된 이미지 확장자. 생략하면 비트맵으로 저장된다. bmp: 비트맵 gif: gif 파일 jpg: jpg 파일 png: 포토샵 파일 wmf: window metafile emf: enhanced metafile
*/
func (p *HShapeObjSaveAsPicture) Ext(v ...string) string {
	return funcParaSetString(p.variant, "Ext", v)
}

/*
HEngineProperties:

	HEngineProperties :  환경 설정 속성
*/
type HEngineProperties struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HEngineProperties) HSet() *HSet {
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
마우스로 두 번 누르기 한 곳에 입력 가능
*/
func (p *HEngineProperties) DoAnyCursorEdit(v ...uint16) int {
	return funcParaSetInt(p.variant, "DoAnyCursorEdit", v)
}

/*
개요 번호 삽입 문단에 개요 스타일 자동 적용
*/
func (p *HEngineProperties) DoOutlineStyle(v ...uint16) int {
	return funcParaSetInt(p.variant, "DoOutlineStyle", v)
}

/*
삽입 잠금
*/
func (p *HEngineProperties) InsertLock(v ...uint16) int {
	return funcParaSetInt(p.variant, "InsertLock", v)
}

/*
표 안에서 <Tab>으로 셀 이동
*/
func (p *HEngineProperties) TabMoveCell(v ...uint16) int {
	return funcParaSetInt(p.variant, "TabMoveCell", v)
}

/*
팩스 드라이버
*/
func (p *HEngineProperties) FaxDriver(v ...string) string {
	return funcParaSetString(p.variant, "FaxDriver", v)
}

/*
PDF 드라이버
*/
func (p *HEngineProperties) PDFDriver(v ...string) string {
	return funcParaSetString(p.variant, "PDFDriver", v)
}

/*
맞춤법 도우미 작동
*/
func (p *HEngineProperties) EnableAutoSpell(v ...uint16) int {
	return funcParaSetInt(p.variant, "EnableAutoSpell", v)
}

/*
새 창으로 불러오기
*/
func (p *HEngineProperties) OpenNewWindow(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenNewWindow", v)
}

/*
HScrollPosInfo:

	HScrollPosInfo :  스크롤바 정보 알아오기
*/
type HScrollPosInfo struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HScrollPosInfo) HSet() *HSet {
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
수평 스크롤바 포지션
*/
func (p *HScrollPosInfo) HorzPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "HorzPos", v)
}

/*
수평 스크롤바 Limit 포지션
*/
func (p *HScrollPosInfo) HorzLimitPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "HorzLimitPos", v)
}

/*
수직 스크롤바 포지션
*/
func (p *HScrollPosInfo) VertPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "VertPos", v)
}

/*
수직 스크롤바 Limit 포지션
*/
func (p *HScrollPosInfo) VertLimitPos(v ...uint32) int {
	return funcParaSetInt(p.variant, "VertLimitPos", v)
}

/*
HMessageSet:

	HMessageSet :  이벤트 메세지 Set
*/
type HMessageSet struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HMessageSet) HSet() *HSet {
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
메세지 (UniCode)
*/
func (p *HMessageSet) String(v ...string) string {
	return funcParaSetString(p.variant, "String", v)
}

/*
HFileXMLSchema:

	HFileXMLSchema :  XML 관련 메뉴
*/
type HFileXMLSchema struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileXMLSchema) HSet() *HSet {
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
XML 노드 생성 옵션
*/
func (p *HFileXMLSchema) NodeCreationOption(v ...uint32) int {
	return funcParaSetInt(p.variant, "NodeCreationOption", v)
}

/*
XML 파일 불러오기/저장 (HXMLOpenSave)
*/
func (p *HFileXMLSchema) XMLOpenSave(v ...*HXMLOpenSave) *HXMLOpenSave {
	var para HXMLOpenSave
	if len(v) == 1 {
		oleutil.CallMethod(p.variant.ToIDispatch(), "SetItem", "XMLOpenSave", v[0].variant.ToIDispatch())
	}
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "XMLOpenSave")
	if err != nil {
		panic(err)
	}
	para.variant = res
	//para.HSet()
	return &para
}

/*
HFormObjInputCodeTable:

	HFormObjInputCodeTable :  양식 개체(Edit)에서만 사용하는 코드표
*/
type HFormObjInputCodeTable struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormObjInputCodeTable) HSet() *HSet {
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
디폴트 문자 코드
*/
func (p *HFormObjInputCodeTable) DefaultCode(v ...uint16) int {
	return funcParaSetInt(p.variant, "DefaultCode", v)
}

/*
입력할 코드 문자열 (리턴값)
*/
func (p *HFormObjInputCodeTable) ResultString(v ...string) string {
	return funcParaSetString(p.variant, "ResultString", v)
}

/*
HFormObjInputHanja:

	HFormObjInputHanja :  양식 개체(Edit))에서만 사용하는 한자 변환
*/
type HFormObjInputHanja struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormObjInputHanja) HSet() *HSet {
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
변환할 문자열
*/
func (p *HFormObjInputHanja) InputString(v ...string) string {
	return funcParaSetString(p.variant, "InputString", v)
}

/*
한 글자만 변환 여부
*/
func (p *HFormObjInputHanja) OneChar(v ...uint16) int {
	return funcParaSetInt(p.variant, "OneChar", v)
}

/*
변환된 문자열(리턴값)
*/
func (p *HFormObjInputHanja) ResultString(v ...string) string {
	return funcParaSetString(p.variant, "ResultString", v)
}

/*
사용된 문자열의 길이(리턴값)
*/
func (p *HFormObjInputHanja) ResultLength(v ...uint32) int {
	return funcParaSetInt(p.variant, "ResultLength", v)
}

/*
HFormObjHanjaBusu:

	HFormObjHanjaBusu :  양식 개체(Edit))에서만 사용하는 한자 부수 입력
*/
type HFormObjHanjaBusu struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormObjHanjaBusu) HSet() *HSet {
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
입력될 한자
*/
func (p *HFormObjHanjaBusu) ResultString(v ...string) string {
	return funcParaSetString(p.variant, "ResultString", v)
}

/*
HFormObjHanjaMean:

	HFormObjHanjaMean :  양식 개체(Edit))에서만 사용하는 한자 새김 입력
*/
type HFormObjHanjaMean struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormObjHanjaMean) HSet() *HSet {
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
입력될 한자
*/
func (p *HFormObjHanjaMean) ResultString(v ...string) string {
	return funcParaSetString(p.variant, "ResultString", v)
}

/*
HFormObjInputIdiom:

	HFormObjInputIdiom :  양식 개체(Edit))에서만 사용하는 상용구 입력
*/
type HFormObjInputIdiom struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFormObjInputIdiom) HSet() *HSet {
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
검색 문자열
*/
func (p *HFormObjInputIdiom) SearchString(v ...string) string {
	return funcParaSetString(p.variant, "SearchString", v)
}

/*
변환된 문자열(리턴값)
*/
func (p *HFormObjInputIdiom) ResultString(v ...string) string {
	return funcParaSetString(p.variant, "ResultString", v)
}

/*
사용된 준말의 길이
*/
func (p *HFormObjInputIdiom) ResultLength(v ...uint32) int {
	return funcParaSetInt(p.variant, "ResultLength", v)
}

/*
HViewStatus:

	HViewStatus :  뷰 상태 정보
*/
type HViewStatus struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HViewStatus) HSet() *HSet {
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
뷰의 상태를 얻기 위한 옵션
*/
func (p *HViewStatus) Type(v ...uint32) int {
	return funcParaSetInt(p.variant, "Type", v)
}

/*
X 위치
*/
func (p *HViewStatus) ViewPosX(v ...int32) int {
	return funcParaSetInt(p.variant, "ViewPosX", v)
}

/*
Y 위치
*/
func (p *HViewStatus) ViewPosY(v ...int32) int {
	return funcParaSetInt(p.variant, "ViewPosY", v)
}

/*
HScriptMacro:

	HScriptMacro :  스크립트 매크로
*/
type HScriptMacro struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HScriptMacro) HSet() *HSet {
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
정의(or 실행)할 매크로의 인텍스
*/
func (p *HScriptMacro) Index(v ...int32) int {
	return funcParaSetInt(p.variant, "Index", v)
}

/*
실행 반복 횟수
*/
func (p *HScriptMacro) RepeatCount(v ...int32) int {
	return funcParaSetInt(p.variant, "RepeatCount", v)
}

/*
매크로 이름
*/
func (p *HScriptMacro) Name(v ...string) string {
	return funcParaSetString(p.variant, "Name", v)
}

/*
매크로 설명
*/
func (p *HScriptMacro) Detail(v ...string) string {
	return funcParaSetString(p.variant, "Detail", v)
}

/*
HFileOpenSave:

	HFileOpenSave :  파일 Open/Save (파일 대화상자)
*/
type HFileOpenSave struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HFileOpenSave) HSet() *HSet {
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
파일 이름
*/
func (p *HFileOpenSave) FileName(v ...string) string {
	return funcParaSetString(p.variant, "FileName", v)
}

/*
format
*/
func (p *HFileOpenSave) Format(v ...string) string {
	return funcParaSetString(p.variant, "Format", v)
}

/*
읽기 전용
*/
func (p *HFileOpenSave) OpenReadOnly(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenReadOnly", v)
}

/*
문서 열기/저장하기 옵션
*/
func (p *HFileOpenSave) OpenFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenFlag", v)
}

/*
해당 format파일을 Open 및 Save할 경우 사용되는 옵션
*/
func (p *HFileOpenSave) Argument(v ...string) string {
	return funcParaSetString(p.variant, "Argument", v)
}

/*
해당 format파일을 Open 및 Save할 경우 사용되는 속성
*/
func (p *HFileOpenSave) Attributes(v ...uint32) int {
	return funcParaSetInt(p.variant, "Attributes", v)
}

/*
HXMLOpenSave:

	HXMLOpenSave :  XML Open/Save (파일 대화상자)
*/
type HXMLOpenSave struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HXMLOpenSave) HSet() *HSet {
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
파일 이름
*/
func (p *HXMLOpenSave) Filename(v ...string) string {
	return funcParaSetString(p.variant, "Filename", v)
}

/*
해당 format파일을 Open 및 Save할 경우 사용되는 옵션
*/
func (p *HXMLOpenSave) Argument(v ...string) string {
	return funcParaSetString(p.variant, "Argument", v)
}

/*
문서 열기/저장하기 옵션
*/
func (p *HXMLOpenSave) OpenFlag(v ...uint16) int {
	return funcParaSetInt(p.variant, "OpenFlag", v)
}

/*
HXMLSchema:

	HXMLSchema :  XML Schema 만들기
*/
type HXMLSchema struct {
	variant *ole.VARIANT
	hset    *HSet
}

/*
HSet:
*/
func (p *HXMLSchema) HSet() *HSet {
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
XML 서식 제목
*/
func (p *HXMLSchema) TemplateName(v ...string) string {
	return funcParaSetString(p.variant, "TemplateName", v)
}

/*
XML 스키마 파일 이름
*/
func (p *HXMLSchema) SchemaFileName(v ...string) string {
	return funcParaSetString(p.variant, "SchemaFileName", v)
}

/*
서식 파일 이름
*/
func (p *HXMLSchema) TemplateFileName(v ...string) string {
	return funcParaSetString(p.variant, "TemplateFileName", v)
}

package gohwp

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
IDHwpCtrlCode:

	IDHwpCtrlCode: 문서 내부의 표, 각주 등의 컨트롤(특수 문자 포함)를 나타내는 오브젝트이다.
	(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpCtrlCode와 동일)
*/
type IDHwpCtrlCode struct {
	variant *ole.VARIANT
}

/*
컨트롤 문자. (읽기전용)

	일반적으로 컨트롤 ID를 사용해 컨트롤의 종류를 판별하지만, 이보다 더 포괄적인 범주를 나타내는 컨트롤 문자로 판별할 수도 있다. 예를 들어 각주와 미주는 ID는 다르지만, 컨트롤 문자는 17로 동일하다. 컨트롤 문자는 1-31 사이의 값을 사용한다.
	Ch    설명
	1     예약
	2     구역/단 정의
	3     필드 시작
	4     필드 끝
	5     예약
	6     예약
	7     예약
	8     예약
	9     탭
	10    강제 줄 나눔
	11    그리기 개체/표
	12    예약
	13    문단 나누기
	14    예약
	15    주석
	16    머리말/꼬리말
	17    각주/미주
	18    자동 번호
	19    예약
	20    예약
	21    쪽바뀜
	22    책갈피 / 찾아보기 표시
	23    덧말 / 글짜 겹침
	24    하이픈
	25    예약
	26    예약
	27    예약
	28    예약
	29    예약
	30    묶음 빈칸
	31    고정 폭 빈칸
*/
func (p *IDHwpCtrlCode) CtrlCh() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CtrlCh")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
컨트롤 ID. (읽기 전용)

	컨트롤 ID는 컨트롤의 종류를 나타내기 위해 할당된 ID로서, 최대 4개의 문자로 구성된 문자열이다. 예를 들어 표는 "tbl", 각주는 "fn"이다. 워디안에서 현재까지 지원되는 모든 컨트롤의 ID는 다음 표 참조.

	ID      Property Set    Initialization Set  설명
	cold    ColDef          ColDef              단
	secd    SecDef          SecDef              구역
	fn      FootnoteShape   FootnoteShape       각주
	en      FootnoteShape   FootnoteShape       미주
	tbl     Table           TableCreation       표
	eqed    EqEdit          EqEdit              수식
	gso     ShapeObject     ShapeObject         그리기 개체
	atno    AutoNum         AutoNum             번호 넣기
	nwno    AutoNum         AutoNum             새 번호로
	pgct    PageNumCtrl     PageNumCtrl         페이지 번호 제어 (97의 홀수 쪽에서 시작)
	pghd    PageHiding      PageHiding          감추기
	pgnp    PageNumPos      PageNumPos          쪽 번호 위치
	head    HeaderFooter    HeaderFooter        머리말
	foot    HeaderFooter    HeaderFooter        꼬리말
	%dte    FieldCtrl       FieldCtrl           현재의 날짜/시간 필드
	%ddt    FieldCtrl       FieldCtrl           파일 작성 날짜/시간 필드
	%pat    FieldCtrl       FieldCtrl           문서 경로 필드
	%bmk    FieldCtrl       FieldCtrl           블록 책갈피
	%mmg    FieldCtrl       FieldCtrl           메일 머지
	%xrf    FieldCtrl       FieldCtrl           상호 참조
	%fmu    FieldCtrl       FieldCtrl           계산식
	%clk    FieldCtrl       FieldCtrl           누름틀
	%smr    FieldCtrl       FieldCtrl           문서 요약 정보 필드
	%usr    FieldCtrl       FieldCtrl           사용자 정보 필드
	%hlk    FieldCtrl       FieldCtrl           하이퍼링크
	bokm    TextCtrl        TextCtrl            책갈피
	idxm    IndexMark       IndexMark           찾아보기
	tdut    Dutmal          Dutmal              덧말
	tcmt    없음            없음                 주석
*/
func (p *IDHwpCtrlCode) CtrlID() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "CtrlID")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
글상자를 지원하는지의 여부(읽기 전용)
*/
func (p *IDHwpCtrlCode) HasList() bool {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HasList")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
다음 컨트롤.(읽기 전용)

	문서 중의 모든 컨트롤(표, 그림등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list중 다음 컨트롤을 나타낸다.
*/
func (p *IDHwpCtrlCode) Next() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Next")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
앞 컨트롤.(읽기 전용)

	문서 중의 모든 컨트롤(표, 그림등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list중 다음 컨트롤을 나타낸다.
*/
func (p *IDHwpCtrlCode) Prev() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Prev")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
컨트롤의 속성을 나타낸다.

	※ 미구현
	모든 컨트롤은 대응하는 parameter set으로 속성을 읽고 쓸 수 있다.
*/
func (p *IDHwpCtrlCode) Properties() interface{} {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Properties")
	if err != nil {
		panic(err)
	}
	res.Value() // 삭제
	curId := p.CtrlID()
	fmt.Println(curId)
	switch curId {
	case "cold":
	case "secd":
	case "fn", "en":
	case "tbl":
	case "eqed":
	case "atno", "nwno":
	case "pgct":
	case "pghd":
	case "pgnp":
	case "head", "foot":
	case "%dte", "%ddt", "%pat", "%bmk", "%mmg", "%xrf", "%fmu", "%clk", "%smr", "%usr", "%hlk":
		fmt.Println("여기에 왔는가?")
	case "bokm":
	case "idxm":
	case "tdut":
	case "tcmt":
		return nil
	}
	return nil
}

/*
컨트롤의 종류를 사용자에게 보여줄 수 있는 localize된 문자열로 나타낸다. (읽기 전용)
*/
func (p *IDHwpCtrlCode) UserDesc() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "UserDesc")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
컨트롤의 anchor의 위치를 리턴한다.

	param: 기준 위치
	0: 바로 상위 리스트에서의 anchor position (default)
	1: 탑레벨 리스트에서의 anchor position
	2: 루트 리스트에서의 anchor position
*/
func (p *IDHwpCtrlCode) GetAnchorPos(param ...int) (list, para, pos int) {
	t := 0
	if len(param) > 0 {
		t = param[0]
	}
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "GetAnchorPos", t)
	if err != nil {
		panic(err)
	}
	ilist, err := oleutil.CallMethod(res.ToIDispatch(), "Item", "List")
	if err != nil {
		panic(err)
	}
	ipara, err := oleutil.CallMethod(res.ToIDispatch(), "Item", "Para")
	if err != nil {
		panic(err)
	}
	ipos, err := oleutil.CallMethod(res.ToIDispatch(), "Item", "Pos")
	if err != nil {
		panic(err)
	}
	return int(ilist.Value().(uint32)), int(ipara.Value().(uint32)), int(ipos.Value().(uint32))
}

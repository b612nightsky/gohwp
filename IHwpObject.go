package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
OLE Automation Object

IHwpObject : 최상위 개체 (모든 Automation Object의 최상위 Object)
*/

/*
문서가 변경되었는지 확인 (읽기전용)

	true: 변경됨
	false: 변경 안됨
*/
func (hwp *IHwpObject) IsModified() bool {
	res, err := oleutil.GetProperty(hwp.dispatch, "IsModified")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)

}

/*
문서가 빈 문서인지 여부 확인 (읽기전용)

	true: 빈문서
	false: 빈문서가 아님
*/
func (hwp *IHwpObject) IsEmpty() bool {
	res, err := oleutil.GetProperty(hwp.dispatch, "IsEmpty")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)

}

/*
현재 편집 모드 확인

	EditMode() : 현재 편집모드 리턴
	EditMode(int) : 현재 편집모드 변경
	0 : 읽기전용
	1 : 일반편집모드
	2 : 양식모드(양식 사용자 모드): Cell과 누름틀 중 양식 모드에서 편집 가능 속성을 가진 것만 편집 가능
	16: 배포용 문서 (SetEditMode로 지정 불가능)
*/
func (hwp *IHwpObject) EditMode(arg ...int) int {
	var em int
	argLen := len(arg)
	switch argLen {
	case 1:
		_, err := oleutil.PutProperty(hwp.dispatch, "EditMode", int32(arg[0]))
		if err != nil {
			panic(err)
		}
		em = arg[0]
	case 0:
		res, err := oleutil.GetProperty(hwp.dispatch, "EditMode")
		if err != nil {
			panic(err)
		}

		em = int(res.Value().(int32))
	default:
		panic("EditMode int 인자 0개 또는 1개여야함.")
	}
	return em
}

/*
문서의 내용이 어떤 Selection 상태인가를 알려준 (읽기전용)

※ 일부 미구현, 일반 블록이 아닌 F3나 F4에 의한 블록 지정 ※

	0 : 블록 되어 있지 않음
	1 : 블록 되어 있음
*/
func (hwp *IHwpObject) SelectionMode() int {
	res, err := oleutil.GetProperty(hwp.dispatch, "SelectionMode")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
캐럿이 위치한 필드의 상태 정보를 구함(읽기 전용)

	[0] : 필드의 종류: (0 : 없음, 1 : 셀, 2 : 누름틀)
	[1] : 필드명의 존재 여부 : (0 : 없음, 1 : 있음)
*/
func (hwp *IHwpObject) CurFieldState() (int, int) {
	res, err := oleutil.GetProperty(hwp.dispatch, "CurFieldState")
	if err != nil {
		panic(err)
	}

	vint := res.Value().(int32)
	fieldType := int(vint & int32(0b1111))    // bit 0~3 필드의 종류
	fieldNameExist := int(vint & int32(1<<4)) // bit 4 필드명 존재 유무
	//tmp := int(vint & int32(0xFFFFFFFF)) // bit 5~31 예약
	return fieldType, fieldNameExist
}

/*
문서 페이지 수 (읽기 전용)
*/
func (hwp *IHwpObject) PageCount() int {
	res, err := oleutil.GetProperty(hwp.dispatch, "PageCount")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
현재 선택되어 있는 표와 셀의 모양 정보를 나타냄

	※ CellShape인데 HCell이 아닌 HTable의 객체를 리턴하네??
*/
func (hwp *IHwpObject) CellShape() *HTable {
	res, err := oleutil.GetProperty(hwp.dispatch, "CellShape")
	if err != nil {
		panic(err)
	}
	var ht HTable
	ht.variant = res
	return &ht
}

/*
현재 selection의 글자 모양을 나타낸다
*/
func (hwp *IHwpObject) CharShape() *HCharShape {
	res, err := oleutil.GetProperty(hwp.dispatch, "CharShape")
	if err != nil {
		panic(err)
	}
	var cs HCharShape
	cs.variant = res
	return &cs
}

/*
문서 중 첫 번째 컨트롤(읽기 전용)
*/
func (hwp *IHwpObject) HeadCtrl() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(hwp.dispatch, "HeadCtrl")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
문서 중 마지막 컨트롤 (읽기 전용)
*/
func (hwp *IHwpObject) LastCtrl() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(hwp.dispatch, "LastCtrl")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
현재 선택되어 있는 컨트롤(읽기 전용)
*/
func (hwp *IHwpObject) CurSelectedCtrl() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(hwp.dispatch, "CurSelectedCtrl")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
현재 Selection의 문단 모양을 나타낸다.
*/
func (hwp *IHwpObject) ParaShape() *HParaShape {
	res, err := oleutil.GetProperty(hwp.dispatch, "ParaShape")
	if err != nil {
		panic(err)
	}
	var ps HParaShape
	ps.variant = res
	return &ps
}

/*
현재 캐럿의 상위 컨트롤(읽기 전용)
*/
func (hwp *IHwpObject) ParentCtrl() *IDHwpCtrlCode {
	res, err := oleutil.GetProperty(hwp.dispatch, "ParentCtrl")
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
뷰의 상태 정보
*/
func (hwp *IHwpObject) ViewProperties() *HViewProperties {
	res, err := oleutil.GetProperty(hwp.dispatch, "ViewProperties")
	if err != nil {
		panic(err)
	}
	var vp HViewProperties
	vp.variant = res
	return &vp
}

/*
문서 파일의 경로
*/
func (hwp *IHwpObject) Path() string {
	res, err := oleutil.GetProperty(hwp.dispatch, "Path")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
문서 파일을 연다.

	Open("절대경로+파일명", "HWPX 등 포멧", "세부옵션")
	포멧 및 세부옵션은 생략 가능 > "" 빈문자열로 지정됨
*/
func (hwp *IHwpObject) Open(argv ...string) {
	path, format, arg := "", "", ""
	arglen := len(argv)
	switch arglen {
	case 3:
		arg = argv[2]
		fallthrough
	case 2:
		format = argv[1]
		fallthrough
	case 1:
		path = argv[0]
	default:
		panic("Open(argv ...string) 변수 갯수 이상")
	}
	_, err := oleutil.CallMethod(hwp.dispatch, "Open", path, format, arg)
	if err != nil {
		panic(err)
	}
}

/*
현재 편집중인 문서를 저장한다.

	arg의 갯수는 없거나 1개. arg 생략시 true
	true: 문서가 변경된 경우에만 저장, false: 변경 여부에 관계없이 무조건 저장
*/
func (hwp *IHwpObject) Save(arg ...bool) bool {
	var b bool
	arglen := len(arg)
	switch arglen {
	case 0:
		b = true
	case 1:
		b = arg[0]
	default:
		panic("Open(arg ...bool) 변수 갯수 이상: 없거나 1개여야 함")
	}
	res, err := oleutil.CallMethod(hwp.dispatch, "Save", b)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
현재 편집중인 문서를 지정한 이름으로 저장한다.

	SaveAs("절대경로+파일명", "HWPX 등 포멧", "세부옵션")
	포멧 생략시 HWP 지정, 세부옵션 생략시 "" 빈문자열
*/
func (hwp *IHwpObject) SaveAs(argv ...string) bool {
	path, format, arg := "", "HWP", ""
	arglen := len(argv)
	switch arglen {
	case 3:
		arg = argv[2]
		fallthrough
	case 2:
		format = argv[1]
		fallthrough
	case 1:
		path = argv[0]
	default:
		panic("SaveAs(argv ...string) 변수 갯수 이상")
	}
	res, err := oleutil.CallMethod(hwp.dispatch, "SaveAs", path, format, arg)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
현재 캐럿 위치에 문서 파일을 삽입한다.

	Insert("절대경로+파일명", "HWPX 등 포멧", "세부옵션")
	포멧 및 세부옵션 생략시 "" 빈문자열
*/
func (hwp *IHwpObject) Insert(argv ...string) {
	path, format, arg := "", "", ""
	arglen := len(argv)
	switch arglen {
	case 3:
		arg = argv[2]
		fallthrough
	case 2:
		format = argv[1]
		fallthrough
	case 1:
		path = argv[0]
	default:
		panic("Insert(argv ...string) 변수 갯수 이상")
	}
	_, err := oleutil.CallMethod(hwp.dispatch, "Insert", path, format, arg)
	if err != nil {
		panic(err)
	}
}

/*
블록을 설정 한다.

	spara: 블록 시작 위치의 문단 번호
	spos: 블록 시작 위치의 문단 중에서 문자의 위치
	epara: 블록 끝 위치의 문단 번호
	epo: 블록 끝 위치의 문단 중에서 문자의 위치, epos가 가리키는 문자는 포함되지 않음
*/
func (hwp *IHwpObject) SelectText(spara, spos, epara, epo int) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "SelectText", spara, spos, epara, epo)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
캐럿의 현재 위치에 누름틀을 생성한다.

	direction: 누름클에 입력이 안된 상태에서 보여지는 안내문/지시문
	memo: 누름틀에 대한 설명/도움말
	name: 누름틀 필드에 대한 필드 이름
*/
func (hwp *IHwpObject) CreateField(direction, memo, name string) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "CreateField", direction, memo, name)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정한 필드로 캐럿을 이동한다.

	MoveToField(field string, text bool, start bool, select bool) bool

	field:  필드 이름. GetFieldText/PutFieldText과 같은 형식으로 이름 뒤에 ‘{{#}}’로 번호를 지정할 수 있다.
	text : 필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True) 누름틀 코드로 이동할지(False)를 지정한다. 누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.
	start : 필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다. select를 True로 지정하면 무시된다. 생략하면 True가 지정된다.
	select : 필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다. 생략하면 False가 지정된다.
	field 외 생략가능, 생략시 text = true, start = true, select = false
*/
func (hwp *IHwpObject) MoveToField(field string, arg ...bool) bool {
	btext, bstart, bselect := true, true, false
	arglen := len(arg)
	switch arglen {
	case 3:
		bselect = arg[2]
		fallthrough
	case 2:
		bstart = arg[1]
		fallthrough
	case 1:
		btext = arg[0]
		fallthrough
	case 0:
		res, err := oleutil.CallMethod(hwp.dispatch, "MoveTofield", field, btext, bstart, bselect)
		if err != nil {
			panic(err)
		}

		return res.Value().(bool)
	default:
		panic("Insert(field string, arg ...bool) 변수 갯수 이상")
	}
}

/*
문서 중에 지정한 데이터 필드가 존재하는지 검사한다.

	field:  필드 이름.
*/
func (hwp *IHwpObject) FieldExist(field string) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "FieldExist", field)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정한 필드에서 문자열을 구한다.

	fieldlist: 필드 이름의 리스트
	한번에 여러 필드를 불러올 때 "\x02"로 구분: "필드이름1\x02"필드이름2"
	문서에 동일 필드 이름 2개 이상 존재할 때 "{{n}}": 필드이름{{n}}
*/
func (hwp *IHwpObject) GetFieldText(fieldlist string) string {
	res, err := oleutil.CallMethod(hwp.dispatch, "GetFieldText", fieldlist)
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
지정한 필드의 내용을 채운다.

	fieldlist: 필드 이름의 리스트, 동일 필드명 2개 이상인 경우 "필드이름{{n}}\x02필드이름{{n+x}}"
	첫번째 번호는 0에서 시작
	textlist: 필드 이름의 리스트
*/
func (hwp *IHwpObject) PutFieldText(fieldlist, textlist string) {
	_, err := oleutil.CallMethod(hwp.dispatch, "PutFieldText", fieldlist, textlist)
	if err != nil {
		panic(err)
	}
}

/*
지정한 필드의 이름을 바꾼다.

	oldname: 이름을 바꿀 필드 이름의 리스트. 형식은 PutFieldText과 동일
	newname: 새로운 필드 이름의 리스트. oldname과 동일한 개수의 필드이름을 \x02로 구분하여 지정
*/
func (hwp *IHwpObject) RenameField(oldname, newname string) {
	_, err := oleutil.CallMethod(hwp.dispatch, "RenameField", oldname, newname)
	if err != nil {
		panic(err)
	}
}

/*
현재 캐럿 위치의 데이터 필드 이름을 구한다.

	0: 0을 지정하면 모두 off, 생략하면 0이 지정
	1: 셀에 부여된 필드 리스트만을 구한다.
	2: 누름틀에 부여된 필드 리스트만을 구한다.
	GetFieldList()의 option 중에 hwpFieldSelection (= 4) 옵션은 사용하지 않는다.
*/
func (hwp *IHwpObject) GetCurFieldName(hwpFieldOption ...int) string {
	opt := 0
	switch len(hwpFieldOption) {
	case 1:
		opt = hwpFieldOption[0]
		fallthrough
	case 0:
		res, err := oleutil.CallMethod(hwp.dispatch, "GetCurFieldName", opt)
		if err != nil {
			panic(err)
		}

		return res.Value().(string)
	default:
		panic("GetCurFieldName(hwpFieldOption ...int)의 인자는 1개이거나 없어야함")
	}
}

/*
현재 캐럿 위치의 데이터 필드 이름을 설정한다.

	opt
		0: 0을 지정하면 모두 off
		1: 셀에 부여된 필드 리스트만을 구한다.
		2: 누름틀에 부여된 필드 리스트만을 구한다.
	direction: 누름틀 필드의 안내문. 누름틀 필드일 때만 유효
	memo: 누름틀 필드의 메모. 누름틀 필드일 때만 유효
	GetFieldList()의 option 중에 hwpFieldSelection (= 4) 옵션은 사용하지 않는다.
*/
func (hwp *IHwpObject) SetCurFieldName(fieldname string, opt int, direction, memo string) bool {

	res, err := oleutil.CallMethod(hwp.dispatch, "SetCurFieldName", fieldname, opt, direction, memo)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
지정한 필드의 속성을 바꾼다.

	※ 어떻게 작동한다는 것인지 이해가 되지 않음.. ※
	field: 속성을 바꿀 필드 이름의 리스트. 형식은 PutFieldText과 동일.
	remove: 제거될 속성
	add: 추가될 속성
	속성: (0: 편집 불가, 1: 편집 가능), remove, add 둘다 0이면 현재 속성을 리턴
	음수가 리턴되면 에러.
*/
func (hwp *IHwpObject) ModifyFieldProperties(field string, remove, add int) int {
	res, err := oleutil.CallMethod(hwp.dispatch, "ModifyFieldProperties", field, remove, add)
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
양식 모드와 읽기 전용모드일때 현재 열린 문서의 필드의 겉보기 속성 (『』표시)을 바꾼다

	※ 에러 없이 설정 속성 값이 리턴되지만 열린 문서에서는 겉보기 속성이 변하는 것을 확인할 수 없음
	1: 『』을 표시하지 않음
	2: 누름틀『』을 빨간색, 개인정보/문서요약/날짜시간/파일경로 『』을 흰색으로 표시
	3: 『』을 흰색으로 표시

	설정된 속성이 리턴, 0이 리턴되면 에러
*/
func (hwp *IHwpObject) SetFieldViewOption(option int) int {
	res, err := oleutil.CallMethod(hwp.dispatch, "SetFieldViewOption", option)
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
문서 중의 필드 리스트를 구한다.

	GetFieldList(number, option int) string

	number: 문서 중 동일 이름의 필드가 다수일 때 구별하는 식별 방법 지정
		0: 아무 기호 없이 순서대로 필드 이름 나열
		1: 필드 이름 뒤에 일련번호 {{#}}
		2: 필드 이름 뒤에 동일 필드 갯수 {{#}}
	option: 다음 옵션을 조합 (생략시 0, 모두 OFF)
		1: 셀에 부여된 필드 리스트만
		2: 누름틀에 부여된 필드 리스트만
		4: 셀렉션 내에 존재하는 필드 리스트 (※주의! 테스트 상 작동하지 않음)
*/
func (hwp *IHwpObject) GetFieldList(arg ...int) string {
	number, option := 0, 0

	switch len(arg) {
	case 2:
		option = arg[1]
		fallthrough
	case 1:
		number = arg[0]
		fallthrough
	case 0:
		res, err := oleutil.CallMethod(hwp.dispatch, "GetFieldList", number, option)
		if err != nil {
			panic(err)
		}

		return res.Value().(string)
	default:
		panic("GetFieldList(number, option int) 함수 인자 갯수 0,1,2개 외 에러")
	}
}

/*
캐럿의 위치를 옮긴다.

	MovePos(moveID, para, pos int) bool

	moveID: 0 루트 리스트, 1 현재 리스트, 2 문서 시작, 3 문서 끝, 등등
	para: 문단 번호
	pos: 문자의 위치
*/
func (hwp *IHwpObject) MovePos(arg ...int) bool {
	moveID, para, pos := 1, 0, 0
	switch len(arg) {
	case 3:
		pos = arg[2]
		fallthrough
	case 2:
		para = arg[1]
		fallthrough
	case 1:
		moveID = arg[0]
		fallthrough
	case 0:
		res, err := oleutil.CallMethod(hwp.dispatch, "MovePos", moveID, para, pos)
		if err != nil {
			panic(err)
		}

		return res.Value().(bool)
	default:
		panic("MovePos(moveID, para, pos) 함수 인자 갯수 0,1,2, 3개 외 에러")
	}
}

/*
문서 검색을 위한 준비 작업을 한다.

	InitScan(option, rang, spara, spos, epara, epos int) bool

	option: 찾을 대상 (컨트롤은 조합 가능)
		0x00: 본문 대상,
		0x01: char 컨트롤(강제줄나눔, 문단끝, 하이픈, 묶음빈칸 ...)
		0x02: inline 컨트롤(누름틀 필드 끝, ..)
		0x04: extende 컨트롤(바탕쪽, 프리젠테이션, 다단, 누름틀 필드 시작, Shape Object등)
	rang: 검색 범위 (조합가능)
		0x0000: 캐럿 위치부터
		0x00?0: 시작위치, 0x000?: 끝 위치
		?: 1 특정, 2 줄, 3 문단, 4 구역, 5 리스트, 6 컨트롤, 7 문서
		0x0ff: 블록으로 제한, 0x0100: 역방향
	spara, spos, epara, epos는 특정(1) rang에 유효 시작s/끝e, 문단para/문자pos
*/
func (hwp *IHwpObject) InitScan(arg ...int) bool {
	option, rang, spara, spos, epara, epos := 0x01|0x02|0x04, 0x0070|0x0007, 0, 0, 0, 0

	switch len(arg) {
	case 6:
		epos = arg[5]
		fallthrough
	case 5:
		epara = arg[4]
		fallthrough
	case 4:
		spos = arg[3]
		fallthrough
	case 3:
		spara = arg[2]
		fallthrough
	case 2:
		rang = arg[1]
		fallthrough
	case 1:
		option = arg[0]
		fallthrough
	case 0:
		res, err := oleutil.CallMethod(hwp.dispatch, "InitScan", option, rang, spara, spos, epara, epos)
		if err != nil {
			panic(err)
		}

		return res.Value().(bool)
	default:
		panic("InitScan(option, rang, spara, spos, epara, epos) 함수 인자 갯수 에러")
	}
}

/*
InitScan()으로 설정된 정보를 초기화 한다.
*/
func (hwp *IHwpObject) ReleaseScan() {
	_, err := oleutil.CallMethod(hwp.dispatch, "ReleaseScan")
	if err != nil {
		panic(err)
	}
}

/*
문서 중에서 텍스트를 얻는다.

	텍스트에서 tab은 '\0x9', 문단 바뀜 CR/LF(0x0D/0X0A)	return 정수 값
	0: 텍스트정보없음, 1: 리스트 끝, 2: 일반텍스트, 3: 다음문단,
	4: 제어문자 내부로 들어감, 5: 제어문자를 빠져나옴
	101: InitScan 안됨, 102: 텍스트 변환 실패
*/
func (hwp *IHwpObject) GetText(text *string) int {
	res, err := oleutil.CallMethod(hwp.dispatch, "GetText", text)
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
현재 캐럿의 위치 정보를 얻어온다.

	list: 리스트 ID, para: 문단 ID, pos: 문단 내 글자 위치
*/
func (hwp *IHwpObject) GetPos(list, para, pos *int) {
	_, err := oleutil.CallMethod(hwp.dispatch, "GetPos", list, para, pos)
	if err != nil {
		panic(err)
	}
}

/*
캐럿을 문서 내 특정 위치로 위치시킨다.

	list: 리스트 ID, para: 문단 ID, pos: 문단 내 글자 위치
*/
func (hwp *IHwpObject) SetPos(list, para, pos int) {
	_, err := oleutil.CallMethod(hwp.dispatch, "SetPos", list, para, pos)
	if err != nil {
		panic(err)
	}
}

/*
상태바에 나타날 정보를 알아낸다.

	seccnt: 총 구역, secno: 현재 구역, prnpageno: 쪽, colno: 단
	line: 줄, pos: 칸, over: (0: 삽입, 1: 수정), ctrlname
*/
func (hwp *IHwpObject) KeyIndicator(seccnt, secno, prnpageno, colno, line, pos *int, over *int16, ctrlname *string) {
	_, err := oleutil.CallMethod(hwp.dispatch, "KeyIndicator",
		seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)
	if err != nil {
		panic(err)
	}
}

/*
지정한 페이지의 이미지를 파일로 생성한다.

	CreatePageImage(path string, pgno, resolution, depth int, format string)

	path: 생성할 이미지 파일의 경로
	pgno: 페이지 번호, 0부터 PageCount-1 까지.
	resolution: 이미지 해상도. DPI 단위 96, 300, 1200 중 하나
	depth: 이미지 파일의 color depth: 1,4,8,24 중 하나
	format: 이미지 파일의 포맷. bmp, gif 중 하나
*/
func (hwp *IHwpObject) CreatePageImage(path string, arg ...interface{}) bool {
	pgno, resolution, depth := 0, 96, 24
	format := "bmp"
	lenArg := len(arg)
	switch lenArg {
	case 4:
		format = arg[3].(string)
		fallthrough
	case 3:
		depth = arg[2].(int)
		fallthrough
	case 2:
		resolution = arg[1].(int)
	case 1:
		pgno = arg[0].(int)
		fallthrough
	case 0:
	default:
		panic("CreatePageImage 함수 인자 갯수 이상")
	}
	res, err := oleutil.CallMethod(hwp.dispatch, "CreatePageImage", path, pgno, resolution, depth, format)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
액션을 실행한다.

	action: 액션 ID

ParameterSet 없이 할 수 있는 모든 액션 실행
*/
func (hwp *IHwpObject) Run(action string) {
	_, err := oleutil.CallMethod(hwp.dispatch, "Run", action)
	if err != nil {
		panic(err)
	}
}

/*
특정 액션이 실행되지 않도록 잠근다.

	actionID: 액션 ID
	lock: true는 lock, false는 unlock
*/
func (hwp *IHwpObject) LockCommand(actionID string, lock bool) {
	_, err := oleutil.CallMethod(hwp.dispatch, "LockCommand", actionID, lock)
	if err != nil {
		panic(err)
	}
}

/*
특정 액션이 잠금 상태인지 여부를 조사한다.

	actionID: 액션 ID
*/
func (hwp *IHwpObject) IsCommandLock(actionID string) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "IsCommandLock", actionID)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
현재 캐럿의 위치에 그림을 삽입한다.

	InsertPicture(path string, embeded bool, sizeoption int, reverse, watermark bool, effect, width, height int) IDHwpCtrlCode

	path: 삽입할 이미지 파일, URL 사용 가능
	embeded: 이미지 파일을 문서내에 포함할지 여부 (True/False).
	sizeoption: 0: 원본 크기, 1: 지정한 크기, 2: 셀의 크기 맞추기, 3: 원본 비율 유지 셀의 크기 맞추기
	reverse: 이미지의 반전 유무 (True/False)
	watermark: watermark효과 유무 (True/False)
	effect: 0: 원본 이미지, 1: 그레이스케일, 2: 흑백효과
	width : 그림의 가로 크기 지정. 단위는 mm
	height : 그림의 높이 크기 지정. 단위는 mm
*/
func (hwp *IHwpObject) InsertPicture(path string, arg ...interface{}) *IDHwpCtrlCode {
	embeded := true
	sizeoption := 0
	reverse, watermark := false, false
	effect, width, height := 0, 50, 50
	lenArg := len(arg)
	switch lenArg {
	case 7:
		height = arg[6].(int)
		fallthrough
	case 6:
		width = arg[5].(int)
		fallthrough
	case 5:
		effect = arg[4].(int)
		fallthrough
	case 4:
		watermark = arg[3].(bool)
		fallthrough
	case 3:
		reverse = arg[2].(bool)
		fallthrough
	case 2:
		sizeoption = arg[1].(int)
		fallthrough
	case 1:
		embeded = arg[0].(bool)
	case 0:
	default:
		panic("InsertPicture 함수 인자 갯수 이상")
	}
	res, err := oleutil.CallMethod(hwp.dispatch, "InsertPicture",
		path, embeded, int16(sizeoption), reverse, watermark, int16(effect), width, height)
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
(셀?) 배경이미지를 넣는다.

	InsertBackgroundPicture(bordertype, path string, embeded bool, filloption int, watermark bool, effect, brightness, contrast int) bool

	bordertype: "SelectedCell": 현재 셀의 배경 변경, "SelectedCellDelete": 현재 셀의 배경 지우기(셀선택되어있어야함)
	path : 삽입할 이미지 파일, URL 사용 가능
	embeded : 이미지 파일을 문서내에 포함할지 여부 (True/False)
	filloption : 삽입할 그림의 크기를 지정하는 옵션 0~14, 0 바둑판식 모두, 5 크기에 맞추어 등
	effect: 0: 원본 이미지, 1: 그레이스케일, 2: 흑백효과
	watermark : watermark효과 유무
	brightness : 밝기 지정(-100 ~ 100), 기본값 : 0
	contrast : 선명도 지정(-100 ~ 100), 기본값 : 0
*/
func (hwp *IHwpObject) InsertBackgroundPicture(bordertype, path string, arg ...interface{}) bool {
	embeded := true
	filloption := 5
	watermartk := false
	effect, brightness, contrast := 0, 0, 0
	lenArg := len(arg)
	switch lenArg {
	case 6:
		contrast = arg[5].(int)
		fallthrough
	case 5:
		brightness = arg[4].(int)
		fallthrough
	case 4:
		effect = arg[3].(int)
		fallthrough
	case 3:
		watermartk = arg[2].(bool)
		fallthrough
	case 2:
		filloption = arg[1].(int)
		fallthrough
	case 1:
		embeded = arg[0].(bool)
	case 0:
	default:
		panic("InsertBackgroundPicture 함수 인자 갯수 이상")
	}
	res, err := oleutil.CallMethod(hwp.dispatch, "InsertBackgroundPicture",
		bordertype, path, embeded, filloption, watermartk, effect, brightness, contrast)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
액션 오브젝트를 생성한다.

	action: 액션 ID
*/
func (hwp *IHwpObject) CreateAction(action string) *IDHwpAction {
	res, err := oleutil.CallMethod(hwp.dispatch, "CreateAction", action)
	if err != nil {
		panic(err)
	}
	var dha IDHwpAction
	dha.variant = res
	return &dha
}

/*
현재 캐럿 위치에 컨트롤을 삽입한다.

	※ 일부 미구현: InsertCtrl(ctrlid string, initparam ParameterSet) Ctrl

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
func (hwp *IHwpObject) InsertCtrl(ctrlid string) *IDHwpCtrlCode {
	res, err := oleutil.CallMethod(hwp.dispatch, "InsertCtrl", ctrlid)
	if err != nil {
		panic(err)
	}
	var hcc IDHwpCtrlCode
	hcc.variant = res
	return &hcc
}

/*
문서내 컨트롤을 삭제한다.

	DeleteCtrl(HWPCtrlCode ctrl)
*/
func (hwp *IHwpObject) DeleteCtrl(ctrl *IDHwpCtrlCode) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "DeleteCtrl", ctrl.variant)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
마우스의 현재 위치를 구한다

	Xrelto: x좌표계의 기준 위치 0: 종이 기준, 1:쪽 기준
	Yrelto: y좌표계의 기준 위치 0: 종이 기준, 1:쪽 기준
	return *ole.IDispatch이나 "MousePos" parameterset
	SetID: XRelTo, YRelTo, Page, X, Y
	X,Y 좌표는 HWPUNIT, Item으로 불러오면 됨
*/
func (hwp *IHwpObject) GetMousePos(Xrelto, Yrelto int) *HMousePos {
	res, err := oleutil.CallMethod(hwp.dispatch, "GetMousePos", Xrelto, Yrelto)
	if err != nil {
		panic(err)
	}
	var mp HMousePos
	mp.variant = res
	return &mp
}

/*
현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

	hwpSaveOption
	0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다.
	1: 문서 내용을 버린다.
	2: 문서가 변경된 경우 저장한다.
	3: 무조건 저장한다.
*/
func (hwp *IHwpObject) Clear(hwpSaveOption ...int) {
	op := int16(0)
	lenOp := len(hwpSaveOption)
	switch lenOp {
	case 1:
		op = int16(hwpSaveOption[0])
	}
	_, err := oleutil.CallMethod(hwp.dispatch, "Clear", op)
	if err != nil {
		panic(err)
	}
}

/*
보안 모듈 관련 문서 참조

	※ 보안승인 메시지가 나타나지 않도록 처리하는 모듈 외 용도를 못 찾음
	module: "FilePathCheckDLL"
	fpcme: 레지스트리에 등록된 "FilePathCheckerModuleExample.dll"값의 문자열 이름
*/
func (hwp *IHwpObject) RegisterModule(module, fpcme string) bool {
	res, err := oleutil.CallMethod(hwp.dispatch, "RegisterModule", module, fpcme)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
최상위 오브젝트 (IHwpObject interface)를 얻는다.
*/
func (hwp *IHwpObject) Application() *IHwpObject {
	res, err := oleutil.GetProperty(hwp.dispatch, "Application")
	if err != nil {
		panic(err)
	}

	var ho IHwpObject
	ho.variant = res
	ho.dispatch = res.Value().(*ole.IDispatch)
	return &ho
}

/*
메시지 박스를 제어하는 XHwpMessageBox Object를 얻는다.
*/
func (hwp *IHwpObject) XHwpMessageBox() *IXHwpMessageBox {
	res, err := oleutil.GetProperty(hwp.dispatch, "XHwpMessageBox")
	if err != nil {
		panic(err)
	}
	var mb IXHwpMessageBox
	mb.variant = res
	return &mb
}

/*
도큐먼트를 관리하는 XHwpDocuments Object를 얻는다.
*/
func (hwp *IHwpObject) XHwpDocuments() *IXHwpDocuments {
	res, err := oleutil.GetProperty(hwp.dispatch, "XHwpDocuments")
	if err != nil {
		panic(err)
	}
	var hds IXHwpDocuments
	hds.variant = res
	return &hds
}

/*
윈도우를 관리하는 XHwpWindows Object를 얻는다.
*/
func (hwp *IHwpObject) XHwpWindows() *IXHwpWindows {
	res, err := oleutil.GetProperty(hwp.dispatch, "XHwpWindows")
	if err != nil {
		panic(err)
	}
	var hw IXHwpWindows
	hw.variant = res
	return &hw
}

/*
파라메터셋 오브젝트를 관리하는 HParameterSet Object를 얻는다.
*/
func (hwp *IHwpObject) HParameterSet() *HParameterSet {
	res, err := oleutil.GetProperty(hwp.dispatch, "HParameterSet")
	if err != nil {
		panic(err)
	}
	var ps HParameterSet
	ps.variant = res
	return &ps
}

/*
Action을 제어하는 HAction Object를 얻는다.
*/
func (hwp *IHwpObject) HAction() *HAction {
	res, err := oleutil.GetProperty(hwp.dispatch, "HAction")
	if err != nil {
		panic(err)
	}
	var ac HAction
	ac.variant = res
	return &ac
}

/*
ODBC로 제어할 수 있는 Object를 얻는다.

	※ 미구현 *ole.IDispatch 리턴
*/
func (hwp *IHwpObject) XHwpODBC() *ole.IDispatch {
	res, err := oleutil.GetProperty(hwp.dispatch, "XHwpODBC")
	if err != nil {
		panic(err)
	}
	return res.Value().(*ole.IDispatch)
}

/*
한글 과 한글 OCX의 버젼 정보를 구한다.읽기 전용.

	※ 메뉴얼에는 아래와 같지만, 실질적으론 "#,#,#,#" string을 리턴하고 있다. #은 숫자.
	byte 3 = 한글의 major version.
	byte 2 = 한글의 minor version.
	byte 1 = 한글 OCX의 major version.
	byte 0 = 한글 OCX의 minor version.
*/
func (hwp *IHwpObject) Version() string {
	res, err := oleutil.GetProperty(hwp.dispatch, "Version")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
HStyleTemplate 파라메터셋 오브젝트에 지정된 스타일을 Export한다.

	※ 미구현.
	BOOL ExportStyle(LPDISPATCH param)
*/
func (hwp *IHwpObject) ExportStyle() {
}

/*
HStyleTemplate 파라메터셋 오브젝트에 지정된 스타일을 Import한다.

	※ 미구현.
	BOOL ImportStyle(LPDISPATCH param)
*/
func (hwp *IHwpObject) ImportStyle() {
}

/*
현재 캐럿의 위치에서 Ctrl을 찾는다.
*/
func (hwp *IHwpObject) FindCtrl() string {
	res, err := oleutil.CallMethod(hwp.dispatch, "FindCtrl")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
현재 Select된 Ctrl의 Selection을 해제한다.
*/
func (hwp *IHwpObject) UnSelectCtrl() {
	_, err := oleutil.CallMethod(hwp.dispatch, "UnSelectCtrl")
	if err != nil {
		panic(err)
	}
}

/*
mm to HWPUNIT
*/
func (hwp *IHwpObject) MiliToHwpUnit(v float32) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "MiliToHwpUnit", v)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
ViewGridOption 액션 및 HGridInfo 파라미터 셋 관련 ViewLine 설정 값

	s: "All", "Vert", "Horz"
*/
func (hwp *IHwpObject) GridViewLine(s string) uint16 {
	res, err := oleutil.CallMethod(hwp.dispatch, "GridViewLine", s)
	if err != nil {
		panic(err)
	}
	return uint16(res.Value().(int32))
}

/*
HCaption 파라미터 셋 관련 Side 설정 값

	s: "Left", "Right", "Top", "Bottom"
*/
func (hwp *IHwpObject) SideType(s string) uint16 {
	res, err := oleutil.CallMethod(hwp.dispatch, "SideType", s)
	if err != nil {
		panic(err)
	}
	return uint16(res.Value().(int32))
}

/*
ShapeDrawLineAttr, HDrawLineAttr 의 선 관련 옵션

	s: "LargeLarge" 등
*/
func (hwp *IHwpObject) EndSize(s string) uint16 {
	res, err := oleutil.CallMethod(hwp.dispatch, "EndSize", s)
	if err != nil {
		panic(err)
	}
	return uint16(res.Value().(int32))
}

/*
ShapeDrawLineAttr, HDrawLineAttr 의 선 관련 옵션

	s: "Box", "Diamond" 등
*/
func (hwp *IHwpObject) EndStyle(s string) uint16 {
	res, err := oleutil.CallMethod(hwp.dispatch, "EndStyle", s)
	if err != nil {
		panic(err)
	}
	return uint16(res.Value().(int32))
}

/*
ShapeDrawLineAttr, HDrawLineAttr 의 선 관련 옵션

	s: "DashDot" 등
*/
func (hwp *IHwpObject) HwpLineType(s string) uint16 {
	res, err := oleutil.CallMethod(hwp.dispatch, "HwpLineType", s)
	if err != nil {
		panic(err)
	}
	return uint16(res.Value().(int32))
}

/*
RGBColor
*/
func (hwp *IHwpObject) RGBColor(r, g, b int32) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "RGBColor", r, g, b)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
PointToHwpUnit
*/
func (hwp *IHwpObject) PointToHwpUnit(p int32) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "PointToHwpUnit", p)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
BrushType
*/
func (hwp *IHwpObject) BrushType(s string) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "BrushType", s)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
HatchStyle
*/
func (hwp *IHwpObject) HatchStyle(s string) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "HatchStyle", s)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
HwpLineWidth

	ex> HwpLineWidth("3.0mm")
*/
func (hwp *IHwpObject) HwpLineWidth(s string) int32 {
	res, err := oleutil.CallMethod(hwp.dispatch, "HwpLineWidth", s)
	if err != nil {
		panic(err)
	}
	return res.Value().(int32)
}

/*
GetHeadingString

	글머리표/문단번호/개요번호를 추출한다.
*/
func (hwp *IHwpObject) GetHeadingString() string {
	res, err := oleutil.CallMethod(hwp.dispatch, "GetHeadingString")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

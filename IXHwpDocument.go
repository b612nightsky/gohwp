package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
※ 미구현 (일부 읽기전용 오브젝트 미구현)
IXHwpDocument:

	IXHwpDocument 도큐먼트 오브젝트
	(Document 개체 - 사용 편이를 위해 제공됨)
*/
type IXHwpDocument struct {
	variant *ole.VARIANT
}

/*
최상위 오브젝트를 얻어옴 (IHwpObject - 읽기 전용)
*/
func (p *IXHwpDocument) Application() *IHwpObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Application")
	if err != nil {
		panic(err)
	}
	var ho IHwpObject
	ho.dispatch = res.Value().(*ole.IDispatch)
	return &ho
}

/*
도큐먼트의 Path를 얻어옴(읽기 전용)
*/
func (p *IXHwpDocument) Path() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Path")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
도큐먼트의 전체 경로를 얻어옴(읽기 전용)
*/
func (p *IXHwpDocument) FullName() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "FullName")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
도큐먼트의 에디트 모드를 설정하거나/얻어옴

	읽기모드: 0, 쓰기모드: 1
	EditMode() : int값으로 리턴
	EdutMode(0) : 읽기모드로 설정하고 int값으로 리턴
	EdutMode(1) : 쓰기모드로 설정하고 int값으로 리턴
*/
func (p *IXHwpDocument) EditMode(arg ...int) int {
	var isEdit int
	lenArg := len(arg)
	switch lenArg {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "EditMode", int32(arg[0]))
		if err != nil {
			panic(err)
		}
		isEdit = arg[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "EditMode")
		if err != nil {
			panic(err)
		}
		isEdit = int(res.Value().(int32))
	default:
		panic("EditMode 함수 인자 갯수 이상")
	}
	return isEdit
}

/*
도큐먼트의 변경 여부를 설정하거나/얻어옴

	미변경: 0, 변경: 1
	Modified() : int값으로 리턴
	Modified(0) : 미변경으로 설정하고 int값으로 리턴
	Modified(1) : 변경으로 설정하고 int값으로 리턴
*/
func (p *IXHwpDocument) Modified(arg ...int) int {
	var isModified int
	lenArg := len(arg)
	switch lenArg {
	case 1:
		_, err := oleutil.PutProperty(p.variant.ToIDispatch(), "Modified", int32(arg[0]))
		if err != nil {
			panic(err)
		}
		isModified = arg[0]
	case 0:
		res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Modified")
		if err != nil {
			panic(err)
		}
		isModified = int(res.Value().(int32))
	default:
		panic("Modified 함수 인자 갯수 이상")
	}
	return isModified
}

/*
도큐먼트의 저장된 포맷을 얻어옴(읽기 전용)
*/
func (p *IXHwpDocument) Format() string {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "Format")
	if err != nil {
		panic(err)
	}
	return res.Value().(string)
}

/*
도큐먼트의 패스워드를 설정(쓰기 전용)

	※ 미구현. IXHwpDocument 객체에서 Password란게 없는 것처럼 오류가 남
*/
func (p *IXHwpDocument) Password() {
}

/*
IXHwpSummaryInfo 문서 요약 정보 오브젝트 (읽기 전용)
*/
func (p *IXHwpDocument) XHwpSummaryInfo() *IXHwpSummaryInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "XHwpSummaryInfo")
	if err != nil {
		panic(err)
	}
	var hsi IXHwpSummaryInfo
	hsi.variant = res
	return &hsi
}

/*
IXHwpDocumentInfo 문서 정보 오브젝트 (읽기 전용)

	※ 미구현, 리턴되는 값이 어떤 object인지 모르겠음.
	IXHwpDocumentInfo에 대한 설명이 메뉴얼에 없음. HDocumentInfo인가 싶어도 해당하는 아이템들이 없음.
*/
func (p *IXHwpDocument) XHwpDocumentInfo() {
}

/*
IXHwpPrint 프린트 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpPrint() {
}

/*
IXHwpRange Range 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpRange() {
}

/*
IXHwpFind 찾기 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFind() {
}

/*
IXHwpSelection 블록 선택 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpSelection() {
}

/*
IXHwpFormPushButtons 양식개체 푸쉬버튼을 관리하는 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFormPushButtons() {
}

/*
IXHwpFormCheckButtons 양식개체 체크박스를 관리하는 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFormCheckButtons() {
}

/*
IXHwpFormRadioButtons 양식개체 라디오버튼을 관리하는 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFormRadioButtons() {
}

/*
IXHwpFormComboBoxs 양식개체 콤보박스를 관리하는 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFormComboBoxs() {
}

/*
IXHwpFormEdits 양식개체 에디트를 관리하는 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpFormEdits() {
}

/*
IXHwpCharacterShape 글자 모양 속성 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpCharacterShape() {
}

/*
IXHwpParagraphShape 문단 모양 속성 오브젝트(읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpParagraphShape() {
}

/*
IXHwpSendMail 메일 보내기 오브젝트 (읽기 전용)

	※ 미구현
*/
func (p *IXHwpDocument) XHwpSendMail() {
}

/*
도큐먼트의 고유 ID(읽기 전용)
*/
func (p *IXHwpDocument) DocumentID() int {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "DocumentID")
	if err != nil {
		panic(err)
	}
	return int(res.Value().(int32))
}

/*
문서를 닫는다.

	isDirty
	true: 변경된 문서는 닫지 않는다
	false: 변경된 문서도 닫는다
*/
func (p *IXHwpDocument) Close(isDirty bool) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Close", isDirty)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
문서를 저장한다.

	save_if_dirty의 갯수는 없거나 1개. save_if_dirty 생략시 true
	true: 문서가 변경된 경우에만 저장, false: 변경 여부에 관계없이 무조건 저장
	문서의 경로가 지정되어 있지 않으면 "새이름으로 저장" 대화상자가 떠서 사용자에게 경로를 묻는다.
*/
func (p *IXHwpDocument) Save(save_if_dirty ...bool) bool {
	var b bool
	arglen := len(save_if_dirty)
	switch arglen {
	case 0:
		b = true
	case 1:
		b = save_if_dirty[0]
	default:
		panic("Open(save_if_dirty ...bool) 변수 갯수 이상: 없거나 1개여야 함")
	}
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Save", b)
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
func (p *IXHwpDocument) SaveAs(argv ...string) bool {
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
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SaveAs", path, format, arg)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
문서에 기록된 Undo Item을 실행한다.

	count : 아이템의 count까지 Undo한다.
*/
func (p *IXHwpDocument) Undo(count int) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Undo", count)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
문서에 기록된 Redo Item을 실행한다.

	count : 아이템의 count까지 Undo한다.
*/
func (p *IXHwpDocument) Redo(count int) bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Redo", count)
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
문서 파일을 연다.

	Open("절대경로+파일명", "HWPX 등 포멧", "세부옵션")
	포멧 및 세부옵션은 생략 가능 > "" 빈문자열로 지정됨
*/
func (p *IXHwpDocument) Open(argv ...string) {
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
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Open", path, format, arg)
	if err != nil {
		panic(err)
	}
}

/*
문서를 브라우저로 내보내기 기능
*/
func (p *IXHwpDocument) SendBrowser() bool {
	res, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SendBrowser")
	if err != nil {
		panic(err)
	}
	return res.Value().(bool)
}

/*
문서를 활성화 상태로 하기
*/
func (p *IXHwpDocument) SetActive_XHwpDocument() {
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "SetActive_XHwpDocument")
	if err != nil {
		panic(err)
	}
}

/*
문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

	option
	0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다.
	1: 문서 내용을 버린다.
	2: 문서가 변경된 경우 저장한다.
	3: 무조건 저장한다.
*/
func (p *IXHwpDocument) Clear(hwpSaveOption ...int) {
	op := int16(0)
	lenOp := len(hwpSaveOption)
	switch lenOp {
	case 1:
		op = int16(hwpSaveOption[0])
	}
	_, err := oleutil.CallMethod(p.variant.ToIDispatch(), "Clear", op)
	if err != nil {
		panic(err)
	}
}

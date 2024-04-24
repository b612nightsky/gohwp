package gohwp

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

/*
HParameterSet::

	HParameterSet: 파라미터 셋 오브젝트를 관리하는 오브젝트
*/
type HParameterSet struct {
	variant *ole.VARIANT
}

/*
구역
*/
func (p *HParameterSet) HSecDef() *HSecDef {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSecDef")
	if err != nil {
		panic(err)
	}
	var para HSecDef
	para.variant = res
	return &para
}

/*
구역 내 용지 설정 서브셋 ID
*/
func (p *HParameterSet) HPageDef() *HPageDef {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPageDef")
	if err != nil {
		panic(err)
	}
	var para HPageDef
	para.variant = res
	return &para
}

/*
다단의 set ID
*/
func (p *HParameterSet) HColDef() *HColDef {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HColDef")
	if err != nil {
		panic(err)
	}
	var para HColDef
	para.variant = res
	return &para
}

/*
글자모양의 set ID
*/
func (p *HParameterSet) HCharShape() *HCharShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCharShape")
	if err != nil {
		panic(err)
	}
	var para HCharShape
	para.variant = res
	return &para
}

/*
뷰 속성의 set ID
*/
func (p *HParameterSet) HViewProperties() *HViewProperties {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HViewProperties")
	if err != nil {
		panic(err)
	}
	var para HViewProperties
	para.variant = res
	return &para
}

/*
어플리케이션 글로벌 상태 정보
*/
func (p *HParameterSet) HAppState() *HAppState {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HAppState")
	if err != nil {
		panic(err)
	}
	var para HAppState
	para.variant = res
	return &para
}

/*
표 속성
*/
func (p *HParameterSet) HShapeObject() *HShapeObject {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HShapeObject")
	if err != nil {
		panic(err)
	}
	var para HShapeObject
	para.variant = res
	return &para
}

/*
Print
*/
func (p *HParameterSet) HPrint() *HPrint {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPrint")
	if err != nil {
		panic(err)
	}
	var para HPrint
	para.variant = res
	return &para
}

/*
탭 정의의 set ID
*/
func (p *HParameterSet) HTabDef() *HTabDef {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTabDef")
	if err != nil {
		panic(err)
	}
	var para HTabDef
	para.variant = res
	return &para
}

/*
문단 모양의 set ID
*/
func (p *HParameterSet) HParaShape() *HParaShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HParaShape")
	if err != nil {
		panic(err)
	}

	var ps HParaShape
	ps.variant = res
	return &ps
}

/*
구역 내 각주/미주 모양 서브셋 ID
*/
func (p *HParameterSet) HFootnoteShape() *HFootnoteShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFootnoteShape")
	if err != nil {
		panic(err)
	}
	var para HFootnoteShape
	para.variant = res
	return &para
}

/*
번호 넣기/새 번호로
*/
func (p *HParameterSet) HAutoNum() *HAutoNum {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HAutoNum")
	if err != nil {
		panic(err)
	}
	var para HAutoNum
	para.variant = res
	return &para
}

/*
페이지번호 제어 (97의 홀수 쪽에서 시작)
*/
func (p *HParameterSet) HPageNumCtrl() *HPageNumCtrl {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPageNumCtrl")
	if err != nil {
		panic(err)
	}
	var para HPageNumCtrl
	para.variant = res
	return &para
}

/*
감추기
*/
func (p *HParameterSet) HPageHiding() *HPageHiding {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPageHiding")
	if err != nil {
		panic(err)
	}
	var para HPageHiding
	para.variant = res
	return &para
}

/*
쪽 번호 위치
*/
func (p *HParameterSet) HPageNumPos() *HPageNumPos {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPageNumPos")
	if err != nil {
		panic(err)
	}
	var para HPageNumPos
	para.variant = res
	return &para
}

/*
머리말/꼬리말
*/
func (p *HParameterSet) HHeaderFooter() *HHeaderFooter {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HHeaderFooter")
	if err != nil {
		panic(err)
	}
	var para HHeaderFooter
	para.variant = res
	return &para
}

/*
메일 머지
*/
func (p *HParameterSet) HMailMergeGenerate() *HMailMergeGenerate {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMailMergeGenerate")
	if err != nil {
		panic(err)
	}
	var para HMailMergeGenerate
	para.variant = res
	return &para
}

/*
서브 리스트의 속성
*/
func (p *HParameterSet) HListProperties() *HListProperties {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HListProperties")
	if err != nil {
		panic(err)
	}
	var para HListProperties
	para.variant = res
	return &para
}

/*
표 속성
*/
func (p *HParameterSet) HTable() *HTable {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTable")
	if err != nil {
		panic(err)
	}
	var para HTable
	para.variant = res
	return &para
}

/*
셀 속성
*/
func (p *HParameterSet) HCell() *HCell {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCell")
	if err != nil {
		panic(err)
	}
	var para HCell
	para.variant = res
	return &para
}

/*
표 생성 시에만 사용
*/
func (p *HParameterSet) HTableCreation() *HTableCreation {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableCreation")
	if err != nil {
		panic(err)
	}
	var para HTableCreation
	para.variant = res
	return &para
}

/*
문서 정보
*/
func (p *HParameterSet) HSummaryInfo() *HSummaryInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSummaryInfo")
	if err != nil {
		panic(err)
	}
	var para HSummaryInfo
	para.variant = res
	return &para
}

/*
스타일
*/
func (p *HParameterSet) HStyleItem() *HStyleItem {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HStyleItem")
	if err != nil {
		panic(err)
	}
	var para HStyleItem
	para.variant = res
	return &para
}

/*
격자 정보
*/
func (p *HParameterSet) HGridInfo() *HGridInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HGridInfo")
	if err != nil {
		panic(err)
	}
	var para HGridInfo
	para.variant = res
	return &para
}

/*
컨트롤 데이터
*/
func (p *HParameterSet) HCtrlData() *HCtrlData {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCtrlData")
	if err != nil {
		panic(err)
	}
	var para HCtrlData
	para.variant = res
	return &para
}

/*
문서 데이터
	미구현
*/
/*
func (p *HParameterSet) HDocData() *HDocData {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDocData")
	if err != nil {
		panic(err)
	}
	var para HDocData
	para.variant = res
	return &para
}
*/

/*
문단 번호 모양 정의
*/
func (p *HParameterSet) HNumberingShape() *HNumberingShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HNumberingShape")
	if err != nil {
		panic(err)
	}
	var para HNumberingShape
	para.variant = res
	return &para
}

/*
불릿 모양 정의
*/
func (p *HParameterSet) HBulletShape() *HBulletShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HBulletShape")
	if err != nil {
		panic(err)
	}
	var para HBulletShape
	para.variant = res
	return &para
}

/*
찾아보기 표식 정의
*/
func (p *HParameterSet) HIndexMark() *HIndexMark {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HIndexMark")
	if err != nil {
		panic(err)
	}
	var para HIndexMark
	para.variant = res
	return &para
}

/*
글자 겹침
*/
func (p *HParameterSet) HChCompose() *HChCompose {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HChCompose")
	if err != nil {
		panic(err)
	}
	var para HChCompose
	para.variant = res
	return &para
}

/*
표 - 셀 나누기
*/
func (p *HParameterSet) HTableSplitCell() *HTableSplitCell {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableSplitCell")
	if err != nil {
		panic(err)
	}
	var para HTableSplitCell
	para.variant = res
	return &para
}

/*
테두리/배경의 일반 속성
*/
func (p *HParameterSet) HBorderFill() *HBorderFill {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HBorderFill")
	if err != nil {
		panic(err)
	}
	var para HBorderFill
	para.variant = res
	return &para
}

/*
UI 구현을 위해 (HBorderFill을 확장)
*/
func (p *HParameterSet) HBorderFillExt() *HBorderFillExt {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HBorderFillExt")
	if err != nil {
		panic(err)
	}
	var para HBorderFillExt
	para.variant = res
	return &para
}

/*
구역의 쪽 테두리/배경
*/
func (p *HParameterSet) HPageBorderFill() *HPageBorderFill {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPageBorderFill")
	if err != nil {
		panic(err)
	}
	var para HPageBorderFill
	para.variant = res
	return &para
}

/*
문서 암호
*/
func (p *HParameterSet) HPassword() *HPassword {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPassword")
	if err != nil {
		panic(err)
	}
	var para HPassword
	para.variant = res
	return &para
}

/*
필드 컨트롤의 공통 데이터
*/
func (p *HParameterSet) HFieldCtrl() *HFieldCtrl {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFieldCtrl")
	if err != nil {
		panic(err)
	}
	var para HFieldCtrl
	para.variant = res
	return &para
}

/*
덧말
*/
func (p *HParameterSet) HDutmal() *HDutmal {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDutmal")
	if err != nil {
		panic(err)
	}
	var para HDutmal
	para.variant = res
	return &para
}

/*
셀 테두리/배경
*/
func (p *HParameterSet) HCellBorderFill() *HCellBorderFill {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCellBorderFill")
	if err != nil {
		panic(err)
	}
	var para HCellBorderFill
	para.variant = res
	return &para
}

/*
표 - 줄/칸 지우기
*/
func (p *HParameterSet) HTableDeleteLine() *HTableDeleteLine {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableDeleteLine")
	if err != nil {
		panic(err)
	}
	var para HTableDeleteLine
	para.variant = res
	return &para
}

/*
표 - 줄/칸 삽입
*/
func (p *HParameterSet) HTableInsertLine() *HTableInsertLine {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableInsertLine")
	if err != nil {
		panic(err)
	}
	var para HTableInsertLine
	para.variant = res
	return &para
}

/*
캡션 속성
*/
func (p *HParameterSet) HCaption() *HCaption {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCaption")
	if err != nil {
		panic(err)
	}
	var para HCaption
	para.variant = res
	return &para
}

/*
하이퍼링크 점프
*/
func (p *HParameterSet) HHyperlinkJump() *HHyperlinkJump {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HHyperlinkJump")
	if err != nil {
		panic(err)
	}
	var para HHyperlinkJump
	para.variant = res
	return &para
}

/*
표 - 표 그리기
*/
func (p *HParameterSet) HTableDrawPen() *HTableDrawPen {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableDrawPen")
	if err != nil {
		panic(err)
	}
	var para HTableDrawPen
	para.variant = res
	return &para
}

/*
수식 - 수식 만들기
*/
func (p *HParameterSet) HEqEdit() *HEqEdit {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HEqEdit")
	if err != nil {
		panic(err)
	}
	var para HEqEdit
	para.variant = res
	return &para
}

/*
갈무리 - 좌표 전달용
*/
func (p *HParameterSet) HCaptureEnd() *HCaptureEnd {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCaptureEnd")
	if err != nil {
		panic(err)
	}
	var para HCaptureEnd
	para.variant = res
	return &para
}

/*
그림으로 저장하기
*/
func (p *HParameterSet) HPrintToImage() *HPrintToImage {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPrintToImage")
	if err != nil {
		panic(err)
	}
	var para HPrintToImage
	para.variant = res
	return &para
}

/*
블록 저장하기
*/
func (p *HParameterSet) HFileSaveBlock() *HFileSaveBlock {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileSaveBlock")
	if err != nil {
		panic(err)
	}
	var para HFileSaveBlock
	para.variant = res
	return &para
}

/*
편지 보내기
*/
func (p *HParameterSet) HFileSendMail() *HFileSendMail {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileSendMail")
	if err != nil {
		panic(err)
	}
	var para HFileSendMail
	para.variant = res
	return &para
}

/*
메시지 박스
*/
func (p *HParameterSet) HHncMessageBox() *HHncMessageBox {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HHncMessageBox")
	if err != nil {
		panic(err)
	}
	var para HHncMessageBox
	para.variant = res
	return &para
}

/*
표에 링크된 차트 정보
*/
func (p *HParameterSet) HTableChartInfo() *HTableChartInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableChartInfo")
	if err != nil {
		panic(err)
	}
	var para HTableChartInfo
	para.variant = res
	return &para
}

/*
표를 문자열로
*/
func (p *HParameterSet) HTableTblToStr() *HTableTblToStr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableTblToStr")
	if err != nil {
		panic(err)
	}
	var para HTableTblToStr
	para.variant = res
	return &para
}

/*
문자열을 표로...
*/
func (p *HParameterSet) HTableStrToTbl() *HTableStrToTbl {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableStrToTbl")
	if err != nil {
		panic(err)
	}
	var para HTableStrToTbl
	para.variant = res
	return &para
}

/*
플래시 삽입 옵션
*/
func (p *HParameterSet) HFlashProperties() *HFlashProperties {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFlashProperties")
	if err != nil {
		panic(err)
	}
	var para HFlashProperties
	para.variant = res
	return &para
}

/*
동영상 삽입 옵션
*/
func (p *HParameterSet) HMovieProperties() *HMovieProperties {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMovieProperties")
	if err != nil {
		panic(err)
	}
	var para HMovieProperties
	para.variant = res
	return &para
}

/*
찾기 바꾸기 옵션
*/
func (p *HParameterSet) HFindReplace() *HFindReplace {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFindReplace")
	if err != nil {
		panic(err)
	}
	var para HFindReplace
	para.variant = res
	return &para
}

/*
표 마당 에서만 사용
*/
func (p *HParameterSet) HTableTemplate() *HTableTemplate {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableTemplate")
	if err != nil {
		panic(err)
	}
	var para HTableTemplate
	para.variant = res
	return &para
}

/*
문서 보호 (배포용 문서)
*/
func (p *HParameterSet) HFileSecurity() *HFileSecurity {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileSecurity")
	if err != nil {
		panic(err)
	}
	var para HFileSecurity
	para.variant = res
	return &para
}

/*
라벨
*/
func (p *HParameterSet) HLabel() *HLabel {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HLabel")
	if err != nil {
		panic(err)
	}
	var para HLabel
	para.variant = res
	return &para
}

/*
형광펜
*/
func (p *HParameterSet) HMarkpenShape() *HMarkpenShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMarkpenShape")
	if err != nil {
		panic(err)
	}
	var para HMarkpenShape
	para.variant = res
	return &para
}

/*
HTML 문서 붙이기 대화상자
*/
func (p *HParameterSet) HPasteHtml() *HPasteHtml {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPasteHtml")
	if err != nil {
		panic(err)
	}
	var para HPasteHtml
	para.variant = res
	return &para
}

/*
OLE 개체 생성 시에만 사용된다.
*/
func (p *HParameterSet) HOleCreation() *HOleCreation {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HOleCreation")
	if err != nil {
		panic(err)
	}
	var para HOleCreation
	para.variant = res
	return &para
}

/*
표 뒤집기
*/
func (p *HParameterSet) HTableSwap() *HTableSwap {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HTableSwap")
	if err != nil {
		panic(err)
	}
	var para HTableSwap
	para.variant = res
	return &para
}

/*
버전 정보
*/
func (p *HParameterSet) HVersionInfo() *HVersionInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HVersionInfo")
	if err != nil {
		panic(err)
	}
	var para HVersionInfo
	para.variant = res
	return &para
}

/*
메모 모양
*/
func (p *HParameterSet) HMemoShape() *HMemoShape {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMemoShape")
	if err != nil {
		panic(err)
	}
	var para HMemoShape
	para.variant = res
	return &para
}

/*
소트
*/
func (p *HParameterSet) HSort() *HSort {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSort")
	if err != nil {
		panic(err)
	}
	var para HSort
	para.variant = res
	return &para
}

/*
그리기 개체 레이아웃
*/
func (p *HParameterSet) HDrawLayOut() *HDrawLayOut {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawLayOut")
	if err != nil {
		panic(err)
	}
	var para HDrawLayOut
	para.variant = res
	return &para
}

/*
그리기 개체 사각형
*/
func (p *HParameterSet) HDrawRectType() *HDrawRectType {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawRectType")
	if err != nil {
		panic(err)
	}
	var para HDrawRectType
	para.variant = res
	return &para
}

/*
그리기 개체 선
*/
func (p *HParameterSet) HDrawLineAttr() *HDrawLineAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawLineAttr")
	if err != nil {
		panic(err)
	}
	var para HDrawLineAttr
	para.variant = res
	return &para
}

/*
그리기 개체 호
*/
func (p *HParameterSet) HDrawArcType() *HDrawArcType {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawArcType")
	if err != nil {
		panic(err)
	}
	var para HDrawArcType
	para.variant = res
	return &para
}

/*
그리기 개체 그림
*/
func (p *HParameterSet) HDrawImageAttr() *HDrawImageAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawImageAttr")
	if err != nil {
		panic(err)
	}
	var para HDrawImageAttr
	para.variant = res
	return &para
}

/*
그리기 개체 늘이기
*/
func (p *HParameterSet) HDrawResize() *HDrawResize {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawResize")
	if err != nil {
		panic(err)
	}
	var para HDrawResize
	para.variant = res
	return &para
}

/*
그리기 개체 회전
*/
func (p *HParameterSet) HDrawRotate() *HDrawRotate {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawRotate")
	if err != nil {
		panic(err)
	}
	var para HDrawRotate
	para.variant = res
	return &para
}

/*
그리기 개체 점 편집
*/
func (p *HParameterSet) HDrawEditDetail() *HDrawEditDetail {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawEditDetail")
	if err != nil {
		panic(err)
	}
	var para HDrawEditDetail
	para.variant = res
	return &para
}

/*
그림 자르기
*/
func (p *HParameterSet) HDrawImageScissoring() *HDrawImageScissoring {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawImageScissoring")
	if err != nil {
		panic(err)
	}
	var para HDrawImageScissoring
	para.variant = res
	return &para
}

/*
그리기 개체 뒤집기
*/
func (p *HParameterSet) HDrawScAction() *HDrawScAction {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawScAction")
	if err != nil {
		panic(err)
	}
	var para HDrawScAction
	para.variant = res
	return &para
}

/*
그리기 개체 하이퍼 링크
*/
func (p *HParameterSet) HDrawCtrlHyperlink() *HDrawCtrlHyperlink {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawCtrlHyperlink")
	if err != nil {
		panic(err)
	}
	var para HDrawCtrlHyperlink
	para.variant = res
	return &para
}

/*
그리기 개체 좌표 정보
*/
func (p *HParameterSet) HDrawCoordInfo() *HDrawCoordInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawCoordInfo")
	if err != nil {
		panic(err)
	}
	var para HDrawCoordInfo
	para.variant = res
	return &para
}

/*
쉬운 개체 선택, 마우스 끌기로 만들기 정보
*/
func (p *HParameterSet) HDrawSoMouseOption() *HDrawSoMouseOption {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawSoMouseOption")
	if err != nil {
		panic(err)
	}
	var para HDrawSoMouseOption
	para.variant = res
	return &para
}

/*
수식 개체 정보
*/
func (p *HParameterSet) HDrawSoEquationOption() *HDrawSoEquationOption {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawSoEquationOption")
	if err != nil {
		panic(err)
	}
	var para HDrawSoEquationOption
	para.variant = res
	return &para
}

/*
그리기 개체 기울기
*/
func (p *HParameterSet) HDrawShear() *HDrawShear {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawShear")
	if err != nil {
		panic(err)
	}
	var para HDrawShear
	para.variant = res
	return &para
}

/*
글맵시
*/
func (p *HParameterSet) HDrawTextart() *HDrawTextart {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDrawTextart")
	if err != nil {
		panic(err)
	}
	var para HDrawTextart
	para.variant = res
	return &para
}

/*
양식 개체 타입 정보
*/
func (p *HParameterSet) HFormGeneral() *HFormGeneral {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormGeneral")
	if err != nil {
		panic(err)
	}
	var para HFormGeneral
	para.variant = res
	return &para
}

/*
양식 개체 공통 속성
*/
func (p *HParameterSet) HFormCommonAttr() *HFormCommonAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormCommonAttr")
	if err != nil {
		panic(err)
	}
	var para HFormCommonAttr
	para.variant = res
	return &para
}

/*
양식 개체 글자모양 속성
*/
func (p *HParameterSet) HFormCharshapeattr() *HFormCharshapeattr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormCharshapeattr")
	if err != nil {
		panic(err)
	}
	var para HFormCharshapeattr
	para.variant = res
	return &para
}

/*
양식 개체 버튼
*/
func (p *HParameterSet) HFormButtonAttr() *HFormButtonAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormButtonAttr")
	if err != nil {
		panic(err)
	}
	var para HFormButtonAttr
	para.variant = res
	return &para
}

/*
양식 개체 콤보 박스
*/
func (p *HParameterSet) HFormComboboxAttr() *HFormComboboxAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormComboboxAttr")
	if err != nil {
		panic(err)
	}
	var para HFormComboboxAttr
	para.variant = res
	return &para
}

/*
양식 개체 에디트 상자
*/
func (p *HParameterSet) HFormEditAttr() *HFormEditAttr {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormEditAttr")
	if err != nil {
		panic(err)
	}
	var para HFormEditAttr
	para.variant = res
	return &para
}

/*
스타일
*/
func (p *HParameterSet) HStyle() *HStyle {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HStyle")
	if err != nil {
		panic(err)
	}
	var para HStyle
	para.variant = res
	return &para
}

/*
문서 열기
*/
func (p *HParameterSet) HFileOpen() *HFileOpen {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileOpen")
	if err != nil {
		panic(err)
	}
	var para HFileOpen
	para.variant = res
	return &para
}

/*
문서 저장
*/
func (p *HParameterSet) HFileSaveAs() *HFileSaveAs {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileSaveAs")
	if err != nil {
		panic(err)
	}
	var para HFileSaveAs
	para.variant = res
	return &para
}

/*
문단 첫 글자 장식
*/
func (p *HParameterSet) HDropCap() *HDropCap {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDropCap")
	if err != nil {
		panic(err)
	}
	var para HDropCap
	para.variant = res
	return &para
}

/*
키 매크로
*/
func (p *HParameterSet) HKeyMacro() *HKeyMacro {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HKeyMacro")
	if err != nil {
		panic(err)
	}
	var para HKeyMacro
	para.variant = res
	return &para
}

/*
모양 복사
*/
func (p *HParameterSet) HShapeCopyPaste() *HShapeCopyPaste {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HShapeCopyPaste")
	if err != nil {
		panic(err)
	}
	var para HShapeCopyPaste
	para.variant = res
	return &para
}

/*
스타일 마당
*/
func (p *HParameterSet) HStyleTemplate() *HStyleTemplate {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HStyleTemplate")
	if err != nil {
		panic(err)
	}
	var para HStyleTemplate
	para.variant = res
	return &para
}

/*
바탕쪽
*/
func (p *HParameterSet) HMasterPage() *HMasterPage {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMasterPage")
	if err != nil {
		panic(err)
	}
	var para HMasterPage
	para.variant = res
	return &para
}

/*
찾아가기
*/
func (p *HParameterSet) HGotoE() *HGotoE {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HGotoE")
	if err != nil {
		panic(err)
	}
	var para HGotoE
	para.variant = res
	return &para
}

/*
텍스트 삽입
*/
func (p *HParameterSet) HInsertText() *HInsertText {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HInsertText")
	if err != nil {
		panic(err)
	}
	var para HInsertText
	para.variant = res
	return &para
}

/*
책갈피
*/
func (p *HParameterSet) HBookMark() *HBookMark {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HBookMark")
	if err != nil {
		panic(err)
	}
	var para HBookMark
	para.variant = res
	return &para
}

/*
하이퍼링크 삽입/고치기
*/
func (p *HParameterSet) HHyperLink() *HHyperLink {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HHyperLink")
	if err != nil {
		panic(err)
	}
	var para HHyperLink
	para.variant = res
	return &para
}

/*
문서 마당 정보
*/
func (p *HParameterSet) HInsertFieldTemplate() *HInsertFieldTemplate {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HInsertFieldTemplate")
	if err != nil {
		panic(err)
	}
	var para HInsertFieldTemplate
	para.variant = res
	return &para
}

/*
상용구
*/
func (p *HParameterSet) HIdiom() *HIdiom {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HIdiom")
	if err != nil {
		panic(err)
	}
	var para HIdiom
	para.variant = res
	return &para
}

/*
문자표
*/
func (p *HParameterSet) HCodeTable() *HCodeTable {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HCodeTable")
	if err != nil {
		panic(err)
	}
	var para HCodeTable
	para.variant = res
	return &para
}

/*
대/소문자 변환
*/
func (p *HParameterSet) HConvertCase() *HConvertCase {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HConvertCase")
	if err != nil {
		panic(err)
	}
	var para HConvertCase
	para.variant = res
	return &para
}

/*
한자, 일어, 구결을 한글로
*/
func (p *HParameterSet) HConvertToHangul() *HConvertToHangul {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HConvertToHangul")
	if err != nil {
		panic(err)
	}
	var para HConvertToHangul
	para.variant = res
	return &para
}

/*
히라가나/가타가나 변환
*/
func (p *HParameterSet) HConvertHiraToGata() *HConvertHiraToGata {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HConvertHiraToGata")
	if err != nil {
		panic(err)
	}
	var para HConvertHiraToGata
	para.variant = res
	return &para
}

/*
전/반각 변환
*/
func (p *HParameterSet) HConvertFullHalf() *HConvertFullHalf {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HConvertFullHalf")
	if err != nil {
		panic(err)
	}
	var para HConvertFullHalf
	para.variant = res
	return &para
}

/*
파일 삽입
*/
func (p *HParameterSet) HInsertFile() *HInsertFile {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HInsertFile")
	if err != nil {
		panic(err)
	}
	var para HInsertFile
	para.variant = res
	return &para
}

/*
Document 필터 리스트
*/
func (p *HParameterSet) HDocFilters() *HDocFilters {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDocFilters")
	if err != nil {
		panic(err)
	}
	var para HDocFilters
	para.variant = res
	return &para
}

/*
빠른 교정
*/
func (p *HParameterSet) HQCorrect() *HQCorrect {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HQCorrect")
	if err != nil {
		panic(err)
	}
	var para HQCorrect
	para.variant = res
	return &para
}

/*
차례 만들기
*/
func (p *HParameterSet) HMakeContents() *HMakeContents {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMakeContents")
	if err != nil {
		panic(err)
	}
	var para HMakeContents
	para.variant = res
	return &para
}

/*
환경 설정
*/
func (p *HParameterSet) HPreference() *HPreference {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HPreference")
	if err != nil {
		panic(err)
	}
	var para HPreference
	para.variant = res
	return &para
}

/*
인터넷
*/
func (p *HParameterSet) HInternet() *HInternet {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HInternet")
	if err != nil {
		panic(err)
	}
	var para HInternet
	para.variant = res
	return &para
}

/*
웹서버로 올리기
*/
func (p *HParameterSet) HFtpUpload() *HFtpUpload {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFtpUpload")
	if err != nil {
		panic(err)
	}
	var para HFtpUpload
	para.variant = res
	return &para
}

/*
파일 변환
*/
func (p *HParameterSet) HFileConvert() *HFileConvert {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileConvert")
	if err != nil {
		panic(err)
	}
	var para HFileConvert
	para.variant = res
	return &para
}

/*
클립보드 Cut/Delete/Paste옵션(셀 블록 상태 일 때)
*/
func (p *HParameterSet) HSelectionOpt() *HSelectionOpt {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSelectionOpt")
	if err != nil {
		panic(err)
	}
	var para HSelectionOpt
	para.variant = res
	return &para
}

/*
문서 정보
*/
func (p *HParameterSet) HDocumentInfo() *HDocumentInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDocumentInfo")
	if err != nil {
		panic(err)
	}
	var para HDocumentInfo
	para.variant = res
	return &para
}

/*
각주/미주 변환
*/
func (p *HParameterSet) HExchangeFootnoteEndNote() *HExchangeFootnoteEndNote {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HExchangeFootnoteEndNote")
	if err != nil {
		panic(err)
	}
	var para HExchangeFootnoteEndNote
	para.variant = res
	return &para
}

/*
주석 저장
*/
func (p *HParameterSet) HSaveFootnote() *HSaveFootnote {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSaveFootnote")
	if err != nil {
		panic(err)
	}
	var para HSaveFootnote
	para.variant = res
	return &para
}

/*
로마자 변환
*/
func (p *HParameterSet) HChangeRome() *HChangeRome {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HChangeRome")
	if err != nil {
		panic(err)
	}
	var para HChangeRome
	para.variant = res
	return &para
}

/*
주소 검색
*/
func (p *HParameterSet) HSearchAddress() *HSearchAddress {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSearchAddress")
	if err != nil {
		panic(err)
	}
	var para HSearchAddress
	para.variant = res
	return &para
}

/*
한자 단어 추가
*/
func (p *HParameterSet) HAddHanjaWord() *HAddHanjaWord {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HAddHanjaWord")
	if err != nil {
		panic(err)
	}
	var para HAddHanjaWord
	para.variant = res
	return &para
}

/*
문서 찾기 결과
*/
func (p *HParameterSet) HDocFindInfo() *HDocFindInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDocFindInfo")
	if err != nil {
		panic(err)
	}
	var para HDocFindInfo
	para.variant = res
	return &para
}

/*
블록 계산(합계/평균/줄 수)
*/
func (p *HParameterSet) HSum() *HSum {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSum")
	if err != nil {
		panic(err)
	}
	var para HSum
	para.variant = res
	return &para
}

/*
현재 캐럿의 위치(GetPos, SetPos API에서 사용)
*/
func (p *HParameterSet) HListParaPos() *HListParaPos {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HListParaPos")
	if err != nil {
		panic(err)
	}
	var para HListParaPos
	para.variant = res
	return &para
}

/*
본문의 내용 얻어오기 (GetText API에서 사용)
*/
func (p *HParameterSet) HGetText() *HGetText {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HGetText")
	if err != nil {
		panic(err)
	}
	var para HGetText
	para.variant = res
	return &para
}

/*
간/번체 변환
*/
func (p *HParameterSet) HConvertJianFan() *HConvertJianFan {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HConvertJianFan")
	if err != nil {
		panic(err)
	}
	var para HConvertJianFan
	para.variant = res
	return &para
}

/*
마우스 위치 얻어오기 (GetMousePos API에서 사용)
*/
func (p *HParameterSet) HMousePos() *HMousePos {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMousePos")
	if err != nil {
		panic(err)
	}
	var para HMousePos
	para.variant = res
	return &para
}

/*
파일 정보 알아오기 (GetFileInfo에서 API에서 사용)
*/
func (p *HParameterSet) HFileInfo() *HFileInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileInfo")
	if err != nil {
		panic(err)
	}
	var para HFileInfo
	para.variant = res
	return &para
}

/*
상호 참조
*/
func (p *HParameterSet) HActionCrossRef() *HActionCrossRef {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HActionCrossRef")
	if err != nil {
		panic(err)
	}
	var para HActionCrossRef
	para.variant = res
	return &para
}

/*
사용자 추가 정보 저장
*/
func (p *HParameterSet) HSaveUserInfoFile() *HSaveUserInfoFile {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HSaveUserInfoFile")
	if err != nil {
		panic(err)
	}
	var para HSaveUserInfoFile
	para.variant = res
	return &para
}

/*
사용자 추가 정보 가져오기
*/
func (p *HParameterSet) HLoadUserInfoFile() *HLoadUserInfoFile {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HLoadUserInfoFile")
	if err != nil {
		panic(err)
	}
	var para HLoadUserInfoFile
	para.variant = res
	return &para
}

/*
사용자 자동 명령 파일 저장/로드
*/
func (p *HParameterSet) HUserQCommandFile() *HUserQCommandFile {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HUserQCommandFile")
	if err != nil {
		panic(err)
	}
	var para HUserQCommandFile
	para.variant = res
	return &para
}

/*
날짜/시간 표시 형식
*/
func (p *HParameterSet) HInputDateStyle() *HInputDateStyle {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HInputDateStyle")
	if err != nil {
		panic(err)
	}
	var para HInputDateStyle
	para.variant = res
	return &para
}

/*
배포용 문서 암호
*/
func (p *HParameterSet) HFileSetSecurity() *HFileSetSecurity {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileSetSecurity")
	if err != nil {
		panic(err)
	}
	var para HFileSetSecurity
	para.variant = res
	return &para
}

/*
문서 연결
*/
func (p *HParameterSet) HLinkDocument() *HLinkDocument {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HLinkDocument")
	if err != nil {
		panic(err)
	}
	var para HLinkDocument
	para.variant = res
	return &para
}

/*
파일 필터에서 사용하는 다이얼로그
*/
func (p *HParameterSet) HDocumentFilterDialog() *HDocumentFilterDialog {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HDocumentFilterDialog")
	if err != nil {
		panic(err)
	}
	var para HDocumentFilterDialog
	para.variant = res
	return &para
}

/*
그리기개체를 그림으로 저장하기
*/
func (p *HParameterSet) HShapeObjSaveAsPicture() *HShapeObjSaveAsPicture {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HShapeObjSaveAsPicture")
	if err != nil {
		panic(err)
	}
	var para HShapeObjSaveAsPicture
	para.variant = res
	return &para
}

/*
환경 설정 속성
*/
func (p *HParameterSet) HEngineProperties() *HEngineProperties {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HEngineProperties")
	if err != nil {
		panic(err)
	}
	var para HEngineProperties
	para.variant = res
	return &para
}

/*
ScrollPos property (GetScrollPosInfo와 OnScroll 이벤트 에서 사용)
*/
func (p *HParameterSet) HScrollPosInfo() *HScrollPosInfo {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HScrollPosInfo")
	if err != nil {
		panic(err)
	}
	var para HScrollPosInfo
	para.variant = res
	return &para
}

/*
Event Message property (NotifyMessageSet 이벤트 에서 사용)
*/
func (p *HParameterSet) HMessageSet() *HMessageSet {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HMessageSet")
	if err != nil {
		panic(err)
	}
	var para HMessageSet
	para.variant = res
	return &para
}

/*
XML 관련 메뉴
*/
func (p *HParameterSet) HFileXMLSchema() *HFileXMLSchema {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileXMLSchema")
	if err != nil {
		panic(err)
	}
	var para HFileXMLSchema
	para.variant = res
	return &para
}

/*
양식 개체(Edit)에서만 사용하는 코드표
*/
func (p *HParameterSet) HFormObjInputCodeTable() *HFormObjInputCodeTable {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormObjInputCodeTable")
	if err != nil {
		panic(err)
	}
	var para HFormObjInputCodeTable
	para.variant = res
	return &para
}

/*
양식 개체(Edit)에서만 사용하는 한자 변환
*/
func (p *HParameterSet) HFormObjInputHanja() *HFormObjInputHanja {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormObjInputHanja")
	if err != nil {
		panic(err)
	}
	var para HFormObjInputHanja
	para.variant = res
	return &para
}

/*
양식 개체(Edit)에서만 사용하는 한자 부수 입력
*/
func (p *HParameterSet) HFormObjHanjaBusu() *HFormObjHanjaBusu {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormObjHanjaBusu")
	if err != nil {
		panic(err)
	}
	var para HFormObjHanjaBusu
	para.variant = res
	return &para
}

/*
양식 개체(Edit)에서만 사용하는 한자 새김 입력
*/
func (p *HParameterSet) HFormObjHanjaMean() *HFormObjHanjaMean {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormObjHanjaMean")
	if err != nil {
		panic(err)
	}
	var para HFormObjHanjaMean
	para.variant = res
	return &para
}

/*
양식 개체(Edit)에서만 사용하는 상용구 입력
*/
func (p *HParameterSet) HFormObjInputIdiom() *HFormObjInputIdiom {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFormObjInputIdiom")
	if err != nil {
		panic(err)
	}
	var para HFormObjInputIdiom
	para.variant = res
	return &para
}

/*
뷰 상태 정보
*/
func (p *HParameterSet) HViewStatus() *HViewStatus {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HViewStatus")
	if err != nil {
		panic(err)
	}
	var para HViewStatus
	para.variant = res
	return &para
}

/*
스크립트 매크로
*/
func (p *HParameterSet) HScriptMacro() *HScriptMacro {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HScriptMacro")
	if err != nil {
		panic(err)
	}
	var para HScriptMacro
	para.variant = res
	return &para
}

/*
파일 불러오기/저장 Dialog
*/
func (p *HParameterSet) HFileOpenSave() *HFileOpenSave {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HFileOpenSave")
	if err != nil {
		panic(err)
	}
	var para HFileOpenSave
	para.variant = res
	return &para
}

/*
XML 파일 불러오기/저장 Dialog
*/
func (p *HParameterSet) HXMLOpenSave() *HXMLOpenSave {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HXMLOpenSave")
	if err != nil {
		panic(err)
	}
	var para HXMLOpenSave
	para.variant = res
	return &para
}

/*
XML Schema 만들기
*/
func (p *HParameterSet) HXMLSchema() *HXMLSchema {
	res, err := oleutil.GetProperty(p.variant.ToIDispatch(), "HXMLSchema")
	if err != nil {
		panic(err)
	}
	var para HXMLSchema
	para.variant = res
	return &para
}

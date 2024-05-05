package gohwp

/*
ActionObject
*/
type ActionObject struct {
	hwp *IHwpObject
}

/*
Frame Action: 창
*/
func (p *ActionObject) FrameActionWindow() *FrameActionWindow {
	var act FrameActionWindow
	act.hwp = p.hwp
	return &act
}

type FrameActionWindow struct {
	hwp *IHwpObject
}

/*
Frame Action: 기타
*/
func (p *ActionObject) FrameActionEtc() *FrameActionEtc {
	var act FrameActionEtc
	act.hwp = p.hwp
	return &act
}

type FrameActionEtc struct {
	hwp *IHwpObject
}

/*
Frame Action: 도움말
*/
func (p *ActionObject) FrameActionHelp() *FrameActionHelp {
	var act FrameActionHelp
	act.hwp = p.hwp
	return &act
}

type FrameActionHelp struct {
	hwp *IHwpObject
}

/*
Frame Action: 파일
*/
func (p *ActionObject) FrameActionFile() *FrameActionFile {
	var act FrameActionFile
	act.hwp = p.hwp
	return &act
}

type FrameActionFile struct {
	hwp *IHwpObject
}

/*
Frame Action: 보기
*/
func (p *ActionObject) FrameActionView() *FrameActionView {
	var act FrameActionView
	act.hwp = p.hwp
	return &act
}

type FrameActionView struct {
	hwp *IHwpObject
}

/*
Action: 파일
*/
func (p *ActionObject) ActionFile() *ActionFile {
	var act ActionFile
	act.hwp = p.hwp
	return &act
}

type ActionFile struct {
	hwp *IHwpObject
}

/*
Action: 편집
*/
func (p *ActionObject) ActionEdit() *ActionEdit {
	var act ActionEdit
	act.hwp = p.hwp
	return &act
}

type ActionEdit struct {
	hwp *IHwpObject
}

/*
Action: 모양
*/
func (p *ActionObject) ActionShape() *ActionShape {
	var act ActionShape
	act.hwp = p.hwp
	return &act
}

type ActionShape struct {
	hwp *IHwpObject
}

/*
Action: 보기
*/
func (p *ActionObject) ActionView() *ActionView {
	var act ActionView
	act.hwp = p.hwp
	return &act
}

type ActionView struct {
	hwp *IHwpObject
}

/*
Action: 입력
*/
func (p *ActionObject) ActionInput() *ActionInput {
	var act ActionInput
	act.hwp = p.hwp
	return &act
}

type ActionInput struct {
	hwp *IHwpObject
}

/*
Action: 도구
*/
func (p *ActionObject) ActionTool() *ActionTool {
	var act ActionTool
	act.hwp = p.hwp
	return &act
}

type ActionTool struct {
	hwp *IHwpObject
}

/*
Action: 표
*/
func (p *ActionObject) ActionTable() *ActionTable {
	var act ActionTable
	act.hwp = p.hwp
	return &act
}

type ActionTable struct {
	hwp *IHwpObject
}

/*
Action: 그리기
*/
func (p *ActionObject) ActionDraw() *ActionDraw {
	var act ActionDraw
	act.hwp = p.hwp
	return &act
}

type ActionDraw struct {
	hwp *IHwpObject
}

/*
Action: 기타
*/
func (p *ActionObject) ActionEtc() *ActionEtc {
	var act ActionEtc
	act.hwp = p.hwp
	return &act
}

type ActionEtc struct {
	hwp *IHwpObject
}

/*
FrameActionWindow:SplitAll
    창 가로 세로 나누기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitAll() {
	p.hwp.Run("SplitAll")
}

/*
FrameActionWindow:NoSplit
    창 나누지 않음
    파라미터셋: 없음
*/
func (p *FrameActionWindow) NoSplit() {
	p.hwp.Run("NoSplit")
}

/*
FrameActionWindow:SplitVert
    창 세로로 나누기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitVert() {
	p.hwp.Run("SplitVert")
}

/*
FrameActionWindow:SplitHorz
    창 가로로 나누기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitHorz() {
	p.hwp.Run("SplitHorz")
}

/*
FrameActionWindow:SplitMemo
    메모창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitMemo() {
	p.hwp.Run("SplitMemo")
}

/*
FrameActionWindow:SplitMainActive
    메모창 활성화
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitMainActive() {
	p.hwp.Run("SplitMainActive")
}

/*
FrameActionWindow:SplitMemoOpen
    메모창 열기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitMemoOpen() {
	p.hwp.Run("SplitMemoOpen")
}

/*
FrameActionWindow:SplitMemoClose
    메모창 닫기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) SplitMemoClose() {
	p.hwp.Run("SplitMemoClose")
}

/*
FrameActionWindow:WindowPrevTab
    이전 창 활성화
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowPrevTab() {
	p.hwp.Run("WindowPrevTab")
}

/*
FrameActionWindow:WindowNextTab
    다음 창 활성화
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowNextTab() {
	p.hwp.Run("WindowNextTab")
}

/*
FrameActionWindow:WindowNextPane
    다음 분할창 활성화
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowNextPane() {
	p.hwp.Run("WindowNextPane")
}

/*
FrameActionWindow:ShowNLPTab
    한글 도우미 작업창 보이기
    파라미터셋: 없음
*/
func (p *FrameActionWindow) ShowNLPTab() {
	p.hwp.Run("ShowNLPTab")
}

/*
FrameActionWindow:WindowList
    창 목록
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowList() {
	p.hwp.Run("WindowList")
}

/*
FrameActionWindow:WindowAlignCascade
    창 겹치게 배열
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowAlignCascade() {
	p.hwp.Run("WindowAlignCascade")
}

/*
FrameActionWindow:WindowAlignTileHorz
    창 가로로 배열
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowAlignTileHorz() {
	p.hwp.Run("WindowAlignTileHorz")
}

/*
FrameActionWindow:WindowAlignTileVert
    창 세로로 배열
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowAlignTileVert() {
	p.hwp.Run("WindowAlignTileVert")
}

/*
FrameActionWindow:WindowMinimizeAll
    창 모두 아이콘으로 배열
    파라미터셋: 없음
*/
func (p *FrameActionWindow) WindowMinimizeAll() {
	p.hwp.Run("WindowMinimizeAll")
}

/*
FrameActionEtc:CharShapeHeight
    글자 크기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CharShapeHeight() {
	p.hwp.Run("CharShapeHeight")
}

/*
FrameActionEtc:CharShapeSpacing
    글자 자간
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CharShapeSpacing() {
	p.hwp.Run("CharShapeSpacing")
}

/*
FrameActionEtc:CharShapeWidth
    글자 장평
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CharShapeWidth() {
	p.hwp.Run("CharShapeWidth")
}

/*
FrameActionEtc:StyleCombo
    글자 스타일
    파라미터셋: 없음
*/
func (p *FrameActionEtc) StyleCombo() {
	p.hwp.Run("StyleCombo")
}

/*
FrameActionEtc:CharShapeLang
    글자 언어
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CharShapeLang() {
	p.hwp.Run("CharShapeLang")
}

/*
FrameActionEtc:CharShapTypeface
    글자 모양
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CharShapTypeface() {
	p.hwp.Run("CharShapTypeface")
}

/*
FrameActionEtc:ParaShapeLineSpace
    문단 모양
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ParaShapeLineSpace() {
	p.hwp.Run("ParaShapeLineSpace")
}

/*
FrameActionEtc:ViewZoomRibon
    화면 확대
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ViewZoomRibon() {
	p.hwp.Run("ViewZoomRibon")
}

/*
FrameActionEtc:MasterPageType
    바탕쪽 종류
    파라미터셋: 없음
*/
func (p *FrameActionEtc) MasterPageType() {
	p.hwp.Run("MasterPageType")
}

/*
FrameActionEtc:TableDrawPenStyle
    표 그리기 선 모양
    파라미터셋: 없음
*/
func (p *FrameActionEtc) TableDrawPenStyle() {
	p.hwp.Run("TableDrawPenStyle")
}

/*
FrameActionEtc:TableDrawPenWidth
    표 그리기 선 굵기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) TableDrawPenWidth() {
	p.hwp.Run("TableDrawPenWidth")
}

/*
FrameActionEtc:NoteNumShape
    주석 번호 모양
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NoteNumShape() {
	p.hwp.Run("NoteNumShape")
}

/*
FrameActionEtc:NotePosition
    각주 위치
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NotePosition() {
	p.hwp.Run("NotePosition")
}

/*
FrameActionEtc:NoteLineLength
    주석 구분선 길이
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NoteLineLength() {
	p.hwp.Run("NoteLineLength")
}

/*
FrameActionEtc:NoteLineShape
    주석 구분선 모양
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NoteLineShape() {
	p.hwp.Run("NoteLineShape")
}

/*
FrameActionEtc:NoteLineWeight
    주석 구분선 굵기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NoteLineWeight() {
	p.hwp.Run("NoteLineWeight")
}

/*
FrameActionEtc:NoteLineColor
    주석 구분선 색
    파라미터셋: 없음
*/
func (p *FrameActionEtc) NoteLineColor() {
	p.hwp.Run("NoteLineColor")
}

/*
FrameActionEtc:EasyFind
    쉬운 찾기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) EasyFind() {
	p.hwp.Run("EasyFind")
}

/*
FrameActionEtc:MarkPenColor
    형광펜 색
    파라미터셋: 없음
*/
func (p *FrameActionEtc) MarkPenColor() {
	p.hwp.Run("MarkPenColor")
}

/*
FrameActionEtc:TopTabFrameClose
    위쪽 작업창 감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) TopTabFrameClose() {
	p.hwp.Run("TopTabFrameClose")
}

/*
FrameActionEtc:BottomTabFrameClose
    아래쪽 작업창 감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) BottomTabFrameClose() {
	p.hwp.Run("BottomTabFrameClose")
}

/*
FrameActionEtc:LeftTabFrameClose
    왼쪽 작업창 감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) LeftTabFrameClose() {
	p.hwp.Run("LeftTabFrameClose")
}

/*
FrameActionEtc:RightTabFrameClose
    오른쪽 작업창 감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) RightTabFrameClose() {
	p.hwp.Run("RightTabFrameClose")
}

/*
FrameActionEtc:ShowFloatTabFrame
    플로팅 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ShowFloatTabFrame() {
	p.hwp.Run("ShowFloatTabFrame")
}

/*
FrameActionEtc:SetWorkSpaceView
    작업창 보기 설정
    파라미터셋: 없음
*/
func (p *FrameActionEtc) SetWorkSpaceView() {
	p.hwp.Run("SetWorkSpaceView")
}

/*
FrameActionEtc:ShowTopWorkspace
    위쪽 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ShowTopWorkspace() {
	p.hwp.Run("ShowTopWorkspace")
}

/*
FrameActionEtc:ShowLeftWorkspace
    왼쪽 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ShowLeftWorkspace() {
	p.hwp.Run("ShowLeftWorkspace")
}

/*
FrameActionEtc:ShowBottomWorkspace
    아래쪽 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ShowBottomWorkspace() {
	p.hwp.Run("ShowBottomWorkspace")
}

/*
FrameActionEtc:ShowRightWorkspace
    오른쪽 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) ShowRightWorkspace() {
	p.hwp.Run("ShowRightWorkspace")
}

/*
FrameActionEtc:HelpAbout
    한글 2005 정보
    파라미터셋: 없음
*/
func (p *FrameActionEtc) HelpAbout() {
	p.hwp.Run("HelpAbout")
}

/*
FrameActionEtc:HancomRoom
    한컴 계약방
    파라미터셋: 없음
*/
func (p *FrameActionEtc) HancomRoom() {
	p.hwp.Run("HancomRoom")
}

/*
FrameActionEtc:PstGradientType
    프리젠테이션 그라데이션 형태
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstGradientType() {
	p.hwp.Run("PstGradientType")
}

/*
FrameActionEtc:PstScrChangeType
    프리젠테이션 화면 전환 형태
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstScrChangeType() {
	p.hwp.Run("PstScrChangeType")
}

/*
FrameActionEtc:PstBlackToWhite
    프리젠테이션 검은색 글자를 흰색으로 변경
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstBlackToWhite() {
	p.hwp.Run("PstBlackToWhite")
}

/*
FrameActionEtc:PstAutoPlay
    프리젠테이션 자동 시연
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstAutoPlay() {
	p.hwp.Run("PstAutoPlay")
}

/*
FrameActionEtc:PstSetupPrevSec
    프리젠테이션 앞 구역 설정
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstSetupPrevSec() {
	p.hwp.Run("PstSetupPrevSec")
}

/*
FrameActionEtc:PstSetupNextSec
    프리젠테이션 뒤 구역 설정
    파라미터셋: 없음
*/
func (p *FrameActionEtc) PstSetupNextSec() {
	p.hwp.Run("PstSetupNextSec")
}

/*
FrameActionEtc:CustRestBtn
    툴바 버튼 처음 상태로 되돌리기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustRestBtn() {
	p.hwp.Run("CustRestBtn")
}

/*
FrameActionEtc:CustRenameBtn
    툴바 버튼 이름 바꾸기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustRenameBtn() {
	p.hwp.Run("CustRenameBtn")
}

/*
FrameActionEtc:CustViewNameBtn
    툴바 버튼 이름만 보이기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustViewNameBtn() {
	p.hwp.Run("CustViewNameBtn")
}

/*
FrameActionEtc:CustViewIconBtn
    툴바 버튼 아이콘만 보이기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustViewIconBtn() {
	p.hwp.Run("CustViewIconBtn")
}

/*
FrameActionEtc:CustViewIconNameBtn
    툴바 버튼 이름과 아이콘 보이기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustViewIconNameBtn() {
	p.hwp.Run("CustViewIconNameBtn")
}

/*
FrameActionEtc:CustInsSepBtn
    툴바 버튼에 구분선 넣기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustInsSepBtn() {
	p.hwp.Run("CustInsSepBtn")
}

/*
FrameActionEtc:CustCopyBtn
    툴바 버튼 복사하기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustCopyBtn() {
	p.hwp.Run("CustCopyBtn")
}

/*
FrameActionEtc:CustCutBtn
    툴바 버튼 오려두기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustCutBtn() {
	p.hwp.Run("CustCutBtn")
}

/*
FrameActionEtc:CustEraseBtn
    툴바 버튼 지우기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustEraseBtn() {
	p.hwp.Run("CustEraseBtn")
}

/*
FrameActionEtc:CustPasteBtn
    툴바 버튼 붙여기
    파라미터셋: 없음
*/
func (p *FrameActionEtc) CustPasteBtn() {
	p.hwp.Run("CustPasteBtn")
}

/*
FrameActionHelp:HelpContents
    내용
    파라미터셋: 없음
*/
func (p *FrameActionHelp) HelpContents() {
	p.hwp.Run("HelpContents")
}

/*
FrameActionHelp:HelpIndex
    찾아보기
    파라미터셋: 없음
*/
func (p *FrameActionHelp) HelpIndex() {
	p.hwp.Run("HelpIndex")
}

/*
FrameActionHelp:HelpWeb
    온라인 고객 지원
    파라미터셋: 없음
*/
func (p *FrameActionHelp) HelpWeb() {
	p.hwp.Run("HelpWeb")
}

/*
FrameActionHelp:MasterWsItemOnOff
    바탕쪽 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionHelp) MasterWsItemOnOff() {
	p.hwp.Run("MasterWsItemOnOff")
}

/*
FrameActionHelp:AppAbout
    한글 2005 정보
    파라미터셋: 없음
*/
func (p *FrameActionHelp) AppAbout() {
	p.hwp.Run("AppAbout")
}

/*
FrameActionFile:FileNew
    새 글
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileNew() {
	p.hwp.Run("FileNew")
}

/*
FrameActionFile:FileNewTab
    새 탭
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileNewTab() {
	p.hwp.Run("FileNewTab")
}

/*
FrameActionFile:FileOpen
    불러오기
    파라미터셋: FileOpenSave
*/
func (p *FrameActionFile) FileOpen(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileOpen", arg...)
}

/*
FrameActionFile:FileSave
    저장하기
    파라미터셋: FileOpenSave
*/
func (p *FrameActionFile) FileSave(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileSave", arg...)
}

/*
FrameActionFile:FileOpenMRU
    최근 작업 문서
    파라미터셋: FileOpenSave
*/
func (p *FrameActionFile) FileOpenMRU(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileOpenMRU", arg...)
}

/*
FrameActionFile:FileSaveAs
    다른 이름으로 저장하기
    파라미터셋: FileOpenSave
*/
func (p *FrameActionFile) FileSaveAs(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileSaveAs", arg...)
}

/*
FrameActionFile:FileSaveAsHwp97
    한글 97문서로 저장하기
    파라미터셋: FileOpenSave
*/
func (p *FrameActionFile) FileSaveAsHwp97(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileSaveAsHwp97", arg...)
}

/*
FrameActionFile:FileSaveAsDRM
    배포용 문서로 저장하기
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileSaveAsDRM() {
	p.hwp.Run("FileSaveAsDRM")
}

/*
FrameActionFile:FileClose
    문서 닫기
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileClose() {
	p.hwp.Run("FileClose")
}

/*
FrameActionFile:TabClose
    현재탭 닫기
    파라미터셋: 없음
*/
func (p *FrameActionFile) TabClose() {
	p.hwp.Run("TabClose")
}

/*
FrameActionFile:FilePreview
    미리 보기
    파라미터셋: Print
*/
func (p *FrameActionFile) FilePreview(arg ...interface{}) {
	p.hwp.ActionWithParameters("FilePreview", arg...)
}

/*
FrameActionFile:FileQuit
    종료
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileQuit() {
	p.hwp.Run("FileQuit")
}

/*
FrameActionFile:FileFind
    문서 찾기
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileFind() {
	p.hwp.Run("FileFind")
}

/*
FrameActionFile:RecentEmpty
    최근 목록 지우기
    파라미터셋: 없음
*/
func (p *FrameActionFile) RecentEmpty() {
	p.hwp.Run("RecentEmpty")
}

/*
FrameActionFile:RecentNoExistDel
    최근 목록에서 존재하지 않는 아이템 지우기
    파라미터셋: 없음
*/
func (p *FrameActionFile) RecentNoExistDel() {
	p.hwp.Run("RecentNoExistDel")
}

/*
FrameActionFile:FileVersionOpen
    버전 불러오기
    파라미터셋: VersionInfo
*/
func (p *FrameActionFile) FileVersionOpen(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileVersionOpen", arg...)
}

/*
FrameActionFile:FileVersionDiff
    버전 비교
    파라미터셋: VersionInfo
*/
func (p *FrameActionFile) FileVersionDiff(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileVersionDiff", arg...)
}

/*
FrameActionFile:FileNextVersionDiff
    버전 비교 : 앞으로 이동
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileNextVersionDiff() {
	p.hwp.Run("FileNextVersionDiff")
}

/*
FrameActionFile:FilePrevVersionDiff
    버전 비교 : 뒤로 이동
    파라미터셋: 없음
*/
func (p *FrameActionFile) FilePrevVersionDiff() {
	p.hwp.Run("FilePrevVersionDiff")
}

/*
FrameActionFile:FileVersionDiffChangeAlign
    버전 비교 : 비교화면 배열 변경 (좌우↔상하)
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileVersionDiffChangeAlign() {
	p.hwp.Run("FileVersionDiffChangeAlign")
}

/*
FrameActionFile:FileVersionDiffSameAlign
    버전 비교 : 비교화면 다시 정렬
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileVersionDiffSameAlign() {
	p.hwp.Run("FileVersionDiffSameAlign")
}

/*
FrameActionFile:FileVersionDiffSyncScroll
    버전 비교 : 비교화면 동시에 이동
    파라미터셋: 없음
*/
func (p *FrameActionFile) FileVersionDiffSyncScroll() {
	p.hwp.Run("FileVersionDiffSyncScroll")
}

/*
FrameActionFile:ACreateSchema
    XML 문서 구조 불러오기
    파라미터셋: XMLSchema
*/
func (p *FrameActionFile) ACreateSchema(arg ...interface{}) {
	p.hwp.ActionWithParameters("ACreateSchema", arg...)
}

/*
FrameActionFile:AOpenXMLDoc
    XML 문서 불러오기
    파라미터셋: XMLOpenSave
*/
func (p *FrameActionFile) AOpenXMLDoc(arg ...interface{}) {
	p.hwp.ActionWithParameters("AOpenXMLDoc", arg...)
}

/*
FrameActionFile:ASaveXMLDoc
    XML 문서 저장하기
    파라미터셋: XMLOpenSave
*/
func (p *FrameActionFile) ASaveXMLDoc(arg ...interface{}) {
	p.hwp.ActionWithParameters("ASaveXMLDoc", arg...)
}

/*
FrameActionView:CustomizeToolbar
    도구상자 사용자 설정
    파라미터셋: 없음
*/
func (p *FrameActionView) CustomizeToolbar() {
	p.hwp.Run("CustomizeToolbar")
}

/*
FrameActionView:FrameFullScreen
    전체 화면
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameFullScreen() {
	p.hwp.Run("FrameFullScreen")
}

/*
FrameActionView:FrameFullScreenEnd
    전체 화면 닫기
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameFullScreenEnd() {
	p.hwp.Run("FrameFullScreenEnd")
}

/*
FrameActionView:FrameHRuler
    가로축 눈금자 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameHRuler() {
	p.hwp.Run("FrameHRuler")
}

/*
FrameActionView:FrameVRuler
    세로축 눈금자 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameVRuler() {
	p.hwp.Run("FrameVRuler")
}

/*
FrameActionView:FrameStatusBar
    상태바 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameStatusBar() {
	p.hwp.Run("FrameStatusBar")
}

/*
FrameActionView:HorzScrollbar
    가로축 스크롤바 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) HorzScrollbar() {
	p.hwp.Run("HorzScrollbar")
}

/*
FrameActionView:VertScrollbar
    세로축 스크롤바 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) VertScrollbar() {
	p.hwp.Run("VertScrollbar")
}

/*
FrameActionView:ViewTabButton
    문서탭 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) ViewTabButton() {
	p.hwp.Run("ViewTabButton")
}

/*
FrameActionView:HwpViewType
    문서창 모양 설정
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpViewType() {
	p.hwp.Run("HwpViewType")
}

/*
FrameActionView:ChangeSkin
    스킨 바꾸기
    파라미터셋: 없음
*/
func (p *FrameActionView) ChangeSkin() {
	p.hwp.Run("ChangeSkin")
}

/*
FrameActionView:ShowAttributeTab
    속성 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) ShowAttributeTab() {
	p.hwp.Run("ShowAttributeTab")
}

/*
FrameActionView:ShowScriptTab
    스크립트 작업창 보이기/감추기
    파라미터셋: 없음
*/
func (p *FrameActionView) ShowScriptTab() {
	p.hwp.Run("ShowScriptTab")
}

/*
FrameActionView:FrameViewZoomRibon
    화면 확대/축소
    파라미터셋: 없음
*/
func (p *FrameActionView) FrameViewZoomRibon() {
	p.hwp.Run("FrameViewZoomRibon")
}

/*
FrameActionView:HwpTabViewHwpDic
    사전 검색 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewHwpDic() {
	p.hwp.Run("HwpTabViewHwpDic")
}

/*
FrameActionView:HwpTabViewNLP
    글 도우미 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewNLP() {
	p.hwp.Run("HwpTabViewNLP")
}

/*
FrameActionView:HwpTabViewOutline
    개요 보기 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewOutline() {
	p.hwp.Run("HwpTabViewOutline")
}

/*
FrameActionView:HwpTabViewMasterPage
    바탕쪽 보기 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewMasterPage() {
	p.hwp.Run("HwpTabViewMasterPage")
}

/*
FrameActionView:HwpTabViewAction
    빠른 실행 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewAction() {
	p.hwp.Run("HwpTabViewAction")
}

/*
FrameActionView:HwpTabViewDistant
    쪽모양 보기 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewDistant() {
	p.hwp.Run("HwpTabViewDistant")
}

/*
FrameActionView:HwpTabViewClipboard
    클립보드 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewClipboard() {
	p.hwp.Run("HwpTabViewClipboard")
}

/*
FrameActionView:HwpTabViewAttribute
    양식 개체 속성 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewAttribute() {
	p.hwp.Run("HwpTabViewAttribute")
}

/*
FrameActionView:HwpTabViewScript
    스크립트 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewScript() {
	p.hwp.Run("HwpTabViewScript")
}

/*
FrameActionView:HwpTabViewXMLTree
    XML 문서 트리 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewXMLTree() {
	p.hwp.Run("HwpTabViewXMLTree")
}

/*
FrameActionView:HwpTabViewXMLSchema
    XML 문서 구조 작업창
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpTabViewXMLSchema() {
	p.hwp.Run("HwpTabViewXMLSchema")
}

/*
FrameActionView:HwpWSDic
    사전 검색 작업창 (Shift + F12)
    파라미터셋: 없음
*/
func (p *FrameActionView) HwpWSDic() {
	p.hwp.Run("HwpWSDic")
}

/*
ActionFile:FileTemplate
    문서마당
    파라미터셋: FileOpen
*/
func (p *ActionFile) FileTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileTemplate", arg...)
}

/*
ActionFile:DocSummaryInfo
    문서 정보
    파라미터셋: SummaryInfo
*/
func (p *ActionFile) DocSummaryInfo(arg ...interface{}) {
	p.hwp.ActionWithParameters("DocSummaryInfo", arg...)
}

/*
ActionFile:DocumentInfo
    현재 문서에 대한 정보
    파라미터셋: DocumentInfo
*/
func (p *ActionFile) DocumentInfo(arg ...interface{}) {
	p.hwp.ActionWithParameters("DocumentInfo", arg...)
}

/*
ActionFile:LinkDocument
    문서 연결
    파라미터셋: LinkDocument
*/
func (p *ActionFile) LinkDocument(arg ...interface{}) {
	p.hwp.ActionWithParameters("LinkDocument", arg...)
}

/*
ActionFile:Print
    인쇄
    파라미터셋: Print
*/
func (p *ActionFile) Print(arg ...interface{}) {
	p.hwp.ActionWithParameters("Print", arg...)
}

/*
ActionFile:Preference
    환경 설정
    파라미터셋: Preference
*/
func (p *ActionFile) Preference(arg ...interface{}) {
	p.hwp.ActionWithParameters("Preference", arg...)
}

/*
ActionFile:SendMailAttach
    편지 보내기-첨부 파일로
    파라미터셋: FileSendMail
*/
func (p *ActionFile) SendMailAttach(arg ...interface{}) {
	p.hwp.ActionWithParameters("SendMailAttach", arg...)
}

/*
ActionFile:SendMailText
    편지 보내기-본문으로
    파라미터셋: FileSendMail
*/
func (p *ActionFile) SendMailText(arg ...interface{}) {
	p.hwp.ActionWithParameters("SendMailText", arg...)
}

/*
ActionFile:SendBrowserText
    브라우저로 보내기
    파라미터셋: 없음
*/
func (p *ActionFile) SendBrowserText() {
	p.hwp.Run("SendBrowserText")
}

/*
ActionFile:PrintToImage
    그림으로 저장하기
    파라미터셋: PrintToImage
*/
func (p *ActionFile) PrintToImage(arg ...interface{}) {
	p.hwp.ActionWithParameters("PrintToImage", arg...)
}

/*
ActionFile:PrintToFax
    UMS(fax)로 인쇄하기
    파라미터셋: Print
*/
func (p *ActionFile) PrintToFax(arg ...interface{}) {
	p.hwp.ActionWithParameters("PrintToFax", arg...)
}

/*
ActionFile:PrintToPDF
    Adobe PDF로 인쇄하기
    파라미터셋: Print
*/
func (p *ActionFile) PrintToPDF(arg ...interface{}) {
	p.hwp.ActionWithParameters("PrintToPDF", arg...)
}

/*
ActionFile:FtpUpload
    웹 서버로 올리기
    파라미터셋: FtpUpload
*/
func (p *ActionFile) FtpUpload(arg ...interface{}) {
	p.hwp.ActionWithParameters("FtpUpload", arg...)
}

/*
ActionFile:FilePassword
    문서 암호
    파라미터셋: Password
*/
func (p *ActionFile) FilePassword(arg ...interface{}) {
	p.hwp.ActionWithParameters("FilePassword", arg...)
}

/*
ActionFile:FileSecurity
    배포용 문서로 저장하기
    파라미터셋: FileSecurity
*/
func (p *ActionFile) FileSecurity(arg ...interface{}) {
	p.hwp.ActionWithParameters("FileSecurity", arg...)
}

/*
ActionFile:SaveBlockAction
    블록 저장하기
    파라미터셋: FileSaveBlock
*/
func (p *ActionFile) SaveBlockAction(arg ...interface{}) {
	p.hwp.ActionWithParameters("SaveBlockAction", arg...)
}

/*
ActionFile:VersionInfo
    버전 관리
    파라미터셋: VersionInfo
*/
func (p *ActionFile) VersionInfo(arg ...interface{}) {
	p.hwp.ActionWithParameters("VersionInfo", arg...)
}

/*
ActionFile:VersionSave
    새 버전으로 저장
    파라미터셋: VersionInfo
*/
func (p *ActionFile) VersionSave(arg ...interface{}) {
	p.hwp.ActionWithParameters("VersionSave", arg...)
}

/*
ActionFile:ASendBrowserText
    웹브라우저로 보내기
    파라미터셋: 없음
*/
func (p *ActionFile) ASendBrowserText() {
	p.hwp.Run("ASendBrowserText")
}

/*
ActionFile:ALoadUserInfoFile
    사용자 환경 파일 읽어오기
    파라미터셋: LoadUserInfoFile
*/
func (p *ActionFile) ALoadUserInfoFile(arg ...interface{}) {
	p.hwp.ActionWithParameters("ALoadUserInfoFile", arg...)
}

/*
ActionEdit:MoveLeft
    왼쪽으로 캐럿 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveLeft() {
	p.hwp.Run("MoveLeft")
}

/*
ActionEdit:MoveRight
    오른쪽으로 캐럿 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveRight() {
	p.hwp.Run("MoveRight")
}

/*
ActionEdit:MoveUp
    위쪽으로 캐럿 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveUp() {
	p.hwp.Run("MoveUp")
}

/*
ActionEdit:MoveDown
    아래쪽으로 캐럿 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveDown() {
	p.hwp.Run("MoveDown")
}

/*
ActionEdit:MoveSelLeft
    블록 설정 상태로 왼쪽으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelLeft() {
	p.hwp.Run("MoveSelLeft")
}

/*
ActionEdit:MoveSelRight
    블록 설정 상태로 오른쪽으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelRight() {
	p.hwp.Run("MoveSelRight")
}

/*
ActionEdit:MoveSelUp
    블록 설정 상태로 위쪽으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelUp() {
	p.hwp.Run("MoveSelUp")
}

/*
ActionEdit:MoveSelDown
    블록 설정 상태로 아래쪽으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelDown() {
	p.hwp.Run("MoveSelDown")
}

/*
ActionEdit:MovePrevChar
    한 글자 앞 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevChar() {
	p.hwp.Run("MovePrevChar")
}

/*
ActionEdit:MoveNextChar
    한 글자 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextChar() {
	p.hwp.Run("MoveNextChar")
}

/*
ActionEdit:MovePrevPos
    한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevPos() {
	p.hwp.Run("MovePrevPos")
}

/*
ActionEdit:MoveNextPos
    한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextPos() {
	p.hwp.Run("MoveNextPos")
}

/*
ActionEdit:MovePrevPosEx
    한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevPosEx() {
	p.hwp.Run("MovePrevPosEx")
}

/*
ActionEdit:MoveNextPosEx
    한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextPosEx() {
	p.hwp.Run("MoveNextPosEx")
}

/*
ActionEdit:MovePrevWord
    한 단어 앞으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevWord() {
	p.hwp.Run("MovePrevWord")
}

/*
ActionEdit:MoveNextWord
    한 단어 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextWord() {
	p.hwp.Run("MoveNextWord")
}

/*
ActionEdit:MoveWordBegin
    현재 위치한 단어의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveWordBegin() {
	p.hwp.Run("MoveWordBegin")
}

/*
ActionEdit:MoveWordEnd
    현재 위치한 단어의 끝으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveWordEnd() {
	p.hwp.Run("MoveWordEnd")
}

/*
ActionEdit:MoveDocBegin
    만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 문서의 시작으로 이동. 현재 서브 리스트 내에 있으면 빠져나간다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveDocBegin() {
	p.hwp.Run("MoveDocBegin")
}

/*
ActionEdit:MoveDocEnd
    만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 문서의 끝으로 이동. 현재 서브 리스트 내에 있으면 빠져나간다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveDocEnd() {
	p.hwp.Run("MoveDocEnd")
}

/*
ActionEdit:MoveLineBegin
    현재 위치한 줄의 시작/끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveLineBegin() {
	p.hwp.Run("MoveLineBegin")
}

/*
ActionEdit:MoveLineEnd
    현재 위치한 줄의 시작/끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveLineEnd() {
	p.hwp.Run("MoveLineEnd")
}

/*
ActionEdit:MoveParaBegin
    현재 위치한 문단의 시작/끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveParaBegin() {
	p.hwp.Run("MoveParaBegin")
}

/*
ActionEdit:MoveParaEnd
    현재 위치한 문단의 시작/끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveParaEnd() {
	p.hwp.Run("MoveParaEnd")
}

/*
ActionEdit:MovePageUp
    앞/뒤 페이지의 시작으로 이동. 현재 탑 레벨 리스트가 아니면 탑 레벨 리스트로 빠져나온다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePageUp() {
	p.hwp.Run("MovePageUp")
}

/*
ActionEdit:MovePageDown
    앞/뒤 페이지의 시작으로 이동. 현재 탑 레벨 리스트가 아니면 탑 레벨 리스트로 빠져나온다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePageDown() {
	p.hwp.Run("MovePageDown")
}

/*
ActionEdit:MovePageBegin
    쪽 맨 처음으로
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePageBegin() {
	p.hwp.Run("MovePageBegin")
}

/*
ActionEdit:MovePageEnd
    쪽 맨 끝으로
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePageEnd() {
	p.hwp.Run("MovePageEnd")
}

/*
ActionEdit:MoveListBegin
    현재 리스트의 시작으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveListBegin() {
	p.hwp.Run("MoveListBegin")
}

/*
ActionEdit:MoveListEnd
    현재 리스트의 끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveListEnd() {
	p.hwp.Run("MoveListEnd")
}

/*
ActionEdit:MoveTopLevelBegin
    탑 레벨 리스트의 시작으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveTopLevelBegin() {
	p.hwp.Run("MoveTopLevelBegin")
}

/*
ActionEdit:MoveTopLevelEnd
    탑 레벨 리스트의 끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveTopLevelEnd() {
	p.hwp.Run("MoveTopLevelEnd")
}

/*
ActionEdit:MovePrevParaEnd
    앞 문단의 끝/다음 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevParaEnd() {
	p.hwp.Run("MovePrevParaEnd")
}

/*
ActionEdit:MoveNextParaBegin
    앞 문단의 끝/다음 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextParaBegin() {
	p.hwp.Run("MoveNextParaBegin")
}

/*
ActionEdit:MovePrevParaBegin
    앞 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevParaBegin() {
	p.hwp.Run("MovePrevParaBegin")
}

/*
ActionEdit:MoveSectionUp
    앞/뒤 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSectionUp() {
	p.hwp.Run("MoveSectionUp")
}

/*
ActionEdit:MoveSectionDown
    앞/뒤 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSectionDown() {
	p.hwp.Run("MoveSectionDown")
}

/*
ActionEdit:MoveParentList
    한 레벨 상위/탑 레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴 한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveParentList() {
	p.hwp.Run("MoveParentList")
}

/*
ActionEdit:MoveTopLevelList
    한 레벨 상위/탑 레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveTopLevelList() {
	p.hwp.Run("MoveTopLevelList")
}

/*
ActionEdit:MoveRootList
    한 레벨 상위/탑 레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴 한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveRootList() {
	p.hwp.Run("MoveRootList")
}

/*
ActionEdit:MoveSelPrevChar
    셀렉션: 이전 글자
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPrevChar() {
	p.hwp.Run("MoveSelPrevChar")
}

/*
ActionEdit:MoveSelNextChar
    셀렉션: 다음 글자
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelNextChar() {
	p.hwp.Run("MoveSelNextChar")
}

/*
ActionEdit:MoveSelPrevPos
    셀렉션: 이전 위치
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPrevPos() {
	p.hwp.Run("MoveSelPrevPos")
}

/*
ActionEdit:MoveSelNextPos
    셀렉션: 다음 위치
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelNextPos() {
	p.hwp.Run("MoveSelNextPos")
}

/*
ActionEdit:MoveSelPrevWord
    셀렉션: 이전 단어
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPrevWord() {
	p.hwp.Run("MoveSelPrevWord")
}

/*
ActionEdit:MoveSelNextWord
    셀렉션: 다음 단어
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelNextWord() {
	p.hwp.Run("MoveSelNextWord")
}

/*
ActionEdit:MoveSelWordBegin
    셀렉션: 단어 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelWordBegin() {
	p.hwp.Run("MoveSelWordBegin")
}

/*
ActionEdit:MoveSelWordEnd
    셀렉션: 단어 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelWordEnd() {
	p.hwp.Run("MoveSelWordEnd")
}

/*
ActionEdit:MoveSelDocBegin
    셀렉션: 문서 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelDocBegin() {
	p.hwp.Run("MoveSelDocBegin")
}

/*
ActionEdit:MoveSelDocEnd
    셀렉션: 문서 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelDocEnd() {
	p.hwp.Run("MoveSelDocEnd")
}

/*
ActionEdit:MoveSelLineBegin
    셀렉션: 줄 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelLineBegin() {
	p.hwp.Run("MoveSelLineBegin")
}

/*
ActionEdit:MoveSelLineEnd
    셀렉션: 줄 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelLineEnd() {
	p.hwp.Run("MoveSelLineEnd")
}

/*
ActionEdit:MoveSelParaBegin
    셀렉션: 문단 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelParaBegin() {
	p.hwp.Run("MoveSelParaBegin")
}

/*
ActionEdit:MoveSelParaEnd
    셀렉션: 문단 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelParaEnd() {
	p.hwp.Run("MoveSelParaEnd")
}

/*
ActionEdit:MoveSelPageUp
    셀렉션: 페이지 업
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPageUp() {
	p.hwp.Run("MoveSelPageUp")
}

/*
ActionEdit:MoveSelPageDown
    셀렉션: 페이지 다운
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPageDown() {
	p.hwp.Run("MoveSelPageDown")
}

/*
ActionEdit:MoveSelListBegin
    셀렉션: 리스트 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelListBegin() {
	p.hwp.Run("MoveSelListBegin")
}

/*
ActionEdit:MoveSelListEnd
    셀렉션: 리스트 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelListEnd() {
	p.hwp.Run("MoveSelListEnd")
}

/*
ActionEdit:MoveSelTopLevelBegin
    셀렉션: 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelTopLevelBegin() {
	p.hwp.Run("MoveSelTopLevelBegin")
}

/*
ActionEdit:MoveSelTopLevelEnd
    셀렉션: 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelTopLevelEnd() {
	p.hwp.Run("MoveSelTopLevelEnd")
}

/*
ActionEdit:MoveSelPrevParaEnd
    셀렉션: 이전 문단 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPrevParaEnd() {
	p.hwp.Run("MoveSelPrevParaEnd")
}

/*
ActionEdit:MoveSelNextParaBegin
    셀렉션: 다음 문단 처음
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelNextParaBegin() {
	p.hwp.Run("MoveSelNextParaBegin")
}

/*
ActionEdit:MoveSelPrevParaBegin
    셀렉션: 이전 문단 시작
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelPrevParaBegin() {
	p.hwp.Run("MoveSelPrevParaBegin")
}

/*
ActionEdit:MoveLineUp
    한 줄 위/아래로 이동한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveLineUp() {
	p.hwp.Run("MoveLineUp")
}

/*
ActionEdit:MoveLineDown
    한 줄 위/아래로 이동한다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveLineDown() {
	p.hwp.Run("MoveLineDown")
}

/*
ActionEdit:MoveViewUp
    현재 뷰의 크기만큼 위/아래로 이동한다. PgUp/PgDn 키의 기능이다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveViewUp() {
	p.hwp.Run("MoveViewUp")
}

/*
ActionEdit:MoveViewDown
    현재 뷰의 크기만큼 위/아래로 이동한다. PgUp/PgDn 키의 기능이다.
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveViewDown() {
	p.hwp.Run("MoveViewDown")
}

/*
ActionEdit:MoveViewBegin
    현재 뷰의 시작/끝에 위치한 곳으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveViewBegin() {
	p.hwp.Run("MoveViewBegin")
}

/*
ActionEdit:MoveViewEnd
    현재 뷰의 시작/끝에 위치한 곳으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveViewEnd() {
	p.hwp.Run("MoveViewEnd")
}

/*
ActionEdit:MovePrevColumn
    앞 단으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MovePrevColumn() {
	p.hwp.Run("MovePrevColumn")
}

/*
ActionEdit:MoveNextColumn
    뒤 단으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveNextColumn() {
	p.hwp.Run("MoveNextColumn")
}

/*
ActionEdit:MoveColumnBegin
    단의 처음으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveColumnBegin() {
	p.hwp.Run("MoveColumnBegin")
}

/*
ActionEdit:MoveColumnEnd
    단의 끝으로 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveColumnEnd() {
	p.hwp.Run("MoveColumnEnd")
}

/*
ActionEdit:MoveScrollPrev
    이전 방향으로 스크롤하면서 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveScrollPrev() {
	p.hwp.Run("MoveScrollPrev")
}

/*
ActionEdit:MoveScrollNext
    다음 방향으로 스크롤하면서 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveScrollNext() {
	p.hwp.Run("MoveScrollNext")
}

/*
ActionEdit:MoveScrollUp
    위 방향으로 스크롤하면서 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveScrollUp() {
	p.hwp.Run("MoveScrollUp")
}

/*
ActionEdit:MoveScrollDown
    아래 방향으로 스크롤하면서 이동
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveScrollDown() {
	p.hwp.Run("MoveScrollDown")
}

/*
ActionEdit:MoveSelLineUp
    셀렉션: 한줄 위
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelLineUp() {
	p.hwp.Run("MoveSelLineUp")
}

/*
ActionEdit:MoveSelLineDown
    셀렉션: 한줄 아래
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelLineDown() {
	p.hwp.Run("MoveSelLineDown")
}

/*
ActionEdit:MoveSelViewUp
    셀렉션: 위
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelViewUp() {
	p.hwp.Run("MoveSelViewUp")
}

/*
ActionEdit:MoveSelViewDown
    셀렉션: 아래
    파라미터셋: 없음
*/
func (p *ActionEdit) MoveSelViewDown() {
	p.hwp.Run("MoveSelViewDown")
}

/*
ActionEdit:Close
    레퍼런스 포인트가 등록되어 있으면 그 포인트로, 없으면 루트 리스트로 이동한다. 나머지특성은 MoveRootList와 동일
    파라미터셋: 없음
*/
func (p *ActionEdit) Close() {
	p.hwp.Run("Close")
}

/*
ActionEdit:CloseEx
    레퍼런스 포인트가 등록되어 있으면 그 포인트로, 없으면 루트 리스트로 이동한다. 이동할 때 캐럿이 위치할 수 있는 곳을 찾아서 리스트를 빠져나간다. 나머지특성은 MoveRootList와 동일
    파라미터셋: 없음
*/
func (p *ActionEdit) CloseEx() {
	p.hwp.Run("CloseEx")
}

/*
ActionEdit:TableLeftCell
    셀이동: 셀 왼쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableLeftCell() {
	p.hwp.Run("TableLeftCell")
}

/*
ActionEdit:TableRightCell
    셀이동: 셀 오른쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableRightCell() {
	p.hwp.Run("TableRightCell")
}

/*
ActionEdit:TableRightCellAppend
    셀이동: 셀 오른쪽에 이어서
    파라미터셋: 없음
*/
func (p *ActionEdit) TableRightCellAppend() {
	p.hwp.Run("TableRightCellAppend")
}

/*
ActionEdit:TableUpperCell
    셀이동: 셀 위
    파라미터셋: 없음
*/
func (p *ActionEdit) TableUpperCell() {
	p.hwp.Run("TableUpperCell")
}

/*
ActionEdit:TableLowerCell
    셀이동: 셀 아래
    파라미터셋: 없음
*/
func (p *ActionEdit) TableLowerCell() {
	p.hwp.Run("TableLowerCell")
}

/*
ActionEdit:TableColBegin
    셀이동: 열 시작
    파라미터셋: 없음
*/
func (p *ActionEdit) TableColBegin() {
	p.hwp.Run("TableColBegin")
}

/*
ActionEdit:TableColEnd
    셀이동: 열 끝
    파라미터셋: 없음
*/
func (p *ActionEdit) TableColEnd() {
	p.hwp.Run("TableColEnd")
}

/*
ActionEdit:TableColPageUp
    셀이동: 페이지 업
    파라미터셋: 없음
*/
func (p *ActionEdit) TableColPageUp() {
	p.hwp.Run("TableColPageUp")
}

/*
ActionEdit:TableColPageDown
    셀이동: 페이지 다운
    파라미터셋: 없음
*/
func (p *ActionEdit) TableColPageDown() {
	p.hwp.Run("TableColPageDown")
}

/*
ActionEdit:TableResizeLeft
    셀 크기 변경
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeLeft() {
	p.hwp.Run("TableResizeLeft")
}

/*
ActionEdit:TableResizeRight
    셀 크기 변경
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeRight() {
	p.hwp.Run("TableResizeRight")
}

/*
ActionEdit:TableResizeUp
    셀 크기 변경
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeUp() {
	p.hwp.Run("TableResizeUp")
}

/*
ActionEdit:TableResizeDown
    셀 크기 변경
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeDown() {
	p.hwp.Run("TableResizeDown")
}

/*
ActionEdit:TableResizeCellLeft
    셀 크기 변경: 셀 왼쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeCellLeft() {
	p.hwp.Run("TableResizeCellLeft")
}

/*
ActionEdit:TableResizeCellRight
    셀 크기 변경: 셀 오른쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeCellRight() {
	p.hwp.Run("TableResizeCellRight")
}

/*
ActionEdit:TableResizeCellUp
    셀 크기 변경: 셀 위
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeCellUp() {
	p.hwp.Run("TableResizeCellUp")
}

/*
ActionEdit:TableResizeCellDown
    셀 크기 변경: 셀 아래
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeCellDown() {
	p.hwp.Run("TableResizeCellDown")
}

/*
ActionEdit:TableResizeLineLeft
    셀 크기 변경: 선 왼쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeLineLeft() {
	p.hwp.Run("TableResizeLineLeft")
}

/*
ActionEdit:TableResizeLineRight
    셀 크기 변경: 선 오른쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeLineRight() {
	p.hwp.Run("TableResizeLineRight")
}

/*
ActionEdit:TableResizeLineUp
    셀 크기 변경: 선 위
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeLineUp() {
	p.hwp.Run("TableResizeLineUp")
}

/*
ActionEdit:TableResizeLineDown
    셀 크기 변경: 선 아래
    파라미터셋: 없음
*/
func (p *ActionEdit) TableResizeLineDown() {
	p.hwp.Run("TableResizeLineDown")
}

/*
ActionEdit:HiddenCredits
    인터넷 정보
    파라미터셋: 없음
*/
func (p *ActionEdit) HiddenCredits() {
	p.hwp.Run("HiddenCredits")
}

/*
ActionEdit:Undo
    되살리기
    파라미터셋: 없음
*/
func (p *ActionEdit) Undo() {
	p.hwp.Run("Undo")
}

/*
ActionEdit:Redo
    다시 실행
    파라미터셋: 없음
*/
func (p *ActionEdit) Redo() {
	p.hwp.Run("Redo")
}

/*
ActionEdit:Cut
    오려 두기
    파라미터셋: SelectionOpt
*/
func (p *ActionEdit) Cut(arg ...interface{}) {
	p.hwp.ActionWithParameters("Cut", arg...)
}

/*
ActionEdit:Copy
    복사하기
    파라미터셋: SelectionOpt
*/
func (p *ActionEdit) Copy(arg ...interface{}) {
	p.hwp.ActionWithParameters("Copy", arg...)
}

/*
ActionEdit:Erase
    지우기
    파라미터셋: SelectionOpt
*/
func (p *ActionEdit) Erase(arg ...interface{}) {
	p.hwp.ActionWithParameters("Erase", arg...)
}

/*
ActionEdit:Paste
    붙이기
    파라미터셋: SelectionOpt
*/
func (p *ActionEdit) Paste(arg ...interface{}) {
	p.hwp.ActionWithParameters("Paste", arg...)
}

/*
ActionEdit:PasteSpecial
    골라 붙이기
    파라미터셋: 없음
*/
func (p *ActionEdit) PasteSpecial() {
	p.hwp.Run("PasteSpecial")
}

/*
ActionEdit:SelectAll
    모두 선택
    파라미터셋: 없음
*/
func (p *ActionEdit) SelectAll() {
	p.hwp.Run("SelectAll")
}

/*
ActionEdit:Select
    블록 설정
    파라미터셋: 없음
*/
func (p *ActionEdit) Select() {
	p.hwp.Run("Select")
}

/*
ActionEdit:SelectColumn
    칸 블록 설정
    파라미터셋: 없음
*/
func (p *ActionEdit) SelectColumn() {
	p.hwp.Run("SelectColumn")
}

/*
ActionEdit:Delete
    Delete
    파라미터셋: 없음
*/
func (p *ActionEdit) Delete() {
	p.hwp.Run("Delete")
}

/*
ActionEdit:DeleteBack
    Backspace
    파라미터셋: 없음
*/
func (p *ActionEdit) DeleteBack() {
	p.hwp.Run("DeleteBack")
}

/*
ActionEdit:DeleteLine
    CTRL-Y (한줄 지우기)
    파라미터셋: 없음
*/
func (p *ActionEdit) DeleteLine() {
	p.hwp.Run("DeleteLine")
}

/*
ActionEdit:DeleteLineEnd
    현재 캐럿 이후 지우기
    파라미터셋: 없음
*/
func (p *ActionEdit) DeleteLineEnd() {
	p.hwp.Run("DeleteLineEnd")
}

/*
ActionEdit:DeleteWord
    단어 지우기 CTRL-T
    파라미터셋: 없음
*/
func (p *ActionEdit) DeleteWord() {
	p.hwp.Run("DeleteWord")
}

/*
ActionEdit:DeleteWordBack
    왼쪽 단어 지우기 CTRL-BS
    파라미터셋: 없음
*/
func (p *ActionEdit) DeleteWordBack() {
	p.hwp.Run("DeleteWordBack")
}

/*
ActionEdit:ToggleOverwrite
    삽입/수정
    파라미터셋: 없음
*/
func (p *ActionEdit) ToggleOverwrite() {
	p.hwp.Run("ToggleOverwrite")
}

/*
ActionEdit:Cancel
    취소 ESC
    파라미터셋: 없음
*/
func (p *ActionEdit) Cancel() {
	p.hwp.Run("Cancel")
}

/*
ActionEdit:FindDlg
    찾기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) FindDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("FindDlg", arg...)
}

/*
ActionEdit:RepeatFind
    다시 찾기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) RepeatFind(arg ...interface{}) {
	p.hwp.ActionWithParameters("RepeatFind", arg...)
}

/*
ActionEdit:ReverseFind
    거꾸로 찾기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) ReverseFind(arg ...interface{}) {
	p.hwp.ActionWithParameters("ReverseFind", arg...)
}

/*
ActionEdit:ForwardFind
    앞으로 찾기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) ForwardFind(arg ...interface{}) {
	p.hwp.ActionWithParameters("ForwardFind", arg...)
}

/*
ActionEdit:BackwardFind
    뒤로 찾기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) BackwardFind(arg ...interface{}) {
	p.hwp.ActionWithParameters("BackwardFind", arg...)
}

/*
ActionEdit:ReplaceDlg
    찾아 바꾸기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) ReplaceDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("ReplaceDlg", arg...)
}

/*
ActionEdit:ExecReplace
    바꾸기(실행)
    파라미터셋: FindReplace
*/
func (p *ActionEdit) ExecReplace(arg ...interface{}) {
	p.hwp.ActionWithParameters("ExecReplace", arg...)
}

/*
ActionEdit:AllReplace
    모두 바꾸기
    파라미터셋: FindReplace
*/
func (p *ActionEdit) AllReplace(arg ...interface{}) {
	p.hwp.ActionWithParameters("AllReplace", arg...)
}

/*
ActionEdit:DocFindInit
    문서 찾기 초기화
    파라미터셋: FindReplace
*/
func (p *ActionEdit) DocFindInit(arg ...interface{}) {
	p.hwp.ActionWithParameters("DocFindInit", arg...)
}

/*
ActionEdit:DocFindNext
    문서 찾기 계속
    파라미터셋: DocFindInfo
*/
func (p *ActionEdit) DocFindNext(arg ...interface{}) {
	p.hwp.ActionWithParameters("DocFindNext", arg...)
}

/*
ActionEdit:DocFindEnd
    문서 찾기 종료
    파라미터셋: FindReplace
*/
func (p *ActionEdit) DocFindEnd(arg ...interface{}) {
	p.hwp.ActionWithParameters("DocFindEnd", arg...)
}

/*
ActionEdit:Goto
    찾아가기
    파라미터셋: GotoE
*/
func (p *ActionEdit) Goto(arg ...interface{}) {
	p.hwp.ActionWithParameters("Goto", arg...)
}

/*
ActionEdit:GotoStyle
    찾아가기-스타일 (찾기/바꾸기 대화상자에서 사용)
    파라미터셋: GotoE
*/
func (p *ActionEdit) GotoStyle(arg ...interface{}) {
	p.hwp.ActionWithParameters("GotoStyle", arg...)
}

/*
ActionEdit:FindForeBackFind
    앞뒤로 찾아가기 : 찾기
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackFind() {
	p.hwp.Run("FindForeBackFind")
}

/*
ActionEdit:FindForeBackPage
    앞뒤로 찾아가기 : 쪽
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackPage() {
	p.hwp.Run("FindForeBackPage")
}

/*
ActionEdit:FindForeBackSection
    앞뒤로 찾아가기 : 구역
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackSection() {
	p.hwp.Run("FindForeBackSection")
}

/*
ActionEdit:FindForeBackLine
    앞뒤로 찾아가기 : 줄
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackLine() {
	p.hwp.Run("FindForeBackLine")
}

/*
ActionEdit:FindForeBackStyle
    앞뒤로 찾아가기 : 스타일
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackStyle() {
	p.hwp.Run("FindForeBackStyle")
}

/*
ActionEdit:FindForeBackCtrl
    앞뒤로 찾아가기 : 조판부호
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackCtrl() {
	p.hwp.Run("FindForeBackCtrl")
}

/*
ActionEdit:FindForeBackBookmark
    앞뒤로 찾아가기 : 책갈피
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackBookmark() {
	p.hwp.Run("FindForeBackBookmark")
}

/*
ActionEdit:FindForeBackSelectCtrl
    앞뒤로 찾아가기 : 조판 부호 찾기 (선택)
    파라미터셋: 없음
*/
func (p *ActionEdit) FindForeBackSelectCtrl() {
	p.hwp.Run("FindForeBackSelectCtrl")
}

/*
ActionEdit:ConvertToHangul
    한글로
    파라미터셋: ConvertToHangul
*/
func (p *ActionEdit) ConvertToHangul(arg ...interface{}) {
	p.hwp.ActionWithParameters("ConvertToHangul", arg...)
}

/*
ActionEdit:ConvertHiraGata
    일어 바꾸기
    파라미터셋: ConvertHiraToGata
*/
func (p *ActionEdit) ConvertHiraGata(arg ...interface{}) {
	p.hwp.ActionWithParameters("ConvertHiraGata", arg...)
}

/*
ActionEdit:ConvertJianFan
    간체/번체 바꾸기
    파라미터셋: ConvertJianFan
*/
func (p *ActionEdit) ConvertJianFan(arg ...interface{}) {
	p.hwp.ActionWithParameters("ConvertJianFan", arg...)
}

/*
ActionEdit:ConvertCase
    대소문자 바꾸기
    파라미터셋: ConvertCase
*/
func (p *ActionEdit) ConvertCase(arg ...interface{}) {
	p.hwp.ActionWithParameters("ConvertCase", arg...)
}

/*
ActionEdit:ConvertFullHalfWidth
    전각 반각 바꾸기
    파라미터셋: ConvertFullHalf
*/
func (p *ActionEdit) ConvertFullHalfWidth(arg ...interface{}) {
	p.hwp.ActionWithParameters("ConvertFullHalfWidth", arg...)
}

/*
ActionEdit:GetSectionApplyString
    유틸리티 액션
    파라미터셋: SectionApply
*/
func (p *ActionEdit) GetSectionApplyString(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetSectionApplyString", arg...)
}

/*
ActionEdit:GetSectionApplyTo
    유틸리티 액션
    파라미터셋: SectionApply
*/
func (p *ActionEdit) GetSectionApplyTo(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetSectionApplyTo", arg...)
}

/*
ActionEdit:GetDocFilters
    유틸리티 액션
    파라미터셋: DocFilters
*/
func (p *ActionEdit) GetDocFilters(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetDocFilters", arg...)
}

/*
ActionEdit:ModifyShapeObject
    고치기-개체 속성
    파라미터셋: 없음
*/
func (p *ActionEdit) ModifyShapeObject() {
	p.hwp.Run("ModifyShapeObject")
}

/*
ActionEdit:ModifyCtrl
    고치기-컨트롤
    파라미터셋: 없음
*/
func (p *ActionEdit) ModifyCtrl() {
	p.hwp.Run("ModifyCtrl")
}

/*
ActionEdit:ModifyLineProperty
    고치기(선 속성 탭으로)
    파라미터셋: 없음
*/
func (p *ActionEdit) ModifyLineProperty() {
	p.hwp.Run("ModifyLineProperty")
}

/*
ActionEdit:ModifyFillProperty
    고치기(채우기 속성 탭으로)
    파라미터셋: 없음
*/
func (p *ActionEdit) ModifyFillProperty() {
	p.hwp.Run("ModifyFillProperty")
}

/*
ActionEdit:SelectCtrlReverse
    개체선택 역방향
    파라미터셋: 없음
*/
func (p *ActionEdit) SelectCtrlReverse() {
	p.hwp.Run("SelectCtrlReverse")
}

/*
ActionEdit:SelectCtrlFront
    개체선택 정방향
    파라미터셋: 없음
*/
func (p *ActionEdit) SelectCtrlFront() {
	p.hwp.Run("SelectCtrlFront")
}

/*
ActionEdit:PasteHtmlDialog
    HTML 문서 붙이기
    파라미터셋: PasteHtml
*/
func (p *ActionEdit) PasteHtmlDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("PasteHtmlDialog", arg...)
}

/*
ActionShape:CharShapeDialog
    글자 모양 대화상자(내부 구현용)
    파라미터셋: CharShape
*/
func (p *ActionShape) CharShapeDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("CharShapeDialog", arg...)
}

/*
ActionShape:CharShape
    글자 모양
    파라미터셋: CharShape
*/
func (p *ActionShape) CharShape(arg ...interface{}) {
	p.hwp.ActionWithParameters("CharShape", arg...)
}

/*
ActionShape:CharShapeNormal
    보통모양 ALT+SHIFT+C
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeNormal() {
	p.hwp.Run("CharShapeNormal")
}

/*
ActionShape:CharShapeBold
    단축키: Alt+L(글자 진하게)
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeBold() {
	p.hwp.Run("CharShapeBold")
}

/*
ActionShape:CharShapeItalic
    이탤릭 ALT + SHIFT + I
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeItalic() {
	p.hwp.Run("CharShapeItalic")
}

/*
ActionShape:CharShapeUnderline
    밑줄 ALT+SHIFT+U
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeUnderline() {
	p.hwp.Run("CharShapeUnderline")
}

/*
ActionShape:CharShapeCenterline
    취소선
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeCenterline() {
	p.hwp.Run("CharShapeCenterline")
}

/*
ActionShape:CharShapeOutline
    외곽선
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeOutline() {
	p.hwp.Run("CharShapeOutline")
}

/*
ActionShape:CharShapeShadow
    그림자
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeShadow() {
	p.hwp.Run("CharShapeShadow")
}

/*
ActionShape:CharShapeEmboss
    양각
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeEmboss() {
	p.hwp.Run("CharShapeEmboss")
}

/*
ActionShape:CharShapeEngrave
    음각
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeEngrave() {
	p.hwp.Run("CharShapeEngrave")
}

/*
ActionShape:CharShapeSuperscript
    위첨자 ALT+SHIFT+P
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeSuperscript() {
	p.hwp.Run("CharShapeSuperscript")
}

/*
ActionShape:CharShapeSubscript
    아래첨자 ALT+SHIFT+S
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeSubscript() {
	p.hwp.Run("CharShapeSubscript")
}

/*
ActionShape:CharShapeSuperSubscript
    첨자 "위첨자→아래첨자→보통" 반복
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeSuperSubscript() {
	p.hwp.Run("CharShapeSuperSubscript")
}

/*
ActionShape:CharShapeHeightIncrease
    크기 크게 ALT+SHIFT+E
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeHeightIncrease() {
	p.hwp.Run("CharShapeHeightIncrease")
}

/*
ActionShape:CharShapeHeightDecrease
    크기 작게 ALT+SHIFT+R
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeHeightDecrease() {
	p.hwp.Run("CharShapeHeightDecrease")
}

/*
ActionShape:CharShapeWidthIncrease
    장평 넓게 ALT+SHIFT+K
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeWidthIncrease() {
	p.hwp.Run("CharShapeWidthIncrease")
}

/*
ActionShape:CharShapeWidthDecrease
    장평 좁게 ALT+SHIFT+J
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeWidthDecrease() {
	p.hwp.Run("CharShapeWidthDecrease")
}

/*
ActionShape:CharShapeSpacingIncrease
    자간 넓게 ALT+SHIFT+W
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeSpacingIncrease() {
	p.hwp.Run("CharShapeSpacingIncrease")
}

/*
ActionShape:CharShapeSpacingDecrease
    자간 좁게 ALT+SHIFT+N
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeSpacingDecrease() {
	p.hwp.Run("CharShapeSpacingDecrease")
}

/*
ActionShape:CharShapeNextFaceName
    다음 글꼴 ALT+SHIFT+F
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeNextFaceName() {
	p.hwp.Run("CharShapeNextFaceName")
}

/*
ActionShape:CharShapePrevFaceName
    이전 글꼴 ALT+SHIFT+G
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapePrevFaceName() {
	p.hwp.Run("CharShapePrevFaceName")
}

/*
ActionShape:MarkPenShape
    형광펜 모양
    파라미터셋: MarkpenShape
*/
func (p *ActionShape) MarkPenShape(arg ...interface{}) {
	p.hwp.ActionWithParameters("MarkPenShape", arg...)
}

/*
ActionShape:ParaShapeDialog
    문단 모양 대화상자(내부 구현용)
    파라미터셋: ParaShape
*/
func (p *ActionShape) ParaShapeDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("ParaShapeDialog", arg...)
}

/*
ActionShape:ParagraphShape
    문단 모양
    파라미터셋: ParaShape
*/
func (p *ActionShape) ParagraphShape(arg ...interface{}) {
	p.hwp.ActionWithParameters("ParagraphShape", arg...)
}

/*
ActionShape:ParagraphShapeAlignJustify
    양쪽 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignJustify() {
	p.hwp.Run("ParagraphShapeAlignJustify")
}

/*
ActionShape:ParagraphShapeAlignLeft
    왼쪽 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignLeft() {
	p.hwp.Run("ParagraphShapeAlignLeft")
}

/*
ActionShape:ParagraphShapeAlignRight
    오른쪽 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignRight() {
	p.hwp.Run("ParagraphShapeAlignRight")
}

/*
ActionShape:ParagraphShapeAlignCenter
    가운데 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignCenter() {
	p.hwp.Run("ParagraphShapeAlignCenter")
}

/*
ActionShape:ParagraphShapeAlignDistribute
    배분 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignDistribute() {
	p.hwp.Run("ParagraphShapeAlignDistribute")
}

/*
ActionShape:ParagraphShapeAlignDivision
    나눔 정렬
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeAlignDivision() {
	p.hwp.Run("ParagraphShapeAlignDivision")
}

/*
ActionShape:ParagraphShapeWithNext
    다음 문단과 함께
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeWithNext() {
	p.hwp.Run("ParagraphShapeWithNext")
}

/*
ActionShape:ParagraphShapeProtect
    문단 보호
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeProtect() {
	p.hwp.Run("ParagraphShapeProtect")
}

/*
ActionShape:ParagraphShapeIncreaseLeftMargin
    왼쪽 여백 키우기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIncreaseLeftMargin() {
	p.hwp.Run("ParagraphShapeIncreaseLeftMargin")
}

/*
ActionShape:ParagraphShapeDecreaseLeftMargin
    왼쪽 여백 줄이기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeDecreaseLeftMargin() {
	p.hwp.Run("ParagraphShapeDecreaseLeftMargin")
}

/*
ActionShape:ParagraphShapeDecreaseRightMargin
    오른쪽 여백 키우기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeDecreaseRightMargin() {
	p.hwp.Run("ParagraphShapeDecreaseRightMargin")
}

/*
ActionShape:ParagraphShapeIncreaseRightMargin
    오른쪽 여백 줄이기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIncreaseRightMargin() {
	p.hwp.Run("ParagraphShapeIncreaseRightMargin")
}

/*
ActionShape:ParagraphShapeIndentAtCaret
    첫 줄 내어 쓰기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIndentAtCaret() {
	p.hwp.Run("ParagraphShapeIndentAtCaret")
}

/*
ActionShape:ParagraphShapeIndentNegative
    첫 줄을 한 글자 내어 씀
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIndentNegative() {
	p.hwp.Run("ParagraphShapeIndentNegative")
}

/*
ActionShape:ParagraphShapeIndentPositive
    첫 줄을 한 글자 들여 씀
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIndentPositive() {
	p.hwp.Run("ParagraphShapeIndentPositive")
}

/*
ActionShape:ParagraphShapeDecreaseLineSpacing
    줄 간격을 점점 좁힘
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeDecreaseLineSpacing() {
	p.hwp.Run("ParagraphShapeDecreaseLineSpacing")
}

/*
ActionShape:ParagraphShapeIncreaseLineSpacing
    줄 간격을 점점 넓힘
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIncreaseLineSpacing() {
	p.hwp.Run("ParagraphShapeIncreaseLineSpacing")
}

/*
ActionShape:ParagraphShapeIncreaseMargin
    왼쪽-오른쪽 여백 키우기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeIncreaseMargin() {
	p.hwp.Run("ParagraphShapeIncreaseMargin")
}

/*
ActionShape:ParagraphShapeDecreaseMargin
    왼쪽-오른쪽 여백 줄이기
    파라미터셋: 없음
*/
func (p *ActionShape) ParagraphShapeDecreaseMargin() {
	p.hwp.Run("ParagraphShapeDecreaseMargin")
}

/*
ActionShape:ParaNumberBullet
    문단번호/글머리표
    파라미터셋: ParaShape
*/
func (p *ActionShape) ParaNumberBullet(arg ...interface{}) {
	p.hwp.ActionWithParameters("ParaNumberBullet", arg...)
}

/*
ActionShape:OutlineNumber
    개요 번호
    파라미터셋: SecDef
*/
func (p *ActionShape) OutlineNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("OutlineNumber", arg...)
}

/*
ActionShape:ParaNumberBulletLevelUp
    문단번호/글머리표 한 수준 위로
    파라미터셋: ParaShape
*/
func (p *ActionShape) ParaNumberBulletLevelUp(arg ...interface{}) {
	p.hwp.ActionWithParameters("ParaNumberBulletLevelUp", arg...)
}

/*
ActionShape:ParaNumberBulletLevelDown
    문단번호/글머리표 한 수준 아래로
    파라미터셋: ParaShape
*/
func (p *ActionShape) ParaNumberBulletLevelDown(arg ...interface{}) {
	p.hwp.ActionWithParameters("ParaNumberBulletLevelDown", arg...)
}

/*
ActionShape:PutParaNumber
    문단번호 달기
    파라미터셋: ParaShape
*/
func (p *ActionShape) PutParaNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("PutParaNumber", arg...)
}

/*
ActionShape:PutNewParaNumber
    문단번호 새 번호 시작하기
    파라미터셋: ParaShape
*/
func (p *ActionShape) PutNewParaNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("PutNewParaNumber", arg...)
}

/*
ActionShape:PutBullet
    글머리표 달기
    파라미터셋: ParaShape
*/
func (p *ActionShape) PutBullet(arg ...interface{}) {
	p.hwp.ActionWithParameters("PutBullet", arg...)
}

/*
ActionShape:PutOutlineNumber
    개요 번호 닫기
    파라미터셋: ParaShape
*/
func (p *ActionShape) PutOutlineNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("PutOutlineNumber", arg...)
}

/*
ActionShape:GetDefaultParaNumber
    문단 번호 디폴트 값
    파라미터셋: ParaShape
*/
func (p *ActionShape) GetDefaultParaNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetDefaultParaNumber", arg...)
}

/*
ActionShape:GetDefaultBullet
    글머리표 디폴트 값
    파라미터셋: ParaShape
*/
func (p *ActionShape) GetDefaultBullet(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetDefaultBullet", arg...)
}

/*
ActionShape:StyleParaNumberBullet
    문단번호/글머리표
    파라미터셋: ParaShape
*/
func (p *ActionShape) StyleParaNumberBullet(arg ...interface{}) {
	p.hwp.ActionWithParameters("StyleParaNumberBullet", arg...)
}

/*
ActionShape:PageBorder
    쪽 테두리/배경
    파라미터셋: SecDef
*/
func (p *ActionShape) PageBorder(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageBorder", arg...)
}

/*
ActionShape:DropCap
    문단 첫 글자 장식
    파라미터셋: DropCap
*/
func (p *ActionShape) DropCap(arg ...interface{}) {
	p.hwp.ActionWithParameters("DropCap", arg...)
}

/*
ActionShape:ShapeCopyPaste
    모양 복사
    파라미터셋: ShapeCopyPaste
*/
func (p *ActionShape) ShapeCopyPaste(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeCopyPaste", arg...)
}

/*
ActionShape:Style
    스타일
    파라미터셋: Style
*/
func (p *ActionShape) Style(arg ...interface{}) {
	p.hwp.ActionWithParameters("Style", arg...)
}

/*
ActionShape:StyleShortcut1
    스타일 단축키(<Ctrl + 1>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut1() {
	p.hwp.Run("StyleShortcut1")
}

/*
ActionShape:StyleShortcut2
    스타일 단축키(<Ctrl + 2>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut2() {
	p.hwp.Run("StyleShortcut2")
}

/*
ActionShape:StyleShortcut3
    스타일 단축키(<Ctrl + 3>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut3() {
	p.hwp.Run("StyleShortcut3")
}

/*
ActionShape:StyleShortcut4
    스타일 단축키(<Ctrl + 4>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut4() {
	p.hwp.Run("StyleShortcut4")
}

/*
ActionShape:StyleShortcut5
    스타일 단축키(<Ctrl + 5>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut5() {
	p.hwp.Run("StyleShortcut5")
}

/*
ActionShape:StyleShortcut6
    스타일 단축키(<Ctrl + 6>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut6() {
	p.hwp.Run("StyleShortcut6")
}

/*
ActionShape:StyleShortcut7
    스타일 단축키(<Ctrl + 7>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut7() {
	p.hwp.Run("StyleShortcut7")
}

/*
ActionShape:StyleShortcut8
    스타일 단축키(<Ctrl + 8>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut8() {
	p.hwp.Run("StyleShortcut8")
}

/*
ActionShape:StyleShortcut9
    스타일 단축키(<Ctrl + 9>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut9() {
	p.hwp.Run("StyleShortcut9")
}

/*
ActionShape:StyleShortcut10
    스타일 단축키(<Ctrl + 0>
    파라미터셋: 없음
*/
func (p *ActionShape) StyleShortcut10() {
	p.hwp.Run("StyleShortcut10")
}

/*
ActionShape:StyleClearCharStyle
    글자 스타일 해제
    파라미터셋: 없음
*/
func (p *ActionShape) StyleClearCharStyle() {
	p.hwp.Run("StyleClearCharStyle")
}

/*
ActionShape:StyleTemplate
    스타일 마당
    파라미터셋: StyleTemplate
*/
func (p *ActionShape) StyleTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("StyleTemplate", arg...)
}

/*
ActionShape:PageSetup
    편집 용지
    파라미터셋: SecDef
*/
func (p *ActionShape) PageSetup(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageSetup", arg...)
}

/*
ActionShape:ModifySection
    구역
    파라미터셋: SecDef
*/
func (p *ActionShape) ModifySection(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifySection", arg...)
}

/*
ActionShape:MasterPage
    바탕쪽 Entry
    파라미터셋: MasterPage
*/
func (p *ActionShape) MasterPage(arg ...interface{}) {
	p.hwp.ActionWithParameters("MasterPage", arg...)
}

/*
ActionShape:MasterPageEntry
    바탕쪽
    파라미터셋: MasterPage
*/
func (p *ActionShape) MasterPageEntry(arg ...interface{}) {
	p.hwp.ActionWithParameters("MasterPageEntry", arg...)
}

/*
ActionShape:MasterPageTypeDlg
    바탕쪽 종류 다이얼로그 띄움
    파라미터셋: MasterPage
*/
func (p *ActionShape) MasterPageTypeDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("MasterPageTypeDlg", arg...)
}

/*
ActionShape:MasterPagePrevSection
    앞 구역 바탕쪽 사용
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPagePrevSection() {
	p.hwp.Run("MasterPagePrevSection")
}

/*
ActionShape:MasterPageExcept
    첫 쪽 제외
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageExcept() {
	p.hwp.Run("MasterPageExcept")
}

/*
ActionShape:MPSectionToPrevious
    이전 구역으로
    파라미터셋: 없음
*/
func (p *ActionShape) MPSectionToPrevious() {
	p.hwp.Run("MPSectionToPrevious")
}

/*
ActionShape:MPSectionToNext
    이후 구역으로
    파라미터셋: 없음
*/
func (p *ActionShape) MPSectionToNext() {
	p.hwp.Run("MPSectionToNext")
}

/*
ActionShape:MasterPageToPrevious
    이전 바탕쪽
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageToPrevious() {
	p.hwp.Run("MasterPageToPrevious")
}

/*
ActionShape:MasterPageToNext
    이후 바탕쪽
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageToNext() {
	p.hwp.Run("MasterPageToNext")
}

/*
ActionShape:MasterPageDelete
    바탕쪽 삭제
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageDelete() {
	p.hwp.Run("MasterPageDelete")
}

/*
ActionShape:MasterPageDuplicate
    기존 바탕쪽과 겹침
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageDuplicate() {
	p.hwp.Run("MasterPageDuplicate")
}

/*
ActionShape:MasterPageFront
    바탕쪽 앞으로 보내기
    파라미터셋: 없음
*/
func (p *ActionShape) MasterPageFront() {
	p.hwp.Run("MasterPageFront")
}

/*
ActionShape:HeaderFooter
    머리말/꼬리말
    파라미터셋: HeaderFooter
*/
func (p *ActionShape) HeaderFooter(arg ...interface{}) {
	p.hwp.ActionWithParameters("HeaderFooter", arg...)
}

/*
ActionShape:HeaderFooterInsField
    코드 넣기
    파라미터셋: HeaderFooter
*/
func (p *ActionShape) HeaderFooterInsField(arg ...interface{}) {
	p.hwp.ActionWithParameters("HeaderFooterInsField", arg...)
}

/*
ActionShape:HeaderFooterToPrev
    이전 머리말
    파라미터셋: 없음
*/
func (p *ActionShape) HeaderFooterToPrev() {
	p.hwp.Run("HeaderFooterToPrev")
}

/*
ActionShape:HeaderFooterToNext
    이후 머리말
    파라미터셋: HeaderFooter
*/
func (p *ActionShape) HeaderFooterToNext(arg ...interface{}) {
	p.hwp.ActionWithParameters("HeaderFooterToNext", arg...)
}

/*
ActionShape:HeaderFooterDelete
    머리말 지우기
    파라미터셋: 없음
*/
func (p *ActionShape) HeaderFooterDelete() {
	p.hwp.Run("HeaderFooterDelete")
}

/*
ActionShape:HeaderFooterModify
    머리말/꼬리말 고치기
    파라미터셋: 없음
*/
func (p *ActionShape) HeaderFooterModify() {
	p.hwp.Run("HeaderFooterModify")
}

/*
ActionShape:MultiColumn
    다단
    파라미터셋: ColDef
*/
func (p *ActionShape) MultiColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("MultiColumn", arg...)
}

/*
ActionShape:PageHiding
    감추기
    파라미터셋: PageHiding
*/
func (p *ActionShape) PageHiding(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageHiding", arg...)
}

/*
ActionShape:PageHidingModify
    감추기 고치기
    파라미터셋: PageHiding
*/
func (p *ActionShape) PageHidingModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageHidingModify", arg...)
}

/*
ActionShape:BreakPage
    쪽 나누기
    파라미터셋: 없음
*/
func (p *ActionShape) BreakPage() {
	p.hwp.Run("BreakPage")
}

/*
ActionShape:BreakColumn
    단 나누기
    파라미터셋: 없음
*/
func (p *ActionShape) BreakColumn() {
	p.hwp.Run("BreakColumn")
}

/*
ActionShape:BreakSection
    구역 나누기
    파라미터셋: 없음
*/
func (p *ActionShape) BreakSection() {
	p.hwp.Run("BreakSection")
}

/*
ActionShape:BreakLine
    줄 나눔
    파라미터셋: 없음
*/
func (p *ActionShape) BreakLine() {
	p.hwp.Run("BreakLine")
}

/*
ActionShape:BreakPara
    문단 나누기
    파라미터셋: 없음
*/
func (p *ActionShape) BreakPara() {
	p.hwp.Run("BreakPara")
}

/*
ActionShape:BreakColDef
    단 정의 삽입
    파라미터셋: 없음
*/
func (p *ActionShape) BreakColDef() {
	p.hwp.Run("BreakColDef")
}

/*
ActionShape:InsertPageNum
    쪽 번호 넣기
    파라미터셋: 없음
*/
func (p *ActionShape) InsertPageNum() {
	p.hwp.Run("InsertPageNum")
}

/*
ActionShape:PageNumPos
    쪽 번호 매기기
    파라미터셋: PageNumPos
*/
func (p *ActionShape) PageNumPos(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageNumPos", arg...)
}

/*
ActionShape:PageNumPosModify
    쪽 번호 매기기
    파라미터셋: PageNumPos
*/
func (p *ActionShape) PageNumPosModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("PageNumPosModify", arg...)
}

/*
ActionShape:NewNumber
    새 번호로 시작
    파라미터셋: AutoNum
*/
func (p *ActionShape) NewNumber(arg ...interface{}) {
	p.hwp.ActionWithParameters("NewNumber", arg...)
}

/*
ActionShape:NewNumberModify
    새 번호 고치기
    파라미터셋: AutoNum
*/
func (p *ActionShape) NewNumberModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("NewNumberModify", arg...)
}

/*
ActionShape:CharShapeTextColorBlack
    글자색을 검정
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorBlack() {
	p.hwp.Run("CharShapeTextColorBlack")
}

/*
ActionShape:CharShapeTextColorRed
    글자색을 빨강
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorRed() {
	p.hwp.Run("CharShapeTextColorRed")
}

/*
ActionShape:CharShapeTextColorBlue
    글자색을 파랑
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorBlue() {
	p.hwp.Run("CharShapeTextColorBlue")
}

/*
ActionShape:CharShapeTextColorViolet
    글자색을 자주
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorViolet() {
	p.hwp.Run("CharShapeTextColorViolet")
}

/*
ActionShape:CharShapeTextColorGreen
    글자색을 초록
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorGreen() {
	p.hwp.Run("CharShapeTextColorGreen")
}

/*
ActionShape:CharShapeTextColorYellow
    글자색을 노랑
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorYellow() {
	p.hwp.Run("CharShapeTextColorYellow")
}

/*
ActionShape:CharShapeTextColorBluish
    글자색을 청록
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorBluish() {
	p.hwp.Run("CharShapeTextColorBluish")
}

/*
ActionShape:CharShapeTextColorWhite
    글자색을 흰색
    파라미터셋: 없음
*/
func (p *ActionShape) CharShapeTextColorWhite() {
	p.hwp.Run("CharShapeTextColorWhite")
}

/*
ActionView:ViewOptionPaper
    쪽 윤곽 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionPaper() {
	p.hwp.Run("ViewOptionPaper")
}

/*
ActionView:ViewOptionGuideLine
    안내선 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionGuideLine() {
	p.hwp.Run("ViewOptionGuideLine")
}

/*
ActionView:ViewOptionPicture
    그림 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionPicture() {
	p.hwp.Run("ViewOptionPicture")
}

/*
ActionView:ViewOptionMemo
    메모 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionMemo() {
	p.hwp.Run("ViewOptionMemo")
}

/*
ActionView:ViewOptionMemoGuideline
    메모 안내선 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionMemoGuideline() {
	p.hwp.Run("ViewOptionMemoGuideline")
}

/*
ActionView:ViewOptionParaMark
    문단 부호 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionParaMark() {
	p.hwp.Run("ViewOptionParaMark")
}

/*
ActionView:ViewOptionCtrlMark
    조판 부호 보기
    파라미터셋: 없음
*/
func (p *ActionView) ViewOptionCtrlMark() {
	p.hwp.Run("ViewOptionCtrlMark")
}

/*
ActionView:ViewZoom
    화면 확대(Ribbon)
    파라미터셋: ViewProperties
*/
func (p *ActionView) ViewZoom(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewZoom", arg...)
}

/*
ActionView:ViewZoomNormal
    화면 확대: 정상
    파라미터셋: ViewProperties
*/
func (p *ActionView) ViewZoomNormal(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewZoomNormal", arg...)
}

/*
ActionView:ViewZoomFitPage
    화면 확대: 페이지에 맞춤
    파라미터셋: ViewProperties
*/
func (p *ActionView) ViewZoomFitPage(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewZoomFitPage", arg...)
}

/*
ActionView:ViewZoomFitWidth
    화면 확대: 폭에 맞춤
    파라미터셋: ViewProperties
*/
func (p *ActionView) ViewZoomFitWidth(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewZoomFitWidth", arg...)
}

/*
ActionView:ViewGridOption
    격자
    파라미터셋: GridInfo
*/
func (p *ActionView) ViewGridOption(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewGridOption", arg...)
}

/*
ActionView:ViewShowGrid
    격자 보이기
    파라미터셋: GridInfo
*/
func (p *ActionView) ViewShowGrid(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewShowGrid", arg...)
}

/*
ActionInput:Idiom
    상용구
    파라미터셋: Idiom
*/
func (p *ActionInput) Idiom(arg ...interface{}) {
	p.hwp.ActionWithParameters("Idiom", arg...)
}

/*
ActionInput:ViewIdiom
    상용구 보기
    파라미터셋: Idiom
*/
func (p *ActionInput) ViewIdiom(arg ...interface{}) {
	p.hwp.ActionWithParameters("ViewIdiom", arg...)
}

/*
ActionInput:InsertIdiom
    상용구 등록
    파라미터셋: Idiom
*/
func (p *ActionInput) InsertIdiom(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertIdiom", arg...)
}

/*
ActionInput:InputCodeChange
    문자 코드 변환
    파라미터셋: 없음
*/
func (p *ActionInput) InputCodeChange() {
	p.hwp.Run("InputCodeChange")
}

/*
ActionInput:InputDateStyle
    날짜/시간 기본 형식 지정
    파라미터셋: InputDateStyle
*/
func (p *ActionInput) InputDateStyle(arg ...interface{}) {
	p.hwp.ActionWithParameters("InputDateStyle", arg...)
}

/*
ActionInput:InsertDateCode
    상용구 코드 넣기 : 만든 날짜
    파라미터셋: 없음
*/
func (p *ActionInput) InsertDateCode() {
	p.hwp.Run("InsertDateCode")
}

/*
ActionInput:InsertUserName
    상용구 코드 넣기 : 만든 사람
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) InsertUserName(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertUserName", arg...)
}

/*
ActionInput:InsertFileName
    상용구 코드 넣기 : 파일 이름
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) InsertFileName(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertFileName", arg...)
}

/*
ActionInput:InsertFilePath
    상용구 코드 넣기 : 파일 경로
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) InsertFilePath(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertFilePath", arg...)
}

/*
ActionInput:InsertDocTitle
    상용구 코드 넣기 : 문서 제목
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) InsertDocTitle(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertDocTitle", arg...)
}

/*
ActionInput:InsertCpNo
    상용구 코드 넣기 : 현재 쪽 번호
    파라미터셋: 없음
*/
func (p *ActionInput) InsertCpNo() {
	p.hwp.Run("InsertCpNo")
}

/*
ActionInput:InsertTpNo
    상용구 코드 넣기 : 현재 쪽수
    파라미터셋: 없음
*/
func (p *ActionInput) InsertTpNo() {
	p.hwp.Run("InsertTpNo")
}

/*
ActionInput:InsertCpTpNo
    상용구 코드 넣기 : 현재 쪽/전체 쪽수
    파라미터셋: 없음
*/
func (p *ActionInput) InsertCpTpNo() {
	p.hwp.Run("InsertCpTpNo")
}

/*
ActionInput:InsertDocInfo
    상용구 코드 넣기 : 만든 사람, 현재 쪽 번호, 만든 날짜
    파라미터셋: 없음
*/
func (p *ActionInput) InsertDocInfo() {
	p.hwp.Run("InsertDocInfo")
}

/*
ActionInput:InsertLastSaveDate
    상용구 코드 넣기 : 마지막 저장한 날짜
    파라미터셋: 없음
*/
func (p *ActionInput) InsertLastSaveDate() {
	p.hwp.Run("InsertLastSaveDate")
}

/*
ActionInput:InsertLastSaveBy
    상용구 코드 넣기 : 마지막 저장한 사람
    파라미터셋: 없음
*/
func (p *ActionInput) InsertLastSaveBy() {
	p.hwp.Run("InsertLastSaveBy")
}

/*
ActionInput:InsertLastPrintDate
    상용구 코드 넣기 : 마지막 인쇄한 날짜
    파라미터셋: 없음
*/
func (p *ActionInput) InsertLastPrintDate() {
	p.hwp.Run("InsertLastPrintDate")
}

/*
ActionInput:ComposeChars
    글자 겹침
    파라미터셋: ChCompose
*/
func (p *ActionInput) ComposeChars(arg ...interface{}) {
	p.hwp.ActionWithParameters("ComposeChars", arg...)
}

/*
ActionInput:DutmalChars
    덧말 넣기
    파라미터셋: Dutmal
*/
func (p *ActionInput) DutmalChars(arg ...interface{}) {
	p.hwp.ActionWithParameters("DutmalChars", arg...)
}

/*
ActionInput:AddHanjaWord
    한자 단어 등록
    파라미터셋: AddHanjaWord
*/
func (p *ActionInput) AddHanjaWord(arg ...interface{}) {
	p.hwp.ActionWithParameters("AddHanjaWord", arg...)
}

/*
ActionInput:EquationCreate
    수식 만들기
    파라미터셋: EqEdit
*/
func (p *ActionInput) EquationCreate(arg ...interface{}) {
	p.hwp.ActionWithParameters("EquationCreate", arg...)
}

/*
ActionInput:EquationPropertyDialog
    수식 고치기 (대화상자)
    파라미터셋: ShapeObject
*/
func (p *ActionInput) EquationPropertyDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("EquationPropertyDialog", arg...)
}

/*
ActionInput:EquationModify
    수식 고치기
    파라미터셋: EqEdit
*/
func (p *ActionInput) EquationModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("EquationModify", arg...)
}

/*
ActionInput:HimKbdChange
    바꾸기
    파라미터셋: 없음
*/
func (p *ActionInput) HimKbdChange() {
	p.hwp.Run("HimKbdChange")
}

/*
ActionInput:Soft
    없음 보기
    파라미터셋: Keyboard
*/
func (p *ActionInput) Soft(arg ...interface{}) {
	p.hwp.ActionWithParameters("Soft", arg...)
}

/*
ActionInput:RunUserKeyLayout
    사용자 글자판 제작 툴
    파라미터셋: 없음
*/
func (p *ActionInput) RunUserKeyLayout() {
	p.hwp.Run("RunUserKeyLayout")
}

/*
ActionInput:Him
    없음 입력기 언어별 환경 설정
    파라미터셋: Config
*/
func (p *ActionInput) Him(arg ...interface{}) {
	p.hwp.ActionWithParameters("Him", arg...)
}

/*
ActionInput:InsertFootnote
    각주 입력
    파라미터셋: 없음
*/
func (p *ActionInput) InsertFootnote() {
	p.hwp.Run("InsertFootnote")
}

/*
ActionInput:InsertEndnote
    미주 입력
    파라미터셋: 없음
*/
func (p *ActionInput) InsertEndnote() {
	p.hwp.Run("InsertEndnote")
}

/*
ActionInput:FootnoteOption
    각주/미주 모양
    파라미터셋: SecDef
*/
func (p *ActionInput) FootnoteOption(arg ...interface{}) {
	p.hwp.ActionWithParameters("FootnoteOption", arg...)
}

/*
ActionInput:NoteNumProperty
    주석 번호 속성
    파라미터셋: 없음
*/
func (p *ActionInput) NoteNumProperty() {
	p.hwp.Run("NoteNumProperty")
}

/*
ActionInput:NoteToPrev
    주석 앞으로 이동
    파라미터셋: 없음
*/
func (p *ActionInput) NoteToPrev() {
	p.hwp.Run("NoteToPrev")
}

/*
ActionInput:NoteToNext
    주석 다음으로 이동
    파라미터셋: 없음
*/
func (p *ActionInput) NoteToNext() {
	p.hwp.Run("NoteToNext")
}

/*
ActionInput:NoteDelete
    주석 지우기
    파라미터셋: 없음
*/
func (p *ActionInput) NoteDelete() {
	p.hwp.Run("NoteDelete")
}

/*
ActionInput:NoteModify
    주석 고치기
    파라미터셋: 없음
*/
func (p *ActionInput) NoteModify() {
	p.hwp.Run("NoteModify")
}

/*
ActionInput:ExchangeFootnoteEndnote
    각주/미주 변환
    파라미터셋: ExchangeFootnoteEndNote
*/
func (p *ActionInput) ExchangeFootnoteEndnote(arg ...interface{}) {
	p.hwp.ActionWithParameters("ExchangeFootnoteEndnote", arg...)
}

/*
ActionInput:SaveFootnote
    주석 저장
    파라미터셋: SaveFootnote
*/
func (p *ActionInput) SaveFootnote(arg ...interface{}) {
	p.hwp.ActionWithParameters("SaveFootnote", arg...)
}

/*
ActionInput:InsertAutoNum
    번호 다시 넣기
    파라미터셋: 없음
*/
func (p *ActionInput) InsertAutoNum() {
	p.hwp.Run("InsertAutoNum")
}

/*
ActionInput:Comment
    숨은 설명
    파라미터셋: 없음
*/
func (p *ActionInput) Comment() {
	p.hwp.Run("Comment")
}

/*
ActionInput:CommentModify
    숨은 설명 고치기
    파라미터셋: 없음
*/
func (p *ActionInput) CommentModify() {
	p.hwp.Run("CommentModify")
}

/*
ActionInput:CommentDelete
    숨은 설명 지우기
    파라미터셋: 없음
*/
func (p *ActionInput) CommentDelete() {
	p.hwp.Run("CommentDelete")
}

/*
ActionInput:Bookmark
    책갈피
    파라미터셋: BookMark
*/
func (p *ActionInput) Bookmark(arg ...interface{}) {
	p.hwp.ActionWithParameters("Bookmark", arg...)
}

/*
ActionInput:ModifyBookmark
    책갈피 고치기
    파라미터셋: BookMark
*/
func (p *ActionInput) ModifyBookmark(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyBookmark", arg...)
}

/*
ActionInput:Hyperlink
    캐럿이 필드 안에 위치했는지 여부에 따라 Insert 또는 Modify 하이퍼 링크 액션
    파라미터셋: HyperLink
*/
func (p *ActionInput) Hyperlink(arg ...interface{}) {
	p.hwp.ActionWithParameters("Hyperlink", arg...)
}

/*
ActionInput:InsertHyperlink
    하이퍼링크 만들기
    파라미터셋: HyperLink
*/
func (p *ActionInput) InsertHyperlink(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertHyperlink", arg...)
}

/*
ActionInput:ModifyHyperlink
    하이퍼링크 고치기
    파라미터셋: HyperLink
*/
func (p *ActionInput) ModifyHyperlink(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyHyperlink", arg...)
}

/*
ActionInput:HyperlinkJump
    하이퍼링크 이동
    파라미터셋: HyperlinkJump
*/
func (p *ActionInput) HyperlinkJump(arg ...interface{}) {
	p.hwp.ActionWithParameters("HyperlinkJump", arg...)
}

/*
ActionInput:HyperlinkBackward
    하이퍼링크 뒤로
    파라미터셋: 없음
*/
func (p *ActionInput) HyperlinkBackward() {
	p.hwp.Run("HyperlinkBackward")
}

/*
ActionInput:HyperlinkForward
    하이퍼링크 앞으로
    파라미터셋: 없음
*/
func (p *ActionInput) HyperlinkForward() {
	p.hwp.Run("HyperlinkForward")
}

/*
ActionInput:ReturnKeyInField
    캐럿이 필드 안에 위치한 상태에서 Return Key에 대한 액션 분기
    파라미터셋: 없음
*/
func (p *ActionInput) ReturnKeyInField() {
	p.hwp.Run("ReturnKeyInField")
}

/*
ActionInput:QuickMarkInsert0
    쉬운 책갈피 - 삽입 0
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert0() {
	p.hwp.Run("QuickMarkInsert0")
}

/*
ActionInput:QuickMarkInsert1
    쉬운 책갈피 - 삽입 1
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert1() {
	p.hwp.Run("QuickMarkInsert1")
}

/*
ActionInput:QuickMarkInsert2
    쉬운 책갈피 - 삽입 2
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert2() {
	p.hwp.Run("QuickMarkInsert2")
}

/*
ActionInput:QuickMarkInsert3
    쉬운 책갈피 - 삽입 3
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert3() {
	p.hwp.Run("QuickMarkInsert3")
}

/*
ActionInput:QuickMarkInsert4
    쉬운 책갈피 - 삽입 4
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert4() {
	p.hwp.Run("QuickMarkInsert4")
}

/*
ActionInput:QuickMarkInsert5
    쉬운 책갈피 - 삽입 5
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert5() {
	p.hwp.Run("QuickMarkInsert5")
}

/*
ActionInput:QuickMarkInsert6
    쉬운 책갈피 - 삽입 6
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert6() {
	p.hwp.Run("QuickMarkInsert6")
}

/*
ActionInput:QuickMarkInsert7
    쉬운 책갈피 - 삽입 7
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert7() {
	p.hwp.Run("QuickMarkInsert7")
}

/*
ActionInput:QuickMarkInsert8
    쉬운 책갈피 - 삽입 8
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert8() {
	p.hwp.Run("QuickMarkInsert8")
}

/*
ActionInput:QuickMarkInsert9
    쉬운 책갈피 - 삽입 9
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkInsert9() {
	p.hwp.Run("QuickMarkInsert9")
}

/*
ActionInput:QuickMarkMove0
    쉬운 책갈피 - 이동 0
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove0() {
	p.hwp.Run("QuickMarkMove0")
}

/*
ActionInput:QuickMarkMove1
    쉬운 책갈피 - 이동 1
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove1() {
	p.hwp.Run("QuickMarkMove1")
}

/*
ActionInput:QuickMarkMove2
    쉬운 책갈피 - 이동 2
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove2() {
	p.hwp.Run("QuickMarkMove2")
}

/*
ActionInput:QuickMarkMove3
    쉬운 책갈피 - 이동 3
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove3() {
	p.hwp.Run("QuickMarkMove3")
}

/*
ActionInput:QuickMarkMove4
    쉬운 책갈피 - 이동 4
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove4() {
	p.hwp.Run("QuickMarkMove4")
}

/*
ActionInput:QuickMarkMove5
    쉬운 책갈피 - 이동 5
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove5() {
	p.hwp.Run("QuickMarkMove5")
}

/*
ActionInput:QuickMarkMove6
    쉬운 책갈피 - 이동 6
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove6() {
	p.hwp.Run("QuickMarkMove6")
}

/*
ActionInput:QuickMarkMove7
    쉬운 책갈피 - 이동 7
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove7() {
	p.hwp.Run("QuickMarkMove7")
}

/*
ActionInput:QuickMarkMove8
    쉬운 책갈피 - 이동 8
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove8() {
	p.hwp.Run("QuickMarkMove8")
}

/*
ActionInput:QuickMarkMove9
    쉬운 책갈피 - 이동 9
    파라미터셋: 없음
*/
func (p *ActionInput) QuickMarkMove9() {
	p.hwp.Run("QuickMarkMove9")
}

/*
ActionInput:InsertCrossReference
    상호 참조 넣기
    파라미터셋: ActionCrossRef
*/
func (p *ActionInput) InsertCrossReference(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertCrossReference", arg...)
}

/*
ActionInput:ModifyCrossReference
    상호 참조 고치기
    파라미터셋: ActionCrossRef
*/
func (p *ActionInput) ModifyCrossReference(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyCrossReference", arg...)
}

/*
ActionInput:InsertFieldTemplate
    문서마당 정보
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) InsertFieldTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertFieldTemplate", arg...)
}

/*
ActionInput:ModifyFieldClickhere
    누름틀 정보 고치기
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) ModifyFieldClickhere(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyFieldClickhere", arg...)
}

/*
ActionInput:ModifyFieldUserInfo
    개인 정보 필드 고치기
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) ModifyFieldUserInfo(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyFieldUserInfo", arg...)
}

/*
ActionInput:ModifyFieldSummary
    문서 요약 필드 고치기
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) ModifyFieldSummary(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyFieldSummary", arg...)
}

/*
ActionInput:ModifyFieldDate
    날짜 필드 고치기
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) ModifyFieldDate(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyFieldDate", arg...)
}

/*
ActionInput:ModifyFieldPath
    문서 경로 필드 고치기
    파라미터셋: InsertFieldTemplate
*/
func (p *ActionInput) ModifyFieldPath(arg ...interface{}) {
	p.hwp.ActionWithParameters("ModifyFieldPath", arg...)
}

/*
ActionInput:InsertFile
    끼워 넣기
    파라미터셋: InsertFile
*/
func (p *ActionInput) InsertFile(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertFile", arg...)
}

/*
ActionInput:InsertFieldMemo
    메모 넣기
    파라미터셋: 없음
*/
func (p *ActionInput) InsertFieldMemo() {
	p.hwp.Run("InsertFieldMemo")
}

/*
ActionInput:DeleteFieldMemo
    메모 지우기
    파라미터셋: 없음
*/
func (p *ActionInput) DeleteFieldMemo() {
	p.hwp.Run("DeleteFieldMemo")
}

/*
ActionInput:MemoShape
    메모 모양
    파라미터셋: SecDef
*/
func (p *ActionInput) MemoShape(arg ...interface{}) {
	p.hwp.ActionWithParameters("MemoShape", arg...)
}

/*
ActionInput:InsertTab
    탭 삽입
    파라미터셋: 없음
*/
func (p *ActionInput) InsertTab() {
	p.hwp.Run("InsertTab")
}

/*
ActionInput:InsertSoftHyphen
    하이픈 삽입
    파라미터셋: 없음
*/
func (p *ActionInput) InsertSoftHyphen() {
	p.hwp.Run("InsertSoftHyphen")
}

/*
ActionInput:InsertText
    텍스트 삽입
    파라미터셋: InsertText
*/
func (p *ActionInput) InsertText(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertText", arg...)
}

/*
ActionInput:InsertSpace
    공백 삽입
    파라미터셋: InsertText
*/
func (p *ActionInput) InsertSpace(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertSpace", arg...)
}

/*
ActionInput:InsertNonBreakingSpace
    묶음 빈칸 삽입
    파라미터셋: 없음
*/
func (p *ActionInput) InsertNonBreakingSpace() {
	p.hwp.Run("InsertNonBreakingSpace")
}

/*
ActionInput:InsertFixedWidthSpace
    고정 폭 빈칸 삽입
    파라미터셋: 없음
*/
func (p *ActionInput) InsertFixedWidthSpace() {
	p.hwp.Run("InsertFixedWidthSpace")
}

/*
ActionInput:RightShiftBlock
    텍스트 블록 상태에서 블록이 문단의 시작위치에서 시작할 경우 블록 왼쪽에 탭을 삽입한다.
    파라미터셋: 없음
*/
func (p *ActionInput) RightShiftBlock() {
	p.hwp.Run("RightShiftBlock")
}

/*
ActionInput:LeftShiftBlock
    텍스트 블록 상태에서 블록 왼쪽에 있는 탭 또는 공백을 지운다.
    파라미터셋: 없음
*/
func (p *ActionInput) LeftShiftBlock() {
	p.hwp.Run("LeftShiftBlock")
}

/*
ActionInput:OleCreateNew
    개체 삽입
    파라미터셋: OleCreation
*/
func (p *ActionInput) OleCreateNew(arg ...interface{}) {
	p.hwp.ActionWithParameters("OleCreateNew", arg...)
}

/*
ActionInput:InsertVoice
    음성 삽입
    파라미터셋: OleCreation
*/
func (p *ActionInput) InsertVoice(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertVoice", arg...)
}

/*
ActionInput:InsertFlash
    플래쉬 삽입
    파라미터셋: OleCreation
*/
func (p *ActionInput) InsertFlash(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertFlash", arg...)
}

/*
ActionInput:InsertMoive
    동영상 삽입
    파라미터셋: OleCreation
*/
func (p *ActionInput) InsertMoive(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertMoive", arg...)
}

/*
ActionInput:PictureInsertDialog
    그림 넣기 (대화상자를 띄워 선택한 이미지 파일을 문서에 삽입하는 액션 : API용)
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureInsertDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureInsertDialog", arg...)
}

/*
ActionInput:ShapeObjDialog
    환경 설정
    파라미터셋: ShapeObject
*/
func (p *ActionInput) ShapeObjDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeObjDialog", arg...)
}

/*
ActionInput:DrawObjTemplateLoad
    그리기 마당에서 불러오기
    파라미터셋: ShapeObject
*/
func (p *ActionInput) DrawObjTemplateLoad(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjTemplateLoad", arg...)
}

/*
ActionInput:DrawObjCreatorTextBox
    글상자
    파라미터셋: ShapeObject
*/
func (p *ActionInput) DrawObjCreatorTextBox(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorTextBox", arg...)
}

/*
ActionInput:DrawObjCreatorConnectLine
    개체 연결선 만들기
    파라미터셋: ShapeObject
*/
func (p *ActionInput) DrawObjCreatorConnectLine(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorConnectLine", arg...)
}

/*
ActionInput:PictureEffect1
    그림 그레이 스케일
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect1(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect1", arg...)
}

/*
ActionInput:PictureEffect2
    그림 흑백으로
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect2(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect2", arg...)
}

/*
ActionInput:PictureEffect3
    그림 워터마크
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect3(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect3", arg...)
}

/*
ActionInput:PictureEffect4
    그림 효과 없음
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect4(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect4", arg...)
}

/*
ActionInput:PictureToOriginal
    그림 원래 그림으로
    파라미터셋: 없음
*/
func (p *ActionInput) PictureToOriginal() {
	p.hwp.Run("PictureToOriginal")
}

/*
ActionInput:PictureEffect5
    그림 밝기 증가
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect5(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect5", arg...)
}

/*
ActionInput:PictureEffect6
    그림 밝기 감소
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect6(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect6", arg...)
}

/*
ActionInput:PictureEffect7
    그림 명암 증가
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect7(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect7", arg...)
}

/*
ActionInput:PictureEffect8
    그림 명암 감소
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureEffect8(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureEffect8", arg...)
}

/*
ActionInput:PictureSave
    그림 빼내기
    파라미터셋: 없음
*/
func (p *ActionInput) PictureSave() {
	p.hwp.Run("PictureSave")
}

/*
ActionInput:PictureLinkedToEmbedded
    연결된 그림을 모두 삽입 그림으로
    파라미터셋: ShapeObject
*/
func (p *ActionInput) PictureLinkedToEmbedded(arg ...interface{}) {
	p.hwp.ActionWithParameters("PictureLinkedToEmbedded", arg...)
}

/*
ActionInput:PictureScissor
    그림 자르기
    파라미터셋: 없음
*/
func (p *ActionInput) PictureScissor() {
	p.hwp.Run("PictureScissor")
}

/*
ActionInput:InsertLine
    선 넣기
    파라미터셋: 없음
*/
func (p *ActionInput) InsertLine() {
	p.hwp.Run("InsertLine")
}

/*
ActionInput:ShapeObjAttachCaption
    캡션 넣기
    파라미터셋: 없음
*/
func (p *ActionInput) ShapeObjAttachCaption() {
	p.hwp.Run("ShapeObjAttachCaption")
}

/*
ActionInput:ShapeObjInsertCaptionNum
    캡션 번호 넣기
    파라미터셋: 없음
*/
func (p *ActionInput) ShapeObjInsertCaptionNum() {
	p.hwp.Run("ShapeObjInsertCaptionNum")
}

/*
ActionInput:ShapeObjDetachCaption
    캡션 없애기
    파라미터셋: 없음
*/
func (p *ActionInput) ShapeObjDetachCaption() {
	p.hwp.Run("ShapeObjDetachCaption")
}

/*
ActionTool:AutoChangeRun
    동작
    파라미터셋: 없음
*/
func (p *ActionTool) AutoChangeRun() {
	p.hwp.Run("AutoChangeRun")
}

/*
ActionTool:AutoChangeHangul
    낱자모 우선
    파라미터셋: 없음
*/
func (p *ActionTool) AutoChangeHangul() {
	p.hwp.Run("AutoChangeHangul")
}

/*
ActionTool:ManualChangeHangul
    한영 수동 전환
    파라미터셋: 없음
*/
func (p *ActionTool) ManualChangeHangul() {
	p.hwp.Run("ManualChangeHangul")
}

/*
ActionTool:QuickCorrect
    빠른 교정
    파라미터셋: QCorrect
*/
func (p *ActionTool) QuickCorrect(arg ...interface{}) {
	p.hwp.ActionWithParameters("QuickCorrect", arg...)
}

/*
ActionTool:QuickCorrect Edit
    빠른 교정 - 내용 편집
    파라미터셋: QCorrect
*/
func (p *ActionTool) QuickCorrectEdit(arg ...interface{}) {
	p.hwp.ActionWithParameters("QuickCorrect Edit", arg...)
}

/*
ActionTool:QuickCorrect Run
    빠른 교정 동작
    파라미터셋: QCorrect
*/
func (p *ActionTool) QuickCorrectRun(arg ...interface{}) {
	p.hwp.ActionWithParameters("QuickCorrect Run", arg...)
}

/*
ActionTool:QuickCorrect Sound
    빠른 교정 - 메뉴에서 효과음 On/Off
    파라미터셋: QCorrect
*/
func (p *ActionTool) QuickCorrectSound(arg ...interface{}) {
	p.hwp.ActionWithParameters("QuickCorrect Sound", arg...)
}

/*
ActionTool:ChangeRome
    로마자 변환
    파라미터셋: ChangeRome
*/
func (p *ActionTool) ChangeRome(arg ...interface{}) {
	p.hwp.ActionWithParameters("ChangeRome", arg...)
}

/*
ActionTool:ChangeRome String
    로마자변환 - 입력받은 스트링 변환
    파라미터셋: ChangeRome
*/
func (p *ActionTool) ChangeRomeString(arg ...interface{}) {
	p.hwp.ActionWithParameters("ChangeRome String", arg...)
}

/*
ActionTool:GetRome String
    로마자 변환 - 입력 받은 스트링 변환을 로마자로 변환해서 넘겨줌
    파라미터셋: ChangeRome
*/
func (p *ActionTool) GetRomeString(arg ...interface{}) {
	p.hwp.ActionWithParameters("GetRome String", arg...)
}

/*
ActionTool:ChangeRome User
    로마자 사용자 데이터
    파라미터셋: ChangeRomeUser
*/
func (p *ActionTool) ChangeRomeUser(arg ...interface{}) {
	p.hwp.ActionWithParameters("ChangeRome User", arg...)
}

/*
ActionTool:ChangeRome User String
    로마자 사용자 데이터 추가
    파라미터셋: ChangeRomeUser
*/
func (p *ActionTool) ChangeRomeUserString(arg ...interface{}) {
	p.hwp.ActionWithParameters("ChangeRome User String", arg...)
}

/*
ActionTool:SearchForeign
    외래어 사전 검색
    파라미터셋: SearchForeign
*/
func (p *ActionTool) SearchForeign(arg ...interface{}) {
	p.hwp.ActionWithParameters("SearchForeign", arg...)
}

/*
ActionTool:SearchAddress
    주소 검색
    파라미터셋: SearchAddress
*/
func (p *ActionTool) SearchAddress(arg ...interface{}) {
	p.hwp.ActionWithParameters("SearchAddress", arg...)
}

/*
ActionTool:IndexMark
    찾아보기 표시
    파라미터셋: IndexMark
*/
func (p *ActionTool) IndexMark(arg ...interface{}) {
	p.hwp.ActionWithParameters("IndexMark", arg...)
}

/*
ActionTool:MakeIndex
    찾아보기 만들기
    파라미터셋: 없음
*/
func (p *ActionTool) MakeIndex() {
	p.hwp.Run("MakeIndex")
}

/*
ActionTool:MakeContents
    차례 만들기
    파라미터셋: MakeContents
*/
func (p *ActionTool) MakeContents(arg ...interface{}) {
	p.hwp.ActionWithParameters("MakeContents", arg...)
}

/*
ActionTool:IndexMarkModify
    찾아보기 표시 고치기
    파라미터셋: IndexMark
*/
func (p *ActionTool) IndexMarkModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("IndexMarkModify", arg...)
}

/*
ActionTool:MarkTitle
    제목 차례 표시
    파라미터셋: 없음
*/
func (p *ActionTool) MarkTitle() {
	p.hwp.Run("MarkTitle")
}

/*
ActionTool:HideTitle
    제목 차례 숨기기
    파라미터셋: 없음
*/
func (p *ActionTool) HideTitle() {
	p.hwp.Run("HideTitle")
}

/*
ActionTool:MacroDefine
    키 매크로 정의
    파라미터셋: KeyMacro
*/
func (p *ActionTool) MacroDefine(arg ...interface{}) {
	p.hwp.ActionWithParameters("MacroDefine", arg...)
}

/*
ActionTool:MacroRepeatDlg
    키 매크로 실행(대화상자)
    파라미터셋: KeyMacro
*/
func (p *ActionTool) MacroRepeatDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("MacroRepeatDlg", arg...)
}

/*
ActionTool:MacroRepeat
    키 매크로 실행
    파라미터셋: 없음
*/
func (p *ActionTool) MacroRepeat() {
	p.hwp.Run("MacroRepeat")
}

/*
ActionTool:MacroPause
    키 매크로 실행 일시 중지 (정의/실행)
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPause() {
	p.hwp.Run("MacroPause")
}

/*
ActionTool:MacroStop
    키 매크로 실행 중지 (정의/실행)
    파라미터셋: 없음
*/
func (p *ActionTool) MacroStop() {
	p.hwp.Run("MacroStop")
}

/*
ActionTool:MacroPlay1
    키 매크로 1
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay1() {
	p.hwp.Run("MacroPlay1")
}

/*
ActionTool:MacroPlay2
    키 매크로 2
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay2() {
	p.hwp.Run("MacroPlay2")
}

/*
ActionTool:MacroPlay3
    키 매크로 3
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay3() {
	p.hwp.Run("MacroPlay3")
}

/*
ActionTool:MacroPlay4
    키 매크로 4
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay4() {
	p.hwp.Run("MacroPlay4")
}

/*
ActionTool:MacroPlay5
    키 매크로 5
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay5() {
	p.hwp.Run("MacroPlay5")
}

/*
ActionTool:MacroPlay6
    키 매크로 6
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay6() {
	p.hwp.Run("MacroPlay6")
}

/*
ActionTool:MacroPlay7
    키 매크로 7
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay7() {
	p.hwp.Run("MacroPlay7")
}

/*
ActionTool:MacroPlay8
    키 매크로 8
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay8() {
	p.hwp.Run("MacroPlay8")
}

/*
ActionTool:MacroPlay9
    키 매크로 9
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay9() {
	p.hwp.Run("MacroPlay9")
}

/*
ActionTool:MacroPlay10
    키 매크로 10
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay10() {
	p.hwp.Run("MacroPlay10")
}

/*
ActionTool:MacroPlay11
    키 매크로 11
    파라미터셋: 없음
*/
func (p *ActionTool) MacroPlay11() {
	p.hwp.Run("MacroPlay11")
}

/*
ActionTool:ScrMacroDefine
    스크립트 매크로 정의
    파라미터셋: ScriptMacro
*/
func (p *ActionTool) ScrMacroDefine(arg ...interface{}) {
	p.hwp.ActionWithParameters("ScrMacroDefine", arg...)
}

/*
ActionTool:ScrMacroRepeatDlg
    스크립트 매크로 실행
    파라미터셋: ScriptMacro
*/
func (p *ActionTool) ScrMacroRepeatDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("ScrMacroRepeatDlg", arg...)
}

/*
ActionTool:ScrMacroRepeat
    스크립트 매크로 실행
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroRepeat() {
	p.hwp.Run("ScrMacroRepeat")
}

/*
ActionTool:ScrMacroPause
    스크립트 매크로 실행 일시 중지 (정의/실행)
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPause() {
	p.hwp.Run("ScrMacroPause")
}

/*
ActionTool:ScrMacroStop
    스크립트 매크로 실행 중지 (정의/실행)
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroStop() {
	p.hwp.Run("ScrMacroStop")
}

/*
ActionTool:ScrMacroPlay1
    스크립트 매크로 1
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay1() {
	p.hwp.Run("ScrMacroPlay1")
}

/*
ActionTool:ScrMacroPlay2
    스크립트 매크로 2
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay2() {
	p.hwp.Run("ScrMacroPlay2")
}

/*
ActionTool:ScrMacroPlay3
    스크립트 매크로 3
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay3() {
	p.hwp.Run("ScrMacroPlay3")
}

/*
ActionTool:ScrMacroPlay4
    스크립트 매크로 4
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay4() {
	p.hwp.Run("ScrMacroPlay4")
}

/*
ActionTool:ScrMacroPlay5
    스크립트 매크로 5
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay5() {
	p.hwp.Run("ScrMacroPlay5")
}

/*
ActionTool:ScrMacroPlay6
    스크립트 매크로 6
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay6() {
	p.hwp.Run("ScrMacroPlay6")
}

/*
ActionTool:ScrMacroPlay7
    스크립트 매크로 7
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay7() {
	p.hwp.Run("ScrMacroPlay7")
}

/*
ActionTool:ScrMacroPlay8
    스크립트 매크로 8
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay8() {
	p.hwp.Run("ScrMacroPlay8")
}

/*
ActionTool:ScrMacroPlay9
    스크립트 매크로 9
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay9() {
	p.hwp.Run("ScrMacroPlay9")
}

/*
ActionTool:ScrMacroPlay10
    스크립트 매크로 10
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay10() {
	p.hwp.Run("ScrMacroPlay10")
}

/*
ActionTool:ScrMacroPlay11
    스크립트 매크로 11
    파라미터셋: 없음
*/
func (p *ActionTool) ScrMacroPlay11() {
	p.hwp.Run("ScrMacroPlay11")
}

/*
ActionTool:LabelAdd
    라벨 추가
    파라미터셋: 없음
*/
func (p *ActionTool) LabelAdd() {
	p.hwp.Run("LabelAdd")
}

/*
ActionTool:LabelTemplate
    라벨 문서 만들기
    파라미터셋: Label
*/
func (p *ActionTool) LabelTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("LabelTemplate", arg...)
}

/*
ActionTool:ManuScriptTemplate
    원고지 쓰기
    파라미터셋: FileOpen
*/
func (p *ActionTool) ManuScriptTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("ManuScriptTemplate", arg...)
}

/*
ActionTool:Presentation
    프레젠테이션
    파라미터셋: Presentation
*/
func (p *ActionTool) Presentation(arg ...interface{}) {
	p.hwp.ActionWithParameters("Presentation", arg...)
}

/*
ActionTool:PresentationSetup
    프레젠테이션 설정
    파라미터셋: Presentation
*/
func (p *ActionTool) PresentationSetup(arg ...interface{}) {
	p.hwp.ActionWithParameters("PresentationSetup", arg...)
}

/*
ActionTool:CaptureHandler
    갈무리 시작
    파라미터셋: 없음
*/
func (p *ActionTool) CaptureHandler() {
	p.hwp.Run("CaptureHandler")
}

/*
ActionTool:CaptureDialog
    갈무리 끝
    파라미터셋: 없음
*/
func (p *ActionTool) CaptureDialog() {
	p.hwp.Run("CaptureDialog")
}

/*
ActionTool:MailMergeField
    메일 머지 필드(표시달기 or 고치기)
    파라미터셋: 없음
*/
func (p *ActionTool) MailMergeField() {
	p.hwp.Run("MailMergeField")
}

/*
ActionTool:MailMergeInsert
    메일 머지 표시 달기
    파라미터셋: FieldCtrl
*/
func (p *ActionTool) MailMergeInsert(arg ...interface{}) {
	p.hwp.ActionWithParameters("MailMergeInsert", arg...)
}

/*
ActionTool:MailMergeModify
    메일 머지 고치기
    파라미터셋: FieldCtrl
*/
func (p *ActionTool) MailMergeModify(arg ...interface{}) {
	p.hwp.ActionWithParameters("MailMergeModify", arg...)
}

/*
ActionTool:MailMergeGenerate
    메일 머지 만들기
    파라미터셋: MailMergeGenerate
*/
func (p *ActionTool) MailMergeGenerate(arg ...interface{}) {
	p.hwp.ActionWithParameters("MailMergeGenerate", arg...)
}

/*
ActionTool:Sort
    소트
    파라미터셋: Sort
*/
func (p *ActionTool) Sort(arg ...interface{}) {
	p.hwp.ActionWithParameters("Sort", arg...)
}

/*
ActionTool:Sum
    블록 합계
    파라미터셋: Sum
*/
func (p *ActionTool) Sum(arg ...interface{}) {
	p.hwp.ActionWithParameters("Sum", arg...)
}

/*
ActionTool:Average
    블록 평균
    파라미터셋: Sum
*/
func (p *ActionTool) Average(arg ...interface{}) {
	p.hwp.ActionWithParameters("Average", arg...)
}

/*
ActionTool:QuickCommand
    입력 자동 명령 동작 - 메뉴에서 동작 체크 On/Off
    파라미터셋: QCorrect
*/
func (p *ActionTool) QuickCommandRun(arg ...interface{}) {
	p.hwp.ActionWithParameters("QuickCommand Run", arg...)
}

/*
ActionTool:AQcommandMerge
    입력 장동 명령
    파라미터셋: UserQCommandFile
*/
func (p *ActionTool) AQcommandMerge(arg ...interface{}) {
	p.hwp.ActionWithParameters("AQcommandMerge", arg...)
}

/*
ActionTable:TableCreate
    표 만들기
    파라미터셋: TableCreation
*/
func (p *ActionTable) TableCreate(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableCreate", arg...)
}

/*
ActionTable:TableFormula
    계산식
    파라미터셋: FieldCtrl
*/
func (p *ActionTable) TableFormula(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableFormula", arg...)
}

/*
ActionTable:TableFormulaSumHor
    가로 합계
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaSumHor() {
	p.hwp.Run("TableFormulaSumHor")
}

/*
ActionTable:TableFormulaSumVer
    세로 합계
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaSumVer() {
	p.hwp.Run("TableFormulaSumVer")
}

/*
ActionTable:TableFormulaAvgHor
    가로 평균
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaAvgHor() {
	p.hwp.Run("TableFormulaAvgHor")
}

/*
ActionTable:TableFormulaAvgVer
    세로 평균
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaAvgVer() {
	p.hwp.Run("TableFormulaAvgVer")
}

/*
ActionTable:TableFormulaProHor
    가로 곱
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaProHor() {
	p.hwp.Run("TableFormulaProHor")
}

/*
ActionTable:TableFormulaProVer
    세로 곱
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaProVer() {
	p.hwp.Run("TableFormulaProVer")
}

/*
ActionTable:TableFormulaSumAuto
    블록 합계
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaSumAuto() {
	p.hwp.Run("TableFormulaSumAuto")
}

/*
ActionTable:TableFormulaAvgAuto
    블록 평균
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaAvgAuto() {
	p.hwp.Run("TableFormulaAvgAuto")
}

/*
ActionTable:TableFormulaProAuto
    블록 곱
    파라미터셋: 없음
*/
func (p *ActionTable) TableFormulaProAuto() {
	p.hwp.Run("TableFormulaProAuto")
}

/*
ActionTable:TablePropertyDialog
    표 고치기
    파라미터셋: ShapeObject
*/
func (p *ActionTable) TablePropertyDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("TablePropertyDialog", arg...)
}

/*
ActionTable:TableCellBlock
    셀 블록
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellBlock() {
	p.hwp.Run("TableCellBlock")
}

/*
ActionTable:TableCellBlockExtend
    셀 블록 연장(F5 + F5)
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellBlockExtend() {
	p.hwp.Run("TableCellBlockExtend")
}

/*
ActionTable:TableCellBlockExtendAbs
    셀 블록 연장(SHIFT + F5)
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellBlockExtendAbs() {
	p.hwp.Run("TableCellBlockExtendAbs")
}

/*
ActionTable:TableCellBlockCol
    셀 블록 (칸)
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellBlockCol() {
	p.hwp.Run("TableCellBlockCol")
}

/*
ActionTable:TableCellBlockRow
    셀 블록 (줄)
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellBlockRow() {
	p.hwp.Run("TableCellBlockRow")
}

/*
ActionTable:TableMergeCell
    셀 합치기
    파라미터셋: 없음
*/
func (p *ActionTable) TableMergeCell() {
	p.hwp.Run("TableMergeCell")
}

/*
ActionTable:TableSplitCell
    셀 나누기
    파라미터셋: TableSplitCell
*/
func (p *ActionTable) TableSplitCell(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableSplitCell", arg...)
}

/*
ActionTable:TableSplitCellRow2
    셀 줄 나누기
    파라미터셋: TableSplitCell
*/
func (p *ActionTable) TableSplitCellRow2(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableSplitCellRow2", arg...)
}

/*
ActionTable:TableSplitCellCol2
    셀 칸 나누기
    파라미터셋: TableSplitCell
*/
func (p *ActionTable) TableSplitCellCol2(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableSplitCellCol2", arg...)
}

/*
ActionTable:TableVAlignTop
    셀 세로 정렬 위
    파라미터셋: 없음
*/
func (p *ActionTable) TableVAlignTop() {
	p.hwp.Run("TableVAlignTop")
}

/*
ActionTable:TableVAlignCenter
    셀 세로정렬 가운데
    파라미터셋: 없음
*/
func (p *ActionTable) TableVAlignCenter() {
	p.hwp.Run("TableVAlignCenter")
}

/*
ActionTable:TableVAlignBottom
    셀 세로정렬 아래
    파라미터셋: 없음
*/
func (p *ActionTable) TableVAlignBottom() {
	p.hwp.Run("TableVAlignBottom")
}

/*
ActionTable:TableDistributeCellWidth
    셀 너비를 같게
    파라미터셋: 없음
*/
func (p *ActionTable) TableDistributeCellWidth() {
	p.hwp.Run("TableDistributeCellWidth")
}

/*
ActionTable:TableDistributeCellHeight
    셀 높이를 같게
    파라미터셋: 없음
*/
func (p *ActionTable) TableDistributeCellHeight() {
	p.hwp.Run("TableDistributeCellHeight")
}

/*
ActionTable:TableDeleteRowColumn
    줄/칸 지우기
    파라미터셋: TableDeleteLine
*/
func (p *ActionTable) TableDeleteRowColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableDeleteRowColumn", arg...)
}

/*
ActionTable:TableDeleteColumn
    칸 지우기
    파라미터셋: TableDeleteLine
*/
func (p *ActionTable) TableDeleteColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableDeleteColumn", arg...)
}

/*
ActionTable:TableDeleteRow
    줄/칸 지우기
    파라미터셋: TableDeleteLine
*/
func (p *ActionTable) TableDeleteRow(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableDeleteRow", arg...)
}

/*
ActionTable:TableInsertRowColumn
    줄/칸 삽입
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableInsertRowColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableInsertRowColumn", arg...)
}

/*
ActionTable:TableInsertLeftColumn
    왼쪽 칸 삽입
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableInsertLeftColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableInsertLeftColumn", arg...)
}

/*
ActionTable:TableInsertRightColumn
    오른쪽 칸 삽입
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableInsertRightColumn(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableInsertRightColumn", arg...)
}

/*
ActionTable:TableInsertUpperRow
    위쪽 줄 삽입
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableInsertUpperRow(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableInsertUpperRow", arg...)
}

/*
ActionTable:TableInsertLowerRow
    아래쪽 줄 삽입
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableInsertLowerRow(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableInsertLowerRow", arg...)
}

/*
ActionTable:TableAppendRow
    줄 추가
    파라미터셋: TableInsertLine
*/
func (p *ActionTable) TableAppendRow(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableAppendRow", arg...)
}

/*
ActionTable:TableSubtractRow
    줄 제거
    파라미터셋: TableDeleteLine
*/
func (p *ActionTable) TableSubtractRow(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableSubtractRow", arg...)
}

/*
ActionTable:CellBorderFill
    셀 테두리 배경
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellBorderFill(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellBorderFill", arg...)
}

/*
ActionTable:CellBorder
    셀 테두리 배경
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellBorder(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellBorder", arg...)
}

/*
ActionTable:CellFill
    셀 배경
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellFill(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellFill", arg...)
}

/*
ActionTable:CellZoneBorderFill
    여러 셀에 걸쳐 적용
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellZoneBorderFill(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellZoneBorderFill", arg...)
}

/*
ActionTable:CellZoneBorder
    셀 테두리 적용
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellZoneBorder(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellZoneBorder", arg...)
}

/*
ActionTable:CellZoneFill
    셀 배경에 적용
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) CellZoneFill(arg ...interface{}) {
	p.hwp.ActionWithParameters("CellZoneFill", arg...)
}

/*
ActionTable:TableDrawPen
    표 그리기
    파라미터셋: TableDrawPen
*/
func (p *ActionTable) TableDrawPen(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableDrawPen", arg...)
}

/*
ActionTable:TableEraser
    표 지우개
    파라미터셋: TableDrawPen
*/
func (p *ActionTable) TableEraser(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableEraser", arg...)
}

/*
ActionTable:TableDeleteCell
    셀 삭제
    파라미터셋: SelectionOpt
*/
func (p *ActionTable) TableDeleteCell(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableDeleteCell", arg...)
}

/*
ActionTable:TableAutoDrawPenStyleWidthDlg
    선 모양
    파라미터셋: TableDrawPen
*/
func (p *ActionTable) TableAutoDrawPenStyleWidthDlg(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableAutoDrawPenStyleWidthDlg", arg...)
}

/*
ActionTable:TableCellShadeInc
    셀 음영 비율 증가
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) TableCellShadeInc(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableCellShadeInc", arg...)
}

/*
ActionTable:TableCellShadeDec
    셀 음영 비율 감소
    파라미터셋: CellBorderFill
*/
func (p *ActionTable) TableCellShadeDec(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableCellShadeDec", arg...)
}

/*
ActionTable:TableSplitTable
    표 나누기
    파라미터셋: 없음
*/
func (p *ActionTable) TableSplitTable() {
	p.hwp.Run("TableSplitTable")
}

/*
ActionTable:TableMergeTable
    표 붙이기
    파라미터셋: 없음
*/
func (p *ActionTable) TableMergeTable() {
	p.hwp.Run("TableMergeTable")
}

/*
ActionTable:TableTableToString
    표를 문자열로
    파라미터셋: TableTblToStr
*/
func (p *ActionTable) TableTableToString(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableTableToString", arg...)
}

/*
ActionTable:TableStringToTable
    문자열을 표로
    파라미터셋: TableStrToTbl
*/
func (p *ActionTable) TableStringToTable(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableStringToTable", arg...)
}

/*
ActionTable:TableAutoFill
    자동 채우기
    파라미터셋: 없음
*/
func (p *ActionTable) TableAutoFill() {
	p.hwp.Run("TableAutoFill")
}

/*
ActionTable:TableAutoFillDlg
    자동 채우기 내용
    파라미터셋: 없음
*/
func (p *ActionTable) TableAutoFillDlg() {
	p.hwp.Run("TableAutoFillDlg")
}

/*
ActionTable:TableCellAlignLeftTop
    셀 왼쪽 위 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignLeftTop() {
	p.hwp.Run("TableCellAlignLeftTop")
}

/*
ActionTable:TableCellAlignLeftCenter
    셀 왼쪽 가운데 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignLeftCenter() {
	p.hwp.Run("TableCellAlignLeftCenter")
}

/*
ActionTable:TableCellAlignLeftBottom
    셀 왼쪽 아래 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignLeftBottom() {
	p.hwp.Run("TableCellAlignLeftBottom")
}

/*
ActionTable:TableCellAlignCenterTop
    셀 가운데 위 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignCenterTop() {
	p.hwp.Run("TableCellAlignCenterTop")
}

/*
ActionTable:TableCellAlignCenterCenter
    셀 가운데 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignCenterCenter() {
	p.hwp.Run("TableCellAlignCenterCenter")
}

/*
ActionTable:TableCellAlignCenterBottom
    셀 가운데 아래 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignCenterBottom() {
	p.hwp.Run("TableCellAlignCenterBottom")
}

/*
ActionTable:TableCellAlignRightTop
    셀 가운데 오른쪽 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignRightTop() {
	p.hwp.Run("TableCellAlignRightTop")
}

/*
ActionTable:TableCellAlignRightCenter
    셀 오른쪽 가운데 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignRightCenter() {
	p.hwp.Run("TableCellAlignRightCenter")
}

/*
ActionTable:TableCellAlignRightBottom
    셀 오른쪽 아래 정렬
    파라미터셋: 없음
*/
func (p *ActionTable) TableCellAlignRightBottom() {
	p.hwp.Run("TableCellAlignRightBottom")
}

/*
ActionTable:TableTemplate
    표 마당
    파라미터셋: TableTemplate
*/
func (p *ActionTable) TableTemplate(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableTemplate", arg...)
}

/*
ActionTable:InsertChart
    차트 만들기
    파라미터셋: OleCreation
*/
func (p *ActionTable) InsertChart(arg ...interface{}) {
	p.hwp.ActionWithParameters("InsertChart", arg...)
}

/*
ActionTable:TableSwap
    표 뒤집기
    파라미터셋: TableSwap
*/
func (p *ActionTable) TableSwap(arg ...interface{}) {
	p.hwp.ActionWithParameters("TableSwap", arg...)
}

/*
ActionTable:TableLowerCell
    아래 셀
    셀 선택 후 위 아래로 올리고 내릴때
*/
func (p *ActionTable) TableLowerCell() {
	p.hwp.Run("TableLowerCell")
}

/*
ActionTable:TableUpperCell
    윗 셀
    셀 선택 후 위 아래로 올리고 내릴때
*/
func (p *ActionTable) TableUpperCell() {
	p.hwp.Run("TableUpperCell")
}

/*
ActionTable:TableCellBorderNo
    표 셀 선 제거
*/
func (p *ActionTable) TableCellBorderNo() {
	p.hwp.Run("TableCellBorderNo")
}

/*
ActionDraw:ShapeObjAttrDialog
    특 속성 환경 설정
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) ShapeObjAttrDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeObjAttrDialog", arg...)
}

/*
ActionDraw:ShapeObjLineProperty
    고치기 대화상자중 line tab
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjLineProperty() {
	p.hwp.Run("ShapeObjLineProperty")
}

/*
ActionDraw:ShapeObjFillProperty
    고치기 대화상자중 fill tab
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjFillProperty() {
	p.hwp.Run("ShapeObjFillProperty")
}

/*
ActionDraw:DrawObjCreatorLine
    선 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorLine(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorLine", arg...)
}

/*
ActionDraw:DrawObjCreatorRectangle
    사각형 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorRectangle(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorRectangle", arg...)
}

/*
ActionDraw:DrawObjCreatorEllipse
    원 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorEllipse(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorEllipse", arg...)
}

/*
ActionDraw:DrawObjCreatorArc
    호 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorArc(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorArc", arg...)
}

/*
ActionDraw:DrawObjCreatorPolygon
    다각형 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorPolygon(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorPolygon", arg...)
}

/*
ActionDraw:DrawObjCreatorCurve
    곡선 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorCurve(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorCurve", arg...)
}

/*
ActionDraw:DrawObjCreatorFreeDrawing
    펜
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorFreeDrawing(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorFreeDrawing", arg...)
}

/*
ActionDraw:FormObjCreatorPushButton
    양식 개체 푸쉬 버튼 넣기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjCreatorPushButton() {
	p.hwp.Run("FormObjCreatorPushButton")
}

/*
ActionDraw:FormObjCreatorRadioButton
    양식 개체 라디오 버튼 넣기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjCreatorRadioButton() {
	p.hwp.Run("FormObjCreatorRadioButton")
}

/*
ActionDraw:FormObjCreatorCheckButton
    양식 개체 체크 박스 넣기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjCreatorCheckButton() {
	p.hwp.Run("FormObjCreatorCheckButton")
}

/*
ActionDraw:FormObjCreatorComboBox
    양식 개체 콤보 박스 넣기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjCreatorComboBox() {
	p.hwp.Run("FormObjCreatorComboBox")
}

/*
ActionDraw:FormObjCreatorEdit
    양식 개체 에디트 박스 넣기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjCreatorEdit() {
	p.hwp.Run("FormObjCreatorEdit")
}

/*
ActionDraw:FormObjectPropertyDialog
    양식 개체 고치기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) FormObjectPropertyDialog(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjectPropertyDialog", arg...)
}

/*
ActionDraw:FormDesignMode
    양식 개체 디자인 모드 변경
    파라미터셋: 없음
*/
func (p *ActionDraw) FormDesignMode() {
	p.hwp.Run("FormDesignMode")
}

/*
ActionDraw:FormObjRadioGroup
    양식 개체 라디오 버튼 그룹 묶기
    파라미터셋: 없음
*/
func (p *ActionDraw) FormObjRadioGroup() {
	p.hwp.Run("FormObjRadioGroup")
}

/*
ActionDraw:FormObjCharShape
    양식 개체 글자 모양 속성 변경
    파라미터셋: CharShape
*/
func (p *ActionDraw) FormObjCharShape(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjCharShape", arg...)
}

/*
ActionDraw:FormObjInputCodeTable
    양식 개체 Edit에서만 사용하는 문자표
    파라미터셋: FormObjInputCodeTable
*/
func (p *ActionDraw) FormObjInputCodeTable(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjInputCodeTable", arg...)
}

/*
ActionDraw:FormObjInputHanja
    양식 개체 Edit에서만 사용하는 한자로 바꾸기
    파라미터셋: FormObjInputHanja
*/
func (p *ActionDraw) FormObjInputHanja(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjInputHanja", arg...)
}

/*
ActionDraw:FormObjHanjaBusu
    양식 개체 Edit에서만 사용하는 한자부수 입력
    파라미터셋: FormObjHanjaBusu
*/
func (p *ActionDraw) FormObjHanjaBusu(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjHanjaBusu", arg...)
}

/*
ActionDraw:FormObjHanjaMean
    양식 개체 Edit에서만 사용하는 한자새김 입력
    파라미터셋: FormObjHanjaMean
*/
func (p *ActionDraw) FormObjHanjaMean(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjHanjaMean", arg...)
}

/*
ActionDraw:FormObjInputIdiom
    양식 개체 Edit에서만 사용하는 상용구 입력
    파라미터셋: FormObjInputIdiom
*/
func (p *ActionDraw) FormObjInputIdiom(arg ...interface{}) {
	p.hwp.ActionWithParameters("FormObjInputIdiom", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiLine
    반복해서 선 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiLine(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiLine", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiRectangle
    반복해서 사각형 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiRectangle(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiRectangle", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiEllipse
    반복해서 원 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiEllipse(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiEllipse", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiArc
    반복해서 호 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiArc(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiArc", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiPolygon
    반복해서 다각형 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiPolygon(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiPolygon", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiCurve
    반복해서 곡선 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiCurve(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiCurve", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiFreeDrawing
    반복해서 펜 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiFreeDrawing(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiFreeDrawing", arg...)
}

/*
ActionDraw:DrawObjCreatorMultiTextBox
    반복해서 글상자 그리기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) DrawObjCreatorMultiTextBox(arg ...interface{}) {
	p.hwp.ActionWithParameters("DrawObjCreatorMultiTextBox", arg...)
}

/*
ActionDraw:DrawObjCancelOneStep
    다각형(곡선) 그리는 중 이전 선 지우기
    파라미터셋: 없음
*/
func (p *ActionDraw) DrawObjCancelOneStep() {
	p.hwp.Run("DrawObjCancelOneStep")
}

/*
ActionDraw:ShapeObjRotater
    자유 각 회전(회전 중심 고정)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjRotater() {
	p.hwp.Run("ShapeObjRotater")
}

/*
ActionDraw:ShapeObjRandomAngleRotater
    자유 각 회전
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) ShapeObjRandomAngleRotater(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeObjRandomAngleRotater", arg...)
}

/*
ActionDraw:ShapeObjRightAngleRotater
    90도 회전
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjRightAngleRotater() {
	p.hwp.Run("ShapeObjRightAngleRotater")
}

/*
ActionDraw:ShapeObjShear
    기울이기
    파라미터셋: ShapeObject
*/
func (p *ActionDraw) ShapeObjShear(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeObjShear", arg...)
}

/*
ActionDraw:DrawObjEditDetail
    그리기 개체 편집
    파라미터셋: 없음
*/
func (p *ActionDraw) DrawObjEditDetail() {
	p.hwp.Run("DrawObjEditDetail")
}

/*
ActionDraw:ShapeObjSelect
    튼 선택 도구
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjSelect() {
	p.hwp.Run("ShapeObjSelect")
}

/*
ActionDraw:ShapeObjNorm
    기본 도형 설정
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjNorm() {
	p.hwp.Run("ShapeObjNorm")
}

/*
ActionDraw:ShapeObjGroup
    틀 묶기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjGroup() {
	p.hwp.Run("ShapeObjGroup")
}

/*
ActionDraw:ShapeObjUngroup
    틀 풀기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjUngroup() {
	p.hwp.Run("ShapeObjUngroup")
}

/*
ActionDraw:ShapeObjBringToFront
    맨 앞으로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjBringToFront() {
	p.hwp.Run("ShapeObjBringToFront")
}

/*
ActionDraw:ShapeObjSendToBack
    맨 뒤로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjSendToBack() {
	p.hwp.Run("ShapeObjSendToBack")
}

/*
ActionDraw:ShapeObjBringForward
    앞으로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjBringForward() {
	p.hwp.Run("ShapeObjBringForward")
}

/*
ActionDraw:ShapeObjSendBack
    뒤로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjSendBack() {
	p.hwp.Run("ShapeObjSendBack")
}

/*
ActionDraw:ShapeObjBringInFrontOfText
    글 앞으로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjBringInFrontOfText() {
	p.hwp.Run("ShapeObjBringInFrontOfText")
}

/*
ActionDraw:ShapeObjCtrlSendBehindText
    글 뒤로
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjCtrlSendBehindText() {
	p.hwp.Run("ShapeObjCtrlSendBehindText")
}

/*
ActionDraw:ShapeObjWrapTopAndBottom
    자리 차지
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjWrapTopAndBottom() {
	p.hwp.Run("ShapeObjWrapTopAndBottom")
}

/*
ActionDraw:ShapeObjWrapSquare
    직사각형
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjWrapSquare() {
	p.hwp.Run("ShapeObjWrapSquare")
}

/*
ActionDraw:ShapeObjPrevObject
    이전 개체로 이동(shift + tab키)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjPrevObject() {
	p.hwp.Run("ShapeObjPrevObject")
}

/*
ActionDraw:ShapeObjNextObject
    이후 개채로 이동(tab키)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjNextObject() {
	p.hwp.Run("ShapeObjNextObject")
}

/*
ActionDraw:ShapeObjHorzFlip
    그리기 개체 좌우 뒤집기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjHorzFlip() {
	p.hwp.Run("ShapeObjHorzFlip")
}

/*
ActionDraw:ShapeObjVertFlip
    그리기 개체 상하 뒤집기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjVertFlip() {
	p.hwp.Run("ShapeObjVertFlip")
}

/*
ActionDraw:ShapeObjLock
    개체 Lock
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjLock() {
	p.hwp.Run("ShapeObjLock")
}

/*
ActionDraw:ShapeObjUnlockAll
    개체 Unlock All
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjUnlockAll() {
	p.hwp.Run("ShapeObjUnlockAll")
}

/*
ActionDraw:ShapeObjAlignTop
    위로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignTop() {
	p.hwp.Run("ShapeObjAlignTop")
}

/*
ActionDraw:ShapeObjAlignMiddle
    중간 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignMiddle() {
	p.hwp.Run("ShapeObjAlignMiddle")
}

/*
ActionDraw:ShapeObjAlignBottom
    아래로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignBottom() {
	p.hwp.Run("ShapeObjAlignBottom")
}

/*
ActionDraw:ShapeObjAlignLeft
    왼쪽으로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignLeft() {
	p.hwp.Run("ShapeObjAlignLeft")
}

/*
ActionDraw:ShapeObjAlignCenter
    가운데로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignCenter() {
	p.hwp.Run("ShapeObjAlignCenter")
}

/*
ActionDraw:ShapeObjAlignRight
    오른쪽으로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignRight() {
	p.hwp.Run("ShapeObjAlignRight")
}

/*
ActionDraw:ShapeObjAlignVertSpacing
    위/아래 일정한 비율로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignVertSpacing() {
	p.hwp.Run("ShapeObjAlignVertSpacing")
}

/*
ActionDraw:ShapeObjAlignHorzSpacing
    왼쪽/오른쪽 일정한 비율로 정렬
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignHorzSpacing() {
	p.hwp.Run("ShapeObjAlignHorzSpacing")
}

/*
ActionDraw:ShapeObjAlignWidth
    폭 맞춤
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignWidth() {
	p.hwp.Run("ShapeObjAlignWidth")
}

/*
ActionDraw:ShapeObjAlignHeight
    높이 맞춤
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignHeight() {
	p.hwp.Run("ShapeObjAlignHeight")
}

/*
ActionDraw:ShapeObjAlignSize
    폭/높이 맞춤
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAlignSize() {
	p.hwp.Run("ShapeObjAlignSize")
}

/*
ActionDraw:ShapeObjMoveLeft
    키로 움직이기(왼쪽)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjMoveLeft() {
	p.hwp.Run("ShapeObjMoveLeft")
}

/*
ActionDraw:ShapeObjMoveRight
    키로 움직이기(오른쪽)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjMoveRight() {
	p.hwp.Run("ShapeObjMoveRight")
}

/*
ActionDraw:ShapeObjMoveUp
    키로 움직이기(위)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjMoveUp() {
	p.hwp.Run("ShapeObjMoveUp")
}

/*
ActionDraw:ShapeObjMoveDown
    키로 움직이기(아래)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjMoveDown() {
	p.hwp.Run("ShapeObjMoveDown")
}

/*
ActionDraw:ShapeObjResizeLeft
    키로 크기 조절(shift + 왼쪽)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjResizeLeft() {
	p.hwp.Run("ShapeObjResizeLeft")
}

/*
ActionDraw:ShapeObjResizeRight
    키로 크기 조절(shift + 오른쪽)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjResizeRight() {
	p.hwp.Run("ShapeObjResizeRight")
}

/*
ActionDraw:ShapeObjResizeUp
    키로 크기 조절(shift + 위)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjResizeUp() {
	p.hwp.Run("ShapeObjResizeUp")
}

/*
ActionDraw:ShapeObjResizeDown
    키로 크기 조절(shift + 아래)
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjResizeDown() {
	p.hwp.Run("ShapeObjResizeDown")
}

/*
ActionDraw:DrawObjTemplateSave
    그리기 마당에 등록
    파라미터셋: 없음
*/
func (p *ActionDraw) DrawObjTemplateSave() {
	p.hwp.Run("DrawObjTemplateSave")
}

/*
ActionDraw:ShapeObjAttachTextBox
    글상자로 만들기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjAttachTextBox() {
	p.hwp.Run("ShapeObjAttachTextBox")
}

/*
ActionDraw:ShapeObjDetachTextBox
    글상자 속성 없애기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjDetachTextBox() {
	p.hwp.Run("ShapeObjDetachTextBox")
}

/*
ActionDraw:ShapeObjTextBoxEdit
    글상자 선택상태에서 편집모드로 들어가기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjTextBoxEdit() {
	p.hwp.Run("ShapeObjTextBoxEdit")
}

/*
ActionDraw:ShapeObjTableSelCell
    테이블 선택 상태에서 첫 번째 셀 선택하기
    파라미터셋: 없음
*/
func (p *ActionDraw) ShapeObjTableSelCell() {
	p.hwp.Run("ShapeObjTableSelCell")
}

/*
ActionDraw:ShapeObjSaveAsPicture
    그리기개체를 그림으로 저장하기
    파라미터셋: ShapeObjSaveAsPicture
*/
func (p *ActionDraw) ShapeObjSaveAsPicture(arg ...interface{}) {
	p.hwp.ActionWithParameters("ShapeObjSaveAsPicture", arg...)
}

/*
ActionDraw:DrawObjOpenClosePolygon
    다각형 열기/닫기
    파라미터셋: 없음
*/
func (p *ActionDraw) DrawObjOpenClosePolygon() {
	p.hwp.Run("DrawObjOpenClosePolygon")
}

/*
ActionDraw:ImageFindPath
    그림 경로 찾기
    파라미터셋: 없음
*/
func (p *ActionDraw) ImageFindPath() {
	p.hwp.Run("ImageFindPath")
}

/*
ActionDraw:FillColorShadeInc
    면색 음영 비율 증가
    파라미터셋: 없음
*/
func (p *ActionDraw) FillColorShadeInc() {
	p.hwp.Run("FillColorShadeInc")
}

/*
ActionDraw:FillColorShadeDec
    면색 음영 비율 감소
    파라미터셋: 없음
*/
func (p *ActionDraw) FillColorShadeDec() {
	p.hwp.Run("FillColorShadeDec")
}

/*
ActionEtc:MessageBox
    메시지 박스
    파라미터셋: MessageBox
*/
func (p *ActionEtc) MessageBox(arg ...interface{}) {
	p.hwp.ActionWithParameters("MessageBox", arg...)
}

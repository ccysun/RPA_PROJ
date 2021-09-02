import sys, os, time
import win32com.client as win32
import win32gui, win32con
import pyautogui, ctypes
import numpy

#+----------------------------------------------+
#| Create Instance HWP 아래한글 인스턴스를 생성 |
#+----------------------------------------------+
def CreateInstanceHWP():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  #아래한글 인스턴스/오브젝트 생성
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule") #보안모듈 자동승인 -> regedit에서 등록해야 함.
    hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제
    return hwp


#+-----------------------------+-------------------------------------------------------------------------+
#| Action Run                  | 아래한글에서 설정된 액션을 수행                                         |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             | action 문자열에 뒤에 공백으로 구분하여 여러개의 action를 수행할수 있다. |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             | 예) MoveUp MoveDown MoveRight                                           |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableMergeCell              | 셀 병합 (셀 병합 하기 위해서는 블럭 선택 모드 상태 이어야 함)           |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             | ( TableCellBlock -> TableCellBlockExtend                                |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableCellBlock              | 셀 블럭 모드 (회색 작은 콩)                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableCellBlockExtend        | 셀 블럭 선택 모드 (빨간 작은 콩)                                        |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableLeftCell               | 셀 왼쪽이동                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableRightCell              | 셀 오른쪽 이동                                                          |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableUpperCell              | 셀 위로 이동                                                            |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableLowerCell              | 셀 아래로 이동                                                          |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableColBegin               | 셀 홈키 이동                                                            |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableColEnd                 | 셀 end키 이동                                                           |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableColPageUp              | 셀 페이지up 키 이동                                                     |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableColPageDown            | 셀 페이지 down 키 이동                                                  |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableResizeCellLeft         | 셀 블럭을 왼쪽으로 밀기                                                 |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableResizeCellRight        | 셀 블럭 오른쪽으로 밀기                                                 |
#+-----------------------------+-------------------------------------------------------------------------+
#| BreakPara                   | 문단나누기(enter)                                                       |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveUp                      | 커서 위로 올리기                                                        |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveDown                    | 커서 아래로                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveRight                   | 커서 오른쪽                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveLeft                    | 커서 왼쪽                                                               |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveLineEnd                 | End키                                                                   |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveLineBegin               | Home키                                                                  |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveViewUp                  | pageup키                                                                |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveViewDown                | pagedown 키                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| MovePageBegin               | 현재 페이지의 시작점으로 이동                                           |
#+-----------------------------+-------------------------------------------------------------------------+
#| MovePageEnd                 | 현재 페이지의 끝점으로 이동                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveParaEnd                 | 현재 위치한 문단의 끝으로 이동                                          |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableCellBorderDiagonalUp   | 대각선 사선 넣기(오른쪽에서 왼쪽 아래로)                                |
#+-----------------------------+-------------------------------------------------------------------------+
#| TableCellBorderDiagonalDown | 대각선 사선 넣기 (왼쪽에서 오른쪽 아래로)                               |
#+-----------------------------+-------------------------------------------------------------------------+
#| ParagraphShapeAlignRight    | 오른쪽 정렬                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| ParagraphShapeAlignLeft     | 왼쪽 정렬                                                               |
#+-----------------------------+-------------------------------------------------------------------------+
#| ParagraphShapeAlignCenter   | 가운데 정렬                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| MoveSelLineEnd              | 현재 라인의 문자열 끝까지 선택                                          |
#+-----------------------------+-------------------------------------------------------------------------+
#| Cancel                      | esc 키                                                                  |
#+-----------------------------+-------------------------------------------------------------------------+
#| CharShapeBold               | 문자 진하게                                                             |
#+-----------------------------+-------------------------------------------------------------------------+
#| Select                      | 블럭 선택                                                               |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             | 3번 Select Select Select 하면 한 문단 전체 선택                         |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             | 블럭 해제시 Cancel                                                      |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             |                                                                         |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             |                                                                         |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             |                                                                         |
#+-----------------------------+-------------------------------------------------------------------------+
#|                             |                                                                         |
#+-----------------------------+-------------------------------------------------------------------------+
def ActionRun(hwpObject, action):
    actionList = action.split(" ")
    for actionStr in actionList:
        hwpObject.HAction.Run(actionStr)


#+------------------------------------+
#| 동일한 Action을 여러번 처리시 사용 |
#+------------------------------------+
def RepeatSameActionRun(hwpObject, action, repeatNum):
    if ( str(action).isdigit() == True ):
        raise ValueError("ERROR 두번째 인자값은 Action 명령어를 정확히 입력했는지 확인 필요")
    if ( str(repeatNum).isdigit() == False ):
        raise ValueError("ERROR 세번째 인자값은 정수 입력 필요")
    for i in range(repeatNum):
        hwpObject.HAction.Run(action)

#+---+--------------------------------------------------------------------------------------+
#|   | Page Move 첫 페이지, 끝 페이지, 특정 페이지로 이동 아래한글에 설정된 value 값을 입력 |
#+---+--------------------------------------------------------------------------------------+
#| 2 | 문서 시작으로 이동                                                                   |
#+---+--------------------------------------------------------------------------------------+
#| 3 | 문서 끝으로 이동                                                                     |
#+---+--------------------------------------------------------------------------------------+
#| 8 | 현재 위치한 단어의 시작으로 이동                                                     |
#+---+--------------------------------------------------------------------------------------+
#| 9 | 현재 위치한 단어의 끝으로 이동                                                       |
#+---+--------------------------------------------------------------------------------------+
def PageMove(hwpObject, value):
    hwpObject.MovePos(value)

#+-----------------------------------------+
#| Insert Text 커서 위치에서 문자열을 입력 |
#+-----------------------------------------+
def InsertString(hwpObject, str):
    hInsTxt = hwpObject.HParameterSet.HInsertText
    hAction = hwpObject.HAction
    hAction.GetDefault("InsertText", hInsTxt.HSet)
    hInsTxt.Text = str
    hAction.Execute("InsertText", hInsTxt.HSet)


#+----------------------------------------+
#| Align String 커서 위치에서 문단을 정렬 |
#+----------------------------------------+
def AlignString(hwpObject, hAlign="Center"):
    hParaShp = hwpObject.HParameterSet.HParaShape
    hAction = hwpObject.HAction
    hAction.GetDefault("ParagraphShape", hParaShp.HSet)
    hParaShp.BreakNonLatinWord = 0
    hParaShp.AlignType = hwpObject.HAlign(hAlign)
    hAction.Execute("ParagraphShape", hParaShp.HSet)

#+-----------------------------------------------------------------+
#| Change Font 선택된 문자열등을 입력한 폰트로 변경 및 사이즈 변경 |
#+-----------------------------------------------------------------+
def ChangeFont(hwpObject, fontName, fontSize):
    # 글자 사이즈 변경
    HCharShp = hwpObject.HParameterSet.HCharShape
    hAction = hwpObject.HAction
    hAction.GetDefault("CharShape", HCharShp.HSet)
    HCharShp.FaceNameUser        = fontName
    HCharShp.FontTypeUser        = hwpObject.FontType("TTF")
    HCharShp.FaceNameSymbol      = fontName
    HCharShp.FontTypeSymbol      = hwpObject.FontType("TTF")
    HCharShp.FaceNameOther       = fontName
    HCharShp.FontTypeOther       = hwpObject.FontType("TTF")
    HCharShp.FaceNameJapanese    = fontName
    HCharShp.FontTypeJapanese    = hwpObject.FontType("TTF")
    HCharShp.FaceNameHanja       = fontName
    HCharShp.FontTypeHanja       = hwpObject.FontType("TTF")
    HCharShp.FaceNameLatin       = fontName
    HCharShp.FontTypeLatin       = hwpObject.FontType("TTF")
    HCharShp.FaceNameHangul      = fontName
    HCharShp.FontTypeHangul      = hwpObject.FontType("TTF")
    HCharShp.Height              = hwpObject.PointToHwpUnit(fontSize)
    hAction.Execute("CharShape", HCharShp.HSet)

#+------------------------------------------------------------------------------------------+
#| Create Table 표를 만든다.                                                                |
#| rows  행 개수, cols 컬럼 개수, maxTableWidth: 표 전체 너비,  maxTableHeight 표 전체 높이 |
#+------------------------------------------------------------------------------------------+
def CreateTable(hwpObject, rows, cols, maxTableWidth, maxTableHeight):
    # 표 만들기
    HTableCrt = hwpObject.HParameterSet.HTableCreation
    hAction = hwpObject.HAction
    hAction.GetDefault("TableCreate", HTableCrt.HSet)
    HTableCrt.Rows = rows
    HTableCrt.Cols = cols
    HTableCrt.WidthType = 2
    HTableCrt.HeightType = 1
    HTableCrt.WidthValue = hwpObject.MiliToHwpUnit(148.0)
    HTableCrt.HeightValue = hwpObject.MiliToHwpUnit(150.0)
    uNum = round(maxTableWidth / cols,1)
    hNum = round(maxTableHeight / rows,1)
    HTableCrt.CreateItemArray("ColWidth", cols)
    # Item() 를 SetItem(index, value) 함수로 변경
    for i in range(cols):
        HTableCrt.ColWidth.SetItem(i,hwpObject.MiliToHwpUnit(uNum))
    HTableCrt.CreateItemArray("RowHeight", rows)
    for i in range(rows):
        HTableCrt.RowHeight.SetItem(i,hwpObject.MiliToHwpUnit(hNum))
    HTableCrt.TableProperties.Width = 41954
    hAction.Execute("TableCreate", HTableCrt.HSet)

#+----------------------------------------------+
#| Treat As Char 생성한 표를 문서에 단어로 취급 |
#+----------------------------------------------+
def TreatAsChr(hwpObject):
    # 글자처럼 취급
    HShapeObj = hwpObject.HParameterSet.HShapeObject
    hAction = hwpObject.HAction
    hAction.GetDefault("TablePropertyDialog", HShapeObj.HSet)
    HShapeObj.HorzRelTo = hwpObject.HorzRel("Para")
    HShapeObj.TreatAsChar = 1
    HShapeObj.HSet.SetItem("ShapeType", 3)
    HShapeObj.HSet.SetItem("ShapeCellSize", 0)
    hAction.Execute("TablePropertyDialog", HShapeObj.HSet)

#+-------------------------------------------+
#| Change Table Cell Width 표 셀 사이즈 변경 |
#+-------------------------------------------+
def ChangeTableCellWidth(hwpObject, size):
    # 표 셀 사이즈 변경
    HShapeObj = hwpObject.HParameterSet.HShapeObject
    hAction = hwpObject.HAction
    hAction.HAction.GetDefault("TablePropertyDialog", HShapeObj.HSet)
    HShapeObj.HSet.SetItem("ShapeType", 3)
    HShapeObj.HSet.SetItem("ShapeCellSize", 1)
    HShapeObj.ShapeTableCell.Width = hwpObject.MiliToHwpUnit(size)
    hAction.Execute("TablePropertyDialog", HShapeObj.HSet)

#+--------------------------------------------------------------------------------+
#| Paste Table Cell Value 엑셀에서 데이터를 복사하여 표에 값만 붙이고 싶을때 사용 |
#+--------------------------------------------------------------------------------+
def PasteTableCellValue(hwpObject, opt):
    # opt 5 : 표의 셀 붙이기 내용만 덮어쓰기, 4 : 표의 셀 붙이기 덮어쓰기 , 6 : 표의 셀에서 셀 안에 표로 넣기
    HSelOpt = hwpObject.HParameterSet.HSelectionOpt
    hAction = hwpObject.HAction
    hAction.GetDefault("Paste", HSelOpt.HSet)
    HSelOpt.option = opt
    hAction.Execute("Paste", HSelOpt.HSet)

#+----------------------------------------------------------+
#| Change Table Cell Border No 표의 테두리 보이지 않게 설정 |
#|                                                          |
#| TableCellBorderNo  테두리 보이지 않게 설정               |
#| TableCellBorderRight  오른쪽 테두리 그리기               |
#| TableCellBorderLeft  왼쪽 테두리 그리기                  |
#+----------------------------------------------------------+
def ChangeTableCellBorderNo(hwpObject):
    hAction = hwpObject.HAction
    hAction.Run("TableCellBlock")                 # 표 셀 블럭 (셀 선택 회색 작은 콩 )
    hAction.Run("TableCellBlockExtend")           # 표 셀 블럭 선택 (셀 선택 회색 빨간 콩)
    hAction.Run("TableCellBlockExtend")           # 전체 선택
    hAction.Run("TableCellBorderNo")              # 셀 테두리 보이지 않게 설정
    hAction.Run("Cancel")


#+------------------------------------------------+
#| Cols Table Split Cell 표의 셀에서 칸을 나눈다. |
#+------------------------------------------------+
def ColsTableSplitCell (hwpObject, cols):
    HTableSplCell = hwpObject.HParameterSet.HTableSplitCell
    hAction = hwpObject.HAction
    hAction.GetDefault("TableSplitCell", HTableSplCell.HSet)
    HTableSplCell.Rows = 0
    HTableSplCell.Cols = cols
    hAction.Execute("TableSplitCell", HTableSplCell.HSet)

#+------------------------------------------------+
#| Rows Table Split Cell 표의 셀에서 줄을 나눈다. |
#+------------------------------------------------+
def RowsTableSplitCell (hwpObject, rows):
    HTableSplCell = hwpObject.HParameterSet.HTableSplitCell
    hAction = hwpObject.HAction
    hAction.GetDefault("TableSplitCell", HTableSplCell.HSet)
    HTableSplCell.Rows = rows
    HTableSplCell.Cols = 0
    hAction.Execute("TableSplitCell", HTableSplCell.HSet)

#+-----------------------+
#| 파일을 다른 이름 저장 |
#+-----------------------+
def FileSaveAs(hwpObject, file):
    hwpObject.SaveAs(file, "HWP")

#+-------------------------------------------+
#| 파일 저장                                 |
#|                                           |
#| Clear(3) 문서를 저장하고 닫는다.          |
#| 1: 문서 내용 버림                         |
#| 2: 문서 변경된 경우 저장                  |
#| 3: 무조건 저장                            |
#|                                           |
#| Save() 문서를 저장한다.                   |
#+-------------------------------------------+
def FileSave(hwpObject):
    #hwpObject.Clear(3)          # 저장하고 현재 문서 파일 닫기?
    hwpObject.Save()


#+------------------------------------+
#| 파일을 닫기 ( 아래한글을 닫는다. ) |
#+------------------------------------+
def HWPClose(hwpObject):
    hwpObject.Quit()


#+-----------------+---------------------------------+
#| 편집 -> 쪽 여백 |                                 |
#+-----------------+---------------------------------+
#| 1               | 기본 (머리말, 꼬리말 여백포함)  |
#+-----------------+---------------------------------+
#| 2               | 좁게                            |
#+-----------------+---------------------------------+
#| 3               | 좁게1 (머리말, 꼬리말 여백포함) |
#+-----------------+---------------------------------+
#| 4               | 좁게2 (머리말, 꼬리말 여백포함) |
#+-----------------+---------------------------------+
#| 5               | 넓게                            |
#+-----------------+---------------------------------+
#| 6               | 넓게1 (머리말, 꼬리말 여백포함) |
#+-----------------+---------------------------------+
def ChangePageMargins(hwpObject, key):
    if ( str(key).isdigit() == False ):
        raise ValueError("ERROR 두번째 인자값은 정수를 입력 필요")
    if ( len(str(key)) > 1 ):
        raise ValueError("ERROR 두번째 인자값은 1자리 정수를 입력 필요")
    if ( key == 0 or key > 6 ):
        raise ValueError("ERROR 두번째 인자값은 1 ~ 6사이의 정수를 입력 필요")
    hsDef = hwpObject.HParameterSet.HSecDef
    hAction = hwpObject.HAction
    hAction.GetDefault("PageSetup", hsDef.HSet)
    if ( key == 2 ):  # 2 좁게
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(10.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(10.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(10.0)       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(10.0)    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(0.0)        # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(0.0)        # 꼬리말 여백
    elif ( key == 3 ):  # 3 좁게1 (머리말, 꼬리말 여백포함)
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(20.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(20.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(15.0)       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(15.0)    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(10.0)       # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(10.0)       # 꼬리말 여백
    elif ( key == 4 ):  # 4 좁게2 (머리말, 꼬리말 여백포함)
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(10.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(10.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(5.0)        # 위 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(5.0)     # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(10.0)       # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(10.0)       # 꼬리말 여백
    elif ( key == 5 ):  # 5 넓게
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(40.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(40.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(35.0)       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(35.0)    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(0.0)        # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(0.0)        # 꼬리말 여백
    elif ( key == 6 ):  # 6 넓게1 (머리말, 꼬리말 여백포함)
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(40.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(40.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(35.0)       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(35.0)    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(10.0)       # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(10.0)       # 꼬리말 여백
    else:               # 1 기본 (머리말, 꼬리말 여백포함)
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(30.0)      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(30.0)     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(20.0)       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(15.0)    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(15.0)       # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(15.0)       # 꼬리말 여백
    hsDef.HSet.SetItem("ApplyClass", 24)
    hsDef.HSet.SetItem("ApplyTo", 3)                    # 0 : 각 셀        1, 3: 전체 셀           2 : 선택된 셀
    hAction.Execute("PageSetup", hsDef.HSet)

#+----------------------------------------------------------------------------------------------------+
#| 쪽 여백 커스텀 설정                                                                                |
#|                                                                                                    |
#| value_str 을 구분자 공백 ' ' 구분하여 '왼쪽 오른쪽 위쪽 아래쪽 머리말 꼬리말여백'  숫자 6개를 입력 |
#| 입력한 순서대로 여백을 지정                                                                        |
#+----------------------------------------------------------------------------------------------------+
def CustomPageMargins(hwpObject, value_str):
    if ( str(value_str).replace(" ","").isdigit() == False ):
        raise ValueError("ERROR 두번째 인자값은 정수로된 6개 숫자 입력 필요")
    key_list = str(value_str).split(' ')
    if ( len(key_list) == 6 ):
        hsDef = hwpObject.HParameterSet.HSecDef
        hAction = hwpObject.HAction
        hAction.GetDefault("PageSetup", hsDef.HSet)
        hsDef.PageDef.LeftMargin = hwpObject.MiliToHwpUnit(key_list[0])      # 왼쪽 여백     단위 (mm)
        hsDef.PageDef.RightMargin = hwpObject.MiliToHwpUnit(key_list[1])     # 오른쪽 여백
        hsDef.PageDef.TopMargin = hwpObject.MiliToHwpUnit(key_list[2])       # 위쪽 여백
        hsDef.PageDef.BottomMargin = hwpObject.MiliToHwpUnit(key_list[3])    # 아래쪽 여백
        hsDef.PageDef.HeaderLen = hwpObject.MiliToHwpUnit(key_list[4])       # 머리말 여백
        hsDef.PageDef.FooterLen = hwpObject.MiliToHwpUnit(key_list[5])       # 꼬리말 여백
        hsDef.HSet.SetItem("ApplyClass", 24)
        hsDef.HSet.SetItem("ApplyTo", 3)                    # 0 : 각 셀        1, 3: 전체 셀           2 : 선택된 셀
        hAction.Execute("PageSetup", hsDef.HSet)
    else:
        raise ValueError("ERROR 두번째 인자 값에 왼쪽 오른쪽 위쪽 아래쪽 머리말 꼬리말여백 순서대로 6개 숫자를 공백으로 구분하여 입력 했는지 확인 필요")

#+---------------------------------------+
#| 아래한글파일 열기 filepath는 절대경로 |
#+---------------------------------------+
def FileOpen(hwpObject, filePath):
    if ( os.path.isfile(filePath) == False ):
        raise ValueError("아래한글 파일 경로가 올바른지 확인 필요")       # 파일 존재 확인
    hAction = hwpObject.HAction
    hFileOpen = hwpObject.HParameterSet.HFileOpenSave
    hAction.GetDefault("FileOpen", hFileOpen.HSet)
    hFileOpen.OpenFlag = 0
    hFileOpen.filename = filePath                                           # 파일명 (파일 절대 경로 포함)
    hFileOpen.OpenReadOnly = 0
    hFileOpen.Attributes = 0
    hAction.Execute("FileOpen", hFileOpen.HSet)


def gen_certificates(collection, file_in, file_out, visible):

  hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  #아래한글 인스턴스/오브젝트 생성
  hwp.RegisterModule("FilePathCheckDLL", "AutomationModule") #보안모듈 자동승인 -> regedit에서 등록해야 함.

  hwp.Open(file_in,"HWP","forceopen:true")  #파일 열기

  if(visible == True):
    hwp.XHwpWindows.Item(0).Visible = True  #실생중 HWP화면 보이게 함
  else:
    hwp.XHwpWindows.Item(0).Visible = False  #실생중 HWP화면 보이게 함

  hwp.MovePos(3)  #문서 끝으로 이동
  page_list = [item for item in hwp.GetFieldList().split('\x02')]

  print(page_list)
  #if(visible == 'True'):
    #msgbox = ctypes.windll.user32.MessageBoxW(0, "확인누름 -> 한글자동화를 실행", "한글 자동화 수행", 1)

  data = numpy.array(collection)

  if data.ndim == 1:
    for cnt, field in enumerate(page_list):
      hwp.PutFieldText(f'{field}{{{{{0}}}}}', collection[cnt])

  elif data.ndim == 2:
    hwp.Run('SelectAll') #문서 전체 선택
    hwp.Run('CopyPage')
    hwp.MovePos(3)  #문서 끝으로 이동

    num_pages = len(collection)
    counts = list(range(1,num_pages))
    for count in counts:
      hwp.Run('PastePage')
      hwp.MovePos(3)
      time.sleep(0.1) #데모용

    pages = list(range(num_pages))

    for page in pages:
      for cnt, field in enumerate(page_list):
        hwp.PutFieldText(f'{field}{{{{{page}}}}}', collection[page][cnt])
  else:
    print("Collection corrupted!")

  if(file_out.endswith('.pdf')):
    hwp.SaveAs(file_out, 'PDF') #다른이름으로 파일저장
  else:
    hwp.SaveAs(file_out, 'HWP') #다른이름으로 파일저장

  #hwp.Save() #연 파일 그대로 저장

#----- 실행결과 데모용 완료된 문서 페이지 스크롤--------------------------------------------
  if(visible == True):
    wndow = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(wndow, win32con.SW_MAXIMIZE) #화면 확대

    hwp.MovePos(2) #문서의 처음으로 이동
    time.sleep(0.25)

    if data.ndim >= 2:
      for page in pages:
        time.sleep(0.25)
        pyautogui.keyDown("alt")
        pyautogui.press('pagedown')
        pyautogui.keyUp("alt")

    time.sleep(1) #Add buffer time

  #hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 현재문서 닫기
  os.system("taskkill /f /im hwp.exe")

#----------------------------------------------------------------
# Register your custom-functions
#----------------------------------------------------------------
def register_functions():

    #Add your function name here
    vars()['gen_certificates'] = gen_certificates

    return

#################################################################
# Activate custom-functions. Do not changes!!!
#################################################################

def custom_func(func, *args):

    if(args[0] != None):
        result = globals()[func](*args)
    else:
        result = globals()[func]()

    return result
#################################################################
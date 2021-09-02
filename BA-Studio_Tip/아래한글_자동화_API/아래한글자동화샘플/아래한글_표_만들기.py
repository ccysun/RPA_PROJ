"""
아래의 148mm는 종이여백 210mm에서 60mm(좌우 각 30mm)를 뺀 150mm에다가,
표 바깥여백 각 1mm를 뺀 148mm이다. (TableProperties.Width = 41954)
각 열의 너비는 5개 기준으로 26mm인데 이는 셀마다 안쪽여백 좌우 각각 1.8mm를 뺀 값으로,
148 - (1.8 x 10 =) 18mm = 130mm
그래서 셀 너비의 총 합은 130이 되어야 한다.
아래의 라인28~32까지 셀너비의 합은 16+36+46+16+16=130
표를 생성하는 시점에는 표 안팎의 여백을 없애거나 수정할 수 없으므로
이는 고정된 값으로 간주해야 한다.
"""
#%%
import win32com.client as win32  # COM 임포트
 
#%%
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 아래아한글 오브젝트 생성
hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제
 
#%%
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)  # 표 생성 시작
hwp.HParameterSet.HTableCreation.Rows = 5  # 행 갯수
hwp.HParameterSet.HTableCreation.Cols = 5  # 열 갯수
hwp.HParameterSet.HTableCreation.WidthType = 2  # 너비 지정(0:단에맞춤, 1:문단에맞춤, 2:임의값)
hwp.HParameterSet.HTableCreation.HeightType = 1  # 높이 지정(0:자동, 1:임의값)
hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(148.0)  # 표 너비
hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(150)  # 표 높이
hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)  # 열 5개 생성
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(16.0))  # 1열
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(36.0))  # 2열
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(46.0))  # 3열
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, hwp.MiliToHwpUnit(16.0))  # 4열
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(4, hwp.MiliToHwpUnit(16.0))  # 5열
hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)  # 행 5개 생성
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(40.0))  # 1행
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(20.0))  # 2행
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2, hwp.MiliToHwpUnit(50.0))  # 3행
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(3, hwp.MiliToHwpUnit(20.0))  # 4행
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(4, hwp.MiliToHwpUnit(20.0))  # 5행
hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(148)  # 표 너비
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)  # 위 코드 실행
 

# hwp.MovePos(3)  # 문서 끝으로 이동



# 저장
hwp.SaveAs(file,"HWP")


# 그리고 저장을 해주고 한글프로그램을 닫습니다. 끝~
hwp.Clear(3)
hwp.Quit()


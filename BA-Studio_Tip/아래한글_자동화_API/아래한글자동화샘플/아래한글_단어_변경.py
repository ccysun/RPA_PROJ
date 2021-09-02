
############################################################### 
# 먼저 파이썬 파일내에 한글을 쓰기 위한 구문을 추가합니다.
#_*_coding:cp949_*_
# 그리고 아래 문구를 한글 액티브액스(open API)를 불러오기 위하여 써줍니다.
import win32com.client as win32
hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)



# 그리고 111.hwp 한글파일을 열어줍니다.
hwp.Open('e:\\한글\\111.hwp',"HWP","forceopen:true")




# 그리고 모두 찾아 바꾸기를 하기 위해 명령은  "AllReplace"를 써줍니다.
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);

# 찾아 바꾸기에서 세부내용을 셋업해줘야 하는데요
# "FindString" 속성에 찾을 문자열을 "ReplaceString" 속성에 새 문자열을 넣어줍니다.
option=hwp.HParameterSet.HFindReplace
option.FindString = "박지성"
option.ReplaceString = "손흥민"

# 한글 파일에서  Ctrl+F 를 치면 나오는 대화상자에 아래와 같이 해주는 동작과 같습니다.
# 그리고 명령실행 시 "진짜 바꾸겠습니까?" 라든지 "총 12번 바꿨습니다" 라는 메시지 상자가 나오는데
# 이를 무시하기 위하여 아래 명령을 써줍니다.
option.IgnoreMessage = 1

# 그리고 실행해줍니다. 실행되면서 박지성이 손흥민으로 바뀝니다!!
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);






# 그리고 저장을 해주고 한글프로그램을 닫습니다. 끝~
hwp.Clear(3)
hwp.Quit()



# 단어 찾기
hwp.HAction.GetDefault("ForwardFind", hwp.HParameterSet.HFindReplace.HSet);
option=hwp.HParameterSet.HFindReplace
option.FindString = "AAA"
option.IgnoreMessage = 1
hwp.HAction.Execute("ForwardFind", hwp.HParameterSet.HFindReplace.HSet);


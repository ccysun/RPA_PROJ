import sys, os
import win32com.client as win32
import time
import pyautogui
import win32gui, win32con
  
def gen_certificates(collection,file):

  print('collection =')
  print(collection)
  
  
  hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  #아래한글 인스턴스/오브젝트 생성
  hwp.RegisterModule("FilePathCheckDLL", "AutomationModule") #보안모듈 자동승인 -> regedit에서 등록해야 함.


  hwp.Open(file,"HWP","forceopen:true")  #파일 열기    
  hwp.XHwpWindows.Item(0).Visible = True  #실생중 HWP화면 보이게 함

  hwp.MovePos(3)  #문서 끝으로 이동
  page_list = [item for item in hwp.GetFieldList().split('\x02')]
  
  num_pages = len(collection)
  counts = list(range(1,num_pages))
  print('count = ',counts)
  
  hwp.Run('SelectAll') #문서 전체 선택
  hwp.Run('CopyPage')
  hwp.MovePos(3)  #문서 끝으로 이동
  for count in counts:
    hwp.Run('PastePage')
    hwp.MovePos(3)
    time.sleep(0.25) #데모용
    
  hwp.MovePos(3)   #데모용
  time.sleep(0.25) #데모용
  pages = list(range(num_pages))
  
  
  for page in pages:
    for cnt, field in enumerate(page_list):
      hwp.PutFieldText(f'{field}{{{{{page}}}}}', collection[page][cnt])

  
  #hwp.SaveAs(file, 'HWP') #다른이름으로 파일저장
  hwp.Save() #파일 저장
  
#----- 실행결과 데모용 완료된 문서 페이지 스크롤--------------------------------------------
  wndow = win32gui.GetForegroundWindow()
  win32gui.ShowWindow(wndow, win32con.SW_MAXIMIZE) #화면 확대
  
  hwp.MovePos(2) #문서의 처음으로 이동
  time.sleep(0.5)
  
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
    #vars()['gen_certificates'] = gen_certificates
    
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









































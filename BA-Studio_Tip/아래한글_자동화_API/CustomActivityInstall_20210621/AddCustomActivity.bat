@ echo off
REM 기존 내용 주석
REM xcopy /y ".\CustomFiles\commands.json" "C:\Program Files (x86)\BATEM_v212\BA-Studio\Resources\commands.json"
REM xcopy /y ".\CustomFiles\UserModule.py" "C:\Program Files (x86)\BATEM_v212\Common\Resource\engine\rpa\modules\UserModule.py"
REM xcopy /y ".\CustomFiles\COMMON.py"   "C:\Program Files (x86)\BATEM_v212\Common\Resource\engine\rpa\modules\COMMON.py"
REM xcopy /y ".\CustomFiles\FilePathCheckerModuleExample.dll"  "C:\Program Files (x86)\HNC\HwpAutomation\Modules"

REM BA-Studio 설치된 경로를 batem_path 변수에 입력 
set batem_path=C:\Program Files (x86)\BATEM_v215s
set hnc_path=C:\Program Files (x86)\HNC
REM 아래한글은 위 경로에 설치되어 있어야 한다. 


REM 파일이 존재하면 파일명 앞에 bck_  붙여 파일명을 변경
if exist "%batem_path%\BA-Studio\Resources\commands.json" move "%batem_path%\BA-Studio\Resources\commands.json" "%batem_path%\BA-Studio\Resources\bck_commands.json"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\UserModule.py" move "%batem_path%\Common\Resource\engine\rpa\modules\UserModule.py" "%batem_path%\Common\Resource\engine\rpa\modules\bck_UserModule.py"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.pyc" move "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.pyc" "%batem_path%\Common\Resource\engine\rpa\modules\bck_COMMON.pyc"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.py"  move "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.py" "%batem_path%\Common\Resource\engine\rpa\modules\bck_COMMON.py"

REM 파일 복사
xcopy /Y ".\CustomFiles\commands.json"  "%batem_path%\BA-Studio\Resources\"
xcopy /Y ".\CustomFiles\UserModule.py"    "%batem_path%\Common\Resource\engine\rpa\modules\"
xcopy /Y ".\CustomFiles\COMMON.py"       "%batem_path%\Common\Resource\engine\rpa\modules\"

REM 폴더가 생성되지 않았으면 생성 후 파일 복사
if not exist "%hnc_path%\HwpAutomation\Modules\" mkdir "%hnc_path%\HwpAutomation\Modules\"
if not exist "%hnc_path%\HwpAutomation\Modules\FilePathCheckerModuleExample.dll" copy ".\CustomFiles\FilePathCheckerModuleExample.dll" "%hnc_path%\HwpAutomation\Modules\"

REM 레지스트리 적용
HwpAutomation.reg /Y



@ echo off
REM ���� ���� �ּ�
REM xcopy /y ".\CustomFiles\commands.json" "C:\Program Files (x86)\BATEM_v212\BA-Studio\Resources\commands.json"
REM xcopy /y ".\CustomFiles\UserModule.py" "C:\Program Files (x86)\BATEM_v212\Common\Resource\engine\rpa\modules\UserModule.py"
REM xcopy /y ".\CustomFiles\COMMON.py"   "C:\Program Files (x86)\BATEM_v212\Common\Resource\engine\rpa\modules\COMMON.py"
REM xcopy /y ".\CustomFiles\FilePathCheckerModuleExample.dll"  "C:\Program Files (x86)\HNC\HwpAutomation\Modules"

REM BA-Studio ��ġ�� ��θ� batem_path ������ �Է� 
set batem_path=C:\Program Files (x86)\BATEM_v212
set hnc_path=C:\Program Files (x86)\HNC
REM �Ʒ��ѱ��� �� ��ο� ��ġ�Ǿ� �־�� �Ѵ�. 


REM ������ �����ϸ� ���ϸ� �տ� bck_  �ٿ� ���ϸ��� ����
if exist "%batem_path%\BA-Studio\Resources\commands.json" move "%batem_path%\BA-Studio\Resources\commands.json" "%batem_path%\BA-Studio\Resources\bck_commands.json"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\UserModule.py" move "%batem_path%\Common\Resource\engine\rpa\modules\UserModule.py" "%batem_path%\Common\Resource\engine\rpa\modules\bck_UserModule.py"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.pyc" move "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.pyc" "%batem_path%\Common\Resource\engine\rpa\modules\bck_COMMON.pyc"
if exist "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.py"  move "%batem_path%\Common\Resource\engine\rpa\modules\COMMON.py" "%batem_path%\Common\Resource\engine\rpa\modules\bck_COMMON.py"

REM ���� ����
xcopy /Y ".\CustomFiles\commands.json"  "%batem_path%\BA-Studio\Resources\"
xcopy /Y ".\CustomFiles\UserModule.py"    "%batem_path%\Common\Resource\engine\rpa\modules\"
xcopy /Y ".\CustomFiles\COMMON.py"       "%batem_path%\Common\Resource\engine\rpa\modules\"

REM ������ �������� �ʾ����� ���� �� ���� ����
if not exist "%hnc_path%\HwpAutomation\Modules\" mkdir "%hnc_path%\HwpAutomation\Modules\"
if not exist "%hnc_path%\HwpAutomation\Modules\FilePathCheckerModuleExample.dll" copy ".\CustomFiles\FilePathCheckerModuleExample.dll" "%hnc_path%\HwpAutomation\Modules\"

REM ������Ʈ�� ����
HwpAutomation.reg /Y



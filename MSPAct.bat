@echo off&color 70&mode con: cols=92 lines=30&title  Microsoft Product Activator (CMD Version) (v1.1) by #dp23>nul
setlocal DisableDelayedExpansion
set "batchPath=%~0"
for %%k in (%0) do set batchName=%%~nk
set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
setlocal EnableDelayedExpansion
:checkPrivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )
:getPrivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)
echo Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
echo args = "ELEV " >> "%vbsGetPrivileges%"
echo For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
echo args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
echo Next >> "%vbsGetPrivileges%"
echo UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
"%SystemRoot%\System32\WScript.exe" "%vbsGetPrivileges%" %*
exit /B
:gotPrivileges
setlocal & pushd .
cd /d %~dp0
if '%1'=='ELEV' (del "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)
:Home
cls
echo               ________________________________________________________________
echo              ^|                                                                ^|
echo              ^|      ___________________________________________________       ^|
echo              ^|                                                                ^|
echo              ^|      [1] Start MSPAct (CMD Version)                            ^|
echo              ^|                                                                ^|
echo              ^|      [2] Backup Windows and Office licenses                    ^|
echo              ^|                                                                ^|
echo              ^|      [3] Restore licenses                                      ^|
echo              ^|                                                                ^|
echo              ^|      [4] License management                                    ^|
echo              ^|      ___________________________________________________       ^|
echo              ^|                                                                ^|
echo              ^|________________________________________________________________^|
echo:
choice /c:1234 /n /m ">_                               Your choice [1, 2, 3, 4] : "                                 
if errorlevel 4 goto LicMgmt
if errorlevel 3 goto CheckRestore                           
if errorlevel 2 goto Backup
if errorlevel 1 goto MSAct
:MSAct
cls
echo:
echo                      MSPAct (Microsoft Product Activator) (CMD Version)                    
echo ____________________________________________________________________________________________
echo:
echo                                          [R] Return                                           
set /p key=" >_ Enter your product key: "
if %key% EQU R goto Home
if %key% EQU r goto Home
call :strLen strlen key
if %strlen% EQU 29 (goto InstallProductKey) else (echo:&echo ^> Please enter a valid product key&goto MSAct)
:InstallProductKey
echo:
echo ^>^> Installing product key: %key%...
set ipk=SoftwareLicensingService
wmic path %ipk% where (Version is not null) call InstallProductKey ProductKey='%key%'>nul 2>&1
call cmd /c exit /b %errorlevel%
set ecode=0x%=exitcode%
set act=SoftwareLicensingProduct
if %errorlevel% EQU 0 (goto Activate)
set ipk=OfficeSoftwareProtectionService
wmic path %ipk% where (Version is not null) call InstallProductKey ProductKey='%key%'>nul 2>&1
set act=OfficeSoftwareProtectionProduct
if %errorlevel% EQU 0 (goto Activate)
echo:
echo ^> Error code: %ecode%. Press any key to try again...
pause>nul
goto MSAct
:Activate
setlocal enabledelayedexpansion
set ptkey=%key:~-5%
echo:
echo ^> Success!
echo:
for /f "tokens=2 delims==" %%a in ('"wmic path %act% where (PartialProductKey='%ptkey%') get Name /value"') do (echo ^>^> Activating: %%a)
wmic path %act% where (PartialProductKey='%ptkey%') call Activate>nul 2>&1
for /f "tokens=2 delims==" %%m in ('"wmic path %act% where (PartialProductKey='%ptkey%') get LicenseStatus /value"') do set "licst=%%m">nul
if %licst% EQU 1 (echo:&echo ^> Success!&goto RefreshLicenseStatus)
call cmd /c exit /b %errorlevel%
set ecode=0x%=exitcode%
echo:
echo ^> Error code: %ecode%
if %ecode% == 0xC004C008 (goto ExportIID)
if %ecode% == 0xC004C020 (goto ExportIID)
goto Notification
:ExportIID
echo:
for /f "tokens=2 delims==" %%a in ('"wmic path %act% where (PartialProductKey='%ptkey%') get OfflineInstallationId /value"') do set "IID=%%a">nul
echo ^> Installation ID: %IID%
goto PromptDepositOfflineConfirmationId
:Notification
echo ____________________________________________________________________________________________
echo:
echo                         Thank you for using MSPAct. Have a nice day!                           
echo                                     [R] Return  [E] Exit                                       
choice /c:RE /n /m ">_ Your choice (R/E): "
if errorlevel 2 exit
if errorlevel 1 goto Home
::==============================================================================================
::  String length function:
::  https://stackoverflow.com/questions/5837418/how-do-you-get-the-string-length-in-a-batch-file
::==============================================================================================
:strlen <resultVar> <stringVar>
(   
    setlocal EnableDelayedExpansion
    (set^ tmp=!%~2!)
    if defined tmp (
        set "len=1"
        for %%P in (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) do (
            if "!tmp:~%%P,1!" NEQ "" ( 
                set /a "len+=%%P"
                set "tmp=!tmp:~%%P!"
            )
        )
    ) ELSE (
        set len=0
    )
)
( 
    endlocal
    set "%~1=%len%"
	exit /b
)
:DepositOfflineConfirmationId
echo:
echo ^>^> Depositing Confirmation ID...
wmic path %act% where (PartialProductKey='%ptkey%') call DepositOfflineConfirmationId InstallationID='%IID%' ConfirmationId='%CID%'>nul 2>&1
if %errorlevel% EQU 0 (echo:&echo ^> Success!&goto RefreshLicenseStatus)
echo:
echo ^> There was an error depositing the Confirmation ID
echo:
choice /c:YN /n /m ">_ Do you want to try again? (Y/N): "  
if errorlevel == 2 goto MSAct
goto PromptDepositOfflineConfirmationId
:PromptDepositOfflineConfirmationId
echo:
set /p ucid=" >_ Enter your Confirmation ID: "
if %ucid% EQU R goto MSAct
if %ucid% EQU r goto MSAct
for /f "tokens=1-20 delims=abcdefghijklmnopqrstuvwxyz!@#$&*()-= " %%a in ("%ucid%") do (set pcid=%%a%%b%%c%%d%%e%%f%%g%%h%%i%%j%%k%%l%%m%%n%%o)
call :strLen strlen pcid
if %strlen% EQU 48 (set CID=%ucid%&goto DepositOfflineConfirmationId) else (echo:&echo ^> Please enter a valid Confirmation ID!&goto PromptDepositOfflineConfirmationId)
:RefreshLicenseStatus
if %ipk% EQU OfficeSoftwareProtectionService (goto Notification)
echo:
echo ^>^> Refreshing license status...
wmic path %ipk% where (Version is not null) call RefreshLicenseStatus>nul 2>&1
goto Notification
:Backup
cls
echo:
echo ^>^> Copying license directory... 
for /f "tokens=4" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion"  /V ProductName  ^|findstr /ri "REG_SZ"') do set WinX=%%a
if %WinX% EQU 7 (set cpWin=%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform) else (if not exist "%windir%\System32\spp\store\2.0" (set cpWin=%windir%\System32\spp\store) else (set cpWin=%windir%\System32\spp\store\2.0))
mkdir MSPBackup\Backup1 >nul 2>&1
xcopy /e /y %cpWin% MSPBackup\Backup1>nul
echo:
echo ^> Done!
if exist "%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform" (
echo:
set cpOffice=%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform
echo ^>^> Copying Office license folder...
mkdir MSPBackup\Backup2 >nul 2>&1
xcopy /e /y !cpOffice! MSPBackup\Backup2>nul
) 
echo:
echo ^> Success!
goto Notification
:CheckRestore
cls
for /F "tokens=3 delims=: " %%h in ('sc query "sppsvc" ^| findstr "        STATE"') do (
  if /I "%%h" EQU "RUNNING" (
   echo:
   echo ^>^> Stopping Software Protection Service...
   net stop sppsvc>nul 2>&1
   for /F "tokens=3 delims=: " %%M in ('sc query "sppsvc" ^| findstr "        STATE"') do (if /I "%%M" EQU "RUNNING" (echo:&echo ^> Failed!&echo:&choice /c:YN /n /m ">_ Do you want to try again? (Y/N): "&if errorlevel == 2 (goto Home) else (goto CheckRestore)))
   echo: & echo ^> Done
  )
)
if exist "%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform" (
for /F "tokens=3 delims=: " %%h in ('sc query "osppsvc" ^| findstr "        STATE"') do (
  if /I "%%h" EQU "RUNNING" (
   echo:
   echo ^>^> Stopping Office Software Protection Service...
   net stop osppsvc>nul 2>&1
   for /F "tokens=3 delims=: " %%M in ('sc query "osppsvc" ^| findstr "        STATE"') do (if /I "%%M" EQU "RUNNING" (echo:&echo ^> Failed!&echo:&choice /c:YN /n /m ">_ Do you want to try again? (Y/N): "&if errorlevel == 2 (goto Home) else (goto CheckRestore)))
   echo: & echo ^> Done
  )
))
if exist MSPBackup (set bRestore=%~dp0MSPBackup&goto Restore) else (goto FolderD)
:FolderD
set "psCommand="(new-object -COM 'Shell.Application')^
.BrowseForFolder(0,'Please choose a backup folder.',0,0).self.path""
for /f "usebackq delims=" %%i in (`powershell %psCommand%`) do set "bRestore=%%i"
call :strLen strlen bRestore
if %strlen% EQU 0 (goto Home)
goto Restore
:Restore
setlocal enabledelayedexpansion
if exist "%bRestore%\Backup1" (
echo:
echo ^>^> Restoring license... 
for /f "tokens=4" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion"  /V ProductName  ^|findstr /ri "REG_SZ"') do set WinX=%%a
if !WinX! EQU 7 (set cpWin=%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform) else (if not exist "%windir%\System32\spp\store\2.0" (set cpWin=%windir%\System32\spp\store) else (set cpWin=%windir%\System32\spp\store\2.0))
xcopy /e /y !bRestore!\Backup1 !cpWin!>nul
echo: & echo ^> Success
)
if exist "%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform" (
echo:
echo ^>^> Restoring Office license...
set cpOffice=%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform
xcopy /e /y !bRestore!\Backup2 !cpOffice!>nul
echo: & echo ^> Success
)
goto Notification
:LicMgmt
cls
setlocal enabledelayedexpansion
for /f "tokens=2 delims==" %%b in ('"wmic path SoftwareLicensingProduct where (PartialProductKey is not null) get PartialProductKey /value" 2^>nul') do (
  for /f "tokens=2 delims==" %%c in ('"wmic path SoftwareLicensingProduct where (PartialProductKey='%%b') get Name /value" 2^>nul') do (echo:&echo Name: %%c)
  echo  ^> PartialProductKey: %%b
  for /f "tokens=2 delims==" %%d in ('"wmic path SoftwareLicensingProduct where (PartialProductKey='%%b') get LicenseStatus /value" 2^>nul') do (
    set /a LicStatus=%%d
    if "!LicStatus!" == "0" (echo  ^> LicenseStatus: Unlicensed)
    if "!LicStatus!" == "1" (echo  ^> LicenseStatus: Licensed)
    if "!LicStatus!" == "2" (echo  ^> LicenseStatus: OOBGrace)
    if "!LicStatus!" == "3" (echo  ^> LicenseStatus: OOTGrace)
    if "!LicStatus!" == "4" (echo  ^> LicenseStatus: NonGenuineGrace)
    if "!LicStatus!" == "5" (echo  ^> LicenseStatus: Notification)
    if "!LicStatus!" == "6" (echo  ^> LicenseStatus: ExtendedGrace)
  )
)
setlocal enabledelayedexpansion
for /f "tokens=2 delims==" %%b in ('"wmic path OfficeSoftwareProtectionProduct where (PartialProductKey is not null) get PartialProductKey /value" 2^>nul') do (
  for /f "tokens=2 delims==" %%c in ('"wmic path OfficeSoftwareProtectionProduct where (PartialProductKey='%%b') get Name /value" 2^>nul') do (echo:&echo Name: %%c)
  echo  ^> PartialProductKey: %%b
  for /f "tokens=2 delims==" %%d in ('"wmic path OfficeSoftwareProtectionProduct where (PartialProductKey='%%b') get LicenseStatus /value" 2^>nul') do (
    set /a LicStatus=%%d
    if "!LicStatus!" == "0" (echo  ^> LicenseStatus: Unlicensed)
    if "!LicStatus!" == "1" (echo  ^> LicenseStatus: Licensed)
    if "!LicStatus!" == "2" (echo  ^> LicenseStatus: OOBGrace)
    if "!LicStatus!" == "3" (echo  ^> LicenseStatus: OOTGrace)
    if "!LicStatus!" == "4" (echo  ^> LicenseStatus: NonGenuineGrace)
    if "!LicStatus!" == "5" (echo  ^> LicenseStatus: Notification)
    if "!LicStatus!" == "6" (echo  ^> LicenseStatus: ExtendedGrace)
  )
)
echo ____________________________________________________________________________________________               
echo                      [1] Uninstall a product     [2] Uninstall KMS keys
echo                      [3] Uninstall grace keys    [4] Return
echo:
choice /c:1234 /n /m ">_ Your choice (1/2/3/4): "
if errorlevel 4 goto Home
if errorlevel 3 goto UninstallGraceKeys
if errorlevel 2 goto UninstallKMSKeys
if errorlevel 1 goto PromptUninstallProductKey
:PromptUninstallProductKey
echo:
set /p keyrm=" >_ Enter the last 5 characters of the key you want to remove: "
echo:&echo ^>^> Uninstalling product key: %keyrm%
set /a kmsrm=0
set /a grcrm=0
goto UninstallProductKey
:UninstallKMSKeys
for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (PartialProductKey is not null and Name like '%%KMS%%') get PartialProductKey /value" 2^>nul') do (
if [%%a] NEQ [] (
echo:&echo ^>^> Uninstalling product key: %%a
set /a kmsrm=1
set keyrm=%%a
goto UninstallProductKey
))
echo:&echo ^> Did not find any KMS products in the SoftwareLicensingProduct class.
for /f "tokens=2 delims==" %%a in ('"wmic path OfficeSoftwareProtectionProduct where (PartialProductKey is not null and Name like '%%KMS%%') get PartialProductKey /value" 2^>nul') do (
if [%%a] NEQ [] (
echo:&echo ^>^> Uninstalling product key: %%a
set /a kmsrm=1
set keyrm=%%a
goto UninstallProductKey
))
echo:&echo ^> Did not find any KMS products in the OfficeSoftwareProtectionProduct class.
pause&goto LicMgmt
:UninstallGraceKeys
for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (PartialProductKey is not null and Name like '%%Grace%%' or LicenseStatus='2' or LicenseStatus='3' or LicenseStatus='4' or LicenseStatus='6') get PartialProductKey /value" 2^>nul') do (
if [%%a] NEQ [] (
echo:&echo ^>^> Uninstalling product key: %%a
set /a grcrm=1
set keyrm=%%a
goto UninstallProductKey
))
echo:&echo ^> Did not find any Grace products in the SoftwareLicensingProduct class.
for /f "tokens=2 delims==" %%a in ('"wmic path OfficeSoftwareProtectionProduct where (PartialProductKey is not null and (Name like '%%Grace%%' or LicenseStatus='2' or LicenseStatus='3' or LicenseStatus='4' or LicenseStatus='6')) get PartialProductKey /value" 2^>nul') do (
if [%%a] NEQ [] (
echo:&echo ^>^> Uninstalling product key: %%a
set /a grcrm=1
set keyrm=%%a
goto UninstallProductKey
))
echo:&echo ^> Did not find any Grace products in the OfficeSoftwareProtectionProduct class.
pause&goto LicMgmt
:UninstallProductKey
wmic path SoftwareLicensingProduct where (PartialProductKey='%keyrm%') call UninstallProductKey >nul 2>&1
wmic path OfficeSoftwareProtectionProduct where (PartialProductKey='%keyrm%') call UninstallProductKey >nul 2>&1
echo:&echo ^> Success
if "!kmsrm!" == "0" (pause&goto LicMgmt)
if "!kmsrm!" == "1" (goto UninstallKMSKeys)
if "!grcrm!" == "0" (pause&goto LicMgmt)
if "!grcrm!" == "1" (goto UninstallGraceKeys)
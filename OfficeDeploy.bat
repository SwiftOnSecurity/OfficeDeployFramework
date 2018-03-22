REM start file - buffer line
:: OfficeDeployFramework
:: Original work by @SwiftOnSecurity https://github.com/SwiftOnSecurity/OfficeDeployFramework

:: Advanced multi-product, multi-generation Microsoft Office installation orchestration with in-depth options and performance optimization for any scenario
:: Supports MSI editions of Office 2007 through Office 2016. Supports removing existing Click-To-Run products
:: Supports retreiving files interally via UNC and publicly via encrypted packages hosted on HTTPS
:: Works both interactively and via command line arguments for deployment via SCCM
:: Gradually built over the period of two years for a very specific combination of requirements in a real enterprise network with roaming users, both with and without admin.
:: 

:: [Downloads/caches/loads directly] installation files for Office suites, then installs them
:: Operates over HTTPS or UNC
:: Also performs cleanup and maintenance tasks to increase system performance

SETLOCAL ENABLEDELAYEDEXPANSION
@echo on
set "EchoMode=echo"
set "debug="

title Microsoft Office installer %1 %3

set Version=65
set ScriptDate=2018-03-20



:: Sanity check to prevent clearing root of drive
if not defined Temp exit
if not defined LocalAppData exit

if /I "%TEMP%"=="" exit
if /I "%TEMP%"=="\" exit
if /I "%TEMP%"=="C:" exit
if /I "%TEMP%"=="C:\" exit
if /I "%TEMP%"=="C:\Windows" exit
if /I "%TEMP%"=="D:\" exit
if /I "%TEMP%"=="E:\" exit
if /I "%TEMP%"=="F:\" exit

:: Set variables
Set "local=%~dp0"
Set "LocalDir=C:\install\OfficeDeploy"
md "%LocalDir%"
Set "Log=%LocalDir%\log.txt"
Set "LogClean=%LocalDir%\logclean.txt"

echo !date! !time!- LOG START >>%log%

:: POSITION 1 - Installation source type (UNC or HTTPS or Liveload)
	set "Protocol=%1"
	echo Flag - Protocol: %Protocol% >>%log%
:: POSITION 2 - Installation source server (UNC only)
	set "SourceServer=%2"
	echo Flag - SourceServer: %SourceServer% >>%log%
:: POSITION 3 - Deliverable suite (#)
	set "InstallType=%3"
	echo Flag - InstallType: %InstallType% >>%log%
:: POSITION 4 - Password to decrypt packages (none to purposely skip, or leave blank for default)
	if not "%4"=="" (
		if not "%4"=="none" (
			set "PackagePassword=%4"
			echo Flag - Password: Included >>%log%
		)
	)
	if not defined PackagePassword (
		set "PackagePassword=%USERDOMAIN%"
		echo Flag - Password: Set natively >>%log%
	)
:: POSITION 5 - Enable jumping to local script file after choices and caching (jump to engage, or none)
	set "LocalJump=%5"
	echo Flag - Jump: %LocalJump% >>%log%
	
	if "%LocalJump%"=="jump" (
		echo ALERT - JUMPING >>%log%
		goto Jump
	)
:: POSITION 6 - Restrict if Office is already installed (restrict or none)
	set "Restrict=%6"
	echo Flag - Restrict: %Restrict% >>%log%
:: POSITION 7 - Note if script pushed onto machine, reboot at end (reboot or none)
	if "%7"=="reboot" (
		set "reboot=reboot"
		echo Flag - Reboot: %reboot% >>%log%
	)

:: Set 32-bit kludge to change 32-bit installs to use 32-bit binaries in some cases
if not defined ProgramFiles(x86) set "32=32"

:: Check Trend Micro
if exist "C:\Program Files (x86)\Trend Micro" (
	echo Uninstall Trend Micro first. Try password 27altmS$.
	echo Uninstall Trend Micro first. >>%log%
	pause
	exit
	)
	
:: ------------------

:Select-Protocol
:: Choose which transfer and caching method to use for files

:: Show user interface if not passed through command line
if not defined Protocol (
	echo =====Protocol=====
	echo   1 = HTTPS
	echo   2 = UNC
	echo   3 = Liveload
	echo.
		set /p ProtocolNumber=Enter your selection:

	:: Take user entry and translate to protocol string
	if "!ProtocolNumber!"=="1" set Protocol=HTTPS
	if "!ProtocolNumber!"=="2" set Protocol=UNC
	if "!ProtocolNumber!"=="3" set Protocol=Liveload

	)
	
	:: Redo title
	title  Microsoft Office installer %Protocol% %3

	:: Return to start of section if user entry invalid
	if not defined Protocol goto Select-Protocol

	:: Log
	echo Protocol: %Protocol% >>%log%

	:Select-Protocol-End

:: ------------------

:Select-SourceServer
:: Choose which internal server to use, if applicable

:: Skip if protocol is HTTPS
REM if "%Protocol%"=="HTTPS" goto Select-SourceServer-End

:: Show user interface if not passed through command line
if not defined SourceServer (
	echo =====SourceServer=====
	echo   1 = SERVER0
	echo   2 = SERVER1
	echo   3 = SERVER2
	echo   4 = SERVER3
	echo.
		set /p SourceServerNumber=Enter your selection:

	:: Take user entry and translate to protocol string
	if "!SourceServerNumber!"=="1" set "SourceServer=\\server0\MicrosoftOffice\OfficeDeploy"
	if "!SourceServerNumber!"=="2" set "SourceServer=\\server1\MicrosoftOffice\OfficeDeploy"
	if "!SourceServerNumber!"=="3" set "SourceServer=\\server2\MicrosoftOffice\OfficeDeploy"
	if "!SourceServerNumber!"=="4" set "SourceServer=\\server3\MicrosoftOffice\OfficeDeploy"

	)

	:: Return to start of section if user entry invalid
	if not defined SourceServer goto Select-SourceServer-End

	:: Log
	echo SourceServer: %SourceServer% >>%log%

	%debug%
	:Select-SourceServer-End

:: ------------------

:Select-InstallType
:: Choose which software to install

:: Show user interface if not passed through command line
if not defined InstallType (
	echo =====InstallType=====
	echo   1 = Office 2016 Standard x86
	echo   2 = Outlook 2016 x86
	echo   3 = Office 2007 + Outlook 2016 x86
	echo   4 = Office 2013 Pro Plus x86
	echo   5 = Access 2007 + Office 2016 Standard x86
	echo   6 = Hollow install
	echo   7 = Office 2007 + Outlook 2016 + preserve Access 2007
	echo   8 = Retrofit Access 2016
	echo   9 = Office 2016 + Access 2016
	echo.
		set /p InstallType=Enter your selection:
	)

	:: Log
	echo InstallType: %InstallType% >>%log%

:: ------------------

:Select-Restriction
:: Skip install if Restrict parameter is passed and Office already installed

if "%Restrict%"=="restrict" (
	:: Check Office 2010
	if exist "C:\Program Files (x86)\Microsoft Office\Office14\Outlook.exe" (
		:: Catch if script resumed from wrong position
		if not "%Restrict%"=="restrict" goto Select-Protocol
		echo Restricted - Office2010 detected >>%log%
		goto Error-Restricted
	)
	:: Check Office 2013
	if exist "C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe" (
		:: Catch if script resumed from wrong position
		if not "%Restrict%"=="restrict" goto Select-Protocol
		echo Restricted - Office2013 detected >>%log%
		goto Error-Restricted
	)
	:: Check Office 2016
	if exist "C:\Program Files (x86)\Microsoft Office\Office16\Outlook.exe" (
		:: Catch if script resumed from wrong position
		if not "%Restrict%"=="restrict" goto Select-Protocol
		echo Restricted - Office2016 detected >>%log%
		goto Error-Restricted
	)
)

:: ------------------

:UpdateLogic
:: Retrieve core bootstrap files from UNC server if specified and cache locally
robocopy "%SourceServer%" "%LocalDir%" /R:0 /FFT /LEV:1 /XF *.exe && goto UpdateLogic-End
:UpdateLogic-End

:UpdateUtil
robocopy "%SourceServer%\util" "%LocalDir%\util" /R:0 /FFT /LEV:1
:UpdateUtil-End


%debug%
:: ------------------

:: Jumping to local script instead of continuing in network script
echo !date! !time! JUMPING - %LocalDir%\installoffice2016.bat %Protocol% %SourceServer% %InstallType% -x- jump >>%log%

%LocalDir%\installoffice2016.bat %Protocol% %SourceServer% %InstallType% %PackagePassword% jump
exit

:Jump

echo !date! !time!- JUMP SUCCESSFUL >>%log%
title  Microsoft Office installer %1 %3
:: ------------------

:: Avaliable files: A2016, DellBios, DCU, EMET, JunkReporter, O2007PP (Office2007ProPlus), O2013PPSP1x86 (Office2013ProPlusSP1x86), O2016Sx86 (Office2016Standardx86), Ou2016Sx86 (Outlook2016Standardx86), S4B2016B, UO2007, UO2016, VisioViewer2016

:: Office 2016 Standard - subflags
if "%InstallType%"=="1" (
	echo !date! !time!- Office 2016 variables loading >>%log%
	set "Exclude=UpdateOffice2007.7z Office2007ProPlus.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Outlook2016x86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,EXPDFXPS,ProPlus,VISVIEW,ACCESSRT,LYNCENTRY,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=1"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=1"
	set "wipe-c2r=1"
	set "clear-2007exe=1"
	set "clear-ost=1"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=0"
	set "install-Office2016Std=1"
	set "install-Outlook2016=0"
	set "install-Access2016=0"
	set "install-Access2016Runtime=1"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: Outlook 2016 - subflags
if "%InstallType%"=="2" (
	echo !date! !time!- Outlook 2016 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z UpdateOffice2007.7z Office2007ProPlus.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,EXPDFXPS,ProPlus,VISVIEW,ACCESSRT,LYNCENTRY,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=0"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=1"
	set "wipe-c2r=1"
	set "clear-2007exe=0"
	set "clear-ost=1"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=0"
	set "install-Office2016Std=0"
	set "install-Outlook2016=1"
	set "install-Access2016=0"
	set "install-Access2016Runtime=1"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: Office 2007 + Outlook 2016 - subflags
if "%InstallType%"=="3" (
	echo !date! !time!- Office 2007 + Outlook 2016 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=0"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=0"
	set "wipe-c2r=0"
	set "clear-2007exe=0"
	set "clear-ost=1"
	set "install-Office2007=1"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=0"
	set "install-Office2016Std=0"
	set "install-Outlook2016=1"
	set "install-Access2016=0"
	set "install-Access2016Runtime=1"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)
	
:: Office 2013 - subflags
if "%InstallType%"=="4" (
	echo !date! !time!- Office 2013 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z UpdateOffice2007.7z Office2007ProPlus.7z Office2016Standardx86.7z Outlook2016x86.7z UpdateOffice2016.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,EXPDFXPS,ProPlus,VISVIEW,ACCESSRT,LYNCENTRY,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=1"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=1"
	set "wipe-c2r=1"
	set "clear-2007exe=1"
	set "clear-ost=1"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=1"
	set "install-Office2016Std=0"
	set "install-Outlook2016=0"
	set "install-Access2016=0"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: Access 2007 + Office 2016 Standard - subflags
if "%InstallType%"=="5" (
	echo !date! !time!- Access 2007 + Office 2016 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z Outlook2016x86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=1"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=1"
	set "wipe-c2r=1"
	set "clear-2007exe=1"
	set "clear-ost=1"
	set "install-Office2007=0"
	set "install-Access2007=1"
	set "skip-patch2007=1"
	set "install-Office2013PP=0"
	set "install-Office2016Std=1"
	set "install-Outlook2016=0"
	set "install-Access2016=0"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: Hollow install - subflags
if "%InstallType%"=="6" (
	echo !date! !time!- Hollow install variables loading >>%log%
	set "Exclude=Access2016Runtime.7z NvidiaMobileDriver.7z Office2007ProPlus.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z Outlook2016x86.7z UpdateOffice2007.7z UpdateOffice2013.7z UpdateOffice2016.7z "
	set "EDITIONS=/Quiet /Log %LOCALDIR%"
	set "wipe-2007=0"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=0"
	set "wipe-c2r=1"
	set "clear-2007exe=0"
	set "clear-ost=0"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=0"
	set "install-Office2016Std=0"
	set "install-Outlook2016=0"
	set "install-Access2016=0"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)
	
:: Office 2007 + Outlook 2016 (preserve Access 2007) - subflags
if "%InstallType%"=="7" (
	echo !date! !time!- Office 2007 + Outlook 2016 + preserve Access 2007 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,CLIENTSUITES,PIA /Quiet /Log %LOCALDIR%"
	set "wipe-2007=0"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=0"
	set "wipe-c2r=0"
	set "clear-2007exe=0"
	set "clear-ost=1"
	set "install-Office2007=1"
	set "install-Access2007=0"
	set "skip-patch2007=0"
	set "install-Office2013PP=0"
	set "install-Office2016Std=0"
	set "install-Outlook2016=1"
	set "install-Access2016=0"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: Retrofit Access 2016
if "%InstallType%"=="8" (
	echo !date! !time!- Retrofit Access 2016 x86 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=ACCESS,ACCESSRT,AccessRuntime /Quiet /Log %LOCALDIR%"
	set "wipe-2007=0"
	set "wipe-2010=0"
	set "wipe-2013=0"
	set "wipe-2016=0"
	set "wipe-c2r=1"
	set "clear-2007exe=0"
	set "clear-ost=0"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=1"
	set "install-Office2013PP=0"
	set "install-Office2016Std=0"
	set "install-Outlook2016=0"
	set "install-Access2016=1"
	set "skip-aux=1"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=0"
	set "install-Visio2016Viewer=0"
	)

:: Office 2016 + Access 2016
if "%InstallType%"=="9" (
	echo !date! !time!- Office 2016 x86 and Access 2016 x86 variables loading >>%log%
	set "Exclude=Access2016Runtime.7z Office2013ProPlusx86.7z UpdateOffice2013.7z Office2016Standardx86.7z NvidiaMobileDriver.7z"
	set "EDITIONS=STANDARD,BASIC,PRO,AccessRuntime,Proof,HomeAndStudent,Enterprise,ProfessionalHybrid,Personal,Ultimate,CLICK2RUN,SmallBusiness,Groove,Outlook,EXPDFXPS,ProPlus,VISVIEW,ACCESS,ACCESSRT,AccessRuntime /Quiet /Log %LOCALDIR%"
	set "wipe-2007=1"
	set "wipe-2010=1"
	set "wipe-2013=1"
	set "wipe-2016=0"
	set "wipe-c2r=1"
	set "clear-2007exe=1"
	set "clear-ost=0"
	set "install-Office2007=0"
	set "install-Access2007=0"
	set "skip-patch2007=1"
	set "install-Office2013PP=0"
	set "install-Office2016Std=1"
	set "install-Outlook2016=0"
	set "install-Access2016=1"
	set "skip-aux=0"
	set "install-Access2016Runtime=0"
	set "install-SkypeForBusiness2016Basic=1"
	set "install-Visio2016Viewer=1"
	)

:: ------------------

:: Activate Windows if not already completed
start "" "cscript.exe" "c:\windows\system32\slmgr.vbs" /ato

	:: Log
	echo !date! !time!- Activate Windows triggered >>%log%

:: ------------------

:: Prevent sleep or screensaver until reboot
taskkill /im "caffeine.exe" /f
start "" /b "%LocalDir%\util\caffeine.exe" -noicon
	
:: ------------------

:: Log computer information
WMIC path Win32_ComputerSystem >> "%LocalDir%\model.txt"
start /b WMIC path Win32_BIOS Get Name >> "%LocalDir%\bios.txt"
start /b WMIC path Win32_PnPEntity Get Name >> "%LocalDir%\hardware.txt"
start /b WMIC path Win32_Product Get Name >> "%LocalDir%\software.txt"
start /b wmic bios get serialnumber >> "%LocalDir%\model.txt"

type "%LocalDir%\model.txt" | FIND "Dell"
	if !ERRORLEVEL! == 0 set HWOEM=Dell
	
type "%LocalDir%\model.txt" | FIND "E6430"
	if !ERRORLEVEL! == 0 set HWMODEL=E6430
	if !ERRORLEVEL! == 0 set VIDEO=NVMOBILE

type "%LocalDir%\model.txt" | FIND "E6420"
	if !ERRORLEVEL! == 0 set HWMODEL=E6420
	if !ERRORLEVEL! == 0 set VIDEO=NVMOBILE
	
:: ------------------

:CLEAN-BEGIN
if "%skip-aux%"=="1" goto CLEAN-END

:: Re-register MSIEXEC, in case it's corrupt
start "" /b /wait "msiexec.exe" /regserver

:: ------------------

:: ------------------
:: ----CLEANING-----

	:: Log
	echo !date! !time!- Begin cleaning, see %logclean% >>%log%

:: Clear unused profiles to free disk space and MFT
start "" /b /wait "%LocalDir%\util\delprof2.exe" /d:45 /u >>%logclean%

:: ------------------

:: Close Office
FOR %%g IN (outlook,lync,ucmapi,msosync,msouc,msoev,msotd,communicator,searchfilterhost,searchindexer,officeclicktorun,lynchtmlconv,iastoricon,iastordatasvc,iastora,iastorv,ocpubmgr) DO (taskkill /IM %%g.exe /T /F >>%log%)

:: Clear OST
if "%clear-ost%"=="1" (
	del /q /f /s "C:\Users\*.ost" >>%logclean%
	for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Outlook\*.ost" >>%logclean%
)

for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.Outlook" >>%logclean%

:: ------------------

:: Clear temp files in user profile and Windows directory
del /q /f "C:\Windows\*.log" >>%logclean%
del "%TEMP%\*.*" /f /s /q >>%logclean%
rd /S /Q "%TEMP%\" >>%logclean%
md "%TEMP%" >>%logclean%
del "C:\Windows\Temp\*.*" /f /s /q >>%logclean%
rd /S /Q "C:\Windows\Temp\" >>%logclean%
del "C:\$RECYCLE.BIN\*.*" /f /s /q >>%logclean%
rd /S /Q "C:\$RECYCLE.BIN\" >>%logclean%

:: ------------------

:: Remove all user temp files
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Temp\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do rd /S /Q "%%a\AppData\Local\Temp\*.*" >>%logclean%

:: ------------------

:: Clear Office caches
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Roaming\Microsoft\Templates\Normal*" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Office\16.0\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Office\16.0" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\microsoft\forms\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\microsoft\forms" >>%logclean%
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Outlook\16\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Outlook\16" >>%logclean%
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Outlook\gliding\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Outlook\gliding" >>%logclean%
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Outlook\Offline Address Books\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Outlook\Offline Address Books" >>%logclean%
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Outlook\RoamCache\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Outlook\RoamCache" >>%logclean%

for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Outlook\*.dat" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Outlook\*.oab" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Outlook\*.obi" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Outlook\*.xml" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Local\Microsoft\Office\16.0\*.*" >>%logclean%

for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Roaming\Microsoft\Outlook\*.dat" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Roaming\Microsoft\Outlook\*.srs" >>%logclean%
for /d %%a in (C:\Users\*) do del /q /f "%%a\AppData\Roaming\Microsoft\Outlook\*.xml" >>%logclean%

:: ------------------

:: Clear Temporary Internet Files
for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Low\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Low\" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.MSO\*.*" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.IE5\*.*" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.Word\*.*" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\Tempor~1\*.*" >>%logclean%
for /d %%a in (C:\Users\*) do RD /s /q "%%a\AppData\Local\Microsoft\Windows\Tempor~1" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Internet Explorer\Recovery\*.*" >>%logclean%

for /d %%a in (C:\Users\*) do del /f /s /q "%%a\AppData\Local\Microsoft\Windows\INetCache\IE\*.*" >>%logclean%

start "" /b /wait RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8

:: ----------------------

:: Clear Windows 10 upgrade files
del /q /f "C:\$WINDOWS.~BT\*.*" >>%logclean%
rd /S /Q "C:\$WINDOWS.~BT\" >>%logclean%

:: ------------------

:: Clear unneeded files in root of drive
attrib -h -s "C:\WinPEpge.sys" >>%logclean%
del "C:\WinPEpge.sys" /f >>%logclean%
del /f /s /q "C:\$Recycle.Bin\" >>%logclean%
rd /s /q "C:\$Recycle.Bin\" >>%logclean%
del /f /s /q "C:\hotfix\*.*" >>%logclean%
rd /s /q "C:\hotfix\" >>%logclean%
del /f /s /q "C:\intel\*.*" >>%logclean%
rd /s /q "C:\intel\" >>%logclean%
del /f /s /q "C:\logs\*.*" >>%logclean%
rd /s /q "C:\logs\" >>%logclean%

:: ------------------

:: Clear old Volume Shadow Copies
start "" /b /wait "net.exe" start VSS >>%logclean%
start "" /b /wait "vssadmin.exe" delete shadows /for=%SystemDrive% /oldest /quiet >>%logclean%

:: ------------------

:: Clean Windows
del "C:\Windows\Logs\CBS\*.*" /f /s /q >>%logclean%
reg import "%LocalDir%\cleanmgr.reg"
::start "" /b /wait "cleanmgr.exe" /sagerun:1337 /verylowdisk

:CLEAN-END

:: ------------------
:: ----DOWNLOADS-----

	:: Log
	echo !date! !time!- Jumping to Protocol %protocol% >>%log%

:: Jump to protocol we'll be using
goto %Protocol%

:: ------------------

:HTTPS

::Set InstallDir to the same as LocalDir
set "InstallDir=%LocalDir%"

	:: Log
	echo !date! !time!- HTTPS download started >>%log%

::Setup wget
::Preload KeyCDN certificate in system store with BITSADMIN download
:: Set random number
set "Rand=%RANDOM%"
start "" /wait "bitsadmin.exe" /transfer KeyCDN%rand% /PRIORITY FOREGROUND /download "https://filestore.kxcdn.com/" "%TEMP%\tempfile%rand%"

:: 		--UNIVERSAL--
echo !date! !time!- HTTPS Universal download starting >>%log%
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/DellBios.7z
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/MSFT_EMET.7z
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/JunkReporter.7z
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/OneDrive.7z
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/VisioViewer2016.7z

::		--HARDWARE--
::Intel
start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/IntelChipset.7z

::Dell
if "%HWOEM%"=="Dell" (
	echo !date! !time!- HTTPS Dell download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/DellCommandUpdate.7z
)

::E6430
if "%VIDEO%"=="NVMOBILE" (
	echo !date! !time!- HTTPS nVidia mobile download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/NvidiaMobileDriver.7z
)


:: 		--Office 2007 Pro Plus x86-- 
if "%install-Office2007%"=="1" (
	:: Download Office2007 Pro Plus x86
	echo !date! !time!- HTTPS Office 2007 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Office2007ProPlus.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2007.7z
)

:: 		--Office 2013 ProPlus SP1 x86--
if "%install-Office2013PP%"=="1" (
	:: Download Office2013 ProPlus SP1 x86
	echo !date! !time!- HTTPS Office 2013 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Office2013ProPlusx86.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2013.7z
)

:: 		--Access 2016 x86-- 
if "%install-Access2016%"=="1" (
	:: Download Access2016 Runtime
	echo !date! !time!- HTTPS Access 2016 x86 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Access2016.7z
)

:: 		--Access 2016 Runtime x86-- 
if "%install-Access2016Runtime%"=="1" (
	:: Download Access2016 Runtime
	echo !date! !time!- HTTPS Access Runtime x86 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Access2016Runtime.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2016.7z
)

:: 		--Skype for Business 2016 Basic x86--
if "%install-SkypeForBusiness2016Basic%"=="1" (
	:: Download Skype for Business 2016 Basic
	echo !date! !time!- HTTPS S4B download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/SkypeForBusiness2016.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2016.7z
)

:: 		--Office 2016 Std x86--
if "%install-Office2016Std%"=="1" (
	:: Download Office 2016 Standard
	echo !date! !time!- HTTPS Office2016 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Office2016Standardx86.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2016.7z
)

:: 		--Outlook 2016 x86--
if "%install-Outlook2016%"=="1" (
	:: Download Outlook 2016
	echo !date! !time!- HTTPS Outlook2016 download starting >>%log%
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/Outlook2016x86.7z
	start "" /b /wait "%LocalDir%\util\wget!32!.exe" --continue --no-cache --no-if-modified-since -N -P "%LocalDir%\Packages\\" https://filestore.kxcdn.com/UpdateOffice2016.7z
)

goto Protocol-End

:: ------------------

title  Microsoft Office installer %1 %3

:: ------------------

:UNC

	:: Log
	echo !date! !time!- Begin UNC copy >>%log%

::Set InstallDir to the same as LocalDir
set "InstallDir=%LocalDir%"

start "" /b /wait "robocopy" "%SourceServer%\packages" "%LocalDir%\packages\\" /MIR /FFT /Z /R:1 /XF %Exclude%
echo !date! !time!- Robocopy errorlevel: !ERRORLEVEL! >>%log%

if "%VIDEO%"=="NVMOBILE" (
	echo !date! !time!- Copying NVMOBILE >>%log%
	start "" /b /wait "xcopy" "%SourceServer%\packages\NvidiaMobileDriver.7z" "%LocalDir%\packages"
)

goto Protocol-End

:: ------------------

:Liveload

	:: Log
	echo !date! !time!- Liveload - no copy needed >>%log%

::Set InstallDir to the server
set "InstallDir=%SourceServer%"

goto Protocol-End

:: ------------------

:Protocol-End

	:: Log
	echo !date! !time!- End of downloads >>%log%

:: ------------------
:: ----EXTRACTION-----

	:: Log
	echo !date! !time!- Extracting >>%log%

cd "%LOCALDIR%\Packages"

for /F "delims=:" %%A IN ('dir /b %LOCALDIR%\Packages\*.7z') DO (
	start "" /b /wait "%LocalDir%\util\7za.exe" -p%PackagePassword% -y x %%A
	echo !date! !time!- Extraction errorlevel: !ERRORLEVEL! >>%log%
)

cd "%LOCALDIR%

:: ------------------

	:: Log
	echo !date! !time!- Closing programs >>%log%

:: Close Office
FOR %%g IN (winword,outlook,powerpnt,mspub,msaccess,excel,lync,ucmapi,msosync,msouc,msoev,msotd,vpreview,groove,onenote,onenotem,firefox,chrome,communicator,makecab,searchfilterhost,searchindexer,officeclicktorun,lynchtmlconv,iexplore,dropbox,iastoricon,iastordatasvc,iastora,iastorv,sfdcmsol,ocpubmgr,skype) DO (taskkill /IM %%g.exe /T /F >>%log%)

:: ------------------

:CONFIG-BEGIN
if "%skip-aux%"=="1" goto CONFIG-END

:: ------------------

	:: Log
	echo !date! !time!- Disabling telemetry1 >>%log%

:: DisableProgramTelemetry1
schtasks /change /TN "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\ProgramDataUpdater" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\StartupAppTask" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\AitAgent" /DISABLE
schtasks /change /TN "\Microsoft\Windows\WindowsBackup\ConfigNotification" /DISABLE
schtasks /change /TN "\Microsoft\Windows\RemoteAssistance\RemoteAssistanceTask" /DISABLE
wevtutil sl AirSpaceChannel /e:false
wevtutil cl AirSpaceChannel

reg add "HKCU\Software\Policies\Microsoft\Windows\AppCompat" /v DisablePCA /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\Windows\AppCompat" /v DisableUAR /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\Windows\AppCompat" /v AITEnable /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Telemetry" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Compatibility-Assistant" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Compatibility-Troubleshooter" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Inventory" /v Enabled /t REG_DWORD /d 0 /f

:: ------------------

	:: Log
	echo !date! !time!- Resetting Windows Search >>%log%

:: Reset Windows Search
sc config WSearch start= disabled >>%log%
net stop WSearch >>%log%
reg.exe delete "HKLM\Software\Microsoft\Windows Search" /v SetupCompletedSuccessfully /f >>%log%
RD /S /Q "C:\ProgramData\Microsoft\Search" >>%logclean%

:: ------------------

	:: Log
	echo !date! !time!- Stopping Windows Update >>%log%

:: Clean Windows Update
sc config wuauserv start= disabled >>%log%
net stop wuauserv >>%log%
if exist "%windir%\softwaredistribution\download" rmdir /s /q "%windir%\softwaredistribution\download"

:: ------------------

	:: Log
	echo !date! !time!- Stopping C2R >>%log%

:: Stop Click2Run and DesktopCentral
sc config ClickToRunSvc start= disabled >>%log%
net stop ClickToRunSvc >>%log%
schtasks /change /TN "\Microsoft\Office\Office Automatic Updates" /DISABLE >>%log%


:: ------------------

	:: Log
	echo !date! !time!- Removing C2R and Telemetry >>%log%

:: Remove Office C2R and Telemetry
if exist "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" start "" /b /wait "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=SkypeforBusinessEntryRetail.16_en-us_x-none culture=en-us version.16=16.0 DisplayLevel=false
if exist "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" start "" /b /wait "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=SkypeforBusinessEntryRetail.16_en-us_x-none culture=en-us version.16=16.0 DisplayLevel=false

:: Remove Office C2R Extensibility
start "" /b /wait "MsiExec.exe" /X{90160000-008C-0000-0000-0000000FF1CE} /passive /norestart
start "" /b /wait "MsiExec.exe" /X{90160000-008F-0000-1000-0000000FF1CE} /passive /norestart
start "" /b /wait "MsiExec.exe" /X{90160000-008C-0409-0000-0000000FF1CE} /passive /norestart
start "" /b /wait "MsiExec.exe" /X{90160000-00DD-0000-1000-0000000FF1CE} /passive /norestart

:: Remove Telemetery
start "" /b /wait "MsiExec.exe" /X{90160000-0132-0409-1000-0000000FF1CE} /passive /norestart

:: Force-remove
WMIC product where "Name like '%%click-to-run%%'" call uninstall /nointeractive

:: ------------------

	:: Log
	echo !date! !time!- Removing Salesforce >>%log%

:: Uninstall Salesforce for Outlook
taskkill /im sfdcmsol.exe /f >>%log%
FOR %%g IN ({41feb4a2-7bd2-4d2a-a260-8e8c0e78850c},{af65dc73-94a0-4e85-8ac2-dba52cad1091},{2f4f88fa-c802-4bb6-8e12-9e8313625475},{115EDFAD-1AFB-46A6-9252-43FBB8186D5A},{80EBD79F-5DE4-4189-8E0D-415C750283BE},{1214FA70-2308-4C8A-92B2-D658BA181770},{3EAE8150-DECE-4D3E-A650-2FDEB6AC06A5},{508C3727-3C5E-403D-A69D-FD58A4759FD8},{ABFAAF4C-37B3-45C0-A48F-41560AC61B16},{8842998B-BEE6-4442-9AC7-827BB108C9CE},{1861F90F-7187-469B-BC93-3F947F09E089},{5A0271E2-384E-4386-B14A-09C900D39C9B},{3D7432D9-F9E6-4A94-AF65-079743221EC5},{502C11EE-AC93-47C1-8819-36345B2F1911},{3B037825-A72D-4B41-BA9F-BC8EDC9254FA},{9EF6B750-497B-4586-A7DF-BDE2CBADB900},{3873EBC6-BD2F-4564-A4FE-CD52643B5379},{D97A761B-27EA-4665-94F2-4EFCA4427728},{B1E177D9-E3C9-48E0-9518-EB21FF60297C},{6ACA47BD-D211-45CC-9FF4-70996A7D36E6},{15D99A8D-399F-4647-B2A6-29BE98FCBABA},{F33CCB78-FC9C-482C-8F1F-AF6F8D175337},{F2CED60E-2E22-4880-8D21-3AAE1B0DE6CD},{79CA5983-8BAC-4F17-A8E8-1734B40BC979},{2F055533-E701-4240-80FE-77EB4A8BDB40},{3C084453-2142-4090-825F-6933FAD183E3},{507CC839-9CAB-4E89-BEA9-2FDD0C656927},{6070D4F8-D063-49D2-AFB1-55306A31D1B2},{0003EB0E-E867-4A53-95DC-09D0C927E417},{DE58EA68-36EA-4D96-AF41-8394A7F26D23},{507CC839-9CAB-4E89-BEA9-2FDD0C656927},{C40BC86E-8631-4848-8664-EF59EF5C9511},{116E6ADA-13A6-4725-B974-E809513EE233},{3A4BF362-96AB-48DE-B770-B5BC584EDE49},{23013471-C07F-429F-A924-1665D8809D9B},{C5E637C6-5AB6-426F-B638-7DC533AE5C75},{116E6ADA-13A6-4725-B974-E809513EE233},{1C2275A8-369E-4351-9468-8046A273B71F},{919EDB7E-78E9-440D-A8D8-49B2FB254D69},{6D6EE834-0773-404A-9C8E-F5C5F4B73406}) DO (MsiExec.exe /X %%g /passive /norestart >>%log%)
reg delete "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\Salesforce for Outlook Side Panel" /f >>%log%
reg delete "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\SalesforceForOutlook" /f >>%log%

reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\Salesforce for Outlook Side Panel" /f >>%log%
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\AddinSidePanel.AddinModule" /f >>%log%
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\ADXForm" /f >>%log%

:: ------------------

:: Remove problematic add-ins

:: Grammarly
start "" /b /wait "MsiExec.exe" /X{919EDB7E-78E9-440D-A8D8-49B2FB254D69} /passive /norestart

:: MeetingBridge
:: if exist "C:\Program Files (x86)\InstallShield Installation Information\{788468B4-686C-44D9-87B7-E641673375F7}\setup.exe" (
::	start "C:\Program Files (x86)\InstallShield Installation Information\{788468B4-686C-44D9-87B7-E641673375F7}\setup.exe" -runfromtemp -l0x0009 -removeonly
::)

:: Office Live Meeting 2007
start "" /b /wait "MsiExec.exe" /X{389F8A7A-8611-42E8-8169-20D2BAF0C595} /passive /norestart

:: ------------------

	:: Log
	echo !date! !time!- Clearing Outlook addin registry >>%log%

:: Clear Outlook add-in data
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\iTunesAddIn.CalendarHelper" /f >>%log%
reg delete "HKCU\Software\Microsoft\Office\12.0\Outlook\Resiliency\DisabledItems" /f >>%log%

:: ------------------

	:: Log
	echo !date! !time!- Setting registry permissions >>%log%

:: Set permissions
start "" /wait "C:\Windows\System32\regini.exe" "%LocalDir%\regini.txt"
echo C:\Windows\System32\regini.exe "%LocalDir%\regini.txt" >>%log%
start "" /wait "C:\Windows\SysWOW64\regini.exe" "%LocalDir%\regini.txt"
echo C:\Windows\SysWOW64\regini.exe "%LocalDir%\regini.txt" >>%log%

:: ------------------

	:: Log
	echo !date! !time!- Setting CSC to clear >>%log%

:: Clear cached network share files in CSC
REG ADD "HKLM\System\CurrentControlSet\Services\CSC\Parameters" /v FormatDatabase /t REG_DWORD /d 1 /f >>%log%

:: ------------------

:: Associate .VBS to correct handler
assoc .vbs=VBSFile
cscript.exe //H:WScript
REG ADD "HKCR\.vbs" /ve /d VBSfile /f

:: ------------------

	:: Log
	echo !date! !time!- Syncing time >>%log%

:: Sync time regardless of network location
sc config w32time start= auto
net stop w32time
w32tm /config /syncfromflags:manual /manualpeerlist:"time.google.com"
net start w32time
w32tm /resync /nowait

:: ------------------

	:: Log
	echo !date! !time!- Repairing WMI if required >>%log%

:: Test and repair WMI
WMIC timezone >NUL
if not !ERRORLEVEL!==0 (
	echo !date! !time!- Repairing WMI >>%log%
	call util\repair_wmi.bat
	echo !date! !time!- Repairing WMI - Complete >>%log%
)
@echo on

:: ------------------

if "%av-stinger%"=="1" (
	:: McAfee Stinger
	taskkill /IM "stinger32.exe" /T /F
	if not exist "util\stinger32.exe" (
		start "" /b /wait "util\curl.exe" -k -o "util\stinger32.exe" https://downloadcenter.mcafee.com/products/mcafee-avert/stinger/stinger32.exe
		start "" /b /wait "util\stinger32.exe" --GO --SILENT --PROGRAM --NOSUB --NOUNZIP --REPORTPATH="c:\logs" --DELETE
	)
)
:: ------------------

if "%av-tdsskiller%"=="1" (
	:: Kaspersky TDSSKiller
	taskkill /IM "TDSSKiller.exe" /T /F
	if not exist "TRON\TDSSKiller.exe" (
		start "" /b /wait "util\curl.exe" -k -o "util\tdsskiller.exe" https://media.kaspersky.com/utilities/VirusUtilities/EN/tdsskiller.exe
		start "" /b /wait "util\TDSSKiller.exe" -l "%TEMP%\tdsskiller.log" -silent -tdlfs -dcexact -accepteula -accepteulaksn
	)
)

:: ------------------

if "%av-kvrt%"=="1" (
	::Kaspersky KVRT
	taskkill /IM "KVRT.exe" /T /F
	if not exist "util\KVRT.exe" (
		start "" /b /wait "util\curl.exe" -k -o "util\KVRT.exe" https://devbuilds.kaspersky-labs.com/devbuilds/KVRT/latest/full/KVRT.exe
		start "" /b /wait "util\KVRT.exe" -d "%TEMP%" -accepteula -adinsilent -silent -processlevel 2 -dontcryptsupportinfo
	)
)

:: ------------------

if "%wipe-2007%"=="1" (

	:: Detect if Office 2007 is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office12\" (

		:: Wipe Office 2007
		echo !date! !time!- Wiping Office2007 >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub07.vbs" %EDITIONS% /OSE /K
		echo !date! !time!- Wiping Office2007 - Complete !ERRORLEVEL! >>%log%
	)
)

if "%wipe-2010%"=="1" (

	:: Detect if Office 2010 is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.exe" (

		:: Wipe Office 2010
		echo !date! !time!- Wiping Office2010 >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub10.vbs" %EDITIONS% /K
		echo !date! !time!- Wiping Office2010 - Complete !ERRORLEVEL! >>%log%
	)
)

if "%wipe-2013%"=="1" (

	:: Wipe Office 2013
	echo !date! !time!- Wiping Office2013 >>%log%
		start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub_O15msi.vbs" %EDITIONS% /K
	echo !date! !time!- Wiping Office2013 - Complete !ERRORLEVEL! >>%log%
	
)

if "%wipe-2016%"=="1" (

	:: Detect if Office 2016 is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office16\" (

		:: Wipe Office 2016
		echo !date! !time!- Wiping Office2016 >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub_O16msi.vbs" %EDITIONS% /K
		echo !date! !time!- Wiping Office2016 - Complete !ERRORLEVEL! >>%log%
	)
)

if "%wipe-c2r%"=="1" (

	:: Detect if OfficeC2R is installed
	if exist "C:\Program Files (x86)\Microsoft Office\root\" (

		:: Wipe Office C2R
		echo !date! !time!- Wiping OfficeC2R >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrubc2r.vbs"
		echo !date! !time!- Wiping OfficeC2R - Complete !ERRORLEVEL! >>%log%
	)
	
)

:: ------------------

:CONFIG-END

:: ------------------

if "%install-Access2007%"=="1" (
	echo !date! !time!- Set to install Access2007 >>%log%

	:: Detect if Office2007x86 already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office12\MSACCESS.exe" (

		:: Office2007x86 not installed, Install Access 2007 x86 - Trigger
		echo !date! !time!- Installing Access2007x86 >>%log%

		:: Detect if install files are avaliable
		if exist "%InstallDir%\Packages\Office_2007_ProPlus\2007-USS-accessonly.MSP" (

			:: Install Access 2007 x86
			start "" /b /wait "%InstallDir%\Packages\Office_2007_ProPlus\setup.exe" /adminfile "%InstallDir%\Packages\Office_2007_ProPlus\2007-USS-accessonly.MSP"
			echo !date! !time!- Installing Access2007x86 - Complete !ERRORLEVEL! >>%log%

		)

	) ELSE (
	
		:: If Office2007x86 already installed, remove and lock-out Office2007x86, other than Access
		echo !date! !time!- Locking out Office2007NonAccess - Trigger >>%log%
		
		:: Detect if install files are avaliable
		if exist "%InstallDir%\Packages\Office_2007_ProPlus\2007-USS-accessonly.MSP" (
		
			:: Lock out Office2007x86, other than Access
			start "" /b /wait "msiexec.exe" /p "%InstallDir%\Packages\Office_2007_ProPlus\2007-USS-accessonly.MSP" /passive /norestart
		echo !date! !time!- Locking out Outlook2007 - Complete !ERRORLEVEL! >>%log%
		)

	)

)

:: ------------------

if "%install-Office2007%"=="1" (
	echo !date! !time!- Set to install Office2007 >>%log%

	:: Detect if Office2007x86 already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office12\GRAPH.exe" (

		:: Install Office 2007 ProPlus x86
		echo !date! !time!- Installing Office2007 >>%log%
			start "" /b /wait "%InstallDir%\Packages\Office_2007_ProPlus\setup.exe" /adminfile "%InstallDir%\Packages\Office_2007_ProPlus\2007-uss-neo.MSP"
		echo !date! !time!- Installing Office2007 - Complete !ERRORLEVEL! >>%log%
	)
	
	:: Only patch is Office 2007 ProPlus x86 installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office12\GRAPH.exe" (
	echo !date! !time!- Office2007 install files detected >>%log%
	
		:: Skip patching Office 2007 ProPlus x86 if option set
		if not "%skip-patch2007%"=="1" (
			echo !date! !time!- Set to patch Office2007 >>%log%

			:: Patch Office 2007
			echo !date! !time!- Patching Office2007 >>%log%
				FOR /F %%G in ('dir /b "%InstallDir%\Packages\Updates-Office2007\*.exe"') do (
					echo !date! !time!- Installing update: %%G >>%log%
					start "" /b /wait "%InstallDir%\Packages\Updates-Office2007\%%G" /passive /norestart /quiet
				)
				FOR /F %%G in ('dir /b "%InstallDir%\Packages\Updates-Office2007\*.msp"') do (
					echo !date! !time!- Installing update: %%G >>%log%
					start "" /b /wait msiexec /p "%InstallDir%\Packages\Updates-Office2007\%%G" /passive /norestart)
				)
			echo !date! !time!- Patching Office2007 - Complete !ERRORLEVEL! >>%log%
		)
	)
)

:: ------------------

if "%clear-2007exe%"=="1" (
	:: Remove abandoned .EXE in root of "Microsoft Office"
	del "C:\Program Files (x86)\Microsoft Office\*.exe" /f >>%logclean%
	
	:: Remove abandoned 2007 shortcuts from user profiles
	del /q /f /s "C:\Users\*Office Excel 2007.lnk" >>%logclean%
	del /q /f /s "C:\Users\*Office Outlook 2007.lnk" >>%logclean%
	del /q /f /s "C:\Users\*Office Word 2007.lnk" >>%logclean%

	:: Remove abandoned 2007 shortcuts from Public profile
	del /q /f /s "C:\ProgramData\Microsoft\Windows\Start Menu\*Office Excel 2007.lnk" >>%logclean%
	del /q /f /s "C:\ProgramData\Microsoft\Windows\Start Menu\*Office Outlook 2007.lnk" >>%logclean%
	del /q /f /s "C:\ProgramData\Microsoft\Windows\Start Menu\*Office Word 2007.lnk" >>%logclean%
)

:: ------------------

if "%install-Office2013PP%"=="1" (
	echo !date! !time!- Set to install Office2013PP >>%log%

	:: Remove Office 2007 Export to PDF since it won't be needed
	start "" /b /wait "MsiExec.exe" /X{90120000-00B2-0409-0000-0000000FF1CE} /qn /norestart REBOOT=ReallySuppress

	:: Detect if Office 2013 ProPlus SP1 x86 is already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.exe" (

		:: Install Office 2013 ProPlus SP1 
		echo !date! !time!- Installing Office2013PP >>%log%
			start "" /b /wait "%InstallDir%\Packages\Office_2013_ProPlus_SP1_x86\setup.exe" /adminfile "%InstallDir%\Packages\Office_2013_ProPlus_SP1_x86\2013-uss.MSP"
		echo !date! !time!- Installing Office2013PP - Complete !ERRORLEVEL! >>%log%
	)

	:: Patch Office 2013
	echo !date! !time!- Patching Office2013 >>%log%
		FOR /F %%G in ('dir /b "%InstallDir%\Packages\Updates-Office2013\*.msp"') do (start "" /b /wait msiexec /p "%InstallDir%\Packages\Updates-Office2013\%%G" /passive /norestart)
	echo !date! !time!- Patching Office2013 - Complete !ERRORLEVEL! >>%log%
)

:: ------------------

if "%install-SkypeForBusiness2016Basic%"=="1" (
	echo !date! !time!- Set to install S4B >>%log%

	:: Detect if S4B x86 is already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office16\lync.exe" (

		:: Install Skype for Business 2016 Basic x86
		echo !date! !time!- Installing S4B Basic x86 >>%log%
		if exist "%InstallDir%\Packages\SkypeForBusiness2016Basic\setup.exe" (
			start "" /b /wait "%InstallDir%\Packages\SkypeForBusiness2016Basic\setup.exe" /adminfile "%InstallDir%\Packages\SkypeForBusiness2016Basic\2016-uss.MSP"
			echo !date! !time!- Installing S4B Basic x86 - Complete !ERRORLEVEL! >>%log%
		)

	)
	
)

:: ------------------

if "%install-Access2016%"=="1" (
	echo !date! !time!- Set to install Access 2016 x86 >>%log%

	:: Remove and lock-out Access 2007
	echo !date! !time!- Locking out Access2007 >>%log%
	if exist "%InstallDir%\Packages\Access_2016_x86\2007-removeaccess.MSP" (
		start "" /b /wait "msiexec.exe" /p "%InstallDir%\Packages\Access_2016_x86\2007-removeaccess.MSP" /passive /norestart
		echo !date! !time!- Locking out Access2007 - Complete !ERRORLEVEL! >>%log%
	)

	:: Detect if Access 2007 Runtime is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office12\MSACCESS.EXE" (

		:: Wipe Access 2007 Runtime
		echo !date! !time!- Wiping Office2007Runtime >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub07.vbs" Access,AccessRuntime,AccessRT /Force
		echo !date! !time!- Wiping Office2007Runtime - Complete !ERRORLEVEL! >>%log%
	)
	
	:: Uninstall Access database engine 2010
	msiexec.exe /x {90140000-00D1-0409-0000-0000000FF1CE} /passive /norestart

	:: Detect if Access 2010 Runtime is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE" (

		:: Wipe Access 2010 Runtime
		echo !date! !time!- Wiping Office2010Runtime >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub10.vbs" Access,AccessRuntime,AccessRT /Force
		echo !date! !time!- Wiping Office2010Runtime - Complete !ERRORLEVEL! >>%log%
	)

	:: Detect if Access 2013 Runtime is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office15\MSACCESS.EXE" (

		:: Wipe Access 2013 Runtime
		echo !date! !time!- Wiping Office2013Runtime >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub_O15msi.vbs" Access,AccessRuntime,AccessRT /Force
		echo !date! !time!- Wiping Office2013Runtime - Complete !ERRORLEVEL! >>%log%
	)

	:: Detect if Access 2016 Runtime is installed
	if exist "C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE" (

		:: Wipe Access 2016 Runtime
		echo !date! !time!- Wiping Access2016Runtime >>%log%
			start "" /b /wait "cscript.exe" "%LocalDir%\OffScrub_O16msi.vbs" Access,AccessRuntime,AccessRT /Force
		echo !date! !time!- Wiping Access2016Runtime - Complete !ERRORLEVEL! >>%log%
	)

	:: Install Access 2016 x86
	echo !date! !time!- Installing Access 2016 x86 >>%log%
		start "" /b /wait "%InstallDir%\Packages\Access_2016_x86\setup.exe" /adminfile "%InstallDir%\Packages\Access_2016_x86\2016access.MSP"
	echo !date! !time!- Installing Access 2016 x86 - Complete !ERRORLEVEL! >>%log%

)

:: ------------------

if "%install-Access2016Runtime%"=="1" (
	echo !date! !time!- Set to install Access 2016 Runtime x86 >>%log%

	:: Install Access 2016 Runtime x86
	echo !date! !time!- Installing Access 2016 Runtime x86 >>%log%
		start "" /b /wait "%InstallDir%\Packages\Access2016Runtime\setup.exe" /config "%InstallDir%\Packages\Access2016Runtime\config.xml"
	echo !date! !time!- Installing Access Runtime - Complete !ERRORLEVEL! >>%log%

)

:: ------------------

if "%install-Office2016Std%"=="1" (
	echo !date! !time!- Set to install Office2016Stdx86 >>%log%

	:: Remove Office 2007 Export to PDF since it won't be needed
	start "" /b /wait "MsiExec.exe" /X{90120000-00B2-0409-0000-0000000FF1CE} /qn /norestart REBOOT=ReallySuppress

	:: Detect if Office 2016 x86 is already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.exe" (

		:: Install Office 2016 Standard x86
		echo !date! !time!- Installing Office2016Stdx86 >>%log%
			start "" /b /wait "%InstallDir%\Packages\Office_2016_Std_x86\setup.exe" /adminfile "%InstallDir%\Packages\Office_2016_Std_x86\2016office.MSP"
		echo !date! !time!- Installing Office2016Stdx86 - Complete !ERRORLEVEL! >>%log%
	)

)

:: ------------------

if "%install-Outlook2016%"=="1" (
	echo !date! !time!- Set to install Outlook2016x86 >>%log%

	:: Detect if Outlook 2016 already installed
	if not exist "C:\Program Files (x86)\Microsoft Office\Office16\outlook.exe" (

		:: Remove and lock-out Outlook 2007
		echo !date! !time!- Locking out Outlook2007 >>%log%
		if exist "%InstallDir%\Packages\Outlook_2016_Std_x86\2007-removeoutlook.MSP" (
			start "" /b /wait "msiexec.exe" /p "%InstallDir%\Packages\Outlook_2016_Std_x86\2007-removeoutlook.MSP" /passive /norestart
			echo !date! !time!- Locking out Outlook2007 - Complete !ERRORLEVEL! >>%log%
		)

		:: Install Outlook 2016
		echo !date! !time!- Installing Outlook2016x86 >>%log%
		if exist "%InstallDir%\Packages\Outlook_2016_Std_x86\setup.exe" (
			start "" /b /wait "%InstallDir%\Packages\Outlook_2016_Std_x86\setup.exe" /adminfile "%InstallDir%\Packages\Outlook_2016_Std_x86\2016outlook.MSP"
			echo !date! !time!- Installing Outlook2016x86 - Complete !ERRORLEVEL! >>%log%
		)

	)

	:: Add WINWORD.EXE to the Office2016 folder to unlock Proofing tools and Picture formatting
	if not exist "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.exe" (
		xcopy "C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.exe" "C:\Program Files (x86)\Microsoft Office\Office16\" /Y >>%log%
		xcopy "C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.exe" "C:\Program Files (x86)\Microsoft Office\Office16\" /Y >>%log%
		xcopy "C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.exe" "C:\Program Files (x86)\Microsoft Office\Office16\" /Y >>%log%
	)

	:: Remove Outlook 2007 shortcuts
	DEL /F /S /Q "C:\Users\*Outlook 2007*.lnk" >>%logclean%
	DEL /F /S /Q "C:\ProgramData\Microsoft\*Outlook 2007*.lnk" >>%logclean%

)

:: ------------------

if "%install-Visio2016Viewer%"=="1" (
	echo !date! !time!- Set to install VisioViewerx86 >>%log%

	:: Install Visio 2016 Viewer
	echo !date! !time!- Installing VisioViewerx86 >>%log%
	if exist "%InstallDir%\Packages\Visio2016Viewer\visioviewer_4339-1001_x86_en-us.exe" (
		start "" /b /wait "%InstallDir%\Packages\Visio2016Viewer\visioviewer_4339-1001_x86_en-us.exe" /passive /norestart /quiet
		echo !date! !time!- Installing VisioViewerx86 - Complete !ERRORLEVEL! >>%log%
	)

)

:: ------------------

:: Install JunkReportingAddinx86
echo !date! !time!- Installing JunkReporterx86 >>%log%
	if exist "%InstallDir%\Packages\JunkReportingAddinx86\JunkReportingAddinx86.msi" (
		start "" /b /wait "msiexec.exe" /i "%InstallDir%\Packages\JunkReportingAddinx86\JunkReportingAddinx86.msi" /passive /norestart REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL="DisableShutdown" MaxMessageSelection=30 BccEmailAddress="junkreports-COMPANY@COMPANY.com"
		echo !date! !time!- Installing JunkReporterx86 - Complete !ERRORLEVEL! >>%log%
	)

:: ------------------

:: Patch Office 2016
echo !date! !time!- Patching Office2016x86 >>%log%
	FOR /F %%G in ('dir /b "%InstallDir%\Packages\Updates-Office2016\*.msp"') do (
		echo !date! !time!- Installing update: %%G >>%log%
		start "" /b /wait msiexec /p "%InstallDir%\Packages\Updates-Office2016\%%G" /passive /norestart
		echo !date! !time!- Update errorlevel: !ERRORLEVEL! >>%log%
	)
	FOR /F %%G in ('dir /b "%InstallDir%\Packages\Updates-Office2016\*.exe"') do (
		echo !date! !time!- Installing update: %%G >>%log%
		start "" /b /wait "%InstallDir%\Packages\Updates-Office2016\%%G" /passive /norestart
	)
echo !date! !time!- Patching Office2016x86 - Complete >>%log%

:: ------------------

:: Restart explorer.exe after grove updates
	if not "%reboot%"=="reboot" (
		start "" /b "explorer.exe"
	)

:: ------------------

:: Activate software
echo !date! !time!- Activating Office >>%log%

if exist "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.vbs" (
	start "" /b /wait "cscript.exe" "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.vbs" /act
)
if exist "C:\Program Files\Microsoft Office\Office16\OSPP.vbs" (
	start "" /b /wait "cscript.exe" "C:\Program Files\Microsoft Office\Office16\OSPP.vbs" /act
)

:: ------------------

:: Install EMET 5.52
:: Per-app config controlled by Group Policy
echo !date! !time!- Installing EMET >>%log%

if exist "%InstallDir%\Packages\EMET\EMETSetup_5.52.msi" (
	start "" /b /wait msiexec /i "%InstallDir%\Packages\EMET\EMETSetup_5.52.msi" /passive /norestart
	echo !date! !time!- Installing EMET - Complete !ERRORLEVEL! >>%log%
)
if exist "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" (
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --delete_all >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --reporting +telemetry +eventlog +trayicon >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --exploitaction stop >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --system Pinning=Disabled >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --deephooks enabled >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --antidetours enabled >>%log%
	start "" /b /wait "C:\Program Files (x86)\EMET 5.5\EMET_Conf.exe" --bannedfunc enabled >>%log%
)

:: ------------------

:: Install OneDrive machine-wide
echo !date! !time!- Installing OneDrive >>%log%

if exist "%InstallDir%\Packages\OneDrive\OneDriveSetup.exe" (
	start "" /b /wait "%InstallDir%\Packages\OneDrive\OneDriveSetup.exe" /silent /PerComputer
	echo !date! !time!- Installing OneDrive - Complete !ERRORLEVEL! >>%log%
)

:: ------------------

:: DisableProgramTelemetry2
schtasks /change /TN "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\ProgramDataUpdater" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\StartupAppTask" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Application Experience\AitAgent" /DISABLE
schtasks /change /TN "\Microsoft\Windows\WindowsBackup\ConfigNotification" /DISABLE
schtasks /change /TN "\Microsoft\Windows\RemoteAssistance\RemoteAssistanceTask" /DISABLE
wevtutil sl AirSpaceChannel /e:false
wevtutil cl AirSpaceChannel

reg add "HKCU\Software\Policies\Microsoft\Windows\AppCompat" /v DisablePCA /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\Windows\AppCompat" /v DisableUAR /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\Windows\AppCompat" /v AITEnable /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Telemetry" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Compatibility-Assistant" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Compatibility-Troubleshooter" /v Enabled /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-Application-Experience/Program-Inventory" /v Enabled /t REG_DWORD /d 0 /f

:: ------------------

:: DisableNetbios
echo !date! !time!- Disabling NetBIOS >>%log%

powershell.exe -command "& {set-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces\tcpip* -Name NetbiosOptions -Value 2}"

:: ------------------

:: UnInstallCCTK
:: Prepare for updating to new version

:: Check if Dell hardware
if "%HWOEM%"=="Dell" (
	echo UnInstalling Dell Client Configuration Toolkit
	echo !date! !time!- UnInstalling Dell Client Configuration Toolkit >>%log%

	::Remove
		set LPATH="%WINDIR%\CCTK\X86"
		if defined ProgramFiles^(x86^) set LPATH="%WINDIR%\CCTK\X86_64"
		cd !LPATH!
		call "HAPI\HAPIUninstall.bat"
		del /q /f /s "%WINDIR%\CCTK\*.*"
		if "%EchoMode%"=="echo" echo on
)

cd %LOCALDIR%

:: ------------------

:: InstallCCTK
:: Allows control of Dell BIOS settings
echo !date! !time!- Installing CCTK >>%log%

:: Check if Dell hardware
if "%HWOEM%"=="Dell" (
	echo Installing Dell Client Configuration Toolkit

	:: Drop CCTK in Windir
	start "" /b /wait "%InstallDir%\Packages\DellBios\7za.exe" x -o%WINDIR% -y "%InstallDir%\Packages\DellBios\CCTK.zip" >>%log%
		set LPATH="%WINDIR%\CCTK\X86"
		if defined ProgramFiles^(x86^) set LPATH="%WINDIR%\CCTK\X86_64"
		cd !LPATH!
		call "HAPI\HAPIInstall.bat"
		if "%EchoMode%"=="echo" echo on
)

cd %LOCALDIR%

:: ------------------

:: BIOS unlock

:: Check if Dell hardware
if "%HWOEM%"=="Dell" (
	echo !date! !time!- Unlocking BIOS >>%log%
	set LPATH="%WINDIR%\CCTK\X86"
	if defined ProgramFiles^(x86^) set LPATH="%WINDIR%\CCTK\X86_64"
	cd !LPATH!
	echo +Removing BIOS setup password temporarily
	cctk --valsetuppwd=00000 --setuppwd= >>%log%
	if "%EchoMode%"=="echo" echo on
)

:: Return to root installer folder
cd %LOCALDIR%

:: ------------------

:: Restore services and tasks
REM This is done before the driver installation in case it bluescreens.
echo !date! !time!- Restoring services >>%log%

sc config wuauserv start= delayed-auto
sc config bits start= delayed-auto
sc config wsearch start= delayed-auto
sc config ClickToRunSvc start= delayed-auto
sc config "ManageEngine Desktop Central - Agent" start= auto
schtasks /change /TN "\Microsoft\Office\Office Automatic Updates" /ENABLE

:: ------------------

:: Update drivers for Office 2016 hardware acceleration compatibility

:: Intel Chipset
	echo !date! !time!- Installing Intel Chipset >>%log%
	:: Intel-Chipset-Install
	if exist "%InstallDir%\Packages\Intel\SetupChipset.exe" (
		start "" /b /wait "%InstallDir%\Packages\Intel\SetupChipset.exe" -s -norestart
		echo !date! !time!- Installing Intel Chipset - Complete !ERRORLEVEL! >>%log%
	)

:: Check if nVidia mobile
if "%VIDEO%"=="NVMOBILE" (
	echo !date! !time!- Installing nVidia mobile drivers >>%log%
	:: NVMOBILE-Install
	if exist "%InstallDir%\Packages\NVMOBILE\390.65-desktop-win8-win7-64bit-international-whql.exe" (
		start "" /b /wait "%InstallDir%\Packages\NVMOBILE\390.65-desktop-win8-win7-64bit-international-whql.exe" -s -n
		echo !date! !time!- Installing Installing nVidia mobile drivers - Complete !ERRORLEVEL! >>%log%
	)
)

:: Check if Dell hardware
if "%HWOEM%"=="Dell" (
	echo !date! !time!- Installing DCU >>%log%
	:: DCU-Install
	if exist "%InstallDir%\Packages\DellCommandUpdate\Dell-Command-Update_X79N4_WIN_2.3.1_A00.EXE" (
		start "" /b /wait "%InstallDir%\Packages\DellCommandUpdate\Dell-Command-Update_X79N4_WIN_2.3.1_A00.EXE" /s >>%log%
	)

	:: DCU-Update
	echo Running Dell Command Update
	echo !date! !time!- Running DCU update >>%log%
	start "" /b /wait "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe" /log "c:\install"

	if "%EchoMode%"=="echo" echo on
)

:: Return to root installer folder
cd %LOCALDIR%

:: ------------------

:: Brand BIOS and configure proper settings

:: Check if Dell hardware
if "%HWOEM%"=="Dell" (

	:: Older method
	set LPATH="%WINDIR%\CCTK\X86
	if defined ProgramFiles^(x86^) set LPATH="%WINDIR%\CCTK\X86_64"
	cd !LPATH!
	echo !date! !time!- Branding BIOS >>%log%
	echo +Branding BIOS
	cctk --propowntag="If found please call" >nul 2>&1
	cctk --asset=COMPANY >nul 2>&1

	:: Newer method
	start "" /b /wait "%InstallDir%\Packages\DellBios\PlatformTags32W.exe" SOT:"Owned by COMPANY."
	start "" /b /wait "%InstallDir%\Packages\DellBios\PlatformTags32W.exe" SAT:"COMPANY"

	cd !LPATH!

	echo +Enabling CPU eXecute Disable ^(XD^) feature support
	cctk --cpuxdsupport=enable >nul 2>&1
	echo +Enabling virtualization
	cctk --virtualization=enable >nul 2>&1
	echo +Enabling VT for Direct I/O
	cctk --vtfordirectio=on >nul 2>&1
	echo +Enabling Wake On Lan ^(WOL^)
	cctk --wakeonlan=enable >nul 2>&1
	cctk --wakeonlanbootovrd=enable >nul 2>&1
	echo +Setting built-in NIC status to PXE
	cctk --embnic1=on >nul 2>&1
	echo +Enabling USB powershare
	cctk --usbpowershare=enable >nul 2>&1
	echo +Setting fastboot
	cctk --fastboot=automatic >nul 2>&1
	echo +Enabling SMART errors
	cctk --smarterrors=enable >nul 2>&1
	echo +Enabling HD free fall protection
	cctk --hdfreefallprotect=enable >nul 2>&1
	echo +Enabling BIOS update signature verification
	cctk --sfuenabled=yes >nul 2>&1
	cctk --chasintrustatus=tripreset >nul 2>&1
	echo +Configuring CPU performance
	cctk --cpucore=all >nul 2>&1
	cctk --cpuxdsupport=enable >nul 2>&1
	cctk --cstatesctrl=enable >nul 2>&1
	cctk --speedstep=automatic >nul 2>&1
	cctk --turbomode=enable >nul 2>&1
	cctk --postmebxkey=on >nul 2>&1
	cctk --biosautorecovery=enable >nul 2>&1
	cctk --biosdowngrade=disable >nul 2>&1
	echo +Clear BIOS log on next boot
	cctk --bioslogclear=enable >nul 2>&1
	echo +Expose all options in BIOS setup
	cctk --biossetupadvmode=enable >nul 2>&1
	echo +Clear Fault Tolerant Memory Log on next boot
	cctk --faulttolerantmemlogclear=enable
	echo +Keyboard backlight timeout on AC to 1min
	cctk --kbdbacklighttimeoutac=1m
	echo +Clear Power Event Log on next boot
	cctk --powerlogclear=enable
)

:: Return to root installer folder
cd %LOCALDIR%

:: ------------------

:: Enroll in Microsoft Update
echo !date! !time!- Enroll in Microsoft Update >>%log%

start "" /b /wait "cscript.exe" "%LocalDir%\MicrosoftUpdate.vbs"

:: ------------------

echo !date! !time!- Updating Group Policy >>%log%

:: Update computer and user Kerberos ticket
klist -li 0x3e7 purge >>%log%
klist purge >>%log%

:: Restore any removed Group Policy settings
gpupdate /force >>%log%

:: ------------------

echo installoffice2016.bat finish !date! !time!>>%log%
echo installoffice2016.bat finish !date! !time!>>c:\install\office2016.txt

:: ------------------

:: Reboot if flag passed at runtime
	if "%reboot%"=="reboot" (
		start "" /b "shutdown.exe" /f /r /t 30
	)
	
:End
exit

:Error
:Error-Restricted
echo.
echo  Message: Office is already installed.
echo.
echo.
pause
exit
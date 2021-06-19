::===============================================================================================================
::===============================================================================================================
	@echo off
	cls
	setLocal EnableExtensions EnableDelayedExpansion
	set "OfficeRToolpath=%~dp0"
	set "OfficeRToolpath=%OfficeRToolpath:~0,-1%"
	set "OfficeRToolname=%~n0.cmd"
	set "pswindowtitle=$Host.UI.RawUI.WindowTitle = 'Administrator: OfficeRTool - 2019/June/08 -'"
	cd /D "%OfficeRToolpath%"
::===============================================================================================================
:: CHECK ADMIN RIGHTS
::===============================================================================================================
	fltmc >nul 2>&1
	if "%errorlevel%" NEQ "0" (goto:UACPrompt) else (goto:GotAdmin)
::===============================================================================================================
::===============================================================================================================
:UACPrompt
	echo:
	echo   Requesting Administrative Privileges...
	echo   Press YES in UAC Prompt to Continue
	echo:
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\GetAdmin.vbs"
	echo args = "ELEV " >> "%temp%\GetAdmin.vbs"
	echo For Each strArg in WScript.Arguments >> "%temp%\GetAdmin.vbs"
	echo args = args ^& UCase(strArg) ^& " "  >> "%temp%\GetAdmin.vbs"
	echo Next >> "%temp%\GetAdmin.vbs"
    echo UAC.ShellExecute "%OfficeRToolname%", args, "%OfficeRToolpath%", "runas", 1 >> "%temp%\GetAdmin.vbs"
    cmd /u /c type "%temp%\GetAdmin.vbs">"%temp%\GetAdminUnicode.vbs"
    cscript //nologo "%temp%\GetAdminUnicode.vbs" %1
    del /f /q "%temp%\GetAdmin.vbs" >nul 2>&1
    del /f /q "%temp%\GetAdminUnicode.vbs" >nul 2>&1
    exit /B
::===============================================================================================================
::===============================================================================================================
:GotAdmin
	if '%1'=='ELEV' shift /1
    if exist "%temp%\GetAdmin.vbs" del /f /q "%temp%\GetAdmin.vbs"
	if exist "%temp%\GetAdminUnicode.vbs" del /f /q "%temp%\GetAdminUnicode.vbs"
::===============================================================================================================
::===============================================================================================================
	cls
	mode con cols=82 lines=48
	color 1F
	echo:
::===============================================================================================================
:: DEFINE SYSTEM ENVIRONMENT
::===============================================================================================================
	for /F "tokens=6 delims=[]. " %%A in ('ver') do set /a win=%%A
	if %win% LSS 7601 (echo:)&&(echo:)&&(echo Unsupported Windows detected)&&(echo:)&&(echo Minimum OS must be Windows 7 SP1 or better)&&(echo:)&&(goto:TheEndIsNear)
	for /F "tokens=2 delims==" %%a in ('wmic path Win32_Processor get AddressWidth /value') do (set winx=win_x%%a)
	set "sls=SoftwareLicensingService"
	set "slp=SoftwareLicensingProduct"
	set "osps=OfficeSoftwareProtectionService"
	set "ospp=OfficeSoftwareProtectionProduct"
	for /F "tokens=2 delims==" %%A in ('"wmic path %sls% get version /VALUE" 2^>nul') do set "slsversion=%%A"
	if %win% LSS 9200 (for /F "tokens=2 delims==" %%A IN ('"wmic path %osps% get version /VALUE" 2^>nul') do set ospsversion=%%A)
	set "downpath=not set"
	set "o16updlocid=not set"
	set "o16arch=x86"
    set "o16lang=en-US"
	set "langtext=Default Language"
    set "o16lcid=1033"
::===============================================================================================================
:: Read OfficeRTool.ini
::===============================================================================================================
	SET /a countx=0
	cd /D "%OfficeRToolpath%"
	for /F "tokens=*" %%a in (OfficeRTool.ini) do (
		SET /a countx=!countx! + 1
		set var!countx!=%%a
	)
	if %countx% LSS 10 (echo:)&&(echo Error in OfficeRTool.ini)&&(echo:)&&(pause)&&(goto:Office16VnextInstall)
	set "inidownpath=%var3%"
	if "%inidownpath:~-1%" EQU " " set "inidownpath=%inidownpath:~0,-1%"
	set "downpath=%inidownpath%"
	set "inidownlang=%var6%"
	if "%inidownlang:~-1%" EQU " " set "inidownlang=%inidownlang:~0,-1%"
	set "o16lang=%inidownlang%"
	set "inidownarch=%var9%"
	if "%inidownarch:~-1%" EQU " " set "inidownarch=%inidownarch:~0,-1%"
	set "o16arch=%inidownarch%"
::===============================================================================================================
::===============================================================================================================
	if "%1" EQU "-NRC" goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:: Check Microsoft Content Delivery Network (CDN) server for new Office releases
::===============================================================================================================
	echo:
	echo Checking public Office distribution channels for new updates
	call :CheckNewVersion Monthly_Channel 492350f6-3a01-4f97-b9c0-c7c6ddf67d60
	call :CheckNewVersion Insider_Channel 5440fd1f-7ecb-4221-8110-145efaa6372f
	call :CheckNewVersion Monthly_Channel_Targeted 64256afe-f5d9-4f86-8936-8840a6a4f5be
	call :CheckNewVersion Semi_Annual_Channel 7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
	call :CheckNewVersion Semi_Annual_Channel_Targeted b8f9b850-328d-4355-9145-c59439a0c4cf
	call :CheckNewVersion Dogfood_DevMain_Channel ea4a4090-de26-49d7-93c1-91bff9e53fc3
	((echo:)&&(echo:)&&(echo:)&&(pause))
::===============================================================================================================
::===============================================================================================================
:Office16VnextInstall
	cd /D "%OfficeRToolpath%"
	mode con cols=82 lines=48
	color 1F
	title OfficeRTool - 2019/June/08 -
	cls
    echo:
    echo ================== OFFICE DOWNLOAD AND INSTALL =============================
	echo ____________________________________________________________________________
	echo:
	echo [D] DOWNLOAD OFFICE OFFLINE INSTALL PACKAGE
    echo:
	echo [I] INSTALL OFFICE SUITES OR SINGLE APPS
	echo:
	echo [C] CONVERT OFFICE RETAIL TO VOLUME
	echo ____________________________________________________________________________
	echo:
    echo [A] SHOW CURRENT ACTIVATION STATUS
	echo:
	echo [K] START KMS ACTIVATION
    echo ____________________________________________________________________________
	echo:
    echo [U] CHANGE OFFICE UPDATE-PATH (SWITCH DISTRIBUTION CHANNEL)
    echo ____________________________________________________________________________
	echo:
	echo [T] DISABLE ACQUISITION AND SENDING OF TELEMETRY DATA
	echo ____________________________________________________________________________
	echo:
    echo [O] CREATE OFFICE ONLINE WEB-INSTALLER LINK
    echo ____________________________________________________________________________
	echo:
    echo [E] END - STOP AND LEAVE PROGRAM
	echo:
    echo ============================================================================
	echo:
    CHOICE /C DICAKUTOEX /N /M "YOUR CHOICE ?"
    if %errorlevel%==1 goto:DownloadO16Offline
	if %errorlevel%==2 goto:InstallO16
    if %errorlevel%==3 goto:Convert16Activate
	if %errorlevel%==4 goto:CheckActivationStatus
	if %errorlevel%==5 goto:StartKMSActivation
    if %errorlevel%==6 goto:ChangeUpdPath
	if %errorlevel%==7 goto:DisableTelemetry
	if %errorlevel%==8 goto:DownloadO16Online
	if %errorlevel%==9 goto:TheEndIsNear
	if %errorlevel%==10 goto:TheEndIsNear
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================	
:CheckNewVersion
	set "o16build=not set "
	set "o16latestbuild=not set"
	if "%1" EQU "Manual_Override" goto:CheckNewVersionSkip1
	if exist "%OfficeRToolpath%\latest_%1_build.txt" (
		set /p "o16build=" <"%OfficeRToolpath%\latest_%1_build.txt"
		)
	set "o16build=%o16build:~0,-1%"
:CheckNewVersionSkip1
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --no-verbose --output-document="%TEMP%\V32.cab" --tries=20 http://officecdn.microsoft.com.edgesuite.net/pr/%2/Office/Data/V32.cab >nul 2>&1
 	if %errorlevel% GEQ 1 goto:ErrCheckNewVersion1
	expand "%TEMP%\V32.cab" -F:VersionDescriptor.xml "%TEMP%" >nul 2>&1
	type "%TEMP%\VersionDescriptor.xml" | find "Available Build" >"%TEMP%\found_office_build.txt"
 	if %errorlevel% GEQ 1 goto:ErrCheckNewVersion2
	set /p "o16latestbuild=" <"%TEMP%\found_office_build.txt"
	set "o16latestbuild=%o16latestbuild:~20,16%"
	set "spaces=     "
	if "%o16latestbuild:~15,1%" EQU " " ((set "o16latestbuild=%o16latestbuild:~0,14%")&&(set "spaces=       "))
	if "%o16latestbuild:~0,3%" NEQ "16." goto:ErrCheckNewVersion3
	(echo --^> Checking channel:      %1)
	if "%1" EQU "Manual_Override" goto:CheckNewVersionSkip2a
	if "%o16build%" NEQ "%o16latestbuild%" (
		powershell -noprofile -command "%pswindowtitle%"; Write-Host "'   'New Build available:'   '" -foreground "White" -nonewline; Write-Host "%o16latestbuild%'%spaces%'" -foreground "Red" -nonewline; Write-Host "LkgBuild:' '" -foreground "White" -nonewline; Write-Host "%o16build%" -foreground "Green"
		echo %o16latestbuild% >"%OfficeRToolpath%\latest_%1_build.txt"
		echo %o16build% >>"%OfficeRToolpath%\latest_%1_build.txt"
		goto:CheckNewVersionSkip2b
		)
:CheckNewVersionSkip2a
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "'   'Last known good Build:' '" -foreground "White" -nonewline; Write-Host "%o16latestbuild%'%spaces%'" -foreground "Green" -nonewline; Write-Host "No newer Build available" -foreground "White"
:CheckNewVersionSkip2b
	del /f /q "%TEMP%\V32.cab"
	del /f /q "%TEMP%\VersionDescriptor.xml"
	del /f /q "%TEMP%\found_office_build.txt"
	goto:eof
:ErrCheckNewVersion1
	echo:
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * channel" -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** No response from Office content delivery server" -foreground "Red"
	echo:
	echo Check Internet connection and/or Channel-ID.
	set "buildcheck=not ok"
	goto:eof
:ErrCheckNewVersion2
	echo:
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * " -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** No Build / Version number found in file: * VersionDescriptor.xml *" -foreground "Red"
	copy "%TEMP%\V32.cab" "%TEMP%\%1_V32.cab" >nul 2>&1
	copy "%TEMP%\VersionDescriptor.xml" "%TEMP%\%1_VersionDescriptor.xml" >nul 2>&1
	echo:
	echo Check file "%TEMP%\%1_V32.cab"
	echo Check file "%TEMP%\%1_VersionDescriptor.xml"
	set "buildcheck=not ok"
	goto:eof
:ErrCheckNewVersion3
	echo:
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * " -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Unsupported Build / Version number detected: * %o16latestbuild% *" -foreground "Red"
	copy "%TEMP%\VersionDescriptor.xml" "%TEMP%\%1_VersionDescriptor.xml" >nul 2>&1
	copy "%TEMP%\found_office_build.txt" "%TEMP%\%1_found_office_build.txt" >nul 2>&1
	echo:
	echo Check file "%TEMP%\%1_VersionDescriptor.xml"
	echo Check file "%TEMP%\%1_found_office_build.txt"
	set "buildcheck=not ok"
	goto:eof
::===============================================================================================================
::===============================================================================================================
:DownloadO16Offline
    cd /D "%OfficeRToolpath%"
	if not defined o16arch set "o16arch=x86"
    set "installtrigger=0"
	set "channeltrigger=0"
	set "o16updlocid=not set"
	set "o16build=not set "
	cls
	echo:
    echo ================== DOWNLOAD OFFICE OFFLINE INSTALL PACKAGE =================
    echo ____________________________________________________________________________
	echo:
	echo DownloadPath: "%inidownpath%"
    echo:
	if "%o16updlocid%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Channel-ID:    %o16updlocid% (Monthly_Channel) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" echo Channel-ID:    %o16updlocid% (Insider_Channel) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Channel-ID:    %o16updlocid% (Monthly_Channel_Targeted) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" echo Channel-ID:    %o16updlocid% (Semi_Annual_Channel) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:    %o16updlocid% (Semi_Annual_Channel_Targeted) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" echo Channel-ID:    %o16updlocid% (Dogfood_DevMain_Channel) && goto:DownOfflineContinue
	if "%o16updlocid%" EQU "not set" echo Channel-ID:    not set && goto:DownOfflineContinue
	echo Channel-ID:    %o16updlocid% (Manual_Override)
::===============================================================================================================
:DownOfflineContinue
	echo:
	echo Office build:  %o16build%
	echo:
	echo Language:      %o16lang% (%langtext%)
    echo:
	echo Architecture:  %o16arch%
    echo ____________________________________________________________________________
	echo:
	echo Set new Office Package download path or press return for
	set /p downpath=DownloadPath^= "%downpath%" ^>
	set "downpath=%downpath:"=%"
	if /I "%downpath%" EQU "X" (set "downpath=not set")&&(goto:Office16VnextInstall)
	set "downdrive=%downpath:~0,2%"
	if "%downdrive:~-1%" NEQ ":" (echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:DownloadO16Offline)
	cd /d %downdrive%\ >nul 2>&1
	if errorlevel 1 (echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:DownloadO16Offline)
	set "downdrive=%downpath:~0,3%"
	if "%downdrive:~-1%" EQU "\" (set "downpath=%downdrive%%downpath:~3%") else (set "downpath=%downdrive:~0,2%\%downpath:~2%")
	if "%downpath:~-1%" EQU "\" set "downpath=%downpath:~0,-1%"
::===============================================================================================================
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	echo:
	if "%inidownpath%" NEQ "%downpath%" ((echo Office install package download path changed)&&(echo old path "%inidownpath%" -- new path "%downpath%")&&(echo:))
	if "%inidownpath%" NEQ "%downpath%" set /p installtrigger=Save new path to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:SkipDownPathSave
	if /I "%installtrigger%" EQU "X" goto:SkipDownPathSave
	set "inidownpath=%downpath%"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download path saved.
::===============================================================================================================
:SkipDownPathSave
	echo:
	echo "Public known" standard distribution channels
	echo Channel Name                                    - Internal Naming   Index-#
	echo ___________________________________________________________________________
	echo:
	echo Monthly_Channel (Retail/RTM)                    - (Production::CC)     (1)
	echo Insider_Channel (Office Insider FAST)           - (Insiders::DevMain)  (2)
	echo Monthly_Channel_Targeted (Office Insider SLOW)  - (Insiders::CC)       (3)
	echo Semi_Annual_Channel (Business)                  - (Production::DC)     (4)
	echo Semi_Annual_Channel_Targeted (Business Insider) - (Insiders::FRDC)     (5)
	echo Manual_Override (set identifier for Channel-ID's not public known)     (M)
	echo Exit to Main Menu                                                      (X)
	echo:
	set /p channeltrigger=Set Channel-Index-# (1,2,3,4,5,M) or X ^>
	if "%channeltrigger%" EQU "1" goto:ChanSel1
	if "%channeltrigger%" EQU "2" goto:ChanSel2
	if "%channeltrigger%" EQU "3" goto:ChanSel3
	if "%channeltrigger%" EQU "4" goto:ChanSel4
	if "%channeltrigger%" EQU "5" goto:ChanSel5
	if /I "%channeltrigger%" EQU "M" ((set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:ChanSelMan))
	if "%channeltrigger%" EQU "0" ((set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall))
	if /I "%channeltrigger%" EQU "X" ((set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall))
	(set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:DownloadO16Offline)
::===============================================================================================================
:ChanSel1
	set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
	call :CheckNewVersion Monthly_Channel %o16updlocid%
	set "o16build=%o16latestbuild%"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel2
	set "o16updlocid=5440fd1f-7ecb-4221-8110-145efaa6372f"
	call :CheckNewVersion Insider_Channel %o16updlocid%
	set "o16build=%o16latestbuild%"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel3
	set "o16updlocid=64256afe-f5d9-4f86-8936-8840a6a4f5be"
	call :CheckNewVersion Monthly_Channel_Targeted %o16updlocid%
	set "o16build=%o16latestbuild%"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel4
	set "o16updlocid=7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
	call :CheckNewVersion Semi_Annual_Channel %o16updlocid%
	set "o16build=%o16latestbuild%"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel5
	set "o16updlocid=b8f9b850-328d-4355-9145-c59439a0c4cf"
	call :CheckNewVersion Semi_Annual_Channel_Targeted %o16updlocid%
	set "o16build=%o16latestbuild%"
	goto:ChannelSelected
::===============================================================================================================
:ChanSelMan
    echo:
	echo "Microsoft Internal Use Only" Beta/Testing distribution channels
	echo Internal Naming           Channel-ID:                               Index-#
	echo ___________________________________________________________________________
	echo:
    echo Dogfood::DevMain   -----^> ea4a4090-de26-49d7-93c1-91bff9e53fc3         (1)
    echo Dogfood::CC        -----^> f3260cf1-a92c-4c75-b02e-d64c0a86a968         (2)
    echo Dogfood::DCEXT     -----^> c4a7726f-06ea-48e2-a13a-9d78849eb706         (3)
    echo Dogfood::FRDC      -----^> 834504cc-dc55-4c6d-9e71-e024d0253f6d         (4)
    echo Microsoft::CC      -----^> 5462eee5-1e97-495b-9370-853cd873bb07         (5)
    echo Microsoft::DC      -----^> f4f024c8-d611-4748-a7e0-02b6e754c0fe         (6)
    echo Microsoft::DevMain -----^> b61285dd-d9f7-41f2-9757-8f61cba4e9c8         (7)
    echo Microsoft::FRDC    -----^> 9a3b7ff2-58ed-40fd-add5-1e5158059d1c         (8)
    echo Production::LTSC   -----^> f2e724c1-748f-4b47-8fb8-8e0d210e9208         (9)
	echo Insider::LTSC      -----^> 2e148de9-61c8-4051-b103-4af54baffbb4         (A)
	echo Exit to Main Menu                                                      (X)
    echo:
	set /p o16updlocid=Set Channel (enter Channel-ID or Index-#) ^>
	if "%o16updlocid%" EQU "not set" goto:DownloadO16Offline
	if /I "%o16updlocid%" EQU "X" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if "%o16updlocid%" EQU "0" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if "%o16updlocid%" EQU "1" set "o16updlocid=ea4a4090-de26-49d7-93c1-91bff9e53fc3"
	if "%o16updlocid%" EQU "2" set "o16updlocid=f3260cf1-a92c-4c75-b02e-d64c0a86a968"
    if "%o16updlocid%" EQU "3" set "o16updlocid=c4a7726f-06ea-48e2-a13a-9d78849eb706"
    if "%o16updlocid%" EQU "4" set "o16updlocid=834504cc-dc55-4c6d-9e71-e024d0253f6d
    if "%o16updlocid%" EQU "5" set "o16updlocid=5462eee5-1e97-495b-9370-853cd873bb07"
    if "%o16updlocid%" EQU "6" set "o16updlocid=f4f024c8-d611-4748-a7e0-02b6e754c0fe"
    if "%o16updlocid%" EQU "7" set "o16updlocid=b61285dd-d9f7-41f2-9757-8f61cba4e9c8"
    if "%o16updlocid%" EQU "8" set "o16updlocid=9a3b7ff2-58ed-40fd-add5-1e5158059d1c"
    if "%o16updlocid%" EQU "9" set "o16updlocid=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
	if /I "%o16updlocid%" EQU "A" set "o16updlocid=2e148de9-61c8-4051-b103-4af54baffbb4"
    call :CheckNewVersion Manual_Override %o16updlocid%
	set "o16build=%o16latestbuild%"
::===============================================================================================================
:ChannelSelected
	set "o16downloadloc=officecdn.microsoft.com.edgesuite.net/pr/%o16updlocid%/Office/Data"
	echo:
	if "%buildcheck%" EQU "not ok" ((pause)&&(set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall))
    set /p o16build=Set Office Build - or press return for %o16build% ^>
	if "%o16build%" EQU "cb" (set "o16build=%o16latestbuild%")&&(goto:DownloadO16Offline)
	if "%o16build%" EQU "not set" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if /I "%o16build%" EQU "X" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
::===============================================================================================================
:LangSelect
	echo:
	echo Possible Language VALUES (enter exactly in small-CAPITAL letters as shown):
    echo ar-SA, bg-BG, cs-CZ, da-DK, de-DE, el-GR, en-US, es-ES, et-EE, fi-FI, fr-FR,
    echo he-IL, hi-IN, hr-HR, hu-HU, id-ID, it-IT, ja-JP, kk-KZ, ko-KR, lt-LT, lv-LV,
    echo ms-MY, nb-NO, nl-NL, pl-PL, pt-BR, pt-PT, ro-RO, ru-RU, sk-SK, sl-SI, sr-latn-RS,
	echo sv-SE, th-TH, tr-TR, uk-UA, vi-VN, zh-CN, zh-TW
	echo:
    set /p o16lang=Set Language Value - or press return for %o16lang% ^>
	call :SetO16Language
	if "%langnotfound%" EQU "TRUE" goto:LangSelect
::===============================================================================================================
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	if "%inidownlang%" NEQ "%o16lang%" ((echo:)&&(echo Office install package download language changed)&&(echo old language "%inidownlang%" -- new language "%o16lang%")&&(echo:))
	if "%inidownlang%" NEQ "%o16lang%" set /p installtrigger=Save new language to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:ArchSelect
	if /I "%installtrigger%" EQU "X" goto:ArchSelect
	set "inidownlang=%o16lang%"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download language saved.
::===============================================================================================================
:ArchSelect
	echo:
	set /p o16arch=Set architecture to download (x86 or x64) - or press return for %o16arch% ^>
	if "%o16arch%" EQU "not set" goto:ArchSelect
	if "%o16arch%" EQU "x86" goto:SkipArchSelect
	if "%o16arch%" EQU "X86" (set "o16arch=x86")&&(goto:SkipArchSelect)
	if "%o16arch%" EQU "x64" goto:SkipArchSelect
	if "%o16arch%" EQU "X64" (set "o16arch=x64")&&(goto:SkipArchSelect)
	set "o16arch=x86"
::===============================================================================================================
:SkipArchSelect
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	echo:
	if "%inidownarch%" NEQ "%o16arch%" ((echo Office install package download architecture changed)&&(echo old architecture "%inidownarch%" -- new architecture "%o16arch%")&&(echo:))
	if "%inidownarch%" NEQ "%o16arch%" set /p installtrigger=Save new architecture to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:SkipDownArchSave
	if /I "%installtrigger%" EQU "X" goto:SkipDownArchSave
	set "inidownarch=%o16arch%"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download architecture saved.
::===============================================================================================================
:SkipDownArchSave
	echo:
	echo ____________________________________________________________________________
    echo:
    echo ================== Pending Download (SUMMARY) ==============================
    echo:
    echo DownloadPath: %downpath%
    echo:
	if "%o16updlocid%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Channel-ID:   %o16updlocid% (Monthly_Channel) && goto:PendDownContinue
	if "%o16updlocid%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" echo Channel-ID:   %o16updlocid% (Insider_Channel) && goto:PendDownContinue
	if "%o16updlocid%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Channel-ID:   %o16updlocid% (Monthly_Channel_Targeted) && goto:PendDownContinue
	if "%o16updlocid%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" echo Channel-ID:   %o16updlocid% (Semi_Annual_Channel) && goto:PendDownContinue
	if "%o16updlocid%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:   %o16updlocid% (Semi_Annual_Channel_Targeted) && goto:PendDownContinue
	if "%o16updlocid%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" echo Channel-ID:   %o16updlocid% (Dogfood_DevMain_Channel) && goto:PendDownContinue
	echo Channel-ID:   %o16updlocid% (Manual_Override)
::===============================================================================================================
:PendDownContinue
	set "installtrigger=0"
	echo Office Build: %o16build%
	echo Language:     %o16lang% (%langtext%)
    echo Architecture: %o16arch%
    echo ____________________________________________________________________________
	echo:
    set /p installtrigger=Start download now? (1/0) ^>
    if "%installtrigger%" EQU "0" goto:DownloadO16Offline
    if "%installtrigger%" EQU "1" goto:Office16VNextDownload
	if /I "%installtrigger%" EQU "X" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	goto:DownloadO16Offline
::===============================================================================================================
::===============================================================================================================
:Office16VNextDownload
	cls
	echo:
	echo ================== DOWNLOADING OFFICE OFFLINE SETUP PACKAGE ================
	echo ____________________________________________________________________________
	if "%o16updlocid%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" set "downbranch=Monthly_Channel" && goto:ContVNextDownload
	if "%o16updlocid%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" set "downbranch=Insider_Channel" && goto:ContVNextDownload
	if "%o16updlocid%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" set "downbranch=Monthly_Channel_Targeted" && goto:ContVNextDownload
	if "%o16updlocid%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" set "downbranch=Semi_Annual_Channel" && goto:ContVNextDownload
	if "%o16updlocid%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" set "downbranch=Semi_Annual_Channel_Targeted" && goto:ContVNextDownload
	if "%o16updlocid%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" set "downbranch=Dogfood_DevMain" && goto:ContVNextDownload
	set "downbranch=Manual_Override"
::===============================================================================================================
:ContVNextDownload
	cd /d "%downdrive%\" >nul 2>&1
	md "%downpath%" >nul 2>&1
	cd /d "%downpath%" >nul 2>&1
	mode con cols=147
	if "%o16arch%" EQU "x64" goto:X64DOWNLOAD
::===============================================================================================================
::	Download x86/32bit Office setup files
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/v32_%o16build%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v32_%o16build%.cab"
	copy %o16build%_%o16lang%_%o16arch%_%downbranch%\Office\Data\v32_%o16build%.cab %o16build%_%o16lang%_%o16arch%_%downbranch%\Office\Data\v32.cab >nul 2>&1
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/stream.x86.%o16lang%.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/stream.x86.%o16lang%.dat"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat"
	goto:GENERALDOWNLOAD
::===============================================================================================================
::	Download x64/64bit Office setup files
:X64DOWNLOAD
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/v64_%o16build%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v64_%o16build%.cab"
	copy %o16build%_%o16lang%_%o16arch%_%downbranch%\Office\Data\v64_%o16build%.cab %o16build%_%o16lang%_%o16arch%_%downbranch%\Office\Data\v64.cab >nul 2>&1
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/stream.x64.%o16lang%.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/stream.x64.%o16lang%.dat"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat"
::===============================================================================================================	
:: Download setup file(s) used in both x86 and x64 architectures
:GENERALDOWNLOAD
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/i320.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/i320.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/i32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/i32%o16lcid%.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/s320.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/s320.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/s32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/s32%o16lcid%.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/i640.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/i640.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/i64%o16lcid%.cab	
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/i64%o16lcid%.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/s640.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/s640.cab"
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --progress=bar:force:noscroll --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_%o16arch%_%downbranch% http://%o16downloadloc%/%o16build%/s64%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/%o16build%/s64%o16lcid%.cab"
::===============================================================================================================	
	echo ____________________________________________________________________________
	if "%downbranch%" EQU "Monthly_Channel" echo Current>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Insider_Channel" echo InsiderFast>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Monthly_Channel_Targeted" echo FirstReleaseCurrent>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Semi_Annual_Channel" echo Deferred>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Semi_Annual_Channel_Targeted" echo FirstReleaseDeferred>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Dogfood_DevMain" echo DogfoodDevMain>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	if "%downbranch%" EQU "Manual_Override" echo ManualOverride>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	echo %o16build%>>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	echo %o16lang%>>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	echo %o16arch%>>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	echo %o16updlocid%>>%o16build%_%o16lang%_%o16arch%_%downbranch%\package.info
	echo:
	echo:
    timeout /t 7
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:WgetError
	set "errortrigger=0"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR downloading: %1" -foreground "Red"
	echo:
	set /p errortrigger=Cancel Download now? (1/0) ^>
	if "%errortrigger%" EQU "1" (
		if exist "%downpath%\%o16build%_%o16lang%_%o16arch%_%downbranch%" rd "%downpath%\%o16build%_%o16lang%_%o16arch%_%downbranch%" /S /Q
		goto:Office16VnextInstall
	)
	echo:
	goto :eof
::===============================================================================================================
::===============================================================================================================
:DownloadO16Online
    cd /D "%OfficeRToolpath%"
    cls
	echo:
	echo ============= DOWNLOAD OFFICE 2016/2019 ONLINE WEB INSTALLER ===============
	echo ____________________________________________________________________________
    set "WebProduct=not set"
    set "o16arch=x86"
    set "o16lang=en-US"
    set "of16install=0"
    set "pr16install=0"
    set "vi16install=0"
    set "of19install=0"
    set "pr19install=0"
    set "vi19install=0"
    set "installtrigger=O"
	echo:
	set /p installtrigger=Generate Office 2016 products setup.exe download-link (1=YES/0=NO) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:WEBOFF2016
	echo:
	set /p installtrigger=Generate Office 2019 products setup.exe download-link (1=YES/0=NO) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:WEBOFF2019
	goto:DownloadO16Online
:WEBOFF2016
	echo:
	echo ____________________________________________________________________________
	echo:
    set /p of16install=Set Office 2016 ProfessionalPlus Install (1/0) ^>
	if "%of16install%" EQU "1" (set "WebProduct=ProPlusRetail")&&(goto:WebArchSelect)
    echo:
    set /p pr16install=Set Project 2016 Professional Install (1/0) ^>
	if "%pr16install%" EQU "1" (set "WebProduct=ProjectProRetail")&&(goto:WebArchSelect)
    echo:
    set /p vi16install=Set Visio 2016 Professional Install (1/0) ^>
	if "%vi16install%" EQU "1" (set "WebProduct=VisioProRetail")&&(goto:WebArchSelect)
	goto:WEBOFFNOTHING
:WEBOFF2019
	echo:
	echo ____________________________________________________________________________
	echo:
    set /p of19install=Set Office 2019 ProfessionalPlus Install (1/0) ^>
	if "%of19install%" EQU "1" (set "WebProduct=ProPlus2019Retail")&&(goto:WebArchSelect)
    echo:
    set /p pr19install=Set Project 2019 Professional Install (1/0) ^>
	if "%pr19install%" EQU "1" (set "WebProduct=ProjectPro2019Retail")&&(goto:WebArchSelect)
    echo:
    set /p vi19install=Set Visio 2019 Professional Install (1/0) ^>
	if "%vi19install%" EQU "1" (set "WebProduct=VisioPro2019Retail")&&(goto:WebArchSelect)
:WEBOFFNOTHING
	echo:
	echo ____________________________________________________________________________
	echo:
	echo Nothing selected - Returning to Main Menu now
	echo:
	pause
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:WebArchSelect
    echo ____________________________________________________________________________
    if "%winx%" EQU "win_x32" (set "o16arch=x86")&&(goto:WebLangSelect)
	echo:
    set /p o16arch=Set architecture to install (enter x86 or x64) - press return for %o16arch% ^>
	echo ____________________________________________________________________________
	
::===============================================================================================================
::===============================================================================================================
:WebLangSelect
	echo:
    echo Possible Language VALUES (enter exactly in small-CAPITAL letters as shown):
    echo ar-SA, bg-BG, cs-CZ, da-DK, de-DE, el-GR, en-US, es-ES, et-EE, fi-FI, fr-FR,
    echo he-IL, hi-IN, hr-HR, hu-HU, id-ID, it-IT, ja-JP, kk-KZ, ko-KR, lt-LT, lv-LV,
    echo ms-MY, nb-NO, nl-NL, pl-PL, pt-BR, pt-PT, ro-RO, ru-RU, sk-SK, sl-SI, sr-latn-RS,
	echo sv-SE, th-TH, tr-TR, uk-UA, vi-VN, zh-CN, zh-TW
    echo:
    set /p o16lang=Set Language Value - or press return for %o16lang% ^>
	call :SetO16Language
	if "%langnotfound%" EQU "TRUE" goto:WebLangSelect
::===============================================================================================================
    echo:
    echo ____________________________________________________________________________
	echo:
    echo Pending Online WEB Install (SUMMARY)
    echo:
    if "%of16install%" EQU "1" echo Install Office 2016 ?      : YES
    if "%pr16install%" EQU "1" echo Install Project 2016 ?     : YES
    if "%vi16install%" EQU "1" echo Install Visio 2016 ?       : YES
    if "%of19install%" EQU "1" echo Install Office 2019 ?      : YES
    if "%pr19install%" EQU "1" echo Install Project 2019 ?     : YES
    if "%vi19install%" EQU "1" echo Install Visio 2019 ?       : YES
    echo:
    echo Install Architecture ?     : %o16arch%
    echo Install Language ?         : %o16lang%
    echo ____________________________________________________________________________
	echo:
    set /p installtrigger=Start Online WEB Install now (1/0)? ^>
    if "%installtrigger%" EQU "0" goto:DownloadO16Online
    if "%installtrigger%" EQU "1" goto:OfficeWebInstall
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
    goto:DownloadO16Online
::===============================================================================================================
::===============================================================================================================
:OfficeWebInstall
    cls
	echo:
    echo ================== DOWNLOAD OFFICE ONLINE WEB INSTALLER ====================
	echo ____________________________________________________________________________
	echo:
    echo Sending generated link to your browser.
    echo:
    echo Save the offered Setup.exe and run it to start Online WEB Install
    start "" "https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/Office/Data/setup%WebProduct%.%o16arch%.%o16lang%.exe
    echo ____________________________________________________________________________
	echo:
    echo:
	timeout /t 7
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckActivationStatus
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	set "CDNBaseUrl=not set"
	set "UpdateUrl=not set"
	set "UpdateBranch=not set"
	cls
	powershell.exe -command "& {$pshost = Get-Host;$pswindow = $pshost.UI.RawUI;$newsize = $pswindow.BufferSize;$newsize.height = 100;$pswindow.buffersize = $newsize;}"
	echo:
	echo ================== SHOW CURRENT ACTIVATION STATUS ==========================
    echo ____________________________________________________________________________
	echo:
	echo Office installation path:
	echo %installpath16%
	echo:
	if "%ProPlusVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%StandardVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%ProjectProVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%VisioProVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%_UWPappINSTALLED%" EQU "YES" ((set "ChannelName=Microsoft Apps Store")&&(set "UpdateUrl=Microsoft Apps Store")&&(goto:CheckActCont))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "CDNBaseUrl" 2^>nul') DO (Set "CDNBaseUrl=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "UpdateUrl" 2^>nul') DO (Set "UpdateUrl=%%B")
	call:DecodeChannelName %UpdateUrl%
::===============================================================================================================
:CheckActCont
	echo Distribution-Channel:
	echo %ChannelName%
	echo:
	echo Updates-Url:
	echo %UpdateUrl%
	echo ____________________________________________________________________________
	echo:
	if "%_ProPlusRetail%" EQU "YES" ((echo Office 2016 ProfessionalPlus --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_ProPlus2019Retail%" EQU "YES" ((echo Office 2019 ProfessionalPlus --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2019))
	if "%_ProPlus2019Volume%" EQU "YES" ((echo Office 2019 ProfessionalPlus --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2019))
	if "%_StandardRetail%" EQU "YES" ((echo Office 2016 Standard --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standart))
	if "%_O365ProPlusRetail%" EQU "YES" ((echo Office 365 ProfessionalPlus --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_O365BusinessRetail%" EQU "YES" ((echo Office 365 Business --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_MondoRetail%" EQU "YES" ((echo Office Mondo Grande Suite --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_WordRetail%" EQU "YES" ((echo Word 2016 SingleApp ------------------ ProductVersion : %o16version%)&&(call :CheckKMSActivation Word))
	if "%_ExcelRetail%" EQU "YES" ((echo Excel 2016 SingleApp ----------------- ProductVersion : %o16version%)&&(call :CheckKMSActivation Excel))
	if "%_PowerPointRetail%" EQU "YES" ((echo PowerPoint 2016 SingleApp ------------ ProductVersion : %o16version%)&&(call :CheckKMSActivation PowerPoint))
	if "%_AccessRetail%" EQU "YES" ((echo Access 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access))
	if "%_OutlookRetail%" EQU "YES" ((echo Outlook 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook))
	if "%_PublisherRetail%" EQU "YES" ((echo Publisher 2016 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher))
	if "%_OneNoteRetail%" EQU "YES" ((echo OneNote 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation OneNote))
	if "%_SkypeForBusinessRetail%" EQU "YES" ((echo Skype For Business 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness))
	if "%_AppxWinword%" EQU "YES" ((echo Word 2016 UWP Appx Desktop App --- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxExcel%" EQU "YES" ((echo Excel 2016 UWP Appx Desktop App --- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxPowerPoint%" EQU "YES" ((echo PowerPoint 2016 UWP Appx Desktop App - ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxAccess%" EQU "YES" ((echo Access 2016 UWP Appx Desktop App ----- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxOutlook%" EQU "YES" ((echo Outlook 2016 UWP Appx Desktop App ---- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxPublisher%" EQU "YES" ((echo Publisher 2016 UWP Appx Desktop App -- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxOneNote%" EQU "YES" ((echo OneNote 2016 UWP Appx Desktop App ---- ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_AppxSkypeForBusiness%" EQU "YES" ((echo Skype 2016 UWP Appx Desktop App ------ ProductVersion : %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_Word2019Retail%" EQU "YES" ((echo Word 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2019))
	if "%_Excel2019Retail%" EQU "YES" ((echo Excel 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2019))
	if "%_PowerPoint2019Retail%" EQU "YES" ((echo PowerPoint 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2019))
	if "%_Access2019Retail%" EQU "YES" ((echo Access 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2019))
	if "%_Outlook2019Retail%" EQU "YES" ((echo Outlook 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2019))
	if "%_Publisher2019Retail%" EQU "YES" ((echo Publisher 2019 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2019))
	if "%_SkypeForBusiness2019Retail%" EQU "YES" ((echo Skype For Business 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2019))
	if "%_Word2019Volume%" EQU "YES" ((echo Word 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2019))
	if "%_Excel2019Volume%" EQU "YES" ((echo Excel 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2019))
	if "%_PowerPoint2019Volume%" EQU "YES" ((echo PowerPoint 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2019))
	if "%_Access2019Volume%" EQU "YES" ((echo Access 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2019))
	if "%_Outlook2019Volume%" EQU "YES" ((echo Outlook 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2019))
	if "%_Publisher2019Volume%" EQU "YES" ((echo Publisher 2019 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2019))
	if "%_SkypeForBusiness2019Volume%" EQU "YES" ((echo Skype For Business 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2019))
	if "%_VisioProRetail%" EQU "YES" ((echo VisioPro 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro))
	if "%_AppxVisio%" EQU "YES" ((echo VisioPro 2016 UWP Appx Desktop App --- ProductVersion : %o16version%)&&(call :CheckKMSActivation VisioPro))
	if "%_ProjectProRetail%" EQU "YES" ((echo Project 2016 Professional --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro))
	if "%_AppxProject%" EQU "YES" ((echo ProjectPro 2016 UWP Appx Desktop App - ProductVersion : %o16version%)&&(call :CheckKMSActivation ProjectPro))
	if "%_VisioPro2019Retail%" EQU "YES" ((echo Visio 2019 Professional ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2019))
	if "%_VisioPro2019Volume%" EQU "YES" ((echo Visio 2019 Professional ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2019))
	if "%_ProjectPro2019Retail%" EQU "YES" ((echo Project 2019 Professional --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2019))
	if "%_ProjectPro2019Volume%" EQU "YES" ((echo Project 2019 Professional --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2019))
	echo:
	echo:
	pause
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckKMSActivation
	set "LicStatus=9"
	set "LicStatusText=(---UNKNOWN---)           "
	set /a "GraceMin=0"
	set "EvalEndDate=00000000"
	set "activationtext=unknown"
	set "PartProdKey=not set"
	if %win% GEQ 9200	(
		for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get LicenseFamily /format:list" 2^>nul') do (set "LicFamily=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get LicenseStatus /value" 2^>nul') do (set "LicStatus=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get GracePeriodRemaining /value" 2^>nul') do (set /a "GraceMin=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get EvaluationEndDate /format:list" 2^>nul') do (set "EvalEndDate=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get PartialProductKey /format:list" 2^>nul') do (set "PartProdKey=%%A")
		)
	if %win% LSS 9200	(
		for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get LicenseFamily /format:list" 2^>nul') do (set /a "LicFamily=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get LicenseStatus /value" 2^>nul') do (set /a "LicStatus=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get GracePeriodRemaining /value" 2^>nul') do (set /a "GraceMin=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get EvaluationEndDate /format:list" 2^>nul') do (set "EvalEndDate=%%A")
		for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where (Name like '%%^%1^%%' and PartialProductKey is not NULL) get PartialProductKey /format:list" 2^>nul') do (set "PartProdKey=%%A")
		)
	set /a GraceDays=%GraceMin%/1440
	set "GraceDays=  %GraceDays%"
	set "GraceDays=%GraceDays:~-3%"
	set "EvalEndDate=%EvalEndDate:~0,8%"
	set "EvalEndDate=%EvalEndDate:~4,2%/%EvalEndDate:~6,2%/%EvalEndDate:~0,4%"
	if "%LicStatus%" EQU "0" (set "LicStatusText=(---UNLICENSED---)        ")
	if "%LicStatus%" EQU "1" (set "LicStatusText=(---LICENSED---)          ")
	if "%LicStatus%" EQU "2" (set "LicStatusText=(---OOB_GRACE---)         ")
	if "%LicStatus%" EQU "3" (set "LicStatusText=(---OOT_GRACE---)         ")
	if "%LicStatus%" EQU "4" (set "LicStatusText=(---NONGENUINE_GRACE---)  ")
	if "%LicStatus%" EQU "5" (set "LicStatusText=(---NOTIFICATIONS---)     ")
	if "%LicStatus%" EQU "6" (set "LicStatusText=(---EXTENDED_GRACE---)    ")
	echo:
	echo License Family: %LicFamily%
	echo:
	echo Activation status: %LicStatus%  %LicStatusText% PartialProductKey: -%PartProdKey%
	if "%EvalEndDate%" NEQ "01/01/1601" (set "activationtext=Product's activation is time-restricted")
	if "%EvalEndDate%" EQU "01/01/1601" (set "activationtext=Product is permanently activated")
	if %LicStatus% EQU 1 if %GraceMin% EQU 0 ((echo:)&&(echo Remaining Retail activation period: %activationtext%))
	if %LicStatus% GEQ 1 if %GraceDays% GEQ 1 (echo:)
	if %LicStatus% GEQ 1 if %GraceDays% GEQ 1 powershell -noprofile -command "%pswindowtitle%"; Write-Host "Remaining KMS activation period: '%GraceDays%' days left '-' License expires at:' '" -nonewline; Get-Date -date $(Get-Date).AddMinutes(%GraceMin%) -Format (Get-Culture).DateTimeFormat.ShortDatePattern
	if "%EvalEndDate%" NEQ "00/00/0000" if "%EvalEndDate%" NEQ "01/01/1601" ((echo:)&&(echo Evaluation-/Beta-Version timebomb active - Product end-of-life: %EvalEndDate%))
	echo ____________________________________________________________________________
	echo:
	goto :eof
::===============================================================================================================
::===============================================================================================================
:ChangeUpdPath
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	set "CDNBaseUrl=not set"
	set "UpdateUrl=not set"
	set "UpdateBranch=not set"
	set "installtrigger=O"
	set "channeltrigger=O"
	set "restrictbuild=newest available"
	set "updatetoversion="
	cls
	echo:
	echo ================== CHANGE INSTALLED OFFICE UPDATE-PATH =====================
    echo ____________________________________________________________________________
	echo:
	if "%ProPlusVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%StandardVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%ProjectProVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%VisioProVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%_UWPappINSTALLED%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for Office UWP Appx Desktop Apps)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%_ProPlusRetail%" EQU "YES"              (echo Office 2016 ProfessionalPlus --------- ProductVersion : %o16version%)
	if "%_ProPlus2019Retail%" EQU "YES"          (echo Office 2019 ProfessionalPlus --------- ProductVersion : %o16version%)
	if "%_ProPlus2019Volume%" EQU "YES"          (echo Office 2019 ProfessionalPlus --------- ProductVersion : %o16version%)
	if "%_O365ProPlusRetail%" EQU "YES"          (echo Office 365 ProfessionalPlus ---------- ProductVersion : %o16version%)
	if "%_O365BusinessRetail%" EQU "YES"         (echo Office 365 Business ------------------ ProductVersion : %o16version%)
	if "%_MondoRetail%" EQU "YES"                (echo Office Mondo Grande Suite ------------ ProductVersion : %o16version%)
	if "%_WordRetail%" EQU "YES"                 (echo Word 2016 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_ExcelRetail%" EQU "YES"                (echo Excel 2016 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPointRetail%" EQU "YES"           (echo PowerPoint 2016 SingleApp ------------ ProductVersion : %o16version%)
	if "%_AccessRetail%" EQU "YES"               (echo Access 2016 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_OutlookRetail%" EQU "YES"              (echo Outlook 2016 SingleApp --------------- ProductVersion : %o16version%)
	if "%_PublisherRetail%" EQU "YES"            (echo Publisher 2016 SingleApp ------------- ProductVersion : %o16version%)
	if "%_OneNoteRetail%" EQU "YES"              (echo OneNote 2016 SingleApp --------------- ProductVersion : %o16version%)
	if "%_SkypeForBusinessRetail%" EQU "YES"     (echo Skype 2016 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_AppxWinword%" EQU "YES"                (echo Word 2016 UWP Appx Desktop App ------- ProductVersion : %o16version%)
	if "%_AppxExcel%" EQU "YES"                  (echo Excel 2016 UWP Appx Desktop App ------ ProductVersion : %o16version%)
	if "%_AppxPowerPoint%" EQU "YES"             (echo PowerPoint 2016 UWP Appx Desktop App - ProductVersion : %o16version%)
	if "%_AppxAccess%" EQU "YES"                 (echo Access 2016 UWP Appx Desktop App ----- ProductVersion : %o16version%)
	if "%_AppxOutlook%" EQU "YES"                (echo Outlook 2016 UWP Appx Desktop App ---- ProductVersion : %o16version%)
	if "%_AppxPublisher%" EQU "YES"              (echo Publisher 2016 UWP Appx Desktop App -- ProductVersion : %o16version%)
	if "%_AppxOneNote%" EQU "YES"                (echo OneNote 2016 UWP Appx Desktop App ---- ProductVersion : %o16version%)
	if "%_AppxSkypeForBusiness%" EQU "YES"       (echo Skype 2016 UWP Appx Desktop App ------ ProductVersion : %o16version%)
	if "%_Word2019Retail%" EQU "YES"             (echo Word 2019 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2019Retail%" EQU "YES"            (echo Excel 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2019Retail%" EQU "YES"       (echo PowerPoint 2019 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2019Retail%" EQU "YES"           (echo Access 2019 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2019Retail%" EQU "YES"          (echo Outlook 2019 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2019Retail%" EQU "YES"        (echo Publisher 2019 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2019Retail%" EQU "YES" (echo Skype 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_Word2019Volume%" EQU "YES"             (echo Word 2019 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2019Volume%" EQU "YES"            (echo Excel 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2019Volume%" EQU "YES"       (echo PowerPoint 2019 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2019Volume%" EQU "YES"           (echo Access 2019 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2019Volume%" EQU "YES"          (echo Outlook 2019 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2019Volume%" EQU "YES"        (echo Publisher 2019 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2019Volume%" EQU "YES" (echo Skype 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_VisioProRetail%" EQU "YES"             (echo VisioPro 2016 ------------------------ ProductVersion : %o16version%)
	if "%_AppxVisio%" EQU "YES"                  (echo VisioPro 2016 UWP Appx Desktop App --- ProductVersion : %o16version%)
	if "%_VisioPro2019Retail%" EQU "YES"         (echo VisioPro 2019 ------------------------ ProductVersion : %o16version%)
	if "%_VisioPro2019Volume%" EQU "YES"         (echo VisioPro 2019 ------------------------ ProductVersion : %o16version%)
	if "%_ProjectProRetail%" EQU "YES"           (echo ProjectPro 2016 ---------------------- ProductVersion : %o16version%)
	if "%_AppxProject%" EQU "YES"                (echo ProjectPro 2016 UWP Appx Desktop App - ProductVersion : %o16version%)
	if "%_ProjectPro2019Retail%" EQU "YES"       (echo ProjectPro 2019 ---------------------- ProductVersion : %o16version%)
	if "%_ProjectPro2019Volume%" EQU "YES"       (echo ProjectPro 2019 ---------------------- ProductVersion : %o16version%)
	echo ____________________________________________________________________________
	echo:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "CDNBaseUrl" 2^>nul') DO (Set "CDNBaseUrl=%%B")
	call:DecodeChannelName %CDNBaseUrl%
	echo Distribution-Channel:
	echo %ChannelName%
	echo:
	echo CDNBase-Url:
	echo %CDNBaseUrl%
	echo:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "UpdateUrl" 2^>nul') DO (Set "UpdateUrl=%%B")
	call:DecodeChannelName %UpdateUrl%
	echo Updates-Channel:
	echo %ChannelName%
	echo:
	echo Updates-Url:
	echo %UpdateUrl%
	echo:
	echo Group-Policy defined UpdateBranch:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate" /v "UpdateBranch" 2^>nul') DO (Set "UpdateBranch=%%B")
	echo %UpdateBranch%
	echo ____________________________________________________________________________
	echo:
	echo Possible Office 2016 Update-Channel ID VALUES:
	echo 1 = Monthly_Channel (Retail/RTM)
	echo 2 = Insider_Channel (Office Insider FAST)
	echo 3 = Monthly_Channel_Targeted (Office Insider SLOW)
	echo 4 = Semi_Annual_Channel (Business)
	echo 5 = Semi_Annual_Channel_Targeted (Business Insider)
	echo 6 = Dogfood_DevMain_Channel (MS Internal Use Only)
	echo X = exit to Main Menu
	echo:
	set /p channeltrigger=Set New Update-Channel-ID (1,2,3,4,5,6) or X ^>
	if "%channeltrigger%" EQU "1" (
		set "latestfile=latest_Monthly_Channel_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
		set "UpdateBranch=Current"
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "2" (
		set "latestfile=latest_Insider_Channel_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
		set "UpdateBranch=InsiderFast"
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "3" (
		set "latestfile=latest_Monthly_Channel_Targeted_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
		set "UpdateBranch=FirstReleaseCurrent"
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "4" (
		set "latestfile=latest_Semi_Annual_Channel_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
		set "UpdateBranch=Deferred"
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "5" (
		set "latestfile=latest_Semi_Annual_Channel_Targeted_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
		set "UpdateBranch=FirstReleaseDeferred"
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "6" (
		set "latestfile=latest_Dogfood_DevMain_Channel_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/ea4a4090-de26-49d7-93c1-91bff9e53fc3"
		set "UpdateBranch=not set"
		goto:UpdateChannelSel
	)
	if /I "%channeltrigger%" EQU "X" (goto:Office16VnextInstall)
	goto:ChangeUpdPath
::===============================================================================================================
:UpdateChannelSel
	echo:
	set /a countx=0
	cd /D "%OfficeRToolpath%"
	for /F "tokens=*" %%a in (!latestfile!) do (
		SET /a countx=!countx! + 1
		set var!countx!=%%a
	)
	set "o16upg1build=%var1%"
	set "o16upg2build=%var2%"
	echo Manually enter any build-nummer such as %o16upg2build%(prior build)
	echo or simply press return for updating to: %o16upg1build%(newest build)
	set /p restrictbuild=Set Office update build ^>
	if "%restrictbuild%" NEQ "newest available" set "updatetoversion=updatetoversion=%restrictbuild%"
	call :DecodeChannelName %UpdateUrl%
	echo ____________________________________________________________________________
	echo:
	echo New Update-Configuration will be set to:
	echo:
	echo Distribution-Channel : %ChannelName%
	echo Update To Version    : %restrictbuild%
	echo:
	set /p installtrigger=Change Configuration and start download of new Office version (1/0)? ^>
    if "%installtrigger%" EQU "0" goto:ChangeUpdPath
    if "%installtrigger%" EQU "1" goto:ChangeUpdateConf
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
    goto:ChangeUpdPath
::===============================================================================================================
:ChangeUpdateConf
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v CDNBaseUrl /d %UpdateUrl% /f >nul 2>&1
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateUrl /d %UpdateUrl% /f >nul 2>&1
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateChannel /d %UpdateUrl% /f >nul 2>&1
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateChannelChanged /d True /f >nul 2>&1
	if "%UpdateBranch%" EQU "not set" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f >nul 2>&1
	if "%UpdateBranch%" NEQ "not set" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %UpdateBranch% /f >nul 2>&1
	reg delete HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateToVersion /f >nul 2>&1
	reg delete HKLM\Software\Microsoft\Office\ClickToRun\Updates /v UpdateToVersion /f >nul 2>&1
	if "%restrictbuild%" NEQ "newest available" (("%CommonProgramFiles%\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user %updatetoversion% updatepromptuser=True displaylevel=True)&&(goto:Office16VnextInstall))
	"%CommonProgramFiles%\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatepromptuser=True displaylevel=True >nul 2>&1
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:DecodeChannelName
	set "ChannelName=%1"
	set "ChannelName=%ChannelName:~-36%"
	if "%ChannelName%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" (set "ChannelName=Monthly_Channel (Retail/RTM)")&&(goto:eof)
	if "%ChannelName%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" (set "ChannelName=Insider (Office Insider FAST)")&&(goto:eof)
	if "%ChannelName%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" (set "ChannelName=Monthly_Channel_Targeted (Office Insider SLOW)")&&(goto:eof)
	if "%ChannelName%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" (set "ChannelName=Semi_Annual_Channel (Business)")&&(goto:eof)
	if "%ChannelName%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" (set "ChannelName=Semi_Annual_Channel_Targeted (Business Insider)")&&(goto:eof)
	if "%ChannelName%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" (set "ChannelName=Dogfood_DevMain_Channel (MS Internal Use Only)")&&(goto:eof)
	set "ChannelName=Non_Standard_Channel (Manual_Override)"
	goto:eof
::===============================================================================================================
::===============================================================================================================
:DisableTelemetry
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	echo ================== DISABLE ACQUISITION OF TELEMETRY DATA ===================
    echo ____________________________________________________________________________
	echo:
	echo Scheduler:  4 Office Telemetry related Tasks were set / changed
	schtasks /Change /TN "Microsoft\Office\Office Automatic Updates" /Disable >nul 2>&1
	schtasks /Change /TN "Microsoft\Office\OfficeTelemetryAgentFallBack2016" /Disable >nul 2>&1
	schtasks /Change /TN "Microsoft\Office\OfficeTelemetryAgentLogOn2016" /Disable >nul 2>&1
	schtasks /Change /TN "Microsoft\Office\Office ClickToRun Service Monitor" /Disable >nul 2>&1
	echo:
	echo Registry:  29 Office Telemetry related User Keys were set / changed
	REG ADD HKCU\Software\Microsoft\Office\Common\ClientTelemetry /v DisableTelemetry /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v sendcustomerdata /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\Feedback /v enabled /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\Feedback /v includescreenshot /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail /v EnableLogging /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Word\Options /v EnableLogging /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v qmenable /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v updatereliabilitydata /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\General /v shownfirstrunoptin /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\General /v skydrivesigninoption /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\ptwatson /v ptwoptin /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\Firstrun /v disablemovie /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v Enablelogging /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v EnableUpload /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v EnableFileObfuscation /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v accesssolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v olksolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v onenotesolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v pptsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v projectsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v publishersolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v visiosolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v wdsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v xlsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v agave /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v appaddins /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v comaddins /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v documentfiles /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v templatefiles /t REG_DWORD /d 1 /f >nul 2>&1
	echo:
	echo Registry:  23 Office Telemetry related Machine Group Policies were set / changed
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common /v qmenable /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common /v updatereliabilitydata /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\General /v shownfirstrunoptin /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\General /v skydrivesigninoption /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\ptwatson /v ptwoptin /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Firstrun /v disablemovie /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v Enablelogging /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v EnableUpload /t REG_DWORD /d 0 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v EnableFileObfuscation /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v accesssolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v olksolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v onenotesolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v pptsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v projectsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v publishersolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v visiosolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v wdsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v xlsolution /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v agave /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v appaddins /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v comaddins /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v documentfiles /t REG_DWORD /d 1 /f >nul 2>&1
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v templatefiles /t REG_DWORD /d 1 /f >nul 2>&1
	echo ____________________________________________________________________________
	echo:
    echo:
	timeout /t 7
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:InstallO16
	set "of16install=0"
	set "of19install=0"
	set "of36install=0"
	set "ofbsinstall=0"
	set "mo16install=0"
	set "wd16disable=0"
	set "ex16disable=0"
	set "pp16disable=0"
	set "ac16disable=0"
	set "ol16disable=0"
	set "pb16disable=0"
	set "on16disable=0"
	set "sk16disable=0"
	set "od16disable=0"
	set "pr16install=0"
	set "vi16install=0"
	set "pr19install=0"
	set "vi19install=0"
	set "wd16install=0"
	set "ex16install=0"
	set "pp16install=0"
	set "ac16install=0"
	set "ol16install=0"
	set "pb16install=0"
	set "on16install=0"
	set "sk16install=0"
	set "wd19install=0"
	set "ex19install=0"
	set "pp19install=0"
	set "ac19install=0"
	set "ol19install=0"
	set "pb19install=0"
	set "sk19install=0"
	set "installtrigger=not set"
	set "createpackage=0"
	set "productstoadd=0"
	set "excludedapps=0"
	set "productkeys=0"
	set "searchdirpattern=16.0"
:InstallO16Loop
	cls
	echo:
	echo ================== INSTALL OFFICE FULL SUITE / SINGLE APPS =================
	echo ____________________________________________________________________________
	echo:
	if "%downpath%" EQU "not set" set /p downpath=Set Office Package Download Path ^>
	set "downpath=%downpath:"=%"
	if /I "%downpath%" EQU "X" ((set "downpath=not set")&&(goto:Office16VnextInstall))
	set "downdrive=%downpath:~0,2%"
	if "%downdrive:~-1%" NEQ ":" ((echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:InstallO16Loop))
	cd /d %downdrive% >nul 2>&1
	if errorlevel 1 (echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:InstallO16Loop)
	set "downdrive=%downpath:~0,3%"
	if "%downdrive:~-1%" EQU "\" (set "downpath=%downdrive%%downpath:~3%") else (set "downpath=%downdrive:~0,2%\%downpath:~2%")
	if "%downpath:~-1%" EQU "\" set "downpath=%downpath:~0,-1%"
::===============================================================================================================
	cd /d "%downdrive%\" >nul 2>&1
	cd /d "%downpath%%" >nul 2>&1
	set /a countx=0
	echo Download-path = "%downpath%"
	echo:
	if "%searchdirpattern%" EQU "not set" ((set /p searchdirpattern=Enter search pattern or enter x to abort ^>)&&(echo:))
	if "%searchdirpattern%" EQU "not set" goto:InstallO16Loop
	if "%searchdirpattern%" EQU "x" (goto:Office16VnextInstall)
	if "%searchdirpattern%" EQU "X" (goto:Office16VnextInstall)
	echo List of available installation packages
	echo:
	echo #   Package
	for /F "tokens=*" %%a in ('dir "%downpath%\" /ad /b 2^>nul ^| findstr /i "%searchdirpattern%"') do (
		echo:
		SET /a countx=!countx! + 1
		set packagelist!countx!=%%a
		echo !countx!   %%a
	)
	if %countx% GTR 0 goto:PackageFound
	echo No install packages found!
	set "searchdirpattern=not set"
	echo:
	pause
	goto:InstallO16Loop
::===============================================================================================================
:PackageFound
	echo:
	echo:
	set /a packnum=0
	set /p packnum=Enter package number # or enter 0 for new search pattern ^>
	if /I "%packnum%" EQU "X" goto:Office16VnextInstall
	if %packnum% EQU 0 ((set "searchdirpattern=not set")&&(goto:InstallO16Loop))
	if %packnum% GTR %countx% ((set "searchdirpattern=not set")&&(goto:InstallO16Loop))
	echo:
	set "installpath=%downpath%\!packagelist%packnum%!"
	if "%installpath:~-1%" EQU "\" set "installpath=%installpath:~0,-1%"
	set countx=0
	cd /d "%installpath%"
	for /F "tokens=*" %%a in (package.info) do (
		SET /a countx=!countx! + 1
		set var!countx!=%%a
	)
	if %countx% LSS 5 (echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:Office16VnextInstall)
	set "distribchannel=%var1%"
	if "%distribchannel:~-1%" EQU " " set "distribchannel=%distribchannel:~0,-1%"
	set "o16build=%var2%"
	set "o16lang=%var3%"
	call :SetO16Language
	set "o16arch=%var4%"
	set "o16updlocid=%var5%"
	if "%winx%" EQU "win_x32" if "%o16arch%" EQU "x64" ((echo:)&&(echo ERROR: You can't install x64/64bit Office on x86/32bit Windows)&&(echo:)&&(pause)&&(goto:InstallO16))
::===============================================================================================================
:InstSuites
	cd /D "%OfficeRToolpath%"
	cls
	set "installtrigger=s"
	echo:
	echo ================== INSTALL OFFICE FULL SUITE / SINGLE APPS =================
	echo ____________________________________________________________________________
	echo:
	echo Using Office Setup Package found in:
	echo %installpath%
	echo:
	set /p installtrigger=Set Install method (OfficeClickToRun.exe = C or ODT Setup.exe = S) ^>
	echo:
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if /I "%installtrigger%" EQU "C" if "%o16arch%" EQU "x86" if "%winx%" EQU "win_x64" ((set "instmethod=XML")&&(echo 32bit Office on 64bit Windows gives install problems with "OfficeClickToRun.exe".)&&(echo Switching to OfficeDeploymentTool "setup.exe" instead.)&&(goto:SelFullSuite))
	if /I "%installtrigger%" EQU "C" ((set "instmethod=C2R")&&(echo Office is installed by using "OfficeClickToRun.exe")&&(goto:SelFullSuite))
	if /I "%installtrigger%" EQU "S" ((set "instmethod=XML")&&(echo Office is installed by using OfficeDeploymentTool "setup.exe")&&(goto:SelFullSuite))
	goto:InstSuites
::===============================================================================================================
:SelFullSuite
	echo:
	echo:
	echo Select full Office Suite for install:
	echo:
	echo 1.) Office 2016 ProfessionalPlus     2.) Office 365 ProfessionalPlus
	echo 3.) Office 365 Business              4.) Office 2016 Mondo
	echo 5.) Office 2019 ProfessionalPlus     6.) Visio 2016 / Project 2016
	echo 7.) Visio 2019 / Project 2019        0.) Single Apps Install (no full suite)
	echo:
	set /p installtrigger=Enter 1...7,0 or x to exit ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "0" goto:SingleAppsInstall
	if "%installtrigger%" EQU "1" ((set "of16install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "2" ((set "of36install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "3" ((set "ofbsinstall=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "4" ((set "mo16install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" if "%distribchannel%" EQU "Current" ((set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" if "%distribchannel%" EQU "InsiderFast" ((set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" if "%distribchannel%" EQU "FirstReleaseCurrent" ((set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" if "%distribchannel%" EQU "ManualOverride" ((set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" if "%distribchannel%" EQU "DogfoodDevMain" ((set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "6" if "%distribchannel%" EQU "Current" (goto:InstVi16Pr16)
	if "%installtrigger%" EQU "6" if "%distribchannel%" EQU "InsiderFast" (goto:InstVi16Pr16)
	if "%installtrigger%" EQU "6" if "%distribchannel%" EQU "FirstReleaseCurrent" (goto:InstVi16Pr16)
	if "%installtrigger%" EQU "6" if "%distribchannel%" EQU "ManualOverride" (goto:InstVi16Pr16)
	if "%installtrigger%" EQU "6" if "%distribchannel%" EQU "DogfoodDevMain" (goto:InstVi16Pr16)
	if "%installtrigger%" EQU "7" if "%distribchannel%" EQU "Current" (goto:InstVi19Pr19)
	if "%installtrigger%" EQU "7" if "%distribchannel%" EQU "InsiderFast" (goto:InstVi19Pr19)
	if "%installtrigger%" EQU "7" if "%distribchannel%" EQU "FirstReleaseCurrent" (goto:InstVi19Pr19)
	if "%installtrigger%" EQU "7" if "%distribchannel%" EQU "ManualOverride" (goto:InstVi19Pr19)
	if "%installtrigger%" EQU "7" if "%distribchannel%" EQU "DogfoodDevMain" (goto:InstVi19Pr19)
	if "%installtrigger%" EQU "5" (
		echo:
		echo Office 2019, Project 2019, Visio 2019 not available in "Semi-Annual" and
		echo "Semi-Annual (Targeted) distribution channnels. Choose another distribution
		echo channel for install.
		echo:
		timeout /t 7
	)
	if "%installtrigger%" EQU "7" (
		echo:
		echo Office 2019, Project 2019, Visio 2019 not available in "Semi-Annual" and
		echo "Semi-Annual (Targeted) distribution channnels. Choose another distribution
		echo channel for install.
		echo:
		timeout /t 7
	)
	goto:InstSuites
::===============================================================================================================
:SingleAppsInstall
	echo:
	set /p installtrigger=Use "2016 version" for Single App install (1/0=2019) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "0" goto:SingleApps2019Install
	if "%installtrigger%" EQU "1" goto:SingleApps2016Install
	goto:InstallO16
:SingleApps2016Install
	echo:
	set /p wd16install=Set Word 2016 Single App Install (1/0) ^>
	if /I "%wd16install%" EQU "X" goto:Office16VnextInstall
	set /p ex16install=Set Excel 2016 Single App Install (1/0) ^>
	if /I "%ex16install%" EQU "X" goto:Office16VnextInstall
	set /p pp16install=Set Powerpoint 2016 Single App Install (1/0) ^>
	if /I "%pp16install%" EQU "X" goto:Office16VnextInstall
	set /p ac16install=Set Access 2016 Single App Install (1/0) ^>
	if /I "%ac16install%" EQU "X" goto:Office16VnextInstall
	set /p ol16install=Set Outlook 2016 Single App Install (1/0) ^>
	if /I "%ol16install%" EQU "X" goto:Office16VnextInstall
	set /p pb16install=Set Publisher 2016 Single App Install (1/0) ^>
	if /I "%pb16install%" EQU "X" goto:Office16VnextInstall
	set /p on16install=Set OneNote 2016 Single App Install (1/0) ^>
	if /I "%on16install%" EQU "X" goto:Office16VnextInstall
	set /p sk16install=Set Skype For Business 2016 Single App Install (1/0) ^>
	if /I "%sk16install%" EQU "X" goto:Office16VnextInstall
	goto:InstallProVis
:SingleApps2019Install
	echo:
	set /p wd19install=Set Word 2019 Single App Install (1/0) ^>
	if /I "%wd19install%" EQU "X" goto:Office16VnextInstall
	set /p ex19install=Set Excel 2019 Single App Install (1/0) ^>
	if /I "%ex19install%" EQU "X" goto:Office16VnextInstall
	set /p pp19install=Set Powerpoint 2019 Single App Install (1/0) ^>
	if /I "%pp19install%" EQU "X" goto:Office16VnextInstall
	set /p ac19install=Set Access 2019 Single App Install (1/0) ^>
	if /I "%ac19install%" EQU "X" goto:Office16VnextInstall
	set /p ol19install=Set Outlook 2019 Single App Install (1/0) ^>
	if /I "%ol19install%" EQU "X" goto:Office16VnextInstall
	set /p pb19install=Set Publisher 2019 Single App Install (1/0) ^>
	if /I "%pb19install%" EQU "X" goto:Office16VnextInstall
	set /p sk19install=Set Skype For Business 2019 Single App Install (1/0) ^>
	if /I "%sk19install%" EQU "X" goto:Office16VnextInstall
	goto:InstallProVis
::===============================================================================================================
:InstallExclusions
	if "%mo16install%" EQU "1" ((set "of16install=0")&&(set "of19install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of16install%" EQU "1" ((set "mo16install=0")&&(set "of19install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of19install%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of36install%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of19install=0")&&(set "ofbsinstall=0"))
	if "%ofbsinstall%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of19install=0")&&(set "of36install=0"))
	echo:
	echo Full Suite Install Exclusion List - Disable not needed Office Programs
	set /p wd16disable=Disable Word Install  (1/0) ^>
	if /I "%wd16disable%" EQU "X" goto:Office16VnextInstall
	set /p ex16disable=Disable Excel Install (1/0) ^>
	if /I "%ex16disable%" EQU "X" goto:Office16VnextInstall
	set /p pp16disable=Disable Powerpoint Install (1/0) ^>
	if /I "%pp16disable%" EQU "X" goto:Office16VnextInstall
	set /p ac16disable=Disable Access Install (1/0) ^>
	if /I "%ac16disable%" EQU "X" goto:Office16VnextInstall
	set /p ol16disable=Disable Outlook Install (1/0) ^>
	if /I "%ol16disable%" EQU "X" goto:Office16VnextInstall
	set /p pb16disable=Disable Publisher Install (1/0) ^>
	if /I "%pb16disable%" EQU "X" goto:Office16VnextInstall
	set /p on16disable=Disable OneNote Install (1/0) ^>
	if /I "%on16disable%" EQU "X" goto:Office16VnextInstall
	set /p sk16disable=Disable Skype For Business Install (1/0) ^>
	if /I "%sk16disable%" EQU "X" goto:Office16VnextInstall
	set /p od16disable=Disable OneDrive For Business Install (1/0) ^>
	if /I "%od16disable%" EQU "X" goto:Office16VnextInstall
::===============================================================================================================
:InstallProVis
	echo ____________________________________________________________________________
	if "%of19install%" EQU "1" goto:InstVi19Pr19
	if "%wd19install%" EQU "1" goto:InstVi19Pr19
	if "%ex19install%" EQU "1" goto:InstVi19Pr19
	if "%pp19install%" EQU "1" goto:InstVi19Pr19 
	if "%ac19install%" EQU "1" goto:InstVi19Pr19
	if "%ol19install%" EQU "1" goto:InstVi19Pr19
	if "%pb19install%" EQU "1" goto:InstVi19Pr19
	if "%sk19install%" EQU "1" goto:InstVi19Pr19
:InstVi16Pr16
	echo:
	set /p vi16install=Set Visio 2016 Install (1/0) ^>
	set /p pr16install=Set Project 2016 Install (1/0) ^>
	goto:InstViPrEnd
:InstVi19Pr19
	echo:
	if "%distribchannel%" EQU "Current" set /p vi19install=Set Visio 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "InsiderFast" set /p vi19install=Set Visio 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "FirstReleaseCurrent" set /p vi19install=Set Visio 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "ManualOverride" set /p vi19install=Set Visio 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "DogfoodDevMain" set /p vi19install=Set Visio 2019 Install (1/0) ^>
::===============================================================================================================
	if "%distribchannel%" EQU "Current" set /p pr19install=Set Project 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "InsiderFast" set /p pr19install=Set Project 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "FirstReleaseCurrent" set /p pr19install=Set Project 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "ManualOverride" set /p pr19install=Set Project 2019 Install (1/0) ^>
	if "%distribchannel%" EQU "DogfoodDevMain" set /p pr19install=Set Project 2019 Install (1/0) ^>
::===============================================================================================================
:InstViPrEnd
	echo ____________________________________________________________________________
	echo:
::===============================================================================================================
	if "%o16updlocid%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" (echo Monthly_Channel - %o16build% -Setup-)&&(goto:PendSetupContinue)
	if "%o16updlocid%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" (echo Insider_Channel - %o16build% -Setup-)&&(goto:PendSetupContinue)
	if "%o16updlocid%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" (echo Monthly_Channel_Targeted - %o16build% -Setup-)&&(goto:PendSetupContinue)
	if "%o16updlocid%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" (echo Semi_Annual_Channel - %o16build% -Setup-)&&(goto:PendSetupContinue)
	if "%o16updlocid%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" (echo Semi_Annual_Channel_Targeted - %o16build% -Setup-)&&(goto:PendSetupContinue)
	if "%o16updlocid%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" (echo Dogfood_DevMain_Channel - %o16build% -Setup-)&&(goto:PendSetupContinue)
	echo Manual_Override %o16updlocid% - %o16build% -Setup-
::===============================================================================================================
:PendSetupContinue
	echo:
	echo The following programs are selected for install:
	echo:
	if "%wd16install%" EQU "1" goto:PendSetupSingleApp
	if "%ex16install%" EQU "1" goto:PendSetupSingleApp
	if "%pp16install%" EQU "1" goto:PendSetupSingleApp
	if "%ac16install%" EQU "1" goto:PendSetupSingleApp
	if "%ol16install%" EQU "1" goto:PendSetupSingleApp
	if "%pb16install%" EQU "1" goto:PendSetupSingleApp
	if "%on16install%" EQU "1" goto:PendSetupSingleApp
	if "%sk16install%" EQU "1" goto:PendSetupSingleApp
	if "%wd19install%" EQU "1" goto:PendSetupSingleApp
	if "%ex19install%" EQU "1" goto:PendSetupSingleApp
	if "%pp19install%" EQU "1" goto:PendSetupSingleApp
	if "%ac19install%" EQU "1" goto:PendSetupSingleApp
	if "%ol19install%" EQU "1" goto:PendSetupSingleApp
	if "%pb19install%" EQU "1" goto:PendSetupSingleApp
	if "%sk19install%" EQU "1" goto:PendSetupSingleApp
::===============================================================================================================
	if "%of16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office 2016 ProfessionalPlus" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office 2019 ProfessionalPlus" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of36install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office 365 ProfessionalPlus" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%ofbsinstall%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office 365 Business" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%mo16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Mondo 2016 Grande Suite" -foreground "Green")&&(goto:PendSetupFullSuite)
	goto:PendSetupProjectVisio
:PendSetupFullSuite
	if "%wd16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Word" -foreground "Red")
	if "%wd16disable%" EQU "0" (echo --^> Enabled:  Word)
	if "%ex16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Excel" -foreground "Red")
	if "%ex16disable%" EQU "0" (echo --^> Enabled:  Excel)
	if "%pp16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Powerpoint" -foreground "Red")
	if "%pp16disable%" EQU "0" (echo --^> Enabled:  PowerPoint)
	if "%ac16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Access" -foreground "Red")
	if "%ac16disable%" EQU "0" (echo --^> Enabled:  Access)
	if "%ol16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Outlook" -foreground "Red")
	if "%ol16disable%" EQU "0" (echo --^> Enabled:  Outlook)
	if "%pb16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Publisher" -foreground "Red")
	if "%pb16disable%" EQU "0" (echo --^> Enabled:  Publisher)
	if "%on16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: OneNote" -foreground "Red")
	if "%on16disable%" EQU "0" (echo --^> Enabled:  OneNote)
	if "%sk16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Skype For Business" -foreground "Red")
	if "%sk16disable%" EQU "0" (echo --^> Enabled:  Skype For Business)
	if "%od16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: OneDrive For Business" -foreground "Red")
	if "%od16disable%" EQU "0" (echo --^> Enabled:  OneDrive For Business)
	goto:PendSetupProjectVisio
::===============================================================================================================
:PendSetupSingleApp	
	if "%wd16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Word 2016 Single App" -foreground "Green")
	if "%ex16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Excel 2016 Single App" -foreground "Green")
	if "%pp16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "PowerPoint 2016 Single App" -foreground "Green")
	if "%ac16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Access 2016 Single App" -foreground "Green")
	if "%ol16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Outlook 2016 Single App" -foreground "Green")
	if "%pb16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Publisher 2016 Single App" -foreground "Green")
	if "%on16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "OneNote 2016 Single App" -foreground "Green")
	if "%sk16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Skype For Business 2016 Single App" -foreground "Green")
	if "%wd19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Word 2019 Single App" -foreground "Green")
	if "%ex19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Excel 2019 Single App" -foreground "Green")
	if "%pp19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "PowerPoint 2019 Single App" -foreground "Green")
	if "%ac19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Access 2019 Single App" -foreground "Green")
	if "%ol19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Outlook 2019 Single App" -foreground "Green")
	if "%pb19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Publisher 2019 Single App" -foreground "Green")
	if "%sk19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Skype For Business 2019 Single App" -foreground "Green")
::===============================================================================================================
:PendSetupProjectVisio
	if "%pr16install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Project 2016 Professional" -foreground "Green")
	if "%vi16install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Visio 2016 Professional" -foreground "Green")
	if "%pr19install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Project 2019 Professional" -foreground "Green")
	if "%vi19install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Visio 2019 Professional" -foreground "Green")
::===============================================================================================================
	echo:
	echo Language:     %o16lang%   (fixed - matches Office download package)
	echo Architecture: %o16arch%     (fixed - matches Office download package)
	echo ____________________________________________________________________________
	echo:
	set /p installtrigger=Start local install now (1/0) or Create Install Package (C) ? ^>
	if "%installtrigger%" EQU "0" goto:InstallO16
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if /I "%installtrigger%" EQU "C" set "createpackage=1"
	if "%installtrigger%" EQU "1" goto:OfficeC2RXMLInstall
	if "%createpackage%" EQU "1" goto:OfficeC2RXMLInstall
	goto:InstallO16
::===============================================================================================================
:OfficeC2RXMLInstall
	cls
    echo:
	echo ================= INSTALL OFFICE FULL SUITE / SINGLE APPS ==================
	echo ____________________________________________________________________________
    echo:
    if "%o16arch%" EQU "x64" (set "o16a=64") else (set "o16a=32")
	if "%instmethod%" EQU "XML" echo Creating setup files "setup.exe", "configure%o16a%.xml" and "start_setup.cmd"
	if "%instmethod%" EQU "C2R" echo Creating setup file "start_setup.cmd"
	echo:
    echo in Installpath: "%installpath%"
    echo:
	if "%instmethod%" EQU "XML" set "oxml=%installpath%\configure%o16a%.xml"
	if "%instmethod%" EQU "XML" copy "%OfficeRToolpath%\OfficeFixes\setup.exe" "%installpath%" /Y >nul 2>&1
	if "%instmethod%" EQU "XML" (set "channel= channel="%distribchannel%"")
	if "%instmethod%" EQU "C2R" if exist "%installpath%\setup.exe" del /s /q "%installpath%\setup.exe" >nul 2>&1
	if "%distribchannel%" EQU "ManualOverride" (set "channel=")
	if "%distribchannel%" EQU "DogfoodDevMain" (set "channel=")
	if exist "%installpath%\configure*.xml" del /s /q "%installpath%\configure*.xml" >nul 2>&1
	set "obat=%installpath%\start_setup.cmd"
	copy "%OfficeRToolpath%\OfficeFixes\start_setup.cmd" "%installpath%" /Y >nul 2>&1
	if "%instmethod%" EQU "C2R" goto:CreateC2RConfig
	if "%instmethod%" EQU "XML" goto:CreateXMLConfig
	goto:InstallO16
::===============================================================================================================
:CreateXMLConfig
    echo ^<Configuration^> >"%oxml%"
	echo     ^<Add DownloadPath="http://officecdn.microsoft.com/pr/%o16updlocid%" OfficeClientEdition="%o16a%" Version="%o16build%"%channel% ^> >>"%oxml%"
	if "%mo16install%" EQU "1" (
        echo         ^<Product ID="MondoRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%" EQU "1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%" EQU "1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%" EQU "1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%" EQU "1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%" EQU "1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%" EQU "1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%" EQU "1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%" EQU "1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="OneDrive"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%of16install%" EQU "1" (
        echo         ^<Product ID="ProPlusRetail"^> >>"%oxml%"
		echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%" EQU "1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%" EQU "1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%" EQU "1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%" EQU "1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%" EQU "1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%" EQU "1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%" EQU "1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%" EQU "1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="OneDrive"/^> >>"%oxml%"
		echo         ^</Product^> >>"%oxml%"
	)
    if "%of19install%" EQU "1" (
        echo         ^<Product ID="ProPlus2019Retail"^> >>"%oxml%"
		echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%" EQU "1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%" EQU "1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%" EQU "1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%" EQU "1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%" EQU "1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%" EQU "1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%" EQU "1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%" EQU "1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="OneDrive"/^> >>"%oxml%"
		echo         ^</Product^> >>"%oxml%"
	)
    if "%of36install%" EQU "1" (
        echo         ^<Product ID="O365ProPlusRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%" EQU "1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%" EQU "1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%" EQU "1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%" EQU "1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%" EQU "1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%" EQU "1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%" EQU "1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%" EQU "1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="OneDrive"/^> >>"%oxml%"
		echo         ^</Product^> >>"%oxml%"
	)
    if "%ofbsinstall%" EQU "1" (
        echo         ^<Product ID="O365BusinessRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%" EQU "1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%" EQU "1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%" EQU "1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%" EQU "1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%" EQU "1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%" EQU "1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%" EQU "1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%" EQU "1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%od16disable%" EQU "1" echo             ^<ExcludeApp ID="OneDrive"/^> >>"%oxml%"
		echo         ^</Product^> >>"%oxml%"
	)
    if "%pr16install%" EQU "1" (
        echo         ^<Product ID="ProjectProRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pr19install%" EQU "1" (
        echo         ^<Product ID="ProjectPro2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%vi16install%" EQU "1" (
        echo         ^<Product ID="VisioProRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%vi19install%" EQU "1" (
        echo         ^<Product ID="VisioPro2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%wd16install%" EQU "1" (
        echo         ^<Product ID="WordRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%wd19install%" EQU "1" (
        echo         ^<Product ID="Word2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ex16install%" EQU "1" (
        echo         ^<Product ID="ExcelRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ex19install%" EQU "1" (
        echo         ^<Product ID="Excel2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pp16install%" EQU "1" (
        echo         ^<Product ID="PowerPointRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pp19install%" EQU "1" (
        echo         ^<Product ID="PowerPoint2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ac16install%" EQU "1" (
        echo         ^<Product ID="AccessRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ac19install%" EQU "1" (
        echo         ^<Product ID="Access2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ol16install%" EQU "1" (
        echo         ^<Product ID="OutlookRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ol19install%" EQU "1" (
        echo         ^<Product ID="Outlook2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
	if "%pb16install%" EQU "1" (
        echo         ^<Product ID="PublisherRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pb19install%" EQU "1" (
        echo         ^<Product ID="Publisher2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%on16install%" EQU "1" (
        echo         ^<Product ID="OneNoteRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%sk16install%" EQU "1" (
        echo         ^<Product ID="SkypeForBusinessRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%sk19install%" EQU "1" (
        echo         ^<Product ID="SkypeForBusiness2019Retail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    echo     ^</Add^> >>"%oxml%"
	echo     ^<Property Name="ForceAppsShutdown" Value="True" /^> >>"%oxml%"
	echo     ^<Property Name="PinIconsToTaskbar" Value="False" /^> >>"%oxml%"
    echo     ^<Display Level="Full" AcceptEula="True" /^> >>"%oxml%"
	echo     ^<Updates Enabled="True" UpdatePath="http://officecdn.microsoft.com/pr/%o16updlocid%"%channel% /^> >>"%oxml%"
	echo ^</Configuration^> >>"%oxml%"
	goto:CreateStartSetupBatch
::===============================================================================================================
::===============================================================================================================
:CreateC2RConfig
	if "%mo16install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|MondoRetail.16_%%instlang%%_x-none"
		set "productID=MondoRetail"
		)
    if "%of16install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|ProPlusRetail.16_%%instlang%%_x-none"
		set "productID=ProPlusRetail"
		)
    if "%of19install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|ProPlus2019Retail.16_%%instlang%%_x-none"
		set "productID=ProPlus2019Retail"
		)
	if "%of36install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|O365ProPlusRetail.16_%%instlang%%_x-none"
		set "productID=O365ProPlusRetail"
		)
	if "%ofbsinstall%" EQU "1" (
        set "productstoadd=!productstoadd!^^|O365BusinessRetail.16_%%instlang%%_x-none"
        set "productID=O365BusinessRetail"
		)
		if "%wd16disable%" EQU "1" set "excludedapps=!excludedapps!,word"
		if "%ex16disable%" EQU "1" set "excludedapps=!excludedapps!,excel"
		if "%pp16disable%" EQU "1" set "excludedapps=!excludedapps!,powerpoint"
		if "%ac16disable%" EQU "1" set "excludedapps=!excludedapps!,access"
		if "%ol16disable%" EQU "1" set "excludedapps=!excludedapps!,outlook"
		if "%pb16disable%" EQU "1" set "excludedapps=!excludedapps!,publisher"
		if "%on16disable%" EQU "1" set "excludedapps=!excludedapps!,onenote"
		if "%sk16disable%" EQU "1" set "excludedapps=!excludedapps!,lync"
		if "%od16disable%" EQU "1" set "excludedapps=!excludedapps!,groove"
		if "%od16disable%" EQU "1" set "excludedapps=!excludedapps!,onedrive"
    )
	if "!excludedapps:~0,2!" EQU "0," (set "excludedapps=%productID%.excludedapps.16^=!excludedapps:~2!") else (set "excludedapps=")
::===============================================================================================================		
    if "%pr16install%" EQU "1" set "productstoadd=!productstoadd!^^|ProjectProRetail.16_%%instlang%%_x-none"
    if "%vi16install%" EQU "1" set "productstoadd=!productstoadd!^^|VisioProRetail.16_%%instlang%%_x-none"
    if "%pr19install%" EQU "1" set "productstoadd=!productstoadd!^^|ProjectPro2019Retail.16_%%instlang%%_x-none"
    if "%vi19install%" EQU "1" set "productstoadd=!productstoadd!^^|VisioPro2019Retail.16_%%instlang%%_x-none"
::===============================================================================================================
    if "%wd16install%" EQU "1" set "productstoadd=!productstoadd!^^|WordRetail.16_%%instlang%%_x-none"
	if "%wd19install%" EQU "1" set "productstoadd=!productstoadd!^^|Word2019Retail.16_%%instlang%%_x-none"
    if "%ex16install%" EQU "1" set "productstoadd=!productstoadd!^^|ExcelRetail.16_%%instlang%%_x-none"
    if "%ex19install%" EQU "1" set "productstoadd=!productstoadd!^^|Excel2019Retail.16_%%instlang%%_x-none"
	if "%pp16install%" EQU "1" set "productstoadd=!productstoadd!^^|PowerPointRetail.16_%%instlang%%_x-none"
    if "%pp19install%" EQU "1" set "productstoadd=!productstoadd!^^|PowerPoint2019Retail.16_%%instlang%%_x-none"
	if "%ac16install%" EQU "1" set "productstoadd=!productstoadd!^^|AccessRetail.16_%%instlang%%_x-none"
    if "%ac19install%" EQU "1" set "productstoadd=!productstoadd!^^|Access2019Retail.16_%%instlang%%_x-none"
    if "%ol16install%" EQU "1" set "productstoadd=!productstoadd!^^|OutlookRetail.16_%%instlang%%_x-none"
    if "%ol19install%" EQU "1" set "productstoadd=!productstoadd!^^|Outlook2019Retail.16_%%instlang%%_x-none"
	if "%pb16install%" EQU "1" set "productstoadd=!productstoadd!^^|PublisherRetail.16_%%instlang%%_x-none"
    if "%pb19install%" EQU "1" set "productstoadd=!productstoadd!^^|Publisher2019Retail.16_%%instlang%%_x-none"
	if "%on16install%" EQU "1" set "productstoadd=!productstoadd!^^|OneNoteRetail.16_%%instlang%%_x-none"
    if "%sk16install%" EQU "1" set "productstoadd=!productstoadd!^^|SkypeForBusinessRetail.16_%%instlang%%_x-none"
	if "%sk19install%" EQU "1" set "productstoadd=!productstoadd!^^|SkypeForBusiness2019Retail.16_%%instlang%%_x-none"
::===============================================================================================================
:CreateStartSetupBatch
	if "%distribchannel%" EQU "Current" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "InsiderFast" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "FirstReleaseCurrent" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "Deferred" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "FirstReleaseDeferred" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "DogfoodDevMain" (
		echo :: Remove Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "ManualOverride" (
		echo :: Remove Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f ^>nul 2^>^&1 >>"%obat%"
	)
	echo ^:^:=============================================================================================================== >>"%obat%"
	if "%instmethod%" EQU "C2R" echo start "" /MIN "%%CommonProgramFiles%%\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" deliverymechanism=%%instupdlocid%% platform=%%instarch1%% productreleaseid=none forcecentcheck= culture=%%instlang%% defaultplatform=False storeid= lcid=%%instlcid%% b= forceappshutdown=True piniconstotaskbar=False scenariosubtype=ODT scenario=unknown updatesenabled.16=True acceptalleulas.16=True updatebaseurl.16=http://officecdn.microsoft.com/pr/%%instupdlocid%% cdnbaseurl.16=http://officecdn.microsoft.com/pr/%%instupdlocid%% version.16=%%instversion%% mediatype.16=Local baseurl.16=%%installfolder%% sourcetype.16=Local flt.downloadappvcab=unknown flt.useclientcabmanager=unknown flt.useexptransportinplacepl=unknown flt.useaddons=unknown flt.useofficehelperaddon=unknown flt.useonedriveclientaddon=unknown productstoadd=!productstoadd:~3! !excludedapps! >>"%obat%"
	if "%instmethod%" EQU "XML" echo start "" /MIN setup.exe /configure configure%%instarch2%%.xml >>"%obat%"
	echo exit >>"%obat%"
	echo ^:^:=============================================================================================================== >>"%obat%"
	if "%createpackage%" EQU "1" goto:InstDone
::===============================================================================================================
	if "%distribchannel%" EQU "Current" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f >nul 2>&1
	if "%distribchannel%" EQU "InsiderFast" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f >nul 2>&1
	if "%distribchannel%" EQU "FirstReleaseCurrent" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f >nul 2>&1
	if "%distribchannel%" EQU "Deferred" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f >nul 2>&1
	if "%distribchannel%" EQU "FirstReleaseDeferred" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f >nul 2>&1
	if "%distribchannel%" EQU "DogfoodDevMain" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f >nul 2>&1
	if "%distribchannel%" EQU "ManualOverride" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f >nul 2>&1
	cd /D "%installpath%"
	start "" /MIN "%obat%"
::===============================================================================================================
:InstDone
	echo ____________________________________________________________________________
    echo:
	echo:
	timeout /t 7
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckOfficeApplications
	set "_ProPlusRetail=NO"
	set "_ProPlus2019Retail=NO"
	set "_ProPlus2019Volume=NO"
	set "_O365ProPlusRetail=NO"
	set "_O365BusinessRetail=NO"
	set "_MondoRetail=NO"
	set "_StandardRetail=NO"
	set "_ProjectProRetail=NO"
	set "_VisioProRetail=NO"
	set "_ProjectPro2019Retail=NO"
	set "_VisioPro2019Retail=NO"
	set "_WordRetail=NO"
	set "_ExcelRetail=NO"
	set "_PowerPointRetail=NO"
	set "_AccessRetail=NO"
	set "_OutlookRetail=NO"
	set "_PublisherRetail=NO"
	set "_OneNoteRetail=NO"
	set "_SkypeForBusinessRetail=NO"
	set "_Word2019Retail=NO"
	set "_Excel2019Retail=NO"
	set "_PowerPoint2019Retail=NO"
	set "_Access2019Retail=NO"
	set "_Outlook2019Retail=NO"
	set "_Publisher2019Retail=NO"
	set "_SkypeForBusiness2019Retail=NO"
	set "_Word2019Volume=NO"
	set "_Excel2019Volume=NO"
	set "_PowerPoint2019Volume=NO"
	set "_Access2019Volume=NO"
	set "_Outlook2019Volume=NO"
	set "_Publisher2019Volume=NO"
	set "_SkypeForBusiness2019Volume=NO"
	set "_UWPappINSTALLED=NO"
	set "_AppxWinword=NO"
	set "_AppxExcel=NO"
	set "_AppxPowerPoint=NO"
	set "_AppxAccess=NO"
	set "_AppxPublisher=NO"
	set "_AppxOutlook=NO"
	set "_AppxSkypeForBusiness=NO"
	set "_AppxOneNote=NO"
	set "_AppxVisio=NO"
	set "_AppxProject=NO"
	set "ProPlusVLFound=NO"
	set "StandardVLFound=NO"
	set "ProjectProVLFound=NO"
	set "VisioProVLFound=NO"
	set "installpath16=not set"
	set "officepath3=not set"
	set "o16version=not set"
	set "o16arch=not set"
	reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOffice16C2R
	reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOfficeVL32onW64
	reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOfficeVL32W32orVL64W64
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msosync.exe" >nul 2>&1
	if %errorlevel% EQU 0 ((set "_UWPappINSTALLED=YES")&&(goto:CheckAppxOffice16UWP))
	(echo:) && (echo Supported Office 2016/2019 product not found) && (echo:) && (pause)
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckAppxOffice16UWP
	for /F "tokens=9 delims=\_() " %%A IN ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msosync.exe" /ve 2^>nul') DO (set "o16version=%%A")
	for /F "tokens=4" %%A IN ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msosync.exe" /ve 2^>nul') DO (set "installpath16=%%A")
	set "installpath16=C:\Program !installpath16!"
	set "installpath16=!installpath16:~0,-21!"
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\winword.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxWinword=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxExcel=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxPowerPoint=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msaccess.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxAccess=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\mspub.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxPublisher=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxOutlook=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\lync.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxSkypeForBusiness=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\onenote.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxOneNote=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\visio.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxVisio=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\winproj.exe" >nul 2>&1
	if %errorlevel% EQU 0 (set "_AppxProject=YES")
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CheckOffice16C2R
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "Platform" 2^>nul') DO (set "o16arch=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" 2^>nul') DO (Set "installpath16=%%B")
	set "officepath3=%installpath16%\Office16"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "ProductReleaseIds" 2^>nul') DO (Set "Office16AppsInstalled=%%B")
	for /F "tokens=1,2,3,4,5,6,7,8,9,10,11,12,13 delims=," %%A IN ("%Office16AppsInstalled%") DO (
	set "_%%A=YES"
	set "_%%B=YES"
	set "_%%C=YES"
	set "_%%D=YES"
	set "_%%E=YES"
	set "_%%F=YES"
	set "_%%G=YES"
	set "_%%H=YES"
	set "_%%I=YES"
	set "_%%J=YES"
	set "_%%K=YES"
	set "_%%L=YES"
	set "_%%M=YES"
	)
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs" /v "ActiveConfiguration" 2^>nul') DO (set "o16activeconf=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\%o16activeconf%" /v "Modifier" 2^>nul') DO (set "o16version=%%B")
	set "o16version=%o16version:~0,16%"
	if "%o16version:~15,1%" EQU "|" (set "o16version=%o16version:~0,14%")
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CheckOfficeVL32onW64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES)&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES)&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_StandardRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_ProjectProRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_VisioProRetail%" EQU "YES" goto:OfficeVL32onW64Path
	goto:Office16VnextInstall
::===============================================================================================================
:OfficeVL32onW64Path
	set "o16arch=x86"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") >nul 2>&1
	set "officepath3=%installpath16%"
	set "checkversionpath=%CommonProgramFiles%"
	set "checkversionpath=%checkversionpath:\=\\%"
	for /F "tokens=2,* delims==" %%A IN ('"wmic datafile where name='%checkversionpath%\\Microsoft Shared\\OFFICE16\\Mso20win32client.dll' get version /format:list" 2^>nul') do set o16version=%%A
	goto:eof
::===============================================================================================================
:CheckOfficeVL32W32orVL64W64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES)&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES)&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32V64Path)
	if "%_StandardRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32V64Path)
	if "%_ProjectProRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	if "%_VisioProRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES)&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES)&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_StandardRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_ProjectProRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_VisioProRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	goto:Office16VnextInstall
::===============================================================================================================
:OfficeVL32VL64Path
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") >nul 2>&1
	set "officepath3=%installpath16%"
	set "checkversionpath=%CommonProgramFiles%"
	set "checkversionpath=%checkversionpath:\=\\%"
	for /F "tokens=2,* delims==" %%A IN ('"wmic datafile where name='%checkversionpath%\\Microsoft Shared\\Office16\\Mso20win32client.dll' get version /format:list" 2^>nul') do set o16version=%%A
	goto:eof
::===============================================================================================================
::===============================================================================================================
:Convert16Activate
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	echo ================== CONVERT OFFICE RETAIL TO VOLUME =========================
    echo ____________________________________________________________________________
	echo:
	echo Installation path:
	echo "%installpath16%"
	echo ____________________________________________________________________________
	echo:
	echo Office Suites:
	set /a countx=0
	echo:
	if "%_ProPlusRetail%" EQU "YES" ((echo Office 2016 ProfessionalPlus              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProPlus2019Retail%" EQU "YES" ((echo Office 2019 ProfessionalPlus              = "FOUND")&&(set /a countx=!countx! + 1)) else if "%_ProPlus2019Volume%" EQU "YES" ((echo Office 2019 ProfessionalPlus              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_StandardRetail%" EQU "YES" ((echo Office 2016 Standard                      = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365ProPlusRetail%" EQU "YES" ((echo Office 365 ProfessionalPlus               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365BusinessRetail%" EQU "YES" ((echo Office 365 Business                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_MondoRetail%" EQU "YES" ((echo Office Mondo Grande Suite                 = "FOUND")&&(set /a countx=!countx! + 1))
	if !countx! EQU 0 (echo Office Full Suite installation            = "NOT FOUND")
	echo ____________________________________________________________________________
	echo:
	echo Additional Apps:
	set /a countx=0
	if "%_VisioProRetail%" EQU "YES" ((echo:)&&(echo VisioPro 2016                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxVisio%" EQU "YES" ((echo:)&&(echo VisioPro 2016 UWP Appx Desktop App        = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2019Retail%" EQU "YES" ((echo:)&&(echo VisioPro 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2019Volume%" EQU "YES" ((echo:)&&(echo VisioPro 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProjectProRetail%" EQU "YES" ((echo:)&&(echo ProjectPro 2016                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_AppxProject%" EQU "YES" ((echo:)&&(echo ProjectPro 2016 UWP Appx Desktop App      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2019Retail%" EQU "YES" ((echo:)&&(echo ProjectPro 2019                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2019Volume%" EQU "YES" ((echo:)&&(echo ProjectPro 2019                           = "FOUND")&&(set /a countx=!countx! + 2))
	if !countx! EQU 0 ((echo:)&&(echo VisioPro and ProjectPro Installation      = "NOT FOUND"))
	if !countx! EQU 1 ((echo:)&&(echo ProjectPro installation                   = "NOT FOUND"))
	if !countx! EQU 2 ((echo:)&&(echo VisioPro installation                     = "NOT FOUND"))
	echo ____________________________________________________________________________
	echo:
	echo Single Apps:
	set /a countx=0
	if "%_WordRetail%" EQU "YES" ((echo:)&&(echo Word 2016                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ExcelRetail%" EQU "YES" ((echo:)&&(echo Excel 2016                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPointRetail%" EQU "YES" ((echo:)&&(echo PowerPoint 2016                           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AccessRetail%" EQU "YES" ((echo:)&&(echo Access 2016                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OutlookRetail%" EQU "YES" ((echo:)&&(echo Outlook 2016                              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PublisherRetail%" EQU "YES" ((echo:)&&(echo Publisher 2016                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OneNoteRetail%" EQU "YES" ((echo:)&&(echo OneNote 2016                              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_SkypeForBusinessRetail%" EQU "YES" ((echo:)&&(echo Skype 2016                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxWinword%" EQU "YES" ((echo:)&&(echo Word 2016 UWP Appx Desktop App            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxExcel%" EQU "YES" ((echo:)&&(echo Excel 2016 UWP Appx Desktop App           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxPowerPoint%" EQU "YES" ((echo:)&&(echo PowerPoint 2016 UWP Appx Desktop App      = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxAccess%" EQU "YES" ((echo:)&&(echo Access 2016 UWP Appx Desktop App          = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxOutlook%" EQU "YES" ((echo:)&&(echo Outlook 2016 UWP Appx Desktop App         = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxPublisher%" EQU "YES" ((echo:)&&(echo Publisher 2016 UWP Appx Desktop App       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxOneNote%" EQU "YES" ((echo:)&&(echo OneNote 2016 UWP Appx Desktop App         = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxSkypeForBusiness%" EQU "YES" ((echo:)&&(echo Skype 2016 UWP Appx Desktop App           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2019Retail%" EQU "YES" ((echo:)&&(echo Word 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2019Retail%" EQU "YES" ((echo:)&&(echo Excel 2019                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2019Retail%" EQU "YES" ((echo:)&&(echo PowerPoint 2019                           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2019Retail%" EQU "YES" ((echo:)&&(echo Access 2019                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2019Retail%" EQU "YES" ((echo:)&&(echo Outlook 2019                              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2019Retail%" EQU "YES" ((echo:)&&(echo Publisher 2019                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_SkypeForBusiness2019Retail%" EQU "YES" ((echo:)&&(echo Skype 2019                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2019Volume%" EQU "YES" ((echo:)&&(echo Word 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2019Volume%" EQU "YES" ((echo:)&&(echo Excel 2019                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2019Volume%" EQU "YES" ((echo:)&&(echo PowerPoint 2019                           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2019Volume%" EQU "YES" ((echo:)&&(echo Access 2019                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2019Volume%" EQU "YES" ((echo:)&&(echo Outlook 2019                              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2019Volume%" EQU "YES" ((echo:)&&(echo Publisher 2019                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_SkypeForBusiness2019Volume%" EQU "YES" ((echo:)&&(echo Skype 2019                                     = "FOUND")&&(set /a countx=!countx! + 1))
	if !countx! EQU 0 ((echo:)&&(echo Single Apps installation                  = "NOT FOUND"))
	echo ____________________________________________________________________________
	echo:
	echo:
	pause
::===============================================================================================================
	cls
	set "VolumeMode=0"
	set "CleanupMode=0"
	echo:
	echo ================== CHECKING INSTALLED OFFICE LICENSES ======================
	echo ____________________________________________________________________________
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 19" 1>nul && set "VolumeMode=1"
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%GRACE%%') get Name /format:list 2>nul | findstr /i /C:"Office 19" 1>nul && set "CleanupMode=1"
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "VolumeMode=1"
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%GRACE%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if "%CleanupMode%" EQU "0" if "%VolumeMode%" EQU "1" goto:Skip_Cleanup
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 19" 1>nul && set "VolumeMode=1"
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%GRACE%%') get Name /format:list 2>nul | findstr /i /C:"Office 19" 1>nul && set "CleanupMode=1"
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "VolumeMode=1"
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%GRACE%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if "%CleanupMode%" EQU "0" if "%VolumeMode%" EQU "1" goto:Skip_Cleanup
::===============================================================================================================
	echo:
	echo:
	echo ==== CLEANUP (Removing Office Retail-/Trial-/Grace-Keys and -Licenses) =====
	echo ____________________________________________________________________________
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -PKey
	echo ____________________________________________________________________________
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -Licenses
	echo ____________________________________________________________________________
	echo:
	echo:
::===============================================================================================================
::	WIN10 Insider/Preview activation failure workaround
	if exist "%windir%\System32\spp\store_test\2.0\tokens.dat"	(
		echo:
		echo ____________________________________________________________________________
		echo:
		echo Windows 10 Insider detected. Workaround for KMS activation needed^!
		echo:
		echo Stopping license service "sppsvc"...
		call :StopService "sppsvc"
		echo:
		echo Deleting license file "tokens.dat"...
		del "%windir%\System32\spp\store_test\2.0\tokens.dat" >nul 2>&1
		echo:
		echo Starting license service "sppsvc"...
		call :StartService "sppsvc"
		echo:
		echo Recreating license file "tokens.dat". Please wait...
		slmgr //B /ato
		echo ____________________________________________________________________________
		echo:
		echo:
	)
	timeout /t 7
::===============================================================================================================
	cls
	echo ===================== CONVERTING OFFICE RETAIL TO VOLUME ===================
	echo ____________________________________________________________________________
	echo:
::===============================================================================================================
	call :Office16ConversionLoop
::===============================================================================================================	
::	During conversion from Retail to Volume Office Client name gets Office 2016 ProfessionalPlus branding
::	Change (if installed) Office 365 client branding name from "Office 2016 ProfessionalPlus" back to "Office 365 ProPlus"
	del "%TEMP%\ONAME_CHANGE*.REG" >nul 2>&1
	if "%_O365ProPlusRetail%" EQU "YES" if "%winx%" EQU "win_x32" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Registration\{3AD61E22-E4FE-497F-BDB1-3E51BD872173} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if "%_O365ProPlusRetail%" EQU "YES" if "%winx%" EQU "win_x64" if "%o16arch%" EQU "x86" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Registration\{3AD61E22-E4FE-497F-BDB1-3E51BD872173} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if "%_O365ProPlusRetail%" EQU "YES" if "%winx%" EQU "win_x64" if "%o16arch%" EQU "x64" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Registration\{3AD61E22-E4FE-497F-BDB1-3E51BD872173} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if exist "%TEMP%\ONAME_CHANGE1.REG"	(
	powershell -noprofile -command "& {Get-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE1.REG" | ForEach-Object { $_ -replace '3AD61E22-E4FE-497F-BDB1-3E51BD872173', '9CAABCCB-61B1-4B4B-8BEC-D10A3C3AC2CE' } | Set-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE2.REG"}" >nul 2>&1
	reg import "%TEMP%\ONAME_CHANGE2.REG" >nul 2>&1
	del "%TEMP%\ONAME_CHANGE*.REG" >nul 2>&1
	)
::	Change (if installed) Office 365 client branding name from "Office 2016 ProfessionalPlus" back to "Office 365 Business"
	if "%_O365BusinessRetail%" EQU "YES" if "%winx%" EQU "win_x32" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Registration\{3D0631E3-1091-416D-92A5-42F84A86D868} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if "%_O365BusinessRetail%" EQU "YES" if "%winx%" EQU "win_x64" if "%o16arch%" EQU "x86" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Registration\{3D0631E3-1091-416D-92A5-42F84A86D868} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if "%_O365BusinessRetail%" EQU "YES" if "%winx%" EQU "win_x64" if "%o16arch%" EQU "x64" (reg export HKLM\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Registration\{3D0631E3-1091-416D-92A5-42F84A86D868} "%TEMP%\ONAME_CHANGE1.REG" /Y) >nul 2>&1
	if exist "%TEMP%\ONAME_CHANGE1.REG"	(
	powershell -noprofile -command "& {Get-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE1.REG" | ForEach-Object { $_ -replace '3D0631E3-1091-416D-92A5-42F84A86D868', '9CAABCCB-61B1-4B4B-8BEC-D10A3C3AC2CE' } | Set-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE2.REG"}" >nul 2>&1
	reg import "%TEMP%\ONAME_CHANGE2.REG" >nul 2>&1
	del "%TEMP%\ONAME_CHANGE*.REG" >nul 2>&1
	)
::	Change name of installed Office 2019 Retail products from Retail to Volume
	if "%_Word2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_Excel2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_PowerPoint2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_Access2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_Outlook2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_Publisher2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_SkypeForBusiness2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_ProPlus2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_VisioPro2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if "%_ProjectPro2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y >nul 2>&1
	if exist "%TEMP%\ONAME_CHANGE1.REG"	(
	powershell -noprofile -command "& {Get-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE1.REG" | ForEach-Object { $_ -replace '2019Retail', '2019Volume' } | Set-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE2.REG"}" >nul 2>&1
	reg delete HKLM\Software\Microsoft\Office\ClickToRun\Configuration /f >nul 2>&1
	reg import "%TEMP%\ONAME_CHANGE2.REG" >nul 2>&1
	del "%TEMP%\ONAME_CHANGE*.REG" >nul 2>&1
	)
::===============================================================================================================
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:Skip_Cleanup
	echo:
	echo All OK. No Conversion or Cleanup required.
	echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:StartKMSActivation
	set "kmstrigger=1"
	set "OfficeRToolKMS=YES"
	set "ExtLocKMS=O"
	set /a "num1=%random% %% 50+30"
	set /a "num2=%random% %% 186+20"
	set "KMSHostIP=10.%num1%.3.%num2%"
	if %win% LEQ 9200 set KMSHostIP=127.0.0.2
	set "KMSPort=16880"
	set "AI=120"
	set "RI=10080"
	set WinDefActive=NO
	sc query WinDefend | find "RUNNING" >nul 2>&1
	if "%errorlevel%" EQU "0" set WinDefActive=YES
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	echo ================= OFFICE PRODUCTS KMS ACTIVATION ===========================
	echo ____________________________________________________________________________
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Activation" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Classes\AppID\slui.exe" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
::===============================================================================================================
	echo:
	echo: 
	set /p kmstrigger=Use internal OfficeRTool KMS server (1/0)? ^>
    if /I "%kmstrigger%" EQU "X" goto:Office16VnextInstall
	if "%kmstrigger%" EQU "0" goto:ChangeKMSParameter
    if "%kmstrigger%" EQU "1" goto:LocalKMS2
    goto:StartKMSActivation
::===============================================================================================================
:ChangeKMSParameter
	set "OfficeRToolKMS=NO"
	echo:
	set /p ExtLocKMS=Use (E)xternal KMS on WAN/LAN or other (L)ocal KMS server (E/L)? ^>
	if /I "%ExtLocKMS%" EQU "X" goto:Office16VnextInstall
	if /I "%ExtLocKMS%" EQU "E" goto:KMSExtern1
	if /I "%ExtLocKMS%" EQU "L" goto:LocalKMS1
	goto:StartKMSActivation
::===============================================================================================================
:LocalKMS1
	echo:
	set /p KMSHostIP=Set KMS host IP ^>
	echo:
	set /p KMSPort=Set KMS host PORT ^>
	echo:
:LocalKMS2
	call :StopService "sppsvc"
	if defined ospsversion (call :StopService "osppsvc")
::===============================================================================================================	
	call :CreateKMSEnvironment
::===============================================================================================================	
	call :StartService "sppsvc"
	if defined ospsversion (call :StartService "osppsvc")
	if "%WinDefActive%" EQU "NO" goto:KMSExtern2
::===============================================================================================================	
:: Add exclusions for KMS activation files to Windows 8.1 and Windows 10 Defender exclusion list
	if %win% GEQ 9600 (
		(echo:)&&(echo Adding KMS activation file exclusions to Windows 10 Defender Exclusion List)
		set "DefExclusion="%OfficeRToolpath%\OfficeFixes\%winx%\FakeClient.exe"
		powershell -noprofile -command Add-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert.dll"
		powershell -noprofile -command Add-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		if "%winx%" EQU "win_x32" set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert32.sys"
		if "%winx%" EQU "win_x64" set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert64.sys"
		powershell -noprofile -command Add-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\vlmcsd.exe"
		powershell -noprofile -command Add-MpPreference -Force -ExclusionPath "$env:DefExclusion"
	)
	goto:KMSExtern2
::===============================================================================================================
:KMSExtern1
	echo:
	set /p KMSHostIP=Set KMS host IP ^>
	echo:
	set /p KMSPort=Set KMS host PORT ^>
	echo:
:KMSExtern2
	echo ____________________________________________________________________________
	if "%_ProPlusRetail%" EQU "YES" call :Office16Activate d450596f-894d-49e0-966a-fd39ed4c4c64
	if "%_ProPlus2019Retail%" EQU "YES" call :Office16Activate 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
	if "%_ProPlus2019Volume%" EQU "YES" call :Office16Activate 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
	if "%_O365ProPlusRetail%" EQU "YES" call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
	if "%_O365BusinessRetail%" EQU "YES" call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
	if "%_MondoRetail%" EQU "YES" call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
	if "%_UWPappINSTALLED%" EQU "YES" call :Office16Activate d450596f-894d-49e0-966a-fd39ed4c4c64
	if "%_StandardRetail%" EQU "YES" call :Office16Activate dedfa23d-6ed1-45a6-85dc-63cae0546de6
	if "%_WordRetail%" EQU "YES" call :Office16Activate bb11badf-d8aa-470e-9311-20eaf80fe5cc
	if "%_ExcelRetail%" EQU "YES" call :Office16Activate c3e65d36-141f-4d2f-a303-a842ee756a29
	if "%_PowerPointRetail%" EQU "YES" call :Office16Activate d70b1bba-b893-4544-96e2-b7a318091c33
	if "%_AccessRetail%" EQU "YES" call :Office16Activate 67c0fc0c-deba-401b-bf8b-9c8ad8395804
	if "%_OutlookRetail%" EQU "YES" call :Office16Activate ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
	if "%_PublisherRetail%" EQU "YES" call :Office16Activate 041a06cb-c5b8-4772-809f-416d03d16654
	if "%_OneNoteRetail%" EQU "YES" call :Office16Activate d8cace59-33d2-4ac7-9b1b-9b72339c51c8
	if "%_SkypeForBusinessRetail%" EQU "YES" call :Office16Activate 83e04ee1-fa8d-436d-8994-d31a862cab77
	if "%_Word2019Retail%" EQU "YES" call :Office16Activate 059834fe-a8ea-4bff-b67b-4d006b5447d3
	if "%_Excel2019Retail%" EQU "YES" call :Office16Activate 237854e9-79fc-4497-a0c1-a70969691c6b
	if "%_PowerPoint2019Retail%" EQU "YES" call :Office16Activate 3131fd61-5e4f-4308-8d6d-62be1987c92c
	if "%_Access2019Retail%" EQU "YES" call :Office16Activate 9e9bceeb-e736-4f26-88de-763f87dcc485
	if "%_Outlook2019Retail%" EQU "YES" call :Office16Activate c8f8a301-19f5-4132-96ce-2de9d4adbd33
	if "%_Publisher2019Retail%" EQU "YES" call :Office16Activate 9d3e4cca-e172-46f1-a2f4-1d2107051444
	if "%_SkypeForBusiness2019Retail%" EQU "YES" call :Office16Activate 734c6c6e-b0ba-4298-a891-671772b2bd1b
	if "%_Word2019Volume%" EQU "YES" call :Office16Activate 059834fe-a8ea-4bff-b67b-4d006b5447d3
	if "%_Excel2019Volume%" EQU "YES" call :Office16Activate 237854e9-79fc-4497-a0c1-a70969691c6b
	if "%_PowerPoint2019Volume%" EQU "YES" call :Office16Activate 3131fd61-5e4f-4308-8d6d-62be1987c92c
	if "%_Access2019Volume%" EQU "YES" call :Office16Activate 9e9bceeb-e736-4f26-88de-763f87dcc485
	if "%_Outlook2019Volume%" EQU "YES" call :Office16Activate c8f8a301-19f5-4132-96ce-2de9d4adbd33
	if "%_Publisher2019Volume%" EQU "YES" call :Office16Activate 9d3e4cca-e172-46f1-a2f4-1d2107051444
	if "%_SkypeForBusiness2019Volume%" EQU "YES" call :Office16Activate 734c6c6e-b0ba-4298-a891-671772b2bd1b
	if "%_VisioProRetail%" EQU "YES" call :Office16Activate 6bf301c1-b94a-43e9-ba31-d494598c47fb
	if "%_VisioPro2019Retail%" EQU "YES" call :Office16Activate 5b5cf08f-b81a-431d-b080-3450d8620565
	if "%_VisioPro2019Volume%" EQU "YES" call :Office16Activate 5b5cf08f-b81a-431d-b080-3450d8620565
	if "%_ProjectProRetail%" EQU "YES" call :Office16Activate 4f414197-0fc2-4c01-b68a-86cbb9ac254c
	if "%_ProjectPro2019Retail%" EQU "YES" call :Office16Activate 2ca2bf3f-949e-446a-82c7-e25a15ec78c4
	if "%_ProjectPro2019Volume%" EQU "YES" call :Office16Activate 2ca2bf3f-949e-446a-82c7-e25a15ec78c4
	if "%_AppxVisio%" EQU "YES" call :Office16Activate 6bf301c1-b94a-43e9-ba31-d494598c47fb
	if "%_AppxProject%" EQU "YES" call :Office16Activate 4f414197-0fc2-4c01-b68a-86cbb9ac254c
::===============================================================================================================
	if /I "%ExtLocKMS%" EQU "E" goto:KMSExtern2
	call :StopService "sppsvc"
	if defined ospsversion (call :StopService "osppsvc")
::===============================================================================================================	
	call :RemoveKMSEnvironment
::===============================================================================================================	
	sc start sppsvc trigger=timer;sessionid=0  >nul 2>&1
	if "%WinDefActive%" EQU "NO" goto:KMSExtern2
:: Remove KMS activation file exclusions from Windows 8.1 and Windows 10 Defender exclusion list
	if %win% GEQ 9600 (
		(echo:)&&(echo Removing KMS activation file exclusions from Windows 10 Defender Exclusion List)
		set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\FakeClient.exe"
		powershell -noprofile -command Remove-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert.dll"
		powershell -noprofile -command Remove-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		if "%winx%" EQU "win_x32" set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert32.sys"
		if "%winx%" EQU "win_x64" set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\WinDivert64.sys"
		powershell -noprofile -command Remove-MpPreference -Force -ExclusionPath "$env:DefExclusion"
		set "DefExclusion=%OfficeRToolpath%\OfficeFixes\%winx%\vlmcsd.exe"
		powershell -noprofile -command Remove-MpPreference -Force -ExclusionPath "$env:DefExclusion"
	)
:KMSExtern2
	if %win% GEQ 9600 reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% GEQ 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% GEQ 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% LSS 9600 reg delete "HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% LSS 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
	echo:
	echo:
	timeout /t 7
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:Office16Activate
	set /a "GraceMin=0"
	if %win% GEQ 9200	(
	wmic path %slp% where ID="%1" call SetKeyManagementServiceMachine %KMSHostIP% >nul 2>&1
	wmic path %slp% where ID="%1" call SetKeyManagementServicePort %KMSPort% >nul 2>&1
	for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where ID="%1" get Name /value"') do ((echo:)&&(echo Activating %%A))
	wmic path %slp% where ID="%1" call Activate >nul 2>&1
	for /F "tokens=2 delims==" %%A in ('"wmic path %slp% where ID="%1" get GracePeriodRemaining /value"') do (set /a "GraceMin=%%A")
	wmic path %slp% where ID="%1" call ClearKeyManagementServiceMachine >nul 2>&1
	wmic path %slp% where ID="%1" call ClearKeyManagementServicePort >nul 2>&1
	)
	if %win% LSS 9200	(
	wmic path %ospp% where ID="%1" call SetKeyManagementServiceMachine %KMSHostIP% >nul 2>&1
	wmic path %ospp% where ID="%1" call SetKeyManagementServicePort %KMSPort% >nul 2>&1
	for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where ID="%1" get Name /value"') do ((echo:)&&(echo Activating %%A))
	wmic path %ospp% where ID="%1" call Activate >nul 2>&1
	for /F "tokens=2 delims==" %%A in ('"wmic path %ospp% where ID="%1" get GracePeriodRemaining /value"') do (set /a "GraceMin=%%A")
	wmic path %ospp% where ID="%1" call ClearKeyManagementServiceMachine >nul 2>&1
	wmic path %ospp% where ID="%1" call ClearKeyManagementServicePort >nul 2>&1
	)
	if %GraceMin% EQU 259200 (echo Activation successful) else (echo Activation failed)
	echo ____________________________________________________________________________
	goto:eof
::===============================================================================================================
::===============================================================================================================
:SetO16Language
    set "langnotfound=FALSE"
	if "%o16lang%" EQU "ar-SA" ((set "o16lcid=1025")&&(set "langtext=Arabic - Saudi Arabia")&&(goto:eof))
    if "%o16lang%" EQU "bg-BG" ((set "o16lcid=1026")&&(set "langtext=Bulgarian")&&(goto:eof))
	if "%o16lang%" EQU "cs-CZ" ((set "o16lcid=1029")&&(set "langtext=Czech")&&(goto:eof))
	if "%o16lang%" EQU "da-DK" ((set "o16lcid=1030")&&(set "langtext=Dansk")&&(goto:eof))
    if "%o16lang%" EQU "de-DE" ((set "o16lcid=1031")&&(set "langtext=German")&&(goto:eof))
    if "%o16lang%" EQU "el-GR" ((set "o16lcid=1032")&&(set "langtext=Greek")&&(goto:eof))
    if "%o16lang%" EQU "en-US" ((set "o16lcid=1033")&&(set "langtext=English")&&(goto:eof))
	if "%o16lang%" EQU "es-ES" ((set "o16lcid=3082")&&(set "langtext=Spanish - Spain (Traditional)")&&(goto:eof))
	if "%o16lang%" EQU "et-EE" ((set "o16lcid=1061")&&(set "langtext=Estonian")&&(goto:eof))
    if "%o16lang%" EQU "fi-FI" ((set "o16lcid=1035")&&(set "langtext=Finnish")&&(goto:eof))
    if "%o16lang%" EQU "fr-FR" ((set "o16lcid=1036")&&(set "langtext=French")&&(goto:eof))
    if "%o16lang%" EQU "he-IL" ((set "o16lcid=1037")&&(set "langtext=Hebrew")&&(goto:eof))
    if "%o16lang%" EQU "hi-IN" ((set "o16lcid=1081")&&(set "langtext=Hindi")&&(goto:eof))
	if "%o16lang%" EQU "hr-HR" ((set "o16lcid=1050")&&(set "langtext=Croatian")&&(goto:eof))
    if "%o16lang%" EQU "hu-HU" ((set "o16lcid=1038")&&(set "langtext=Hungarian")&&(goto:eof))
	if "%o16lang%" EQU "id-ID" ((set "o16lcid=1057")&&(set "langtext=Indonesian")&&(goto:eof))
    if "%o16lang%" EQU "it-IT" ((set "o16lcid=1040")&&(set "langtext=Italian")&&(goto:eof))
    if "%o16lang%" EQU "ja-JP" ((set "o16lcid=1041")&&(set "langtext=Japanese")&&(goto:eof))
    if "%o16lang%" EQU "ko-KR" ((set "o16lcid=1042")&&(set "langtext=Korean")&&(goto:eof))
	if "%o16lang%" EQU "kz-KZ" ((set "o16lcid=1087")&&(set "langtext=Kazakh")&&(goto:eof))
    if "%o16lang%" EQU "lt-LT" ((set "o16lcid=1063")&&(set "langtext=Lithuanian")&&(goto:eof))
    if "%o16lang%" EQU "lv-LV" ((set "o16lcid=1062")&&(set "langtext=Latvian")&&(goto:eof))
	if "%o16lang%" EQU "ms-MY" ((set "o16lcid=1086")&&(set "langtext=Malay - Malaysia")&&(goto:eof))
	if "%o16lang%" EQU "nb-NO" ((set "o16lcid=1044")&&(set "langtext=Norwegian - Bokml")&&(goto:eof))
	if "%o16lang%" EQU "nl-NL" ((set "o16lcid=1043")&&(set "langtext=Dutch - Netherlands")&&(goto:eof))
    if "%o16lang%" EQU "pl-PL" ((set "o16lcid=1045")&&(set "langtext=Polish")&&(goto:eof))
    if "%o16lang%" EQU "pt-BR" ((set "o16lcid=1046")&&(set "langtext=Portuguese - Brazil")&&(goto:eof))
    if "%o16lang%" EQU "pt-PT" ((set "o16lcid=2070")&&(set "langtext=Portuguese - Portugal")&&(goto:eof))
    if "%o16lang%" EQU "ro-RO" ((set "o16lcid=1048")&&(set "langtext=Romanian - Romania")&&(goto:eof))
    if "%o16lang%" EQU "ru-RU" ((set "o16lcid=1049")&&(set "langtext=Russian")&&(goto:eof))
	if "%o16lang%" EQU "sk-SK" ((set "o16lcid=1051")&&(set "langtext=Slovak")&&(goto:eof))
	if "%o16lang%" EQU "sl-SI" ((set "o16lcid=1060")&&(set "langtext=Slovenian")&&(goto:eof))
	if "%o16lang%" EQU "sr-latn-RS" ((set "o16lcid=9242")&&(set "langtext=Serbian - Latin")&&(goto:eof))
    if "%o16lang%" EQU "sv-SE" ((set "o16lcid=1053")&&(set "langtext=Swedish - Sweden")&&(goto:eof))
    if "%o16lang%" EQU "th-TH" ((set "o16lcid=1054")&&(set "langtext=Thai")&&(goto:eof))
    if "%o16lang%" EQU "tr-TR" ((set "o16lcid=1055")&&(set "langtext=Turkish")&&(goto:eof))
    if "%o16lang%" EQU "uk-UA" ((set "o16lcid=1058")&&(set "langtext=Ukrainian")&&(goto:eof))
    if "%o16lang%" EQU "vi-VN" ((set "o16lcid=1066")&&(set "langtext=Vietnamese")&&(goto:eof))
	if "%o16lang%" EQU "zh-CN" ((set "o16lcid=2052")&&(set "langtext=Chinese - China")&&(goto:eof))
    if "%o16lang%" EQU "zh-TW" ((set "o16lcid=1028")&&(set "langtext=Chinese - Taiwan")&&(goto:eof))
	set "langnotfound=TRUE"
    goto:eof
::===============================================================================================================
::===============================================================================================================
:ConvertOffice16
	cls
	echo:
	echo ================= %1 found ========================================
	echo ____________________________________________________________________________
	echo:
	if %win% GEQ 9200    (    
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-pl.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-phn.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ppd.xrm-ms"
	)
	if %win% LSS 9200    (    
    cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
    cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
    cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
    cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-pl.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-phn.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-oob.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ppd.xrm-ms"
    )
    echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:eof
::===============================================================================================================
::===============================================================================================================
:ConvertGeneral16
	cls
	echo ================= Office General Client found ==============================
	echo ____________________________________________________________________________
	echo:
	if %win% GEQ 9200    (    
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	)
	if %win% LSS 9200    (    
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	)
	echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:eof
::===============================================================================================================
::===============================================================================================================
:Office16ConversionLoop
	call :ConvertGeneral16
	if "%_ProPlusRetail%" EQU "YES" call :ConvertOffice16 ProPlus
	if "%_ProPlus2019Retail%" EQU "YES" call :ConvertOffice16 ProPlus2019 _AE
	if "%_ProPlus2019Volume%" EQU "YES" call :ConvertOffice16 ProPlus2019 _AE
	if "%_O365ProPlusRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_O365BusinessRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_MondoRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_UWPappINSTALLED%" EQU "YES" call :ConvertOffice16 ProPlus
	if "%_StandardRetail%" EQU "YES" call :ConvertOffice16 Standard
	if "%_WordRetail%" EQU "YES" call :ConvertOffice16 Word
	if "%_ExcelRetail%" EQU "YES" call :ConvertOffice16 Excel
	if "%_PowerPointRetail%" EQU "YES" call :ConvertOffice16 PowerPoint
	if "%_AccessRetail%" EQU "YES" call :ConvertOffice16 Access
	if "%_OutlookRetail%" EQU "YES" call :ConvertOffice16 Outlook
	if "%_PublisherRetail%" EQU "YES" call :ConvertOffice16 Publisher
	if "%_OneNoteRetail%" EQU "YES" call :ConvertOffice16 OneNote
	if "%_SkypeForBusinessRetail%" EQU "YES" call :ConvertOffice16 SkypeForBusiness
	if "%_Word2019Retail%" EQU "YES" call :ConvertOffice16 Word2019 _AE
	if "%_Excel2019Retail%" EQU "YES" call :ConvertOffice16 Excel2019 _AE
	if "%_PowerPoint2019Retail%" EQU "YES" call :ConvertOffice16 PowerPoint2019 _AE
	if "%_Access2019Retail%" EQU "YES" call :ConvertOffice16 Access2019 _AE
	if "%_Outlook2019Retail%" EQU "YES" call :ConvertOffice16 Outlook2019 _AE
	if "%_Publisher2019Retail%" EQU "YES" call :ConvertOffice16 Publisher2019 _AE
	if "%_SkypeForBusiness2019Retail%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2019 _AE
	if "%_Word2019Volume%" EQU "YES" call :ConvertOffice16 Word2019 _AE
	if "%_Excel2019Volume%" EQU "YES" call :ConvertOffice16 Excel2019 _AE
	if "%_PowerPoint2019Volume%" EQU "YES" call :ConvertOffice16 PowerPoint2019 _AE
	if "%_Access2019Volume%" EQU "YES" call :ConvertOffice16 Access2019 _AE
	if "%_Outlook2019Volume%" EQU "YES" call :ConvertOffice16 Outlook2019 _AE
	if "%_Publisher2019Volume%" EQU "YES" call :ConvertOffice16 Publisher2019 _AE
	if "%_SkypeForBusiness2019Volume%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2019 _AE
	if "%_VisioProRetail%" EQU "YES" call :ConvertOffice16 VisioPro
	if "%_AppxVisio%" EQU "YES" call :ConvertOffice16 VisioPro
	if "%_VisioPro2019Retail%" EQU "YES" call :ConvertOffice16 VisioPro2019 _AE
	if "%_VisioPro2019Volume%" EQU "YES" call :ConvertOffice16 VisioPro2019 _AE
	if "%_ProjectProRetail%" EQU "YES" call :ConvertOffice16 ProjectPro
	if "%_AppxProject%" EQU "YES" call :ConvertOffice16 ProjectPro
	if "%_ProjectPro2019Retail%" EQU "YES" call :ConvertOffice16 ProjectPro2019 _AE
	if "%_ProjectPro2019Volume%" EQU "YES" call :ConvertOffice16 ProjectPro2019 _AE
	cls
	echo ================= INSTALLING GVLK ==========================================
	echo ____________________________________________________________________________
	if %win% GEQ 9200 if "%_ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 ProfessionalPlus"
	if %win% LSS 9200 if "%_ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 ProfessionalPlus"
	if %win% GEQ 9200 if "%_ProPlus2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office 2019 ProfessionalPlus"
	if %win% LSS 9200 if "%_ProPlus2019Retail%" EQU "YES" call :OfficeGVLKInstall "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office 2019 ProfessionalPlus"
	if %win% GEQ 9200 if "%_ProPlus2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office 2019 ProfessionalPlus"
	if %win% LSS 9200 if "%_ProPlus2019Volume%" EQU "YES" call :OfficeGVLKInstall "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office 2019 ProfessionalPlus"
	if %win% GEQ 9200 if "%_O365ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office 365 ProfessionalPlus"
	if %win% LSS 9200 if "%_O365ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office 365 ProfessionalPlus"
	if %win% GEQ 9200 if "%_O365BusinessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office 365 Business"
	if %win% LSS 9200 if "%_O365BusinessRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office 365 Business"
	if %win% GEQ 9200 if "%_MondoRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% LSS 9200 if "%_MondoRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% GEQ 9200 if "%_UWPappINSTALLED%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 UWP Appx Desktop Apps"
	if %win% LSS 9200 if "%_UWPappINSTALLED%" EQU "YES" call :OfficeGVLKInstall "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 UWP Appx Desktop Apps"
	if %win% GEQ 9200 if "%_StandardRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office 2016 Standard"
	if %win% LSS 9200 if "%_StandardRetail%" EQU "YES" call :OfficeGVLKInstall "JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office 2016 Standard"
	if %win% GEQ 9200 if "%_WordRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% LSS 9200 if "%_WordRetail%" EQU "YES" call :OfficeGVLKInstall "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% GEQ 9200 if "%_ExcelRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% LSS 9200 if "%_ExcelRetail%" EQU "YES" call :OfficeGVLKInstall "9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% GEQ 9200 if "%_PowerPointRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% LSS 9200 if "%_PowerPointRetail%" EQU "YES" call :OfficeGVLKInstall "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% GEQ 9200 if "%_AccessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% LSS 9200 if "%_AccessRetail%" EQU "YES" call :OfficeGVLKInstall "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% GEQ 9200 if "%_OutlookRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% LSS 9200 if "%_OutlookRetail%" EQU "YES" call :OfficeGVLKInstall "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% GEQ 9200 if "%_PublisherRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% LSS 9200 if "%_PublisherRetail%" EQU "YES" call :OfficeGVLKInstall "F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% GEQ 9200 if "%_OneNoteRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% LSS 9200 if "%_OneNoteRetail%" EQU "YES" call :OfficeGVLKInstall "DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% GEQ 9200 if "%_SkypeForBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	if %win% LSS 9200 if "%_SkypeForBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	if %win% GEQ 9200 if "%_Word2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% LSS 9200 if "%_Word2019Retail%" EQU "YES" call :OfficeGVLKInstall "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% GEQ 9200 if "%_Excel2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% LSS 9200 if "%_Excel2019Retail%" EQU "YES" call :OfficeGVLKInstall "TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% GEQ 9200 if "%_PowerPoint2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% LSS 9200 if "%_PowerPoint2019Retail%" EQU "YES" call :OfficeGVLKInstall "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% GEQ 9200 if "%_Access2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% LSS 9200 if "%_Access2019Retail%" EQU "YES" call :OfficeGVLKInstall "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% GEQ 9200 if "%_Outlook2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% LSS 9200 if "%_Outlook2019Retail%" EQU "YES" call :OfficeGVLKInstall "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% GEQ 9200 if "%_Publisher2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% LSS 9200 if "%_Publisher2019Retail%" EQU "YES" call :OfficeGVLKInstall "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% GEQ 9200 if "%_SkypeForBusiness2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NCJ33-JHBBY-HTK98-MYCV8-HMKHJ","Skype For Business 2019"
	if %win% LSS 9200 if "%_SkypeForBusiness2019Retail%" EQU "YES" call :OfficeGVLKInstall "NCJ33-JHBBY-HTK98-MYCV8-HMKHJ","Skype For Business 2019"
	if %win% GEQ 9200 if "%_Word2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% LSS 9200 if "%_Word2019Volume%" EQU "YES" call :OfficeGVLKInstall "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% GEQ 9200 if "%_Excel2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% LSS 9200 if "%_Excel2019Volume%" EQU "YES" call :OfficeGVLKInstall "TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% GEQ 9200 if "%_PowerPoint2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% LSS 9200 if "%_PowerPoint2019Volume%" EQU "YES" call :OfficeGVLKInstall "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% GEQ 9200 if "%_Access2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% LSS 9200 if "%_Access2019Volume%" EQU "YES" call :OfficeGVLKInstall "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% GEQ 9200 if "%_Outlook2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% LSS 9200 if "%_Outlook2019Volume%" EQU "YES" call :OfficeGVLKInstall "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% GEQ 9200 if "%_Publisher2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% LSS 9200 if "%_Publisher2019Volume%" EQU "YES" call :OfficeGVLKInstall "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% GEQ 9200 if "%_SkypeForBusiness2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NCJ33-JHBBY-HTK98-MYCV8-HMKHJ","Skype For Business 2019"
	if %win% LSS 9200 if "%_SkypeForBusiness2019Volume%" EQU "YES" call :OfficeGVLKInstall "NCJ33-JHBBY-HTK98-MYCV8-HMKHJ","Skype For Business 2019"
	if %win% GEQ 9200 if "%_ProjectProRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project 2016 Professional"
	if %win% LSS 9200 if "%_ProjectProRetail%" EQU "YES" call :OfficeGVLKInstall "YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project 2016 Professional"
	if %win% GEQ 9200 if "%_VisioProRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio 2016 Professional"
	if %win% LSS 9200 if "%_VisioProRetail%" EQU "YES" call :OfficeGVLKInstall "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio 2016 Professional"
	if %win% GEQ 9200 if "%_AppxProject%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","ProjectPro 2016 UWP Appx Desktop App"
	if %win% LSS 9200 if "%_AppxProject%" EQU "YES" call :OfficeGVLKInstall "YG9NW-3K39V-2T3HJ-93F3Q-G83KT","ProjectPro 2016 UWP Appx Desktop App"
	if %win% GEQ 9200 if "%_AppxVisio%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","VisioPro 2016 UWP Appx Desktop App"
	if %win% LSS 9200 if "%_AppxVisio%" EQU "YES" call :OfficeGVLKInstall "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","VisioPro 2016 UWP Appx Desktop App"
	if %win% GEQ 9200 if "%_ProjectPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project 2019 Professional"
	if %win% LSS 9200 if "%_ProjectPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project 2019 Professional"
	if %win% GEQ 9200 if "%_ProjectPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project 2019 Professional"
	if %win% LSS 9200 if "%_ProjectPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project 2019 Professional"
	if %win% GEQ 9200 if "%_VisioPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio 2019 Professional"
	if %win% LSS 9200 if "%_VisioPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio 2019 Professional"
	if %win% GEQ 9200 if "%_VisioPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio 2019 Professional"
	if %win% LSS 9200 if "%_VisioPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio 2019 Professional"
	echo:
	echo:
	timeout /t 7
	goto:eof
::===============================================================================================================
::===============================================================================================================
:OfficeGVLKInstall
	echo:
	if %win% GEQ 9200    (
    echo %4
	wmic path %~1 where version='%~2' call InstallProductKey ProductKey="%~3" >nul 2>&1
    if %errorlevel% EQU 0 ((echo:)&&(echo Successfully installed %~3)&&(echo:))
    if %errorlevel% NEQ 0 ((echo:)&&(echo Installing %~3 failed)&&(echo:))
    )
    if %win% LSS 9200    (
    echo %2
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inpkey:%~1 >nul 2>&1
    if %errorlevel% EQU 0 ((echo:)&&(echo Successfully installed %~1)&&(echo:))
    if %errorlevel% NEQ 0 ((echo:)&&(echo Installing %~1 Failed)&&(echo:))
    )
	echo ____________________________________________________________________________
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CreateKMSEnvironment
	wmic path %sls% where version='%slsversion%' call SetKeyManagementServiceMachine MachineName="%KMSHostIP%" >nul 2>&1
	wmic path %sls% where version='%slsversion%' call SetKeyManagementServicePort %KMSPort% >nul 2>&1
	if "%OfficeRToolKMS%" EQU "NO" goto:eof
	netsh advfirewall firewall add rule name="VLMCSD" dir=in program="%OfficeRToolpath%\OfficeFixes\%winx%\vlmcsd.exe" profile=any localport=%KMSPort% protocol=TCP action=allow remoteip=any >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out program="%OfficeRToolpath%\OfficeFixes\%winx%\vlmcsd.exe" profile=any localport=%KMSPort% protocol=TCP action=allow remoteip=any >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=in program="%OfficeRToolpath%\OfficeFixes\%winx%\FakeClient.exe" profile=any protocol=TCP action=allow remoteip=any >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out program="%OfficeRToolpath%\OfficeFixes\%winx%\FakeClient.exe" profile=any protocol=TCP action=allow remoteip=any >nul 2>&1
	start "KMS VLMCSD" /D "%OfficeRToolpath%\OfficeFixes\%winx%\" /MIN "%OfficeRToolpath%\OfficeFixes\%winx%\vlmcsd.exe" -P %KMSPort% -A %AI% -R %RI% -C %o16lcid% -De -r1
	if %win% GEQ 9600 (
		route add %KMSHostIP% 0.0.0.0 IF 1 >nul 2>&1
		start /B "" /D "%OfficeRToolpath%\OfficeFixes\%winx%\" /MIN "%OfficeRToolpath%\OfficeFixes\%winx%\FakeClient.exe" %KMSHostIP% >nul 2>&1
	)
	goto:eof
::===============================================================================================================
::===============================================================================================================
:RemoveKMSEnvironment
	if "%OfficeRToolKMS%" EQU "NO" goto:eof
	tasklist | find /i "vlmcsd.exe" >nul 2>&1
	if %errorlevel% NEQ 0 goto:eof
	taskkill /t /f /IM "vlmcsd.exe" >nul 2>&1
	if %win% GEQ 9600 (
		taskkill /t /f /IM "FakeClient.exe" >nul 2>&1
		route delete %KMSHostIP% 0.0.0.0 >nul 2>&1
	)
	netsh advfirewall firewall delete rule name="VLMCSD" >nul 2>&1
	goto:eof
::===============================================================================================================
::===============================================================================================================
:StopService
	sc query "%~1" | findstr /i "RUNNING" >nul 2>&1
	if %errorlevel% EQU 0 sc stop "%~1" >nul 2>&1
	sc query "%~1" | findstr /i "STOPPED" >nul 2>&1
	if %errorlevel% NEQ 0 goto:StopService
	goto:eof
::===============================================================================================================
::===============================================================================================================
	:StartService
	sc query "%~1" | findstr /i "STOPPED" >nul 2>&1
	if %errorlevel% EQU 0 sc start "%~1" >nul 2>&1
	sc query "%~1" | findstr /i "RUNNING" >nul 2>&1
	if %errorlevel% NEQ 0 goto:StartService
	goto:eof
::===============================================================================================================
::===============================================================================================================
:TheEndIsNear
	echo:
	echo:
	echo Ending OfficeRTool ...
	timeout /t 7
	exit
::===============================================================================================================
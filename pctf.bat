@echo off
setlocal enabledelayedexpansion
setlocal enableextensions
echo ==================================================
echo    ___  _______________
echo   / _ \/ ___/_  __/ __/
echo  / ___/ /__  / / / _/   
echo /_/   \___/ /_/ /_/    
echo.
echo PC Transfer Tool v1.3.1 -- Developed by Samuel Bunko
echo ==================================================
if "%1"=="export" (
	echo Export Agent
	echo ==================================================
	echo Relaunched with elevated privileges...
    set "worDir=%2"
	set "uName=%3"
	set "opt=%4"

    goto exportUser
)
if "%1"=="importUser" (
	echo Import Agent
	echo ==================================================
    echo Relaunched with elevated privileges...
	set "uName=%2"
	set "opt=%3"
    goto importUser
)
echo.
echo 1. Import current user
echo 2. Import another user
echo 3. Export current user
echo 4. Export another user
echo.

:: Print script options
:getOptions
set /p opt=What would you like to do? 
if "%opt%" equ "1" (
	goto checkLocalUser
) else if "%opt%" equ "2" (
	goto selectLocalUser	
) else if "%opt%" equ "3" (
	goto checkLocalUser
) else if "%opt%" equ "4" (
	goto selectLocalUser
) else (
	echo "Invalid option. Please select one of the options listed above (1, 2, 3, or 4)."
	goto getOptions
)

:: Identify target for operation
:selectLocalUser
set /p "uName=Type the username of the account: "
	if not exist "C:\Users\!uName!" (
		echo ERROR: C:\Users\!uName! could not be found; please try again.
		goto selectLocalUser
	) else (
		if exist "C:\Users\!uName!" (
			echo Working in directory C:\Users\!uName!..
			if "%opt%" equ "1" goto importUser
			if "%opt%" equ "2" goto importUser
			if "%opt%" equ "3" goto exportUser
			if "%opt%" equ "4" goto exportUser
		)
	)

:: Check user account to verify it is registered on the workstation
:checkLocalUser
if /i "%userprofile%" equ "C:\Windows\system32\config\systemprofile" (
	set /p "uName=ERROR: User account not found - enter the desired username below: "
	if not exist "C:\Users\!uName!" (
		goto checkLocalUser
	) else (
		if exist "C:\Users\!uName!" (
			echo Working in directory C:\Users\!uName!..
			if "%opt%" equ "1" goto importUser
			if "%opt%" equ "2" goto importUser
			if "%opt%" equ "3" goto exportUser
			if "%opt%" equ "4" goto exportUser
		)
	)
)
if exist "C:\Users\%USERNAME%" (
	echo Working in directory %userprofile%..
	set "uName=%USERNAME%"
	if "%opt%" equ "1" goto importUser
	if "%opt%" equ "2" goto importUser
	if "%opt%" equ "3" goto exportUser
	if "%opt%" equ "4" goto exportUser
)

:: Configuring working directory for user export
:exportDirectory
set /p worDir=Enter a directory you wish to export the user data to (or press ENTER to default to C:\aitsys\!uName!_%COMPUTERNAME%_Export): 
if "%worDir%"=="" (
    set "worDir=C:\aitsys\!uName!_%COMPUTERNAME%_Export"
)

:: Tidy work directory of illegal characters to avoid issues exporting content
set "cleanDir=%worDir%"
for %%C in (^< ^> ^" ^| ^? ^* " ") do (
    set "worDir=!cleanDir:%%C=!"
	if not exist "%worDir%" (
		mkdir "%worDir%"
        echo Created directory %worDir% for export
        goto exportUser
    )
	echo Working in directory %worDir%
	goto exportUser
)

:: Function for exporting user data
:exportUser
:: Error check to confirm work directory has been specified
if not defined worDir goto exportDirectory
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'export', '%worDir%', '%uName%', '%opt%' -Verb runAs"
    exit /b
)
echo %uName% > %worDir%\oldUsername.txt
:: Check if Downloads Folder is required
%SystemRoot%\System32\choice.exe /C YN /N /M "Is the Downloads folder required? (y/n) "
if errorlevel 1 set dlReq=y
if errorlevel 2 set dlReq=n

:: Check if Printer Export is required
%SystemRoot%\System32\choice.exe /C YN /N /M "Would you like to export printers? (y/n) "
if errorlevel 1 set prReq=y
if errorlevel 2 set prReq=n
echo ==================================================
echo Commencing export - stand back!

::Set variables
set "roaming=C:\Users\%uName%\AppData\Roaming"
set "local=C:\Users\%uName%\AppData\Local"

:: Export M365 Data
echo Exporting Microsoft 365 data (templates, signatures, User MRU, Microsoft Notes, Outlook Categories)
taskkill /f /im Microsoft.Notes.exe 2> nul
mkdir "%worDir%\MicrosoftNotes" 2> nul
robocopy "%roaming%\Microsoft\Signatures" "%worDir%\Signatures" /E /MT:16 /R:3 /W:1 /XJ > nul
robocopy "%roaming%\Microsoft\Templates" "%worDir%\Templates" /E /MT:16 /R:3 /W:1 /XJ > nul
robocopy "%local%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState" "%worDir%\MicrosoftNotes" plum.sqlite settings.json > nul
if "%opt%" equ "3" (
	reg export "HKCU\Software\Microsoft\Office\16.0\Word\User MRU"  "%worDir%\Word_MRU.reg" /y > nul 2>&1
	reg export "HKCU\Software\Microsoft\Office\16.0\Excel\User MRU" "%worDir%\Excel_MRU.reg" /y > nul 2>&1
	reg export "HKCU\Software\Microsoft\Office\16.0\PowerPoint\User MRU" "%worDir%\PowerPoint_MRU.reg" /y > nul 2>&1
	reg export "HKCU\Software\Microsoft\Office\16.0\OneNote\User MRU" "%worDir%\OneNote_MRU.reg" /y > nul 2>&1
	reg export "HKCU\Software\Microsoft\Office\16.0\Common\Categories" "%worDir%\Outlook_Categories.reg" /y > nul 2>&1
)
if "%opt%" equ "4" (
	for /f "delims=" %%S in ('powershell -NoProfile -Command "$p='C:\Users\!uName!'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.ProfileImagePath -ieq $p } | Select-Object -ExpandProperty PSChildName"') do set "SID=%%S"
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\Word\User MRU" "%worDir%\Word_MRU.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\Excel\User MRU" "%worDir%\Excel_MRU.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\PowerPoint\User MRU" "%worDir%\PowerPoint_MRU.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\OneNote\User MRU" "%worDir%\OneNote_MRU.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\Common\Categories" "%worDir%\Outlook_Categories.reg" /y > nul 2>&1
)

:: Export Quick Access / Favorites (File Explorer)
echo Exporting Quick Access / Favorites (File Explorer)
if not exist "%worDir%\QuickAccess" mkdir "%worDir%\QuickAccess"
robocopy "%roaming%\Microsoft\Windows\Recent\AutomaticDestinations" "%worDir%\QuickAccess\AutomaticDestinations" *.* /E > nul
robocopy "%roaming%\Microsoft\Windows\Recent\CustomDestinations" "%worDir%\QuickAccess\CustomDestinations" *.* /E > nul

:: Export Google Chrome Data (Default and additional profiles)
echo Exporting Google Chrome data..
robocopy "%local%\Google\Chrome\User Data\Default" "%worDir%\Chrome\Default" /E /MT:16 /R:3 /W:1 /XJ /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB" "GPUCache" "ShaderCache" "Network" "Safe Browsing Network" "Sessions" /XF "Cookies" "Cookies-journal" "Network Persistent State" "History-journal" "History Provider Cache" "Session_*" "Tabs_*" > nul 2>&1
for /D %%G in ("%local%\Google\Chrome\User Data\Profile *") do (
	robocopy "%%G" "%worDir%\Chrome\%%~nxG" /E /MT:16 /R:3 /W:1 /XJ /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB" "GPUCache" "ShaderCache" "Network" "Safe Browsing Network" "Sessions" /XF "Cookies" "Cookies-journal" "Network Persistent State" "History-journal" "History Provider Cache" "Session_*" "Tabs_*" > nul 2>&1
)

:: Export Microsoft Edge Data
echo Exporting Microsoft Edge data..
robocopy "%local%\Microsoft\Edge\User Data\Default" "%worDir%\Edge" /E /MT:16 /R:3 /W:1 /XJ /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB" "GPUCache" "ShaderCache" "Network" "Safe Browsing Network" "Sessions" /XF "Cookies" "Cookies-journal" "Network Persistent State" "History-journal" "History Provider Cache" "Session_*" "Tabs_*" > nul
for /D %%G in ("%local%\Microsoft\Edge\User Data\Profile *") do (
	robocopy "%%G" "%worDir%\Edge\%%~nxG" /E /MT:16 /R:3 /W:1 /XJ /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB" "GPUCache" "ShaderCache" "Network" "Safe Browsing Network" "Sessions" /XF "Cookies" "Cookies-journal" "Network Persistent State" "History-journal" "History Provider Cache" "Session_*" "Tabs_*" > nul 2>&1
)

:: Export Mozilla Firefox Data
echo Exporting Mozilla Firefox data..
copy /Y "%roaming%\Mozilla\Firefox\installs.ini" "%worDir%\Firefox\installs.ini" > nul
copy /Y "%roaming%\Mozilla\Firefox\profiles.ini" "%worDir%\Firefox\profiles.ini" > nul
robocopy "%roaming%\Mozilla\Firefox\Profiles" "%worDir%\Firefox\Profiles" /E /MT:16 /R:3 /W:1 /XJ /XD "shader-cache" "saved-telemetry-pings" "crashes" > nul

:: Export Adobe Stamps & signatures
if exist "%roaming%\Adobe" (
	echo Exporting Adobe Acrobat Stamps and Signature data..
	mkdir "%worDir%\Adobe"
	mkdir "%worDir%\Adobe\Stamps"
	mkdir "%worDir%\Adobe\Security"
	if "%opt%" equ "3" (
		reg export "HKCU\Software\Adobe\Adobe Acrobat\DC\Annots" "%worDir%\Adobe\Acrobat_Annots.reg" /y > nul 2>&1
		reg export "HKCU\Software\Adobe\Adobe Acrobat\DC\Security" "%worDir%\Adobe\Acrobat_Security.reg" /y > nul 2>&1
	)
	if "%opt%" equ "4" (
		for /f "delims=" %%S in ('powershell -NoProfile -Command "$p='C:\Users\!uName!'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.ProfileImagePath -ieq $p } | Select-Object -ExpandProperty PSChildName"') do set "SID=%%S"
		echo !SID!
		reg export "HKU\!SID!\Software\Adobe\Adobe Acrobat\DC\Annots" "%worDir%\Adobe\Acrobat_Annots.reg" /y > nul 2>&1
		reg export "HKU\!SID!\Software\Adobe\Adobe Acrobat\DC\Security" "%worDir%\Adobe\Acrobat_Security.reg" /y > nul 2>&1
	)
	robocopy "%roaming%\Adobe\Acrobat\DC\Stamps" "%worDir%\Adobe\Stamps" *.* /E > nul
	robocopy "%roaming%\Adobe\Acrobat\DC\Security" "%worDir%\Adobe\Security" appearance.acrodata > nul
)

:: Export Power Plan & Lid Configurations
echo Exporting Power Plan and Lid Configurations..
for /f "tokens=2 delims=:" %%A in ('powercfg /getactivescheme') do set "REST=%%A"
for /f "tokens=1 delims=() " %%B in ("!REST!") do set "ACTIVE_GUID=%%B"
set "ACTIVE_GUID=%ACTIVE_GUID: =%"
mkdir "%worDir%\Power" 2>nul
powercfg /export "%worDir%\Power\active_powerplan.pow" %ACTIVE_GUID%
powercfg /q > "%worDir%\Power\powercfg_dump.txt"

:: Export Accessibility niceties - StickyKeys, ToggleKeys, FilterKeys, MouseKeys, HighContrast, SoundSentry, keyboard response timings, Pointer scheme & per-cursor file overrides, Caret blink rate/width, menu/show delays, font smoothing, scaling preferences, Ease-of-Access app preferences.
echo Exporting Accessibility niceties..
mkdir "%worDir%\Accessibility" 2>nul
if "%opt%" equ "3" (
	reg export "HKCU\Control Panel\Accessibility" "%worDir%\Accessibility\Accessibility.reg" /y > nul 2>&1
	reg export "HKCU\Control Panel\Cursors" "%worDir%\Accessibility\Cursors.reg" /y > nul 2>&1
	reg export "HKCU\Control Panel\Desktop" "%worDir%\Accessibility\Desktop.reg" /y > nul 2>&1
	reg export "HKCU\Software\Microsoft\Accessibility" "%worDir%\Accessibility\MS_Accessibility.reg" /y > nul 2>&1
)
if "%opt%" equ "4" (
	for /f "delims=" %%S in ('powershell -NoProfile -Command "$p='C:\Users\!uName!'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.ProfileImagePath -ieq $p } | Select-Object -ExpandProperty PSChildName"') do set "SID=%%S"
	echo !SID!
	reg export "HKU\!SID!\Control Panel\Accessibility" "%worDir%\Accessibility\Accessibility.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Control Panel\Cursors" "%worDir%\Accessibility\Cursors.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Control Panel\Desktop" "%worDir%\Accessibility\Desktop.reg" /y > nul 2>&1
	reg export "HKU\!SID!\Software\Microsoft\Accessibility" "%worDir%\Accessibility\MS_Accessibility.reg" /y > nul 2>&1
)

:: Export Downloads Folder
if "%dlReq%" equ "y" (
    echo Exporting Downloads folder..
	robocopy "C:\Users\%uName%\Downloads" "%worDir%\Downloads" /E /MT:16 /R:3 /W:1 /XJ > nul
)

:: Export third-party font files
echo Copying font files..
if not exist "%worDir%\Fonts" mkdir "%worDir%\Fonts"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$fontList=((Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/sbaitsys/pctf/main/fonts.txt' -UseBasicParsing).Content -split '\r?\n' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ }); $dest=if([string]::IsNullOrWhiteSpace('%worDir%')){ Join-Path $PWD 'Fonts' } else { Join-Path '%worDir%' 'Fonts' }; if(-not (Test-Path -LiteralPath $dest)){ New-Item -ItemType Directory -Path $dest | Out-Null }; Get-ChildItem 'C:\Windows\Fonts' -Include *.ttf,*.ttc,*.otf,*.fon -File -Force | ForEach-Object { if($fontList -notcontains $_.Name.ToLower()){ Copy-Item -LiteralPath $_.FullName -Destination $dest -Force -ErrorAction SilentlyContinue } }"

:: Export Wallpaper Data
echo Exporting Wallpaper..
if exist "%roaming%\Microsoft\Windows\Themes\TranscodedWallpaper" (
	robocopy "%roaming%\Microsoft\Windows\Themes" "%worDir%\Wallpaper" /E /MT:16 /R:3 /W:1 /XJ > nul
	cd %worDir%\Wallpaper
    ren TranscodedWallpaper TranscodedWallpaper.jpg
) else (
    echo No wallpaper found. Skipping.
)

:: Export Wi-Fi Profiles
echo Exporting WiFi profiles..
if not exist "%worDir%\WiFiProfiles" mkdir "%worDir%\WiFiProfiles"
netsh wlan export profile key=clear folder="%worDir%\WiFiProfiles" > nul
if "%prReq%" equ "y" (
	echo Moving onto printer export..
    goto printerExpo
) else (
    goto zip
)

:: Export Printer Drivers/Configurations
:printerExpo
set "DRIVER_DIR=%worDir%\Printers\Drivers"
set "USED_LIST=%worDir%\Printers\used_drivers.txt"
set "PRINTER_LIST=%worDir%\Printers\printers.txt"
set "CONFIG_FILE=%worDir%\Printers\PrinterConfig.csv"

if not exist "%DRIVER_DIR%" mkdir "%DRIVER_DIR%" >nul
echo Exporting printer installations..
powershell -NoProfile -Command "Get-Printer | Where-Object { $_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe|DYMO|Brother QL|Generic' } | Select-Object -ExpandProperty Name | Set-Content -Path '%PRINTER_LIST%' -Encoding ASCII"
powershell -Command "Get-Printer | Where-Object { $_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe|Zebra|DYMO|Brother QL|Generic' } | Export-Csv -Path \"%CONFIG_FILE%\" -NoTypeInformation"
echo Exporting third-party drivers..
powershell -Command "Get-PrinterDriver | Where-Object { $_.InfPath -and (Test-Path $_.InfPath) -and ($_.Name -notmatch 'Microsoft|PDF|OneNote|Zebra|Remote Desktop|Adobe|DYMO|Brother QL|Generic') } | ForEach-Object { $source = Split-Path $_.InfPath -Parent; $dest = Join-Path \"%DRIVER_DIR%\" $_.Name; Copy-Item -Path $source -Destination $dest -Recurse -Force }"
echo Exporting printer configurations..
if not exist "%worDir%\Printers\Configurations" mkdir "%worDir%\Printers\Configurations"
for /f "usebackq delims=" %%A in ("%PRINTER_LIST%") do (
    call :expoSpecificPrinter "%%~A"
)
goto zip

:expoSpecificPrinter
set "PN=%~1"
:: strip quotes and commas
set "PN=%PN:"=%"
set "PN=%PN:,=%"
:: replace filesystem-invalid chars in output filename only
set "FN=%PN%"
set "FN=%FN:\=_%"
set "FN=%FN:/=_%"
set "FN=%FN::=_%"
set "FN=%FN:?=_%"
set "FN=%FN:<=_%"
set "FN=%FN:>=_%"
set "FN=%FN:|=_%"
echo Exporting configuration for "%PN%"
rundll32 printui.dll,PrintUIEntry /Ss /n "%PN%" /a "%worDir%\Printers\Configurations\%FN%.dat" m c u d g
exit /b

:: Function for importing user data
:importUser
:: Administration elevation
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'importUser', '%uName%', '%opt%' -Verb runAs"
    exit /b
)

:: Check if Import Directory exists
set /p "importDir=Enter the FULL filepath containing the exported user data: "
if not exist "!importDir!" (
	echo Unable to locate directory; try again
	goto importUser
	)
echo ==================================================
echo Commencing import - stand back!

:: Set variables
set "roaming=C:\Users\%uName%\AppData\Roaming"
set "local=C:\Users\%uName%\AppData\Local"
set /p oldUsername=<"%importDir%\oldUsername.txt"

:: Creating compatibility junction for Quick Access, MRU and other if old username is different
if not %oldusername% == %uName% (
	mklink /J "C:\Users\%oldUsername%" "C:\Users\%uName%"
)

:: Close any applications correlated with imported data
%SystemRoot%\System32\choice.exe /C YN /N /M "To import succesfully, all relevant applications must be closed. Force close applications? (y/n) "
if errorlevel 1 set killTasks=y
if errorlevel 2 set killTasks=n
if "%killTasks%" equ "y" (
    taskkill /f /im explorer.exe 2> nul
	taskkill /f /im Microsoft.Notes.exe 2> nul
	taskkill /f /im chrome.exe 2> nul
	taskkill /f /im msedge.exe 2> nul
	taskkill /f /im firefox.exe 2> nul
	taskkill /f /im outlook.exe 2> nul
	taskkill /f /im winword.exe 2> nul
	taskkill /f /im excel.exe 2> nul
	taskkill /f /im powerpnt.exe 2> nul 
	taskkill /f /im Acrord32.exe 2> nul
) else (
	echo Unable to complete import; please save any open files within Chrome, Edge, Firefox, Outlook, Word, Excel or PowerPoint and try again.
)

:: Import M365 Data
echo Importing Microsoft 365 data (templates, signatures, User MRU)
robocopy "%importDir%\Signatures" "%roaming%\Microsoft\Signatures" /E /MT:16 /R:3 /W:1 /XJ > nul
robocopy "%importDir%\Templates" "%roaming%\Microsoft\Templates" /E /MT:16 /R:3 /W:1 /XJ > nul 
robocopy "%importDir%\MicrosoftNotes" "%local%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState" plum.sqlite settings.json > nul
if "%opt%" equ "1" (
	reg import "%importDir%\Word_MRU.reg" > nul 2>&1
	reg import "%importDir%\Excel_MRU.reg" > nul 2>&1
	reg import "%importDir%\PowerPoint_MRU.reg" > nul 2>&1
	reg import "%importDir%\OneNote_MRU.reg" > nul 2>&1
	reg import "%importDir%\Outlook_Categories.reg" > nul 2>&1
)
if "%opt%" equ "2" (
    reg load "HKU\TempHive" "C:\Users\%uName%\NTUSER.DAT" > nul 2>&1
	reg import "%importDir%\Word_MRU.reg" > nul 2>&1
	reg import "%importDir%\Excel_MRU.reg" > nul 2>&1
	reg import "%importDir%\PowerPoint_MRU.reg" > nul 2>&1
	reg import "%importDir%\OneNote_MRU.reg" > nul 2>&1
	reg import "%importDir%\Outlook_Categories.reg" > nul 2>&1
	reg unload "HKU\TempHive" > nul 2>&1
)

:: Import Quick Access / Favorites (File Explorer)
echo Importing Quick Access / Favorites (File Explorer)
mkdir "%roaming%\Microsoft\Windows\Recent\AutomaticDestinations" > nul 2>&1
mkdir "%roaming%\Microsoft\Windows\Recent\CustomDestinations" > nul 2>&1
robocopy "%importDir%\QuickAccess\AutomaticDestinations" "%roaming%\Microsoft\Windows\Recent\AutomaticDestinations" *.* /E > nul 
robocopy "%importDir%\QuickAccess\CustomDestinations" "%roaming%\Microsoft\Windows\Recent\CustomDestinations" *.* /E > nul
start explorer.exe

:: Import Google Chrome Data
echo Importing Google Chrome Data..
robocopy "%importDir%\Chrome" "%local%\Google\Chrome\User Data" /E /MT:16 /R:3 /W:1 /XJ > nul 

:: Import Microsoft Edge Data
echo Importing Microsoft Edge Data..
robocopy "%importDir%\Edge" "%local%\Microsoft\Edge\User Data\Default" /E /MT:16 /R:3 /W:1 /XJ > nul

:: Import Mozilla Firefox Data
echo Importing Mozilla Firefox Data..
copy "%importDir%\Firefox\installs.ini" "%roaming%\Mozilla\Firefox" /E > nul
copy "%importDir%\Firefox\profiles.ini" "%roaming%\Mozilla\Firefox" /E > nul
robocopy "%importDir%\Firefox\Profiles" "%roaming%\Mozilla\Firefox\Profiles" /E /MT:16 /R:3 /W:1 /XJ > nul 

:: Export Adobe Stamps & signatures
if exist "%roaming%\Adobe" (
	echo Exporting Adobe Acrobat Stamps and Signature data..
	robocopy "%importDir%\Adobe\Stamps" "%roaming%\Adobe\Acrobat\DC\Stamps" *.* /E > nul
	robocopy "%importDir%\Adobe\Security" "%roaming%\Adobe\Acrobat\DC\Security" appearance.acrodata > nul
	if "%opt%" equ "1" (
		reg import "%importDir%\Acrobat_Annots.reg" > nul 2>&1
		reg import "%importDir%\Acrobat_Security.reg" > nul 2>&1
	)
	if "%opt%" equ "2" (
		reg load "HKU\TempHive" "C:\Users\%uName%\NTUSER.DAT" > nul 2>&1
		reg import "%importDir%\Acrobat_Annots.reg" > nul 2>&1
		reg import "%importDir%\Acrobat_Security.reg" > nul 2>&1
		reg unload "HKU\TempHive" > nul 2>&1
	)
)

:: Import Power Plan & Lid Configurations
for /f "tokens=4" %%G in ('powercfg /import "%importDir%\Power\active_powerplan.pow" ^&^& powercfg /l ^| findstr /i "Imported"') do set NEWGUID=%%G
powercfg /setactive %NEWGUID% 2>nul
powercfg /setacvalueindex %NEWGUID% SUB_BUTTONS LIDACTION 0 2>nul :: Configure on Lid Close to do nothing (on power)
powercfg /setdcvalueindex %NEWGUID% SUB_BUTTONS LIDACTION 1 2>nul :: Configure on Lid Close to Sleep (on battery)
powercfg /setacvalueindex %NEWGUID% SUB_BUTTONS PBUTTONACTION 2 2>nul :: Configure Power Button to initiate shutdown (on power)
powercfg /setdcvalueindex %NEWGUID% SUB_BUTTONS PBUTTONACTION 2 2>nul :: Configure Power Button to initiate shutdown (on battery)
powercfg /S %NEWGUID% 2>nul

:: Import Accessibility niceties
echo Importing Accessibility niceties..
for %%F in ("%importDir%\Accessibility\*.reg") do reg import "%%~fF" > nul

:: Import WiFi Profiles
echo Importing WiFi Profiles
for /r "%importDir%\WiFiProfiles" %%w in (*.xml) do (
    echo Adding profile from: "%%~nxw"
    netsh wlan add profile filename="%%w" user=all > nul
)

:: Import Wallpapers & Install activIT Theme Pack
echo Importing Wallpaper data and downloading activIT Theme Pack
echo Downloading activIT Theme Pack..
curl -s -L -A "Mozilla/5.0" "https://www.aitsys.com.au/internal-use/AITSYS2023theme.deskthemepack" --output "C:\aitsys\aitsys.deskthemepack" > nul
C:\aitsys\aitsys.deskthemepack
"%localappdata%\Microsoft\Windows\Themes\AITSYS 20\AITSYS 20.theme"
echo Applied activIT Theme Pack.
if exist "%importDir%\Wallpaper\TranscodedWallpaper.jpg" (
	move "%importDir%\Wallpaper\TranscodedWallpaper.jpg" "C:\aitsys\Wallpaper.jpg" > nul
	powershell -command "set-itemproperty -path 'HKCU:Control Panel\Desktop' -name WallPaper -value C:\aitsys\Wallpaper.jpg" > nul
	powershell RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True > nul
	echo Updated wallpaper.
)

:: Import Downloads Folder
if exist "%importDir%\Downloads" (
	echo Importing Downloads Folder 
	robocopy "%importDir%\Downloads" "C:\Users\%uName%\Downloads" /E /MT:16 /R:3 /W:1 /XJ > nul 
)

:: Import third party Font files
echo Installing font files..
powershell -Command "Get-ChildItem '%importDir%\Fonts' -File | ForEach-Object { Copy-Item $_.FullName -Destination 'C:\Windows\Fonts' -Force; New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name $_.BaseName -PropertyType String -Value $_.Name -Force }" > nul

:: Import Printer Drivers and Configurations
if not exist "%importDir%\Printers" (
	echo No Printer Data found; skipping import.
	goto postMDT
) else (
	goto printerImpo
)

:: Import Printer Drivers/Configurations
:printerImpo
set DRIVER_DIR=%importDir%\Printers\Drivers
set CONFIG_FILE=%importDir%\Printers\PrinterConfig.csv
set PRINTER_CONFS=%importDir%\Printers\Configurations
set "PRINTER_LIST=%importDir%\Printers\printers.txt"

echo Importing Printer settings
echo Adding printer drivers to driver store...
for /R "%DRIVER_DIR%" %%f in (*.inf) do (
    echo Installing driver: %%f
    pnputil /add-driver "%%f" /install >nul 2>&1 || echo Failed to install %%f
)
echo Recreating printers..
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Csv '%CONFIG_FILE%' | ForEach-Object { if (-not (Get-Printer -Name $_.Name -ErrorAction SilentlyContinue)) { try { if (-not (Get-PrinterPort -Name $_.PortName -ErrorAction SilentlyContinue)) { if ($_.PortName -like 'IP_*') { $ip = $_.PortName -replace '^IP_', ''; Add-PrinterPort -Name $_.PortName -PrinterHostAddress $ip } } } catch {} try { if (-not (Get-PrinterDriver -Name $_.DriverName -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $_.DriverName } } catch {} try { Add-Printer -Name $_.Name -DriverName $_.DriverName -PortName $_.PortName; Write-Host ('Added printer: ' + $_.Name) } catch { Write-Host ('Failed to add printer: ' + $_.Name + ' Error: ' + $_.Exception.Message) } } else { Write-Host ('Printer already exists: ' + $_.Name) } }"
if not exist "%PRINTER_CONFS%" (
    echo No configurations folder found: %PRINTER_CONFS%
    goto postMDT
)

echo Importing Printer Settings..
set "RUNDLL=%SystemRoot%\System32\rundll32.exe"
if exist %SystemRoot%\Sysnative\rundll32.exe set "RUNDLL=%SystemRoot%\Sysnative\rundll32.exe"
for %%P in ("%PRINTER_CONFS%\*.dat") do (
	call :impoSpecificPrinter "%%~fP"
)

goto postMDT

:impoSpecificPrinter
if exist "%~1" (
    "%RUNDLL%" printui.dll,PrintUIEntry /Sr /q /n "%~n1" /a "%~1" m c d g >nul 2>&1
	if %errorlevel%==0 (
		echo Restored configuration for %~1
	) else (
		echo Error
	)
) else (
    echo No configuration file for "%~1"
)
exit /b

:postMDT
:: Disable Administrator users
echo Disabling Administrator users
if not exist "C:\aitsys" mkdir "C:\aitsys"
net localgroup Administrators > C:\aitsys\admin_members.txt
for /f "skip=6 tokens=*" %%a in ('type C:\aitsys\admin_members.txt') do (
    set "line=%%a"
    for /f "tokens=* delims=" %%b in ("!line!") do set "line=%%b"
    if "!line!"=="The command completed successfully." goto endPostMDT
    
	if not "!line!"=="" (
        if /I not "!line!"=="%username%" (
			if /I "!line:AzureAD=!"=="!line!" (
				echo Disabling user: !line!
				net user "!line!" /active:no
			)
        )
    )
)

:endPostMDT
del C:\aitsys\admin_members.txt
echo "Disabled all administrator users (excluding currently logged in)."

:: Configure English (AUS) as primary language pack
powershell Set-WinSystemLocale en-AU
powershell Set-WinUserLanguageList en-AU -Force
echo Configured English (Australia) as default language profile
:: Remove activIT WiFi profiles
netsh wlan delete profile name="activIT Systems" >nul 2>&1
netsh wlan delete profile name="activIT Systems - Guest" >nul 2>&1
netsh wlan delete profile name="activIT Systems - Workbench" >nul 2>&1
echo Removed activIT Systems WiFi Profiles
echo.
:: Announce script completion
echo Workstation Import complete. Closing script.
timeout /t 30
exit

:: ZIP Export
:zip
for /d %%G in ("C:\Users\%uName%\OneDrive - *") do set "OneDriveDir=%%~fG"
if exist "%OneDriveDir%" (
	echo OneDrive Directory Located: %OneDriveDir%
	%SystemRoot%\System32\choice.exe /C YN /N /M "Would you like to ZIP contents there instead? (y/n)"
	if errorlevel 2 set "backupOneDrive=n"
	if errorlevel 1 set "backupOneDrive=y"
)
for %%F in ("%worDir%") do set "zipName=%%~nF.zip"
if /I "%backupOneDrive%"=="y" (
	echo Creating .zip archive for %OneDriveDir%\%zipName%..
	cd "%OneDriveDir%"
	tar.exe -a -c -f  "%OneDriveDir%\%zipName%" -C "%worDir%" *
	
	if errorlevel 1 (
		echo [ERROR] tar.exe failed to create the archive. Aborting.
		exit /b 1
	)

	set "exportLoc=%OneDriveDir%\%zipName%"
	rmdir /s /q "%worDir%"
) 
if /I "%backupOneDrive%"=="n" (
	echo Creating .zip archive for %worDir%\%zipName%..
	cd "!worDir!\.."
	tar.exe -a -c -f  "%zipName%" -C "%worDir%" *
	set "exportLoc=%worDir%\%zipName%"
	cd ..
	
	if errorlevel 1 (
		echo [ERROR] tar.exe failed to create the archive. Aborting.
		exit /b 1
	)

	rmdir /s /q "%worDir%"
)

:: Announce export completion.

echo Export complete - export saved to %exportLoc%.
echo Ensure to copy the .zip created above to the new workstation and extract the archive prior to commencing the import portion of this script.
timeout /t 30
exit
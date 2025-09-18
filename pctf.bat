@echo off
setlocal enabledelayedexpansion
setlocal enableextensions
if "%1"=="printerExpo" (
    set "worDir=%2"
    echo Relaunched with elevated privileges...
    goto printerExpo
)
if "%1"=="postMDT" (
    echo Relaunched with elevated privileges...
    goto postMDT
)
if "%1"=="printerImpo" (
    echo Relaunched with elevated privileges...
	set "importDir=%2"
    goto printerImpo
)
echo    ___  _______________
echo   / _ \/ ___/_  __/ __/
echo  / ___/ /__  / / / _/   
echo /_/   \___/ /_/ /_/    
echo.
echo PC Transfer Tool v1.1.6 -- Developed by Samuel Bunko
echo.
echo 1. Import current user
echo 2. Import another user
echo 3. Export current user
echo 4. Export another user
echo.

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
	echo "Invalid option. Please select one of the options listed above (1, 2, 3, 4 or 5)."
	goto getOptions
)


:selectLocalUser
set /p "uName=Type the user account you wish to import: "
	if not exist "C:\Users\!uName!" (
		echo ERROR: C:\Users\!uName! could not be found; please try again.
		goto selectLocalUser
	) else (
		if exist "C:\Users\!uName!" (
			echo Working in directory C:\Users\!uName!..
			if "%opt%" equ "1" goto :importUser
			if "%opt%" equ "2" goto :importUser
			if "%opt%" equ "3" goto :exportUser
			if "%opt%" equ "4" goto :exportUser
		)
	)
	
:postMDT
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'postMDT' -Verb runAs"
    exit /b
)
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
echo "Disabled all administrator accounts (excluding currently logged in)."
powershell Set-WinSystemLocale en-AU
powershell Set-WinUserLanguageList en-AU -Force
echo Configured English (Australia) as default language profile
netsh wlan delete profile name="activIT Systems"
netsh wlan delete profile name="activIT Systems - Guest"
netsh wlan delete profile name="activIT Systems - Workbench"
echo Removed activIT Systems WiFi Profiles
echo.
echo Workstation Import complete. Closing script.
timeout /t 50
exit

:checkLocalUser
if /i "%userprofile%" equ "C:\Windows\system32\config\systemprofile" (
	set /p "uName=ERROR: User account not found - enter the desired username below: "
	if not exist "C:\Users\!uName!" (
		goto checkLocalUser
	) else (
		if exist "C:\Users\!uName!" (
			echo Working in directory C:\Users\!uName!..
			if "%opt%" equ "1" goto :importUser
			if "%opt%" equ "2" goto :importUser
			if "%opt%" equ "3" goto :exportUser
			if "%opt%" equ "4" goto :exportUser
		)
	)
) 
if exist "C:\Users\%USERNAME%\" (
		echo Working in directory %userprofile%..
		set uName=%USERNAME%
		if "%opt%" equ "1" goto :importUser
		if "%opt%" equ "2" goto :importUser
		if "%opt%" equ "3" goto :exportUser
		if "%opt%" equ "4" goto :exportUser
		PAUSE
	)
)

:exportDirectory
set /p worDir=Enter a directory you wish to export the user data to (or press ENTER to default to C:\aitsys\!uName!-%COMPUTERNAME%-Export): 
set "cleanDir=%worDir%"
for %%C in (^< ^> ^: ^" ^| ^? ^* " ") do (
    set "cleanDir=!cleanDir:%%C=!"
)
if "%worDir%"=="" (
    set "cleanDir=C:\aitsys\!uName!-%COMPUTERNAME%-Export"
)
if exist "%cleanDir%" (
    echo Working in directory %cleanDir%
    set "worDir=%cleanDir%"
    goto :exportUser
) else (
    mkdir "%cleanDir%" 2>nul
    if exist "%cleanDir%" (
        echo Created directory %cleanDir% for export
        set "worDir=%cleanDir%"
        goto :exportUser
    ) else (
        echo ERROR: Directory does not exist or contains invalid characters. Enter a valid directory to continue.
        goto :exportDirectory
    )
)


:exportUser
if not defined worDir goto :exportDirectory
%SystemRoot%\System32\choice.exe /C YN /N /M "Is the Downloads folder required? (y/n) "
if errorlevel 1 set dlReq=y
if errorlevel 2 set dlReq=n
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Signatures" "%worDir%\Signatures" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Templates" "%worDir%\Templates" /E
robocopy "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default" "%worDir%\Chrome" /E /XD "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default\Service Worker" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default\WebStorage" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default\Cache" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default\Code Cache" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default\IndexedDB"
robocopy "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default" "%worDir%\Edge" /E /XD "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default\Service Worker" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default\WebStorage" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default\Cache" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default\Code Cache" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default\IndexedDB"
copy "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\installs.ini" "%worDir%\Firefox" /E
copy "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\profiles.ini" "%worDir%\Firefox" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" "%worDir%\Firefox\Profiles" /E
if exist "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper" (
    robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes" "%worDir%\Wallpaper" /E
	cd %worDir%\Wallpaper
    ren TranscodedWallpaper TranscodedWallpaper.jpg
    if "%dlReq%" equ "y" (
	    robocopy "C:\Users\%uName%\Downloads" "%worDir%\Downloads" /E
    )
) else (
    echo No wallpaper found. Skipping.
)
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes" "%worDir%\Wallpaper" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" "%worDir%\TaskbarPins\Shortcuts" /E
mkdir %worDir%\TaskbarPins 2> nul
reg export "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" %worDir%\TaskbarPins\Taskband.reg
if not exist "%worDir%\WiFiProfiles" (
		mkdir "%worDir%\WiFiProfiles"
	)
netsh wlan export profile key=clear folder="%worDir%\WiFiProfiles"
goto printerExpo
exit

:importUser
set /p importDir=Enter the FULL filepath containing the exported user data: 
if not exist "!importDir!" (
	echo Unable to locate directory; try again
	goto importUser
	)
%SystemRoot%\System32\choice.exe /C YN /N /M "To import succesfully, all relevant applications must be closed. Force close applications? (y/n) "
if errorlevel 1 set killTasks=y
if errorlevel 2 set killTasks=n
if "%killTasks%" equ "y" (
	taskkill /f /im explorer.exe 2> nul
	taskkill /f /im chrome.exe 2> nul
	taskkill /f /im msedge.exe 2> nul
	taskkill /f /im firefox.exe 2> nul
	taskkill /f /im outlook.exe 2> nul
	taskkill /f /im winword.exe 2> nul
	taskkill /f /im excel.exe 2> nul
	taskkill /f /im powerpnt.exe 2> nul
) else (
	echo Unable to complete import; please save any open files within Chrome, Edge, Firefox, Outlook, Word, Excel or PowerPoint and try again.
	timeout /t 15
	exit
)
robocopy "%importDir%\Signatures" "C:\Users\%uName%\AppData\Roaming\Microsoft\Signatures" /E
robocopy "%importDir%\Templates" "C:\Users\%uName%\AppData\Roaming\Microsoft\Templates" /E
robocopy "%importDir%\Chrome" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default" /E
robocopy "%importDir%\Edge" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default" /E
copy "%importDir%\Firefox\installs.ini" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox" /E
copy "%importDir%\Firefox\profiles.ini" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox" /E
robocopy "%importDir%\Firefox\Profiles" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" /E
robocopy "%importDir%\TaskbarPins\Shortcuts" "C:\Users\%uName%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" /E
reg import "%importDir%\TaskbarPins\Taskband.reg"
start explorer.exe
ie4uinit.exe -ClearIconCache
for /r "%importDir%\WiFiProfiles" %%w in (*.xml) do (
    echo Adding profile from: "%%w"
    netsh wlan add profile filename="%%w" user=all
)
mkdir %importDir%\Wallpaper 2>nul
echo Downloading activIT Theme Pack..
curl -s -L -A "Mozilla/5.0" "https://www.aitsys.com.au/internal-use/AITSYS2023theme.deskthemepack" --output "%importDir%\Wallpaper\aitsys.deskthemepack"
%importDir%\Wallpaper\aitsys.deskthemepack
%localappdata%\Microsoft\Windows\Themes\AITSYS 20\AITSYS 20.theme
echo Applied activIT Theme Pack.
if exist "%importDir%\Wallpaper\TranscodedWallpaper.jpg" (
	move "%importDir%\Wallpaper\TranscodedWallpaper.jpg" "C:\aitsys\Wallpaper.jpg"
	powershell -command "set-itemproperty -path 'HKCU:Control Panel\Desktop' -name WallPaper -value C:\aitsys\Wallpaper.jpg"
	powershell RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True
	echo Updated wallpaper.
)
robocopy "%importDir%\Downloads" "C:\Users\%uName%\Downloads" /E

goto printerImpo
exit

:zip
set /p opt1=Zip the export? (y/n) 
if "%opt1%" equ "y" (
    set "zipName=!uName!_!COMPUTERNAME!.zip"
    for %%Z in (^< ^> ^: ^" ^| ^? ^* " ") do (
        set "zipName=!zipName:%%Z=!"
    )
    echo Creating .zip archive for !zipName!..
    cd !worDir!
    tar.exe -a -c -f  "!worDir!\..\!zipName!" Signatures Templates Chrome Edge Firefox Wallpaper TaskbarPins WiFiProfiles Downloads .
    cd ..
    rmdir /s /q "!worDir!"
    echo -=- Export complete -=-
    echo Ensure to copy the .zip created above to the new workstation and extract the archive prior to commencing the import portion of this script.
    timeout /t 15
)
exit

:printerExpo
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'printerExpo', '%worDir%' -Verb runAs"
    exit /b
)
set "DRIVER_DIR=%worDir%\Printers\Drivers"
set "USED_LIST=%worDir%\Printers\used_drivers.txt"
set "PRINTER_LIST=%worDir%\Printers\printers.txt"
set "CONFIG_FILE=%worDir%\Printers\PrinterConfig.csv"
echo Creating export folders..
mkdir "%DRIVER_DIR%" >nul
echo Exporting printer configurations..
powershell -Command "Get-Printer | Where-Object { $_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe' } | Export-Csv -Path \"%CONFIG_FILE%\" -NoTypeInformation"
echo Exporting third-party drivers..
powershell -Command "Get-PrinterDriver | Where-Object { $_.InfPath -and (Test-Path $_.InfPath) -and ($_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe') } | ForEach-Object { $source = Split-Path $_.InfPath -Parent; $dest = Join-Path \"%DRIVER_DIR%\" $_.Name; Copy-Item -Path $source -Destination $dest -Recurse -Force }"
echo All drivers exported to Drivers folder.
goto zip
exit

:printerImpo
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'printerImpo', '%importDir%' -Verb runAs"
    exit /b
)
set DRIVER_DIR=%importDir%\Printers\Drivers
set CONFIG_FILE=%importDir%\Printers\PrinterConfig.csv
echo Adding printer drivers to driver store...
for /R "%DRIVER_DIR%" %%f in (*.inf) do (
    echo Installing driver: %%f
    pnputil /add-driver "%%f" /install
)
echo Recreating printers from config...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Csv '%CONFIG_FILE%' | ForEach-Object { if (-not (Get-Printer -Name $_.Name -ErrorAction SilentlyContinue)) { try { if (-not (Get-PrinterPort -Name $_.PortName -ErrorAction SilentlyContinue)) { if ($_.PortName -like 'IP_*') { $ip = $_.PortName -replace '^IP_', ''; Add-PrinterPort -Name $_.PortName -PrinterHostAddress $ip } elseif ($_.PortName -like 'WSD*') { Add-PrinterPort -Name $_.PortName } else { Add-PrinterPort -Name $_.PortName } } } catch {} try { if (-not (Get-PrinterDriver -Name $_.DriverName -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $_.DriverName } } catch {} try { Add-Printer -Name $_.Name -DriverName $_.DriverName -PortName $_.PortName; Write-Host ('Added printer: ' + $_.Name) } catch { Write-Host ('Failed to add printer: ' + $_.Name + ' Error: ' + $_.Exception.Message) } } else { Write-Host ('Printer already exists: ' + $_.Name) } }"
goto postMDT

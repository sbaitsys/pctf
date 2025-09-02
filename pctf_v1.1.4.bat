@echo off
setlocal enabledelayedexpansion
echo    ___  _______________
echo   / _ \/ ___/_  __/ __/
echo  / ___/ /__  / / / _/   
echo /_/   \___/ /_/ /_/    
echo.
echo PC Transfer Tool v1.1.4 -- Developed by activIT Systems
echo.
echo 1. Import current user
echo 2. Import another user
echo 3. Export current user
echo 4. Export another user
echo 5. Perform post-MDT optimizations (run as administrator)
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
) else if "%opt%" equ "5" (
	goto postMDT
) else (
	echo "Invalid option. Please select one of the options listed above (1, 2, 3, 4 or 5)."
	goto getOptions
)

:selectLocalUser
set /p "uName=Type the user account you wish to export: "
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
echo Disabled all administrator accounts (excluding currently logged in).
powershell Set-WinSystemLocale en-AU
powershell Set-WinUserLanguageList en-AU -Force
echo Configured English (Australia) as default language profile
netsh wlan delete profile name="activIT Systems"
netsh wlan delete profile name="activIT Systems - Guest"
netsh wlan delete profile name="activIT Systems - Workbench"
echo Removed activIT Systems WiFi Profiles
timeout /t 15
exit

:checkLocalUser
if /i "%userprofile%" equ "C:\Windows\system32\config\systemprofile" (
	set /p "uName=ERROR: User account not found - enter the desired exported username below: "
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
robocopy "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data" "%worDir%\Chrome" /E
robocopy "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data" "%worDir%\Edge" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" "%worDir%\Firefox" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes" "%worDir%\Wallpaper" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" "%worDir%\TaskbarPins" /E
if not exist "%worDir%\WiFiProfiles" (
		mkdir "%worDir%\WiFiProfiles"
	)
netsh wlan export profile key=clear folder="%worDir%\WiFiProfiles"
cd %worDir%\Wallpaper
ren TranscodedWallpaper TranscodedWallpaper.jpg
if "%dlReq%" equ "y" (
	robocopy "C:\Users\%uName%\Downloads" "%worDir%\Downloads" /E
)

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
	exit
)
robocopy "%importDir%\Signatures" "C:\Users\%uName%\AppData\Roaming\Microsoft\Signatures" /E
robocopy "%importDir%\Templates" "C:\Users\%uName%\AppData\Roaming\Microsoft\Templates" /E
robocopy "%importDir%\Chrome" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data" /E
robocopy "%importDir%\Edge" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data" /E
robocopy "%importDir%\Firefox" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" /E
robocopy "%importDir%\TaskbarPins" "C:\Users\%uName%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" /E
for /r "%importDir%\WiFiProfiles" %%w in (*.xml) do (
    echo Adding profile from: "%%w"
    netsh wlan add profile filename="%%w" user=all
)
move "%importDir%\Wallpaper\TranscodedWallpaper.jpg" "C:\aitsys\Wallpaper.jpg"
powershell set-itemproperty -path "HKCU:Control Panel\Desktop" -name WallPaper -value C:\aitsys\Wallpaper.jpg
powershell RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True
robocopy "%importDir%\Downloads" "C:\Users\%uName%\Downloads" /E

start explorer.exe

echo Import complete. Closing script.
timeout /t 5
exit
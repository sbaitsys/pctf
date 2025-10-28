@echo off
setlocal enabledelayedexpansion
setlocal enableextensions
echo    ___  _______________
echo   / _ \/ ___/_  __/ __/
echo  / ___/ /__  / / / _/   
echo /_/   \___/ /_/ /_/    
echo.
echo PC Transfer Tool v1.1.7 -- Developed by Samuel Bunko
if "%1"=="printerExpo" (
    set "worDir=%2"
	set "uName=%3"
    echo Relaunched with elevated privileges...
    goto printerExpo
)
if "%1"=="importUser" (
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
set /p "uName=Type the username of the account: "
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
if "%worDir%"=="" (
    set "worDir=C:\aitsys\!uName!_%COMPUTERNAME%_Export"
)
set "cleanDir=%worDir%"
for %%C in (^< ^> ^" ^| ^? ^* " ") do (
    set "worDir=!cleanDir:%%C=!"
	if not exist "%worDir%" (
        echo Created directory %worDir% for export
        goto :exportUser
    )
	echo Working in directory %worDir%
	goto :exportUser
)

:exportUser
if not defined worDir goto :exportDirectory
%SystemRoot%\System32\choice.exe /C YN /N /M "Is the Downloads folder required? (y/n) "
if errorlevel 1 set dlReq=y
if errorlevel 2 set dlReq=n
echo Exporting Microsoft 365 data (templates, signatures, User MRU)
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Signatures" "%worDir%\Signatures" /E
robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Templates" "%worDir%\Templates" /E
if "%opt%" equ "3" (
	reg export "HKCU\Software\Microsoft\Office\16.0\Word\User MRU"  "%worDir%\Word_MRU.reg" /y
	reg export "HKCU\Software\Microsoft\Office\16.0\Excel\User MRU" "%worDir%\Excel_MRU.reg" /y
	reg export "HKCU\Software\Microsoft\Office\16.0\PowerPoint\User MRU" "%worDir%\PowerPoint_MRU.reg" /y
	reg export "HKCU\Software\Microsoft\Office\16.0\OneNote\User MRU" "%worDir%\OneNote_MRU.reg" /y    	
)
if "%opt%" equ "4" (
	for /f "delims=" %%S in ('powershell -NoProfile -Command "$p='C:\Users\!uName!'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.ProfileImagePath -ieq $p } | Select-Object -ExpandProperty PSChildName"') do set "SID=%%S"
	echo !SID!
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\Word\User MRU" "%worDir%\Word_MRU.reg" /y
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\Excel\User MRU" "%worDir%\Excel_MRU.reg" /y
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\PowerPoint\User MRU" "%worDir%\PowerPoint_MRU.reg" /y
	reg export "HKU\!SID!\Software\Microsoft\Office\16.0\OneNote\User MRU" "%worDir%\OneNote_MRU.reg" /y
)
echo Exporting Google Chrome data..
robocopy "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default" "%worDir%\Chrome" /E /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB"
echo Exporting Microsoft Edge data..
robocopy "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default" "%worDir%\Edge" /E /XD "Service Worker" "WebStorage" "Cache" "Code Cache" "IndexedDB"
echo Exporting Mozilla Firefox data..
copy /Y "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\installs.ini" "%worDir%\Firefox\installs.ini"
copy /Y "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\profiles.ini" "%worDir%\Firefox\profiles.ini"
robocopy "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" "%worDir%\Firefox\Profiles" /E /XD "shader-cache" "saved-telemetry-pings" "crashes"
if "%dlReq%" equ "y" (
    echo Exporting Downloads folder..
	robocopy "C:\Users\%uName%\Downloads" "%worDir%\Downloads" /E
)
echo Copying font files..
mkdir "%worDir%\Fonts"
powershell -NoProfile -Command "foreach ($f in Get-ChildItem 'C:\\Windows\\Fonts' -File) { if (@('8514fix.fon','8514fixe.fon','8514fixg.fon','8514fixr.fon','8514fixt.fon','8514oem.fon','8514oeme.fon','8514oemg.fon','8514oemr.fon','8514oemt.fon','8514sys.fon','8514syse.fon','8514sysg.fon','8514sysr.fon','8514syst.fon','85775.fon','85855.fon','85f1255.fon','85f1256.fon','85f1257.fon','85f874.fon','85s1255.fon','85s1256.fon','85s1257.fon','85s874.fon','AGENCYB.TTF','AGENCYR.TTF','ALGER.TTF','ANTQUAB.TTF','ANTQUABI.TTF','ANTQUAI.TTF','app775.fon','app850.fon','app852.fon','app855.fon','app857.fon','app866.fon','app932.fon','app936.fon','app949.fon','app950.fon','arial.ttf','arialbd.ttf','arialbi.ttf','ariali.ttf','ARIALN.TTF','ARIALNB.TTF','ARIALNBI.TTF','ARIALNI.TTF','ariblk.ttf','ARLRDBD.TTF','bahnschrift.ttf','BASKVILL.TTF','BAUHS93.TTF','BELL.TTF','BELLB.TTF','BELLI.TTF','BERNHC.TTF','BKANT.TTF','BOD_B.TTF','BOD_BI.TTF','BOD_BLAI.TTF','BOD_BLAR.TTF','BOD_CB.TTF','BOD_CBI.TTF','BOD_CI.TTF','BOD_CR.TTF','BOD_I.TTF','BOD_PSTC.TTF','BOD_R.TTF','BOOKOS.TTF','BOOKOSB.TTF','BOOKOSBI.TTF','BOOKOSI.TTF','BRADHITC.TTF','BRITANIC.TTF','BRLNSB.TTF','BRLNSDB.TTF','BRLNSR.TTF','BROADW.TTF','BRUSHSCI.TTF','BSSYM7.TTF','c8514fix.fon','c8514oem.fon','c8514sys.fon','calibri.ttf','calibrib.ttf','calibrii.ttf','calibril.ttf','calibrili.ttf','calibriz.ttf','CALIFB.TTF','CALIFI.TTF','CALIFR.TTF','CALIST.TTF','CALISTB.TTF','CALISTBI.TTF','CALISTI.TTF','cambria.ttc','cambriab.ttf','cambriai.ttf','cambriaz.ttf','Candara.ttf','Candarab.ttf','Candarai.ttf','Candaral.ttf','Candarali.ttf','Candaraz.ttf','CASTELAR.TTF','CENSCBK.TTF','CENTAUR.TTF','CENTURY.TTF','cga40737.fon','cga40850.fon','cga40852.fon','cga40857.fon','cga40866.fon','cga40869.fon','cga40woa.fon','cga80737.fon','cga80850.fon','cga80852.fon','cga80857.fon','cga80866.fon','cga80869.fon','cga80woa.fon','CHILLER.TTF','COLONNA.TTF','comic.ttf','comicbd.ttf','comici.ttf','comicz.ttf','consola.ttf','consolab.ttf','consolai.ttf','consolaz.ttf','constan.ttf','constanb.ttf','constani.ttf','constanz.ttf','COOPBL.TTF','COPRGTB.TTF','COPRGTL.TTF','corbel.ttf','corbelb.ttf','corbeli.ttf','corbell.ttf','corbelli.ttf','corbelz.ttf','coue1255.fon','coue1256.fon','coue1257.fon','couf1255.fon','couf1256.fon','couf1257.fon','cour.ttf','courbd.ttf','courbi.ttf','coure.fon','couree.fon','coureg.fon','courer.fon','couret.fon','courf.fon','courfe.fon','courfg.fon','courfr.fon','courft.fon','couri.ttf','CURLZ___.TTF','cvgafix.fon','cvgasys.fon','desktop.ini','DFHeiA1.ttf','DFKai71.ttf','DFMinC1.ttf','DFMinE1.ttf','dfmw5.ttf','DFPop91.ttf','dos737.fon','dos869.fon','dosapp.fon','DUBAI-BOLD.TTF','DUBAI-LIGHT.TTF','DUBAI-MEDIUM.TTF','DUBAI-REGULAR.TTF','ebrima.ttf','ebrimabd.ttf','ega40737.fon','ega40850.fon','ega40852.fon','ega40857.fon','ega40866.fon','ega40869.fon','ega40woa.fon','ega80737.fon','ega80850.fon','ega80852.fon','ega80857.fon','ega80866.fon','ega80869.fon','ega80woa.fon','ELEPHNT.TTF','ELEPHNTI.TTF','ENGR.TTF','ERASBD.TTF','ERASDEMI.TTF','ERASLGHT.TTF','ERASMD.TTF','FELIXTI.TTF','FORTE.TTF','FRABK.TTF','FRABKIT.TTF','FRADM.TTF','FRADMCN.TTF','FRADMIT.TTF','FRAHV.TTF','FRAHVIT.TTF','framd.ttf','FRAMDCN.TTF','framdit.ttf','FREE3OF9.TTF','FREESCPT.TTF','FRSCRIPT.TTF','FTLTLT.TTF','Gabriola.ttf','gadugi.ttf','gadugib.ttf','GARA.TTF','GARABD.TTF','GARAIT.TTF','georgia.ttf','georgiab.ttf','georgiai.ttf','georgiaz.ttf','GIGI.TTF','GILBI___.TTF','GILB____.TTF','GILC____.TTF','GILI____.TTF','GILLUBCD.TTF','GILSANUB.TTF','GIL_____.TTF','GLECB.TTF','GlobalMonospace.CompositeFont','GlobalSansSerif.CompositeFont','GlobalSerif.CompositeFont','GlobalUserInterface.CompositeFont','GLSNECB.TTF','GOTHIC.TTF','GOTHICB.TTF','GOTHICBI.TTF','GOTHICI.TTF','GOUDOS.TTF','GOUDOSB.TTF','GOUDOSI.TTF','GOUDYSTO.TTF','h8514fix.fon','h8514oem.fon','h8514sys.fon','HARLOWSI.TTF','HARNGTON.TTF','HATTEN.TTF','himalaya.ttf','HTOWERT.TTF','HTOWERTI.TTF','hvgafix.fon','hvgasys.fon','impact.ttf','IMPRISHA.TTF','INFROMAN.TTF','Inkfree.ttf','ITCBLKAD.TTF','ITCEDSCR.TTF','ITCKRIST.TTF','j8514fix.fon','j8514oem.fon','j8514sys.fon','javatext.ttf','JOKERMAN.TTF','jsmalle.fon','jsmallf.fon','JUICE___.TTF','jvgafix.fon','jvgasys.fon','KUNSTLER.TTF','LATINWD.TTF','LBRITE.TTF','LBRITED.TTF','LBRITEDI.TTF','LBRITEI.TTF','LCALLIG.TTF','LeelaUIb.ttf','LEELAWAD.TTF','LEELAWDB.TTF','LeelawUI.ttf','LeelUIsl.ttf','LFAX.TTF','LFAXD.TTF','LFAXDI.TTF','LFAXI.TTF','LHANDW.TTF','LSANS.TTF','LSANSD.TTF','LSANSDI.TTF','LSANSI.TTF','LTYPE.TTF','LTYPEB.TTF','LTYPEBO.TTF','LTYPEO.TTF','lucon.ttf','l_10646.ttf','MAGNETOB.TTF','MAIAN.TTF','malgun.ttf','malgunbd.ttf','malgunsl.ttf','marlett.ttf','MATURASC.TTF','micross.ttf','mingliub.ttc','MISTRAL.TTF','mmrtext.ttf','mmrtextb.ttf','MOD20.TTF','modern.fon','monbaiti.ttf','msgothic.ttc','msjh.ttc','msjhbd.ttc','msjhl.ttc','MSUIGHUB.TTF','MSUIGHUR.TTF','msyh.ttc','msyhbd.ttc','msyhl.ttc','msyi.ttf','MTCORSVA.TTF','MTEXTRA.TTF','mvboli.ttf','NIAGENG.TTF','NIAGSOL.TTF','Nirmala.ttc','ntailu.ttf','ntailub.ttf','OCR-a___.ttf','OCR-b___.ttf','OCRAEXT.TTF','OLDENGL.TTF','ONYX.TTF','OUTLOOK.TTF','pala.ttf','palab.ttf','palabi.ttf','palai.ttf','PALSCRI.TTF','PAPYRUS.TTF','PARCHM.TTF','PERBI___.TTF','PERB____.TTF','PERI____.TTF','PERTIBD.TTF','PERTILI.TTF','PER_____.TTF','phagspa.ttf','phagspab.ttf','PLAYBILL.TTF','POORICH.TTF','PRISTINA.TTF','RAGE.TTF','RAVIE.TTF','REFSAN.TTF','REFSPCL.TTF','ROCCB___.TTF','ROCC____.TTF','ROCK.TTF','ROCKB.TTF','ROCKBI.TTF','ROCKEB.TTF','ROCKI.TTF','roman.fon','s8514fix.fon','s8514oem.fon','s8514sys.fon','SansSerifCollection.ttf','SCHLBKB.TTF','SCHLBKBI.TTF','SCHLBKI.TTF','script.fon','SCRIPTBL.TTF','segmdl2.ttf','SegoeIcons.ttf','segoepr.ttf','segoeprb.ttf','segoesc.ttf','segoescb.ttf','segoeui.ttf','segoeuib.ttf','segoeuii.ttf','segoeuil.ttf','segoeuisl.ttf','segoeuiz.ttf','seguibl.ttf','seguibli.ttf','seguiemj.ttf','seguihis.ttf','seguili.ttf','seguisb.ttf','seguisbi.ttf','seguisli.ttf','seguisym.ttf','SegUIVar.ttf','sere1255.fon','sere1256.fon','sere1257.fon','serf1255.fon','serf1256.fon','serf1257.fon','serife.fon','serifee.fon','serifeg.fon','serifer.fon','serifet.fon','seriff.fon','seriffe.fon','seriffg.fon','seriffr.fon','serifft.fon','SHOWG.TTF','simsun.ttc','simsunb.ttf','SimsunExtG.ttf','SitkaVF-Italic.ttf','SitkaVF.ttf','smae1255.fon','smae1256.fon','smae1257.fon','smaf1255.fon','smaf1256.fon','smaf1257.fon','smalle.fon','smallee.fon','smalleg.fon','smaller.fon','smallet.fon','smallf.fon','smallfe.fon','smallfg.fon','smallfr.fon','smallft.fon','SNAP____.TTF','ssee1255.fon','ssee1256.fon','ssee1257.fon','ssee874.fon','ssef1255.fon','ssef1256.fon','ssef1257.fon','ssef874.fon','sserife.fon','sserifee.fon','sserifeg.fon','sserifer.fon','sserifet.fon','sseriff.fon','sseriffe.fon','sseriffg.fon','sseriffr.fon','sserifft.fon','STENCIL.TTF','svgafix.fon','svgasys.fon','sylfaen.ttf','symbol.ttf','tahoma.ttf','tahomabd.ttf','taile.ttf','taileb.ttf','TCBI____.TTF','TCB_____.TTF','TCCB____.TTF','TCCEB.TTF','TCCM____.TTF','TCMI____.TTF','TCM_____.TTF','TEMPSITC.TTF','times.ttf','timesbd.ttf','timesbi.ttf','timesi.ttf','trebuc.ttf','trebucbd.ttf','trebucbi.ttf','trebucit.ttf','tt0001m_.ttf','tt0002m_.ttf','tt0003c_.ttf','tt0003m_.ttf','tt0004c_.ttf','tt0004m_.ttf','tt0005c_.ttf','tt0005m_.ttf','tt0006c_.ttf','tt0006m_.ttf','tt0007m_.ttf','tt0009m_.ttf','tt0010m_.ttf','tt0035m_.ttf','tt0036m_.ttf','tt0037m_.ttf','tt0038m_.ttf','tt0047m_.ttf','tt0048m_.ttf','tt0049m_.ttf','tt0050m_.ttf','tt0102m_.ttf','tt0132m_.ttf','tt0140m_.ttf','tt0141m_.ttf','tt0142m_.ttf','tt0143m_.ttf','tt0144m_.ttf','tt0145m_.ttf','tt0171m_.ttf','tt0172m_.ttf','tt0173m_.ttf','tt0246m_.ttf','tt0247m_.ttf','tt0248m_.ttf','tt0249m_.ttf','tt0282m_.ttf','tt0283m_.ttf','tt0284m_.ttf','tt0288m_.ttf','tt0289m_.ttf','tt0290m_.ttf','tt0291m_.ttf','tt0292m_.ttf','tt0293m_.ttf','tt0308m_.ttf','tt0309m_.ttf','tt0310m_.ttf','tt0311m_.ttf','tt0319m_.ttf','tt0320m_.ttf','tt0351m_.ttf','tt0365m_.ttf','tt0371m_.ttf','tt0375m_.ttf','tt0503m_.ttf','tt0504m_.ttf','tt0524m_.ttf','tt0586m_.ttf','tt0588m_.ttf','tt0627m_.ttf','tt0628m_.ttf','tt0663m_.ttf','tt0726m_.ttf','tt0855m_.ttf','tt0857m_.ttf','tt0861m_.ttf','tt0868m_.ttf','tt0869m_.ttf','tt0939m_.ttf','tt1018m_.ttf','tt1019m_.ttf','tt1020m_.ttf','tt1041m_.ttf','tt1057m_.ttf','tt1106m_.ttf','tt1107m_.ttf','tt1159m_.ttf','tt1160m_.ttf','tt1161m_.ttf','tt1180m_.ttf','tt1181m_.ttf','tt1182m_.ttf','tt1183m_.ttf','tt1184m_.ttf','tt1185m_.ttf','tt6804m_.ttf','tt6805m_.ttf','tt6806m_.ttf','tt6807m_.ttf','verdana.ttf','verdanab.ttf','verdanai.ttf','verdanaz.ttf','vga737.fon','vga775.fon','vga850.fon','vga852.fon','vga855.fon','vga857.fon','vga860.fon','vga861.fon','vga863.fon','vga865.fon','vga866.fon','vga869.fon','vga932.fon','vga936.fon','vga949.fon','vga950.fon','vgaf1255.fon','vgaf1256.fon','vgaf1257.fon','vgaf874.fon','vgafix.fon','vgafixe.fon','vgafixg.fon','vgafixr.fon','vgafixt.fon','vgaoem.fon','vgas1255.fon','vgas1256.fon','vgas1257.fon','vgas874.fon','vgasys.fon','vgasyse.fon','vgasysg.fon','vgasysr.fon','vgasyst.fon','VINERITC.TTF','VIVALDII.TTF','VLADIMIR.TTF','webdings.ttf','wingding.ttf','WINGDNG2.TTF','WINGDNG3.TTF','YuGothB.ttc','YuGothL.ttc','YuGothM.ttc','YuGothR.ttc') -notcontains $f.Name.ToLower()) { Copy-Item -LiteralPath $f.FullName -Destination \"%worDir%\\Fonts\" -Force } }"
if exist "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper" (
    echo Exporting Wallpaper..
	robocopy "C:\Users\%uName%\AppData\Roaming\Microsoft\Windows\Themes" "%worDir%\Wallpaper" /E > nul
	cd %worDir%\Wallpaper
    ren TranscodedWallpaper TranscodedWallpaper.jpg
) else (
    echo No wallpaper found. Skipping.
)
if not exist "%worDir%\WiFiProfiles" (
		mkdir "%worDir%\WiFiProfiles"
	)
echo Exporting WiFi profiles..
netsh wlan export profile key=clear folder="%worDir%\WiFiProfiles" > nul
goto printerExpo
exit

:importUser
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'importUser', '%uName%', '%opt%' -Verb runAs"
    exit /b
)
set /p importDir=Enter the FULL filepath containing the exported user data: 
if not exist "!importDir!" (
	echo Unable to locate directory; try again
	goto importUser
	)
%SystemRoot%\System32\choice.exe /C YN /N /M "To import succesfully, all relevant applications must be closed. Force close applications? (y/n) "
if errorlevel 1 set killTasks=y
if errorlevel 2 set killTasks=n
if "%killTasks%" equ "y" (
	taskkill /f /im chrome.exe 2> nul
	taskkill /f /im msedge.exe 2> nul
	taskkill /f /im firefox.exe 2> nul
	taskkill /f /im outlook.exe 2> nul
	taskkill /f /im winword.exe 2> nul
	taskkill /f /im excel.exe 2> nul
	taskkill /f /im powerpnt.exe 2> nul
) else (
	echo Unable to complete import; please save any open files within Chrome, Edge, Firefox, Outlook, Word, Excel or PowerPoint and try again.
)
echo Importing Microsoft 365 data (templates, signatures, User MRU)
robocopy "%importDir%\Signatures" "C:\Users\%uName%\AppData\Roaming\Microsoft\Signatures" /E
robocopy "%importDir%\Templates" "C:\Users\%uName%\AppData\Roaming\Microsoft\Templates" /E
if "%opt%" equ "1" (
	reg import "%importDir%\Word_MRU.reg"
	reg import "%importDir%\Excel_MRU.reg"
	reg import "%importDir%\PowerPoint_MRU.reg"
	reg import "%importDir%\OneNote_MRU.reg"
)
if "%opt%" equ "2" (
    reg load "HKU\TempHive" "C:\Users\%uName%\NTUSER.DAT"
	reg import "%importDir%\Word_MRU.reg"
	reg import "%importDir%\Excel_MRU.reg"
	reg import "%importDir%\PowerPoint_MRU.reg"
	reg import "%importDir%\OneNote_MRU.reg"
	reg unload "HKU\TempHive"
)
echo Importing Google Chrome data
robocopy "%importDir%\Chrome" "C:\Users\%uName%\AppData\Local\Google\Chrome\User Data\Default" /E
echo Importing Microsoft Edge data
robocopy "%importDir%\Edge" "C:\Users\%uName%\AppData\Local\Microsoft\Edge\User Data\Default" /E
echo Importing Mozilla Firefox data
copy "%importDir%\Firefox\installs.ini" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox" /E
copy "%importDir%\Firefox\profiles.ini" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox" /E
robocopy "%importDir%\Firefox\Profiles" "C:\Users\%uName%\AppData\Roaming\Mozilla\Firefox\Profiles" /E
echo Importing WiFi Profiles
for /r "%importDir%\WiFiProfiles" %%w in (*.xml) do (
    echo Adding profile from: "%%w"
    netsh wlan add profile filename="%%w" user=all
)
echo Importing Wallpaper data & downloading activIT Theme Pack
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
if exist "%importDir%\Downloads" (
	echo Importing Downloads folder 
	robocopy "%importDir%\Downloads" "C:\Users\%uName%\Downloads" /E
)
goto printerImpo
exit

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

:printerExpo
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -ArgumentList 'printerExpo', '%worDir%', '%uName%' -Verb runAs"
    exit /b
)
set "DRIVER_DIR=%worDir%\Printers\Drivers"
set "USED_LIST=%worDir%\Printers\used_drivers.txt"
set "PRINTER_LIST=%worDir%\Printers\printers.txt"
set "CONFIG_FILE=%worDir%\Printers\PrinterConfig.csv"
echo Creating export folders..
mkdir "%DRIVER_DIR%" >nul
echo Exporting printer configurations..
powershell -Command "Get-Printer | Where-Object { $_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe|DYMO' } | Export-Csv -Path \"%CONFIG_FILE%\" -NoTypeInformation"
echo Exporting third-party drivers..
powershell -Command "Get-PrinterDriver | Where-Object { $_.InfPath -and (Test-Path $_.InfPath) -and ($_.Name -notmatch 'Microsoft|PDF|OneNote|Remote Desktop|Adobe|DYMO') } | ForEach-Object { $source = Split-Path $_.InfPath -Parent; $dest = Join-Path \"%DRIVER_DIR%\" $_.Name; Copy-Item -Path $source -Destination $dest -Recurse -Force }"
echo All drivers exported to Drivers folder.
goto zip
exit

:printerImpo
set DRIVER_DIR=%importDir%\Printers\Drivers
set CONFIG_FILE=%importDir%\Printers\PrinterConfig.csv
echo Adding printer drivers to driver store...
for /R "%DRIVER_DIR%" %%f in (*.inf) do (
    echo Installing driver: %%f
    pnputil /add-driver "%%f" /install
)
echo Recreating printers from config...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Csv '%CONFIG_FILE%' | ForEach-Object { if (-not (Get-Printer -Name $_.Name -ErrorAction SilentlyContinue)) { try { if (-not (Get-PrinterPort -Name $_.PortName -ErrorAction SilentlyContinue)) { if ($_.PortName -like 'IP_*') { $ip = $_.PortName -replace '^IP_', ''; Add-PrinterPort -Name $_.PortName -PrinterHostAddress $ip } } } catch {} try { if (-not (Get-PrinterDriver -Name $_.DriverName -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $_.DriverName } } catch {} try { Add-Printer -Name $_.Name -DriverName $_.DriverName -PortName $_.PortName; Write-Host ('Added printer: ' + $_.Name) } catch { Write-Host ('Failed to add printer: ' + $_.Name + ' Error: ' + $_.Exception.Message) } } else { Write-Host ('Printer already exists: ' + $_.Name) } }"
echo Installing font files..
powershell -Command "Get-ChildItem '%importDir%\Fonts' -File | ForEach-Object { Copy-Item $_.FullName -Destination 'C:\Windows\Fonts' -Force; New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name $_.BaseName -PropertyType String -Value $_.Name -Force }"
goto postMDT

:zip
for %%F in ("%worDir%") do set "zipName=%%~nF.zip"
echo Creating .zip archive for !zipName!..
cd !worDir!\..
tar.exe -a -c -f  "!worDir!.zip" -C "%worDir%" *
cd ..
rmdir /s /q "!worDir!"
echo Export complete.
echo Ensure to copy the .zip created above to the new workstation and extract the archive prior to commencing the import portion of this script.
timeout /t 15
)
exit
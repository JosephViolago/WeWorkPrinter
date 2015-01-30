@echo off

setlocal
cd /d %~dp0

Call :DownloadFile "http://weprint.wework.com/Konica-Drivers.zip"   "C:\Temp\Konica-Drivers.zip"
Call :DownloadFile "http://weprint.wework.com/PaperCut-Windows.zip" "C:\Temp\PaperCut-Windows.zip"

Call :UnZipFile "C:\Temp\" "C:\Temp\Konica-Drivers.zip"
Call :UnZipFile "C:\Temp\" "C:\Temp\PaperCut-Windows.zip"

"C:\Temp\PaperCut Windows\SETUP.exe" /S

:: Install Driver
rundll32 printui.dll,PrintUIEntry /ia /m "KONICA MINOLTA C364SeriesFax" /f "C:\Temp\Konica Drivers\Windows 7 and 8\Drivers\Fax\EN\Win_x64\KOAYQS__.INF"
rundll32 printui.dll,PrintUIEntry /ia /m "KONICA MINOLTA C364SeriesPCL" /f "C:\Temp\Konica Drivers\Windows 7 and 8\Drivers\PCL\EN\Win_x64\KOAYQJ__.INF"
rundll32 printui.dll,PrintUIEntry /ia /m "KONICA MINOLTA C364SeriesPS"  /f "C:\Temp\Konica Drivers\Windows 7 and 8\Drivers\PS\EN\Win_x64\KOAYQA__.INF"

:: Add printer
rundll32 printui.dll,PrintUIEntry /b "WeWork" /n "KONICA MINOLTA C364 at WeWork" /ii /f "C:\Temp\Konica Drivers\Windows 7 and 8\Drivers\PS\EN\Win_x64\KOAYQA__.INF" /r "ipp://p.wework.com/printers/WeWork" /m "KONICA MINOLTA C364Series" /z

:: DEBUG; Get printui options window
:: rundll32 printui.dll

Call :Message "Congratulations! You have completed the installation. Now go waste some paper!"

exit /b


::--------------------------------------------------------
::-- Function section starts below here
::--------------------------------------------------------

:DownloadFile <DownloadFrom> <newzipfile>
If Not Exist %2 (
    bitsadmin.exe /transfer WeWorkPrinter /download /priority normal %1 %2
)
goto:eof

:Message <message>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>>%vbs% echo MsgBox %1
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%
goto:eof

:UnZipFile <ExtractTo> <newzipfile>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>>%vbs% echo Set fso = CreateObject("Scripting.FileSystemObject")
>>%vbs% echo If NOT fso.FolderExists(%1) Then
>>%vbs% echo fso.CreateFolder(%1)
>>%vbs% echo End If
>>%vbs% echo set objShell = CreateObject("Shell.Application")
>>%vbs% echo set FilesInZip=objShell.NameSpace(%2).items
>>%vbs% echo objShell.NameSpace(%1).CopyHere(FilesInZip)
>>%vbs% echo Set fso = Nothing
>>%vbs% echo Set objShell = Nothing
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%
goto:eof


::--------------------------------------------------------
::-- Resources
::--------------------------------------------------------
::-- Unzip
:: http://stackoverflow.com/questions/21704041
::-- Printer Administration
:: http://www.tech-recipes.com/rx/45529/install-network-printers-via-batch-file-or-command-line-in-windows-78-and-server-2008/
:: https://www.novell.com/coolsolutions/trench/607.html
:: http://smallbusiness.chron.com/delete-tcp-ip-printer-ports-via-command-line-56676.html
::-- Message
:: http://stackoverflow.com/questions/774175

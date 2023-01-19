@echo off
:start
cls
color 0c
echo       MENU
echo.
echo 1. Afficher IP
echo 2. Afficher Partage
echo 3. Afficher Programme
echo 4. Fermer
echo.
echo.
echo.
set /p choix=Quelle option ?
(
if not %choix%=='' set choix=%choix:~0,1%
if %choix%==1 goto ip
if %choix%==2 goto partage
if %choix%==3 goto programme
if %choix%==4 goto fin
)
goto start
:ip

C:\Windows\system32\cscript.exe ##PATH HERE
goto start
:partage
C:\Windows\system32\cscript.exe ##PATH HERE
goto start
:programme
C:\Windows\system32\cscript.exe ##PATH HERE
goto start
:pause
:fin
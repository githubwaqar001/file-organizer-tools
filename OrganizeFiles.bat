@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title File Organizer
color 0A

echo ============================================
echo          FILE ORGANIZER TOOL
echo ============================================
echo.
echo  Current folder: %CD%
echo.
echo  Choose a sorting method:
echo.
echo   [1] By Extension    (e.g. Images, Documents, Videos...)
echo   [2] By Date         (e.g. 2024-01, 2024-02...)
echo   [3] By First Letter (e.g. A, B, C...)
echo   [4] By Size         (e.g. Tiny, Small, Medium, Large, Huge)
echo   [5] Exit
echo.
echo ============================================

set /p "choice=  Enter your choice (1-5): "

if "%choice%"=="1" goto SortByExtension
if "%choice%"=="2" goto SortByDate
if "%choice%"=="3" goto SortByName
if "%choice%"=="4" goto SortBySize
if "%choice%"=="5" goto End
echo  Invalid choice. Please try again.
timeout /t 2 >nul
goto :eof

REM ============================================
REM  SORT BY FILE EXTENSION
REM ============================================
:SortByExtension
echo.
echo  Sorting by file extension...
echo.

set "count=0"
for %%F in (*.*) do (
    if /i not "%%~nxF"=="%~nx0" (
        set "ext=%%~xF"
        if defined ext (
            set "ext=!ext:~1!"
            call :GetCategoryForExt "!ext!" folder
            if not exist "!folder!" mkdir "!folder!"
            move "%%F" "!folder!\" >nul 2>&1
            if !errorlevel! equ 0 (
                echo   Moved: %%~nxF  --^>  !folder!\
                set /a count+=1
            )
        )
    )
)
goto Done

:GetCategoryForExt
set "_ext=%~1"
set "_ext=!_ext: =!"

for %%E in (jpg jpeg png gif bmp svg webp ico tiff tif) do if /i "!_ext!"=="%%E" set "%2=Images" & goto :eof
for %%E in (mp4 avi mkv mov wmv flv webm m4v mpg mpeg) do if /i "!_ext!"=="%%E" set "%2=Videos" & goto :eof
for %%E in (mp3 wav flac aac ogg wma m4a) do if /i "!_ext!"=="%%E" set "%2=Audio" & goto :eof
for %%E in (doc docx pdf txt rtf odt xls xlsx csv ppt pptx) do if /i "!_ext!"=="%%E" set "%2=Documents" & goto :eof
for %%E in (zip rar 7z tar gz bz2 xz iso cab) do if /i "!_ext!"=="%%E" set "%2=Archives" & goto :eof
for %%E in (exe msi bat cmd ps1 com scr) do if /i "!_ext!"=="%%E" set "%2=Programs" & goto :eof
for %%E in (py js ts html css java c cpp cs go rb php sql sh json xml yaml yml toml ini cfg conf) do if /i "!_ext!"=="%%E" set "%2=Code" & goto :eof
for %%E in (ttf otf woff woff2 eot) do if /i "!_ext!"=="%%E" set "%2=Fonts" & goto :eof
for %%E in (dll sys drv inf reg dat log bak tmp) do if /i "!_ext!"=="%%E" set "%2=System" & goto :eof
for %%E in (psd ai eps indd xd fig sketch) do if /i "!_ext!"=="%%E" set "%2=Design" & goto :eof
for %%E in (torrent lnk url) do if /i "!_ext!"=="%%E" set "%2=Shortcuts" & goto :eof
set "%2=Other"
goto :eof

REM ============================================
REM  SORT BY DATE (Year-Month)
REM ============================================
:SortByDate
echo.
echo  Sorting by last modified date...
echo.

set "count=0"
for %%F in (*.*) do (
    if /i not "%%~nxF"=="%~nx0" (
        set "filedate=%%~tF"
        REM Parse date - handles DD/MM/YYYY and MM/DD/YYYY via generic split
        for /f "tokens=1-3 delims=/ " %%A in ("!filedate!") do (
            set "part1=%%A"
            set "part2=%%B"
            set "part3=%%C"
        )
        REM Detect format: if part3 is 4-digit year, use it; otherwise try part1
        set "yr=!part3!"
        if "!yr:~0,2!"=="20" (
            set "folder=!yr!-!part2!"
        ) else if "!yr:~0,2!"=="19" (
            set "folder=!yr!-!part2!"
        ) else (
            REM Fallback: try part1 as year
            set "yr=!part1!"
            if "!yr:~0,2!"=="20" (
                set "folder=!yr!-!part2!"
            ) else (
                set "folder=Unknown-Date"
            )
        )
        if not exist "!folder!" mkdir "!folder!"
        move "%%F" "!folder!\" >nul 2>&1
        if !errorlevel! equ 0 (
            echo   Moved: %%~nxF  --^>  !folder!\
            set /a count+=1
        )
    )
)
goto Done

REM ============================================
REM  SORT BY FIRST LETTER
REM ============================================
:SortByName
echo.
echo  Sorting by first letter of filename...
echo.

set "count=0"
for %%F in (*.*) do (
    if /i not "%%~nxF"=="%~nx0" (
        set "fname=%%~nF"
        set "letter=!fname:~0,1!"

        REM Check if it starts with a number
        echo !letter!| findstr /r "[0-9]" >nul 2>&1
        if !errorlevel! equ 0 (
            set "folder=#Numbers"
        ) else (
            REM Convert to uppercase for consistency
            for %%L in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
                if /i "!letter!"=="%%L" set "folder=%%L"
            )
            if not defined folder set "folder=#Symbols"
        )

        if not exist "!folder!" mkdir "!folder!"
        move "%%F" "!folder!\" >nul 2>&1
        if !errorlevel! equ 0 (
            echo   Moved: %%~nxF  --^>  !folder!\
            set /a count+=1
        )
        set "folder="
    )
)
goto Done

REM ============================================
REM  SORT BY FILE SIZE
REM ============================================
:SortBySize
echo.
echo  Sorting by file size...
echo   Tiny     = under 10 KB
echo   Small    = 10 KB - 1 MB
echo   Medium   = 1 MB - 100 MB
echo   Large    = 100 MB - 1 GB
echo   Huge     = over 1 GB
echo.

set "count=0"
for %%F in (*.*) do (
    if /i not "%%~nxF"=="%~nx0" (
        set "fsize=%%~zF"

        if !fsize! LSS 10240 (
            set "folder=Tiny (under 10KB)"
        ) else if !fsize! LSS 1048576 (
            set "folder=Small (10KB-1MB)"
        ) else if !fsize! LSS 104857600 (
            set "folder=Medium (1MB-100MB)"
        ) else if !fsize! LSS 1073741824 (
            set "folder=Large (100MB-1GB)"
        ) else (
            set "folder=Huge (over 1GB)"
        )

        if not exist "!folder!" mkdir "!folder!"
        move "%%F" "!folder!\" >nul 2>&1
        if !errorlevel! equ 0 (
            echo   Moved: %%~nxF  --^>  !folder!\
            set /a count+=1
        )
    )
)
goto Done

REM ============================================
REM  DONE
REM ============================================
:Done
echo.
echo ============================================
echo  Done! Moved !count! file(s).
echo ============================================
echo.
pause
goto :eof

:End
echo.
echo  Goodbye!
timeout /t 1 >nul
goto :eof

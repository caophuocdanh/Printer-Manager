@echo off
title Build Printer Registry Manager
color 0A

:: --- CẤU HÌNH ---
set PYTHON_FILE=PrinterManager.py
set EXE_NAME=PrinterManager
set ICON_FILE=printer.ico

cls
echo ========================================================
echo      TOOL DONG GOI EXE - PRINTER REGISTRY MANAGER
echo ========================================================
echo.

:: 1. KIỂM TRA FILE NGUỒN
if not exist "%PYTHON_FILE%" (
    color 0C
    echo [LOI] Khong tim thay file ma nguon: %PYTHON_FILE%
    echo Vui long kiem tra lai ten file.
    pause
    exit
)

if not exist "%ICON_FILE%" (
    color 0E
    echo [CANH BAO] Khong tim thay file icon: %ICON_FILE%
    echo Chuong trinh se build voi icon mac dinh.
    echo.
)

:: 2. DỌN DẸP TRƯỚC KHI BUILD
echo [BUOC 1/4] Dang don dep file rac cu...
if exist build (
    echo   - Xoa thu muc 'build'...
    rmdir /s /q build
)
if exist dist (
    echo   - Xoa thu muc 'dist'...
    rmdir /s /q dist
)
if exist *.spec (
    echo   - Xoa file '.spec'...
    del /q *.spec
)
if exist __pycache__ (
    echo   - Xoa cache Python...
    rmdir /s /q __pycache__
)
echo [OK] Da don dep xong.
echo.

:: 3. THỰC HIỆN BUILD
echo [BUOC 2/4] Dang chay PyInstaller...
echo Lenh: pyinstaller --noconsole --onefile --icon="%ICON_FILE%" --add-data "%ICON_FILE%;." --name="%EXE_NAME%" "%PYTHON_FILE%"
echo --------------------------------------------------------

pyinstaller --noconsole --onefile --icon="%ICON_FILE%" --add-data "%ICON_FILE%;." --name="%EXE_NAME%" "%PYTHON_FILE%"

echo --------------------------------------------------------
if %errorlevel% neq 0 (
    color 0C
    echo [LOI] Qua trinh Build gap su co! Vui long kiem tra lai code.
    pause
    exit
)
echo [OK] Build thanh cong.
echo.

:: 4. DI CHUYỂN FILE EXE VÀ DỌN DẸP SAU KHI BUILD
echo [BUOC 3/4] Dang xu ly file thanh pham...

if exist "dist\%EXE_NAME%.exe" (
    echo   - Di chuyen '%EXE_NAME%.exe' ra thu muc goc...
    move /y "dist\%EXE_NAME%.exe" . >nul
) else (
    echo [LOI] Khong tim thay file EXE trong thu muc dist!
    pause
    exit
)

echo.
echo [BUOC 4/4] Don dep file tam sau khi build...
if exist build (
    echo   - Xoa thu muc 'build'...
    rmdir /s /q build
)
if exist dist (
    echo   - Xoa thu muc 'dist'...
    rmdir /s /q dist
)
if exist "%EXE_NAME%.spec" (
    echo   - Xoa file '%EXE_NAME%.spec'...
    del /q "%EXE_NAME%.spec"
)

echo.
echo ========================================================
echo        HOAN TAT! FILE EXE DA SAN SANG.
echo ========================================================
echo File cua ban: %CD%\%EXE_NAME%.exe
echo.
pause
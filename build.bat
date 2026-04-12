@echo off
title SGPilot - Build EXE
echo.
echo  ========================================
echo   SGPilot - Gerando executavel...
echo  ========================================
echo.

echo  Instalando PyInstaller...
pip install pyinstaller --quiet
echo.

echo  Gerando SGPilot.exe...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "SGPilot" ^
    --icon "sgpilot.ico" ^
    --hidden-import "customtkinter" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL._tkinter_finder" ^
    --collect-all "customtkinter" ^
    --collect-all "selenium" ^
    main.py

echo.
if exist "dist\SGPilot.exe" (
    echo  ========================================
    echo   Build concluido com sucesso!
    echo   Executavel em: dist\SGPilot.exe
    echo  ========================================
    echo.
    echo   Copie para a MESMA PASTA do .exe:
    echo     - logo.png
    echo     - logo2.png
    echo     - chrome_debug.bat
) else (
    echo  ========================================
    echo   ERRO: Build falhou!
    echo   Verifique os erros acima.
    echo  ========================================
)
echo.
pause

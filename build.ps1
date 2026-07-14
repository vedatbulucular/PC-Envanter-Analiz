param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"

# PC_Envanter_Analiz - Production Build (Windowed / Optimized)
# Sadece openpyxl kasitli olarak DAHIL edilmistir (Excel raporu icin zorunludur).
# Kodda kullanilmayan agir kutuphaneler dislandi: pandas, numpy, scipy, matplotlib, PyQt5,
# IPython, jupyter, torch, sklearn, tensorflow, transformers, cv2

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --uac-admin `
    --name "PC_Envanter_Analiz" `
    --collect-all customtkinter `
    --collect-all darkdetect `
    --hidden-import=wmi `
    --hidden-import=pythoncom `
    --hidden-import=pywintypes `
    --hidden-import=win32api `
    --hidden-import=win32com `
    --hidden-import=win32com.client `
    --hidden-import=win32security `
    --hidden-import=winreg `
    --hidden-import=psutil `
    --exclude-module pandas `
    --exclude-module numpy `
    --hidden-import=openpyxl `
    --hidden-import=openpyxl.styles `
    --hidden-import=openpyxl.utils `
    --hidden-import=openpyxl.workbook `
    "--hidden-import=openpyxl.writer.excel" `
    --hidden-import=et_xmlfile `
    --hidden-import=tkinter `
    "--hidden-import=tkinter.messagebox" `
    --hidden-import=packaging `
    "--hidden-import=packaging.version" `
    --hidden-import=PIL `
    --hidden-import=PIL.Image `
    --exclude-module=scipy `
    --exclude-module=matplotlib `
    --exclude-module=PyQt5 `
    --exclude-module=PyQt6 `
    --exclude-module=IPython `
    --exclude-module=jupyter `
    --exclude-module=torch `
    --exclude-module=sklearn `
    --exclude-module=tensorflow `
    --exclude-module=transformers `
    --exclude-module=cv2 `
    --exclude-module=notebook `
    --exclude-module=numba `
    --exclude-module=sympy `
    --exclude-module=PySide2 `
    --exclude-module=PySide6 `
    --exclude-module=wx `
    --exclude-module=gtk `
    --exclude-module=gi `
    arayuz.py

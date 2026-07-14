<#!
.SYNOPSIS
Builds a shareable Windows desktop app for a teacher.

.DESCRIPTION
Run once from PowerShell on a computer that has Python. The finished app is
dist\TeacherPdfConverter\TeacherPdfConverter.exe. Microsoft Word still needs
to be installed and activated on each computer that runs the app.
#>

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --windowed --name TeacherPdfConverter `
  --collect-all fitz `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --hidden-import win32com.client `
  teacher_pdf_converter.py

Write-Host ''
Write-Host 'Build complete:' -ForegroundColor Green
Write-Host (Join-Path $PSScriptRoot 'dist\TeacherPdfConverter\TeacherPdfConverter.exe')

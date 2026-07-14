<#
.SYNOPSIS
Tạo một tệp .exe tiếng Việt, có thể gửi trực tiếp cho giáo viên.

.DESCRIPTION
Tệp kết quả là dist\ChuyenDoiPDFSangWord.exe. Tệp này không cần cài Python
trên máy của giáo viên, nhưng mỗi máy vẫn cần có Microsoft Word bản desktop
đã kích hoạt.
#>

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (Get-Command py -ErrorAction SilentlyContinue) {
    $python = 'py'
    $pythonArgs = @('-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $python = 'python'
    $pythonArgs = @()
} else {
    throw 'Không tìm thấy Python. Hãy cài Python 3.10 hoặc mới hơn rồi chạy lại.'
}

& $python @pythonArgs -m pip install --upgrade pip
& $python @pythonArgs -m pip install -r requirements.txt pyinstaller
& $python @pythonArgs -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name ChuyenDoiPDFSangWord `
    --collect-all fitz `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32com.client `
    teacher_pdf_converter_vi.py

Write-Host ''
Write-Host 'Đã tạo ứng dụng:' -ForegroundColor Green
Write-Host (Join-Path $PSScriptRoot 'dist\ChuyenDoiPDFSangWord.exe')

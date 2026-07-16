<#
.SYNOPSIS
Cài ứng dụng tiếng Việt cho người dùng hiện tại, không cần quyền quản trị.

.PARAMETER Source
Đường dẫn đến ChuyenDoiPDFSangWord.exe. Mặc định dùng tệp trong thư mục dist.
#>

param(
    [string]$Source
)

$ErrorActionPreference = 'Stop'
if (-not $Source) {
    $candidates = @(
        (Join-Path $PSScriptRoot 'ChuyenDoiPDFSangWord.exe'),
        (Join-Path $PSScriptRoot 'release\ChuyenDoiPDFSangWord.exe'),
        (Join-Path $PSScriptRoot 'dist\ChuyenDoiPDFSangWord.exe')
    )
    $Source = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
}
if (-not $Source) {
    throw 'Không tìm thấy ChuyenDoiPDFSangWord.exe. Hãy tải bản phát hành đầy đủ thay vì Source code của GitHub.'
}
$Source = (Resolve-Path -LiteralPath $Source).Path
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
    throw "Không tìm thấy tệp ứng dụng: $Source"
}

$installDir = Join-Path $env:LOCALAPPDATA 'ChuyenDoiPDFSangWord'
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
$target = Join-Path $installDir 'ChuyenDoiPDFSangWord.exe'
Copy-Item -LiteralPath $Source -Destination $target -Force

$shortcutCreated = $false
$desktopShortcutCreated = $false
try {
    $shell = New-Object -ComObject WScript.Shell
    $startMenu = [Environment]::GetFolderPath([Environment+SpecialFolder]::Programs)
    if (-not $startMenu) {
        throw 'Windows did not return a Start Menu Programs folder.'
    }
    New-Item -ItemType Directory -Path $startMenu -Force | Out-Null
    # Use an ASCII filename so the shortcut is created reliably on all Windows locale/code-page settings.
    $shortcut = $shell.CreateShortcut((Join-Path $startMenu 'Chuyen PDF sang Word - thay Bao.lnk'))
    $shortcut.TargetPath = $target
    $shortcut.WorkingDirectory = $installDir
    $shortcut.Description = 'Chuyển PDF sang Word cho thầy Bảo'
    $shortcut.Save()
    $shortcutCreated = $true

    $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
    if (-not $desktop) {
        throw 'Windows did not return a Desktop folder.'
    }
    New-Item -ItemType Directory -Path $desktop -Force | Out-Null
    $desktopShortcut = $shell.CreateShortcut((Join-Path $desktop 'Chuyen PDF sang Word - thay Bao.lnk'))
    $desktopShortcut.TargetPath = $target
    $desktopShortcut.WorkingDirectory = $installDir
    $desktopShortcut.Description = 'Chuyển PDF sang Word cho thầy Bảo'
    $desktopShortcut.Save()
    $desktopShortcutCreated = $true
} catch {
    Write-Warning "Ứng dụng đã được cài, nhưng không tạo được một hoặc nhiều lối tắt: $($_.Exception.Message)"
}

Write-Host 'Đã cài xong.' -ForegroundColor Green
if ($shortcutCreated) {
    Write-Host 'Mở Start Menu và tìm: Chuyen PDF sang Word - thay Bao'
} else {
    Write-Host "Mở ứng dụng trực tiếp tại: $target"
}
if ($desktopShortcutCreated) {
    Write-Host 'Đã thêm lối tắt trên Desktop.'
}

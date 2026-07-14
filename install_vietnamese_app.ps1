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
        (Join-Path $PSScriptRoot 'dist\ChuyenDoiPDFSangWord.exe')
    )
    $Source = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
}
if (-not $Source) {
    throw 'Cannot find ChuyenDoiPDFSangWord.exe in the application folder.'
}
$Source = (Resolve-Path -LiteralPath $Source).Path
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
    throw "Không tìm thấy tệp ứng dụng: $Source"
}

$installDir = Join-Path $env:LOCALAPPDATA 'ChuyenDoiPDFSangWord'
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
$target = Join-Path $installDir 'ChuyenDoiPDFSangWord.exe'
Copy-Item -LiteralPath $Source -Destination $target -Force

$shell = New-Object -ComObject WScript.Shell
$startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
$shortcut = $shell.CreateShortcut((Join-Path $startMenu 'Chuyển PDF sang Word.lnk'))
$shortcut.TargetPath = $target
$shortcut.WorkingDirectory = $installDir
$shortcut.Description = 'Chuyển PDF sang Word có thể chỉnh sửa'
$shortcut.Save()

Write-Host 'Đã cài xong.' -ForegroundColor Green
Write-Host 'Mở Start Menu và tìm: Chuyển PDF sang Word'

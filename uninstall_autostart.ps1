$startup = [System.Environment]::GetFolderPath('Startup')
$lnkPath = Join-Path $startup 'VbeComplete.lnk'
if (Test-Path $lnkPath) {
    Remove-Item $lnkPath -Force
    Write-Host ("Removed startup shortcut: " + $lnkPath)
} else {
    Write-Host "No VbeComplete startup shortcut found."
}

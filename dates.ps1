param([string]$folder)

if (-not $folder) {
    $folder = "test"
    Write-Host "Using default folder '$folder'." -ForegroundColor Red
}

if (-not (Test-Path $folder -PathType Container)) {
    Write-Host "Folder '$folder' does not exist." -ForegroundColor Red
    exit 1
}

dotnet script dates.csx $folder

Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

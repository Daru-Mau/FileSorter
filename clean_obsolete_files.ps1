# Clean up obsolete files and build artifacts from the FileSorter project
# This script removes build artifacts, cache directories, and obsolete files

# Stop on errors
$ErrorActionPreference = "Stop"

Write-Host "Starting cleanup of obsolete files..." -ForegroundColor Green

# 1. Remove the original monolithic file if it exists
$monoFile = Join-Path -Path $PSScriptRoot -ChildPath "src\excel_file_renamer.py"
if (Test-Path $monoFile) {
    Write-Host "Removing obsolete monolithic file: $monoFile" -ForegroundColor Yellow
    Remove-Item -Path $monoFile -Force
}

# 2. Remove build directories
$buildDirs = @(
    "build",
    "dist",
    "FileSorter_clean_build"
)

foreach ($dir in $buildDirs) {
    $path = Join-Path -Path $PSScriptRoot -ChildPath $dir
    if (Test-Path $path) {
        Write-Host "Removing build directory: $dir" -ForegroundColor Yellow
        Remove-Item -Path $path -Recurse -Force
    }
}

# 3. Remove __pycache__ directories
Write-Host "Removing Python cache directories..." -ForegroundColor Yellow
Get-ChildItem -Path $PSScriptRoot -Directory -Include "__pycache__" -Recurse | 
    ForEach-Object {
        Write-Host "  Removing: $($_.FullName)" -ForegroundColor Gray
        Remove-Item -Path $_.FullName -Recurse -Force
    }

# 4. Remove .pyc files
Write-Host "Removing Python compiled files..." -ForegroundColor Yellow
Get-ChildItem -Path $PSScriptRoot -Include "*.pyc","*.pyo","*.pyd" -Recurse |
    ForEach-Object {
        Write-Host "  Removing: $($_.FullName)" -ForegroundColor Gray
        Remove-Item -Path $_.FullName -Force
    }

# 5. Remove other specific obsolete files
$obsoleteFiles = @(
    # Add other specific obsolete files here
)

foreach ($file in $obsoleteFiles) {
    $path = Join-Path -Path $PSScriptRoot -ChildPath $file
    if (Test-Path $path) {
        Write-Host "Removing obsolete file: $file" -ForegroundColor Yellow
        Remove-Item -Path $path -Force
    }
}

Write-Host "Cleanup completed successfully!" -ForegroundColor Green

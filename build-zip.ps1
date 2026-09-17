# ============================================================
# build-zip.ps1 — сборка access-course-project.zip (Windows)
# Использование:
#   cd C:\path\to\access-course
#   powershell -ExecutionPolicy Bypass -File .\build-zip.ps1
# ============================================================

$ErrorActionPreference = 'Stop'

# --- 1. Перейти в каталог проекта (каталог, где лежит этот скрипт) ---
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $projectDir
Write-Host "Каталог проекта: $projectDir" -ForegroundColor Cyan

# --- 2. Имя архива и удаление старого ---
$zipName = 'access-course-project.zip'
$zipPath = Join-Path -Path $projectDir -ChildPath $zipName
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
    Write-Host "Старый архив удалён: $zipName"
}

# --- 3. Исключения: git, временные, системные, кэш ---
$excludeDirs  = @('.git', 'node_modules', '.vscode', '.idea', '.cache', 'dist', 'build', 'import')
$excludeFiles = @('Thumbs.db', '.DS_Store', 'Desktop.ini')
$excludeExt   = @('*.tmp', '*.temp', '*.log', '*.swp', '*.bak', '*.zip', '*.accdb', '*.mdb', '*.laccdb', '*.csv', '*.xlsx')

function Test-Excluded([System.IO.FileSystemInfo]$item) {
    # CSV репозитория — часть учебника и входят в архив: datasets/calendar.csv
    # и starter-data/*.csv (стартовый комплект). Пользовательские *.csv
    # (выгрузки Olist, import/) исключаются ниже.
    if ($item.Name -eq 'calendar.csv') { return $false }
    if ($item.FullName -like '*\starter-data\*.csv') { return $false }
    foreach ($d in $excludeDirs)  { if ($item.FullName -like "*\$d\*")     { return $true } }
    foreach ($f in $excludeFiles) { if ($item.Name -eq $f)                { return $true } }
    foreach ($e in $excludeExt)   { if ($item.Name -like $e)              { return $true } }
    return $false
}

# --- 4. Собрать список нужных файлов ---
$allFiles = Get-ChildItem -Path $projectDir -Recurse -File |
    Where-Object { -not (Test-Excluded $_) }

if ($allFiles.Count -eq 0) {
    Write-Error 'Файлы для архивирования не найдены — проверьте содержимое каталога.'
}

# --- 5. Создать ZIP (через .NET — без внешних зависимостей) ---
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in $allFiles) {
        $relative = $file.FullName.Substring($projectDir.Length + 1) -replace '\\', '/'
        $entryName = 'access-course/' + $relative
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
}
finally {
    $zip.Dispose()
}

# --- 6. Проверить наличие архива ---
if (-not (Test-Path $zipPath)) {
    Write-Error 'Архив не создан — сборка не удалась.'
}

# --- 7. Вывести список файлов в архиве ---
Write-Host ''
Write-Host '=== Содержимое архива ===' -ForegroundColor Cyan
$zipRead = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    $zipRead.Entries |
        Sort-Object FullName |
        ForEach-Object { Write-Host ("{0,10}  {1}" -f $_.Length, $_.FullName) }
    $count = $zipRead.Entries.Count
}
finally {
    $zipRead.Dispose()
}

# --- 8. Размер архива ---
$sizeBytes = (Get-Item $zipPath).Length
$sizeKB = [math]::Round($sizeBytes / 1KB, 1)
$sizeMB = [math]::Round($sizeBytes / 1MB, 2)
Write-Host ''
Write-Host "Файлов в архиве : $count"
if ($sizeMB -ge 1) { Write-Host ("Размер архива   : {0} МБ ({1} байт)" -f $sizeMB, $sizeBytes) }
else               { Write-Host ("Размер архива   : {0} КБ ({1} байт)" -f $sizeKB, $sizeBytes) }

# --- 9. Успешное завершение ---
Write-Host ''
Write-Host "ГОТОВО: $zipName создан успешно." -ForegroundColor Green
Write-Host 'Проверка: Expand-Archive access-course-project.zip -DestinationPath ..\check -Force'

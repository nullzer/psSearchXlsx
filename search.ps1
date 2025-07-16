# === НАСТРОЙКИ ===
$folderPath = "C:\Путь\К\Папке"               # Папка с Excel-файлами
$inputFile = "C:\Путь\К\input.txt"            # Файл со строками для поиска
$outputCsv = "C:\Путь\К\results.csv"          # Файл для вывода CSV

# === ЗАПУСК ТАЙМЕРА ===
$startTime = Get-Date

# === ЧТЕНИЕ СТРОК ДЛЯ ПОИСКА ===
if (!(Test-Path $inputFile)) {
    Write-Error "Файл со строками поиска не найден: $inputFile"
    exit
}

$searchValues = @()
foreach ($line in Get-Content -Path $inputFile) {
    $cleaned = $line.Trim()
    if ($cleaned -ne "") {
        $normalized = ($cleaned -replace "ё", "е" -replace "Ё", "Е").ToLower()
        $searchValues += $normalized
    }
}

if ($searchValues.Count -eq 0) {
    Write-Error "Файл '$inputFile' не содержит подходящих строк для поиска."
    exit
}

# === ПОДГОТОВКА ===
$excelFiles = Get-ChildItem -Path $folderPath -Re

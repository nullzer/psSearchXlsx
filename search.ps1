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
$excelFiles = Get-ChildItem -Path $folderPath -Recurse -Include *.xlsx, *.xls -File
$totalFiles = $excelFiles.Count
$fileIndex = 0
$results = @()

# === ИНИЦИАЛИЗАЦИЯ EXCEL ===
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($file in $excelFiles) {
    $fileIndex++
    $progressPercent = [int](($fileIndex / $totalFiles) * 100)

    Write-Progress -Activity "Обработка Excel-файлов..." `
                   -Status "Файл $fileIndex из $totalFiles: $($file.Name)" `
                   -PercentComplete $progressPercent

    try {
        $workbook = $excel.Workbooks.Open($file.FullName, $null, $true)  # read-only

        foreach ($sheet in $workbook.Sheets) {
            $usedRange = $sheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count

            for ($row = 1; $row -le $rowCount; $row++) {
                $rowText = ""
                for ($col = 1; $col -le $colCount; $col++) {
                    $cell = $sheet.Cells.Item($row, $col)
                    if ($cell -ne $null) {
                        $rowText += ($cell.Text + "`t")
                    }
                }

                $rowNormalized = ($rowText -replace "ё", "е" -replace "Ё", "Е").ToLower()

                foreach ($searchValue in $searchValues) {
                    Write-Host "Поиск по строке: '$searchValue'" -ForegroundColor Cyan

                    if ($rowNormalized -like "*$searchValue*") {
                        $results += [pscustomobject]@{
                            Файл              = $file.FullName
                            ДатаСоздания      = $file.CreationTime
                            ДатаИзменения     = $file.LastWriteTime
                            Лист              = $sheet.Name
                            НомерСтроки       = $row
                            Найдено           = $searchValue
                            СодержимоеСтроки  = $rowText.Trim()
                        }
                        break
                    }
                }
            }
        }

        $workbook.Close($false)
    } catch {
        Write-Warning "Ошибка при обработке файла $($file.FullName): $_"
    }
}

# === ЗАВЕРШЕНИЕ РАБОТЫ EXCEL ===
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# === СОХРАНЕНИЕ РЕЗУЛЬТАТОВ В CSV ===
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

# === ВЫВОД СТАТИСТИКИ ===
$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host "`nПоиск завершён. Найдено: $($results.Count) совпадений."
Write-Host "Результаты сохранены в файл: $outputCsv"
Write-Host "Общее время выполнения: $($duration.ToString())"

# === НАСТРОЙКИ ===
$folderPath = "C:\Путь\К\Папке"               # Путь к папке с Excel-файлами
$inputFile = "C:\Путь\К\input.txt"            # Файл со строками для поиска
$outputCsv = "C:\Путь\К\results.csv"          # Файл для CSV-результатов

# === ПОДГОТОВКА ===
if (!(Test-Path $inputFile)) {
    Write-Error "Файл со строками поиска не найден: $inputFile"
    exit
}

$searchValuesRaw = Get-Content -Path $inputFile | Where-Object { $_.Trim() -ne "" }
$searchValues = $searchValuesRaw | ForEach-Object {
    ($_ -replace "ё", "е" -replace "Ё", "Е").ToLower()
}

if ($searchValues.Count -eq 0) {
    Write-Error "Файл '$inputFile' не содержит строк для поиска."
    exit
}

# Создаем таблицу для хранения результатов
$results = @()

# Получаем список Excel-файлов
$excelFiles = Get-ChildItem -Path $folderPath -Recurse -Include *.xlsx, *.xls -File

# Запускаем Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($file in $excelFiles) {
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

# Завершаем работу Excel
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Сохраняем в CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

Write-Host "`nПоиск завершён. Результаты сохранены в файл: $outputCsv"

# === НАСТРОЙКИ ===
$folderPath = "C:\Путь\К\Папке"               # Путь к папке с Excel-файлами
$inputFile = "C:\Путь\К\input.txt"            # Путь к файлу со строками для поиска
$outputFile = "C:\Путь\К\results.txt"         # Путь к файлу для вывода результатов

# === ПОДГОТОВКА ===
if (!(Test-Path $inputFile)) {
    Write-Error "Файл со строками поиска не найден: $inputFile"
    exit
}

$searchValues = Get-Content -Path $inputFile | Where-Object { $_.Trim() -ne "" }

if ($searchValues.Count -eq 0) {
    Write-Error "Файл '$inputFile' не содержит строк для поиска."
    exit
}

"Результаты поиска (дата: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))" | Out-File -FilePath $outputFile -Encoding UTF8
"" | Out-File -FilePath $outputFile -Append

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

                foreach ($searchValue in $searchValues) {
                    if ($rowText -like "*$searchValue*") {
                        $result = @"
Файл: $($file.FullName)
Дата создания: $($file.CreationTime)
Дата изменения: $($file.LastWriteTime)
Лист: $($sheet.Name), Строка: $row
Найдено: '$searchValue'
Содержимое строки: $rowText
--------------------------------------------------------------------------------
"@
                        $result | Out-File -FilePath $outputFile -Append -Encoding UTF8
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

Write-Host "`nПоиск завершён. Результаты сохранены в файл: $outputFile"

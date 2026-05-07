$sumatra = "C:\Users\11119223\AppData\Local\SumatraPDF\SumatraPDF.exe"
$printer = "\\ha-19-print\Follow_Me_BW"

Get-ChildItem -Filter *.pdf | ForEach-Object {
    Write-Host "Printing: $($_.Name)"
    & $sumatra -print-to $printer -print-settings "monochrome" -silent $_.FullName
    Start-Sleep -Seconds 3
}

Write-Host "Done."
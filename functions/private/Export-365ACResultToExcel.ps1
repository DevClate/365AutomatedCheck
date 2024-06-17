function Export-365ACResultToExcel {
    param (
        [array]$Results,
        [string]$ExcelFilePath,
        [int]$TotalTests,
        [int]$PassedTests,
        [int]$FailedTests,
        [string]$TestedProperty
    )

    $results | Export-Excel -Path $ExcelFilePath -WorkSheetname 'Results' -AutoSize -FreezePane 7, 1 -NoHeader -StartRow 7 -ConditionalText (New-ConditionalText -Text 'Yes' -BackgroundColor Green -ForegroundColor White), (New-ConditionalText -Text 'No' -BackgroundColor Red -ForegroundColor White)

    $excelPackage = Open-ExcelPackage -Path $ExcelFilePath
    $resultSheet = $excelPackage.Workbook.Worksheets['Results']

    # Adding title to the Results Sheet
    $resultSheet.InsertRow(1, 0)
    $resultSheet.Cells["A1"].Value = "365AutomatedCheck Results"
    $resultSheet.Cells["A1:B1"].Merge = $true
    $resultSheet.Cells["A1:B1"].Style.Font.Size = 20
    $resultSheet.Cells["A1:B1"].Style.Font.Bold = $true
    $resultSheet.Cells["A1:B1"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $resultSheet.Cells["A1:B1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Black)
    $resultSheet.Cells["A1:B1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

    # Centering text in column B
    $resultSheet.Column(2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    
    # Adding summary information to the Results Sheet
    $resultSheet.Cells["A2"].Value = "Total tests"
    $resultSheet.Cells["B2"].Value = $TotalTests
    $resultSheet.Cells["A3"].Value = "Passed"
    $resultSheet.Cells["B3"].Value = $PassedTests
    $resultSheet.Cells["A4"].Value = "Failed"
    $resultSheet.Cells["B4"].Value = $FailedTests
    $resultSheet.Cells["A5"].Value = "Not tested"
    $resultSheet.Cells["B5"].Value = 0

    $resultSheet.Cells["A2:A5"].Style.Font.Size = 16
    $resultSheet.Cells["A2:A5"].Style.Font.Bold = $true
    $resultSheet.Cells["A2:A5"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["A2:A5"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gray)
    $resultSheet.Cells["A2:A5"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

    $resultSheet.Cells["B2:B5"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
    $resultSheet.Cells["B3"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Green)
    $resultSheet.Cells["B4"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
    $resultSheet.Cells["B5"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gray)
    $resultSheet.Cells["B2:B5"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $resultSheet.Cells["B2:B5"].Style.Font.Size = 16
    $resultSheet.Cells["B2:B5"].Style.Font.Bold = $true

    # Formatting the Results Sheet headers
    $resultSheet.Cells["A6:B6"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["A6:B6"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Black)
    $resultSheet.Cells["A6"].Value = "User Display Name"
    $resultSheet.Cells["B6"].Value = $TestedProperty
    $resultSheet.Cells["A6:B6"].Style.Font.Bold = $true
    $resultSheet.Cells["A6:B6"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

    # Applying conditional formatting to the results
    $startRow = 7
    for ($i = 0; $i -lt $Results.Count; $i++) {
        $row = $i + $startRow
        if ($Results[$i].$TestedProperty -eq 'TRUE') {
            $resultSheet.Cells["A$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["A$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Green)
            $resultSheet.Cells["A$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $resultSheet.Cells["B$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["B$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Green)
            $resultSheet.Cells["B$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        }
        else {
            $resultSheet.Cells["A$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["A$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
            $resultSheet.Cells["A$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $resultSheet.Cells["B$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["B$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
            $resultSheet.Cells["B$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        }
    }
    
    Close-ExcelPackage $excelPackage
}

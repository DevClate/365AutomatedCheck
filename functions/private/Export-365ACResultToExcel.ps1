function Export-365ACResultToExcel {
    param (
        [array]$Results,
        [string]$ExcelFilePath,
        [int]$TotalTests,
        [int]$PassedTests,
        [int]$FailedTests,
        [string]$TestedProperty
    )

    # Exporting the results to an Excel file
    $results | Export-Excel -Path $ExcelFilePath -WorkSheetname 'Results' -AutoSize -FreezePane 8,1 -NoHeader -StartRow 8

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

    # Adding black line divider
    $resultSheet.Cells["A6:B6"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["A6:B6"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Black)

    # Formatting the Results Sheet headers
    $resultSheet.Cells["A7"].Value = "User Display Name"
    $resultSheet.Cells["B7"].Value = "$TestedProperty"
    $resultSheet.Cells["A7:B7"].Style.Font.Bold = $true
    $resultSheet.Cells["A7:B7"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $resultSheet.Cells["A7:B7"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gray)
    $resultSheet.Cells["A7:B7"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

    # Applying conditional formatting to the results
    $startRow = 8
    for ($i = 0; $i -lt $Results.Count; $i++) {
        $row = $i + $startRow
        $hasMobilePhone = [System.Convert]::ToBoolean($Results[$i]."$TestedProperty")
        
        $resultSheet.Cells["A$row"].Value = $Results[$i].'User Display Name'
        $resultSheet.Cells["B$row"].Value = $Results[$i]."$TestedProperty"

        if ($hasMobilePhone) {
            $resultSheet.Cells["A$row:B$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["A$row:B$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Green)
            $resultSheet.Cells["A$row:B$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        } else {
            $resultSheet.Cells["A$row:B$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $resultSheet.Cells["A$row:B$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
            $resultSheet.Cells["A$row:B$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        }
    }

    Close-ExcelPackage $excelPackage
}

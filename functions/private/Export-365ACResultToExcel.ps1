<#
.SYNOPSIS
Exports the 365AutomatedCheck results to an Excel file.

.DESCRIPTION
The Export-365ACResultToExcel function takes an array of results, along with other parameters, and exports the results to an Excel file. It applies formatting, adds summary information, and applies conditional formatting based on the tested property.

.PARAMETER Results
The array of results to be exported to the Excel file.

.PARAMETER OutputExcelFilePath
The path of the output Excel file.

.PARAMETER TotalTests
The total number of tests.

.PARAMETER PassedTests
The number of tests that passed.

.PARAMETER FailedTests
The number of tests that failed.

.PARAMETER TestedProperty
The name of the property that was tested.

.EXAMPLE
$Results = @(
    [PSCustomObject]@{
        "UserDisplayName" = "John Doe"
        "TestedProperty" = $true
    },
    [PSCustomObject]@{
        "UserDisplayName" = "Jane Smith"
        "TestedProperty" = $false
    }
)

$OutputExcelFilePath = "C:\Results.xlsx"
$TotalTests = $Results.Count
$PassedTests = ($Results | Where-Object { $_.TestedProperty }).Count
$FailedTests = ($Results | Where-Object { -not $_.TestedProperty }).Count
$TestedProperty = "TestedProperty"

Export-365ACResultToExcel -Results $Results -OutputExcelFilePath $OutputExcelFilePath -TotalTests $TotalTests -PassedTests $PassedTests -FailedTests $FailedTests -TestedProperty $TestedProperty

This example exports the results to an Excel file named "Results.xlsx" and applies formatting and conditional formatting based on the "TestedProperty" property.

.NOTES
Author: Clayton Tyger
Date: 06/20/2024
#>
function Export-365ACResultToExcel {
    param (
        [array]$Results,
        [string]$OutputExcelFilePath,
        [int]$TotalTests,
        [int]$PassedTests,
        [int]$FailedTests,
        [string]$TestedProperty
    )

    $results | Export-Excel -Path $OutputExcelFilePath -WorkSheetname 'Results' -AutoSize -FreezePane 7, 1 -NoHeader -StartRow 7 -ConditionalText (New-ConditionalText -Text 'Yes' -BackgroundColor Green -ForegroundColor White), (New-ConditionalText -Text 'No' -BackgroundColor Red -ForegroundColor White)

    $excelPackage = Open-ExcelPackage -Path $OutputExcelFilePath
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

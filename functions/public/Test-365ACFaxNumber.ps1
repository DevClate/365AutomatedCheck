<#
.SYNOPSIS
    This function tests whether users have a fax number and exports the results to Excel, HTML, or the console.

.DESCRIPTION
    The Test-365ACFaxNumber function tests whether users have a fax number by checking the FaxNumber property of each user. It accepts an array of users as input and outputs the results as a custom object with the user's display name and a boolean value indicating whether they have a fax number.

.PARAMETER Users
    Specifies the array of users to test. If not provided, it retrieves all users using the Get-MgUser cmdlet.

.PARAMETER ExcelFilePath
    Specifies the path to export the results as an Excel file. If this parameter is provided, the Export-365ACResultToExcel function is called to export the results.

.PARAMETER HtmlFilePath
    Specifies the path to export the results as an HTML file. If this parameter is provided, the Export-365ACResultToHtml function is called to export the results.

.EXAMPLE
    Test-365ACFaxNumber -Users $users -ExcelFilePath "C:\Results.xlsx"
    Tests the specified users for fax numbers and exports the results to an Excel file.

.EXAMPLE
    Test-365ACFaxNumber -HtmlFilePath "C:\Results.html"
    Tests all users for fax numbers and exports the results to an HTML file.

.NOTES
    - This function requires the ImportExcel module to export results to Excel. If the module is not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions are assumed to be defined elsewhere in the script.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>

Function Test-365ACFaxNumber {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, FaxNumber | Select-Object displayname, FaxNumber),
        [ValidatePattern('\.xlsx$')]
        [string]$ExcelFilePath,
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath
    )
    BEGIN {
        if ($ExcelFilePath -and !(Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
            Write-Error "Export-Excel cmdlet not found. Please install the ImportExcel module."
            return
        }
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $hasFaxNumber = [bool]($user.FaxNumber)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Fax Number'    = $hasFaxNumber
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Fax Number' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Fax Number'
        }
        else {
            Write-Output $results
        }
    }
}
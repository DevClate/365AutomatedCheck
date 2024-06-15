<#
.SYNOPSIS
    This function tests if users have a job title assigned and exports the results to Excel, HTML, or the console.

.DESCRIPTION
    The Test-365ACJobTitle function tests if users have a job title assigned. It takes an array of users as input and checks if each user has a job title. The function returns a collection of results indicating whether each user has a job title or not.

.PARAMETER Users
    Specifies the array of users to test. If not provided, all users will be tested.

.PARAMETER ExcelFilePath
    Specifies the path to export the results to an Excel file. If this parameter is specified, the function will use the Export-365ACResultToExcel function to export the results.

.PARAMETER HtmlFilePath
    Specifies the path to export the results to an HTML file. If this parameter is specified, the function will use the Export-365ACResultToHtml function to export the results.

.EXAMPLE
    Test-365ACJobTitle -Users (Get-MgUser -All) -ExcelFilePath "C:\Results.xlsx"
    Tests all users and exports the results to an Excel file located at "C:\Results.xlsx".

.EXAMPLE
    Test-365ACJobTitle -Users $users -HtmlFilePath "C:\Results.html"
    Tests the specified users and exports the results to an HTML file located at "C:\Results.html".

.NOTES
    - This function requires the ImportExcel module to export results to Excel. If the module is not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions are assumed to be defined elsewhere in the script.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
function Test-365ACJobTitle {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [array]$Users = (Get-MgUser -All),
        
        [ValidatePattern('\.xlsx$')]
        [string]$ExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,

        [string]$TestedProperty = 'Has Job Title'
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
            $hasJobTitle = [bool]($user.JobTitle)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Job Title' = $hasJobTitle
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Job Title' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
        } elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
        } else {
            Write-Output $results
        }
    }
}
<#
.SYNOPSIS
    Tests if users have a job title assigned and exports the results.

.DESCRIPTION
    The Test-365ACJobTitle function checks if users in Microsoft 365 have a job title assigned. It can validate job titles against a list of valid titles if provided. The results can be exported to Excel, HTML, or output to the console.

.PARAMETER Users
    Specifies an array of users to test. If not provided, the function tests all users in Microsoft 365.

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid job titles. If provided, the function validates each user's job title against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to export the results to an Excel file. If specified, the function exports the results using the Export-365ACResultToExcel function.

.PARAMETER HtmlFilePath
    Specifies the path to export the results to an HTML file. If specified, the function exports the results using the Export-365ACResultToHtml function.

.PARAMETER TestedProperty
    Specifies the property being tested. Defaults to 'Has Job Title'.

.EXAMPLE
    Test-365ACJobTitle -Users (Get-MgUser -All) -OutputExcelFilePath "C:\Results.xlsx"
    Tests all users for a job title and exports the results to an Excel file.

.EXAMPLE
    Test-365ACJobTitle -Users $users -HtmlFilePath "C:\Results.html"
    Tests the specified users for a job title and exports the results to an HTML file.

.NOTES
    - Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
function Test-365ACJobTitle {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, JobTitle),
        
        [ValidatePattern('\.xlsx$')]
        [string]$ValidationExcelFilePath,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Has Job Title'
    )
    BEGIN {
        $validJobTitles = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
                return
            }
            # Import the Excel file to get valid job titles
            $validJobTitles = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty Title -Unique
        }
        $results = @()
    }
    PROCESS {
        $columnName = $ValidationExcelFilePath ? 'Has Valid Job Title' : 'Has Job Title'
        foreach ($user in $Users) {
            $hasJobTitle = [bool]($user.JobTitle)
            if ($ValidationExcelFilePath) {
                $hasJobTitle = $user.JobTitle -in $validJobTitles
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName         = $hasJobTitle
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.$columnName }).Count
        $failedTests = $totalTests - $passedTests
        if ($OutputExcelFilePath) {
            Export-365ACResultToExcel -Results $results -OutputExcelFilePath $OutputExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        } elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        } else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
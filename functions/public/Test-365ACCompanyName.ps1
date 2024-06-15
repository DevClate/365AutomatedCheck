<#
.SYNOPSIS
    This function tests whether users have a company name associated with their account and exports the results to Excel, HTML, or the console.
.DESCRIPTION
    The Test-365ACCompanyName function checks whether users have a company name associated with their account. It takes an array of users, an optional Excel file path as input. It then iterates through each user, determines if they have a company name, and generates a result object for each user. The function returns the results as an array of objects or exports them to an Excel file or an HTML file.
.PARAMETER Users
    Specifies an array of users to test. Each user should have a display name and a company name property.
.PARAMETER ExcelFilePath
    Specifies the file path to export the results to an Excel file. If this parameter is provided, the Export-Excel module must be installed.
.PARAMETER HtmlFilePath
    Specifies the file path to export the results to an HTML file.
.EXAMPLE
    Test-365ACCompanyName -Users $users -ExcelFilePath "C:\Results.xlsx"
    This example tests the company names of the users in the $users array and exports the results to an Excel file located at "C:\Results.xlsx".
.EXAMPLE
    Test-365ACCompanyName -Users $users -HtmlFilePath "C:\Results.html"
    This example tests the company names of the users in the $users array and exports the results to an HTML file located at "C:\Results.html".
#>
Function Test-365ACCompanyName {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, CompanyName | Select-Object displayname, CompanyName),
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
            $hasCompanyName = [bool]($user.CompanyName)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Company Name'  = $hasCompanyName
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Company Name' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Company Name'
        }
        else {
            Write-Output $results
        }
    }
}
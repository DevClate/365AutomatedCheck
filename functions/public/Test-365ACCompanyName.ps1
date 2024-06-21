<#
.SYNOPSIS
    Tests whether users have a company name and optionally validates it against a list of valid company names.

.DESCRIPTION
    The Test-365ACCompanyName function tests whether users have a company name and optionally validates it against a list of valid company names. It generates a report of the test results, which can be exported to an Excel file or an HTML file.

.PARAMETER Users
    Specifies the users to be tested. If not specified, all users in the organization will be tested.

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid company names. If specified, the function will validate the company names of the users against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to save the test results as an Excel file. If not specified, the test results will be displayed in the console.

.PARAMETER HtmlFilePath
    Specifies the path to save the test results as an HTML file. If not specified, the test results will be displayed in the console.

.PARAMETER TestedProperty
    Specifies the name of the tested property. This will be used as the column name in the test results.

.EXAMPLE
    Test-365ACCompanyName -Users (Get-MgUser -All) -ValidationExcelFilePath "C:\Validation.xlsx" -OutputExcelFilePath "C:\Results.xlsx" -TestedProperty "Has Valid Company Name"
    Tests all users in the organization, validates their company names against the list of valid company names in the "Validation.xlsx" file, and saves the test results to the "Results.xlsx" file.

.EXAMPLE
    Test-365ACCompanyName -Users (Get-MgUser -All) -HtmlFilePath "C:\Results.html" -TestedProperty "Has Company Name"
    Tests all users in the organization and saves the test results as an HTML file named "Results.html".

#>
Function Test-365ACCompanyName {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, CompanyName | Select-Object DisplayName, CompanyName),
        
        [ValidatePattern('\.xlsx$')]
        [string]$ValidationExcelFilePath,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,

        [string]$TestedProperty = 'Has Company Name'
    )
    BEGIN {
        
        $validCompanyNames = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Write-Error "Import-Excel cmdlet not found. Please install the ImportExcel module."
                return@
            }
            # Import the Excel file to get valid company names
            $validCompanyNames = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty CompanyName -Unique
        }
        $results = @()
    }
    PROCESS {
        $columnName = $ValidationExcelFilePath ? 'Has Valid Company Name' : 'Has Company Name'
        foreach ($user in $Users) {
            $hasCompanyName = [bool]($user.CompanyName)
            if ($ValidationExcelFilePath) {
                $hasCompanyName = $user.CompanyName -in $validCompanyNames
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName         = $hasCompanyName
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
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        }
        else {
            Write-Output $results
        }
    }
}
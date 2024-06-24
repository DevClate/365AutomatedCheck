<#
.SYNOPSIS
    Tests whether users in Microsoft 365 have a company name assigned and optionally validates against a list of valid company names.

.DESCRIPTION
    The Test-365ACCompanyName function checks if users in Microsoft 365 have a company name assigned to their profile. It can also validate these company names against a provided list of valid company names. The results of this test can be exported to an Excel file, an HTML file, or displayed in the console.

.PARAMETER Users
    Specifies the users to be tested. If not specified, the function will test all users in the organization by retrieving their DisplayName and CompanyName properties.

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid company names. If specified, the function will validate the company names of the users against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to save the test results as an Excel file. If not specified, the test results will be displayed in the console. The path must end with '.xlsx'.

.PARAMETER HtmlFilePath
    Specifies the path to save the test results as an HTML file. If not specified, the test results will be displayed in the console. The path must end with '.html'.

.PARAMETER TestedProperty
    Specifies the name of the tested property. This will be used as the column name in the test results. Defaults to 'Has Company Name' but changes to 'Has Valid Company Name' if a validation list is provided.

.EXAMPLE
    Test-365ACCompanyName -Users (Get-MgUser -All) -ValidationExcelFilePath "C:\Validation.xlsx" -OutputExcelFilePath "C:\Results.xlsx"
    Tests all users in the organization, validates their company names against the list of valid company names in "Validation.xlsx", and saves the test results to "Results.xlsx".

.EXAMPLE
    Test-365ACCompanyName -Users (Get-MgUser -All) -HtmlFilePath "C:\Results.html"
    Tests all users in the organization and saves the test results as an HTML file named "Results.html".

.NOTES
    Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
    The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results to Excel or HTML, respectively.

.LINK
    https://github.com/DevClate/365AutomatedCheck
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
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
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
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
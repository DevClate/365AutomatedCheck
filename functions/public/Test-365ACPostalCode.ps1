<#
.SYNOPSIS
Tests if users have a postal code assigned in their Microsoft 365 profiles.

.DESCRIPTION
The Test-365ACPostalCode function checks each user retrieved from Microsoft 365 for a postal code value. It processes a list of users, determining whether each user has a postal code assigned. The results can be exported to an Excel file, an HTML file, or output directly.

.PARAMETER Users
An array of users to be tested. This can be piped in or specified directly. Each user should have DisplayName and PostalCode properties.

.PARAMETER OutputExcelFilePath
The file path where the Excel report will be saved. The file must have an .xlsx extension.

.PARAMETER HtmlFilePath
The file path where the HTML report will be saved. The file must have an .html extension.

.PARAMETER TestedProperty
The property being tested. This is set to 'Has Postal Code' by default.

.EXAMPLE
PS> $users = Get-MgUser -All -Property DisplayName, PostalCode | Select-Object DisplayName, PostalCode
PS> Test-365ACPostalCode -Users $users -OutputExcelFilePath "C:\Reports\PostalCodeReport.xlsx"

This example retrieves all users with their DisplayName and PostalCode, then tests each user to see if they have a postal code assigned. The results are exported to an Excel file.

.EXAMPLE
PS> Get-MgUser -All -Property DisplayName, PostalCode | Select-Object DisplayName, PostalCode | Test-365ACPostalCode -HtmlFilePath "C:\Reports\PostalCodeReport.html"

This example pipes a list of users directly into Test-365ACPostalCode, which then checks if each user has a postal code assigned. The results are exported to an HTML file.

.NOTES
This function requires the Microsoft Graph PowerShell SDK to retrieve user information from Microsoft 365.

.LINK
https://github.com/DevClate/365AutomatedCheck

#>
Function Test-365ACPostalCode {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, PostalCode | Select-Object DisplayName, PostalCode),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Has Postal Code'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $hasPostalCode = [bool]($user.PostalCode)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasPostalCode
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.$TestedProperty }).Count
        $failedTests = $totalTests - $passedTests
        if ($OutputExcelFilePath) {
            Export-365ACResultToExcel -Results $results -OutputExcelFilePath $OutputExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
        }
        else {
            Write-Output $results
        }
    }
}
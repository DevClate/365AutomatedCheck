<#
.SYNOPSIS
Tests if users' accounts are enabled in their Microsoft 365 profiles.

.DESCRIPTION
The Test-365ACAccountEnabled function checks each user retrieved from Microsoft 365 to determine if their account is enabled. It processes a list of users, determining the enabled status of each user's account. The results can be exported to an Excel file, an HTML file, or output directly.

.PARAMETER Users
An array of users to be tested. This can be piped in or specified directly. Each user should have DisplayName and AccountEnabled properties.

.PARAMETER OutputExcelFilePath
The file path where the Excel report will be saved. The file must have an .xlsx extension.

.PARAMETER HtmlFilePath
The file path where the HTML report will be saved. The file must have an .html extension.

.PARAMETER TestedProperty
The property being tested. This is set to 'Is Account Enabled' by default.

.EXAMPLE
PS> $users = Get-MgUser -All -Property DisplayName, AccountEnabled | Select-Object DisplayName, AccountEnabled
PS> Test-365ACAccountEnabled -Users $users -OutputExcelFilePath "C:\Reports\AccountEnabledReport.xlsx"

This example retrieves all users with their DisplayName and AccountEnabled status, then tests each user to see if their account is enabled. The results are exported to an Excel file.

.EXAMPLE
PS> Get-MgUser -All -Property DisplayName, AccountEnabled | Select-Object DisplayName, AccountEnabled | Test-365ACAccountEnabled -HtmlFilePath "C:\Reports\AccountEnabledReport.html"

This example pipes a list of users directly into Test-365ACAccountEnabled, which then checks if each user's account is enabled. The results are exported to an HTML file.

.NOTES
This function requires the Microsoft Graph PowerShell SDK to retrieve user information from Microsoft 365.

.LINK
https://docs.microsoft.com/powershell/module/microsoft.graph.users/get-mguser

#>
Function Test-365ACAccountEnabled {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, AccountEnabled | Select-Object DisplayName, AccountEnabled),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Is Account Enabled'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $isAccountEnabled = [bool]($user.AccountEnabled)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $isAccountEnabled
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
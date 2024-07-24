<#
.SYNOPSIS
Tests if users have a mail nickname assigned in their Microsoft 365 profiles.

.DESCRIPTION
The Test-365ACMailNickname function checks each user retrieved from Microsoft 365 for a mail nickname value. It processes a list of users, determining whether each user has a mail nickname assigned. The results can be exported to an Excel file, an HTML file, or output directly.

.PARAMETER Users
An array of users to be tested. This can be piped in or specified directly. Each user should have DisplayName and MailNickname properties.

.PARAMETER OutputExcelFilePath
The file path where the Excel report will be saved. The file must have an .xlsx extension.

.PARAMETER HtmlFilePath
The file path where the HTML report will be saved. The file must have an .html extension.

.PARAMETER TestedProperty
The property being tested. This is set to 'Has Mail Nickname' by default.

.EXAMPLE
PS> $users = Get-MgUser -All -Property DisplayName, MailNickname | Select-Object DisplayName, MailNickname
PS> Test-365ACMailNickname -Users $users -OutputExcelFilePath "C:\Reports\MailNicknameReport.xlsx"

This example retrieves all users with their DisplayName and MailNickname, then tests each user to see if they have a mail nickname assigned. The results are exported to an Excel file.

.EXAMPLE
PS> Get-MgUser -All -Property DisplayName, MailNickname | Select-Object DisplayName, MailNickname | Test-365ACMailNickname -HtmlFilePath "C:\Reports\MailNicknameReport.html"

This example pipes a list of users directly into Test-365ACMailNickname, which then checks if each user has a mail nickname assigned. The results are exported to an HTML file.

.NOTES
This function requires the Microsoft Graph PowerShell SDK to retrieve user information from Microsoft 365.

.LINK
https://docs.microsoft.com/powershell/module/microsoft.graph.users/get-mguser

.LINK
https://github.com/DevClate/365AutomatedCheck

#>
Function Test-365ACMailNickname {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, MailNickname | Select-Object DisplayName, MailNickname),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Has Mail Nickname'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $hasMailNickname = [bool]($user.MailNickname)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasMailNickname
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
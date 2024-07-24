<#
.SYNOPSIS
Tests if users are hidden from the Address Book in their Microsoft 365 profiles.

.DESCRIPTION
The Test-365ACAddressBook function checks each user retrieved from Microsoft 365 to determine if they are hidden from the Address Book. It processes a list of users, checking the 'HiddenFromAddressListsEnabled' property. The results can be exported to an Excel file, an HTML file, or output directly.

.PARAMETER Users
An array of users to be tested. This can be piped in or specified directly. Each user should have DisplayName and HiddenFromAddressListsEnabled properties.

.PARAMETER OutputExcelFilePath
The file path where the Excel report will be saved. The file must have an .xlsx extension.

.PARAMETER HtmlFilePath
The file path where the HTML report will be saved. The file must have an .html extension.

.PARAMETER TestedProperty
The property being tested. This is set to 'HiddenFromAddressBook' by default.

.EXAMPLE
$users = Get-MgUser -All -Property DisplayName, HiddenFromAddressListsEnabled | Select-Object DisplayName, HiddenFromAddressListsEnabled
Test-365ACAddressBook -Users $users -OutputExcelFilePath "C:\Reports\AddressBookReport.xlsx"
This example retrieves all users with their DisplayName and HiddenFromAddressListsEnabled properties, then tests each user to see if they are hidden from the Address Book. The results are exported to an Excel file.

.EXAMPLE
Get-MgUser -All -Property DisplayName, HiddenFromAddressListsEnabled | Select-Object DisplayName, HiddenFromAddressListsEnabled | Test-365ACAddressBook -HtmlFilePath "C:\Reports\AddressBookReport.html"
This example pipes a list of users directly into Test-365ACAddressBook, which then checks if each user is hidden from the Address Book. The results are exported to an HTML file.

.NOTES
This function requires the Microsoft Graph PowerShell SDK to retrieve user information from Microsoft 365.

.LINK
https://docs.microsoft.com/powershell/module/microsoft.graph.users/get-mguser
#>
Function Test-365ACAddressBook {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, HiddenFromAddressListsEnabled | Select-Object DisplayName, HiddenFromAddressListsEnabled),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'HiddenFromAddressBook'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $isHiddenFromAddressBook = [bool]($user.HiddenFromAddressListsEnabled)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $isHiddenFromAddressBook
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

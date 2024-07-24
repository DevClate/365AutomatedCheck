<#
.SYNOPSIS
    Tests whether users have a street address and generates a report.

.DESCRIPTION
    The Test-365ACStreetAddress function tests whether users have a street address and generates a report. 
    It takes a list of users as input and checks if each user has a street address. 
    The function generates a report that includes the user's display name and whether they have a street address.
    The report can be exported to an Excel file or an HTML file.

.PARAMETER Users
    Specifies the list of users to test. If not provided, it retrieves all users from the Microsoft Graph API.

.PARAMETER OutputExcelFilePath
    Specifies the file path to export the report as an Excel file. The file must have a .xlsx extension.

.PARAMETER HtmlFilePath
    Specifies the file path to export the report as an HTML file. The file must have a .html extension.

.PARAMETER TestedProperty
    Specifies the name of the property being tested. Default value is 'Has Street Address'.

.EXAMPLE
    Test-365ACStreetAddress -Users $users -OutputExcelFilePath 'C:\Reports\StreetAddressReport.xlsx'

    This example tests the specified list of users for street addresses and exports the report to an Excel file.

.EXAMPLE
    Test-365ACStreetAddress -OutputExcelFilePath 'C:\Reports\StreetAddressReport.xlsx'

    This example retrieves all users from the Microsoft Graph API, tests them for street addresses, and exports the report to an Excel file.

#>

Function Test-365ACStreetAddress {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, StreetAddress | Select-Object DisplayName, StreetAddress),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Has Street Address'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $hasStreetAddress = [bool]($user.StreetAddress)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasStreetAddress
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
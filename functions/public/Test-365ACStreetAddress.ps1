<#
.SYNOPSIS
Tests whether users have a street address and generates a report.

.DESCRIPTION
The Test-365ACStreetAddress function tests whether users have a street address and generates a report. 
It takes a list of users as input and checks if each user has a street address. 
The function generates a report that includes the user's display name and whether they have a street address.
The report can be exported to an Excel file, an HTML file, or a Markdown file.

.PARAMETER Users
Specifies the list of users to test. If not provided, it retrieves all users from the Microsoft Graph API.

.PARAMETER TenantID
The ID of the tenant to connect to. Required if using app-only authentication.

.PARAMETER ClientID
The ID of the client to use for app-only authentication. Required if using app-only authentication.

.PARAMETER CertificateThumbprint
The thumbprint of the certificate to use for app-only authentication. Required if using app-only authentication.

.PARAMETER AccessToken
The access token to use for authentication. Required if using app-only authentication.

.PARAMETER InteractiveLogin
Specifies whether to use interactive login. If this switch is present, interactive login will be used. Otherwise, app-only authentication will be used.

.PARAMETER OutputExcelFilePath
Specifies the file path to export the report as an Excel file. The file must have a .xlsx extension.

.PARAMETER OutputHtmlFilePath
Specifies the file path to export the report as an HTML file. The file must have a .html extension.

.PARAMETER OutputMarkdownFilePath
Specifies the file path to export the report as a Markdown file. The file must have a .md extension.

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
        
        [Parameter(Mandatory = $false)]
        [string]$TenantID,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientID,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false)]
        [string]$AccessToken,
        
        [Parameter(Mandatory = $false)]
        [switch]$InteractiveLogin,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$OutputHtmlFilePath,
        
        [ValidatePattern('\.md$')]
        [string]$OutputMarkdownFilePath,
        
        [string]$TestedProperty = 'Has Street Address'
    )
    BEGIN {
        if ($InteractiveLogin) {
            Write-PSFMessage "Using interactive login..." -Level Host
            Connect-MgGraph -Scopes "User.Read.All", "AuditLog.read.All"  -NoWelcome
        }
        else {
            Write-PSFMessage "Using app-only authentication..." -Level Host
            Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint -Scopes "User.Read.All", "AuditLog.Read.All"
        }
        
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
            Write-PSFMessage "Excel report saved to $OutputExcelFilePath" -Level Host
        }
        elseif ($OutputHtmlFilePath) {
            Export-365ACResultToHtml -Results $results -OutputHtmlFilePath $OutputHtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
            Write-PSFMessage "HTML report saved to $OutputHtmlFilePath" -Level Host
        }
        elseif ($OutputMarkdownFilePath) {
            Export-365ACResultToMarkdown -Results $results -OutputMarkdownFilePath $OutputMarkdownFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
            Write-PSFMessage "Markdown report saved to $OutputMarkdownFilePath" -Level Host
        }
        else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
<#
.SYNOPSIS
Tests if users have a state assigned in their Microsoft 365 profiles.

.DESCRIPTION
The Test-365ACState function checks each user retrieved from Microsoft 365 for a state value. It processes a list of users, determining whether each user has a state assigned. The results can be exported to an Excel file, an HTML file, a Markdown file, or output directly.

.PARAMETER Users
An array of users to be tested. This can be piped in or specified directly. Each user should have DisplayName and State properties.

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
The file path where the Excel report will be saved. The file must have an .xlsx extension.

.PARAMETER OutputHtmlFilePath
The file path where the HTML report will be saved. The file must have an .html extension.

.PARAMETER OutputMarkdownFilePath
The file path where the Markdown report will be saved. The file must have an .md extension.

.PARAMETER TestedProperty
The property being tested. This is set to 'Has State' by default.

.EXAMPLE
PS> $users = Get-MgUser -All -Property DisplayName, State | Select-Object DisplayName, State
PS> Test-365ACState -Users $users -OutputExcelFilePath "C:\Reports\StateReport.xlsx"
This example retrieves all users with their DisplayName and State, then tests each user to see if they have a state assigned. The results are exported to an Excel file.

.EXAMPLE
PS> Get-MgUser -All -Property DisplayName, State | Select-Object DisplayName, State | Test-365ACState -OutputHtmlFilePath "C:\Reports\StateReport.html"
This example pipes a list of users directly into Test-365ACState, which then checks if each user has a state assigned. The results are exported to an HTML file.

.NOTES
This function requires the Microsoft Graph PowerShell SDK to retrieve user information from Microsoft 365.

.LINK
https://docs.microsoft.com/powershell/module/microsoft.graph.users/get-mguser
#>
Function Test-365ACState {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, State | Select-Object DisplayName, State),
        
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
        
        [string]$TestedProperty = 'Has State'
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
            $hasState = [bool]($user.State)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasState
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
<#
.SYNOPSIS
Tests the last login activity of users in Microsoft 365.

.DESCRIPTION
The Test-365ACLastLogin function tests the last login activity of users in Microsoft 365. It retrieves the required properties of users, calculates the number of inactive days, and determines if a user has logged in within a specified number of days. The function then generates a report of the test results.

.PARAMETER TenantID
The ID of the tenant to connect to. This parameter is optional.

.PARAMETER ClientID
The ID of the client to use for app-only authentication. This parameter is optional.

.PARAMETER CertificateThumbprint
The thumbprint of the certificate to use for app-only authentication. This parameter is optional.

.PARAMETER AccessToken
The access token to use for authentication. This parameter is optional.

.PARAMETER InteractiveLogin
Specifies whether to use interactive login. If this switch is present, interactive login will be used. Otherwise, app-only authentication will be used.

.PARAMETER Days
The number of days to consider a user as inactive. The default value is 30.

.PARAMETER TestedProperty
The name of the tested property. The default value is "Has Logged In Last <Days> Days".

.PARAMETER OutputExcelFilePath
The path to the output Excel file. The file should have a .xlsx extension. This parameter is optional.

.PARAMETER HtmlFilePath
The path to the output HTML file. The file should have a .html extension. This parameter is optional.

.EXAMPLE
Test-365ACLastLogin -TenantID "contoso.onmicrosoft.com" -ClientID "12345678-1234-1234-1234-1234567890ab" -CertificateThumbprint "AB12AB34AB56AB78AB90AB12AB34AB56AB78AB90" -Days 60 -OutputExcelFilePath "C:\Reports\LastLoginReport.xlsx"
Tests the last login activity of users in the "contoso.onmicrosoft.com" tenant using app-only authentication with the specified client ID and certificate thumbprint. Users who have not logged in within the last 60 days will be considered inactive. The test results will be exported to an Excel file located at "C:\Reports\LastLoginReport.xlsx".

.EXAMPLE
Test-365ACLastLogin -InteractiveLogin -Days 90 -OutputHtmlFilePath "C:\Reports\LastLoginReport.html"
Tests the last login activity of users using interactive login. Users who have not logged in within the last 90 days will be considered inactive. The test results will be exported to an HTML file located at "C:\Reports\LastLoginReport.html".
#>

Function Test-365ACLastLogin {
    [CmdletBinding()]
    param
    (
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
        
        [int]$Days = 30,
        
        [string]$TestedProperty = "Has Logged In Last $($Days) Days",
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$OutputHtmlFilePath,

        [ValidatePattern('\.md$')]
        [string]$OutputMarkdownFilePath
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
        $Count = 0
        $PrintedUser = 0
        $columnName = $TestedProperty
        #Retrieve users
        $RequiredProperties = @('UserPrincipalName', 'DisplayName', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle', 'SigninActivity')
        Get-MgBetaUser -All -Property $RequiredProperties | Select-Object $RequiredProperties | ForEach-Object {
            $Print = 1  # Initialize $Print to 1 at the start of each iteration
            $Count++
            # $UPN = $_.UserPrincipalName - use in future
            $LastSuccessfulSigninDate = $_.SignInActivity.lastSuccessfulSignInDateTime
            # $AccountEnabled = $_.AccountEnabled - use in future
            $DisplayName = $_.DisplayName
            #Calculate Inactive days
            if ($null -eq $LastSuccessfulSigninDate) { 
                $hasLoggedIn = $false
            }
            else {
                $InactiveUserDays = (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
                $hasLoggedIn = $InactiveUserDays -le $Days
            }
                
            if ($Print -eq 1) {
                $PrintedUser++   
            
                $result = [PSCustomObject]@{
                    "User Display Name" = $DisplayName
                    $columnName         = $hasLoggedIn
                }
                $results += $result
            }
        }
    }
    END {
        $TotalTests = $results.Count
        $PassedTests = ($results | Where-Object { $_.$columnName -eq $true }).Count
        $FailedTests = $TotalTests - $PassedTests
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
            <# Action when this condition is true #>
        }
        else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
        Disconnect-MgGraph
    }
}
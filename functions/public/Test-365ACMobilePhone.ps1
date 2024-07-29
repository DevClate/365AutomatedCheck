<#
.SYNOPSIS
    Tests whether users in Microsoft 365 have a valid mobile phone number.

.DESCRIPTION
    The Test-365ACMobilePhone function checks if users in Microsoft 365 have a mobile phone number and optionally validates these numbers against a provided list of valid numbers. It supports outputting the results to an Excel file, an HTML file, or the console.

.PARAMETER Users
    Specifies an array of users to test. If not provided, the function retrieves all users in Microsoft 365 with their DisplayName and MobilePhone properties.

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

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid mobile phone numbers. If provided, the function validates the users' mobile phone numbers against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to an Excel file where the results will be exported. Requires the ImportExcel module.

.PARAMETER OutputHtmlFilePath
    Specifies the path to an HTML file where the results will be exported. Requires the Export-365ACResultToHtml function to be defined.

.PARAMETER TestedProperty
    Specifies the property that is being tested. Defaults to 'Has Valid Mobile Phone' when validating against a list, otherwise 'Has Mobile Phone'.

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All) -OutputExcelFilePath "C:\Results.xlsx"
    Retrieves all users in Microsoft 365 and exports the validation results to an Excel file.

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All) -OutputHtmlFilePath "C:\Results.html"
    Retrieves all users in Microsoft 365 and exports the validation results to an HTML file.

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All)
    Retrieves all users in Microsoft 365 and outputs the validation results to the console.

.NOTES
    - Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results to Excel or HTML, respectively.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
function Test-365ACMobilePhone {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, MobilePhone | Select-Object DisplayName, MobilePhone),
        
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
        [string]$ValidationExcelFilePath,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$OutputHtmlFilePath,
        
        [string]$TestedProperty = 'Has Valid Mobile Phone'
    )
    BEGIN {
        $validMobilePhones = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
                return
            }
            # Import the Excel file to get valid mobile phone numbers
            $validMobilePhones = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty MobilePhone -Unique
        }

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
        $columnName = $ValidationExcelFilePath ? 'Has Valid Mobile Phone' : 'Has Mobile Phone'
        foreach ($user in $Users) {
            $hasMobilePhone = [bool]($user.MobilePhone)
            if ($ValidationExcelFilePath) {
                $hasMobilePhone = $user.MobilePhone -in $validMobilePhones
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName         = $hasMobilePhone
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
        } elseif ($OutputHtmlFilePath) {
            Export-365ACResultToHtml -Results $results -OutputHtmlFilePath $OutputHtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        } else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
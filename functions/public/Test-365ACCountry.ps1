<#
.SYNOPSIS
Tests whether users in Microsoft 365 have a country assigned and optionally validates against a list of valid countries.

.DESCRIPTION
The Test-365ACCountry function checks if users in Microsoft 365 have a country assigned to their profile. It can also validate these countries against a provided list of valid countries. The results of this test can be exported to an Excel file, an HTML file, or displayed in the console.
.PARAMETER Users
Specifies the users to be tested. If not specified, the function will test all users in the organization by retrieving their DisplayName and Country properties.

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
Specifies the path to an Excel file containing a list of valid countries. If specified, the function will validate the countries of the users against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to save the test results as an Excel file. If not specified, the test results will be displayed in the console. The path must end with '.xlsx'.

.PARAMETER OutputHtmlFilePath
Specifies the path to save the test results as an HTML file. If not specified, the test results will be displayed in the console. The path must end with '.html'.

.PARAMETER TestedProperty
    Specifies the name of the tested property. This will be used as the column name in the test results. Defaults to 'Has Country' but changes to 'Has Valid Country' if a validation list is provided.

.EXAMPLE
Test-365ACCountry -Users (Get-MgUser -All) -ValidationExcelFilePath "C:\Validation.xlsx" -OutputExcelFilePath "C:\Results.xlsx"
Tests all users in the organization, validates their countries against the list of valid countries in "Validation.xlsx", and saves the test results to "Results.xlsx".

.EXAMPLE
    Test-365ACCountry -Users (Get-MgUser -All) -OutputHtmlFilePath "C:\Results.html"
    Tests all users in the organization and saves the test results as an HTML file named "Results.html".

.NOTES
Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results to Excel or HTML, respectively.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
Function Test-365ACCountry {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, Country | Select-Object DisplayName, Country),
        
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
        [string]$TestedProperty = 'Has Country'
    )
    BEGIN {
        
        $validCountries = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
                return@
            }
            # Import the Excel file to get valid countries
            $validCountries = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty Country -Unique
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
        $columnName = $ValidationExcelFilePath ? 'Has Valid Country' : 'Has Country'
        foreach ($user in $Users) {
            $hasCountry = [bool]($user.Country)
            if ($ValidationExcelFilePath) {
                $hasCountry = $user.Country -in $validCountries
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName         = $hasCountry
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
        elseif ($OutputHtmlFilePath) {
            Export-365ACResultToHtml -Results $results -OutputHtmlFilePath $OutputHtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        }
        else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
<#
.SYNOPSIS
    Tests whether users in Microsoft 365 have a fax number assigned and exports the results.

.DESCRIPTION
    The Test-365ACFaxNumber function checks if users in Microsoft 365 have a fax number by examining the FaxNumber property. It can validate fax numbers against a provided list of valid numbers. The results can be exported to Excel, HTML, or output to the console.

.PARAMETER Users
    Specifies an array of users to test. If not provided, the function retrieves all users in Microsoft 365.

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
    Specifies the path to an Excel file containing a list of valid fax numbers. If provided, the function validates the users' fax numbers against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to export the results to an Excel file. If specified, the function exports the results using the Export-365ACResultToExcel function.

.PARAMETER OutputHtmlFilePath
    Specifies the path to export the results to an HTML file. If specified, the function exports the results using the Export-365ACResultToHtml function.

.PARAMETER TestedProperty
    Specifies the property being tested. Defaults to 'Has Fax Number' but changes to 'Has Valid Fax Number' if a validation list is provided.

.EXAMPLE
    Test-365ACFaxNumber -Users (Get-MgUser -All) -OutputExcelFilePath "C:\Results.xlsx"
    Tests all users for a fax number and exports the results to an Excel file.

.EXAMPLE
    Test-365ACFaxNumber -Users $users -OutputHtmlFilePath "C:\Results.html"
    Tests the specified users for a fax number and exports the results to an HTML file.

.NOTES
    - Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
Function Test-365ACFaxNumber {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, FaxNumber | Select-Object DisplayName, FaxNumber),
        
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
        
        [string]$TestedProperty = 'Has Fax Number'
    )
    BEGIN {
        $validFaxNumbers = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
                return
            }
            # Import the Excel file to get valid fax numbers
            $validFaxNumbers = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty FaxNumber -Unique
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
        $columnName = $ValidationExcelFilePath ? 'Has Valid Fax Number' : 'Has Fax Number'
        foreach ($user in $Users) {
            $hasFaxNumber = [bool]($user.FaxNumber)
            if ($ValidationExcelFilePath) {
                $hasFaxNumber = $user.FaxNumber -in $validFaxNumbers
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName         = $hasFaxNumber
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
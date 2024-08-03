<#
.SYNOPSIS
Tests whether users in Microsoft 365 have a department assigned and optionally validates against a list of valid departments.

.DESCRIPTION
The Test-365ACDepartment function checks if users in Microsoft 365 have a department assigned and can validate these departments against a provided list of valid departments. It supports exporting the results to an Excel file, an HTML file, or outputting directly to the console.

.PARAMETER Users
Specifies an array of users to test. If not provided, the function retrieves all users in Microsoft 365 with their DisplayName and Department properties.

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
Specifies the path to an Excel file containing a list of valid departments. If provided, the function validates the users' departments against this list.

.PARAMETER OutputExcelFilePath
Specifies the path to an Excel file where the results will be exported. Requires the ImportExcel module.

.PARAMETER OutputHtmlFilePath
Specifies the path to an HTML file where the results will be exported. Requires the Export-365ACResultToHtml function to be defined.

.PARAMETER OutputMarkdownFilePath
Specifies the path to a Markdown file where the results will be exported. Requires the Export-365ACResultToMarkdown function to be defined.

.PARAMETER TestedProperty
Specifies the property that is being tested. Defaults to 'Has Department' but changes to 'Has Valid Department' if a validation list is provided.

.EXAMPLE
Test-365ACDepartment -Users (Get-MgUser -All) -OutputExcelFilePath "C:\Results.xlsx"
Retrieves all users in Microsoft 365 and exports the validation results to an Excel file.

.EXAMPLE
Test-365ACDepartment -Users (Get-MgUser -All) -OutputHtmlFilePath "C:\Results.html"
Retrieves all users in Microsoft 365 and exports the validation results to an HTML file.

.EXAMPLE
Test-365ACDepartment -Users (Get-MgUser -All)
Retrieves all users in Microsoft 365 and outputs the validation results to the console.

.NOTES
- Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
- The Export-365ACResultToExcel, Export-365ACResultToHtml, and Export-365ACResultToMarkdown functions must be defined for exporting results to Excel, HTML, or Markdown, respectively.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
Function Test-365ACDepartment {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, Department | Select-Object DisplayName, Department),
        
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
        
        [ValidatePattern('\.md$')]
        [string]$OutputMarkdownFilePath,
        
        [string]$TestedProperty = 'Has Department'
    )
    BEGIN {
        $validDepartments = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Stop-PSFFunction -Message "Import-Excel cmdlet not found. Please install the ImportExcel module." -ErrorRecord $_ -Tag 'MissingModule'
                return
            }
            # Import the Excel file to get valid department names
        }
        if ($InteractiveLogin) {
            Write-PSFMessage "Using interactive login..." -Level Host
            Connect-MgGraph -Scopes "User.Read.All", "AuditLog.read.All" -NoWelcome
        } else {
            Write-PSFMessage "Using app-only authentication..." -Level Host
            Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint -Scopes "User.Read.All", "AuditLog.Read.All"
        }
        $results = @()
    }
    PROCESS {
        $results = @()
        foreach ($user in $Users) {
            $hasDepartment = [bool]($user.Department)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasDepartment
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
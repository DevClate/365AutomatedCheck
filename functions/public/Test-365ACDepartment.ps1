<#
.SYNOPSIS
    Tests whether users in Microsoft 365 have a department assigned and optionally validates against a list of valid departments.

.DESCRIPTION
    The Test-365ACDepartment function checks if users in Microsoft 365 have a department assigned and can validate these departments against a provided list of valid departments. It supports exporting the results to an Excel file, an HTML file, or outputting directly to the console.

.PARAMETER Users
    Specifies an array of users to test. If not provided, the function retrieves all users in Microsoft 365 with their DisplayName and Department properties.

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid departments. If provided, the function validates the users' departments against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to an Excel file where the results will be exported. Requires the ImportExcel module.

.PARAMETER HtmlFilePath
    Specifies the path to an HTML file where the results will be exported. Requires the Export-365ACResultToHtml function to be defined.

.PARAMETER TestedProperty
    Specifies the property that is being tested. Defaults to 'Has Department' but changes to 'Has Valid Department' if a validation list is provided.

.EXAMPLE
    Test-365ACDepartment -Users (Get-MgUser -All) -OutputExcelFilePath "C:\Results.xlsx"
    Retrieves all users in Microsoft 365 and exports the validation results to an Excel file.

.EXAMPLE
    Test-365ACDepartment -Users (Get-MgUser -All) -HtmlFilePath "C:\Results.html"
    Retrieves all users in Microsoft 365 and exports the validation results to an HTML file.

.EXAMPLE
    Test-365ACDepartment -Users (Get-MgUser -All)
    Retrieves all users in Microsoft 365 and outputs the validation results to the console.

.NOTES
    - Requires the ImportExcel module for exporting results to Excel. If not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions must be defined for exporting results to Excel or HTML, respectively.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
Function Test-365ACDepartment {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, Department | Select-Object DisplayName, Department),
        
        [ValidatePattern('\.xlsx$')]
        [string]$ValidationExcelFilePath,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
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
            $validDepartments = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty Department -Unique
        }
        $results = @()
    }
    PROCESS {
        $columnName = $ValidationExcelFilePath ? 'Has Valid Department' : 'Has Department'
        foreach ($user in $Users) {
            $hasDepartment = [bool]($user.Department)
            if ($ValidationExcelFilePath) {
                $hasDepartment = $user.Department -in $validDepartments
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName = $hasDepartment
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
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $columnName
        }
        else {
            Write-PSFMessage -Level Output -Message ($results | Out-String)
        }
    }
}
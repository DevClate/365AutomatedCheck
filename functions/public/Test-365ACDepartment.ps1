<#
.SYNOPSIS
    This function tests the department property of users and exports the results to Excel, HTML, or the console.

.DESCRIPTION
    The Test-365ACDepartment function tests the department property of users and generates test results. It takes an array of users as input and checks if each user has a department specified. The function then generates test results indicating whether each user has a department or not.

.PARAMETER Users
    Specifies the array of users to be tested. The users should have the 'DisplayName' and 'Department' properties.

.PARAMETER ValidationExcelFilePath
    Specifies the path to an Excel file containing a list of valid Departments. If specified, the function will validate the Departments of the users against this list.

.PARAMETER OutputExcelFilePath
    Specifies the path to the Excel file where the test results will be exported. If this parameter is specified, the function will use the Export-365ACResultToExcel function to export the results to an Excel file.

.PARAMETER HtmlFilePath
    Specifies the path to the HTML file where the test results will be exported. If this parameter is specified, the function will use the Export-365ACResultToHtml function to export the results to an HTML file.

.PARAMETER TestedProperty
    Specifies the property that is being tested. Default is 'Has Department'.

.EXAMPLE
    Test-365ACDepartment -Users $users -OutputExcelFilePath "C:\TestResults.xlsx"
    Tests the department property of the specified users and exports the test results to an Excel file.

.EXAMPLE
    Test-365ACDepartment -Users $users -HtmlFilePath "C:\TestResults.html"
    Tests the department property of the specified users and exports the test results to an HTML file.

.NOTES
    - This function requires the ImportExcel module to export results to Excel. If the module is not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions are assumed to be defined elsewhere in the script.

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
                Write-Error "Import-Excel cmdlet not found. Please install the ImportExcel module."
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
            Write-Output $results
        }
    }
}
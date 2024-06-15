<#
.SYNOPSIS
    This function tests the department property of users and generates test results.

.DESCRIPTION
    The Test-365ACDepartment function tests the department property of users and generates test results. It takes an array of users as input and checks if each user has a department specified. The function then generates test results indicating whether each user has a department or not.

.PARAMETER Users
    Specifies the array of users to be tested. The users should have the 'DisplayName' and 'Department' properties.

.PARAMETER ExcelFilePath
    Specifies the path to the Excel file where the test results will be exported. If this parameter is specified, the function will use the Export-365ACResultToExcel function to export the results to an Excel file.

.PARAMETER HtmlFilePath
    Specifies the path to the HTML file where the test results will be exported. If this parameter is specified, the function will use the Export-365ACResultToHtml function to export the results to an HTML file.

.EXAMPLE
    Test-365ACDepartment -Users $users -ExcelFilePath "C:\TestResults.xlsx"
    Tests the department property of the specified users and exports the test results to an Excel file.

.EXAMPLE
    Test-365ACDepartment -Users $users -HtmlFilePath "C:\TestResults.html"
    Tests the department property of the specified users and exports the test results to an HTML file.

.NOTES
    This function requires the ImportExcel module to be installed in order to export the results to an Excel file.

    The function generates test results indicating whether each user has a department or not. The total number of tests, number of passed tests, and number of failed tests are also calculated and displayed.

    If neither the ExcelFilePath nor the HtmlFilePath parameters are specified, the function will output the test results to the console.
#>
Function Test-365ACDepartment
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, Department | Select-Object displayname, Department),
        [ValidatePattern('\.xlsx$')]
        [string]$ExcelFilePath,
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath
    )
    BEGIN
    {
        if ($ExcelFilePath -and !(Get-Command Export-Excel -ErrorAction SilentlyContinue))
        {
            Write-Error "Export-Excel cmdlet not found. Please install the ImportExcel module."
            return
        }
        $results = @()
    }
    PROCESS
    {
        foreach ($user in $Users)
        {
            #Write-Output "Checking user $($user.DisplayName) with department $($user.Department)"
            $hasDepartment = [bool]($user.Department)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Department' = $hasDepartment
            }
            $results += $result
        }
    }
    END
    {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Department' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath)
        {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath)
        {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Department'
        }
        else
        {
            Write-Output $results
        }
    }
}
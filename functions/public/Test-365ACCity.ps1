<#
.SYNOPSIS
    Tests if users have a city property and generates test results.

.DESCRIPTION
    The Test-365ACCity function tests if users have a city property and generates test results. It takes an array of users as input and checks if each user has a city property. The test results are stored in an array of custom objects, which include the user's display name and the result of the test.

.PARAMETER Users
    Specifies the array of users to test. Each user should have a DisplayName and City property.

.PARAMETER OutputExcelFilePath
    Specifies the path to the output Excel file. If provided, the test results will be exported to an Excel file.

.PARAMETER HtmlFilePath
    Specifies the path to the output HTML file. If provided, the test results will be exported to an HTML file.

.PARAMETER TestedProperty
    Specifies the name of the tested property. Default value is 'Has City'.

.INPUTS
    An array of users with DisplayName and City properties.

.OUTPUTS
    If OutputExcelFilePath or HtmlFilePath is not provided, the function outputs an array of custom objects with the user's display name and the result of the test.

.EXAMPLE
    $users = Get-MgUser -All -Property DisplayName, City | Select-Object DisplayName, City
    Test-365ACCity -Users $users -OutputExcelFilePath 'C:\TestResults.xlsx'

    This example tests if the users have a city property and exports the test results to an Excel file.

#>
Function Test-365ACCity {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All -Property DisplayName, City | Select-Object DisplayName, City),
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,
        
        [string]$TestedProperty = 'Has City'
    )
    BEGIN {
        $results = @()
    }
    PROCESS {
        foreach ($user in $Users) {
            $hasCity = [bool]($user.City)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $TestedProperty     = $hasCity
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
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty $TestedProperty
        }
        else {
            Write-Output $results
        }
    }
}
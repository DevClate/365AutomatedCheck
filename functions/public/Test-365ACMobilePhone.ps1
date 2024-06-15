<#
.SYNOPSIS
    This function tests whether users have a mobile phone number associated with their account.

.DESCRIPTION
    The Test-365ACMobilePhone function tests whether users have a mobile phone number associated with their account. It retrieves a list of users and checks if each user has a mobile phone number. The function outputs the results in the form of a custom object with the user's display name and a boolean value indicating whether they have a mobile phone.

.PARAMETER Users
    Specifies an array of users to test. If not provided, the function retrieves all users in Microsoft 365.

.PARAMETER ExcelFilePath
    Specifies the path to an Excel file where the results will be exported. If this parameter is provided, the function requires the ImportExcel module to be installed.

.PARAMETER HtmlFilePath
    Specifies the path to an HTML file where the results will be exported. If this parameter is provided, the function requires the Export-365ACResultToHtml function to be available.

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All) -ExcelFilePath "C:\Results.xlsx"
    Retrieves all users in Microsoft 365 and exports the results to an Excel file located at "C:\Results.xlsx".

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All) -HtmlFilePath "C:\Results.html"
    Retrieves all users in Microsoft 365 and exports the results to an HTML file located at "C:\Results.html".

.EXAMPLE
    Test-365ACMobilePhone -Users (Get-MgUser -All)
    Retrieves all users in Microsoft 365 and outputs the results to the console.

.NOTES
    - This function requires the ImportExcel module to export results to Excel. If the module is not installed, an error will be displayed.
    - The Export-365ACResultToExcel and Export-365ACResultToHtml functions are assumed to be defined elsewhere in the script.

.LINK
    https://github.com/DevClate/365AutomatedCheck
#>
function Test-365ACMobilePhone {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (Get-MgUser -All),

        [ValidatePattern('\.xlsx$')]
        [string]$ExcelFilePath,

        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath
    )

    BEGIN {
        if ($ExcelFilePath -and !(Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
            Write-Error "Export-Excel cmdlet not found. Please install the ImportExcel module."
            return
        }
        $results = @()
    }

    PROCESS {
        foreach ($user in $Users) {
            $hasMobilePhone = [bool]($user.MobilePhone)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Mobile Phone'  = $hasMobilePhone
            }
            $results += $result
        }
    }

    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Mobile Phone' }).Count
        $failedTests = $totalTests - $passedTests

        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Mobile Phone'
        }
        else {
            Write-Output $results
        }
    }
}
function Test-365ACMobilePhone {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
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
                'Has Mobile Phone' = $hasMobilePhone
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
        } elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Mobile Phone'
        } else {
            Write-Output $results
        }
    }
}
function Test-365ACJobTitle {
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
            $hasJobTitle = [bool]($user.JobTitle)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Job Title' = $hasJobTitle
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Job Title' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        } elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Job Title'
        } else {
            Write-Output $results
        }
    }
}
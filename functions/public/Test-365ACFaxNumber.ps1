Function Test-365ACFaxNumber {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, FaxNumber | Select-Object displayname, FaxNumber),
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
            $hasFaxNumber = [bool]($user.FaxNumber)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Fax Number'    = $hasFaxNumber
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Fax Number' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath) {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath) {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Fax Number'
        }
        else {
            Write-Output $results
        }
    }
}
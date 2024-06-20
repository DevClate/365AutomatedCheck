Function Test-365ACCompanyName {
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, CompanyName | Select-Object DisplayName, CompanyName),
        
        [ValidatePattern('\.xlsx$')]
        [string]$ValidationExcelFilePath,
        
        [ValidatePattern('\.xlsx$')]
        [string]$OutputExcelFilePath,
        
        [ValidatePattern('\.html$')]
        [string]$HtmlFilePath,

        [string]$TestedProperty = 'Has Company Name'
    )
    BEGIN {
        $validCompanyNames = @()
        if ($ValidationExcelFilePath) {
            if (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
                Write-Error "Import-Excel cmdlet not found. Please install the ImportExcel module."
                return@
            }
            # Import the Excel file to get valid company names
            $validCompanyNames = Import-Excel -Path $ValidationExcelFilePath | Select-Object -ExpandProperty CompanyName -Unique
        }
        $results = @()
    }
    PROCESS {
        $columnName = $ValidationExcelFilePath ? 'Has Valid Company Name' : 'Has Company Name'
        foreach ($user in $Users) {
            $hasCompanyName = [bool]($user.CompanyName)
            if ($ValidationExcelFilePath) {
                $hasCompanyName = $user.CompanyName -in $validCompanyNames
            }
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                $columnName = $hasCompanyName
            }
            $results += $result
        }
    }
    END {
        $totalTests = $results.Count
        # Update passed and failed tests calculation based on the new logic
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
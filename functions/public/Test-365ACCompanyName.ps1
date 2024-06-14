Function Test-365ACCompanyName
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Users = (get-mguser -all -Property DisplayName, CompanyName | Select-Object displayname, CompanyName),
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
            $hasCompanyName = [bool]($user.CompanyName)
            $result = [PSCustomObject]@{
                'User Display Name' = $user.DisplayName
                'Has Company Name' = $hasCompanyName
            }
            $results += $result
        }
    }
    END
    {
        $totalTests = $results.Count
        $passedTests = ($results | Where-Object { $_.'Has Company Name' }).Count
        $failedTests = $totalTests - $passedTests
        if ($ExcelFilePath)
        {
            Export-365ACResultToExcel -Results $results -ExcelFilePath $ExcelFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests
        }
        elseif ($HtmlFilePath)
        {
            Export-365ACResultToHtml -Results $results -HtmlFilePath $HtmlFilePath -TotalTests $totalTests -PassedTests $passedTests -FailedTests $failedTests -TestedProperty 'Has Company Name'
        }
        else
        {
            Write-Output $results
        }
    }
}
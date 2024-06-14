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
<#
.SYNOPSIS
Exports 365AutomatedCheck results to a Markdown file.
.DESCRIPTION
The Export-365ACResultToMarkdown function takes an array of results, along with other parameters, and generates a Markdown file that displays the results in a table format. The function calculates the total number of tests, the number of passed tests, the number of failed tests, and the number of tests that were not tested.
.PARAMETER Results
The array of results containing the test data. Each element in the array should be an object with properties representing the test results.
.PARAMETER OutputMarkdownFilePath
The path to the Markdown file where the results will be exported.
.PARAMETER TotalTests
The total number of tests.
.PARAMETER PassedTests
The number of tests that passed.
.PARAMETER FailedTests
The number of tests that failed.
.PARAMETER TestedProperty
The name of the property in the test results object that indicates whether a test was passed or failed.
.EXAMPLE
$Results = @(
    [PSCustomObject]@{
        'User Display Name' = 'John Doe'
        'Test Property' = 'Yes'
    },
    [PSCustomObject]@{
        'User Display Name' = 'Jane Smith'
        'Test Property' = 'No'
    }
)
$OutputMarkdownFilePath = '/path/to/results.md'
$TotalTests = $Results.Count
$PassedTests = ($Results | Where-Object { $_.'Test Property' -eq 'Yes' }).Count
$FailedTests = ($Results | Where-Object { $_.'Test Property' -eq 'No' }).Count
$TestedProperty = 'Test Property'
Export-365ACResultToMarkdown -Results $Results -OutputMarkdownFilePath $OutputMarkdownFilePath -TotalTests $TotalTests -PassedTests $PassedTests -FailedTests $FailedTests -TestedProperty $TestedProperty
.NOTES
This function requires PowerShell version 5.1 or above.
#>
function Export-365ACResultToMarkdown {
    param (
        [array]$Results,
        [string]$OutputMarkdownFilePath,
        [int]$TotalTests,
        [int]$PassedTests,
        [int]$FailedTests,
        [string]$TestedProperty
    )
    $markdown = @"
# 365AutomatedCheck Results

## Summary
- **Total tests:** $TotalTests
- **Passed:** $PassedTests
- **Failed:** $FailedTests
- **Not tested:** 0

## Results

| User Display Name | $TestedProperty |
|-------------------|-----------------|
"@
    # Write the initial markdown content to the file
    Set-Content -Path $OutputMarkdownFilePath -Value $markdown

    # Append each result to the file
    foreach ($result in $Results) {
        $userName = $result.'User Display Name' -replace '\|', '\\|'
        $testProperty = $result.$TestedProperty -replace '\|', '\\|'
        $line = "| $userName | $testProperty |"
        Add-Content -Path $OutputMarkdownFilePath -Value $line
    }
}
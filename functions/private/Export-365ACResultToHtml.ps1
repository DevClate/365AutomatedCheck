<#
.SYNOPSIS
Exports 365AutomatedCheck results to an HTML file.

.DESCRIPTION
The Export-365ACResultToHtml function takes an array of results, along with other parameters, and generates an HTML file that displays the results in a table format. The function calculates the total number of tests, the number of passed tests, the number of failed tests, and the number of tests that were not tested. It also provides filter buttons to toggle the visibility of different types of test results.

.PARAMETER Results
The array of results containing the test data. Each element in the array should be an object with properties representing the test results.

.PARAMETER OutputHtmlFilePath
The path to the HTML file where the results will be exported.

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

$OutputHtmlFilePath = '/path/to/results.html'
$TotalTests = $Results.Count
$PassedTests = ($Results | Where-Object { $_.'Test Property' -eq 'Yes' }).Count
$FailedTests = ($Results | Where-Object { $_.'Test Property' -eq 'No' }).Count
$TestedProperty = 'Test Property'

Export-365ACResultToHtml -Results $Results -OutputHtmlFilePath $OutputHtmlFilePath -TotalTests $TotalTests -PassedTests $PassedTests -FailedTests $FailedTests -TestedProperty $TestedProperty

.NOTES
This function requires PowerShell version 5.1 or above.
#>

function Export-365ACResultToHtml {
    param (
        [array]$Results,
        [string]$OutputHtmlFilePath,
        [int]$TotalTests,
        [int]$PassedTests,
        [int]$FailedTests,
        [string]$TestedProperty
    )

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>365AutomatedCheck Results</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #1e1e1e;
            color: #fff;
        }
        .container {
            width: 90%;
            margin: auto;
            overflow: hidden;
        }
        header {
            background: #000000;
            color: white;
            padding-top: 30px;
            min-height: 70px;
            border-bottom: #e8491d 3px solid;
        }
        header h1 {
            padding: 5px 0;
            text-align: center;
        }
        .summary-box {
            display: flex;
            justify-content: space-around;
            padding: 20px;
            background-color: #2c2c2c;
            margin-bottom: 20px;
        }
        .summary-item {
            text-align: center;
            flex: 1;
            margin: 10px;
        }
        .summary-item h2 {
            margin: 0;
            font-size: 2em;
        }
        .summary-item p {
            margin: 5px 0 0 0;
            font-size: 1.2em;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background-color: #333;
            color: #fff;
        }
        th, td {
            padding: 15px;
            text-align: left;
            border-bottom: 1px solid #444;
        }
        .success {
            background-color: #4CAF50;
            color: white;
        }
        .failure {
            background-color: #f44336;
            color: white;
        }
        .not-tested {
            background-color: #777;
            color: white;
        }
        .filter-buttons {
            text-align: center;
            margin-bottom: 20px;
        }
        .filter-buttons button {
            padding: 10px 20px;
            margin: 0 5px;
            background-color: #444;
            color: #fff;
            border: none;
            cursor: pointer;
        }
        .filter-buttons button:hover {
            background-color: #666;
        }
        .nowrap {
            white-space: nowrap;
        }
    </style>
    <script>
        function filterTests(filter) {
            var rows = document.querySelectorAll('table tr.test-row');
            rows.forEach(function(row) {
                if (filter === 'all' || row.classList.contains(filter)) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }

                var failedMessageCell = row.querySelector('.failure-message');
                if (failedMessageCell) {
                    if (filter === 'Success') {
                        failedMessageCell.classList.add('nowrap');
                    } else {
                        failedMessageCell.classList.remove('nowrap');
                    }
                }
            });
        }
    </script>
</head>
<body>
    <header>
        <div class="container">
            <h1>365AutomatedCheck Results</h1>
        </div>
    </header>
    <div class="container">
        <div class="summary-box">
            <div class="summary-item">
                <h2>$TotalTests</h2>
                <p>Total tests</p>
            </div>
            <div class="summary-item">
                <h2>$PassedTests</h2>
                <p>Passed</p>
            </div>
            <div class="summary-item">
                <h2>$FailedTests</h2>
                <p>Failed</p>
            </div>
            <div class="summary-item">
                <h2>0</h2>
                <p>Not tested</p>
            </div>
        </div>
        <div class="filter-buttons">
            <button onclick="filterTests('all')">All</button>
            <button onclick="filterTests('success')">Passed</button>
            <button onclick="filterTests('failure')">Failed</button>
        </div>
        <table>
            <tr>
                <th>User Display Name</th>
                <th>$TestedProperty</th>
            </tr>
"@
    foreach ($result in $Results) {
        $html += "<tr class='test-row "
        $html += if ($result.$TestedProperty -eq 'Yes') { 'success' } else { 'failure' }
        $html += "'>"
        $html += "<td>$($result.'User Display Name')</td>"
        if ($result.$TestedProperty -eq 'Yes') {
            $html += "<td class='success'>$($result.$TestedProperty)</td>"
        } else {
            $html += "<td class='failure'>$($result.$TestedProperty)</td>"
        }
        $html += "</tr>"
    }
    $html += @"
        </table>
    </div>
</body>
</html>
"@
    Set-Content -Path $OutputHtmlFilePath -Value $html
}

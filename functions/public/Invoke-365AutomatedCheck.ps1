function Invoke-365AutomatedCheck {
    [CmdletBinding()]
    param (
        [hashtable] $PesterConfiguration,
        [ValidateSet('None', 'Normal', 'Detailed', 'Diagnostic')]
        [string] $Verbosity = 'None',
        [string] $XsltPath = "$RootPath/functions/private/DefaultReportConfig.xslt",
        [string] $XmlPath,
        [string] $Path = "$RootPath/tests/",
        [string] $OutputHtmlPath,
        [bool] $PassThru = $false,
        [string[]] $Tag,
        [string[]] $ExcludeTag,
        [string] $ExcelFilePath
    )

    #Requires -Module Pester, ImportExcel

    $XmlPath = Set-365ACDefaultPath -Path $XmlPath -DefaultPath '365ACReport.xml'
    Write-Host "Using XML Path: $XmlPath"

    $OutputHtmlPath = Set-365ACDefaultPath -Path $OutputHtmlPath -DefaultPath '365ACReport.html'
    Write-Host "Using HTML Output Path: $OutputHtmlPath"

    $pesterConfig = Get-365ACPesterConfiguration -Path $Path -Tag $Tag -ExcludeTag $ExcludeTag -XmlPath $XmlPath -PesterConfiguration $PesterConfiguration -Verbosity $Verbosity -PassThru $PassThru

    $env:ExcelFilePath = $ExcelFilePath

    Invoke-Pester -Configuration $pesterConfig

    Start-Sleep -Seconds 2

    Convert-365ACXmlToHtml -XmlPath $XmlPath -XsltPath $XsltPath -OutputHtmlPath $OutputHtmlPath
}
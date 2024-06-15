<#
.SYNOPSIS
    Invokes the 365AutomatedCheck function to perform automated checks on Office 365.

.DESCRIPTION
    The Invoke-365AutomatedCheck function is used to perform automated checks on Office 365. It runs Pester tests and generates an HTML report based on the results.

.PARAMETER PesterConfiguration
    Specifies the Pester configuration hashtable to be used for running the tests.

.PARAMETER Verbosity
    Specifies the verbosity level for the Pester output. Valid values are 'None', 'Normal', 'Detailed', and 'Diagnostic'. The default value is 'None'.

.PARAMETER XsltPath
    Specifies the path to the XSLT file used for transforming the XML report to HTML. The default path is "$RootPath/functions/private/DefaultReportConfig.xslt".

.PARAMETER XmlPath
    Specifies the path to the XML report file generated by the Pester tests. If not provided, the default path is set to '365ACReport.xml'.

.PARAMETER Path
    Specifies the path to the Pester test files. The default path is "$RootPath/tests/".

.PARAMETER OutputHtmlPath
    Specifies the path to save the generated HTML report. If not provided, the default path is set to '365ACReport.html'.

.PARAMETER PassThru
    Indicates whether to pass the Pester test results through the pipeline. By default, it is set to $false.

.PARAMETER Tag
    Specifies the tags to include when running the Pester tests. Multiple tags can be specified using an array.

.PARAMETER ExcludeTag
    Specifies the tags to exclude when running the Pester tests. Multiple tags can be specified using an array.

.PARAMETER ExcelFilePath
    Specifies the path to the Excel file used for additional reporting. This parameter is optional.

.EXAMPLE
    Invoke-365AutomatedCheck -PesterConfiguration $config -Verbosity 'Detailed' -XmlPath 'C:\Reports\365ACReport.xml' -OutputHtmlPath 'C:\Reports\365ACReport.html' -Tag 'Basic', 'HR'

    This example runs the 365AutomatedCheck function with a detailed verbosity level, specifying the XML report path and the output HTML path. It also includes tests with the 'Basic' and 'HR' tags.

.NOTES
    This function requires the Pester and ImportExcel modules to be installed.

#>

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
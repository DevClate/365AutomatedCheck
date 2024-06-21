<#
.SYNOPSIS
    Retrieves the Pester configuration for the 365AutomatedCheck module.

.DESCRIPTION
    The Get-365ACPesterConfiguration function retrieves the Pester configuration for the 365AutomatedCheck module. It allows you to customize the configuration by specifying various parameters such as the path to run the tests, tags to include or exclude, XML output path, Pester configuration hashtable, verbosity level, and whether to pass through the configuration object.

.PARAMETER Path
    The path where the Pester tests should be run.

.PARAMETER Tag
    An array of tags to include when running the tests.

.PARAMETER ExcludeTag
    An array of tags to exclude when running the tests.

.PARAMETER XmlPath
    The path where the XML test results should be saved.

.PARAMETER PesterConfiguration
    A hashtable containing additional Pester configuration settings.

.PARAMETER Verbosity
    The verbosity level for the Pester output.

.PARAMETER PassThru
    Specifies whether to return the Pester configuration object.

.OUTPUTS
    PesterConfiguration
        The Pester configuration object.

.EXAMPLE
    PS> Get-365ACPesterConfiguration -Path "C:\Tests" -Tag "Unit" -XmlPath "C:\TestResults.xml" -Verbosity "Detailed"

    This example retrieves the Pester configuration for the 365AutomatedCheck module. It sets the path to run the tests to "C:\Tests", includes only tests with the "Unit" tag, saves the XML test results to "C:\TestResults.xml", and sets the verbosity level to "Detailed".

#>
function Get-365ACPesterConfiguration {
    param (
        [string] $Path,
        [string[]] $Tag,
        [string[]] $ExcludeTag,
        [string] $XmlPath,
        [hashtable] $PesterConfiguration,
        [string] $Verbosity,
        [bool] $PassThru
    )
    
    $config = [PesterConfiguration]::Default
    $config.Run.Path = $Path
    $config.Run.PassThru = $PassThru
    $config.Output.Verbosity = $Verbosity
    $config.TestResult.Enabled = $true
    $config.TestResult.OutputPath = $XmlPath
    $config.TestResult.OutputFormat = 'NUnitXml'

    if ($Tag) { $config.Filter.Tag = $Tag }
    if ($ExcludeTag) { $config.Filter.ExcludeTag = $ExcludeTag }

    return $config
}

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

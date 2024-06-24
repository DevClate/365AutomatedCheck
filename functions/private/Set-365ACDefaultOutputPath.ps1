<#
.SYNOPSIS
Sets the default output path for 365 Automated Check reports.

.DESCRIPTION
The Set-365ACDefaultOutputPath function generates a default output path for saving reports. If a path is not provided, it creates a directory based on the current date and a provided default path name. This ensures that reports are organized by date and are easily locatable.

.PARAMETER Path
The custom path where the report should be saved. If this parameter is not provided, a default path is generated.

.PARAMETER DefaultPath
A default directory name to be used when generating the default path. This is appended to the date-based directory name.

.EXAMPLE
Set-365ACDefaultOutputPath -DefaultPath "MyReport"
This example generates a default output path for "MyReport" using the current date and saves the report in the "365ACReports" directory under the generated path.

.EXAMPLE
Set-365ACDefaultOutputPath -Path "C:\CustomReports\MyReport" -DefaultPath "IgnoredInThisCase"
This example sets the output path to "C:\CustomReports\MyReport". The DefaultPath parameter is ignored since a custom path is provided.

.INPUTS
None. You cannot pipe input to this function.

.OUTPUTS
String. Returns the full path where the report will be saved.

.NOTES
This function checks if the specified directory exists and creates it if it does not. It uses the current date to generate a unique directory for each day, ensuring that reports are not overwritten and are easy to find based on the date.
#>
function Set-365ACDefaultOutputPath {
    param (
        [string] $Path,
        [string] $DefaultPath
    )
    if (-not $Path) {
        # today's date and time in the format yyyyMMdd-HHmm
        $todaysDate = (Get-Date).ToString("yyyyMMdd-HHmm")
        
        # Ensure cross-platform compatibility by using Join-Path for each segment
        $reportsPath = Join-Path -Path (Get-Location) -ChildPath "365ACReports"
        $datePath = "$todaysDate-$DefaultPath"
        $Path = Join-Path -Path $reportsPath -ChildPath $datePath
        
        $folderPath = [System.IO.Path]::GetDirectoryName($Path)

        if (-not (Test-Path -Path $folderPath)) {
            Write-PSFMessage -Level Host -Message "Creating directory: $folderPath"
            New-Item -ItemType Directory -Path $folderPath -Force
        } else {
            Write-PSFMessage -Level Host -Message "Directory already exists: $folderPath"
        }
    }

    return $Path
}
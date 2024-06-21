<#
.SYNOPSIS
Sets the default output path for 365ACReports.

.DESCRIPTION
This function sets the default output path for 365ACReports if the provided path is empty. It creates a directory with the current date and time appended to the default path.

.PARAMETER Path
The path to set as the output path. If not provided, the function will generate a path based on the default path and the current date and time.

.PARAMETER DefaultPath
The default path to use if the provided path is empty.

.EXAMPLE
Set-365ACDefaultOutputPath -Path "C:\Reports" -DefaultPath "Output"

This example sets the output path to "C:\Reports" if provided, otherwise it generates a path based on the default path "Output" and the current date and time.

.OUTPUTS
System.String
The output path that was set.

.NOTES
Author: Clayton Tyger
Date: 06/20/2024
#>
function Set-365ACDefaultOutputPath {
    param (
        [string] $Path,
        [string] $DefaultPath
    )
    if (-not $Path) {
        # today's date and time in the format yyyyMMdd hhmmss
        $todaysDate = (Get-Date).ToString("yyyyMMdd-HHmm")
        $Path = Join-Path -Path (Get-Location) -ChildPath "365ACReports/$todaysDate-$DefaultPath"
        $folderPath = [System.IO.Path]::GetDirectoryName($Path)

        if (-not (Test-Path -Path $folderPath)) {
            Write-Host "Creating directory: $folderPath"
            New-Item -ItemType Directory -Path $folderPath -Force
        } else {
            Write-Host "Directory already exists: $folderPath"
        }
    }
    return $Path
}
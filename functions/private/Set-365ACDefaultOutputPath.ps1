<#
.SYNOPSIS
Sets the default output path for the 365ACReports.

.DESCRIPTION
This function sets the default output path for the 365ACReports. If a specific path is provided, it checks if it is an absolute path. If not, it appends it to the base reports path. If no path is provided, it creates a new path based on the current date and the default path.

.PARAMETER Path
The specific path to set as the default output path. If not provided, a new path will be created based on the current date and the default path.

.PARAMETER DefaultPath
The default path to use when creating a new path based on the current date.

.EXAMPLE
Set-365ACDefaultOutputPath -Path "C:\Reports\" -DefaultPath "Output"

This example sets the default output path to "C:\Reports".

.OUTPUTS
The function returns the final output path.

.NOTES
This function requires the Write-PSFMessage function from the PowerShell Framework module.

.LINK
https://github.com/DevClate/365AutomatedCheck

#>
function Set-365ACDefaultOutputPath {
    param (
        [string] $Path,
        [ValidateNotNullOrEmpty()]
        [string] $DefaultPath
    )

    # Initialize the base reports path
    $reportsPath = Join-Path -Path (Get-Location) -ChildPath "365ACReports"

    if (-not $Path) {
        $todaysDate = (Get-Date).ToString("yyyyMMdd-HHmm")
        $datePath = "$todaysDate-$DefaultPath"
        $Path = Join-Path -Path $reportsPath -ChildPath $datePath
    }
    else {
        # Check if the provided Path is an absolute path
        if ([System.IO.Path]::IsPathRooted($Path)) {
            # If it's an absolute path, use it directly
        }
        else {
            # If it's not an absolute path, append it to the $reportsPath
            $Path = Join-Path -Path $reportsPath -ChildPath $Path
        }
    }

    $folderPath = [System.IO.Path]::GetDirectoryName($Path)

    try {
        if (-not (Test-Path -Path $folderPath)) {
            Write-PSFMessage -Level Host -Message "Creating directory: $folderPath"
            New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
        }
        else {
            Write-PSFMessage -Level Host -Message "Using existing directory: $folderPath"
        }
    }
    catch {
        Write-PSFMessage -Level Error -Message "Failed to create or access directory: $folderPath. Error: $_"
        throw
    }

    return $Path
}
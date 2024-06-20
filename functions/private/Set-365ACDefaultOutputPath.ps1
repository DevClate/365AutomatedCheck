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
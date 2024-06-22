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
            Write-Host "Creating directory: $folderPath"
            New-Item -ItemType Directory -Path $folderPath -Force
        } else {
            Write-Host "Directory already exists: $folderPath"
        }
    }

    return $Path
}
function Set-365ACDefaultPath {
    param (
        [string] $Path,
        [string] $DefaultPath
    )
    if (-not $Path) {
        $todaysDate = (Get-Date).ToString("yyyyMMdd")
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
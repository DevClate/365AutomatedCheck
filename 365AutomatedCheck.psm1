$RootPath = Split-Path $MyInvocation.MyCommand.Path

$FunctionFiles = $("$PSScriptRoot\functions\public\","$PSScriptRoot\functions\private\") | Get-Childitem -file -Recurse -Include "*.ps1" -ErrorAction SilentlyContinue

$functions = @()
foreach($FunctionFile in $FunctionFiles){
    try {
        . $FunctionFile.FullName
        $functions += Get-Command -Module $MyInvocation.MyCommand.ModuleName -CommandType Function
    }
    catch {
        Write-Error -Message "Failed to import function: '$($FunctionFile.FullName)': $_"
    }
}

Export-ModuleMember -Function ($functions.Name)
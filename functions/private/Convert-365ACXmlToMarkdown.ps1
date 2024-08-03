<#
.SYNOPSIS
Converts an XML file to a Markdown file using a specified XSLT transformation.
.PARAMETER XmlPath
The path to the input XML file.
.PARAMETER XsltPath
The path to the XSLT file used for the transformation.
.PARAMETER OutputMarkdownPath
The path to the output Markdown file that will be generated.
.PARAMETER IncludeButtons
Specifies whether to include buttons in the output Markdown file.
.EXAMPLE
Convert-365ACXmlToMarkdown -XmlPath "C:\Path\To\Input.xml" -XsltPath "C:\Path\To\Transform.xslt" -OutputMarkdownPath "C:\Path\To\Output.md" -IncludeButtons $false
This example demonstrates how to use the Convert-365ACXmlToMarkdown function to convert an XML file to a Markdown file using a specified XSLT transformation, excluding buttons.
.NOTES
This function requires the System.Xml.Xsl.XslCompiledTransform and System.Xml.XmlDocument classes from the .NET Framework.
.LINK
https://docs.microsoft.com/en-us/dotnet/api/system.xml.xsl.xslcompiledtransform?view=net-5.0
#>
function Convert-365ACXmlToMarkdown {
    param (
        [string] $XmlPath,
        [string] $OutputMarkdownPath,
        [string] $XsltPath,
        [bool] $IncludeButtons = $true
    )

    if (-not (Test-Path -Path $XmlPath)) {
        Write-PSFMessage -Level Error -Message "XML file not found: $XmlPath"
        return
    }

    Write-PSFMessage -Level Host -Message "XML file found: $XmlPath"

    if ([string]::IsNullOrEmpty($XsltPath)) {
        Write-PSFMessage -Level Error -Message "XSLT file path is null or empty."
        return
    }

    # Ensure XsltPath is an absolute path
    $absoluteXsltPath = [System.IO.Path]::GetFullPath($XsltPath)

    if (-not (Test-Path -Path $absoluteXsltPath)) {
        Write-PSFMessage -Level Error -Message "XSLT file not found: $absoluteXsltPath"
        return
    }

    Write-PSFMessage -Level Host -Message "XSLT file found: $absoluteXsltPath"

    # Create XsltArgumentList and add the parameter
    $xsltArgs = New-Object System.Xml.Xsl.XsltArgumentList
    $includeButtonsValue = if ($IncludeButtons) { 'true' } else { 'false' }
    $xsltArgs.AddParam('includeButtons', '', $includeButtonsValue)

    try {
        $xslt = New-Object System.Xml.Xsl.XslCompiledTransform
        $xslt.Load($absoluteXsltPath)

        $writer = [System.Xml.XmlWriter]::Create($OutputMarkdownPath)
        $xslt.Transform($XmlPath, $xsltArgs, $writer)
        $writer.Close()
    }
    catch {
        Write-PSFMessage -Level Error -Message "An error occurred: $_"
    }
}
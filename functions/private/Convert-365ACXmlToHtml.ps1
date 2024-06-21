<#
.SYNOPSIS
Converts an XML file to an HTML file using an XSLT transformation.

.DESCRIPTION
The Convert-365ACXmlToHtml function takes an XML file, an XSLT file, and an output HTML file path as input parameters. It performs an XSLT transformation on the XML file using the specified XSLT file and generates an HTML report at the specified output path.

.PARAMETER XmlPath
The path to the XML file that needs to be transformed.

.PARAMETER XsltPath
The path to the XSLT file that defines the transformation rules.

.PARAMETER OutputHtmlPath
The path to the output HTML file that will be generated.

.EXAMPLE
Convert-365ACXmlToHtml -XmlPath "C:\Path\To\Input.xml" -XsltPath "C:\Path\To\Transform.xslt" -OutputHtmlPath "C:\Path\To\Output.html"
This example demonstrates how to use the Convert-365ACXmlToHtml function to convert an XML file to an HTML file using a specified XSLT transformation.

.NOTES
This function requires the System.Xml.Xsl.XslCompiledTransform and System.Xml.XmlDocument classes from the .NET Framework.

.LINK
https://docs.microsoft.com/en-us/dotnet/api/system.xml.xsl.xslcompiledtransform?view=net-5.0
#>
function Convert-365ACXmlToHtml {
    param (
        [string] $XmlPath,
        [string] $XsltPath,
        [string] $OutputHtmlPath
    )

    if (Test-Path -Path $XmlPath) {
        Write-Host "XML file generated: $XmlPath"

        $xslt = New-Object System.Xml.Xsl.XslCompiledTransform
        $xslt.Load($XsltPath)

        $xml = New-Object System.Xml.XmlDocument
        $xml.Load($XmlPath)

        $writer = New-Object System.IO.StringWriter
        $xslt.Transform($xml, $null, $writer)

        $writer.ToString() | Out-File $OutputHtmlPath
        Write-Host "HTML report generated: $OutputHtmlPath"
    } else {
        Write-Error "The XML file '$XmlPath' does not exist."
    }
}
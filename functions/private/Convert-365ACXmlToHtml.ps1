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
# Transform extracted Office metadata into a line-by-line CSV using XSLT.

try {
    $resultFolder = $env:SELECTED_RESULT_FOLDER

    if (-not $resultFolder -or
        -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
        Write-Host "No valid result folder selected." -ForegroundColor Red
        return
    }

    # Step 1: Determine the document type
    if (Test-Path -LiteralPath (Join-Path $resultFolder "word") -PathType Container) {
        $filetype = "docx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "xl") -PathType Container) {
        $filetype = "xlsx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "ppt") -PathType Container) {
        $filetype = "pptx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "meta.xml") -PathType Leaf) {
        $filetype = "odt"
    }
    else {
        Write-Host "No supported Office or OpenDocument file found." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "Document type: $filetype"
    Write-Host "Folder: $resultFolder"
    Write-Host "Reading metadata..." -ForegroundColor Cyan

    # Step 2: Load the metadata XML files into one XML document
    $combinedXml = New-Object System.Xml.XmlDocument
    $combinedXml.XmlResolver = $null

    $root = $combinedXml.CreateElement("Metadata")
    $combinedXml.AppendChild($root) | Out-Null

    $metadataFiles = switch ($filetype) {
        "odt" {
            @(
                @{
                    Path = "meta.xml"
                    FullPath = Join-Path $resultFolder "meta.xml"
                }
            )
        }
        default {
            @(
                @{
                    Path = "docProps/core.xml"
                    FullPath = Join-Path $resultFolder "docProps\core.xml"
                },
                @{
                    Path = "docProps/app.xml"
                    FullPath = Join-Path $resultFolder "docProps\app.xml"
                }
            )
        }
    }

    $filesFound = 0

    foreach ($metadataFile in $metadataFiles) {
        if (-not (Test-Path -LiteralPath $metadataFile.FullPath -PathType Leaf)) {
            Write-Host "Metadata file not found: $($metadataFile.FullPath)" `
                -ForegroundColor Yellow
            continue
        }

        try {
            $sourceXml = New-Object System.Xml.XmlDocument
            $sourceXml.XmlResolver = $null
            $sourceXml.Load($metadataFile.FullPath)

            $section = $combinedXml.CreateElement("Section")
            $section.SetAttribute("path", $metadataFile.Path)
            $root.AppendChild($section) | Out-Null

            foreach ($node in $sourceXml.DocumentElement.ChildNodes) {
                if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    $importedNode = $combinedXml.ImportNode($node, $true)
                    $section.AppendChild($importedNode) | Out-Null
                }
            }

            $filesFound++
        }
        catch {
            Write-Host "Could not read $($metadataFile.Path): $_" `
                -ForegroundColor Yellow
        }
    }

    if ($filesFound -eq 0) {
        Write-Host "No metadata files could be read." -ForegroundColor Red
        return
    }

    # Step 3: Define the XSLT transformation.
    # Each leaf metadata element becomes one CSV row.
    # All fields are quoted, and internal quotes are escaped.


    $xsltText = @'
<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <xsl:output method="text" encoding="utf-8"/>

    <xsl:template name="csv-escape">
        <xsl:param name="text"/>
        <xsl:choose>
            <xsl:when test="contains($text, '&quot;')">
                <xsl:value-of select="substring-before($text, '&quot;')"/>
                <xsl:text>""</xsl:text>
                <xsl:call-template name="csv-escape">
                    <xsl:with-param name="text"
                        select="substring-after($text, '&quot;')"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$text"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="csv-field">
        <xsl:param name="text"/>
        <xsl:text>"</xsl:text>
        <xsl:call-template name="csv-escape">
            <xsl:with-param name="text" select="$text"/>
        </xsl:call-template>
        <xsl:text>"</xsl:text>
    </xsl:template>

    <xsl:template name="write-row">
        <xsl:param name="path"/>
        <xsl:param name="property"/>
        <xsl:param name="value"/>

        <xsl:call-template name="csv-field">
            <xsl:with-param name="text" select="$path"/>
        </xsl:call-template>
        <xsl:text>,</xsl:text>
        <xsl:call-template name="csv-field">
            <xsl:with-param name="text" select="$property"/>
        </xsl:call-template>
        <xsl:text>,</xsl:text>
        <xsl:call-template name="csv-field">
            <xsl:with-param name="text" select="$value"/>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
    </xsl:template>

    <xsl:template match="/">
        <xsl:text>"Path","Property","Value"&#10;</xsl:text>

        <!-- Extract element values, including nested metadata elements -->
        <xsl:for-each select="/Metadata/Section//*[not(*)]">
            <xsl:if test="normalize-space(.) != ''">
                <xsl:call-template name="write-row">
                    <xsl:with-param name="path" select="../../@path"/>
                    <xsl:with-param name="property" select="local-name()"/>
                    <xsl:with-param name="value" select="normalize-space(.)"/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>

        <!-- Extract metadata attributes, such as ODT document statistics -->
        <xsl:for-each select="/Metadata/Section//@*">
            <xsl:if test="normalize-space(.) != ''">
                <xsl:call-template name="write-row">
                    <xsl:with-param name="path" select="../../@path"/>
                    <xsl:with-param name="property"
                        select="concat(local-name(..), '@', local-name())"/>
                    <xsl:with-param name="value" select="normalize-space(.)"/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>

</xsl:stylesheet>
'@

    # Step 4: Execute the XSLT transformation
    $stylesheet = New-Object System.Xml.Xsl.XslCompiledTransform

    $stylesheetSettings = New-Object System.Xml.XmlReaderSettings
    $stylesheetSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $stylesheetSettings.XmlResolver = $null

    $stringReader = New-Object System.IO.StringReader($xsltText)
    $xmlReader = [System.Xml.XmlReader]::Create(
        $stringReader,
        $stylesheetSettings
    )

    try {
        $stylesheet.Load($xmlReader)
    }
    finally {
        $xmlReader.Dispose()
        $stringReader.Dispose()
    }

    $outputPath = Join-Path (Split-Path $resultFolder -Parent) "output.csv"

    $outputWriter = New-Object System.IO.StringWriter
    try {
        $stylesheet.Transform($combinedXml, $null, $outputWriter)
        $csvContent = $outputWriter.ToString()
    }
    finally {
        $outputWriter.Dispose()
    }

    # Step 5: Save the CSV using UTF-8 encoding
    $utf8 = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($outputPath, $csvContent, $utf8)

    $rowCount = @(
        $csvContent -split "\r?\n" |
        Select-Object -Skip 1 |
        Where-Object { $_.Trim() }
    ).Count

    Write-Host ""
    Write-Host "CSV created successfully!" -ForegroundColor Green
    Write-Host "File: $outputPath"
    Write-Host "Metadata rows: $rowCount"
}
catch {
    Write-Host ""
    Write-Host "METADATA CONVERSION FAILED" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Location: $($_.InvocationInfo.ScriptName)" -ForegroundColor Yellow
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Full error details:" -ForegroundColor Cyan
    Write-Host ($_ | Out-String)
}
finally {
    Write-Host ""
    Read-Host "Press Enter to return to the menu"
}

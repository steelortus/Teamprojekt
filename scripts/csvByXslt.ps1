# Transform extracted Office metadata into output.csv using XSLT.
# Produces the same CSV as printCsv.ps1.

try {
    $resultFolder = $env:SELECTED_RESULT_FOLDER

    if (-not $resultFolder -or
        -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
        Write-Host "No valid result folder selected." -ForegroundColor Red
        return
    }

    # Step 1: Determine the document type (same logic as printCsv.ps1)
    if (Test-Path -LiteralPath (Join-Path $resultFolder "word") -PathType Container) {
        $filetype = "docx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "xl") -PathType Container) {
        $filetype = "xlsx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "ppt") -PathType Container) {
        $filetype = "pptx"
    }
    elseif (Test-Path -LiteralPath (Join-Path $resultFolder "mimetype") -PathType Leaf) {
        # OpenDocument file
        $mimetype = (Get-Content -LiteralPath (Join-Path $resultFolder "mimetype") -Raw).Trim()
        $type = $mimetype.Split(".")[-1]

        if ($type -eq "text") {
            $filetype = "odt"
        }
        elseif ($type -eq "spreadsheet") {
            $filetype = "ods"
        }
        else {
            Write-Host "No supported file format found." -ForegroundColor Red
            return
        }
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
    $root.SetAttribute("filetype", $filetype)
    $combinedXml.AppendChild($root) | Out-Null

    $metadataFiles = switch ($filetype) {
        { $_ -in @("odt", "ods") } {
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

            foreach ($node in $sourceXml.DocumentElement.ChildNodes) {
                if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    $importedNode = $combinedXml.ImportNode($node, $true)
                    $root.AppendChild($importedNode) | Out-Null
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
    # It produces the same single-row CSV as printCsv.ps1: the columns depend
    # on the document type, all fields are quoted, internal quotes are escaped.

    $xsltText = @'
<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <xsl:output method="text" encoding="utf-8"/>

    <xsl:variable name="type" select="/Metadata/@filetype"/>

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

    <!-- One quoted field followed by a comma, or by a newline if last -->
    <xsl:template name="f">
        <xsl:param name="v" select="''"/>
        <xsl:param name="last" select="false()"/>
        <xsl:text>"</xsl:text>
        <xsl:call-template name="csv-escape">
            <xsl:with-param name="text" select="string($v)"/>
        </xsl:call-template>
        <xsl:text>"</xsl:text>
        <xsl:choose>
            <xsl:when test="$last"><xsl:text>&#10;</xsl:text></xsl:when>
            <xsl:otherwise><xsl:text>,</xsl:text></xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="/">
        <xsl:choose>

            <!-- docx: all columns; xlsx/pptx: same columns, text statistics empty -->
            <xsl:when test="$type = 'docx' or $type = 'xlsx' or $type = 'pptx'">
                <xsl:text>"Filetype","Author","Language","Revision","Application","Version","Created","Modified","TotalTime","Pages","Words","Characters","CharactersWithSpaces","Paragraphs"&#10;</xsl:text>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$type"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='creator']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='language']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='revision']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='Application']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='AppVersion']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='created']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='modified']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='TotalTime']"/></xsl:call-template>
                <xsl:if test="$type = 'docx'">
                    <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='Pages']"/></xsl:call-template>
                    <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='Words']"/></xsl:call-template>
                    <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='Characters']"/></xsl:call-template>
                    <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='CharactersWithSpaces']"/></xsl:call-template>
                    <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='Paragraphs']"/><xsl:with-param name="last" select="true()"/></xsl:call-template>
                </xsl:if>
                <xsl:if test="$type != 'docx'">
                    <xsl:call-template name="f"/>
                    <xsl:call-template name="f"/>
                    <xsl:call-template name="f"/>
                    <xsl:call-template name="f"/>
                    <xsl:call-template name="f"><xsl:with-param name="last" select="true()"/></xsl:call-template>
                </xsl:if>
            </xsl:when>

            <!-- odt -->
            <xsl:when test="$type = 'odt'">
                <xsl:variable name="stats" select="/Metadata/*[local-name()='document-statistic']"/>
                <xsl:text>"Filetype","Revision","Application","Created","Modified","Duration","Tables","Objects","Pages","Words","Characters","CharactersWithSpaces","Paragraphs","Images"&#10;</xsl:text>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$type"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='editing-cycles']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='generator']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='creation-date']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='date']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='editing-duration']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='table-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='object-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='page-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='word-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='non-whitespace-character-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='character-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='paragraph-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='image-count']"/><xsl:with-param name="last" select="true()"/></xsl:call-template>
            </xsl:when>

            <!-- ods -->
            <xsl:otherwise>
                <xsl:variable name="stats" select="/Metadata/*[local-name()='document-statistic']"/>
                <xsl:text>"Filetype","Revision","Application","Created","Modified","Duration","Tables","Objects","Cells"&#10;</xsl:text>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$type"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='editing-cycles']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='generator']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='creation-date']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='date']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="/Metadata/*[local-name()='editing-duration']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='table-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='object-count']"/></xsl:call-template>
                <xsl:call-template name="f"><xsl:with-param name="v" select="$stats/@*[local-name()='cell-count']"/><xsl:with-param name="last" select="true()"/></xsl:call-template>
            </xsl:otherwise>

        </xsl:choose>
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

    Write-Host ""
    Write-Host "CSV created successfully!" -ForegroundColor Green
    Write-Host "File: $outputPath"
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

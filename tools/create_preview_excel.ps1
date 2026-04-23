$ErrorActionPreference = 'Stop'

$outDir = 'C:\Projects\FillDoc\test_data'
$outPath = Join-Path $outDir 'preview_test.xlsx'
$tempRoot = Join-Path $outDir 'preview_test_xlsx_tmp'

if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $tempRoot '_rels') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $tempRoot 'docProps') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $tempRoot 'xl') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $tempRoot 'xl\_rels') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $tempRoot 'xl\worksheets') | Out-Null

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Escape-Xml([string]$Text) {
    if ($null -eq $Text) {
        return ''
    }

    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}

function Get-ColumnName([int]$Number) {
    $name = ''
    while ($Number -gt 0) {
        $Number--
        $name = [char](65 + ($Number % 26)) + $name
        $Number = [math]::Floor($Number / 26)
    }
    return $name
}

function New-InlineStringCell([int]$Column, [int]$Row, [string]$Value, [int]$StyleId = 0) {
    $cellRef = '{0}{1}' -f (Get-ColumnName $Column), $Row
    $escaped = Escape-Xml $Value
    return '<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>' -f $cellRef, $StyleId, $escaped
}

function New-RowXml([int]$RowNumber, [string[]]$Values, [int]$StyleId) {
    $cells = for ($i = 0; $i -lt $Values.Count; $i++) {
        New-InlineStringCell -Column ($i + 1) -Row $RowNumber -Value $Values[$i] -StyleId $StyleId
    }
    return '<row r="{0}">{1}</row>' -f $RowNumber, ($cells -join '')
}

function New-SheetXml([string[]]$Headers, [object[]]$Rows) {
    $rowXml = New-Object System.Collections.Generic.List[string]
    $rowXml.Add((New-RowXml -RowNumber 1 -Values $Headers -StyleId 1))

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $rowXml.Add((New-RowXml -RowNumber ($i + 2) -Values $Rows[$i] -StyleId 0))
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:G$($Rows.Count + 1)"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="4" width="20" customWidth="1"/>
    <col min="5" max="5" width="14" customWidth="1"/>
    <col min="6" max="6" width="14" customWidth="1"/>
    <col min="7" max="7" width="16" customWidth="1"/>
  </cols>
  <sheetData>
    $($rowXml -join "`n    ")
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"@
}

$headers = @(
    'Project Name',
    'Case Number',
    'Creditor',
    'Debtor',
    'Debt Amount',
    'Status',
    'filldoc_id'
)

$currentRows = @(
    @('Test Project 1', 'A40-12345/2026', 'Alpha LLC', 'Petrov IE', '150000', 'New', 'demo-001'),
    @('Test Project 2', 'A40-54321/2026', 'Development Bank', 'Beta LLC', '275500', 'In Progress', 'demo-002'),
    @('Test Project 3', 'A40-77777/2026', 'Gamma LLC', 'Vector JSC', '99999', 'Review', 'demo-003')
)

$archiveRows = @(
    @('Archived Project', 'A40-00001/2025', 'Archive LLC', 'Old Client LLC', '50000', 'Archived', 'arch-001')
)

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@

$rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$coreProps = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-04-23T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-23T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"@

$appProps = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>
"@

$workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Current" sheetId="1" r:id="rId1"/>
    <sheet name="Archive" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"@

$workbookRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"@

Write-Utf8NoBom -Path (Join-Path $tempRoot '[Content_Types].xml') -Content $contentTypes
Write-Utf8NoBom -Path (Join-Path $tempRoot '_rels\.rels') -Content $rootRels
Write-Utf8NoBom -Path (Join-Path $tempRoot 'docProps\core.xml') -Content $coreProps
Write-Utf8NoBom -Path (Join-Path $tempRoot 'docProps\app.xml') -Content $appProps
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\workbook.xml') -Content $workbook
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRels
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\styles.xml') -Content $styles
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\worksheets\sheet1.xml') -Content (New-SheetXml -Headers $headers -Rows $currentRows)
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\worksheets\sheet2.xml') -Content (New-SheetXml -Headers $headers -Rows $archiveRows)

if (Test-Path -LiteralPath $outPath) {
    Remove-Item -LiteralPath $outPath -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outPath)
Remove-Item -LiteralPath $tempRoot -Recurse -Force

Write-Output $outPath

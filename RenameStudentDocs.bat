@echo off
setlocal
chcp 65001 >nul 2>&1
title Student Document Renamer
color 0B

echo ============================================
echo     STUDENT DOCUMENT RENAMER
echo ============================================
echo.
echo  Folder: %CD%
echo.
echo  This tool will:
echo   1. Read each PDF, DOCX, and TXT file
echo   2. Find the Student Name and Reg Number
echo   3. Rename the file as: Name - RegNumber.ext
echo.
echo  Expected labels in documents:
echo   Name / Student Name / Student's Name
echo   Reg / Reg No / Registration Number / Roll No
echo.
echo ============================================
echo.

set /p "confirm=  Proceed? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo  Cancelled.
    timeout /t 2 >nul
    goto :eof
)

echo.
echo  Processing files...
echo.

REM --- Set paths as environment variables so PowerShell can read them safely ---
set "_PS=%TEMP%\_StuRename_%RANDOM%.ps1"
set "_BAT=%~f0"
set "_WD=%CD%"
set "_BN=%~nx0"

REM --- Extract the embedded PowerShell script (after the marker) to a temp file ---
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=[IO.File]::ReadAllText($env:_BAT);$m='#__PS'+'SCRIPT__#';$i=$c.LastIndexOf($m);if($i-ge0){$s=$c.Substring($i+$m.Length);$s=$s.Replace([char]0x2018,[char]39).Replace([char]0x2019,[char]39).Replace([char]0x201C,[char]34).Replace([char]0x201D,[char]34);[IO.File]::WriteAllText($env:_PS,$s,[System.Text.Encoding]::UTF8)}"

REM --- Run the extracted script ---
powershell -NoProfile -ExecutionPolicy Bypass -File "%_PS%"

REM --- Cleanup ---
del "%_PS%" >nul 2>&1

echo.
pause
exit /b
#__PSSCRIPT__#

Set-Location $env:_WD
$batName = $env:_BN

$totalProcessed = 0
$totalRenamed = 0
$totalSkipped = 0

# =============================================
#  TEXT EXTRACTION FUNCTIONS
# =============================================

function Extract-TextFromTxt($path) {
    $text = $null
    # Try UTF8 first, then default encoding
    $text = Get-Content $path -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $text) {
        $text = Get-Content $path -Raw -ErrorAction SilentlyContinue
    }
    return $text
}

function Extract-TextFromDocx($path) {
    # Method 1: Word COM (best quality - handles tables, headers, footers)
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($path, $false, $true)
        $text = $doc.Content.Text
        # Also grab headers/footers
        foreach ($section in $doc.Sections) {
            foreach ($header in $section.Headers) {
                $text += " " + $header.Range.Text
            }
            foreach ($footer in $section.Footers) {
                $text += " " + $footer.Range.Text
            }
        }
        $doc.Close($false)
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        if ($text -and $text.Trim().Length -gt 0) { return $text }
    } catch {}

    # Method 2: ZIP/XML fallback (no Word installed)
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        $allText = ""
        # Read main document, headers, footers
        $xmlFiles = $zip.Entries | Where-Object {
            $_.FullName -match ‘^word/(document|header\d*|footer\d*)\.xml$’
        }
        foreach ($entry in $xmlFiles) {
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $xml = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            $xml = $xml -replace ‘</w:p>’, "`n"
            $xml = $xml -replace ‘</w:tr>’, "`n"
            $xml = $xml -replace ‘<w:tab/>’, " "
            $xml = $xml -replace ‘<[^>]+>’, ‘’
            $allText += $xml + "`n"
        }
        $zip.Dispose()
        $allText = $allText -replace ‘[ \t]+’, ‘ ‘
        $allText = $allText -replace ‘(\r?\n){3,}’, "`n`n"
        if ($allText.Trim().Length -gt 0) { return $allText.Trim() }
    } catch {}

    return $null
}

function Extract-TextFromPdf($path) {
    # Method 1: Word COM (opens PDF as Word doc - most reliable)
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($path, $false, $true)
        $text = $doc.Content.Text
        $doc.Close($false)
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        if ($text -and $text.Trim().Length -gt 10) { return $text }
    } catch {}

    # Method 2: Raw PDF text extraction (BT/ET blocks + Tj/TJ operators)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
        $texts = @()

        $btMatches = [regex]::Matches($raw, ‘BT\s*(.*?)\s*ET’, ‘Singleline’)
        foreach ($bm in $btMatches) {
            $block = $bm.Groups[1].Value
            # Match text in parentheses: Tj operator
            $tjMatches = [regex]::Matches($block, ‘\(([^)]*)\)’)
            foreach ($tj in $tjMatches) {
                $t = $tj.Groups[1].Value
                $t = $t -replace ‘\\n’, "`n"
                $t = $t -replace ‘\\r’, ""
                $t = $t -replace ‘\\t’, " "
                $t = $t -replace ‘\\(.)’, ‘$1’
                if ($t.Trim().Length -gt 0) { $texts += $t.Trim() }
            }
            # Match hex strings: <hex> Tj
            $hexMatches = [regex]::Matches($block, ‘<([0-9A-Fa-f]+)>’)
            foreach ($hm in $hexMatches) {
                $hex = $hm.Groups[1].Value
                if ($hex.Length -ge 4 -and $hex.Length % 2 -eq 0) {
                    $decoded = ""
                    for ($i = 0; $i -lt $hex.Length; $i += 2) {
                        $byte = [Convert]::ToInt32($hex.Substring($i, 2), 16)
                        if ($byte -ge 32 -and $byte -le 126) {
                            $decoded += [char]$byte
                        }
                    }
                    if ($decoded.Trim().Length -gt 0) { $texts += $decoded.Trim() }
                }
            }
        }
        if ($texts.Count -gt 0) {
            $result = $texts -join ‘ ‘
            if ($result.Trim().Length -gt 10) { return $result }
        }

        # Method 3: Brute force - extract all readable text from raw PDF
        $readable = $raw -replace ‘[^\x20-\x7E\r\n]’, ‘ ‘
        $readable = $readable -replace ‘\s+’, ‘ ‘
        if ($readable.Trim().Length -gt 50) { return $readable.Trim() }
    } catch {}

    return $null
}

# =============================================
#  STUDENT INFO EXTRACTION (FULL DOC SEARCH)
# =============================================

function Find-StudentInfo($text) {
    if (-not $text) { return $null }
    $name = $null
    $reg = $null

    # --- FIND NAME ---
    $apos = "[‘`’" + [char]0x2018 + [char]0x2019 + "]"
    $namePatterns = @(
        "(?i)(?:student${apos}?s?\s*name|full\s*name|name\s*of\s*(?:the\s*)?(?:student|candidate)|candidate${apos}?s?\s*name|submitted\s*by|prepared\s*by|author)\s*[:=\-]\s*([A-Za-z][A-Za-z\s.]{2,50})",
        "(?i)\bname\s*[:=\-]\s*([A-Za-z][A-Za-z\s.]{2,50})",
        "(?i)(?:student${apos}?s?\s*name|full\s*name|name\s*of\s*(?:the\s*)?(?:student|candidate)|candidate${apos}?s?\s*name)\s+([A-Z][A-Za-z\s.]{2,50})",
        "(?i)\bname\s+([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){1,4})"
    )
    foreach ($p in $namePatterns) {
        $m = [regex]::Match($text, $p)
        if ($m.Success) {
            $name = $m.Groups[1].Value.Trim()
            $name = $name -replace "\s+", " "
            $name = ($name -replace "(?i)\s*(reg|roll|registration|class|section|semester|date|department|dept|faculty|subject|course|program|matric|enrollment|student\s*id|submitted|assignment|project|report|lab|marks|grade|gpa|cgpa|batch|session|year|contact|phone|email|address|father|cnic|id\s*no|signature).*$", "").Trim()
            $name = $name -replace "[.:,;]+$", ""
            if ($name.Length -ge 3 -and $name -match "[A-Za-z]") { break }
            else { $name = $null }
        }
    }

    # --- FIND REG NUMBER ---
    $regPatterns = @(
        "(?i)(?:roll|reg(?:istration)?)\s*(?:no\.?|number|#)?\s*[.:=\-]\s*([A-Za-z0-9][\w\-/.]{2,30})",
        "(?i)(?:student|enrollment|matric)\s*(?:no\.?|number|id|#)?\s*[.:=\-]\s*([A-Za-z0-9][\w\-/.]{2,30})",
        "(?i)(?:cms|sap|erp|id)\s*(?:no\.?|number|id|#)?\s*[.:=\-]\s*([A-Za-z0-9][\w\-/.]{2,30})",
        "(?i)(?:roll|reg(?:istration)?)\s*(?:no\.?|number|#)?\s+([A-Z0-9][\w\-/.]{2,30})",
        "(?i)(?:student|enrollment|matric)\s*(?:no\.?|number|id|#)?\s+([A-Z0-9][\w\-/.]{2,30})"
    )
    foreach ($p in $regPatterns) {
        $m = [regex]::Match($text, $p)
        if ($m.Success) {
            $reg = $m.Groups[1].Value.Trim()
            $reg = $reg -replace "[.:,;\s]+$", ""
            if ($reg.Length -ge 3) { break }
            else { $reg = $null }
        }
    }

    if ($name -or $reg) {
        return @{ Name = $name; Reg = $reg }
    }
    return $null
}

# =============================================
#  FILENAME SANITIZER
# =============================================

function Sanitize-FileName($s) {
    $s = $s -replace "[/\\]", "-"
    $s = $s -replace "[<>:""|?*]", ""
    $s = $s -replace "\s+", " "
    return $s.Trim()
}

# =============================================
#  MAIN
# =============================================

$files = Get-ChildItem -File | Where-Object {
    $_.Extension -match ‘^\.(txt|docx|pdf)$’ -and $_.Name -ne $batName
}

if (-not $files -or @($files).Count -eq 0) {
    Write-Host ‘  No PDF, DOCX, or TXT files found in this folder.’ -ForegroundColor Yellow
    exit
}

$fileList = @($files)
Write-Host ("  Found {0} file(s) to process..." -f $fileList.Count)
Write-Host

foreach ($file in $fileList) {
    $totalProcessed++
    $text = $null
    $ext = $file.Extension.ToLower()

    Write-Host ("  Reading: {0} ..." -f $file.Name) -ForegroundColor Gray

    switch ($ext) {
        ‘.txt’  { $text = Extract-TextFromTxt $file.FullName }
        ‘.docx’ { $text = Extract-TextFromDocx $file.FullName }
        ‘.pdf’  { $text = Extract-TextFromPdf $file.FullName }
    }

    if (-not $text) {
        Write-Host ("  SKIP: {0}" -f $file.Name) -ForegroundColor Yellow
        Write-Host "        (could not read any text from this file)" -ForegroundColor DarkYellow
        $totalSkipped++
        continue
    }

    $info = Find-StudentInfo $text

    if (-not $info) {
        Write-Host ("  SKIP: {0}" -f $file.Name) -ForegroundColor Yellow
        Write-Host "        (could not find name or reg in the text below)" -ForegroundColor DarkYellow
        # Show first 300 chars of extracted text so user can see what was read
        $preview = $text.Substring(0, [Math]::Min(300, $text.Length)) -replace ‘\r?\n’, ‘ ‘
        Write-Host ("        TEXT: {0}..." -f $preview) -ForegroundColor DarkGray
        $totalSkipped++
        continue
    }

    # Build new filename
    if ($info.Name -and $info.Reg) {
        $baseName = "{0} - {1}" -f $info.Name, $info.Reg
    } elseif ($info.Name) {
        $baseName = $info.Name
    } else {
        $baseName = $info.Reg
    }
    $newName = (Sanitize-FileName $baseName) + $ext
    $newPath = Join-Path $file.DirectoryName $newName

    # Handle duplicates
    if (Test-Path $newPath) {
        $counter = 2
        do {
            $newName = (Sanitize-FileName ("{0} ({1})" -f $baseName, $counter)) + $ext
            $newPath = Join-Path $file.DirectoryName $newName
            $counter++
        } while (Test-Path $newPath)
    }

    try {
        Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
        Write-Host ("  OK:   {0}" -f $file.Name) -ForegroundColor Green
        Write-Host ("    --> {0}" -f $newName) -ForegroundColor Cyan
        $totalRenamed++
    } catch {
        Write-Host ("  FAIL: {0}" -f $file.Name) -ForegroundColor Red
        Write-Host ("        {0}" -f $_.Exception.Message) -ForegroundColor DarkRed
        $totalSkipped++
    }
}

Write-Host
Write-Host ‘  ============================================’
Write-Host ("  Total files:    {0}" -f $totalProcessed)
Write-Host ("  Renamed:        {0}" -f $totalRenamed) -ForegroundColor Green
Write-Host ("  Skipped:        {0}" -f $totalSkipped) -ForegroundColor Yellow
Write-Host ‘  ============================================’

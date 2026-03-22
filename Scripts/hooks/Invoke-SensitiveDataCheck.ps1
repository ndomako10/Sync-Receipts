#Requires -Version 5.0
<#
.SYNOPSIS
    Checks file content for sensitive data patterns and returns findings.

.DESCRIPTION
    Pure helper called by Invoke-PreCommitCheck.ps1. Accepts injected content
    and patterns so it can be unit-tested without disk I/O.

    Also contains Get-AccountsLast4, a pure XML parser that extracts Last4
    account identifiers from an Accounts.xlsx worksheet for use as dynamic
    patterns.

.NOTES
    Lines ending with '# nocheck' are skipped for all patterns.
    A match is suppressed if the line contains any entry in the pattern's
    allowlist array.
#>

function Invoke-SensitiveDataCheck {
    <#
    .SYNOPSIS
        Scans staged file content for sensitive data patterns.

    .PARAMETER Content
        The full text content of the staged file (from git show :file).

    .PARAMETER FileName
        The file path as returned by git (e.g. 'Docs/ADRs/ADR-009.md').
        Used to filter patterns by their fileTypes list.

    .PARAMETER Patterns
        Array of pattern objects loaded from SensitivePatterns.json (or
        built dynamically). Each object must have: name, regex, fileTypes,
        message, allowlist.

    .OUTPUTS
        [PSCustomObject[]] with fields PatternName, LineNumber, Line, Message.
        Returns an empty array when no findings are detected.

    .EXAMPLE
        $findings = Invoke-SensitiveDataCheck -Content $staged -FileName $path -Patterns $patterns
        if ($findings) { $findings | ForEach-Object { Write-Host $_.Message } }
    #>
    [CmdletBinding()]
    param(
        [string]$Content,
        [string]$FileName,
        [array]$Patterns = @()
    )

    $findings = @()
    if (-not $Content -or -not $Patterns -or $Patterns.Count -eq 0) { return $findings }

    $ext   = [System.IO.Path]::GetExtension($FileName).ToLower()
    $lines = $Content -split "`n"

    foreach ($pattern in $Patterns) {
        # Filter by fileTypes if the pattern specifies them
        $fileTypes = @($pattern.fileTypes)
        if ($fileTypes.Count -gt 0 -and $ext -notin $fileTypes) { continue }

        $allowlist = @($pattern.allowlist)

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # Skip lines with the suppression token
            if ($line -match '#\s*nocheck\s*$') { continue }

            # Skip lines that contain an allowlist entry
            $suppressed = $false
            foreach ($entry in $allowlist) {
                if ($line -match [regex]::Escape($entry)) {
                    $suppressed = $true
                    break
                }
            }
            if ($suppressed) { continue }

            if ($line -match $pattern.regex) {
                $findings += [PSCustomObject]@{
                    PatternName = $pattern.name
                    LineNumber  = $i + 1
                    Line        = $line
                    Message     = $pattern.message
                }
            }
        }
    }

    return $findings
}

function Get-AccountsLast4 {
    <#
    .SYNOPSIS
        Extracts Last4 account identifiers from Accounts.xlsx XML content.

    .DESCRIPTION
        Pure function -- no disk I/O, no COM. The caller opens Accounts.xlsx
        as a ZIP archive, reads xl/sharedStrings.xml and
        xl/worksheets/sheet1.xml, then passes both as strings to this function.

        Row 1 (the header row) is skipped. Only values that are exactly 4
        digits are returned; other column A values are ignored.

    .PARAMETER SharedStringsXml
        Raw XML content of xl/sharedStrings.xml from the xlsx archive.

    .PARAMETER WorksheetXml
        Raw XML content of xl/worksheets/sheet1.xml from the xlsx archive.

    .OUTPUTS
        [string[]] of 4-digit Last4 values. Returns an empty array on parse
        failure or when no qualifying values are found.

    .EXAMPLE
        $zip = [System.IO.Compression.ZipFile]::OpenRead($accountsPath)
        $ssXml = [System.IO.StreamReader]::new($zip.GetEntry('xl/sharedStrings.xml').Open()).ReadToEnd()
        $wsXml = [System.IO.StreamReader]::new($zip.GetEntry('xl/worksheets/sheet1.xml').Open()).ReadToEnd()
        $zip.Dispose()
        $last4s = Get-AccountsLast4 -SharedStringsXml $ssXml -WorksheetXml $wsXml
    #>
    [CmdletBinding()]
    param(
        [string]$SharedStringsXml,
        [string]$WorksheetXml
    )

    $results = @()
    if (-not $SharedStringsXml -or -not $WorksheetXml) { return $results }

    try {
        # Build shared strings lookup array (index -> value)
        $ssDoc = [xml]$SharedStringsXml
        $sharedStrings = @($ssDoc.sst.si | ForEach-Object {
            # Shared string text may be a plain <t> or composed of <r><t> runs
            if ($null -ne $_.t) {
                [string]$_.t
            } elseif ($null -ne $_.r) {
                ($_.r | ForEach-Object { if ($null -ne $_.t) { [string]$_.t } }) -join ''
            } else {
                ''
            }
        })

        # Parse worksheet and collect column A values from data rows (row > 1)
        $wsDoc = [xml]$WorksheetXml
        $rows  = @($wsDoc.worksheet.sheetData.row)

        foreach ($row in $rows) {
            # Skip header row
            if ([int]$row.r -eq 1) { continue }

            # Find the column A cell (r attribute matches A followed by digits)
            $cellA = @($row.c) | Where-Object { $_.r -match '^A\d+$' } | Select-Object -First 1
            if ($null -eq $cellA) { continue }

            $rawValue = [string]$cellA.v
            if (-not $rawValue) { continue }

            # Resolve shared string reference (t="s") or use the literal value
            $value = if ($cellA.t -eq 's') {
                $idx = [int]$rawValue
                if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) { $sharedStrings[$idx] } else { '' }
            } else {
                $rawValue
            }

            # Collect only 4-digit identifiers
            if ($value -match '^\d{4}$') {
                $results += $value
            }
        }
    } catch {
        # Return empty array on any parse failure; caller may log the exception
    }

    return $results
}

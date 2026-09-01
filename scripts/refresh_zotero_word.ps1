param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputDocx,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [Parameter()][ValidateRange(5, 120)][int]$WaitSeconds = 45
)

$ErrorActionPreference = 'Stop'

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Write-NewUtf8 {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($Value)
    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::None
    )
    try {
        $stream.Write($bytes, 0, $bytes.Length)
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ZoteroFieldInventory {
    param([Parameter(Mandatory = $true)]$Document)

    $counts = [ordered]@{
        total = 0
        zotero_citations = 0
        zotero_bibliographies = 0
        endnote_fields = 0
        other = 0
    }

    for ($storyType = 1; $storyType -le 17; $storyType++) {
        try {
            $range = $Document.StoryRanges.Item($storyType)
        }
        catch {
            continue
        }

        $rangeCount = 0
        while ($null -ne $range) {
            $rangeCount++
            if ($rangeCount -gt 512) {
                throw "Word story range traversal exceeded the safety limit for story type $storyType."
            }

            foreach ($field in @($range.Fields)) {
                $counts.total++
                try {
                    $code = ([string]$field.Code.Text).Trim()
                }
                catch {
                    $code = ''
                }

                if ($code.StartsWith(
                    'ADDIN ZOTERO_ITEM CSL_CITATION',
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    $counts.zotero_citations++
                }
                elseif ($code.StartsWith(
                    'ADDIN ZOTERO_BIBL',
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    $counts.zotero_bibliographies++
                }
                elseif ($code.StartsWith(
                    'ADDIN EN.',
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    $counts.endnote_fields++
                }
                else {
                    $counts.other++
                }
            }
            $range = $range.NextStoryRange
        }
    }

    return [pscustomobject]$counts
}

function Get-UnexpectedFieldDecreases {
    param(
        [Parameter(Mandatory = $true)]$Before,
        [Parameter(Mandatory = $true)]$After,
        [Parameter(Mandatory = $true)][string]$Stage
    )

    $decreases = @()
    foreach ($fieldName in @('zotero_citations', 'zotero_bibliographies')) {
        $beforeCount = [int]$Before.$fieldName
        $afterCount = [int]$After.$fieldName
        if ($afterCount -lt $beforeCount) {
            $decreases += [pscustomobject]@{
                stage = $Stage
                field = $fieldName
                before = $beforeCount
                after = $afterCount
            }
        }
    }
    return $decreases
}

$resolvedInput = (Resolve-Path -LiteralPath $InputDocx).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputDocx)
$resolvedReport = [System.IO.Path]::GetFullPath($ReportPath)
if ($resolvedInput -ieq $resolvedOutput) {
    throw 'InputDocx and OutputDocx must be different files.'
}
if ($resolvedReport -ieq $resolvedInput -or $resolvedReport -ieq $resolvedOutput) {
    throw 'ReportPath must be different from InputDocx and OutputDocx.'
}

if (Test-Path -LiteralPath $resolvedOutput) {
    throw 'OutputDocx already exists. Choose a new path.'
}
if (Test-Path -LiteralPath $resolvedReport) {
    throw 'ReportPath already exists. Choose a new path.'
}

$outputDirectory = Split-Path -Parent $resolvedOutput
[System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory((Split-Path -Parent $resolvedReport)) | Out-Null

$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutput)
$workingPath = Join-Path $outputDirectory (
    ".$outputBaseName.zotero-refresh.$([Guid]::NewGuid().ToString('N')).partial.docx"
)

$word = $null
$bootstrapDocument = $null
$document = $null
$verificationDocument = $null
$ownsWordProcess = $false
$failure = $null
$processingSucceeded = $false
$committedOutputSha256 = $null
$sourceHashBefore = Get-Sha256 -Path $resolvedInput

$report = [ordered]@{
    input = $resolvedInput
    output = $resolvedOutput
    staging_path = $workingPath
    quarantine_path = $null
    source_sha256_before = $sourceHashBefore
    source_sha256_after = $null
    output_sha256_before = $null
    output_sha256_after = $null
    success = $false
    deliverable = $false
    status = 'processing'
    output_committed = $false
    working_copy_created = $false
    macro = $null
    word_pid = $null
    zotero_template_loaded = $false
    templates = @()
    field_inventory_before = $null
    field_inventory_after_macro = $null
    field_inventory_after_wait = $null
    field_inventory_after_save = $null
    field_inventory_after_reopen = $null
    unexpected_field_decreases = @()
    document_saved_immediately_after_macro = $null
    observed_dirty_during_wait = $false
    document_saved_after_wait = $null
    wait_seconds = $WaitSeconds
    document_hash_changed = $null
    saved = $false
    reopen_read_only_requested = $false
    reopen_read_only = $null
    reopen_read_only_success = $false
    error_type = $null
    error_message = $null
    quarantine_error = $null
}

try {
    Copy-Item -LiteralPath $resolvedInput -Destination $workingPath
    $report.working_copy_created = $true
    $report.output_sha256_before = Get-Sha256 -Path $workingPath

    $wordPidsBefore = @(
        Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue | ForEach-Object Id
    )
    $word = New-Object -ComObject Word.Application
    if ($null -eq ('ZoteroCitationRefresh.NativeMethods' -as [type])) {
        Add-Type -Namespace ZoteroCitationRefresh -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
    }

    $bootstrapDocument = $word.Documents.Add()
    [IntPtr]$wordWindowHandle = $bootstrapDocument.ActiveWindow.Hwnd
    [uint32]$wordProcessId = 0
    [void][ZoteroCitationRefresh.NativeMethods]::GetWindowThreadProcessId(
        $wordWindowHandle,
        [ref]$wordProcessId
    )
    if ($wordWindowHandle -eq [IntPtr]::Zero -or $wordProcessId -le 0) {
        throw 'Could not map the Word COM window handle to a process ID.'
    }
    $wordProcess = Get-Process -Id $wordProcessId -ErrorAction Stop
    if ($wordProcess.ProcessName -ine 'WINWORD') {
        throw "The mapped COM process is not WINWORD: $($wordProcess.ProcessName)"
    }
    if ($wordProcessId -in $wordPidsBefore) {
        throw "Word COM reused pre-existing WINWORD PID $wordProcessId; ownership was not established."
    }
    $ownsWordProcess = $true
    $report.word_pid = [int]$wordProcessId
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $bootstrapDocument.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrapDocument)
    $bootstrapDocument = $null

    $templateNames = @()
    foreach ($template in @($word.Templates)) {
        $templateNames += [string]$template.FullName
    }
    $report.templates = $templateNames
    $report.zotero_template_loaded = [bool]($templateNames | Where-Object {
        [System.IO.Path]::GetFileName($_) -ieq 'Zotero.dotm'
    })
    if (-not $report.zotero_template_loaded) {
        throw 'Zotero.dotm is not loaded in the isolated Word instance.'
    }

    $document = $word.Documents.Open($workingPath, $false, $false, $false)
    $document.Activate()
    $report.field_inventory_before = Get-ZoteroFieldInventory -Document $document

    $macroErrors = @()
    foreach ($macro in @('ZoteroRefresh', 'Zotero.dotm!ZoteroRefresh')) {
        try {
            [void]$word.Run($macro)
            $report.macro = $macro
            break
        }
        catch {
            $macroErrors += "$macro :: $($_.Exception.Message)"
        }
    }
    if ($null -eq $report.macro) {
        throw ('No Zotero refresh macro succeeded. ' + ($macroErrors -join ' | '))
    }

    $report.field_inventory_after_macro = Get-ZoteroFieldInventory -Document $document
    $report.document_saved_immediately_after_macro = [bool]$document.Saved
    if (-not $document.Saved) {
        $report.observed_dirty_during_wait = $true
    }

    $deadline = [DateTime]::UtcNow.AddSeconds($WaitSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 500
        [void]$document.Fields.Count
        if (-not $document.Saved) {
            $report.observed_dirty_during_wait = $true
        }
    }

    $report.field_inventory_after_wait = Get-ZoteroFieldInventory -Document $document
    $report.document_saved_after_wait = [bool]$document.Saved
    $document.Save()
    $report.saved = $true
    $report.field_inventory_after_save = Get-ZoteroFieldInventory -Document $document

    $document.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null

    $report.output_sha256_after = Get-Sha256 -Path $workingPath
    $report.document_hash_changed = (
        $report.output_sha256_before -ne $report.output_sha256_after
    )
    $report.source_sha256_after = Get-Sha256 -Path $resolvedInput
    if ($report.source_sha256_after -ne $report.source_sha256_before) {
        throw 'The source DOCX hash changed during refresh.'
    }
    if (-not $report.document_hash_changed) {
        throw 'ZoteroRefresh produced an unchanged candidate; integration was not proven and it is not deliverable.'
    }

    $report.reopen_read_only_requested = $true
    $verificationDocument = $word.Documents.Open($workingPath, $false, $true, $false)
    $report.reopen_read_only = [bool]$verificationDocument.ReadOnly
    if (-not $verificationDocument.ReadOnly) {
        throw 'The refreshed candidate did not reopen read-only in Word.'
    }
    $report.field_inventory_after_reopen = Get-ZoteroFieldInventory -Document $verificationDocument
    $report.reopen_read_only_success = $true

    $decreases = @()
    $decreases += @(Get-UnexpectedFieldDecreases -Before $report.field_inventory_before -After $report.field_inventory_after_wait -Stage 'after-wait')
    $decreases += @(Get-UnexpectedFieldDecreases -Before $report.field_inventory_before -After $report.field_inventory_after_save -Stage 'after-save')
    $decreases += @(Get-UnexpectedFieldDecreases -Before $report.field_inventory_before -After $report.field_inventory_after_reopen -Stage 'after-reopen')
    $report.unexpected_field_decreases = @($decreases)
    if ($decreases.Count -gt 0) {
        $details = @($decreases | ForEach-Object {
            "$($_.stage):$($_.field) $($_.before)->$($_.after)"
        })
        throw ('Zotero citation/bibliography field count decreased unexpectedly: ' + ($details -join '; '))
    }

    $verificationDocument.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($verificationDocument)
    $verificationDocument = $null
    $processingSucceeded = $true
}
catch {
    $failure = $_.Exception
    $report.error_type = $failure.GetType().FullName
    $report.error_message = $failure.Message
}
finally {
    if ($null -ne $verificationDocument) {
        $verificationDocument.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($verificationDocument)
    }
    if ($null -ne $bootstrapDocument) {
        $bootstrapDocument.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrapDocument)
    }
    if ($null -ne $document) {
        $document.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        if ($ownsWordProcess) {
            $word.Quit()
        }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($null -eq $failure -and $processingSucceeded) {
    try {
        $report.source_sha256_after = Get-Sha256 -Path $resolvedInput
        if ($report.source_sha256_after -ne $report.source_sha256_before) {
            throw 'The source DOCX hash changed before output commit.'
        }

        Move-Item -LiteralPath $workingPath -Destination $resolvedOutput
        $committedOutputSha256 = $report.output_sha256_after
        if ((Get-Sha256 -Path $resolvedOutput) -ne $committedOutputSha256) {
            throw 'The committed output no longer matches the refreshed candidate hash.'
        }
        $report.output_committed = $true
        $report.success = $true
        $report.deliverable = $true
        $report.status = 'deliverable'
    }
    catch {
        $failure = $_.Exception
        $report.error_type = $failure.GetType().FullName
        $report.error_message = $failure.Message
    }
}

if ($null -ne $failure) {
    $report.success = $false
    $report.deliverable = $false
    $report.output_committed = $false

    if ($null -eq $report.source_sha256_after -and (Test-Path -LiteralPath $resolvedInput)) {
        $report.source_sha256_after = Get-Sha256 -Path $resolvedInput
    }
    if ((Test-Path -LiteralPath $workingPath) -and $null -eq $report.output_sha256_after) {
        $report.output_sha256_after = Get-Sha256 -Path $workingPath
        if ($null -ne $report.output_sha256_before) {
            $report.document_hash_changed = (
                $report.output_sha256_before -ne $report.output_sha256_after
            )
        }
    }

    if (Test-Path -LiteralPath $workingPath) {
        $quarantineName = (
            "$outputBaseName.quarantine.$([DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))." +
            "$([Guid]::NewGuid().ToString('N')).docx"
        )
        $quarantinePath = Join-Path $outputDirectory $quarantineName
        try {
            Move-Item -LiteralPath $workingPath -Destination $quarantinePath
            $report.quarantine_path = $quarantinePath
            $report.status = 'quarantined-not-deliverable'
        }
        catch {
            $report.quarantine_error = $_.Exception.Message
            $report.status = 'failed-partial-not-deliverable'
        }
    }
    else {
        $report.status = 'failed-before-working-copy-not-deliverable'
    }

}

try {
    Write-NewUtf8 -Path $resolvedReport -Value (
        ($report | ConvertTo-Json -Depth 8) + [Environment]::NewLine
    )
}
catch {
    $reportFailure = $_.Exception
    if ($report.output_committed -and (Test-Path -LiteralPath $resolvedOutput)) {
        try {
            $currentOutputSha256 = Get-Sha256 -Path $resolvedOutput
        }
        catch {
            throw [System.InvalidOperationException]::new(
                ('Report publication failed and committed output ownership could not be verified; ' +
                 'the script did not remove the output. Report error: ' + $reportFailure.Message +
                 ' Hash error: ' + $_.Exception.Message),
                $reportFailure
            )
        }
        if ($null -eq $committedOutputSha256 -or
            $currentOutputSha256 -ne $committedOutputSha256) {
            throw [System.InvalidOperationException]::new(
                ('Report publication failed and committed output ownership/hash mismatch; ' +
                 'the output was preserved. Expected SHA256: ' + $committedOutputSha256 +
                 '; current SHA256: ' + $currentOutputSha256 +
                 '. Report error: ' + $reportFailure.Message),
                $reportFailure
            )
        }
        Remove-Item -LiteralPath $resolvedOutput
    }
    throw $reportFailure
}

$report | ConvertTo-Json -Depth 8
if ($null -ne $failure) {
    exit 1
}


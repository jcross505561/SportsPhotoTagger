param(
    [ValidateSet('Preview', 'Apply', 'Reset')]
    [string]$Action = 'Preview',
    [switch]$RosterTest,
    [string]$Folder = '.',
    [string]$School = '',
    [string]$Sport = '',
    [string]$Season = '',
    [string]$Squad = '',
    [string]$Opponent = '',
    [switch]$Fresh,
    [switch]$AllPhotos,
    [ValidateSet('Inspect','Tag','Undo','ImportRoster','FindRoster','SaveRoster')][string]$GuidedAction,
    [string]$RequestFile
)

$ErrorActionPreference = 'Stop'
$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
if ($RosterTest -or $GuidedAction) {
    $selectedFolder = Get-Item -LiteralPath $Folder
    if (-not $selectedFolder.PSIsContainer -or $selectedFolder.PSProvider.Name -ne 'FileSystem') {
        throw 'Use -Folder to select a local photo folder.'
    }
    $gameFolder = $selectedFolder.FullName
    $folder = $gameFolder
    $originals = Join-Path $gameFolder 'PHOTOS\ORIGINALS'
    if (Test-Path -LiteralPath $originals -PathType Container) {
        $folder = $originals
    } elseif ($selectedFolder.Name -eq 'ORIGINALS' -and $selectedFolder.Parent.Name -eq 'PHOTOS') {
        $gameFolder = $selectedFolder.Parent.Parent.FullName
    }
    $stateName = if ($GuidedAction) { '00_SportsPhotoTagger' } else { '00_RosterKeywordTest' }
    $state = Join-Path $gameFolder $stateName
    if ($GuidedAction) {
        # Reuse saved runs and backups in game folders created before the rename.
        $legacyState = Join-Path $gameFolder '00_GameKeywords'
        if (Test-Path -LiteralPath $legacyState) {
            $state = $legacyState
        }
    }
    $referenceRoot = Join-Path $workspace '06_Sports Reference\Photo Keyword Reference'
} else {
    $folder = (Get-Item -LiteralPath (Join-Path $workspace '01_Incoming\Folder Inspection Test')).FullName
    $gameFolder = $folder
    $state = Join-Path $folder '00_KeywordTest'
}
$previewPath = Join-Path $state 'Preview.json'
$activePath = Join-Path $state 'ActiveRun.json'
$writer = Join-Path $PSScriptRoot 'Add-KeywordToTestImage.ps1'
$encoding = [Text.UTF8Encoding]::new($false)

function Assert-Local($Path) {
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if ($item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "Choose a regular local file or folder: $Path"
    }
}

function Get-JpegPath($Name) {
    if ($Name -isnot [string] -or -not $Name -or [IO.Path]::GetFileName($Name) -cne $Name -or
        $Name.Contains(':') -or [IO.Path]::GetExtension($Name) -notin '.jpg','.jpeg') {
        throw "Invalid JPEG filename: $Name"
    }
    $path = Join-Path $folder $Name
    Assert-Local $path
    if (-not [IO.File]::Exists($path)) { throw "JPEG is missing: $path" }
    return $path
}

function Get-Digest($Path) {
    $digest = [Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [IO.File]::OpenRead($Path)
        return [BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-', '')
    } finally {
        if ($stream) { $stream.Dispose() }
        $digest.Dispose()
    }
}

function Save-Json($Path, $Value) {
    Assert-Local $Path
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($temporary, (ConvertTo-Json -InputObject $Value -Depth 10), $encoding)
        if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temporary, $Path, [NullString]::Value) }
        else { [IO.File]::Move($temporary, $Path) }
    } finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function Read-Json($Path) {
    Assert-Local $Path
    if (-not [IO.File]::Exists($Path)) { throw "Missing file: $Path" }
    ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($Path))
}

function Write-PreviewImage($Source, $Destination) {
    $stream = [IO.File]::Open($Source, 'Open', 'Read', 'ReadWrite')
    try {
        $frame = [Windows.Media.Imaging.JpegBitmapDecoder]::new($stream, 'PreservePixelFormat', 'OnDemand').Frames[0]
        $scale = [Math]::Min(1.0, 1200.0 / [Math]::Max($frame.PixelWidth, $frame.PixelHeight))
        $orientation = 1
        if ($frame.Metadata) {
            $tag = $frame.Metadata.GetQuery('/app1/ifd/{ushort=274}')
            if ($tag) { $orientation = [int]$tag }
        }
        $rotation = switch ($orientation) {
            { $_ -in 5,6 } { 90 }
            { $_ -in 3,4 } { 180 }
            { $_ -in 7,8 } { 270 }
            default { 0 }
        }
        $horizontal = if ($orientation -in 2,4,5,7) { -$scale } else { $scale }
        $transform = [Windows.Media.TransformGroup]::new()
        $transform.Children.Add([Windows.Media.ScaleTransform]::new($horizontal, $scale))
        $transform.Children.Add([Windows.Media.RotateTransform]::new($rotation))
        $image = [Windows.Media.Imaging.TransformedBitmap]::new($frame, $transform)
        $encoder = [Windows.Media.Imaging.JpegBitmapEncoder]::new()
        $encoder.QualityLevel = 84
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($image, $null, $null, $null))
        $target = [IO.File]::Open($Destination, 'CreateNew', 'Write', 'None')
        try { $encoder.Save($target) } finally { $target.Dispose() }
    } finally { $stream.Dispose() }
}

function Show-Plan($Rows) {
    foreach ($row in $Rows) {
        $words = if ($row.keywords.Count) { $row.keywords -join ' | ' } else { 'No proposed keywords' }
        Write-Host "  $($row.file): $words"
        if ($row.evidence) { Write-Host "    $($row.evidence)" }
    }
}

function Get-JpegInventory {
    foreach ($file in (Get-ChildItem -LiteralPath $folder -File | Sort-Object Name)) {
        if ($file.Extension -notin '.jpg','.jpeg' -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint)) { continue }
        [pscustomobject]@{ Name = $file.Name; Path = $file.FullName; Hash = Get-Digest $file.FullName }
    }
}

function Get-DistinctJpegs($Inventory) {
    $seen = @{}
    foreach ($file in $Inventory) {
        if (-not $seen.ContainsKey($file.Hash)) {
            $seen[$file.Hash] = $true
            $file
        }
    }
}

function Get-SampleJpegs {
    $unique = @(Get-DistinctJpegs @(Get-JpegInventory))
    if ($unique.Count -lt 5) { throw 'Place at least five distinct JPEGs directly in the photo folder.' }
    for ($i = 0; $i -lt 5; $i++) {
        $unique[[int][Math]::Round($i * ($unique.Count - 1) / 4)]
    }
}

function Invoke-GameCodex([string[]]$Arguments, [string]$InputText = '') {
    # Windows PowerShell jobs turn native stderr into error records, even for a successful Codex run.
    $ErrorActionPreference = 'Continue'
    $diagnostics = @($InputText | & codex.cmd @Arguments 2>&1 | Where-Object {
        $_ -is [Management.Automation.ErrorRecord]
    } | ForEach-Object { $_.Exception.Message })
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        $errorLine = $diagnostics | Where-Object { $_ -match '^\s*Error(?:\s|:)' } | Select-Object -Last 1
        $detail = if ($errorLine) { ' ' + $errorLine.Trim() } else { '' }
        throw "Codex request failed (exit $code). Check your sign-in or connection and retry.$detail"
    }
}
function Get-ModelObservations([string]$SchoolName, [string]$SportName, $Uniforms, $Activities, $Selected) {
    if (-not (Get-Command codex.cmd -ErrorAction SilentlyContinue)) {
        throw 'Codex CLI is missing. Use a PowerShell window with a working Codex sign-in.'
    }
    $scratch = Join-Path $state ('Preview-' + [guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($scratch) | Out-Null
    $schemaPath = Join-Path $scratch 'schema.json'
    $resultPath = Join-Path $scratch 'result.json'
    $attachments = @()
    try {
        $schema = @{
            type = 'object'; additionalProperties = $false; required = @('frames')
            properties = @{ frames = @{
                type = 'array'; items = @{
                    type = 'object'; additionalProperties = $false
                    required = @('frame','sport_seen','school_seen','activity','numbers','number_uncertain','evidence')
                    properties = @{
                        frame = @{ type = 'integer' }
                        sport_seen = @{ type = 'boolean' }
                        school_seen = @{ type = 'boolean' }
                        activity = @{ type = 'string' }
                        numbers = @{ type = 'array'; items = @{ type = 'string' } }
                        number_uncertain = @{ type = 'boolean' }
                        evidence = @{ type = 'string' }
                    }
                }
            } }
        }
        Save-Json $schemaPath $schema
        $prompt = @(
            "Inspect exactly $($selected.Count) image attachments for a sports keyword test, in attachment order 1 to $($selected.Count)."
            'Return one JSON frames entry per image. Use visible evidence from the images.'
            "Set sport_seen true for $SportName action and school_seen true when a pictured player or activity group clearly belongs to $SchoolName."
            "Uniform examples for $SchoolName in this sport are $($Uniforms -join ' / '). Use visible affiliation and game context."
            "If the main subject is one of these activities, use its exact keyword in activity: $($Activities -join ', '). Otherwise use the empty string."
            "Record up to five $SchoolName jersey numbers as digit strings only when both the digits and the school are clear."
            'Keep leading zeroes. For wide sideline or group photos include only main subjects.'
            'Set number_uncertain true if a relevant jersey number is partly covered, unclear, or unreadable. Give a short evidence note per photo.'
            'Leave the opponent and player names unresolved. Use only the attachments, without using tools or editing files.'
        ) -join ' '
        $codexArgs = @('exec',$prompt,'--model','gpt-6.2-sol','-c','model_reasoning_effort=medium',
            '--skip-git-repo-check','--sandbox','read-only','--ephemeral','--cd',$scratch,
            '--output-schema',$schemaPath,'--output-last-message',$resultPath)
        Write-Host "Analyzing $($selected.Count) JPEG previews with Codex:"
        for ($i = 0; $i -lt $selected.Count; $i++) {
            $photo = $selected[$i]
            $attachment = Join-Path $scratch ("{0}.jpg" -f ($i + 1))
            $attachments += $attachment
            Write-Host "  $($i + 1). $($photo.Name)"
            Write-PreviewImage $photo.Path $attachment
            if ((Get-Digest $photo.Path) -ne $photo.Hash) { throw "Photo changed during preview: $($photo.Path)" }
            $codexArgs += @('--image', $attachment)
        }
        Invoke-GameCodex $codexArgs
        $result = Read-Json $resultPath
        if (@($result.frames).Count -ne $selected.Count) {
            throw 'The result must contain one observation for each attached JPEG.'
        }
        for ($i = 0; $i -lt $selected.Count; $i++) {
            $matchesForFrame = @($result.frames | Where-Object frame -eq ($i + 1))
            if ($matchesForFrame.Count -ne 1) { throw "Unexpected observations for photo $($i + 1)." }
            $observation = $matchesForFrame[0]
            if ($observation.sport_seen -isnot [bool] -or $observation.school_seen -isnot [bool] -or
                $observation.number_uncertain -isnot [bool] -or $observation.activity -isnot [string] -or
                ($observation.activity -and $observation.activity -cnotin $Activities) -or
                $observation.numbers -isnot [array] -or @($observation.numbers).Count -gt 5 -or
                @($observation.numbers | Where-Object { $_ -isnot [string] -or $_ -notmatch '^[0-9]{1,2}$' }).Count) {
                throw "Invalid observation for photo $($i + 1)."
            }
            $keywords = @()
            if ($observation.activity) { $keywords += $observation.activity }
            elseif ($observation.sport_seen) { $keywords += $SportName }
            if ($observation.school_seen) {
                $keywords += $SchoolName
                if ($observation.sport_seen -and -not $observation.activity) {
                    $keywords += @($observation.numbers | Select-Object -Unique | ForEach-Object { "$SchoolName #$_" })
                }
            }
            @{
                file = $selected[$i].Name; sha256 = $selected[$i].Hash
                keywords = $keywords; evidence = [string]$observation.evidence
                number_uncertain = $observation.number_uncertain
            }
        }
    } finally {
        foreach ($path in @($attachments) + @($schemaPath,$resultPath)) {
            if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
        }
        if ([IO.Directory]::Exists($scratch) -and -not (Get-ChildItem -LiteralPath $scratch -Force | Select-Object -First 1)) {
            [IO.Directory]::Delete($scratch)
        }
    }
}

function Invoke-Preview {
    if ([IO.File]::Exists($activePath)) { throw 'Reset the active trial before previewing another batch.' }
    if ([IO.File]::Exists($previewPath)) { Save-Json $previewPath @{ version = 0 } }
    if ($RosterTest) {
        Invoke-RosterPreview
        return
    }
    $rows = @(Get-ModelObservations 'Franklin High School' 'Volleyball' @(
        'Maroon with white lettering and numbers', 'White with maroon lettering, numbers, and trim for the libero'
    ) @() @(Get-SampleJpegs))
    Save-Json $previewPath @{ version = 1; photos = $rows }
    Write-Host 'Keyword preview:'
    Show-Plan $rows
    Write-Host "Saved preview to $previewPath"
}

function Read-Keywords($Path) {
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        $metadata = [Windows.Media.Imaging.JpegBitmapDecoder]::new($stream, 'PreservePixelFormat', 'OnDemand').Frames[0].Metadata
        if ($metadata) { $metadata.Keywords }
    } finally { $stream.Dispose() }
}

function New-TrialRun($PreviousRun) {
    if (-not [IO.File]::Exists($previewPath)) { throw 'Run Preview before Apply.' }
    $preview = Read-Json $previewPath
    $expectedVersion = if ($RosterTest -or $GuidedAction) { 2 } else { 1 }
    $photoCount = @($preview.photos).Count
    $validCount = if (($RosterTest -or $GuidedAction) -and $preview.scope -eq 'all') { $photoCount -gt 0 } else { $photoCount -eq 5 }
    if ($preview.version -ne $expectedVersion -or -not $validCount) {
        throw 'Run a successful Preview before Apply.'
    }
    $paths = @{}
    $approved = @(foreach ($row in $preview.photos) {
        $path = Get-JpegPath $row.file
        if ($paths.ContainsKey($path) -or $row.sha256 -notmatch '^[A-F0-9]{64}$' -or
            (Get-Digest $path) -ne $row.sha256) { throw "Preview is out of date or invalid for $path. Run Preview again." }
        $paths[$path] = $true
        $invalid = @($row.keywords | Where-Object {
            if ($_ -isnot [string] -or -not $_.Trim()) { return $true }
            if ($RosterTest -or $GuidedAction) { return $_ -cnotin $preview.permitted_keywords }
            return $_ -cnotmatch '^(Volleyball|Franklin High School( #[0-9]{1,2})?)$'
        })
        if ($row.keywords -isnot [array] -or $invalid.Count) {
            throw "The preview contains an unsupported keyword for $path."
        }
        if ($row.keywords.Count) { $row }
    })
    if (-not $approved.Count) { throw 'The preview has no keywords to apply.' }
    Show-Plan $approved
    $id = [guid]::NewGuid().ToString('N')
    $runFolder = Join-Path $state $id
    [IO.Directory]::CreateDirectory($runFolder) | Out-Null
    $rows = @(for ($i = 0; $i -lt $approved.Count; $i++) {
        $row = $approved[$i]
        $number = $i + 1
        $original = Join-Path $runFolder "$number-before.jpg"
        $tagged = Join-Path $runFolder "$number-after.jpg"
        [IO.File]::Copy((Get-JpegPath $row.file), $original)
        if ((Get-Digest $original) -ne $row.sha256) { throw "Photo changed during backup: $($row.file)" }
        [IO.File]::Copy($original, $tagged)
        foreach ($keyword in $row.keywords) {
            & $writer -Image $tagged -Keyword $keyword -Apply | Out-Null
        }
        $readBack = @(Read-Keywords $tagged)
        foreach ($keyword in $row.keywords) {
            if ($keyword -cnotin $readBack) { throw "Keyword readback failed for $($row.file): $keyword" }
        }
        @{
            number = $number; file = $row.file; keywords = $row.keywords
            before = $row.sha256; after = Get-Digest $tagged
        }
    })
    $run = @{ version = 1; id = $id; photos = $rows }
    if ($PreviousRun) { Save-Json (Join-Path $runFolder 'PreviousRun.json') $PreviousRun }
    Save-Json $activePath $run
    return $run
}

function Set-TrialPhotos($Run, [bool]$Resetting) {
    $photoCount = @($Run.photos).Count
    if ($Run.version -ne 1 -or $Run.id -notmatch '^[a-f0-9]{32}$' -or
        $photoCount -lt 1 -or (-not ($RosterTest -or $GuidedAction) -and $photoCount -gt 5)) { throw 'Invalid keyword trial state.' }
    $runFolder = Join-Path $state $Run.id
    Assert-Local $runFolder
    $paths = @{}
    $pending = @(foreach ($row in $Run.photos) {
        if ($row.number -notmatch '^[1-9][0-9]*$' -or [int]$row.number -gt $photoCount -or
            $row.before -notmatch '^[A-F0-9]{64}$' -or
            $row.after -notmatch '^[A-F0-9]{64}$') { throw 'Invalid keyword trial entry.' }
        $path = Get-JpegPath $row.file
        if ($paths.ContainsKey($path)) { throw "Duplicate trial entry: $path" }
        $paths[$path] = $true
        $desired = if ($Resetting) { $row.before } else { $row.after }
        $other = if ($Resetting) { $row.after } else { $row.before }
        $suffix = if ($Resetting) { 'before' } else { 'after' }
        $source = Join-Path $runFolder "$($row.number)-$suffix.jpg"
        Assert-Local $source
        if (-not [IO.File]::Exists($source) -or (Get-Digest $source) -ne $desired) {
            throw "A saved test copy is missing or changed: $source"
        }
        $current = Get-Digest $path
        if ($current -ne $desired -and $current -ne $other) {
            throw "The JPEG was edited outside this trial: $path. Its saved copies are in $runFolder."
        }
        if ($current -ne $desired) {
            @{ path = $path; source = $source; current = $current; desired = $desired }
        }
    })
    foreach ($item in $pending) {
        $stage = Join-Path $folder ('.keyword-' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            [IO.File]::Copy($item.source, $stage)
            if ((Get-Digest $stage) -ne $item.desired -or (Get-Digest $item.path) -ne $item.current) {
                throw "The file changed during the trial: $($item.path)"
            }
            [IO.File]::Replace($stage, $item.path, [NullString]::Value)
            if ((Get-Digest $item.path) -ne $item.desired) { throw "JPEG verification failed: $($item.path)" }
        } finally {
            if ([IO.File]::Exists($stage)) { [IO.File]::Delete($stage) }
        }
    }
}

for ($directory = Get-Item -LiteralPath $folder; $directory; $directory = $directory.Parent) {
    Assert-Local $directory.FullName
}
if ($RosterTest -and $Action -eq 'Preview') {
    Assert-Local $referenceRoot
    Assert-Local (Join-Path $referenceRoot 'Rosters')
    . (Join-Path $PSScriptRoot 'RosterKeywordReview.ps1')
}
Assert-Local $state
Assert-Local $previewPath
Assert-Local $activePath
Add-Type -AssemblyName PresentationCore, WindowsBase
$mutex = [Threading.Mutex]::new($false, 'Local\SportsPhotographyFolderKeywordTrial')
$acquired = $false
try {
    try { $acquired = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $acquired = $true }
    if (-not $acquired) { throw 'Another keyword trial command is running.' }
    Write-Host "JPEG test folder: $folder"
    if ($GuidedAction) {
        . (Join-Path $PSScriptRoot 'PhotoCaptureTime.ps1')
        . (Join-Path $PSScriptRoot 'GuidedSportsPhotoTagger.ps1')
        Invoke-GuidedAction
        return
    }
    switch ($Action) {
        'Preview' {
            [IO.Directory]::CreateDirectory($state) | Out-Null
            Invoke-Preview
        }
        'Apply' {
            $run = if ([IO.File]::Exists($activePath)) { Read-Json $activePath } else { New-TrialRun }
            Set-TrialPhotos $run $false
            Write-Host "Applied and verified the preview keywords in $(@($run.photos).Count) JPEGs. Press / in Photo Mechanic, then I on a photo to check Keywords."
        }
        'Reset' {
            if (-not [IO.File]::Exists($activePath)) { Write-Host 'There is no applied trial to reset.'; break }
            $run = Read-Json $activePath
            Set-TrialPhotos $run $true
            $archive = Join-Path (Join-Path $state $run.id) 'Reset.json'
            [IO.File]::Move($activePath, $archive)
            Write-Host "Restored $(@($run.photos).Count) JPEGs to their original bytes. Press / in Photo Mechanic to rescan."
        }
    }
} finally {
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}

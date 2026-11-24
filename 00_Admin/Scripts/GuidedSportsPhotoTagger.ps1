$jpeg = '.jpg','.jpeg'
. (Join-Path $PSScriptRoot 'RosterIdentity.ps1')
. (Join-Path $PSScriptRoot 'GuidedRosterSources.ps1')

function Get-GuidedInventory {
    $shellFolder = $null
    try { $shellFolder = (New-Object -ComObject Shell.Application).NameSpace($folder) } catch {}
    $frames = @(foreach ($item in (Get-JpegInventory)) {
        $item | Add-Member -NotePropertyName Captured -NotePropertyValue (Get-CaptureTime @((Get-Item -LiteralPath $item.Path)) $shellFolder)
        $item
    })
    $frames = @($frames | Sort-Object @{Expression={-not $_.Captured}}, Captured, Name)
    $index = 0
    foreach ($item in $frames) {
        $item | Add-Member -NotePropertyName Number -NotePropertyValue (++$index)
        $item
    }
}

function Invoke-GameVision($Photos, [string]$Instructions, [bool]$Inspect) {
    if (-not (Get-Command codex.cmd -ErrorAction SilentlyContinue)) {
        throw 'Install Codex CLI and sign in to ChatGPT with codex.cmd login.'
    }
    $scratch = Join-Path $state ('Vision-' + [guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($scratch) | Out-Null
    $schemaFile = Join-Path $scratch 'schema.json'
    $resultFile = Join-Path $scratch 'result.json'
    $files = @($schemaFile,$resultFile)
    try {
        $text = @{type='string'}
        $yes = @{type='boolean'}
        $numbers = @{type='array';items=$text}
        $rowSchema = @{
            type='object'; additionalProperties=$false
            required=@('frame','activity','side1','side2','numbers1','numbers2','uncertain','evidence')
            properties=@{frame=@{type='integer'};activity=$text;side1=$yes;side2=$yes;
                numbers1=$numbers;numbers2=$numbers;uncertain=$yes;evidence=$text}
        }
        $fields = @{photos=@{type='array';items=$rowSchema}}
        $required = @('photos')
        if ($Inspect) {
            foreach ($key in 'sport','school1','school2','uniform1','uniform2','evidence') { $fields[$key] = $text; $required += $key }
            $fields.mixed = $yes
            $required += 'mixed'
        }
        Save-Json $schemaFile @{type='object';additionalProperties=$false;required=$required;properties=$fields}
        $argsForCodex = @('exec',$Instructions,'--model','gpt-6.2-sol','-c','model_reasoning_effort=medium',
            '--skip-git-repo-check','--sandbox','read-only','--ephemeral','--cd',$scratch,
            '--output-schema',$schemaFile,'--output-last-message',$resultFile)
        for ($index = 0; $index -lt $Photos.Count; $index++) {
            $attachment = Join-Path $scratch "$($index + 1).jpg"
            $files += $attachment
            Write-PreviewImage $Photos[$index].Path $attachment
            if ((Get-Digest $Photos[$index].Path) -ne $Photos[$index].Hash) {
                throw "The JPEG changed during inspection: $($Photos[$index].Name)"
            }
            $argsForCodex += @('--image',$attachment)
        }
        Invoke-GameCodex $argsForCodex
        $result = Read-Json $resultFile
        if (@($result.photos).Count -ne $Photos.Count) { throw 'The photo review returned an incomplete batch. Try again.' }
        $rows = @(for ($index = 0; $index -lt $Photos.Count; $index++) {
            $row = @($result.photos | Where-Object frame -eq ($index + 1))
            if ($row.Count -ne 1) { throw 'The photo review returned an ambiguous frame number.' }
            $entry = $row[0]
            $numbers1 = @($entry.numbers1 | Where-Object { $_ -match '^[0-9]{1,2}$' } | Select-Object -Unique -First 5)
            $remaining = 5 - $numbers1.Count
            $numbers2 = @($entry.numbers2 | Where-Object { $_ -match '^[0-9]{1,2}$' } | Select-Object -Unique -First $remaining)
            @{file=$Photos[$index].Name; sha256=$Photos[$index].Hash; activity=$entry.activity
              side1=$entry.side1; side2=$entry.side2; uncertain=$entry.uncertain; evidence=$entry.evidence
              numbers1=$numbers1;numbers2=$numbers2}
        })
        @{game=$result; photos=$rows}
    } finally {
        foreach ($path in $files) { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
        if (-not (Get-ChildItem -LiteralPath $scratch -Force | Select-Object -First 1)) { [IO.Directory]::Delete($scratch) }
    }
}

function Get-GuidedSchool([string]$Name) {
    $text = $Name.Trim()
    if (-not $text -or $text -match '^(unknown|unresolved)( school| team)?$') { return $null }
    $matches = @($catalog | Where-Object { $_.name -eq $text -or $text -in $_.aliases })
    if ($matches.Count -eq 1) { return $matches[0].name }
    return $text
}

function Get-GuidedRosters($Game) {
    $rosterSport = Resolve-RosterSport $Game.sport $labels
    $rosterSeason = Resolve-RosterSeason $Game.season $Game.date
    foreach ($side in @($Game.sides)) {
        $matches = @(foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $referenceRoot 'Rosters') -File -Filter '*.json')) {
            $roster = Read-Json $file.FullName
            if ($side.school -and $roster.school -eq $side.school -and $roster.sport -eq $rosterSport -and
                $roster.season -eq $rosterSeason -and $roster.squad -eq $side.squad) {
                @{path=$file.FullName; data=$roster; version=(Get-Digest $file.FullName)}
            }
        })
        @{ matches=$matches }
    }
}

function Invoke-GuidedInspection($Request) {
    $frames = @(Get-GuidedInventory)
    if (-not $frames.Count) { throw 'This folder has no JPEGs. Choose the folder with the JPEGs you use in Photo Mechanic.' }
    $inspectionFrames = $frames
    $selectedRange = $null
    if ($Request.more -and $Request.range -match '^([1-9][0-9]{0,7})-([1-9][0-9]{0,7})$') {
        $first = [int]$Matches[1]
        $last = [int]$Matches[2]
        if ($last -le $frames.Count -and $first -le $last) {
            $inspectionFrames = @($frames | Where-Object { $_.Number -ge $first -and $_.Number -le $last })
            $selectedRange = "$first-$last"
        }
    }
    $inventory = @(Get-DistinctJpegs $inspectionFrames)
    $dated = @($frames | Where-Object Captured)
    $groups = @()
    $start = 1
    for ($i = 1; $i -lt $dated.Count; $i++) {
        if (($dated[$i].Captured - $dated[$i-1].Captured).TotalHours -ge 2) {
            $groups += "$start-$i"
            $start = $i + 1
        }
    }
    if ($groups.Count) { $groups += "$start-$($dated.Count)" }
    if ($groups.Count -and $dated.Count -lt $frames.Count) { $groups += "$($dated.Count + 1)-$($frames.Count)" }
    $sampleSize = if ($Request.more) { 20 } else { 16 }
    $count = [Math]::Min($sampleSize,$inventory.Count)
    $sample = if ($count -eq 1) { @($inventory[0]) } else {
        @(for ($i=0;$i -lt $count;$i++) { $inventory[[int][Math]::Round($i * ($inventory.Count-1) / ($count-1))] })
    }
    $priorPath = Join-Path $state 'Game.json'
    $notes = @()
    if (-not $Request.more -and [IO.File]::Exists($priorPath)) {
        $prior = Read-Json $priorPath
        $game = $prior.game
        $notes += 'Loaded your saved game details. Confirm the photo range for this run.'
        $sampleRows = @($prior.examples)
        $mixed = [bool]$prior.mixed
        $evidence = $prior.evidence
    } else {
        $selectedDays = @($inspectionFrames | Where-Object Captured | Group-Object { $_.Captured.ToString('yyyy-MM-dd') })
        $date = if ($selectedDays.Count -eq 1) { $selectedDays[0].Name } else { '' }
        $season = if ($date) { $date.Substring(0,4) } else { '' }
        $game = @{date=$date;season=$season;sport='';sides=@(@{school='';uniform='';squad=''},@{school='';uniform='';squad=''})}
        $text = "Inspect the $count attached sports JPEGs, in attachment order. Give sport, both school names, typical jersey descriptions, and evidence. Use an empty name for an unresolved school. Assign distinct sides using play direction, visible branding, and uniform variants such as a contrasting libero or goalkeeper. Color alone does not establish a school. In each photo row, Side 1 means school1 or uniform1 and Side 2 means school2 or uniform2. Set mixed true only for evidence of multiple games. Known school names: $($catalog.name -join ', ')."
        Write-Host "Reading capture times for $($frames.Count) JPEGs and inspecting $count sample previews."
        $sampleRows = @()
        $mixed = $false
        $evidence = ''
        try {
            $answer = Invoke-GameVision @($sample) ($text + ' ' + (Get-PhotoInstructions)) $true
            $game.sport=Resolve-RosterSport $answer.game.sport $labels
            $game.sides[0].school=Get-GuidedSchool $answer.game.school1
            $game.sides[1].school=Get-GuidedSchool $answer.game.school2
            $game.sides[0].uniform=$answer.game.uniform1
            $game.sides[1].uniform=$answer.game.uniform2
            $sampleRows=$answer.photos
            $mixed=[bool]$answer.game.mixed
            $evidence=$answer.game.evidence
        } catch {
            $notes += "Visual inspection unavailable. $($_.Exception.Message)"
            $notes += 'Enter the game details yourself to continue.'
        }
    }
    $otherFiles = @(Get-ChildItem -LiteralPath $folder -File | Where-Object Extension -in '.arw','.cr2','.cr3','.dng','.nef','.nrw','.orf','.raf','.rw2')
    if ($otherFiles.Count) { $notes += "$($otherFiles.Count) RAW files are outside this JPEG tagging run." }
    if ($groups.Count -or $mixed) { $notes += 'Check for multiple games. Use the numbered list to pick one range or explicitly select All.' }
    $examples = @($sampleRows | Select-Object -First 4)
    $result = @{game=$game; examples=$examples; sample=$sampleRows; groups=$groups; mixed=($mixed -or $groups.Count -gt 0); evidence=$evidence
        selected_range=$selectedRange; notes=$notes; photos=@($frames | ForEach-Object { @{number=$_.Number;file=$_.Name;sha256=$_.Hash
            captured=$(if ($_.Captured) { $_.Captured.ToString('yyyy-MM-dd HH:mm:ss') }) } })}
    Save-Json $responseFile $result
}

function Get-PhotoInstructions {
    "Return one photos row per attached image, numbered in attachment order. side1 and side2 report visible affiliation. Use number lists only for clearly readable jerseys tied to that side, at most five total. Include background players when clear, and focus on main subjects in wide groups. Use activity only when the subject is one of these exact activities: $($labels.activities -join ', '). Otherwise use an empty activity. Mark uncertain when a relevant number or team attribution is unclear. Provide short concrete evidence. Use the attachments alone and do not change files."
}

function Invoke-GuidedTag($Request) {
    $game = $Request.game
    $scope = $Request.scope
    if (@($game.sides).Count -ne 2 -or -not $scope -or -not $game.sport) { throw 'Confirm a sport, two sides, and a photo range first.' }
    $preferred = @($labels.sports | Where-Object { $_ -eq $game.sport })
    if ($preferred.Count -eq 1) { $game.sport = $preferred[0] }
    foreach ($side in $game.sides) { $side.school = Get-GuidedSchool $side.school }
    if ($game.sides[0].school -and $game.sides[0].school -eq $game.sides[1].school) {
        throw 'Choose different schools for the two sides.'
    }
    $seen = @{}
    $inventory = @(foreach ($row in $scope) {
        $path = Get-JpegPath $row.file
        if ($seen.ContainsKey($path) -or (Get-Digest $path) -ne $row.sha256) {
            throw "The selected JPEGs changed. Inspect the folder again: $($row.file)"
        }
        $seen[$path]=$true
        @{Name=$row.file;Path=$path;Hash=$row.sha256}
    })
    if ([IO.File]::Exists($activePath)) {
        $priorRun = Read-Json $activePath
        foreach ($entry in $priorRun.photos) {
            if ((Get-Digest (Get-JpegPath $entry.file)) -ne $entry.after) {
                throw 'An earlier tagging run is unfinished or its JPEGs changed. Use Undo Last Run before applying a new batch.'
            }
        }
    } else { $priorRun = $null }

    $cachePath = Join-Path $state 'Observations.json'
    $identity = "$($game.sport)|$($game.sides[0].uniform)|$($game.sides[1].uniform)"
    $sideIdentity = @($game.sides | ForEach-Object { [string]$_.school })
    $byHash = @{}
    if ([IO.File]::Exists($cachePath)) {
        $cache = Read-Json $cachePath
        $sameSides = $true
        for ($i=0;$i -lt 2;$i++) {
            if ($cache.schools[$i] -and $sideIdentity[$i] -and $cache.schools[$i] -ne $sideIdentity[$i]) { $sameSides=$false }
        }
        if ($cache.identity -eq $identity -and $sameSides) {
            foreach ($entry in $cache.photos) { $byHash[$entry.sha256]=$entry }
        }
    }
    if ($priorRun -and $byHash.Count) {
        foreach ($item in $priorRun.photos) {
            if ($byHash.ContainsKey($item.before)) {
                $source = $byHash[$item.before]
                $byHash[$item.after] = @{
                    file=$source.file;sha256=$item.after;activity=$source.activity
                    side1=$source.side1;side2=$source.side2;uncertain=$source.uncertain;evidence=$source.evidence
                    numbers1=@($source.numbers1);numbers2=@($source.numbers2)
                }
            }
        }
    }
    $proposal = $Request.proposed_game
    if ($proposal -and $proposal.sport -eq $game.sport -and @($proposal.sides).Count -eq 2) {
        $sameUniforms = $true
        for ($i=0;$i -lt 2;$i++) {
            $before = $proposal.sides[$i]
            $after = $game.sides[$i]
            if ($before.uniform -ne $after.uniform -or ($before.school -and $before.school -ne $after.school)) {
                $sameUniforms = $false
            }
        }
        if ($sameUniforms) {
            foreach ($sample in $Request.sample) {
                if ($sample.sha256 -match '^[A-F0-9]{64}$' -and -not $byHash.ContainsKey($sample.sha256)) {
                    $byHash[$sample.sha256]=$sample
                }
            }
        }
    }
    $pending = @(Get-DistinctJpegs @($inventory | Where-Object { -not $byHash.ContainsKey($_.Hash) }))
    $hint = "These are photos of $($game.sport). Side 1 has school $($game.sides[0].school) and uniform $($game.sides[0].uniform). Side 2 has school $($game.sides[1].school) and uniform $($game.sides[1].uniform). A missing school can still have a side and numbers when the uniform or visible play supports it. Use no school guesses."
    Write-Host "$($inventory.Count) JPEGs selected. $($pending.Count) distinct photos need visual analysis."
    for ($start=0;$start -lt $pending.Count;$start+=12) {
        $batch=@($pending | Select-Object -Skip $start -First 12)
        Write-Host "Analyzing photos $($start+1) through $($start+$batch.Count) of $($pending.Count)."
        $answer=Invoke-GameVision $batch ($hint+' '+(Get-PhotoInstructions)) $false
        foreach ($entry in $answer.photos) { $byHash[$entry.sha256]=$entry }
        Save-Json $cachePath @{identity=$identity;schools=$sideIdentity;photos=@($byHash.Values)}
    }
    $rosters = @(Get-GuidedRosters $game)
    $rows = @(foreach ($item in $inventory) {
        $observation = $byHash[$item.Hash]
        $activity = if ($observation.activity -cin $labels.activities) { $observation.activity } else { $game.sport }
        $words = @($activity)
        $issues = @()
        for ($i=0;$i -lt 2;$i++) {
            $side = $game.sides[$i]
            $sideNumber = $i + 1
            $present = if ($i -eq 0) { $observation.side1 } else { $observation.side2 }
            $numbers = if ($i -eq 0) { $observation.numbers1 } else { $observation.numbers2 }
            if (-not $present) { continue }
            if (-not $side.school) {
                $unknown = if (@($numbers | Where-Object {$_}).Count) { ' (jerseys ' + ($numbers -join ', ') + ')' } else { '' }
                $issues += "Side $sideNumber school is unresolved$unknown"
                continue
            }
            $words += $side.school
            if ($activity -ne $game.sport) { continue }
            foreach ($number in $numbers) {
                if (-not $number) { continue }
                $words += "$($side.school) #$number"
                if ($rosters[$i].matches.Count -eq 1) {
                    $players = @($rosters[$i].matches[0].data.players | Where-Object { $_.number -ceq $number -and $_.name })
                    if ($players.Count -eq 1) { $words += $players[0].name }
                    else { $issues += "$($side.school) #$number has $($players.Count) roster matches" }
                }
            }
        }
        if ($observation.uncertain) { $issues += 'Jersey number or team unclear' }
        if (-not $observation.side1 -and -not $observation.side2) { $issues += 'School unclear' }
        @{file=$item.Name;sha256=$item.Hash;keywords=@($words | Select-Object -Unique);evidence=$observation.evidence;reasons=$issues}
    })
    $issues = @(foreach ($i in 0,1) {
        if (-not $game.sides[$i].school) { "Side $($i+1) school unresolved" }
        elseif ($rosters[$i].matches.Count -ne 1) { "$($game.sides[$i].school) roster needs review" }
    })
    if ($priorRun) {
        $priorByName = @{}
        foreach ($item in $priorRun.photos) { $priorByName[$item.file]=$item }
        foreach ($row in $rows) {
            $prior = $priorByName[$row.file]
            if ($prior -and @($prior.keywords | Where-Object { $_ -cnotin $row.keywords }).Count) {
                throw "The revised game or roster changes previously applied keywords on $($row.file). Use Undo Last Run before replacing those keywords."
            }
        }
    }
    $report = @{game=$issues;photos=@($rows | Where-Object {$_.reasons.Count})}
    Save-Json $cachePath @{identity=$identity;schools=$sideIdentity;photos=@($byHash.Values)}
    $gameReview = @{game=$game;examples=@($Request.examples);mixed=$Request.mixed;evidence=$Request.evidence
        scope=@($scope.file);rosters=@($rosters | ForEach-Object { @($_.matches | ForEach-Object { @{file=$_.path;version=$_.version} }) })}
    $records = Join-Path $state 'Games'
    Assert-Local $records
    [IO.Directory]::CreateDirectory($records) | Out-Null
    $identifier = [Security.Cryptography.SHA256]::Create()
    try {
        $scopeKey = [BitConverter]::ToString($identifier.ComputeHash($encoding.GetBytes((($scope.file | Sort-Object) -join [char]0)))).Replace('-','')
    } finally { $identifier.Dispose() }
    Save-Json (Join-Path $records "$scopeKey.json") $gameReview
    Save-Json (Join-Path $state 'Game.json') $gameReview
    Save-Json $previewPath @{version=2;scope='all';photos=$rows
        permitted_keywords=@($rows | ForEach-Object {$_.keywords} | Select-Object -Unique);manual_review=$report}
    Save-Json (Join-Path $state 'Manual Review.json') $report

    Write-Host "Writing and verifying keywords in $($rows.Count) JPEGs."
    $run = New-TrialRun $priorRun
    Set-TrialPhotos $run $false
    Save-Json $responseFile @{count=$run.photos.Count;review=$report;files=$run.photos}
}

function Invoke-GuidedAction {
    Assert-Local $referenceRoot
    $catalog = Read-Json (Join-Path $referenceRoot 'Teams.json')
    $labels = Read-Json (Join-Path $referenceRoot 'Keywords.json')
    if (-not $RequestFile) { throw 'Open the SportsPhotoTagger window to select the game and start the run.' }
    $request = Read-Json $RequestFile
    $responseFile = "$RequestFile.result"
    [IO.Directory]::CreateDirectory($state) | Out-Null
    switch ($GuidedAction) {
        'Inspect' { Invoke-GuidedInspection $request }
        'Tag' { Invoke-GuidedTag $request }
        'ImportRoster' { Import-GuidedRoster $request }
        'FindRoster' { Find-GuidedRoster $request }
        'SaveRoster' { Save-OnlineRoster $request }
        'Undo' {
            if (-not [IO.File]::Exists($activePath)) { Save-Json $responseFile @{count=0}; return }
            $run = Read-Json $activePath
            Set-TrialPhotos $run $true
            $runFolder = Join-Path $state $run.id
            $previous = Join-Path $runFolder 'PreviousRun.json'
            [IO.File]::Move($activePath,(Join-Path $runFolder 'Undo.json'))
            if ([IO.File]::Exists($previous)) { Save-Json $activePath (Read-Json $previous) }
            Save-Json $responseFile @{count=$run.photos.Count}
        }
    }
}

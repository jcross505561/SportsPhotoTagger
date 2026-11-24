function Get-ReferenceSchool([string]$Name, $Catalog) {
    $matches = @($Catalog | Where-Object { $_.name -eq $Name -or $Name -in $_.aliases })
    if ($matches.Count -ne 1) {
        throw "Enter a unique school name or alias from the reference catalog: $Name"
    }
    return $matches[0]
}

function Show-ManualReview($GameReview, $PhotoReview) {
    Write-Host 'Manual review:'
    foreach ($issue in $GameReview) { Write-Host "  Game: $issue" }
    foreach ($item in $PhotoReview) {
        Write-Host "  $($item.file): $($item.reasons -join '; ')"
        if ($item.evidence) { Write-Host "    $($item.evidence)" }
    }
    if (-not $GameReview.Count -and -not $PhotoReview.Count) { Write-Host '  No open items.' }
}

function Get-AllPhotoObservations($Saved, $TeamName, $SportName, $Uniforms, $Activities) {
    $inventory = @(Get-JpegInventory)
    if (-not $inventory.Count) { throw 'Place JPEGs directly in the selected photo folder.' }
    $byHash = @{}
    foreach ($row in $Saved) {
        if ($row.sha256 -match '^[A-F0-9]{64}$') { $byHash[$row.sha256] = $row }
    }
    $missing = @($inventory | Where-Object { -not $byHash.ContainsKey($_.Hash) })
    $pending = @(Get-DistinctJpegs $missing)
    Write-Host "All-photo preview covers $($inventory.Count) JPEGs. Reusing observations for $($inventory.Count - $missing.Count) files."
    if ($pending.Count) {
        Write-Host "Sending $($pending.Count) distinct new previews to Codex."
        $batchSize = 12
        for ($offset = 0; $offset -lt $pending.Count; $offset += $batchSize) {
            $batch = @($pending | Select-Object -Skip $offset -First $batchSize)
            foreach ($row in (Get-ModelObservations $TeamName $SportName $Uniforms $Activities $batch)) {
                $byHash[$row.sha256] = $row
            }
        }
    }
    foreach ($file in $inventory) {
        $observation = $byHash[$file.Hash]
        if (-not $observation) { throw "Missing observation for $($file.Name)." }
        @{
            file = $file.Name
            sha256 = $file.Hash
            keywords = @($observation.keywords)
            evidence = $observation.evidence
            number_uncertain = [bool]$observation.number_uncertain
        }
    }
}

function Invoke-RosterPreview {
    if (-not $School.Trim() -or -not $Sport.Trim() -or -not $Season.Trim() -or -not $Squad.Trim()) {
        throw 'Preview needs -School, -Sport, -Season, and -Squad for the game under review.'
    }
    $schoolCatalog = Read-Json (Join-Path $referenceRoot 'Teams.json')
    $labelCatalog = Read-Json (Join-Path $referenceRoot 'Keywords.json')
    $team = Get-ReferenceSchool $School.Trim() $schoolCatalog
    $teamName = $team.name
    $sportLabel = @($labelCatalog.sports | Where-Object { $_ -eq $Sport.Trim() })
    if ($sportLabel.Count -ne 1) { throw "Enter a sport from the reference keyword list: $Sport" }
    $sportName = $sportLabel[0]
    $seasonName = $Season.Trim()
    $squadName = $Squad.Trim()
    $opponentName = $null
    if ($Opponent.Trim()) {
        $opponentName = (Get-ReferenceSchool $Opponent.Trim() $schoolCatalog).name
        if ($opponentName -eq $teamName) { throw 'Choose a different school for the opponent.' }
    }

    $rosterFolder = Join-Path $referenceRoot 'Rosters'
    $rosterMatches = @(foreach ($file in (Get-ChildItem -LiteralPath $rosterFolder -File -Filter '*.json')) {
        Assert-Local $file.FullName
        $text = [IO.File]::ReadAllText($file.FullName)
        $definition = ConvertFrom-Json -InputObject $text
        if ($definition.school -eq $teamName -and $definition.sport -eq $sportName -and
            $definition.season -eq $seasonName -and $definition.squad -eq $squadName) {
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $digest = [BitConverter]::ToString($sha.ComputeHash($encoding.GetBytes($text))).Replace('-', '')
            } finally { $sha.Dispose() }
            @{ File = $file.Name; Digest = $digest; Definition = $definition }
        }
    })
    $roster = if ($rosterMatches.Count -eq 1) { $rosterMatches[0] } else { $null }
    $gameReview = @()
    if (-not $opponentName) { $gameReview += 'Opponent unresolved.' }
    if (-not $roster) {
        $gameReview += if ($rosterMatches.Count) { 'More than one roster matches this school, sport, season, and squad.' }
            else { 'No roster matches this school, sport, season, and squad.' }
    }
    Write-Host "Game for review: $teamName | $sportName | $seasonName | $squadName"
    $opponentDisplay = if ($opponentName) { $opponentName } else { 'Unresolved' }
    Write-Host "Opponent: $opponentDisplay"
    $uniforms = @($team.uniforms | Where-Object sport -eq $sportName | ForEach-Object {
        "$($_.role) - $($_.jersey)"
    })
    foreach ($uniform in $uniforms) { Write-Host "  Uniform reference: $uniform" }
    if ($roster) {
        Write-Host "Roster for review: $($roster.Definition.source)"
        Write-Host "  Reference file: $($roster.File) | Source checked: $($roster.Definition.source_checked)"
    }

    $samplePath = Join-Path $state 'Observations.json'
    $observationsPath = if ($AllPhotos) { Join-Path $state 'AllObservations.json' } else { $samplePath }
    $legacyPath = Join-Path $gameFolder '00_KeywordTest\Preview.json'
    $observations = $null
    if (-not $Fresh) {
        foreach ($path in (@($observationsPath, $samplePath) | Select-Object -Unique)) {
            if (-not [IO.File]::Exists($path)) { continue }
            $saved = Read-Json $path
            if ($saved.version -eq 1 -and $saved.school -eq $teamName -and $saved.sport -eq $sportName) {
                $observations = $saved
                break
            }
        }
    }
    if (-not $Fresh -and -not $observations -and $teamName -eq 'Franklin High School' -and
        $sportName -eq 'Volleyball' -and [IO.File]::Exists($legacyPath)) {
        Assert-Local (Split-Path -Parent $legacyPath)
        $legacy = Read-Json $legacyPath
        if ($legacy.version -eq 1 -and @($legacy.photos).Count -eq 5) {
            $observations = @{ version = 1; school = $teamName; sport = $sportName; photos = @($legacy.photos) }
        }
    }
    if ($AllPhotos) {
        $rows = @(Get-AllPhotoObservations $observations.photos $teamName $sportName $uniforms $labelCatalog.activities)
        $observations = @{ version = 1; school = $teamName; sport = $sportName; photos = $rows }
    } elseif ($observations) {
        Write-Host 'Reusing five saved photo observations.'
    } else {
        $rows = @(Get-ModelObservations $teamName $sportName $uniforms $labelCatalog.activities @(Get-SampleJpegs))
        $observations = @{ version = 1; school = $teamName; sport = $sportName; photos = $rows }
    }
    if (-not $AllPhotos -and @($observations.photos).Count -ne 5) { throw 'The observation set must contain five photos.' }

    $numberPattern = '^' + [regex]::Escape($teamName) + ' #([0-9]{1,2})$'
    $observedTags = @($teamName, $sportName) + @($labelCatalog.activities)
    $seen = @{}
    $rows = @(foreach ($observation in $observations.photos) {
        $path = Get-JpegPath $observation.file
        if ($seen.ContainsKey($path) -or $observation.sha256 -notmatch '^[A-F0-9]{64}$' -or
            (Get-Digest $path) -ne $observation.sha256) {
            throw "The JPEG differs from the saved observation: $path. Reset any active keyword trial or run Preview with -Fresh."
        }
        $seen[$path] = $true
        $numbers = @($observation.keywords | ForEach-Object {
            if ($_ -cmatch $numberPattern) { $Matches[1] }
        })
        $keywords = @($observation.keywords | Where-Object { $_ -cin $observedTags })
        $matchesForPhoto = @()
        $reasons = @()
        if ($teamName -cin $keywords) {
            $keywords += @($numbers | ForEach-Object { "$teamName #$_" })
            if ($observation.number_uncertain -or ($sportName -cin $keywords -and -not $numbers.Count)) {
                $reasons += 'Jersey number partly covered or unreadable'
            }
            foreach ($number in $numbers) {
                if (-not $roster) { continue }
                $players = @($roster.Definition.players | Where-Object { $_.number -ceq $number -and $_.name })
                if ($players.Count -eq 1) {
                    $name = $players[0].name.Trim()
                    if ($name) {
                        $keywords += $name
                        $matchesForPhoto += @{ number = $number; name = $name }
                    }
                } elseif ($players.Count -gt 1) {
                    $reasons += "More than one roster entry for #$number"
                } else {
                    $reasons += "No roster entry for #$number"
                }
            }
        } elseif ($sportName -cin $keywords) {
            $reasons += 'Target school unresolved in this photo'
        } elseif (-not $keywords.Count) {
            $reasons += 'Sport or activity unresolved in this photo'
        }
        @{
            file = $observation.file; sha256 = $observation.sha256; evidence = $observation.evidence
            keywords = @($keywords | Select-Object -Unique); player_matches = $matchesForPhoto
            reasons = $reasons
        }
    })
    $photoReview = @($rows | Where-Object { $_.reasons.Count })
    $scope = if ($AllPhotos) { 'all' } else { 'sample' }
    $plan = @{
        version = 2; photos = $rows; scope = $scope
        permitted_keywords = @($rows | ForEach-Object { $_.keywords } | Select-Object -Unique)
        game = @{
            school = $teamName; sport = $sportName; season = $seasonName; squad = $squadName
            opponent = $opponentName; roster_file = $roster.File; roster_version = $roster.Digest
        }
        manual_review = @{ game = $gameReview; photos = @($photoReview | ForEach-Object {
            @{ file = $_.file; reasons = $_.reasons; evidence = $_.evidence }
        }) }
    }
    Save-Json $observationsPath $observations
    Save-Json $previewPath $plan
    Write-Host 'Keyword preview:'
    Show-Plan $rows
    foreach ($row in ($rows | Where-Object { $_.player_matches.Count })) {
        $matchesText = @($row.player_matches | ForEach-Object { "#$($_.number) = $($_.name)" }) -join ', '
        Write-Host "  Roster match for $($row.file): $matchesText"
    }
    Show-ManualReview $gameReview $photoReview
    Write-Host "Saved preview to $previewPath"
}

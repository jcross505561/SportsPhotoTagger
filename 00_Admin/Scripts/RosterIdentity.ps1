function Resolve-RosterSport([string]$Value, $Vocabulary) {
    $text = $Value.Trim()
    if ($text -eq 'American football') { $text = 'Football' }
    $preferred = @($Vocabulary.sports | Where-Object { $_ -eq $text })
    if ($preferred.Count -eq 1) { return $preferred[0] }
    return $text
}

function Resolve-RosterSeason([string]$Value, [string]$GameDate) {
    $text = $Value.Trim()
    if ($text -match '^(19|20)\d{2}(?:-(?:\d{2}|(?:19|20)\d{2}))?$') { return $text }
    if ($GameDate -match '^((?:19|20)\d{2})-\d{2}-\d{2}$') { return $Matches[1] }
    return $text
}

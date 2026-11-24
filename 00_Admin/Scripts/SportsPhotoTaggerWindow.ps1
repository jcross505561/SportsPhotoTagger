param([switch]$SmokeTest)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
. (Join-Path $PSScriptRoot 'RosterIdentity.ps1')
[Windows.Forms.Application]::EnableVisualStyles()
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$engine=Join-Path $PSScriptRoot 'Test-FolderKeywords.ps1'
$references=Join-Path $root '06_Sports Reference\Photo Keyword Reference'
$schools=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText((Join-Path $references 'Teams.json')))
$vocabulary=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText((Join-Path $references 'Keywords.json')))
$form=[Windows.Forms.Form]::new()
$form.Text='SportsPhotoTagger'
$form.StartPosition='CenterScreen'
$form.ClientSize=[Drawing.Size]::new(980,638)
$form.FormBorderStyle='FixedDialog'
$form.MaximizeBox=$false
$form.Font=[Drawing.Font]::new('Segoe UI',10)
$form.AutoScaleMode='Dpi'
function Put($item,$x,$y,$w,$h) {
    $item.SetBounds($x,$y,$w,$h)
    $form.Controls.Add($item)
    $item
}
function Text-Label($text,$x,$y,$w) {
    $item=[Windows.Forms.Label]::new()
    $item.Text=$text
    Put $item $x $y $w 26
}
function Text-Field($x,$y,$w) { Put ([Windows.Forms.TextBox]::new()) $x $y $w 26 }
function Choice($x,$y,$w,$values) {
    $item=[Windows.Forms.ComboBox]::new()
    $item.AutoCompleteSource='ListItems'
    $item.AutoCompleteMode='SuggestAppend'
    foreach($value in $values){[void]$item.Items.Add($value)}
    Put $item $x $y $w 26
}
function Make-Button($text,$x,$y,$w,$click) {
    $item=[Windows.Forms.Button]::new()
    $item.Text=$text
    $item.Add_Click($click)
    Put $item $x $y $w 38
}
function Notice($message) {
    [void][Windows.Forms.MessageBox]::Show($form,$message,'SportsPhotoTagger','OK','Information')
}
$script:job=$null
$script:session=$null
$script:operation=''
$script:requestFile=''
$script:inspectionText=''
$script:showingReport=$false
$script:populating=$false
$script:log=@()
[void](Text-Label 'Photo folder' 20 14 120)
$folderField=Text-Field 20 39 690
$folderField.ReadOnly=$true
$choose=Make-Button 'Choose Folder...' 722 32 142 {
    $dialog=[Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description='Choose the game folder or the folder containing the JPEGs'
    if($folderField.Text){$dialog.SelectedPath=$folderField.Text}
    try {
        if($dialog.ShowDialog($form) -eq 'OK'){
            $folderField.Text=$dialog.SelectedPath
            $script:session=$null
            $script:showingReport=$false
            $details.Clear()
            $exampleChoice.Items.Clear()
            Show-Example
            $status.Text='Ready to inspect the selected folder.'
            Refresh-Actions
        }
    }finally{$dialog.Dispose()}
}
$inspect=Make-Button 'Inspect' 875 32 85 { Start-Work 'Inspect' @{} }
[void](Text-Label 'Game date' 20 88 155)
[void](Text-Label 'Sport' 190 88 210)
[void](Text-Label 'Season' 415 88 130)
[void](Text-Label 'Photo range' 555 88 250)
$date=Text-Field 20 111 155
$sport=Choice 190 111 210 $vocabulary.sports
$season=Text-Field 415 111 125
$scope=Choice 555 111 405 @()
[void](Text-Label 'Side and school' 20 151 250)
[void](Text-Label 'Jersey or uniform description' 299 151 370)
[void](Text-Label 'Roster squad' 672 151 105)
$school1=Choice 20 175 265 $schools.name
$school2=Choice 20 234 265 $schools.name
$uniform1=Text-Field 299 175 357
$uniform2=Text-Field 299 234 357
$squad1=Choice 672 175 102 @('Girls Varsity','Boys Varsity','Varsity','Girls JV','Boys JV','JV','Freshman')
$squad2=Choice 672 234 102 $squad1.Items
$find1=Make-Button 'Find Online' 781 169 86 { Find-Roster 0 }
$find2=Make-Button 'Find Online' 781 228 86 { Find-Roster 1 }
$import1=Make-Button 'Upload CSV' 874 169 86 { Choose-Roster 0 }
$import2=Make-Button 'Upload CSV' 874 228 86 { Choose-Roster 1 }
$roster1=Text-Label 'Side 1' 20 201 810
$roster2=Text-Label 'Side 2' 20 260 810
$details=Put ([Windows.Forms.TextBox]::new()) 20 295 655 256
$details.Multiline=$true
$details.ReadOnly=$true
$details.ScrollBars='Vertical'
[void](Text-Label 'Example photo' 687 295 273)
$exampleChoice=Choice 687 320 273 @()
$exampleChoice.DropDownStyle='DropDownList'
$exampleImage=Put ([Windows.Forms.PictureBox]::new()) 687 355 273 196
$exampleImage.SizeMode='Zoom'
$exampleImage.BorderStyle='FixedSingle'
$exampleChoice.Add_SelectedIndexChanged({ Show-Example })
$status=Text-Label 'Choose a photo folder to start.' 20 562 940
$open=Make-Button 'Open Folder' 20 588 117 { Start-Process explorer.exe -ArgumentList ('"' + $folderField.Text + '"') }
$undo=Make-Button 'Undo Last Run' 146 588 134 {
    if ([Windows.Forms.MessageBox]::Show($form,'Restore the JPEGs from the last tagging run?','Undo Last Run','YesNo','Question','Button2') -eq 'Yes'){
        Start-Work 'Undo' @{}
    }
}
$more=Make-Button 'Inspect More' 645 588 134 { Start-Work 'Inspect' @{more=$true;range=$scope.Text.Trim()} }
$tag=Make-Button 'Confirm && Tag' 792 584 168 { Confirm-Game }
$form.AcceptButton=$inspect
$timer=[Windows.Forms.Timer]::new()
$timer.Interval=250

function Show-Example {
    if($exampleImage.Image){ $exampleImage.Image.Dispose(); $exampleImage.Image=$null }
    $name=[string]$exampleChoice.SelectedItem
    if(-not $name -or [IO.Path]::GetFileName($name) -ne $name){ return }
    $photoFolder=$folderField.Text
    $originals=Join-Path $photoFolder 'PHOTOS\ORIGINALS'
    if(Test-Path -LiteralPath $originals -PathType Container){ $photoFolder=$originals }
    try {
        $image=[Drawing.Image]::FromFile((Join-Path $photoFolder $name))
        try { $exampleImage.Image=[Drawing.Bitmap]::new($image) } finally { $image.Dispose() }
    }catch{}
}

function Get-GameInput {
    @{date=$date.Text.Trim();sport=$sport.Text.Trim();season=$season.Text.Trim()
        sides=@(@{school=$school1.Text.Trim();uniform=$uniform1.Text.Trim();squad=$squad1.Text.Trim()},
                @{school=$school2.Text.Trim();uniform=$uniform2.Text.Trim();squad=$squad2.Text.Trim()})}
}

function Refresh-Actions {
    $ready=-not $script:job
    $hasFolder=[bool]$folderField.Text
    $inspected=[bool]$script:session
    $choose.Enabled=$ready
    $inspect.Enabled=$ready -and $hasFolder
    $open.Enabled=$ready -and $hasFolder
    $undo.Enabled=$ready -and $hasFolder
    $tag.Enabled=$ready -and $inspected
    $more.Enabled=$ready -and $inspected
    foreach($control in $date,$sport,$season,$scope,$school1,$school2,$uniform1,$uniform2,$squad1,$squad2,$find1,$find2,$import1,$import2){
        $control.Enabled=$ready -and $inspected
    }
}

function Update-Review {
    if(-not $script:session -or $script:populating){return}
    $game=Get-GameInput
    $rosterFields=@($roster1,$roster2)
    $rosterSport=Resolve-RosterSport $game.sport $vocabulary
    $rosterSeason=Resolve-RosterSeason $game.season $game.date
    $rosters=@(foreach($file in (Get-ChildItem -LiteralPath (Join-Path $references 'Rosters') -File -Filter '*.json')){
        ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($file.FullName))
    })
    for($i=0;$i -lt 2;$i++){
        $side=$game.sides[$i]
        $canonical=@($schools | Where-Object { $_.name -eq $side.school -or $side.school -in $_.aliases } | Select-Object -ExpandProperty name)
        $key=if($canonical.Count -eq 1){$canonical[0]}else{$side.school}
        $matches=@($rosters | Where-Object { $_.school -eq $key -and $_.sport -eq $rosterSport -and $_.season -eq $rosterSeason -and $_.squad -eq $side.squad })
        $label=if(-not $key){'School unresolved'}elseif(-not $side.squad){'Choose the squad to check for a roster'}
            elseif($rosterSeason -notmatch '^(19|20)\d{2}'){'Enter the roster season year'}elseif($matches.Count -eq 1){'Roster found - ' + $matches[0].source}elseif($matches.Count -gt 1){'Several rosters match'}else{'Roster unresolved'}
        $rosterFields[$i].Text="Side $($i+1)  |  $label"
    }
    if($script:showingReport){return}
    $lines=@($script:inspectionText)
    $lines+='Sample keywords'
    $examples=@($script:session.examples | Where-Object {$_})
    foreach($example in $examples){
        $sportTag=if($example.activity -cin $vocabulary.activities){$example.activity}else{$game.sport}
        $words=@($sportTag)
        if($example.side1 -and $game.sides[0].school){$words+=$game.sides[0].school}
        if($example.side2 -and $game.sides[1].school){$words+=$game.sides[1].school}
        $lines+="  $($example.file)  |  $($words -join ', ')"
    }
    if(-not $examples.Count){$lines+="  $($script:session.photos[0].file)  |  $($game.sport)"}
    if($script:session.mixed){
        $lines+=''; $lines+='Numbered photos for choosing the game range'
        $lines+=@($script:session.photos | ForEach-Object {"  $($_.number). $($_.captured)  $($_.file)"})
    }
    $details.Lines=[string[]]$lines
}

function Confirm-Game {
    $game=Get-GameInput
    $range=$scope.Text.Trim()
    if(-not $game.sport){Notice 'Enter the sport before tagging.';return}
    if($game.date){
        $parsed=[datetime]::MinValue
        if(-not [datetime]::TryParseExact($game.date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)){
            Notice 'Enter the date as YYYY-MM-DD, or leave it blank.';return
        }
    }
    if($range -eq 'All'){
        $rows=@($script:session.photos)
    }elseif($range -match '^([1-9][0-9]{0,7})-([1-9][0-9]{0,7})$'){
        $first=[int]$Matches[1];$last=[int]$Matches[2]
        $rows=@($script:session.photos | Where-Object { $_.number -ge $first -and $_.number -le $last })
        if($last -gt $script:session.photos.Count){$rows=@()}
    }else{$rows=@()}
    if(-not $rows.Count){Notice 'Choose All or a photo range such as 1-20.';return}
    Start-Work 'Tag' @{game=$game;scope=$rows;examples=@($script:session.examples);mixed=$script:session.mixed
        evidence=$script:session.evidence;proposed_game=$script:session.game;sample=@($script:session.sample)}
}

function Get-RosterRequest([int]$Index) {
    $game=Get-GameInput
    $side=$game.sides[$Index]
    $rosterSport=Resolve-RosterSport $game.sport $vocabulary
    $rosterSeason=Resolve-RosterSeason $game.season $game.date
    if(-not $side.school -or -not $rosterSport -or $rosterSeason -notmatch '^(19|20)\d{2}(?:-(?:\d{2}|(?:19|20)\d{2}))?$' -or -not $side.squad){
        Notice 'Enter the school, sport, season year (such as 2026), and squad first.'
        return
    }
    @{school=$side.school;sport=$rosterSport;season=$rosterSeason;squad=$side.squad;date=$game.date}
}

function Find-Roster([int]$Index) {
    $request=Get-RosterRequest $Index
    if($request){Start-Work 'FindRoster' $request}
}

function Choose-Roster([int]$Index) {
    $request=Get-RosterRequest $Index
    if(-not $request){return}
    $dialog=[Windows.Forms.OpenFileDialog]::new()
    $dialog.Filter='Roster CSV (*.csv)|*.csv'
    $dialog.Title='Choose a roster with Number and Name columns'
    try {
        if($dialog.ShowDialog($form) -eq 'OK'){
            $label="$($request.school) $($request.sport) $($request.season) $($request.squad)"
            if([Windows.Forms.MessageBox]::Show($form,"Use this CSV as the shared roster for $label?",'Upload CSV','YesNo','Question','Button2') -eq 'Yes'){
                $request.csv=$dialog.FileName
                Start-Work 'ImportRoster' $request
            }
        }
    }finally{$dialog.Dispose()}
}

function Show-RosterCandidate($Roster) {
    $review=[Windows.Forms.Form]::new()
    $review.Text='Review Online Roster'
    $review.Font=$form.Font
    $review.ClientSize=[Drawing.Size]::new(720,490)
    $review.StartPosition='CenterParent'
    $review.FormBorderStyle='FixedDialog'
    $review.MaximizeBox=$false
    $review.MinimizeBox=$false
    $target=$Roster.selection
    $heading=[Windows.Forms.Label]::new()
    $heading.Text="$($target.school)  |  $($target.sport)  |  $($target.season)  |  $($target.squad)"
    $heading.SetBounds(18,15,680,25)
    $review.Controls.Add($heading)
    $source=[Windows.Forms.LinkLabel]::new()
    $source.Text="$($Roster.source)  |  $($Roster.source_url)"
    $source.AutoEllipsis=$true
    $source.SetBounds(18,44,680,27)
    $source.Add_LinkClicked({
        $uri=$null
        if([Uri]::TryCreate($Roster.source_url,[UriKind]::Absolute,[ref]$uri) -and $uri.Scheme -eq 'https'){
            Start-Process -FilePath $uri.AbsoluteUri
        }
    }.GetNewClosure())
    $review.Controls.Add($source)
    $notes=[Windows.Forms.TextBox]::new()
    $notes.Multiline=$true
    $notes.ReadOnly=$true
    $notes.ScrollBars='Vertical'
    $notes.SetBounds(18,75,680,65)
    $notes.Text=[string]$Roster.notes
    if($Roster.conflicts.Count){
        $notes.AppendText([Environment]::NewLine+"Duplicate jersey numbers need review: $($Roster.conflicts -join ', ').")
    }
    $review.Controls.Add($notes)
    $players=[Windows.Forms.ListView]::new()
    $players.View='Details'
    $players.FullRowSelect=$true
    $players.GridLines=$true
    $players.SetBounds(18,149,680,282)
    [void]$players.Columns.Add('Number',90)
    [void]$players.Columns.Add('Player',558)
    foreach($player in ($Roster.players | Sort-Object { [int]$_.number }, name)){
        $item=[Windows.Forms.ListViewItem]::new([string]$player.number)
        [void]$item.SubItems.Add([string]$player.name)
        [void]$players.Items.Add($item)
    }
    $review.Controls.Add($players)
    $cancel=[Windows.Forms.Button]::new()
    $cancel.Text='Cancel'
    $cancel.DialogResult='Cancel'
    $cancel.SetBounds(468,442,100,35)
    $review.Controls.Add($cancel)
    $save=[Windows.Forms.Button]::new()
    $save.Text='Save Roster'
    $save.DialogResult='OK'
    $save.SetBounds(579,442,119,35)
    $review.Controls.Add($save)
    $review.CancelButton=$cancel
    try { return $review.ShowDialog($form) -eq 'OK' }
    finally { $review.Dispose() }
}

function Start-Work($kind,$request) {
    if($script:job){return}
    try {
        $script:requestFile=Join-Path ([IO.Path]::GetTempPath()) ('SportsPhotoTagger-'+[guid]::NewGuid().ToString('N')+'.json')
        [IO.File]::WriteAllText($script:requestFile,(ConvertTo-Json -InputObject $request -Depth 12),[Text.UTF8Encoding]::new($false))
        $script:operation=$kind
        $script:log=@()
        $status.Text='Working...'
        $script:job=Start-Job -ArgumentList $engine,$kind,$folderField.Text,$script:requestFile -ScriptBlock {
            param($scriptPath,$operation,$photoFolder,$requestPath)
            $ErrorActionPreference='Stop'
            & $scriptPath -GuidedAction $operation -Folder $photoFolder -RequestFile $requestPath
        }
        Refresh-Actions
        $timer.Start()
    }catch{Notice $_.Exception.Message}
}

function Show-Inspection($result) {
    $script:populating=$true
    try {
        $script:session=$result
        $script:showingReport=$false
        $date.Text=[string]$result.game.date
        $sport.Text=[string]$result.game.sport
        $season.Text=[string]$result.game.season
        $school1.Text=[string]$result.game.sides[0].school
        $school2.Text=[string]$result.game.sides[1].school
        $uniform1.Text=[string]$result.game.sides[0].uniform
        $uniform2.Text=[string]$result.game.sides[1].uniform
        $squad1.Text=[string]$result.game.sides[0].squad
        $squad2.Text=[string]$result.game.sides[1].squad
        $exampleChoice.Items.Clear()
        $examples=@($result.examples | Where-Object {$_} | ForEach-Object {$_.file})
        if(-not $examples.Count){$examples=@($result.photos[0].file)}
        foreach($name in $examples){[void]$exampleChoice.Items.Add($name)}
        if($exampleChoice.Items.Count){$exampleChoice.SelectedIndex=0}
        $scope.Items.Clear()
        [void]$scope.Items.Add('All')
        foreach($group in $result.groups){[void]$scope.Items.Add($group)}
        $scope.Text=if($result.selected_range){$result.selected_range}elseif($result.mixed){''}else{'All'}
        $script:inspectionText=(@($result.notes | Where-Object {$_})+@($result.evidence | Where-Object {$_})) -join [Environment]::NewLine
    }finally{$script:populating=$false}
    Update-Review
    $status.Text='Check the game details and sample keywords before tagging.'
}

$timer.Add_Tick({
    if(-not $script:job){return}
    $output=@(Receive-Job -Job $script:job -ErrorAction SilentlyContinue *>&1)
    if($output.Count){
        $script:log+=@($output | ForEach-Object {$_.ToString()})
        $status.Text=$script:log[-1]
    }
    if($script:job.State -notin 'Completed','Failed','Stopped'){return}
    $timer.Stop()
    $answerPath="$script:requestFile.result"
    $candidate=$null
    try {
        if($script:job.State -ne 'Completed' -or -not [IO.File]::Exists($answerPath)){
            throw (($script:log | Select-Object -Last 5) -join [Environment]::NewLine)
        }
        $result=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($answerPath))
        switch($script:operation){
            'Inspect' {Show-Inspection $result}
            'Tag' {
                $script:showingReport=$true
                $lines=@('Manual review')+@($result.review.game)
                $lines+=@($result.review.photos | ForEach-Object {"$($_.file)  |  $($_.reasons -join ' / ')"})
                if($lines.Count -eq 1){$lines+='No open items.'}
                $details.Lines=[string[]]$lines
                foreach($change in $result.files){
                    foreach($photo in ($script:session.photos | Where-Object file -eq $change.file)){ $photo.sha256=$change.after }
                }
                $status.Text="Tagged $($result.count) JPEGs. In Photo Mechanic, press / to rescan and I to check Keywords."
            }
            'Undo' {
                $script:session=$null
                $details.Clear()
            $exampleChoice.Items.Clear()
            Show-Example
                $status.Text=if($result.count){"Restored $($result.count) JPEGs. Inspect again to start another run."}else{'There is no tagging run to undo.'}
            }
            {$_ -in 'ImportRoster','SaveRoster'} {$status.Text="Saved $($result.count) roster names."; Update-Review}
            'FindRoster' {
                if($result.found){
                    $status.Text='Review the source and numbered players.'
                    $candidate=$result
                } else {
                    $status.Text='Roster unresolved. You can upload a CSV.'
                    $reason=if($result.notes){[Environment]::NewLine+[string]$result.notes}else{''}
                    Notice "The search found no usable roster for this school, season, and squad.$reason"
                }
            }
        }
    }catch {
        $status.Text='Action stopped. See the details below.'
        $details.Text=$_.Exception.Message
    }finally{
        Remove-Job -Job $script:job
        $script:job=$null
        foreach($path in $script:requestFile,$answerPath){if([IO.File]::Exists($path)){[IO.File]::Delete($path)}}
        Refresh-Actions
    }
    if($candidate){
        if(Show-RosterCandidate $candidate){ Start-Work 'SaveRoster' $candidate }
        else { $status.Text='Roster review closed. The current roster is still in use.' }
    }
})
foreach($control in $date,$sport,$season,$school1,$school2,$squad1,$squad2){
    $control.Add_TextChanged({ Update-Review })
}
$form.Add_FormClosing({param($sender,$eventArgs)
    if($script:job){$eventArgs.Cancel=$true;Notice 'Wait for the current operation to finish before closing.'}
})
Refresh-Actions
if($SmokeTest){ Write-Output 'SportsPhotoTagger window initialized.'; $form.Dispose(); return }
[void]$form.ShowDialog()
if($exampleImage.Image){$exampleImage.Image.Dispose()}
$form.Dispose()

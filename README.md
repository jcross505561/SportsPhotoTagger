# SportsPhotoTagger

SportsPhotoTagger is a Windows desktop tool for reviewing and writing keywords to sports JPEGs. It can use a roster to match a jersey number to a player name, and you can check the written keywords in Photo Mechanic.

## Run the app

1. Clone or download this repository and open **SportsPhotoTagger.cmd** on Windows. The launcher uses Windows PowerShell.
2. Install Codex CLI and sign in with `codex.cmd login` to enable **Inspect** and **Find Online**. Inspect sends resized photo previews to Codex for analysis.
3. Choose an existing game folder and click **Inspect**. Review the date, sport, teams, and sample keywords. If the folder contains several games, select the photo range you intend to tag.
4. Click **Confirm & Tag** when the preview looks right. The app lists photos that need manual review. **Undo Last Run** uses the saved backup for the selected game.
5. In Photo Mechanic, press **/** to rescan and **I** to check a photo's Keywords field.

## Use a roster

Enter the school, sport, season, and squad for the relevant side. Select **Upload CSV** to import a roster or **Find Online** to review a suggested source before choosing **Save Roster**. You can type school names directly into the app.

This fictional CSV shows the required columns. Save your own roster CSV outside the checkout.

```csv
Number,Name
7,Avery Example
22,Jordan Sample
```

Imported and saved rosters stay in the checkout's ignored **Rosters** folder. Backups and review records stay in the selected game folder. Keep game photos outside the checkout and use the folder picker to open them. The repository ships with an empty team catalog and no roster records.

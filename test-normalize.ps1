$testFiles = @(
    'Holiday/Various Artists/A Rhythm & Blues Christmas - Vol. 1/(02) Run Rudolph Run - Chuck Berry.mp3',
    'Holiday/Various Artists/Have a Fun Christmas Vol. 1/03 - Run Rudolph Run - Chuck Berry.mp3',
    'Holiday/Various Artists/Rock Christmas Vol. 8/21 - Bryan Adams - Run Rudolph Run.mp3'
)

. .\Shuffle-M3UPlaylist.ps1

foreach ($file in $testFiles) {
    $normalized = Get-NormalizedSongTitle $file
    Write-Host "$file -> $normalized"
}

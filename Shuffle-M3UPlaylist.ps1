<#
.SYNOPSIS
Shuffles a CSV playlist while keeping different versions of the same song spread apart.

.DESCRIPTION
This script reads a CSV playlist file and creates a shuffled version that attempts to
avoid playing different versions of the same song close together. For example, if the
playlist has "Jingle Bells" by 5 different artists, this script ensures they won't
play consecutively or too close to each other.

The CSV file should have columns: Filename, Weight, Missing

The script:
1. Extracts normalized song titles from filenames (removing artist info, version info, etc.)
2. Groups songs by similarity
3. Shuffles using an algorithm that distributes similar songs throughout the playlist
4. Saves the result as a new M3U file

.PARAMETER InputCSV
Path to the input CSV playlist file.

.PARAMETER OutputM3U
Path to save the shuffled M3U playlist. If not specified, defaults to the input filename
with "_shuffled.m3u" appended.

.PARAMETER MinimumDistance
Minimum number of songs that should appear between different versions of the same song.
Default is 5. Higher values spread similar songs further apart.

.PARAMETER Seed
Random seed for reproducible shuffles. If not specified, uses a random seed.

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "X:\Holiday\Playlists\Christmas.csv"

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "X:\Holiday\Playlists\Christmas.csv" -OutputM3U "X:\Holiday\Playlists\Christmas_Shuffled.m3u" -MinimumDistance 8

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "Christmas.csv" -Seed 42
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $InputCSV,

    [Parameter(Mandatory = $false)]
    [string] $OutputM3U,

    [Parameter(Mandatory = $false)]
    [string[]] $GrepGroup,

    [Parameter(Mandatory = $false)]
    [string[]] $DebugParts,

    [Parameter(Mandatory = $false)]
    [int] $MinimumDistance = 5,

    [Parameter(Mandatory = $false)]
    [int] $ShowGroupCount = 20,

    [Parameter(Mandatory = $false)]
    [int] $Seed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Set random seed if provided
if ($PSBoundParameters.ContainsKey('Seed'))
{
    Get-Random -SetSeed $Seed
}

function Read-CSVFile
{
    param([string] $Path)
    
    if (!(Test-Path -LiteralPath $Path))
    {
        throw "CSV file not found: $Path"
    }

    $csv = Import-Csv -LiteralPath $Path
    $tracks = @()
    
    foreach ($row in $csv)
    {
        if (![string]::IsNullOrWhiteSpace($row.Filename))
        {
            $tracks += [PSCustomObject]@{
                Filename = $row.Filename
                Weight = $row.Weight
                Missing = $row.Missing
            }
        }
    }

    if ($tracks.Count -eq 0)
    {
        throw "No tracks found in CSV file: $Path"
    }

    Write-Host "Loaded $($tracks.Count) tracks from $Path" -ForegroundColor Cyan
    return $tracks
}

function NormalizePart
{
    param([string] $Part)
    
    # Remove trailing artist info after delimiters like "by", "ft.", "feat."
    $Part = $Part -replace '\s+by\s+.*$', ''
    $Part = $Part -replace '\s+ft[\.\s]+.*$', ''
    $Part = $Part -replace '\s+feat[\.\s]+.*$', ''
    
    # - Parenthetical info (e.g., "(Live)", "(Remix)", "(Remastered)")
    $Part = $Part -replace '\s*\([^\)]*\)\s*', ''
    $Part = $Part -replace '\s*\[[^\]]*\]\s*', ''

    # - leading track numbers
    $Part = $Part -replace ' +\d+[\s\.\-]+', ''
    $Part = $Part -replace ' +\(\d+\) *', ''

    # - Special characters and extra whitespace
    $Part = $Part -replace '[_\-:]+', ' '
    $Part = $Part -replace '\s+', ' '
    $Part = $Part.Trim()
    
    # Convert to lowercase for case-insensitive comparison
    return $Part.ToLower()
}


function Get-PathContext
{
    param([string] $FilePath)
    
    # Extract potential artist and album names from the directory structure
    # Common patterns: ...\Artist\Album\Song.mp3 or ...\Album\Song.mp3
    $directory = [System.IO.Path]::GetDirectoryName($FilePath)
    $pathParts = $directory -split '[/\\]'
    
    # Get the last 2-3 directory levels (might contain artist, album, compilation info)
    $context = @{
        Artist   = $null
        Album    = $null
        Keywords = @()
    }
    
    if ($pathParts.Count -ge 1)
    {
        $lastDir = $pathParts[-1]
        $context.Album = $lastDir
        $context.Keywords += $lastDir
    }
    
    if ($pathParts.Count -ge 2)
    {
        $secondLastDir = $pathParts[-2]
        $context.Artist = $secondLastDir
        $context.Keywords += $secondLastDir
    }
    
    return $context
}

function partMatchesSongPattern
{
    param([string] $part)

    # Common song starting words that indicate it's likely a song title
    $matchValues = '^(if|my|let''s|the|a|an|my|your|i|we|let|don''t|it''s|you''re|all|happy|merry|jingle|deck|winter|white|blue|silent|holy|little|santa|christmas|feliz|merry|las)\s+'

    if ($part -match $matchValues)
    {
        return $true
    }
    return $false
}

function shouldChoosePart
{
    param([string] $part, [string] $partOther)

    if (partMatchesSongPattern $part)
    {
        return $true
    }

    # if we are longer than the other part, AND the other part is not a common song starting word, prefer this part
    if ($part.Length -gt $partOther.Length -and -not (partMatchesSongPattern $partOther))
    {
        return $true
    }
    return $false
}

function chooseWinnerFromTwoParts
{
    param([string] $firstPart, [string] $lastPart, [bool] $debugThis)

    if (shouldChoosePart $firstPart $lastPart)
    {
        if ($debugThis)
        {
            Write-Host "      Choosing between parts $($firstPart)/$($lastPart): chose $($firstPart)"
        }
        return $firstPart
    }
    elseif (shouldChoosePart $lastPart $firstPart)
    {
        if ($debugThis)
        {
            Write-Host "      Choosing between parts $($firstPart)/$($lastPart): chose $($lastPart)"
        }
        return $lastPart
    }
    # Default: prefer last part (most common pattern: "Artist - Song")
    else
    {
        if ($debugThis)
        {
            Write-Host "      Choosing between parts $($firstPart)/$($lastPart): chose by default $($lastPart)"
        }
        return $lastPart
    }
}

function Get-NormalizedSongTitle
{
    param([string] $FilePath,
        [hashtable] $groups,
        [string[]] $DebugParts)

    $debugThis = $false

    if ($DebugParts -and ($DebugParts | Where-Object { $FilePath -match $_ }))
    {
        $debugThis = $true;
    }

    # Extract filename without extension
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    
    # Get context from path (artist/album names)
    $pathContext = Get-PathContext $FilePath
    
    # Remove common patterns:
    # - Track numbers (e.g., "01 - ", "01. ", "01-", "1. ", "01 ")
    $normalized = $filename -replace '^\d+[\s\.\-]+', ''
    $normalized = $normalized -replace '^\(\d+\) *', ''

    # get rid of multiple spaces and collapse into 1 space
    $normalized = $normalized -replace '  +', ''

    $normalized = $normalized -replace '^\d+$', ''  # Just a number with nothing after
    
    # Remove disc/CD numbers (e.g., "Disc 1 - ", "CD2-", "(Disc 1)")
    $normalized = $normalized -replace '^(Disc|CD)\s*\d+[\s\.\-]+', ''
    $normalized = $normalized -replace '\(Disc\s*\d+\)', ''
    
    # Parse multi-part filenames (e.g., "Artist - Album - Track - Song Title - Extra Info")
    # Split by " - " and try to identify which part is the song title
    $parts = @($normalized -split '\s+-\s+')
    
    if ($debugThis)
    {
        Write-Host "here! filepath: $($FilePath), filename: $($filename), normalized so far: $($normalized), Parts: '$($parts -join ', ')'"
    }

    if ($parts.Count -gt 1)
    {
        # Common patterns:
        # - "Artist - Song" -> use last part
        # - "Artist - Album - Song" -> use last part  
        # - "Artist - Album - TrackNum - Song" -> use last part (tracknum already removed)
        # - "Album (Disc X) - TrackNum - Song - Artist" -> prefer middle parts
        # - "Album - Song - Artist" -> use middle part
        
        # Filter out parts that match path context (artist/album)
        $candidateParts = @()
        $candidateIndices = @()
        
        for ($i = 0; $i -lt $parts.Count; $i++)
        {
            $part = $parts[$i].Trim()
            if ($debugThis)
            {
                Write-Host "      Processing part: $($part)"
            }
            $part = $part -replace '^\d+ *', ''

            $isContext = $false
            
            if ($debugThis)
            {
                Write-Host "      Processing part: $($part)"
            }
            
            # Skip empty parts
            if ([string]::IsNullOrWhiteSpace($part))
            {
                if ($debugThis)
                {
                    Write-Host "      Skipping empty part: $($part)"
                }
                continue
            }
            
            # Skip parts that are just numbers (track numbers, years, etc.)
            if ($part -match '^\d+$')
            {
                if ($debugThis)
                {
                    Write-Host "      Skipping numeric part: $($part)"
                }
                continue
            }
            
            # Skip very short parts (likely not song titles)
            if ($part.Length -lt 2)
            {
                if ($debugThis)
                {
                    Write-Host "      Skipping short part: $($part)"
                }
                continue
            }
            
            # Skip parts that look like album disc references
            if ($part -match '(?i)(disc|cd)\s*\d+')
            {
                if ($debugThis)
                {
                    Write-Host "      Skipping disc reference part: $($part)"
                }
                continue
            }
  
            # don't skip parts that strongly match a song pattern
            if (-not (partMatchesSongPattern $part))
            {
                if ($debugThis)
                {
                    Write-Host "      Checking context for part: $($part)"
                }

                # Check if this part matches path context
                foreach ($keyword in $pathContext.Keywords)
                {
                    if (![string]::IsNullOrWhiteSpace($keyword))
                    {
                        # Check if this part matches the path context (case-insensitive, allowing for variations)
                        # Match if the part contains the keyword or vice versa
                        if ($part -match "(?i)$([regex]::Escape($keyword))" -or $keyword -match "(?i)$([regex]::Escape($part))")
                        {
                            if ($debugThis)
                            {
                                Write-Host "      Skipping context part: $($part) (matches keyword: $($keyword))"
                            }
                        
                            $isContext = $true
                            break
                        }
                    }
                }
            }        
            if (-not $isContext)
            {
                $candidateParts += $part
                $candidateIndices += $i
            }
        }
        
        # Choose the best candidate:
        # Common patterns after filtering out context:
        # - ["Artist", "Song"] -> use last (Song)
        # - ["Song", "Artist"] -> use first (Song) 
        # - ["Song"] -> use it
        # - ["Track", "Song", "Artist"] -> prefer middle (Song)
        if ($candidateParts.Count -gt 0)
        {
            if ($candidateParts.Count -eq 1)
            {
                if ($debugThis)
                {
                    Write-Host "      Chose the only part: $($candidateParts[0])"
                }
                $normalized = $candidateParts[0]
            }
            else
            {
                $partIndex = 0
                $lastPartMatched = -1
                $lastPartWeight = -1

                while ($partIndex -lt $candidateParts.Count)
                {
                    $part = $candidateParts[$partIndex]
                    $normalizedPart = NormalizePart($part)
                    if ($groups.ContainsKey($normalizedPart))
                    {
                        $weight = $groups[$normalizedPart].Count
                        if ($weight -gt $lastPartWeight)
                        {
                            $lastPartWeight = $weight
                            $lastPartMatched = $partIndex
                        }
                    }

                    $partIndex++
                }
                if ($lastPartMatched -ne -1)
                {
                    $normalized = $candidateParts[$lastPartMatched]
                    if ($debugThis)
                    {
                        Write-Host "      Chose known part from groups: $($normalized) with weight $($lastPartWeight)"
                    }
                    return $normalized
                }

                # No known song matched, use heuristics below
                if ($candidateParts.Count -eq 2)
                {
                    # With 2 candidates, use heuristics:
                    # - Prefer the part that looks more like a song title
                    # - Song titles often have: "The", "A", "An", common words, are longer
                    # - Artist names often have: proper capitalization, shorter, person names
                    $normalized = chooseWinnerFromTwoParts $candidateParts[0] $candidateParts[1] $debugThis
                }
                else
                {
                    # binary search for a winner

                    $index = 0
                    $length = 1

                    # while we have pairs to compare
                    while ($index + $length -lt $candidateParts.Count)
                    {
                        while ($index -lt $candidateParts.Count - $length)
                        {
                            # if we don't have one to compare with, then we win by default (which means we don't change)
                            if ($index + $length -lt $candidateParts.Count)
                            {
                                $candidateParts[$index] = chooseWinnerFromTwoParts $candidateParts[$index] $candidateParts[$index + $length] $debugThis
                            }
                            $index += $length
                        }
                        $length *= 2
                    }

                    # after reducing, our answer is in candidateParts[0]
                    $normalized = $candidateParts[0]

                    if ($debugThis)
                    {
                        Write-Host "      Chose winner in multi-candidate: $($normalized). CandidateParts: '$($candidateParts -join ', ')', chose: $($normalized)"
                    }
                }
            }
        }
        else
        {
            # No non-context parts found, use the last part
            if ($debugThis)
            {
                Write-Host "      No non-context parts found, using whole parts: $($parts[-1])"
            }
            
            $normalized = $parts[-1]
        }
    }
    if ($debugThis)
    {
        Write-Host "       Final result!  CandidateParts: '$($candidateParts -join ', ')', chose: $($normalized)"
    }
    return NormalizePart($normalized)
}

function Group-TracksBySimilarity
{
    param([string[]] $Tracks, [string[]] $GrepGroup, [string[]] $DebugParts, [int] $ShowGroupCount)
    
    # first pass just collects groups for the next pass
    $groupsPass1 = @{}
    $groups = @{}
    $trackInfo = @()
    
    # first pass just collects groups for the next pass
    
    Write-Host "`nBuilding similarity groups..." -ForegroundColor Cyan

    foreach ($track in $Tracks)
    {
        # yes, pass in groups here -- we always want the empty hash table
        $normalized = Get-NormalizedSongTitle $track $groups $DebugParts
        
        if (!$groupsPass1.ContainsKey($normalized))
        {
            $groupsPass1[$normalized] = @()
        }
        $groupsPass1[$normalized] += $track
    }

    # second pass builds the actual track info
    Write-Host "`nBuilding track info with gathered group information..." -ForegroundColor Cyan

    foreach ($track in $Tracks)
    {
        $normalized = Get-NormalizedSongTitle $track $groupsPass1 $DebugParts
        
        if (!$groups.ContainsKey($normalized))
        {
            $groups[$normalized] = @()
        }
        $groups[$normalized] += $track
        
        $trackInfo += [PSCustomObject]@{
            Path            = $track
            NormalizedTitle = $normalized
        }
    }
    
    
    #return;

    # Report duplicates/variants
    $variantGroups = @($groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 } | Sort-Object { $_.Value.Count } -Descending)
    
    if ($variantGroups)
    {
        Write-Host "`nFound $($variantGroups.Count) songs with multiple versions:" -ForegroundColor Yellow
        $displayCount = [Math]::Min($ShowGroupCount, $variantGroups.Count)
        foreach ($group in ($variantGroups | Select-Object -First $displayCount))
        {
            if ($GrepGroup -and ($group.Key -notin $GrepGroup))
            {
                continue
            }

            Write-Host "  '$($group.Key)' - $($group.Value.Count) versions" -ForegroundColor DarkYellow
            # Show a few examples of the actual filenames for the top groups
            if ($group.Value.Count -gt 1)
            {
                $exampleCount = [Math]::Min(3, $group.Value.Count)
                for ($i = 0; $i -lt $exampleCount; $i++)
                {
                    $exampleFile = [System.IO.Path]::GetFileName($group.Value[$i])
                    Write-Host "    e.g.: $exampleFile" -ForegroundColor DarkGray
                }
            }
        }
        if ($variantGroups.Count -gt 20)
        {
            Write-Host "  ... and $($variantGroups.Count - 20) more" -ForegroundColor DarkGray
        }
    }
    
    return , $trackInfo
}

function Invoke-SmartShuffle
{
    param(
        [PSCustomObject[]] $TrackInfo,
        [int] $MinDistance
    )
    
    Write-Host "`nShuffling with minimum distance of $MinDistance between similar songs..." -ForegroundColor Cyan
    
    # Create a list of available tracks (not yet placed)
    $available = New-Object System.Collections.Generic.List[PSCustomObject]
    $available.AddRange($TrackInfo)
    
    # Result list
    $shuffled = New-Object System.Collections.Generic.List[PSCustomObject]
    
    # Track when each normalized title was last used (index in shuffled list)
    $lastUsed = @{}
    
    $attempts = 0
    $maxAttemptsPerTrack = 100
    
    while ($available.Count -gt 0)
    {
        $placed = $false
        $attempts++
        
        # Shuffle the available list to randomize selection
        $availableArray = $available.ToArray()
        $availableArray = $availableArray | Sort-Object { Get-Random }
        
        # Try to find a track that respects the minimum distance
        foreach ($track in $availableArray)
        {
            $canPlace = $true
            
            # Check if this normalized title was used recently
            if ($lastUsed.ContainsKey($track.NormalizedTitle))
            {
                $distance = $shuffled.Count - $lastUsed[$track.NormalizedTitle]
                if ($distance -lt $MinDistance)
                {
                    $canPlace = $false
                }
            }
            
            if ($canPlace)
            {
                # Place this track
                [void]$shuffled.Add($track)
                $lastUsed[$track.NormalizedTitle] = $shuffled.Count - 1
                [void]$available.Remove($track)
                $placed = $true
                $attempts = 0
                break
            }
        }
        
        # If we couldn't place any track, relax the constraint temporarily
        if (!$placed)
        {
            if ($attempts -gt $maxAttemptsPerTrack)
            {
                Write-Warning "Could not maintain minimum distance for all songs. Placing next available track."
                $track = $available[0]
                [void]$shuffled.Add($track)
                $lastUsed[$track.NormalizedTitle] = $shuffled.Count - 1
                [void]$available.RemoveAt(0)
                $attempts = 0
            }
        }
    }
    
    Write-Host "Shuffle complete. Generated playlist with $($shuffled.Count) tracks." -ForegroundColor Green
    # Use comma operator to prevent array unrolling
    return , $shuffled
}

function Test-ShuffleQuality
{
    param(
        [array] $ShuffledTracks,
        [int] $MinDistance
    )
    
    Write-Host "`nAnalyzing shuffle quality..." -ForegroundColor Cyan
    
    $violations = 0
    $minDistanceFound = [int]::MaxValue
    
    for ($i = 0; $i -lt $ShuffledTracks.Count; $i++)
    {
        $currentTrack = $ShuffledTracks[$i]
        $currentTitle = $currentTrack.NormalizedTitle
        
        # Look ahead for the same normalized title
        for ($j = $i + 1; $j -lt [Math]::Min($i + $MinDistance, $ShuffledTracks.Count); $j++)
        {
            $compareTrack = $ShuffledTracks[$j]
            if ($compareTrack.NormalizedTitle -eq $currentTitle)
            {
                $distance = $j - $i
                $violations++
                if ($distance -lt $minDistanceFound)
                {
                    $minDistanceFound = $distance
                }
                Write-Warning "Similar songs at positions $($i+1) and $($j+1) (distance: $distance) - '$currentTitle'"
            }
        }
    }
    
    if ($violations -eq 0)
    {
        Write-Host "Perfect shuffle! No similar songs within $MinDistance tracks of each other." -ForegroundColor Green
    }
    else
    {
        Write-Host "Found $violations constraint violation(s). Minimum distance found: $minDistanceFound" -ForegroundColor Yellow
        Write-Host "Consider reducing -MinimumDistance or running again for a different random shuffle." -ForegroundColor Yellow
    }
}

function Write-M3UFile
{
    param(
        [string] $Path,
        [array] $ShuffledTracks
    )
    
    $lines = @()
    $lines += "#EXTM3U"
    
    foreach ($track in $ShuffledTracks)
    {
        $lines += $track.Path
    }
    
    # Ensure directory exists
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    if ($dir -and !(Test-Path -LiteralPath $dir))
    {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    
    $lines | Out-File -FilePath $Path -Encoding UTF8
    Write-Host "`nSaved shuffled playlist to: $Path" -ForegroundColor Green
}

# ----- Main Script -----

Write-Host "=== CSV Playlist Smart Shuffler ===" -ForegroundColor Cyan
Write-Host ""

# Determine output path
if ([string]::IsNullOrWhiteSpace($OutputM3U))
{
    $dir = [System.IO.Path]::GetDirectoryName($InputCSV)
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($InputCSV)
    if ([string]::IsNullOrWhiteSpace($dir))
    {
        $OutputM3U = "${nameWithoutExt}_shuffled.m3u"
    }
    else
    {
        $OutputM3U = Join-Path $dir "${nameWithoutExt}_shuffled.m3u"
    }
}

# Read input file
$trackObjects = Read-CSVFile -Path $InputCSV

# Extract just the filenames for processing (for now, ignore Weight and Missing)
$tracks = @($trackObjects | ForEach-Object { $_.Filename })

# Group tracks by similarity
$trackInfo = Group-TracksBySimilarity -Tracks $tracks -GrepGroup $GrepGroup -DebugParts $DebugParts -ShowGroupCount $ShowGroupCount

# Perform smart shuffle
$shuffledTracks = Invoke-SmartShuffle -TrackInfo $trackInfo -MinDistance $MinimumDistance

# Test shuffle quality
Test-ShuffleQuality -ShuffledTracks $shuffledTracks -MinDistance $MinimumDistance

# Write output file
Write-M3UFile -Path $OutputM3U -ShuffledTracks $shuffledTracks

Write-Host "`nDone!" -ForegroundColor Green

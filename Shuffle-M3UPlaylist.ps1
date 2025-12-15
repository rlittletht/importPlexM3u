<#
.SYNOPSIS
Shuffles an M3U playlist while keeping different versions of the same song spread apart.

.DESCRIPTION
This script reads an M3U playlist file and creates a shuffled version that attempts to
avoid playing different versions of the same song close together. For example, if the
playlist has "Jingle Bells" by 5 different artists, this script ensures they won't
play consecutively or too close to each other.

The script:
1. Extracts normalized song titles from filenames (removing artist info, version info, etc.)
2. Groups songs by similarity
3. Shuffles using an algorithm that distributes similar songs throughout the playlist
4. Saves the result as a new M3U file

.PARAMETER InputM3U
Path to the input M3U playlist file.

.PARAMETER OutputM3U
Path to save the shuffled M3U playlist. If not specified, defaults to the input filename
with "_shuffled" appended.

.PARAMETER MinimumDistance
Minimum number of songs that should appear between different versions of the same song.
Default is 5. Higher values spread similar songs further apart.

.PARAMETER Seed
Random seed for reproducible shuffles. If not specified, uses a random seed.

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputM3U "X:\Holiday\Playlists\Christmas.m3u"

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputM3U "X:\Holiday\Playlists\Christmas.m3u" -OutputM3U "X:\Holiday\Playlists\Christmas_Shuffled.m3u" -MinimumDistance 8

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputM3U "Christmas.m3u" -Seed 42
#>

param(
    [Parameter(Mandatory=$true)]
    [string] $InputM3U,

    [Parameter(Mandatory=$false)]
    [string] $OutputM3U,

    [Parameter(Mandatory=$false)]
    [int] $MinimumDistance = 5,

    [Parameter(Mandatory=$false)]
    [int] $Seed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Set random seed if provided
if ($PSBoundParameters.ContainsKey('Seed')) {
    Get-Random -SetSeed $Seed
}

function Read-M3UFile {
    param([string] $Path)
    
    if (!(Test-Path -LiteralPath $Path)) {
        throw "M3U file not found: $Path"
    }

    $lines = Get-Content -LiteralPath $Path -Encoding UTF8
    $tracks = @()
    
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed.Length -eq 0) { continue }
        if ($trimmed.StartsWith('#')) { continue }
        $tracks += $trimmed
    }

    if ($tracks.Count -eq 0) {
        throw "No tracks found in M3U file: $Path"
    }

    Write-Host "Loaded $($tracks.Count) tracks from $Path" -ForegroundColor Cyan
    return $tracks
}

function Get-PathContext {
    param([string] $FilePath)
    
    # Extract potential artist and album names from the directory structure
    # Common patterns: ...\Artist\Album\Song.mp3 or ...\Album\Song.mp3
    $directory = [System.IO.Path]::GetDirectoryName($FilePath)
    $pathParts = $directory -split '[/\\]'
    
    # Get the last 2-3 directory levels (might contain artist, album, compilation info)
    $context = @{
        Artist = $null
        Album = $null
        Keywords = @()
    }
    
    if ($pathParts.Count -ge 1) {
        $lastDir = $pathParts[-1]
        $context.Album = $lastDir
        $context.Keywords += $lastDir
    }
    
    if ($pathParts.Count -ge 2) {
        $secondLastDir = $pathParts[-2]
        $context.Artist = $secondLastDir
        $context.Keywords += $secondLastDir
    }
    
    # Also add parent directory if it looks relevant
    if ($pathParts.Count -ge 3) {
        $context.Keywords += $pathParts[-3]
    }
    
    return $context
}

function Get-NormalizedSongTitle {
    param([string] $FilePath)
    
    # Extract filename without extension
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    
    # Get context from path (artist/album names)
    $pathContext = Get-PathContext $FilePath
    
    # Remove common patterns:
    # - Track numbers (e.g., "01 - ", "01. ", "01-", "1. ", "01 ")
    $normalized = $filename -replace '^\d+[\s\.\-]+', ''
    $normalized = $normalized -replace '^\d+$', ''  # Just a number with nothing after
    
    # Remove disc/CD numbers (e.g., "Disc 1 - ", "CD2-", "(Disc 1)")
    $normalized = $normalized -replace '^(Disc|CD)\s*\d+[\s\.\-]+', ''
    $normalized = $normalized -replace '\(Disc\s*\d+\)', ''
    
    # Parse multi-part filenames (e.g., "Artist - Album - Track - Song Title - Extra Info")
    # Split by " - " and try to identify which part is the song title
    $parts = $normalized -split '\s+-\s+'
    
    if ($parts.Count -gt 1) {
        # Common patterns:
        # - "Artist - Song" -> use last part
        # - "Artist - Album - Song" -> use last part  
        # - "Artist - Album - TrackNum - Song" -> use last part (tracknum already removed)
        # - "Album (Disc X) - TrackNum - Song - Artist" -> prefer middle parts
        # - "Album - Song - Artist" -> use middle part
        
        # Filter out parts that match path context (artist/album)
        $candidateParts = @()
        $candidateIndices = @()
        
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $part = $parts[$i].Trim()
            $isContext = $false
            
            # Skip empty parts
            if ([string]::IsNullOrWhiteSpace($part)) {
                continue
            }
            
            # Skip parts that are just numbers (track numbers, years, etc.)
            if ($part -match '^\d+$') {
                continue
            }
            
            # Skip very short parts (likely not song titles)
            if ($part.Length -lt 2) {
                continue
            }
            
            # Skip parts that look like album disc references
            if ($part -match '(?i)(disc|cd)\s*\d+') {
                continue
            }
            
            # Check if this part matches path context
            foreach ($keyword in $pathContext.Keywords) {
                if (![string]::IsNullOrWhiteSpace($keyword)) {
                    # Check if this part matches the path context (case-insensitive, allowing for variations)
                    # Match if the part contains the keyword or vice versa
                    if ($part -match "(?i)$([regex]::Escape($keyword))" -or $keyword -match "(?i)$([regex]::Escape($part))") {
                        $isContext = $true
                        break
                    }
                }
            }
            
            if (-not $isContext) {
                $candidateParts += $part
                $candidateIndices += $i
            }
        }
        
        # Choose the best candidate:
        # 1. If we have candidates, prefer one in the middle (likely song title)
        # 2. Otherwise use the last non-context part
        if ($candidateParts.Count -gt 0) {
            # If there are multiple candidates, prefer the one closest to the middle
            # but not the last one (which is often artist name)
            if ($candidateParts.Count -eq 1) {
                $normalized = $candidateParts[0]
            } elseif ($candidateParts.Count -eq 2) {
                # With 2 candidates, prefer the first (song over artist)
                $normalized = $candidateParts[0]
            } else {
                # With 3+ candidates, take the middle one or second one
                $middleIndex = [Math]::Floor($candidateParts.Count / 2)
                if ($middleIndex -gt 0) {
                    $middleIndex = 1
                }
                $normalized = $candidateParts[$middleIndex]
            }
        } else {
            # No non-context parts found, use the last part
            $normalized = $parts[-1]
        }
    }
    
    # Remove trailing artist info after delimiters like "by", "ft.", "feat."
    $normalized = $normalized -replace '\s+by\s+.*$', ''
    $normalized = $normalized -replace '\s+ft[\.\s]+.*$', ''
    $normalized = $normalized -replace '\s+feat[\.\s]+.*$', ''
    
    # - Parenthetical info (e.g., "(Live)", "(Remix)", "(Remastered)")
    $normalized = $normalized -replace '\s*\([^\)]*\)\s*', ''
    $normalized = $normalized -replace '\s*\[[^\]]*\]\s*', ''
    
    # - Special characters and extra whitespace
    $normalized = $normalized -replace '[_\-:]+', ' '
    $normalized = $normalized -replace '\s+', ' '
    $normalized = $normalized.Trim()
    
    # Convert to lowercase for case-insensitive comparison
    return $normalized.ToLower()
}

function Group-TracksBySimilarity {
    param([string[]] $Tracks)
    
    $groups = @{}
    $trackInfo = @()
    
    foreach ($track in $Tracks) {
        $normalized = Get-NormalizedSongTitle $track
        
        if (!$groups.ContainsKey($normalized)) {
            $groups[$normalized] = @()
        }
        $groups[$normalized] += $track
        
        $trackInfo += [PSCustomObject]@{
            Path = $track
            NormalizedTitle = $normalized
        }
    }
    
    # Report duplicates/variants
    $variantGroups = $groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 } | Sort-Object { $_.Value.Count } -Descending
    
    if ($variantGroups) {
        Write-Host "`nFound $($variantGroups.Count) songs with multiple versions:" -ForegroundColor Yellow
        $displayCount = [Math]::Min(10, $variantGroups.Count)
        foreach ($group in ($variantGroups | Select-Object -First $displayCount)) {
            Write-Host "  '$($group.Key)' - $($group.Value.Count) versions" -ForegroundColor DarkYellow
            # Show a few examples of the actual filenames for the top groups
            if ($group.Value.Count -gt 5) {
                $exampleCount = [Math]::Min(3, $group.Value.Count)
                for ($i = 0; $i -lt $exampleCount; $i++) {
                    $exampleFile = [System.IO.Path]::GetFileName($group.Value[$i])
                    Write-Host "    e.g.: $exampleFile" -ForegroundColor DarkGray
                }
            }
        }
        if ($variantGroups.Count -gt 10) {
            Write-Host "  ... and $($variantGroups.Count - 10) more" -ForegroundColor DarkGray
        }
    }
    
    return ,$trackInfo
}

function Invoke-SmartShuffle {
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
    
    while ($available.Count -gt 0) {
        $placed = $false
        $attempts++
        
        # Shuffle the available list to randomize selection
        $availableArray = $available.ToArray()
        $availableArray = $availableArray | Sort-Object { Get-Random }
        
        # Try to find a track that respects the minimum distance
        foreach ($track in $availableArray) {
            $canPlace = $true
            
            # Check if this normalized title was used recently
            if ($lastUsed.ContainsKey($track.NormalizedTitle)) {
                $distance = $shuffled.Count - $lastUsed[$track.NormalizedTitle]
                if ($distance -lt $MinDistance) {
                    $canPlace = $false
                }
            }
            
            if ($canPlace) {
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
        if (!$placed) {
            if ($attempts -gt $maxAttemptsPerTrack) {
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
    return ,$shuffled
}

function Test-ShuffleQuality {
    param(
        [array] $ShuffledTracks,
        [int] $MinDistance
    )
    
    Write-Host "`nAnalyzing shuffle quality..." -ForegroundColor Cyan
    
    $violations = 0
    $minDistanceFound = [int]::MaxValue
    
    for ($i = 0; $i -lt $ShuffledTracks.Count; $i++) {
        $currentTrack = $ShuffledTracks[$i]
        $currentTitle = $currentTrack.NormalizedTitle
        
        # Look ahead for the same normalized title
        for ($j = $i + 1; $j -lt [Math]::Min($i + $MinDistance, $ShuffledTracks.Count); $j++) {
            $compareTrack = $ShuffledTracks[$j]
            if ($compareTrack.NormalizedTitle -eq $currentTitle) {
                $distance = $j - $i
                $violations++
                if ($distance -lt $minDistanceFound) {
                    $minDistanceFound = $distance
                }
                Write-Warning "Similar songs at positions $($i+1) and $($j+1) (distance: $distance) - '$currentTitle'"
            }
        }
    }
    
    if ($violations -eq 0) {
        Write-Host "Perfect shuffle! No similar songs within $MinDistance tracks of each other." -ForegroundColor Green
    } else {
        Write-Host "Found $violations constraint violation(s). Minimum distance found: $minDistanceFound" -ForegroundColor Yellow
        Write-Host "Consider reducing -MinimumDistance or running again for a different random shuffle." -ForegroundColor Yellow
    }
}

function Write-M3UFile {
    param(
        [string] $Path,
        [array] $ShuffledTracks
    )
    
    $lines = @()
    $lines += "#EXTM3U"
    $lines += "# Shuffled playlist generated by Shuffle-M3UPlaylist.ps1"
    $lines += "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += ""
    
    foreach ($track in $ShuffledTracks) {
        $lines += $track.Path
    }
    
    # Ensure directory exists
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    if ($dir -and !(Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    
    $lines | Out-File -FilePath $Path -Encoding UTF8
    Write-Host "`nSaved shuffled playlist to: $Path" -ForegroundColor Green
}

# ----- Main Script -----

Write-Host "=== M3U Playlist Smart Shuffler ===" -ForegroundColor Cyan
Write-Host ""

# Determine output path
if ([string]::IsNullOrWhiteSpace($OutputM3U)) {
    $dir = [System.IO.Path]::GetDirectoryName($InputM3U)
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($InputM3U)
    $ext = [System.IO.Path]::GetExtension($InputM3U)
    if ([string]::IsNullOrWhiteSpace($dir)) {
        $OutputM3U = "${nameWithoutExt}_shuffled${ext}"
    } else {
        $OutputM3U = Join-Path $dir "${nameWithoutExt}_shuffled${ext}"
    }
}

# Read input file
$tracks = Read-M3UFile -Path $InputM3U

# Group tracks by similarity
$trackInfo = Group-TracksBySimilarity -Tracks $tracks

# Perform smart shuffle
$shuffledTracks = Invoke-SmartShuffle -TrackInfo $trackInfo -MinDistance $MinimumDistance

# Test shuffle quality
Test-ShuffleQuality -ShuffledTracks $shuffledTracks -MinDistance $MinimumDistance

# Write output file
Write-M3UFile -Path $OutputM3U -ShuffledTracks $shuffledTracks

Write-Host "`nDone!" -ForegroundColor Green

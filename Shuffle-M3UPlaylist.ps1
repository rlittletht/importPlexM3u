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

function Get-NormalizedSongTitle {
    param([string] $FilePath)
    
    # Extract filename without extension
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    
    # Remove common patterns:
    # - Track numbers (e.g., "01 - ", "01. ", "01-", "1. ")
    $normalized = $filename -replace '^\d+[\s\.\-]+', ''
    
    # - Artist delimiter patterns (e.g., " - ", " by ", " ft. ", " feat. ")
    # Keep only the song title (typically before the first delimiter)
    $normalized = $normalized -replace '\s+-\s+.*$', ''
    $normalized = $normalized -replace '\s+by\s+.*$', ''
    $normalized = $normalized -replace '\s+ft[\.\s]+.*$', ''
    $normalized = $normalized -replace '\s+feat[\.\s]+.*$', ''
    
    # - Parenthetical info (e.g., "(Live)", "(Remix)", "(Remastered)")
    $normalized = $normalized -replace '\s*\([^\)]*\)\s*', ''
    $normalized = $normalized -replace '\s*\[[^\]]*\]\s*', ''
    
    # - Special characters and extra whitespace
    $normalized = $normalized -replace '[_\-]+', ' '
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

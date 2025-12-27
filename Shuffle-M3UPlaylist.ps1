<#
.SYNOPSIS
Shuffles a CSV playlist while keeping different versions of the same song spread apart.

.DESCRIPTION
This script reads a CSV playlist file and creates a shuffled version that attempts to
avoid playing different versions of the same song close together. For example, if the
playlist has "Jingle Bells" by 5 different artists, this script ensures they won't
play consecutively or too close to each other.

The CSV file should have columns: Filename, [Weight]

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

.PARAMETER TargetSongCount
Target number of songs in the output playlist. If specified and less than the total tracks,
only this many tracks will be randomly selected before shuffling. If not specified, all tracks
from the input CSV will be included.

.PARAMETER MinimumDistance
Minimum number of songs that should appear between different versions of the same song.
Default is 5. Higher values spread similar songs further apart.

.PARAMETER Seed
Random seed for reproducible shuffles. If not specified, uses a random seed.

.PARAMETER DistributionFile
Path to save a CSV file containing distribution statistics about the shuffle.

.PARAMETER DontWriteM3u
Skip writing the M3U output file. Only valid when -DistributionFile is specified.
Useful for testing distribution without generating the actual playlist file.

.PARAMETER PassThrough
Skip shuffling and write tracks in their original order from the CSV file.
When combined with -TargetSongCount, outputs the first N tracks without shuffling.

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "X:\Holiday\Playlists\Christmas.csv"

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "X:\Holiday\Playlists\Christmas.csv" -OutputM3U "X:\Holiday\Playlists\Christmas_Shuffled.m3u" -MinimumDistance 8

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "Christmas.csv" -TargetSongCount 100 -Seed 42

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "Christmas.csv" -TargetSongCount 100 -DistributionFile "dist.csv" -DontWriteM3u

.EXAMPLE
.\Shuffle-M3UPlaylist.ps1 -InputCSV "Christmas.csv" -TargetSongCount 100 -PassThrough
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $InputCSV,

    [Parameter(Mandatory = $false)]
    [string] $OutputM3U,

    [Parameter(Mandatory = $false)]
    [int] $TargetSongCount,

    [Parameter(Mandatory = $false)]
    [string[]] $GrepGroup,

    [Parameter(Mandatory = $false)]
    [string[]] $DebugParts,

    [Parameter(Mandatory = $false)]
    [int] $MinimumDistance = 5,

    [Parameter(Mandatory = $false)]
    [int] $ShowGroupCount = 20,

    [Parameter(Mandatory = $false)]
    [int] $Seed,

    [Parameter(Mandatory = $false)]
    [string] $DistributionFile,

    [Parameter(Mandatory = $false)]
    [switch] $DontWriteM3u,

    [Parameter(Mandatory = $false)]
    [switch] $PassThrough
)
    
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Validate that PassThrough is not used with shuffle-related parameters
if ($PassThrough)
{
    $invalidParams = @()
    if ($PSBoundParameters.ContainsKey('TargetSongCount')) { $invalidParams += 'TargetSongCount' }
    if ($PSBoundParameters.ContainsKey('MinimumDistance')) { $invalidParams += 'MinimumDistance' }
    if ($PSBoundParameters.ContainsKey('Seed')) { $invalidParams += 'Seed' }
    if ($PSBoundParameters.ContainsKey('DistributionFile')) { $invalidParams += 'DistributionFile' }
    if ($DontWriteM3u) { $invalidParams += 'DontWriteM3u' }
    
    if ($invalidParams.Count -gt 0)
    {
        throw "The -PassThrough parameter cannot be used with: $($invalidParams -join ', '). These parameters are only valid for shuffled output."
    }
}

# Validate that DontWriteM3u is only used with DistributionFile
if ($DontWriteM3u -and [string]::IsNullOrWhiteSpace($DistributionFile))
{
    throw "The -DontWriteM3u parameter can only be used when -DistributionFile is specified."
}

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
            # Default weight to 1 if not present or empty
            $weight = 1
            if ($row.PSObject.Properties.Name -contains 'Weight' -and ![string]::IsNullOrWhiteSpace($row.Weight))
            {
                $weight = $row.Weight
            }
            
            $tracks += [PSCustomObject]@{
                Filename = $row.Filename
                Weight   = $weight
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
    
    # remove leading "."
    $Part = $Part -replace '^ *\. *', ''

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

function partMatchesNotSongPattern
{
    param([string] $part)

    $matchValuesNotSong = '^(stevie|jose)\s+'
    if ($part -match $matchValuesNotSong)
    {
        return $true
    }

    return $false
}

function partMatchesSongPattern
{
    param([string] $part)

    if (partMatchesNotSongPattern $part)
    {
        return $false
    }

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
    if ($part.Length -gt $partOther.Length -and -not (partMatchesSongPattern $partOther) -and -not (partMatchesNotSongPattern $part))
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
    elseif ((partMatchesNotSongPattern $firstPart) -and -not (partMatchesNotSongPattern $lastPart))
    {
        if ($debugThis)
        {
            Write-Host "      Choosing not not song from parts $($firstPart)/$($lastPart): chose $($lastPart)"
        }
        return $lastPart

    }
    elseif ((partMatchesNotSongPattern $lastPart) -and -not (partMatchesNotSongPattern $firstPart))
    {
        if ($debugThis)
        {
            Write-Host "      Choosing not not song from parts $($firstPart)/$($lastPart): chose $($firstPart)"
        }
        return $firstPart

    }
    else
    {
        if ($debugThis)
        {
            Write-Host "      Choosing between parts $($firstPart)/$($lastPart): chose by default $($firstPart)"
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
            $part = $part -replace '^ *\. *', ''

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

    $normalizedPartReturn = NormalizePart($normalized)
    if ($debugThis)
    {
        Write-Host "       Final result!  CandidateParts: '$($candidateParts -join ', ')', chose: $($normalized) normalized to $($normalizedPartReturn)"
    }
    return NormalizePart($normalizedPartReturn)
}

function Group-TracksBySimilarity
{
    param([PSCustomObject[]] $TrackObjects, [string[]] $GrepGroup, [string[]] $DebugParts, [int] $ShowGroupCount)
    
    # first pass just collects groups for the next pass
    $groupsPass1 = @{}
    $groups = @{}
    $trackInfo = @()
    
    # first pass just collects groups for the next pass
    
    Write-Host "`nBuilding similarity groups for $($TrackObjects.Count) tracks..." -ForegroundColor Cyan

    foreach ($trackObject in $TrackObjects)
    {
        $track = $trackObject.Filename

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

    foreach ($trackObject in $TrackObjects)
    {
        $track = $trackObject.Filename

        $normalized = Get-NormalizedSongTitle $track $groupsPass1 $DebugParts
        
        Write-Verbose "Normalized: $track to $normalized"

        if (!$groups.ContainsKey($normalized))
        {
            $groups[$normalized] = @()
        }
        $groups[$normalized] += $track
        
        $trackInfo += [PSCustomObject]@{
            Path            = $track
            NormalizedTitle = $normalized
            OriginalObject  = $trackObject
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
    
    Write-Verbose "`nFinal track info: $($trackInfo | Format-Table | Out-String)"
    return , $trackInfo
}

function CanPlaceTrack
{
    param([PSCustomObject] $Track, [hashtable] $LastUsed, [int]$CurrentTargetIndex, [int] $MinDistance)
    # Check if this normalized title was used recently
    if ($LastUsed.ContainsKey($Track.NormalizedTitle))
    {
        $distance = $CurrentTargetIndex - $LastUsed[$Track.NormalizedTitle]
        if ($distance -lt $MinDistance)
        {
            return $false
        }
    }

    return $true
}

function BuildVirtualIndexMap
{
    param(
        [PSCustomObject[]] $TrackInfo
    )

    Write-Verbose "Building virtual index map for weighted random selection from $($TrackInfo.Count) tracks..."
    $virtualIndexMap = @()

    $virtualIndexMap += 0
    $virtualIndexNext = 0

    # Build a virtual index map to allow random weighted selection
    for ($i = 0; $i -lt $TrackInfo.Count; $i++)
    {
        $weight = [Math]::Max([int]$TrackInfo[$i].OriginalObject.Weight, 1)
        $virtualIndexNext += $weight
        $virtualIndexMap += $virtualIndexNext
    }

    return , $virtualIndexMap
}

function GetIndexFromVirtualIndex
{
    param(
        [int[]] $VirtualIndexMap,
        [int] $VirtualIndex
    )

    # Binary search to find the real index corresponding to the virtual index
    $low = 0
    $high = $VirtualIndexMap.Count - 1

    while ($low -lt $high)
    {
        $mid = [Math]::Floor(($low + $high) / 2)
        if ($VirtualIndexMap[$mid] -le $VirtualIndex)
        {
            # this might be our match if the next item is greater than VirtualIndex
            $nextIndex = $mid + 1

            if (($nextIndex -ge $high) -or ($VirtualIndexMap[$nextIndex] -gt $VirtualIndex))
            {
                return $mid
            }

            $low = $nextIndex
        }
        else
        {
            $high = $mid
        }
    }

    return $low - 1
}

function TestIndexFromVirtualIndex
{
    param(
        [int[]] $VirtualIndexMap,
        [int] $VirtualIndex,
        [int] $ExpectedRealIndex
    )
    $realIndex = GetIndexFromVirtualIndex -VirtualIndexMap $VirtualIndexMap -VirtualIndex $VirtualIndex
    if ($realIndex -ne $ExpectedRealIndex)
    {
        throw "Test failed for VirtualIndex $(VirtualIndex): expected $ExpectedRealIndex, got $realIndex"
    }
}

function DoTestsForVirtualIndexMapping
{
    # Example virtual index map for testing
    $virtualIndexMap = @(0, 3, 7, 10)  # 3 items with weights 3, 4, 3

    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 0 -ExpectedRealIndex 0
    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 2 -ExpectedRealIndex 0
    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 3 -ExpectedRealIndex 1
    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 6 -ExpectedRealIndex 1
    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 7 -ExpectedRealIndex 2
    TestIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex 9 -ExpectedRealIndex 2
}

function Invoke-SmartShuffle
{
    param(
        [PSCustomObject[]] $TrackInfo,
        [int] $MinDistance,
        [int] $TargetCount,
        [string] $DistributionFile
    )
    Write-Host "`nShuffling $TargetCount songs with minimum distance of $MinDistance between similar songs for $($TrackInfo.Count) tracks..." -ForegroundColor Cyan
    
    # Result list
    $shuffled = New-Object System.Collections.Generic.List[PSCustomObject]
    
    # Track when each normalized title was last used (index in shuffled list)
    $lastUsed = @{}
    $frequencies = @{}

    $virtualIndexMap = BuildVirtualIndexMap -TrackInfo $TrackInfo
    Write-Verbose "Virtual index map: $($virtualIndexMap -join ', ')"

    $virtualTrackCount = $virtualIndexMap[$virtualIndexMap.Count - 1]

    $attempts = 0
    $maxAttemptsPerTrack = 100

    while ($TargetCount -gt 0)
    {
        $attempts++
        
        # Shuffle the available list to randomize selection
        $nextVirtualIndex = Get-Random -Maximum $virtualTrackCount

        # get the real index from that index
        $nextIndex = GetIndexFromVirtualIndex -VirtualIndexMap $virtualIndexMap -VirtualIndex $nextVirtualIndex

        Write-Verbose "Selected virtual index: $nextVirtualIndex, real index: $nextIndex"

        # Try to find a track that respects the minimum distance
        $track = $TrackInfo[$nextIndex]
        #        Write-Verbose "Considering track: $($track | Format-Table | Out-String)"
        # see of we can place that track. if we can't just skip
        if ((CanPlaceTrack -Track $track -LastUsed $lastUsed -CurrentTargetIndex $shuffled.Count -MinDistance $MinDistance) -or $attempts -gt $maxAttemptsPerTrack)
        {
            if ($attempts -gt $maxAttemptsPerTrack)
            {
                Write-Verbose "Relaxing constraints to place track: $($track.Path)"
            }
            #            Write-Verbose "Placing track: $($track.Path)"
            # Place this track
            [void]$shuffled.Add($track)
            if ($frequencies.ContainsKey($track.Path))
            {
                $frequencies[$track.Path].Count += 1
                $frequencies[$track.Path].Indexes += $shuffled.Count - 1
            }
            else
            {
                $frequencies[$track.Path] = [PSCustomObject]@{ 
                    Count   = 1; 
                    Indexes = @($shuffled.Count - 1) 
                }
            }
            $lastUsed[$track.NormalizedTitle] = $shuffled.Count - 1
            # tabulate by path since path is what is weighted
            $attempts = 0
            --$TargetCount
        }
    }

    Write-Host "Shuffle complete. Generated playlist with $($shuffled.Count) tracks." -ForegroundColor Green

    # Output distribution to file if specified
    if ($DistributionFile)
    {
        $lines = @()
        $lines += "NormalizedTitle,Count,Expected,Distribution,FirstDelta,LastDelta,AverageDelta,Indexes"

        foreach ($track in $TrackInfo)
        {
            if ($frequencies.ContainsKey($track.Path))
            {
                $frequency = $frequencies[$track.Path]
                $count = $frequency.Count
                $firstPlayDelta = if ($frequency.Indexes.Count -gt 0) { $frequency.Indexes[0] } else { -1 }
                $lastPlayDelta = if ($frequency.Indexes.Count -gt 0) { $shuffled.Count - $frequency.Indexes[-1] } else { -1 }

                $sumDeltas = 0
                $lastIndex = 0
                for ($i = 1; $i -lt $frequency.Indexes.Count; $i++)
                {
                    $index = $frequency.Indexes[$i]
                    $sumDeltas += $frequency.Indexes[$i] - $frequency.Indexes[$i - 1]
                }
                $averageDelta = if ($frequency.Indexes.Count -gt 1) { $sumDeltas / ($frequency.Indexes.Count - 1) } else { -1 }
            }
            else
            {
                $count = 0
                $firstPlayDelta = -1
                $lastPlayDelta = -1
                $averageDelta = -1
            }
            
            $expectedDistribution = [Math]::Round((([int]$track.OriginalObject.Weight) / $virtualTrackCount) * 1000, 2) / 1000;
            $distribution = [Math]::Round(($count / $shuffled.Count) * 1000, 2) / 1000
            $indexes = if ($frequencies.ContainsKey($track.Path)) { $frequencies[$track.Path].Indexes -join ';' } else { '' }
            $line = """$($track.Path)"", $count, $expectedDistribution, $distribution, $firstPlayDelta, $lastPlayDelta, $averageDelta, $indexes"
            $lines += $line
        }
        $lines | Out-File -FilePath $DistributionFile -Encoding UTF8
        Write-Host "Saved track count distribution to: $DistributionFile" -ForegroundColor Green
    }
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

DoTestsForVirtualIndexMapping

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

$targetCount = $trackObjects.Count

# Apply TargetSongCount if specified
if ($TargetSongCount -and $TargetSongCount -gt 0)
{
    Write-Verbose "Applying TargetSongCount: $TargetSongCount"
    $targetCount = $TargetSongCount
}

Write-Verbose "Using target count: $targetCount with total tracks: $($trackObjects.Count)"

if ($PassThrough)
{
    Write-Host "`nPass-through mode: Writing tracks in original order (no shuffling)..." -ForegroundColor Cyan
    
    # Create trackInfo objects without grouping
    $trackInfo = @()
    foreach ($trackObj in $trackObjects)
    {
        $trackInfo += [PSCustomObject]@{
            Path            = $trackObj.Filename
            NormalizedTitle = ""
            Weight          = 0
        }
    }
    
    $shuffledTracks = $trackInfo
    
    Write-Host "Selected $($shuffledTracks.Count) tracks for output." -ForegroundColor Cyan
}
else
{
    # Group tracks by similarity
    $trackInfo = Group-TracksBySimilarity -TrackObjects $trackObjects -GrepGroup $GrepGroup -DebugParts $DebugParts -ShowGroupCount $ShowGroupCount

    Write-Verbose "Total grouped tracks: $($trackInfo.Count)"

    # Perform smart shuffle
    $shuffledTracks = Invoke-SmartShuffle -TrackInfo $trackInfo -MinDistance $MinimumDistance -TargetCount $targetCount -DistributionFile $DistributionFile

    # Test shuffle quality
    Test-ShuffleQuality -ShuffledTracks $shuffledTracks -MinDistance $MinimumDistance
}

# Write output file
if (-not $DontWriteM3u)
{
    Write-M3UFile -Path $OutputM3U -ShuffledTracks $shuffledTracks
}
else
{
    Write-Host "`nSkipping M3U file creation (DontWriteM3u specified)." -ForegroundColor Yellow
}

Write-Host "`nDone!" -ForegroundColor Green

<#
Imports one or more .m3u playlists into Plex as "dumb" audio playlists,
preserving the order of tracks in each .m3u file.

Requires:
- Plex base URL (e.g. https://<plex>.plex.direct:32400)
- X-Plex-Token
- machineIdentifier
- music library section id (your "key" = 6)

To get the token, using plex web, get info for an item then click on the XML link to open the XML info in a new tab. Copy the token from the end of the URL
To get the plex base URL, do that same as above, but get to root server address after https
To get the machine identifier, run curl https://<base-url>/identity?X-Plex-Token=<token>
To get the section id, run https://<base-url>/library/sections?X-Plex-Token=<token>
(note that the key will come from the "key" attribute on the <Directory> element for your music library)

Sample:
.\Import-PlexM3UPlaylists.ps1 -PlexBaseUrl https://192-168-1-45.<idvalue>.plex.direct:32400 -PlexToken <token> -MachineIdentifier <machineid> -SectionId <key> -M3UPaths @("X:\Holiday\Playlists\ChristmasCurated_Short2.m3u", "X:\Holiday\Playlists\ChristmasCurated_Short3.m3u")
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $PlexBaseUrl,            # e.g. https://<plex>.plex.direct:32400

    [Parameter(Mandatory = $true)]
    [string] $PlexToken,

    [Parameter(Mandatory = $true)]
    [string] $MachineIdentifier,      # e.g. 87e186a3...

    [Parameter(Mandatory = $true)]
    [int] $SectionId,                 # e.g. 6

    [Parameter(Mandatory = $true)]
    [string[]] $M3UPaths,             # one or more .m3u files

    [Parameter(Mandatory = $false)]
    [string] $MediaContainerFile,     # optional path to load cached Plex MediaContainer XML responses

    [int] $LibraryPageSize = 5000,    # paging size when enumerating tracks
    [int] $ChunkSize = 200,           # playlist append chunk size (tune as desired)
    [switch] $SkipCertificateCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function UrlEncode([string] $s) { [System.Net.WebUtility]::UrlEncode($s) }

function New-PlexHeaders
{
    @{
        'Accept'       = 'application/json'
        'X-Plex-Token' = $PlexToken
    }
}

function Invoke-Plex([ValidateSet('GET', 'POST', 'PUT')] [string] $Method, [string] $PathAndQuery)
{
    $uri = ($PlexBaseUrl.TrimEnd('/') + $PathAndQuery)
    $headers = New-PlexHeaders

    try
    {
        if ($SkipCertificateCheck)
        {
            return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -SkipCertificateCheck
        }
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
    }
    catch
    {
        # Give helpful diagnostics on HTTP failures
        $ex = $_.Exception
        if ($ex.Response -and $ex.Response.StatusCode)
        {
            $status = [int]$ex.Response.StatusCode
            $desc = $ex.Response.StatusDescription
            Write-Error "HTTP $status $desc calling: $Method $uri"
        }
        else
        {
            Write-Error "Error calling: $Method $uri :: $($_.Exception.Message)"
        }
        throw
    }
}

function Read-M3UTracks([string] $m3uPath)
{
    if (!(Test-Path -LiteralPath $m3uPath))
    {
        throw "M3U not found: $m3uPath"
    }

    # Most .m3u are UTF-8; if yours isn't, fix it upstream.
    $lines = Get-Content -LiteralPath $m3uPath -Encoding UTF8

    $tracks = @()
    foreach ($line in $lines)
    {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t.StartsWith('#')) { continue }
        $tracks += $t
    }

    if ($tracks.Count -eq 0)
    {
        throw "No track paths found in: $m3uPath"
    }

    return $tracks
}

function Build-FilePathToRatingKeyMap
{
    param(
        [string] $MediaContainerFile
    )

    Write-Host "Building Plex file->ratingKey map from section $SectionId ..." -ForegroundColor Cyan

    $map = @{}
    $start = 0
    $total = $null
    $allResponses = @()

    # Query Plex directly
    while ($true)
    {
        # type=10 => Track
        
        # Check if we should load from file instead of querying Plex
        if (![string]::IsNullOrWhiteSpace($MediaContainerFile) -and (Test-Path -LiteralPath $MediaContainerFile))
        {
            Write-Host "Loading MediaContainer from file: $MediaContainerFile" -ForegroundColor Yellow
            $jsonContent = Get-Content -LiteralPath $MediaContainerFile -Raw
            # Use Invoke-RestMethod's XML parsing to convert to PSCustomObject like the live API does
            $resp = ConvertFrom-Json $jsonContent
            $mc = $resp.MediaContainer
            Write-Host "Loaded MediaContainer from File: $mc"
        }
        else
        {
            $q = "/library/sections/$SectionId/all?type=10&X-Plex-Container-Start=$start&X-Plex-Container-Size=$LibraryPageSize"
            $resp = Invoke-Plex GET $q
            $mc = $resp.MediaContainer
            Write-Host "Loaded MediaContainer from Plex: $mc"
        }
        if ($null -eq $total)
        {
            $total = [int]$mc.totalSize
            Write-Host "Plex reports $total tracks in section." -ForegroundColor DarkCyan
        }

        $items = @()
        if ($mc.Metadata) { $items = $mc.Metadata }

        foreach ($it in $items)
        {
            $rk = $it.ratingKey
            if ($null -eq $rk) { continue }

            if ($it.Media)
            {
                foreach ($m in $it.Media)
                {
                    if ($m.Part)
                    {
                        foreach ($p in $m.Part)
                        {
                            $file = $p.file
                            if (![string]::IsNullOrWhiteSpace($file) -and -not $map.ContainsKey($file))
                            {
                                $map[$file] = [string]$rk
                            }
                        }
                    }
                }
            }
        }

        $start += $LibraryPageSize
        Write-Host ("Mapped {0} files so far..." -f $map.Count) -ForegroundColor DarkGray

        if ($start -ge $total) { break }
        if ($items.Count -eq 0) { break }

        if (![string]::IsNullOrWhiteSpace($MediaContainerFile))
        {
            # If loading from file, only do one pass. we should have broken above
            Write-Host "Loaded from file, but not all records were in the file. Total was $($total), but file only had $($items.Count)." -ForegroundColor Yellow
            break
        }
    }

    if ($map.Count -eq 0)
    {
        throw "No mappings created. Are there tracks in section $SectionId?"
    }

    return $map
}

function Get-UriForRatingKeys([string[]] $ratingKeys)
{
    $rkList = ($ratingKeys -join ",")
    "server://$MachineIdentifier/com.plexapp.plugins.library/library/metadata/$rkList"
}

function New-PlexPlaylistWithFirstChunk([string] $title, [string[]] $firstChunkRatingKeys)
{
    if ($firstChunkRatingKeys.Count -lt 1)
    {
        throw "Cannot create playlist '$title' with zero items."
    }

    $titleEnc = UrlEncode $title
    $uriEnc = UrlEncode (Get-UriForRatingKeys $firstChunkRatingKeys)

    $createPath = "/playlists?type=audio&smart=0&title=$titleEnc&uri=$uriEnc"
    $createResp = Invoke-Plex POST $createPath

    $plist = $createResp.MediaContainer.Metadata
    if ($null -eq $plist -or $plist.Count -eq 0)
    {
        throw "Playlist create did not return playlist metadata for '$title'."
    }

    $playlistId = $plist[0].ratingKey
    if ([string]::IsNullOrWhiteSpace($playlistId))
    {
        throw "Could not read playlistId after creating '$title'."
    }

    return $playlistId
}

function Add-PlaylistItemsChunk([string] $playlistId, [string[]] $chunkRatingKeys)
{
    if ($chunkRatingKeys.Count -lt 1) { return }

    $uriEnc = UrlEncode (Get-UriForRatingKeys $chunkRatingKeys)

    # Plex servers differ; try the most common, then fallback.
    $pathsToTry = @(
        "/playlists/$playlistId?uri=$uriEnc",
        "/playlists/$playlistId/items?uri=$uriEnc"
    )

    $lastError = $null
    foreach ($p in $pathsToTry)
    {
        try
        {
            [void](Invoke-Plex PUT $p)
            return
        }
        catch
        {
            $lastError = $_
        }
    }

    throw $lastError
}

function Split-IntoChunks([string[]] $arr, [int] $size)
{
    if ($size -lt 1) { throw "ChunkSize must be >= 1." }
    for ($i = 0; $i -lt $arr.Count; $i += $size)
    {
        $take = [Math]::Min($size, $arr.Count - $i)
        , ($arr[$i..($i + $take - 1)])
    }
}

# ----- Main -----
$PlexBaseUrl = $PlexBaseUrl.TrimEnd('/')

$fileToRk = Build-FilePathToRatingKeyMap -MediaContainerFile $MediaContainerFile

foreach ($m3u in $M3UPaths)
{
    $title = [System.IO.Path]::GetFileNameWithoutExtension($m3u)
    Write-Host "Processing $m3u -> Plex playlist '$title'" -ForegroundColor Green

    $trackFiles = Read-M3UTracks $m3u

    $missing = New-Object System.Collections.Generic.List[string]
    $ratingKeysOrdered = New-Object System.Collections.Generic.List[string]

    foreach ($f in $trackFiles)
    {
        if ($fileToRk.ContainsKey($f))
        {
            $ratingKeysOrdered.Add($fileToRk[$f])
        }
        else
        {
            $missing.Add($f)
        }
    }

    if ($missing.Count -gt 0)
    {
        Write-Warning ("{0} tracks were not found in Plex (showing first 20):" -f $missing.Count)
        $missing | Select-Object -First 20 | ForEach-Object { Write-Warning "  MISSING: $_" }
        Write-Warning "Skipping '$title'. Fix missing paths (must match Plex's Part.file exactly) and re-run."
        continue
    }

    $all = $ratingKeysOrdered.ToArray()
    $chunks = @(Split-IntoChunks $all $ChunkSize)

    Write-Host ("Creating playlist with first chunk ({0} tracks)..." -f $chunks[0].Count) -ForegroundColor Cyan
    $playlistId = New-PlexPlaylistWithFirstChunk -title $title -firstChunkRatingKeys $chunks[0]

    for ($c = 1; $c -lt $chunks.Count; $c++)
    {
        Write-Host ("Appending chunk {0}/{1} ({2} tracks)..." -f ($c + 1), $chunks.Count, $chunks[$c].Count) -ForegroundColor DarkCyan
        Add-PlaylistItemsChunk -playlistId $playlistId -chunkRatingKeys $chunks[$c]
    }

    Write-Host ("Created '$title' (id=$playlistId) with {0} tracks." -f $all.Length) -ForegroundColor Cyan
}

Write-Host "Done." -ForegroundColor Green

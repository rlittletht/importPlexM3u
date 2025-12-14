<#
Imports one or more .m3u playlists into Plex as "dumb" audio playlists,
preserving the order of tracks in each .m3u file.

Requires:
- Plex base URL (e.g. https://<plex>.plex.direct:32400)
- X-Plex-Token
- machineIdentifier
- music library section id (your "key" = 6)
#>

param(
    [Parameter(Mandatory=$true)]
    [string] $PlexBaseUrl,            # e.g. https://192-168-1-45....plex.direct:32400

    [Parameter(Mandatory=$true)]
    [string] $PlexToken,

    [Parameter(Mandatory=$true)]
    [string] $MachineIdentifier,      # e.g. 87e186a3...

    [Parameter(Mandatory=$true)]
    [int] $SectionId,                 # e.g. 6

    [Parameter(Mandatory=$true)]
    [string[]] $M3UPaths,             # one or more .m3u files

    [int] $PageSize = 5000,           # bump if you have huge libraries
    [switch] $SkipCertificateCheck    # only if TLS validation fails for you
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function UrlEncode([string] $s) {
    return [System.Net.WebUtility]::UrlEncode($s)
}

function New-PlexHeaders {
    $h = @{
        'Accept' = 'application/json'
        'X-Plex-Token' = $PlexToken
    }
    return $h
}

function Invoke-PlexJsonGet([string] $pathAndQuery) {
    $uri = "$PlexBaseUrl$pathAndQuery"
    $headers = New-PlexHeaders
    Write-Host "GET $uri"

    if ($SkipCertificateCheck) {
        return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -SkipCertificateCheck
    }
    return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
}

function Invoke-PlexJsonPost([string] $pathAndQuery) {
    $uri = "$PlexBaseUrl$pathAndQuery"
    $headers = New-PlexHeaders
    Write-Host "POST $uri"

    if ($SkipCertificateCheck) {
        return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -SkipCertificateCheck
    }
    return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers
}

function Invoke-PlexJsonPut([string] $pathAndQuery) {
    $uri = "$PlexBaseUrl$pathAndQuery"
    $headers = New-PlexHeaders
    if ($SkipCertificateCheck) {
        return Invoke-RestMethod -Method Put -Uri $uri -Headers $headers -SkipCertificateCheck
    }
    return Invoke-RestMethod -Method Put -Uri $uri -Headers $headers
}

function Read-M3UTracks([string] $m3uPath) {
    if (!(Test-Path -LiteralPath $m3uPath)) {
        throw "M3U not found: $m3uPath"
    }

    $lines = Get-Content -LiteralPath $m3uPath -Encoding UTF8

    $tracks = @()
    foreach ($line in $lines) {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t.StartsWith('#')) { continue }  # comments / #EXTM3U / #EXTINF
        $tracks += $t
    }

    if ($tracks.Count -eq 0) {
        throw "No track paths found in: $m3uPath"
    }

    return $tracks
}

function Build-FilePathToRatingKeyMap {
    Write-Host "Building Plex track map from section $SectionId ..." -ForegroundColor Cyan

    # /library/sections/{id}/all?type=10 lists tracks in a music library section (type=10 = Track).
    # We'll page using X-Plex-Container-Start / X-Plex-Container-Size.
    $map = @{}
    $start = 0
    $total = $null

    while ($true) {
        $q = "/library/sections/$SectionId/all?type=10&X-Plex-Container-Start=$start&X-Plex-Container-Size=$PageSize"
        $resp = Invoke-PlexJsonGet $q

        $mc = $resp.MediaContainer
        if ($null -eq $total) {
            $total = [int]$mc.totalSize
            Write-Host "Total tracks reported by Plex: $total" -ForegroundColor DarkCyan
        }

        $items = @()
        if ($mc.Metadata) { $items = $mc.Metadata }

        foreach ($it in $items) {
            # ratingKey identifies the track in Plex
            $rk = $it.ratingKey
            if ($null -eq $rk) { continue }

            # Track has Media[] -> Part[] -> file
            if ($it.Media) {
                foreach ($m in $it.Media) {
                    if ($m.Part) {
                        foreach ($p in $m.Part) {
                            $file = $p.file
                            if (![string]::IsNullOrWhiteSpace($file)) {
                                # Keep first seen mapping; duplicates are rare but can happen with multi-part items.
                                if (-not $map.ContainsKey($file)) {
                                    $map[$file] = [string]$rk
                                }
                            }
                        }
                    }
                }
            }
        }

        $start += $PageSize
        Write-Host ("Mapped {0} files so far..." -f $map.Count) -ForegroundColor DarkGray

        if ($start -ge $total) { break }
        if ($items.Count -eq 0) { break } # safety
    }

    if ($map.Count -eq 0) {
        throw "Failed to build any file->ratingKey mappings. Are there tracks in section $SectionId?"
    }

    return $map
}

function Ensure-PlaylistFromRatingKeys([string] $title, [string[]] $ratingKeys) {
    if ($ratingKeys.Count -lt 1) {
        throw "Cannot create playlist '$title' with zero items (Plex returns 400)."
    }

    # Create playlist WITH ITEMS (uri is required on many PMS versions)
    $tEnc = UrlEncode $title
    $rkList = ($ratingKeys -join ",")

    # NOTE: uri format must be:
    # server://{machineId}/com.plexapp.plugins.library/library/metadata/{rk1,rk2,...}
    $uri = "server://$MachineIdentifier/com.plexapp.plugins.library/library/metadata/$rkList"
    $uriEnc = UrlEncode $uri

    $createPath = "/playlists?type=audio&smart=0&title=$tEnc&uri=$uriEnc"
    $createResp = Invoke-PlexJsonPost $createPath

    $plist = $createResp.MediaContainer.Metadata
    if ($null -eq $plist -or $plist.Count -eq 0) {
        throw "Playlist create did not return playlist metadata for '$title'."
    }

    $playlistId = $plist[0].ratingKey
    if ([string]::IsNullOrWhiteSpace($playlistId)) {
        throw "Could not read playlistId/ratingKey after creating '$title'."
    }

    return $playlistId
}

# --- Main ---
# Normalize base URL (no trailing slash)
$PlexBaseUrl = $PlexBaseUrl.TrimEnd('/')

# Build mapping once
$fileToRk = Build-FilePathToRatingKeyMap

foreach ($m3u in $M3UPaths) {
    $title = [System.IO.Path]::GetFileNameWithoutExtension($m3u)
    Write-Host "Processing: $m3u  => Playlist '$title'" -ForegroundColor Green

    $trackFiles = Read-M3UTracks $m3u

    $missing = New-Object System.Collections.Generic.List[string]
    $ratingKeysOrdered = New-Object System.Collections.Generic.List[string]

    foreach ($f in $trackFiles) {
        if ($fileToRk.ContainsKey($f)) {
            $ratingKeysOrdered.Add($fileToRk[$f])
        } else {
            $missing.Add($f)
        }
    }

    if ($missing.Count -gt 0) {
        Write-Warning ("{0} tracks from '{1}' were not found in Plex section {2}. Showing first 20:" -f $missing.Count, $m3u, $SectionId)
        $missing | Select-Object -First 20 | ForEach-Object { Write-Warning "  MISSING: $_" }
        Write-Warning "Fix the paths (must match Plex's Part.file exactly) and re-run."
        continue
    }

    $playlistId = Ensure-PlaylistFromRatingKeys -title $title -ratingKeys $ratingKeysOrdered.ToArray()
    Write-Host "Created playlist '$title' (id=$playlistId) with $($ratingKeysOrdered.Count) tracks." -ForegroundColor Cyan
}

Write-Host "Done." -ForegroundColor Green

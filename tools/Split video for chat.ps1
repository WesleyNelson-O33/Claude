<#
    Split video for chat
    --------------------
    Cuts a long recording into pieces small enough to upload in a chat,
    instead of putting it through GitHub.

    It always makes AUDIO pieces. Audio is tiny, and audio is all that is
    needed for a transcript. A ninety minute recording comes to about twenty
    megabytes in total, in four or five files.

    Add -AlsoVideo if the screen has to be seen as well. Video pieces are much
    bigger, so there will be a lot more of them.

    How to run it
      1. Put this file in the same folder as the recording.
      2. Double click the .bat file of the same name.
         Or right click this file and choose Run with PowerShell.

    Nothing is installed. Nothing on the computer is changed. The original
    recording is not touched. Everything new goes into a folder called for-chat.
#>

param(
    # The recording. Leave it out and it takes the newest video in this folder.
    [string] $Video,

    # How many minutes in each audio piece.
    [int] $AudioMinutes = 20,

    # Also cut the video up, not just the audio.
    [switch] $AlsoVideo,

    # How many minutes in each video piece.
    [int] $VideoMinutes = 5,

    # The most any one piece may be, in megabytes. Chat takes 30, so 25 is safe.
    [int] $MaxMB = 25
)

$ErrorActionPreference = 'Stop'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

function Say($t)  { Write-Host $t }
function Step($t) { Write-Host ''; Write-Host $t -ForegroundColor Cyan }
function Oops($t) { Write-Host ''; Write-Host $t -ForegroundColor Red }

try {
    # ------------------------------------------------------------ the recording
    Step 'Looking for the recording'
    if (-not $Video) {
        $kinds = '*.mp4', '*.mov', '*.mkv', '*.m4v', '*.avi', '*.wmv', '*.webm'
        $found = Get-ChildItem -Path (Join-Path $here '*') -Include $kinds -File `
                     -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if (-not $found) {
            throw "No video in this folder. Put this script next to the recording and run it again."
        }
        $Video = $found[0].FullName
    }
    if (-not (Test-Path -LiteralPath $Video)) { throw "Cannot find $Video" }
    $file = Get-Item -LiteralPath $Video
    Say ('  ' + $file.Name)
    Say ('  ' + [math]::Round($file.Length / 1MB, 1) + ' MB')

    # ------------------------------------------------------------------ ffmpeg
    Step 'Getting the tool that does the cutting'
    $toolDir = Join-Path $here 'ffmpeg-tool'
    $ffmpeg = $null

    $onPath = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($onPath) {
        $ffmpeg = $onPath.Source
        Say '  Already on this computer.'
    } else {
        $already = Get-ChildItem -Path $toolDir -Filter 'ffmpeg.exe' -Recurse `
                       -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($already) {
            $ffmpeg = $already.FullName
            Say '  Downloaded before, using that.'
        }
    }

    if (-not $ffmpeg) {
        Say '  Downloading it. About 40 MB, so give it a minute.'
        $zip = Join-Path $env:TEMP 'ffmpeg-for-chat.zip'
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
        $sources = @(
            'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip',
            'https://github.com/BtbN/FFmpeg-Builds/releases/latest/download/ffmpeg-master-latest-win64-gpl.zip'
        )
        $got = $false
        foreach ($url in $sources) {
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
                $got = $true
                break
            } catch {
                Say ('  That one did not work. Trying another.')
            }
        }
        if (-not $got) {
            throw ("The download was blocked, which is usually the work network. " +
                   "Ask IT for ffmpeg, or put the recording on GitHub the way we did last time.")
        }
        New-Item -ItemType Directory -Force -Path $toolDir | Out-Null
        Expand-Archive -LiteralPath $zip -DestinationPath $toolDir -Force
        Remove-Item -LiteralPath $zip -Force
        $exe = Get-ChildItem -Path $toolDir -Filter 'ffmpeg.exe' -Recurse | Select-Object -First 1
        if (-not $exe) { throw 'Could not find ffmpeg.exe after unzipping.' }
        $ffmpeg = $exe.FullName
    }
    Say ('  ' + $ffmpeg)

    # ---------------------------------------------------------------- how long
    Step 'Reading how long the recording is'
    $seconds = 0
    $probe = Join-Path (Split-Path $ffmpeg) 'ffprobe.exe'
    if (Test-Path -LiteralPath $probe) {
        $out = & $probe -v error -show_entries format=duration -of csv=p=0 $file.FullName 2>$null
        if ($out) { $seconds = [double](($out | Select-Object -First 1) -replace '[^\d\.]', '') }
    }
    if ($seconds -le 0) {
        $text = (& $ffmpeg -hide_banner -i $file.FullName 2>&1) | Out-String
        if ($text -match 'Duration:\s*(\d+):(\d+):(\d+)') {
            $seconds = [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]
        }
    }
    if ($seconds -gt 0) {
        Say ('  ' + [math]::Floor($seconds / 60) + ' minutes ' +
             [math]::Round($seconds % 60) + ' seconds')
    } else {
        Say '  Could not read the length. Carrying on anyway.'
    }

    $outDir = Join-Path $here 'for-chat'
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    Get-ChildItem -Path $outDir -File -ErrorAction SilentlyContinue | Remove-Item -Force

    # ------------------------------------------------------------------- audio
    Step 'Making the audio pieces'
    Say ('  One piece every ' + $AudioMinutes + ' minutes.')
    & $ffmpeg -hide_banner -loglevel error -y -i $file.FullName `
        -vn -ac 1 -ar 16000 -c:a libmp3lame -b:a 32k `
        -f segment -segment_time ($AudioMinutes * 60) -reset_timestamps 1 `
        (Join-Path $outDir 'audio-%02d.mp3')
    if ($LASTEXITCODE -ne 0) { throw 'The audio step failed.' }

    # ------------------------------------------------------------------- video
    if ($AlsoVideo) {
        Step 'Making the video pieces'
        # Pick a bitrate that keeps each piece under the cap, with room to spare.
        $audioKbit = 48
        $budgetKbit = [math]::Floor(($MaxMB * 0.90 * 8 * 1024) / ($VideoMinutes * 60))
        $videoKbit = [math]::Max(200, $budgetKbit - $audioKbit)
        Say ('  One piece every ' + $VideoMinutes + ' minutes, aiming at ' +
             $MaxMB + ' MB each. This part is slow.')
        & $ffmpeg -hide_banner -loglevel error -y -i $file.FullName `
            -vf "scale='min(1280,iw)':-2" -r 15 `
            -c:v libx264 -preset veryfast -b:v ('{0}k' -f $videoKbit) `
            -maxrate ('{0}k' -f ([math]::Floor($videoKbit * 1.3))) `
            -bufsize ('{0}k' -f ($videoKbit * 2)) `
            -pix_fmt yuv420p -c:a aac -b:a ('{0}k' -f $audioKbit) -ac 1 `
            -f segment -segment_time ($VideoMinutes * 60) -reset_timestamps 1 `
            -movflags '+faststart' `
            (Join-Path $outDir 'video-%02d.mp4')
        if ($LASTEXITCODE -ne 0) { Oops 'The video step failed. The audio pieces are still fine.' }
    }

    # -------------------------------------------------------------- the result
    Step 'Done'
    $pieces = @(Get-ChildItem -Path $outDir -File | Sort-Object Name)
    $total = 0
    foreach ($p in $pieces) {
        $mb = [math]::Round($p.Length / 1MB, 1)
        $total += $mb
        $flag = ''
        if ($mb -gt $MaxMB) { $flag = '   ** too big, run it again with fewer minutes **' }
        Say ('  ' + $p.Name.PadRight(18) + ($mb.ToString() + ' MB').PadLeft(10) + $flag)
    }
    Say ''
    Say ('  ' + $pieces.Count + ' files, ' + [math]::Round($total, 1) + ' MB in total')
    Say ('  They are in: ' + $outDir)
    Say ''
    Say '  Upload the audio files first, in order, audio-00 to begin with.'
    Say '  They are all that is needed for the transcript.'
}
catch {
    Oops $_.Exception.Message
}
finally {
    Write-Host ''
    Read-Host 'Press Enter to close'
}

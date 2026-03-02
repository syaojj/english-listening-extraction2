$ErrorActionPreference = 'Stop'
$v = '7.5'
$d = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$src = $PSScriptRoot
$parent = Split-Path -Parent $src
$dest = Join-Path $parent ("English Listening Data Extraction2_backup_v" + $v + "_" + $d)
robocopy $src $dest /E /XD node_modules .git /NFL /NDL /NJH /NJS
Write-Output "BACKUP_DEST=$dest"

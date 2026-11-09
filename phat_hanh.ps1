param(
  [Parameter(Mandatory=$true)][string]$Dotm,
  [Parameter(Mandatory=$true)][string]$Version,
  [string]$Notes = ""
)
# Phat hanh ban moi: chep file, dong goi, ghi version.json, commit + push.
$ErrorActionPreference = 'Stop'
$repo = 'anhtu604/addin-the-thuc-van-ban'
$root = $PSScriptRoot
if ($Version -notmatch '^\d{4}\.\d{1,2}\.\d{1,2}\.\d+$') { throw "Phien ban phai dang 2026.11.09.3" }
$old = (Get-Content (Join-Path $root 'version.json') -Raw | ConvertFrom-Json).version
$a = $old.Split('.') | % { [int]$_ }; $b = $Version.Split('.') | % { [int]$_ }
$newer = $false; for ($i=0; $i -lt 4; $i++) { if ($b[$i] -gt $a[$i]) { $newer=$true; break }; if ($b[$i] -lt $a[$i]) { break } }
if (-not $newer) { throw "Phien ban $Version khong lon hon ban dang phat hanh $old" }

$distName = "ChuanHoa_VanBan_Rieng_$Version.dotm"
Copy-Item $Dotm (Join-Path $root "dist\$distName") -Force
Copy-Item $Dotm (Join-Path $root 'bo_cai_dat\ChuanHoa_VanBan_Rieng.dotm') -Force
$zip = Join-Path $root 'Bo_Cai_Dat_AddIn.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $root 'bo_cai_dat\*') -DestinationPath $zip

$sha = (Get-FileHash (Join-Path $root "dist\$distName") -Algorithm SHA256).Hash.ToLower()
$manifest = [ordered]@{
  version     = $Version
  word_url    = "https://raw.githubusercontent.com/$repo/main/dist/$distName"
  word_sha256 = $sha
  notes_ascii = $Notes
}
$json = $manifest | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $root 'version.json'), $json + "`n", (New-Object Text.UTF8Encoding $false))

git -C $root add -A
git -C $root commit -m "Phat hanh $Version"
git -C $root push origin main
"Da phat hanh $Version (sha256 $sha)"

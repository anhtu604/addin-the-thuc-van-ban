param(
  [Parameter(Mandatory=$true)][string]$Version,
  [Parameter(Mandatory=$true)][string]$Notes
)
# Dung va phat hanh ban moi bang MOT lenh (can Windows + Word + Python + git).
# Dieu kien: dang o nhanh main, da bat "Trust access to the VBA project object model".
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$ng   = Join-Path $root 'ma_nguon'
$repo = 'anhtu604/addin-the-thuc-van-ban'
$tmp  = Join-Path $env:TEMP 'addin_build'

$nhanh = (git -C $root rev-parse --abbrev-ref HEAD).Trim()
if ($nhanh -ne 'main') { throw "Dang o nhanh '$nhanh'. Hay chuyen sang main (da gop ma moi) roi chay lai." }
if (-not (Get-Command python -ErrorAction SilentlyContinue)) { throw "Khong thay python trong PATH." }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw "Hay dong het cua so Word truoc khi dung ban." }

New-Item -ItemType Directory -Force $tmp | Out-Null
# Xoa ket qua cu: neu buoc nap ma that bai, khong de file cu bi coi la ban moi.
Remove-Item "$tmp\a.bas","$tmp\b.bas","$tmp\c.bas","$tmp\tam.dotm","$tmp\final.dotm" -Force -ErrorAction SilentlyContinue
function Py { param([string[]]$a) & python @a; if ($LASTEXITCODE -ne 0) { throw "python $($a -join ' ') loi" } }

Push-Location $ng
try {
  Py 'va_1_the_thuc.py','Module1_ban_goc_2026.11.07.5.bas',"$tmp\a.bas"
  Py 'va_2_ban_rieng.py',"$tmp\a.bas","$tmp\b.bas"
  Py 'va_3_tien_ich.py',"$tmp\b.bas","$tmp\c.bas"
  Py 'va_4_phat_hanh.py',"$tmp\c.bas","$ng\Module1.bas",$repo,$Version
  & "$ng\nap_ma_vao_dotm.ps1" -Src "$ng\khung_goc_2026.11.07.5.dotm" -Bas "$ng\Module1.bas" -Out "$tmp\tam.dotm"
  if (-not (Test-Path "$tmp\tam.dotm")) { throw "Nap ma vao .dotm that bai." }
  Py 'va_5_ribbon.py',"$tmp\tam.dotm","$tmp\final.dotm"
} finally { Pop-Location }

Write-Host "Da dung xong $tmp\final.dotm. Hay MO THU file nay trong Word (chuan hoa 1 van ban mau) truoc khi phat hanh."
$ok = Read-Host "Da thu OK, phat hanh luon? (go Y de tiep tuc)"
if ($ok -ne 'Y') { Write-Host "Dung lai, chua phat hanh."; return }
& "$root\phat_hanh.ps1" -Dotm "$tmp\final.dotm" -Version $Version -Notes $Notes

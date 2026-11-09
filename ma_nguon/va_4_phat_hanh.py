"""Ban phat hanh cua Vuong Dinh Anh Tu.
python va_phat_hanh.py <vao.bas> <ra.bas> <owner/repo> <phien_ban>
- Bo dang ky may voi may chu cua tac gia goc (da duoc tac gia dong y).
- Ghi cong: nguoi duy tri Vuong Dinh Anh Tu, dua tren add-in cua VPDU xa Thanh Phu.
- Kenh tu cap nhat: https://raw.githubusercontent.com/<owner/repo>/main/version.json
"""
import re, sys
src, dst, repo, ver = sys.argv[1:5]
s = open(src, encoding='utf-8').read()

def rep(o, n, count=1):
    global s
    c = s.count(o)
    if c != count:
        raise SystemExit(f'found {c}x want {count}: {o[:70]!r}')
    s = s.replace(o, n)

# 1) license / machine registration removed
s, n = re.subn(r'Private Function LicApiUrl\(\) As String\n.*?\nEnd Function\n', '', s, flags=re.S)
assert n == 1, 'LicApiUrl'
s, n = re.subn(r'Private Function CheckLicenseActivation\(\) As Boolean\n.*?\nEnd Function\n',
               'Private Function CheckLicenseActivation() As Boolean\nCheckLicenseActivation = True\nEnd Function\n',
               s, flags=re.S)
assert n == 1, 'CheckLicenseActivation'

# 2) attribution
credit = '"Duy tri: Vuong Dinh Anh Tu - dua tren add-in cua VPDU xa Thanh Phu"'
rep('"Created by VPDU xa Thanh Phu"', credit, 2)

# 3) update channel + version
s, n = re.subn(r'CN_PhienBan = "[^"]*"', f'CN_PhienBan = "{ver}"', s)
assert n == 1
rep('CN_ManifestUrl = ""', f'CN_ManifestUrl = "https://raw.githubusercontent.com/{repo}/main/version.json"')
rep('"Khong ket noi duoc may chu cap nhat (GitHub). Kiem tra mang roi thu lai."',
    '"Khong ket noi duoc kenh cap nhat (GitHub " & "' + repo + '" & "). Kiem tra mang roi thu lai."')

assert 'script.google.com' not in s
open(dst, 'w', encoding='utf-8').write(s)
print('phat hanh ok', ver, repo)

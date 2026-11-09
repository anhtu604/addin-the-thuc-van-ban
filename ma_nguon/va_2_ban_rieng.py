import sys
s=open(sys.argv[1],encoding='utf-8').read()
def rep(o,n):
    global s
    assert s.count(o)==1,o[:60]; s=s.replace(o,n)
rep('CN_PhienBan = "2026.11.07.5"','CN_PhienBan = "2026.11.09.1"')
rep('CN_ManifestUrl = "https://github.com/CuongBC195/add-in-office/releases/latest/download/version.json"','CN_ManifestUrl = ""')
rep('CN_DungFileAddIn = (n = "chuanhoa_vanbandang.dotm")','CN_DungFileAddIn = (n = "chuanhoa_vanban_rieng.dotm")')
rep('''Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
''','''Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
If Len(CN_ManifestUrl()) = 0 Then
daDocDuoc = True
If hoiKhiKhongCo Then MsgBox "Ban quan ly rieng - phien ban " & CN_PhienBan() & "." & vbCrLf & "Ban nay KHONG tu cap nhat tu kenh cua ban goc. Khi co ban sua moi, cai lai bang bo cai dat.", vbInformation, "Phien ban Add-in"
Exit Sub
End If
''')

rep("""If Not CheckLicenseActivation() Then Exit Sub
End If
Application.ScreenUpdating = False
On Error Resume Next
Dim doc As Document""","""If Not CheckLicenseActivation() Then Exit Sub
End If
Application.ScreenUpdating = False
On Error Resume Next
Application.UndoRecord.StartCustomRecord "Chuan hoa the thuc"
Dim doc As Document""")
rep("""strWarnings = KiemTra_LogicTheThuc(doc)
On Error Resume Next
Application.ScreenUpdating = True""","""strWarnings = KiemTra_LogicTheThuc(doc)
On Error Resume Next
Application.UndoRecord.EndCustomRecord
Application.ScreenUpdating = True""")
rep(""""Van ban dat chuan 100%, khong phat hien loi logic." & vbCrLf & vbCrLf & _""",""""Van ban dat chuan, khong phat hien loi logic." & vbCrLf & "Muon tra lai nhu cu: bam Ctrl+Z mot lan." & vbCrLf & vbCrLf & _""")
rep("""strWarnings & vbCrLf & vbCrLf & _
"Created by VPDU xa Thanh Phu", vbExclamation""","""strWarnings & vbCrLf & vbCrLf & "Muon tra lai nhu cu: bam Ctrl+Z mot lan." & vbCrLf & vbCrLf & _
"Created by VPDU xa Thanh Phu", vbExclamation""")

open(sys.argv[2],'w',encoding='utf-8').write(s); print('clone patch ok')

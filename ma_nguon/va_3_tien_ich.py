import sys
s = open(sys.argv[1], encoding='utf-8').read()
def rep(o, n, count=1):
    global s
    c = s.count(o)
    if c != count:
        raise SystemExit(f'found {c}x want {count}: {o[:70]!r}')
    s = s.replace(o, n)

rep("Private mTTQPPL As Boolean\n", "Private mTTQPPL As Boolean\nPrivate mTTNotes As String\n")
rep('CN_PhienBan = "2026.11.09.1"', 'CN_PhienBan = "2026.11.09.2"')

# reset notes + typo fix at start
rep("mTTQPPL = TT_DetectQPPL(doc)\n", "mTTQPPL = TT_DetectQPPL(doc)\nmTTNotes = \"\"\n")
rep("Call SafeReplaceSoftBreaksInRange(doc.Content)\n",
    "Call SafeReplaceSoftBreaksInRange(doc.Content)\nCall TT_SuaLoiGo(doc)\n")

# canonical Quoc hieu text
rep("""If InStr(1, pAsc, "CONG HOA XA HOI") > 0 Then
pCell.Range.Case = wdUpperCase""",
"""If InStr(1, pAsc, "CONG HOA XA HOI") > 0 Then
If Len(pAsc) <= 45 And CleanParaText(pCell.Range.Text) <> TT_TextQuocHieu() Then Call SafeSetParagraphText(doc, pCell, TT_TextQuocHieu())
pCell.Range.Case = wdUpperCase""")
# "UỶ" -> "ỦY" in organ name lines (spelling used in ND 30 appendix)
rep("""pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = TT_N("ORGAN_SIZE")""",
"""pCell.Format.Alignment = wdAlignParagraphCenter
Call TT_ChuanUY(pCell.Range)
pCell.Range.Font.Size = TT_N("ORGAN_SIZE")""")

# widow/orphan control for body text
rep(""".SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
If TT_B("LEGAL_BASIS_ITALIC")""",
""".SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
.WidowControl = True
End With
If TT_B("LEGAL_BASIS_ITALIC")""")

# page-layout helpers after all formatting, before checks
rep("""Call TT_DebugLog("BUOC 6 bat dau")""",
"""Call TT_XuLyBangNoiDung(doc)
Call TT_GiuChuKyCungTrang(doc)
Call TT_EpTrang(doc)
Call TT_DebugLog("BUOC 6 bat dau")""")

# show notes
rep("""TT_ReportHeader() & vbCrLf & vbCrLf & _""", """TT_ReportHeader() & TT_NotesText() & vbCrLf & vbCrLf & _""", 2)
rep('''TT_ReportHeader() & vbCrLf & "Canh bao:" & strWarnings''',
    '''TT_ReportHeader() & vbCrLf & "Canh bao:" & strWarnings & vbCrLf & "Tu dong:" & mTTNotes''')

NEW = r'''
Private Sub TT_Note(ByVal msg As String)
mTTNotes = mTTNotes & vbCrLf & "- " & msg
End Sub
Private Function TT_NotesText() As String
TT_NotesText = ""
If Len(mTTNotes) > 0 Then TT_NotesText = vbCrLf & vbCrLf & "Da tu dong xu ly:" & mTTNotes
End Function
Private Function TT_IsLetterCode(ByVal c As Long) As Boolean
TT_IsLetterCode = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 192 And c <= 7929 And c <> 215 And c <> 247)
End Function
Private Function TT_SuaCachTrongDoan(ByVal doc As Document, ByVal p As Paragraph) As Boolean
On Error Resume Next
TT_SuaCachTrongDoan = False
If p.Range.Fields.Count > 0 Or p.Range.InlineShapes.Count > 0 Then Exit Function
Dim t As String, n As Long, i As Long, j As Long, base As Long, c As String, nx As String, lastIdx As Long
t = p.Range.Text
If Len(t) <> p.Range.End - p.Range.Start Then Exit Function
If p.Range.ShapeRange.Count > 0 Then Exit Function
n = Len(t)
Do While n > 0
c = Mid(t, n, 1)
If c = vbCr Or c = Chr(7) Or c = Chr(11) Then n = n - 1 Else Exit Do
Loop
If n < 2 Then Exit Function
lastIdx = n
base = p.Range.Start
For i = n To 1 Step -1
c = Mid(t, i, 1)
If c = " " Then
j = i
Do While j < n And Mid(t, j + 1, 1) = " "
j = j + 1
Loop
nx = ""
If j < n Then nx = Mid(t, j + 1, 1)
Dim doDel As Boolean: doDel = False
If Len(nx) > 0 And InStr(1, ",;:!?)", nx) > 0 And i > 1 Then doDel = True
If nx = "." And j + 1 = lastIdx And i > 1 Then
If Mid(t, i - 1, 1) <> "." And Mid(t, i - 1, 1) <> " " Then doDel = True
End If
Dim k As Long: k = i
Do While k > 1 And Mid(t, k - 1, 1) = " "
k = k - 1
Loop
If k > 1 And Mid(t, k - 1, 1) = "(" Then doDel = True
If doDel Then
If doc.Range(base + i - 1, base + i).Text = " " Then
doc.Range(base + i - 1, base + i).Delete
TT_SuaCachTrongDoan = True
End If
End If
ElseIf (c = "," Or c = ";") And i < n Then
Dim cc As Long: cc = AscW(Mid(t, i + 1, 1))
If cc < 0 Then cc = cc + 65536
If TT_IsLetterCode(cc) Then
If doc.Range(base + i - 1, base + i).Text = c Then
doc.Range(base + i, base + i).InsertAfter " "
TT_SuaCachTrongDoan = True
End If
End If
End If
Next i
End Function
Private Sub TT_SuaLoiGo(ByVal doc As Document)
On Error Resume Next
Dim p As Paragraph, hit As Boolean, idx As Long, total As Long
total = doc.Paragraphs.Count
If total > 3000 Then Exit Sub
Set p = doc.Paragraphs(1)
For idx = 1 To total
If p Is Nothing Then Exit For
If TT_SuaCachTrongDoan(doc, p) Then hit = True
Set p = p.Next
Next idx
If hit Then Call TT_Note("Sua dau cach quanh dau cau (bo cach truoc , ; : ! ? ) va dau cham cuoi doan, bo cach sau dau mo ngoac, them cach sau , ;).")
End Sub
Private Sub TT_ChuanUY(ByVal rng As Range)
On Error Resume Next
Dim t As String, pos As Long, base As Long
t = rng.Text
base = rng.Start
pos = InStrRev(t, "U" & ChrW(7926))
Do While pos > 0
rng.Document.Range(base + pos - 1, base + pos + 1).Text = ChrW(7910) & "Y"
If pos <= 1 Then Exit Do
pos = InStrRev(t, "U" & ChrW(7926), pos - 1)
Loop
End Sub
Private Function TT_FindFooterTable(ByVal doc As Document) As Table
On Error Resume Next
Dim t As Table
Set TT_FindFooterTable = Nothing
For Each t In doc.Tables
If IsFooterSignatureTable(t, ToAscii(t.Range.Text)) Then Set TT_FindFooterTable = t
Next t
End Function
Private Function TT_IsLayoutTable(ByVal doc As Document, ByVal t As Table) As Boolean
Dim a As String
a = ToAscii(t.Range.Text)
TT_IsLayoutTable = True
If t.Range.Start = doc.Tables(1).Range.Start Then Exit Function
If InStr(1, a, "CONG HOA XA HOI") > 0 Or InStr(1, a, "DANG CONG SAN") > 0 Then Exit Function
If IsKinhGuiTable(t, a) Or IsFooterSignatureTable(t, a) Then Exit Function
TT_IsLayoutTable = False
End Function
Private Function TT_PageAt(ByVal doc As Document, ByVal pos As Long) As Long
TT_PageAt = doc.Range(pos, pos).Information(wdActiveEndPageNumber)
End Function
Private Function TT_PrevTextPara(ByVal doc As Document, ByVal pos As Long) As Paragraph
On Error Resume Next
Set TT_PrevTextPara = Nothing
If pos <= 0 Then Exit Function
Dim p As Paragraph
Set p = doc.Range(pos - 1, pos - 1).Paragraphs(1)
Do While Not p Is Nothing
If Not p.Range.Information(wdWithInTable) Then
If Len(CleanParaText(p.Range.Text)) > 0 Then
Set TT_PrevTextPara = p
Exit Function
End If
End If
Set p = p.Previous
Loop
End Function
Private Sub TT_XuLyBangNoiDung(ByVal doc As Document)
On Error Resume Next
Dim t As Table, n As Integer
For Each t In doc.Tables
If Not TT_IsLayoutTable(doc, t) And t.Rows.Count >= 2 Then
t.Rows(1).HeadingFormat = True
t.Rows.AllowBreakAcrossPages = False
t.PreferredWidthType = wdPreferredWidthPercent
t.PreferredWidth = 100
t.Rows.LeftIndent = 0
n = n + 1
End If
Next t
If n > 0 Then Call TT_Note("Bang noi dung (" & n & "): lap lai dong tieu de khi sang trang, khong cat ngang dong, vua khit le.")
End Sub
Private Sub TT_GiuChuKyCungTrang(ByVal doc As Document)
On Error Resume Next
Dim ft As Table
Set ft = TT_FindFooterTable(doc)
If ft Is Nothing Then Exit Sub
ft.Rows.AllowBreakAcrossPages = False
Dim p1 As Paragraph
Set p1 = TT_PrevTextPara(doc, ft.Range.Start)
If p1 Is Nothing Then Exit Sub
Dim q As Paragraph
Set q = p1
Do While Not q Is Nothing
If q.Range.Start >= ft.Range.Start Then Exit Do
q.Format.KeepWithNext = True
Set q = q.Next
Loop
p1.Format.WidowControl = True
If p1.Range.ComputeStatistics(wdStatisticLines) <= 3 Then p1.Format.KeepTogether = True
End Sub
Private Sub TT_EpTrang(ByVal doc As Document)
On Error Resume Next
Dim ft As Table
Set ft = TT_FindFooterTable(doc)
If ft Is Nothing Then Exit Sub
doc.Repaginate
Dim pg0 As Long
pg0 = doc.ComputeStatistics(wdStatisticPages)
If pg0 < 2 Then Exit Sub
Dim lastPg As Long
lastPg = TT_PageAt(doc, ft.Range.Start)
If lastPg < pg0 Then Exit Sub
Dim p As Paragraph, nLines As Long
For Each p In doc.Paragraphs
If p.Range.Start >= ft.Range.Start Then Exit For
If Not p.Range.Information(wdWithInTable) Then
If Len(CleanParaText(p.Range.Text)) > 0 Then
If TT_PageAt(doc, p.Range.End - 1) = lastPg Then nLines = nLines + p.Range.ComputeStatistics(wdStatisticLines)
End If
End If
Next p
If nLines > 6 Then Exit Sub
Dim chg1 As New Collection
For Each p In doc.Paragraphs
If p.Range.Start >= ft.Range.Start Then Exit For
If Not p.Range.Information(wdWithInTable) Then
If p.Format.SpaceBefore >= 6 And p.Format.SpaceAfter >= 6 Then
p.Format.SpaceAfter = 0
chg1.Add p
End If
End If
Next p
doc.Repaginate
If doc.ComputeStatistics(wdStatisticPages) < pg0 Then
Call TT_Note("Trang cuoi chi con " & nLines & " dong + chu ky: da thu khoang cach doan con 6pt (dung muc toi thieu ND 30) de gon " & pg0 - 1 & " trang.")
Exit Sub
End If
Dim big As Single
big = TT_N("BODY_SIZE")
Dim chg2 As New Collection
For Each p In doc.Paragraphs
If p.Range.Start >= ft.Range.Start Then Exit For
If Not p.Range.Information(wdWithInTable) Then
If p.Range.Font.Size = big Then
p.Range.Font.Size = big - 1
chg2.Add p
End If
End If
Next p
doc.Repaginate
If doc.ComputeStatistics(wdStatisticPages) < pg0 Then
Call TT_Note("Trang cuoi chi con " & nLines & " dong + chu ky: da thu khoang cach doan 6pt va ha co chu noi dung " & big & " -> " & big - 1 & " (trong khung cho phep) de gon " & pg0 - 1 & " trang.")
Exit Sub
End If
For Each p In chg2
p.Range.Font.Size = big
Next p
For Each p In chg1
p.Format.SpaceAfter = 6
Next p
doc.Repaginate
If nLines <= 2 Then Call TT_Note("Trang cuoi chi co " & nLines & " dong noi dung + chu ky; khong the thu gon ma van dung the thuc. Can nhac rut gon noi dung.")
End Sub
'''
anchor = "Private Function IsFooterBreakerLine("
assert s.count(anchor) == 1
s = s.replace(anchor, NEW.strip('\n') + "\n" + anchor)
# ---------- V/v (trich yeu cong van): co 12, nho hon dong So (13) theo mau ND 30 ----------
rep('Case "CV_SUMMARY_SIZE": TT_RawN = 13', 'Case "CV_SUMMARY_SIZE": TT_RawN = 12', 4)

# ---------- Kinh gui: tu 2 don vi ttro len, moi don vi mot dong ----------
rep("""Call SafeReplaceSoftBreaksInRange(cKG.Range)
On Error Resume Next
cKG.Range.Font.Bold = False
cKG.Range.Font.Italic = False
""", """Call SafeReplaceSoftBreaksInRange(cKG.Range)
Dim kgTxtCu As String, kgTxtMoi As String, pKgT As Paragraph
kgTxtCu = ""
For Each pKgT In cKG.Range.Paragraphs
If Len(CleanParaText(pKgT.Range.Text)) > 0 Then
If Len(kgTxtCu) > 0 Then kgTxtCu = kgTxtCu & vbCr
kgTxtCu = kgTxtCu & CleanParaText(pKgT.Range.Text)
End If
Next pKgT
kgTxtMoi = TT_TachDonViNhan(kgTxtCu)
If kgTxtMoi <> kgTxtCu Then cKG.Range.Text = kgTxtMoi
On Error Resume Next
cKG.Range.Font.Bold = False
cKG.Range.Font.Italic = False
""")
NEW2 = """
Private Function TT_TachDonViNhan(ByVal txt As String) As String
Dim ln() As String, i As Long, j As Long, outT As String, cur As String, piece As String, ch As String
ln = Split(txt, vbCr)
outT = ""
For i = 0 To UBound(ln)
cur = Trim(Replace(ln(i), ChrW(160), " "))
If Len(cur) > 0 Then
piece = ""
For j = 1 To Len(cur)
ch = Mid(cur, j, 1)
piece = piece & ch
If ch = ";" Then
If Len(Trim(Mid(cur, j + 1))) > 0 Then
outT = outT & Trim(piece) & vbCr
piece = ""
End If
End If
Next j
If Len(Trim(piece)) > 0 Then outT = outT & Trim(piece) & vbCr
End If
Next i
If Len(outT) > 0 Then outT = Left(outT, Len(outT) - 1)
TT_TachDonViNhan = outT
End Function
"""
anchor2 = "Private Function IsFooterBreakerLine("
assert s.count(anchor2) == 1
s = s.replace(anchor2, NEW2.strip('\n') + "\n" + anchor2)
open(sys.argv[2], 'w', encoding='utf-8').write(s)
print('tien ich ok')

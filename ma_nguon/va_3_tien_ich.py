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
# ---------- Kinh gui >= 2 don vi: "Kinh gui:" mot dong rieng, moi don vi mot dong ben duoi ----------
rep("""If HasKinhGuiOrTrinh(cAscKG) Then
seenKinhGuiCell = True
""", """If HasKinhGuiOrTrinh(cAscKG) Then
If tbl.Range.Cells.Count = 1 Then GoTo NextCellKG
seenKinhGuiCell = True
""")
rep("""NextCellKG:
Next cKG
End If
""", """NextCellKG:
Next cKG
Call TT_KinhGuiNhieuDonVi(doc, tbl)
End If
""")
NEW3 = """
Private Function TT_LaGachDau(ByVal ch As String) As Boolean
TT_LaGachDau = False
If Len(ch) = 0 Then Exit Function
Select Case AscW(ch)
Case 45, 43, 42, 183, 8210, 8211, 8212, 8226, 8259, 8722: TT_LaGachDau = True
End Select
End Function
Private Function TT_KinhGuiDinhDang(ByVal items As String) As String
Dim ln() As String, i As Long, t As String, cleaned As String, outT As String, n As Long
ln = Split(items, vbCr)
cleaned = ""
For i = 0 To UBound(ln)
t = Trim(Replace(ln(i), ChrW(160), " "))
Do While Len(t) > 0
If TT_LaGachDau(Left(t, 1)) Then t = Trim(Mid(t, 2)) Else Exit Do
Loop
Do While Len(t) > 0
If InStr(".;, ", Right(t, 1)) > 0 Then t = Left(t, Len(t) - 1) Else Exit Do
Loop
If Len(t) > 0 Then
If Len(cleaned) > 0 Then cleaned = cleaned & vbCr
cleaned = cleaned & t
End If
Next i
ln = Split(cleaned, vbCr)
n = UBound(ln)
outT = ""
For i = 0 To n
If i > 0 Then outT = outT & vbCr
outT = outT & "- " & ln(i) & IIf(i = n, ".", ";")
Next i
TT_KinhGuiDinhDang = outT
End Function
Private Sub TT_KinhGuiNhieuDonVi(ByVal doc As Document, ByVal tbl As Table)
On Error Resume Next
Dim nCells As Long, leftW As Single, prefixT As String, rawItems As String
Dim p As Paragraph, c As Cell, t As String, first As Boolean, i As Long
nCells = tbl.Range.Cells.Count
rawItems = ""
first = True
If nCells = 2 And tbl.Rows.Count = 1 Then
leftW = tbl.Cell(1, 1).Width
prefixT = KinhGuiPrefixFor(ToAscii(tbl.Cell(1, 1).Range.Text))
For Each p In tbl.Cell(1, 2).Range.Paragraphs
t = CleanParaText(p.Range.Text)
If Len(t) > 0 Then
If Len(rawItems) > 0 Then rawItems = rawItems & vbCr
rawItems = rawItems & t
End If
Next p
ElseIf nCells = 1 Then
leftW = tbl.Cell(1, 1).Range.Paragraphs(1).LeftIndent
If tbl.Cell(1, 1).Range.Paragraphs.Count >= 2 Then leftW = tbl.Cell(1, 1).Range.Paragraphs(2).LeftIndent - MillimetersToPoints(3)
For Each p In tbl.Cell(1, 1).Range.Paragraphs
t = CleanParaText(p.Range.Text)
If Len(t) > 0 Then
If first Then
prefixT = KinhGuiPrefixFor(ToAscii(t))
first = False
Else
If Len(rawItems) > 0 Then rawItems = rawItems & vbCr
rawItems = rawItems & t
End If
End If
Next p
Else
Exit Sub
End If
If Len(rawItems) = 0 Then Exit Sub
rawItems = TT_TachDonViNhan(rawItems)
If UBound(Split(rawItems, vbCr)) < 1 Then Exit Sub
If leftW < MillimetersToPoints(20) Then leftW = MillimetersToPoints(52)
If nCells = 2 Then tbl.Cell(1, 1).Merge MergeTo:=tbl.Cell(1, 2)
Set c = tbl.Cell(1, 1)
c.Range.Text = prefixT & vbCr & TT_KinhGuiDinhDang(rawItems)
c.Borders.Enable = False
c.VerticalAlignment = wdCellAlignVerticalTop
c.Range.Font.Name = "Times New Roman"
c.Range.Font.Size = 14
c.Range.Font.Bold = False
c.Range.Font.Italic = False
c.Range.Font.ColorIndex = wdBlack
c.Range.Font.Underline = wdUnderlineNone
i = 0
For Each p In c.Range.Paragraphs
i = i + 1
With p.Format
.RightIndent = 0
.SpaceBefore = 3
.SpaceAfter = 3
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
If i = 1 Then
.Alignment = wdAlignParagraphLeft
.LeftIndent = leftW
.FirstLineIndent = 0
Else
.Alignment = wdAlignParagraphLeft
.LeftIndent = leftW + MillimetersToPoints(3)
.FirstLineIndent = -MillimetersToPoints(3)
End If
End With
If i = 1 Then p.Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
Next p
' ND 30: cac gach dau dong thang hang duoi dau hai cham cua "Kinh gui:".
' Dat "Kinh gui:" ket thuc dung tai cot gach dau dong (leftW).
Dim wKG As Single, r1 As Range, r2 As Range
Set r1 = c.Range.Paragraphs(1).Range.Characters.First
Set r2 = c.Range.Paragraphs(1).Range.Characters.Last
r1.Collapse wdCollapseStart
r2.Collapse wdCollapseStart
wKG = r2.Information(wdHorizontalPositionRelativeToPage) - r1.Information(wdHorizontalPositionRelativeToPage)
If wKG < 20 Or wKG > 200 Then wKG = Len(prefixT) * 6.2
If leftW - wKG < 0 Then
leftW = wKG
For i = 2 To c.Range.Paragraphs.Count
c.Range.Paragraphs(i).LeftIndent = leftW + MillimetersToPoints(3)
c.Range.Paragraphs(i).FirstLineIndent = -MillimetersToPoints(3)
Next i
End If
c.Range.Paragraphs(1).LeftIndent = leftW - wKG
Call TT_Note("Kinh gui co tu 2 don vi tro len: de 'Kinh gui:' mot dong rieng, moi don vi xuong mot dong (ND 30).")
End Sub
"""
anchor3 = "Private Function IsFooterBreakerLine("
assert s.count(anchor3) == 1
s = s.replace(anchor3, NEW3.strip('\n') + "\n" + anchor3)
# ---------- Kinh gui: nhan don vi cuoi khong co dau cau ----------
# Dong don vi truoc ket thuc bang ";" hoac "," (danh sach con tiep) thi dong
# ke tiep duoc nhan la don vi cuoi, du khong co gach dau dong / dau cau.
# VD: "Trung tam ... thanh pho;" + "Benh vien ... Hai Duong".
rep("""Dim scanCount As Integer: scanCount = 0
""", """Dim scanCount As Integer: scanCount = 0
Dim kgPrevSemi As Boolean: kgPrevSemi = False
""")
rep("""If IsKinhGuiStart(sRaw, sAsc) Then
Set pStart = p
Set pEnd = p
foundStart = True
""", """If IsKinhGuiStart(sRaw, sAsc) Then
Set pStart = p
Set pEnd = p
foundStart = True
kgPrevSemi = TT_KetThucChamPhay(sRaw)
""")
rep("""If IsKinhGuiRecipientItem(sRaw, sAsc) Then
Set pEnd = p
Else
Exit For
End If
""", """If IsKinhGuiRecipientItem(sRaw, sAsc) Then
Set pEnd = p
kgPrevSemi = TT_KetThucChamPhay(sRaw)
ElseIf kgPrevSemi And TT_LaDonViCuoi(sRaw, sAsc) Then
Set pEnd = p
Exit For
Else
Exit For
End If
""")
rep("""Dim lastPInTbl As Paragraph
Set lastPInTbl = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count)
""", """Dim lastPInTbl As Paragraph
Set lastPInTbl = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count)
Dim kgAbsSemi As Boolean
kgAbsSemi = TT_KetThucChamPhay(TT_DongCuoiCo(tbl.Range))
""")
rep("""If IsKinhGuiRecipientItem(sAfterRaw, sAfterAsc) Then
If pAbsorbStart Is Nothing Then Set pAbsorbStart = pAfterTbl
Set pAbsorbEnd = pAfterTbl
Else
Exit Do
End If
""", """If IsKinhGuiRecipientItem(sAfterRaw, sAfterAsc) Then
If pAbsorbStart Is Nothing Then Set pAbsorbStart = pAfterTbl
Set pAbsorbEnd = pAfterTbl
kgAbsSemi = TT_KetThucChamPhay(sAfterRaw)
ElseIf kgAbsSemi And TT_LaDonViCuoi(sAfterRaw, sAfterAsc) Then
If pAbsorbStart Is Nothing Then Set pAbsorbStart = pAfterTbl
Set pAbsorbEnd = pAfterTbl
Exit Do
Else
Exit Do
End If
""")
NEW4 = """
Private Function TT_KetThucChamPhay(ByVal s As String) As Boolean
Dim t As String
t = Trim(Replace(CleanParaText(s), ChrW(160), " "))
TT_KetThucChamPhay = (Right(t, 1) = ";" Or Right(t, 1) = ",")
End Function
Private Function TT_DongCuoiCo(ByVal rng As Range) As String
Dim p As Paragraph, t As String
TT_DongCuoiCo = ""
For Each p In rng.Paragraphs
t = CleanParaText(p.Range.Text)
If Len(Trim(t)) > 0 Then TT_DongCuoiCo = t
Next p
End Function
Private Function TT_LaDonViCuoi(ByVal sRaw As String, ByVal sAsc As String) As Boolean
Dim t As String
TT_LaDonViCuoi = False
t = Trim(CleanParaText(sRaw))
If Len(t) = 0 Or Len(t) > 150 Then Exit Function
If Right(t, 1) = ":" Then Exit Function
Dim a As String
a = Trim(sAsc)
If a Like "DE *" Or a Like "NHAM *" Or a Like "THUC HIEN*" Or a Like "CAN CU*" Or a Like "THEO *" Or _
a Like "TRA LOI*" Or a Like "PHUC DAP*" Or a Like "HIEN NAY*" Or a Like "TRONG *" Or a Like "NGAY *" Or _
a Like "VAO *" Or a Like "SAU *" Or a Like "TIEP THEO*" Or a Like "QUA *" Or a Like "VE VIEC*" Or _
a Like "KINH DE NGHI*" Or a Like "XIN *" Or a Like "TRAN TRONG*" Or a Like "DUOC *" Or a Like "TU *" Then Exit Function
If IsSeparatorLine(t) Or IsKinhGuiStart(t, sAsc) Or IsPartyDocTitle(sAsc) Then Exit Function
If IsStrictBodyOrHeadingStart(t, sAsc) Or IsBodyStart(t, sAsc) Then Exit Function
If sAsc Like "NOI NHAN*" Or IsSignatureRoleOrAuth(sAsc) Then Exit Function
TT_LaDonViCuoi = True
End Function
"""
anchor4 = "Private Function IsFooterBreakerLine("
assert s.count(anchor4) == 1
s = s.replace(anchor4, NEW4.strip('\n') + "\n" + anchor4)

# ---- Cap nhat (Windows): khong ket im lang khi Word con chay an ----
# Ban cu: file .cmd cho den khi KHONG con WINWORD.EXE nao. Neu Word chay an
# (khong cua so) thi cho mai, file khong duoc thay, mo lai Word van ban cu va
# bam [Cap nhat] lai tai lai tu dau. Ban moi: script PowerShell cho dong Word,
# neu chi con Word chay an thi hoi dong no; chep xong kiem tra SHA-256 va bao
# "Da cai xong"; trong luc dang cho, bam [Cap nhat] bao "dang cho dong Word".
rep("msg = CN_TaiVaCai(url, sha)\n", "msg = CN_TaiVaCai(url, sha, ver)\n")
rep("Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String) As String\n",
    "Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String, Optional ByVal ver As String = \"\") As String\n")
rep("""Dim ans As VbMsgBoxResult
ans = MsgBox("Co phien ban moi: " & ver""",
"""If CN_DangChoCai() Then
If hoiKhiKhongCo Then MsgBox "Ban " & ver & " da tai ve va dang CHO BAN DONG HET Word de cai (dang dung " & CN_PhienBan() & ")." & vbCrLf & vbCrLf & _
"Hay luu va dong tat ca cua so Word. Khi cai xong se co thong bao 'Da cai xong', luc do mo lai Word.", vbInformation, "Cap nhat Add-in"
Exit Sub
End If
Dim ans As VbMsgBoxResult
ans = MsgBox("Co phien ban moi: " & ver""")

OLD_WIN = r'''Dim tgtW As String
tgtW = target
If Len(Environ("APPDATA")) > 0 And LCase(Left(tgtW, Len(Environ("APPDATA")))) = LCase(Environ("APPDATA")) Then
tgtW = "%APPDATA%" & Mid(tgtW, Len(Environ("APPDATA")) + 1)
End If
Dim cmdPath As String
cmdPath = Environ("TEMP") & "\cn_capnhat_WORD.cmd"
Dim f As Integer
f = FreeFile
Open cmdPath For Output As #f
Print #f, "@echo off"
Print #f, ":cho"
Print #f, "tasklist /FI ""IMAGENAME eq WINWORD.EXE"" 2>nul | find /I ""WINWORD.EXE"" >nul"
Print #f, "if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto cho)"
Print #f, "copy /Y ""%TEMP%\cn_capnhat_WORD.tmp"" """ & tgtW & """ >nul"
Print #f, "powershell -NoProfile -ExecutionPolicy Bypass -Command ""Unblock-File -LiteralPath '" & tgtW & "'"" >nul 2>&1"
Print #f, "del ""%TEMP%\cn_capnhat_WORD.tmp"" >nul 2>&1"
Print #f, "(goto) 2>nul & del ""%~f0"""
Close #f
Shell "cmd.exe /c """ & cmdPath & """", vbHide
CN_TaiVaCai = "Da tai ban moi. Hay DONG HET cua so Word: ban moi se tu cai dat, mo lai Word la dung duoc."
Exit Function
'''
PS_LINES = r'''$ErrorActionPreference = 'SilentlyContinue'
$tmp = Join-Path $env:TEMP 'cn_capnhat_WORD.tmp'
@@TGT@@
$sha = '@@SHA@@'
$ver = '@@VER@@'
Add-Type -AssemblyName System.Windows.Forms
$nl = [Environment]::NewLine
$tAn = $null
$daHoi = $false
while (@(Get-Process WINWORD).Count -gt 0) {
if (@(Get-Process WINWORD | Where-Object { $_.MainWindowHandle -ne 0 }).Count -gt 0) { $tAn = $null }
elseif ($tAn -eq $null) { $tAn = Get-Date }
elseif ((-not $daHoi) -and (((Get-Date) - $tAn).TotalSeconds -ge 15)) {
$daHoi = $true
$m = 'Cac cua so Word da dong nhung Word van con CHAY AN, nen chua cai duoc ban ' + $ver + ' cua Add-in The thuc.' + $nl + $nl + 'Dong Word dang chay an de cai ngay?' + $nl + '(Chon No neu ban dang tron thu/in an tu phan mem khac.)'
if ([System.Windows.Forms.MessageBox]::Show($m, 'Cap nhat Add-in', 'YesNo', 'Question') -eq 'Yes') { Get-Process WINWORD | Stop-Process -Force; Start-Sleep -Seconds 2 }
}
Start-Sleep -Seconds 2
}
if (Test-Path -LiteralPath $tmp) {
$ok = $false
for ($i = 0; ($i -lt 10) -and (-not $ok); $i++) {
Copy-Item -LiteralPath $tmp -Destination $tgt -Force
if ((Get-FileHash -LiteralPath $tgt -Algorithm SHA256).Hash.ToLower() -eq $sha) { $ok = $true } else { Start-Sleep -Seconds 2 }
}
if ($ok) {
Unblock-File -LiteralPath $tgt
Remove-Item -LiteralPath $tmp -Force
[void][System.Windows.Forms.MessageBox]::Show('Da cai xong Add-in The thuc ban ' + $ver + '.' + $nl + 'Mo lai Word de dung ban moi.', 'Cap nhat Add-in', 'OK', 'Information')
} else {
[void][System.Windows.Forms.MessageBox]::Show('Chua thay duoc file Add-in (file dang bi khoa). Hay dong het Word roi bam [Cap nhat] lai, hoac cai bang bo cai dat.', 'Cap nhat Add-in', 'OK', 'Warning')
}
}
Remove-Item -LiteralPath $PSCommandPath -Force'''
def vba_print(line):
    q = line.replace('"', '""')
    q = q.replace("@@SHA@@", '" & sha & "').replace("@@VER@@", '" & Replace(ver, "\'", "") & "')
    if line == "@@TGT@@":
        return "Print #f, tgtW"
    return 'Print #f, "' + q + '"'
NEW_WIN = r'''Dim tgtW As String
If Len(Environ("APPDATA")) > 0 And LCase(Left(target, Len(Environ("APPDATA")))) = LCase(Environ("APPDATA")) Then
tgtW = "$tgt = Join-Path $env:APPDATA '" & Replace(Mid(target, Len(Environ("APPDATA")) + 2), "'", "''") & "'"
Else
tgtW = "$tgt = '" & Replace(target, "'", "''") & "'"
End If
Dim psPath As String
psPath = Environ("TEMP") & "\cn_capnhat_WORD.ps1"
Dim f As Integer
f = FreeFile
Open psPath For Output As #f
''' + "\n".join(vba_print(l) for l in PS_LINES.split("\n")) + r'''
Close #f
Shell "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psPath & """", vbHide
CN_TaiVaCai = "Da tai ban " & ver & ". Ban moi CHI co hieu luc sau khi ban DONG HET cua so Word." & vbCrLf & vbCrLf & _
"Hay luu va dong tat ca cua so Word. Vai giay sau se co thong bao 'Da cai xong', luc do mo lai Word."
Exit Function
'''
rep(OLD_WIN, NEW_WIN)
rep("Private Function CN_TaiVaCai(", r'''Private Function CN_DangChoCai() As Boolean
CN_DangChoCai = False
#If Mac Then
#Else
On Error Resume Next
Dim wmi As Object, ds As Object, p As Object
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set ds = wmi.ExecQuery("Select ProcessId From Win32_Process Where Name='powershell.exe' And CommandLine Like '%cn_capnhat_WORD.ps1%'")
For Each p In ds
CN_DangChoCai = True
Exit For
Next p
On Error GoTo 0
#End If
End Function
Private Function CN_TaiVaCai(''')
open(sys.argv[2], 'w', encoding='utf-8').write(s)
print('tien ich ok')

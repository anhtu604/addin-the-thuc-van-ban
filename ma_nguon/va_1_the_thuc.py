import sys, io
src, dst = sys.argv[1], sys.argv[2]
s = open(src, encoding='utf-8').read()
# drop olevba separator line
if s.startswith('- - -'):
    s = s.split('\n', 1)[1]

def rep(old, new, count=1):
    global s
    n = s.count(old)
    if n != count:
        raise SystemExit(f'pattern found {n}x (want {count}): {old[:80]!r}')
    s = s.replace(old, new)

# ---------- E0: module state ----------
rep("Private mTTRibbon As Object\n",
    "Private mTTRibbon As Object\nPrivate mTTQPPL As Boolean\n")

# ---------- E1: strong detection overrides fixed block (Dang <-> nha nuoc) ----------
rep("""Else
mTTProfile = setting
mTTSource = "CHON"
End If
TT_Resolve = mTTProfile""",
"""ElseIf mTTDetectedStrong And Len(mTTDetected) > 0 And Len(mTTForceProfile) = 0 And _
((TT_RawS(mTTDetected, "MOTTO_KIND") = "DANG") <> (TT_RawS(setting, "MOTTO_KIND") = "DANG")) Then
mTTProfile = mTTDetected
mTTSource = "GHI_DE"
Else
mTTProfile = setting
mTTSource = "CHON"
End If
TT_Resolve = mTTProfile""")

rep("""Case "MAC_DINH": s = s & " (khong nhan dien duoc, dung khoi mac dinh)"
""",
"""Case "MAC_DINH": s = s & " (khong nhan dien duoc, dung khoi mac dinh)"
Case "GHI_DE": s = s & " (tu dong chuyen vi tieu ngu tren van ban khac khoi dang chon tren Ribbon)"
""")
rep("""TT_ReportHeader = s
End Function""",
"""If mTTQPPL Then s = s & vbCrLf & "Loai: VAN BAN QUY PHAM PHAP LUAT - trinh bay Chuong/Muc/Dieu theo Luat Ban hanh VBQPPL nam 2025 va Nghi dinh huong dan (can doi chieu van ban goc)."
TT_ReportHeader = s
End Function
Private Function TT_DetectQPPL(ByVal doc As Document) As Boolean
TT_DetectQPPL = False
If TT_S("MOTTO_KIND") = "DANG" Then Exit Function
On Error Resume Next
Dim p As Paragraph
Dim n As Integer
Dim a As String
For Each p In doc.Paragraphs
a = Trim(ToAscii(CleanParaText(p.Range.Text)))
If Len(a) > 0 Then
n = n + 1
If IsDocNumberLineAsc(a) Then
If Replace(a, " ", "") Like "*/####/*-*" Then TT_DetectQPPL = True
Exit Function
End If
If n >= 15 Then Exit Function
End If
Next p
End Function
Private Function TT_SectionKeyword(ByVal sAsc As String) As String
TT_SectionKeyword = ""
Dim t As String: t = Trim(sAsc)
Dim sp As Long: sp = InStr(1, t, " ")
If sp = 0 Then Exit Function
Dim w1 As String: w1 = Left(t, sp - 1)
If w1 <> "PHAN" And w1 <> "CHUONG" And w1 <> "MUC" Then Exit Function
Dim w2 As String: w2 = Trim(Mid(t, sp + 1))
Dim sp2 As Long: sp2 = InStr(1, w2, " ")
If sp2 > 0 Then w2 = Left(w2, sp2 - 1)
w2 = Replace(Replace(w2, ".", ""), ":", "")
If Len(w2) = 0 Then Exit Function
If w2 Like "#*" Or (w1 = "PHAN" And w2 = "THU") Then
TT_SectionKeyword = w1
Exit Function
End If
Dim k As Long
For k = 1 To Len(w2)
If InStr(1, "IVXLC", Mid(w2, k, 1)) = 0 Then Exit Function
Next k
TT_SectionKeyword = w1
End Function
Private Function TT_SectionKeywordText(ByVal kw As String) As String
Select Case kw
Case "PHAN": TT_SectionKeywordText = "Ph" & ChrW(7847) & "n"
Case "CHUONG": TT_SectionKeywordText = "Ch" & ChrW(432) & ChrW(417) & "ng"
Case "MUC": TT_SectionKeywordText = "M" & ChrW(7909) & "c"
End Select
End Function
Private Function TT_NextTextAsc(ByVal p As Paragraph) As String
TT_NextTextAsc = ""
On Error Resume Next
Dim pN As Paragraph
Set pN = p.Next
Do While Not pN Is Nothing
If pN.Range.Information(wdWithInTable) Then Exit Function
Dim t As String: t = CleanParaText(pN.Range.Text)
If Len(t) > 0 Then
TT_NextTextAsc = Trim(ToAscii(t))
Exit Function
End If
Set pN = pN.Next
Loop
End Function
Private Function IsAuthorityLine(ByVal p As Paragraph, ByVal s As String, ByVal sAsc As String) As Boolean
IsAuthorityLine = False
If Len(s) < 6 Or Len(s) > 150 Then Exit Function
If UCase(s) <> s Then Exit Function
If Not (sAsc Like "*[A-Z][A-Z]*") Then Exit Function
If IsSeparatorLine(s) Or IsStrictBodyOrHeadingStart(s, sAsc) Or IsPartyDocTitle(sAsc) Then Exit Function
Dim nx As String: nx = TT_NextTextAsc(p)
If nx Like "CAN CU*" Or nx Like "- CAN CU*" Or nx Like "XET DE NGHI*" Or nx Like "THEO DE NGHI*" Or _
nx = "QUYET DINH:" Or nx = "QUYET DINH :" Or nx Like "TRONG *" Then IsAuthorityLine = True
End Function
Private Sub FormatAuthorityLine(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub FormatSectionHeading(ByVal p As Paragraph, ByVal upper As Boolean)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = IIf(upper, 0, 12)
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
.KeepWithNext = True
End With
p.Range.Font.Name = "Times New Roman"
If upper Then p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub TT_AddSummaryRule(ByVal doc As Document, ByVal p As Paragraph)
On Error Resume Next
If TT_S("MOTTO_KIND") = "DANG" Then
p.Range.InsertParagraphAfter
Dim pN As Paragraph
Set pN = p.Next
doc.Range(pN.Range.Start, pN.Range.Start).InsertAfter "-----"
Call FormatCenteredSeparator(pN)
Exit Sub
End If
Dim areaW As Single
With p.Range.Sections(1).PageSetup
areaW = .PageWidth - .LeftMargin - .RightMargin
End With
If areaW < 300 Or areaW > 800 Then areaW = MillimetersToPoints(165)
Dim nL As Long
nL = 1
nL = p.Range.ComputeStatistics(wdStatisticLines)
If nL < 1 Then nL = 1
Dim textW As Single
If nL > 1 Then
textW = areaW * 0.85
Else
textW = TT_EstimateTextWidth(p)
End If
If textW > areaW Or textW <= 0 Then textW = areaW * 0.85
Dim lineW As Single
lineW = textW * 0.4
If lineW < MillimetersToPoints(20) Then lineW = MillimetersToPoints(20)
Dim lPos As Single
lPos = (areaW - lineW) / 2
Dim sz As Single
sz = p.Range.Font.Size
If sz <= 0 Or sz > 40 Then sz = 14
Dim tPos As Single
tPos = nL * sz * 1.17 + 3
If p.Format.SpaceAfter < 12 Then p.Format.SpaceAfter = 12
Dim newL As Shape
Set newL = doc.Shapes.AddLine(lPos, tPos, lPos + lineW, tPos, p.Range)
With newL
.Name = "LineTrichYeu"
.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
.Left = lPos
.Top = tPos
.Line.Weight = 0.75
.Line.ForeColor.RGB = RGB(0, 0, 0)
.Line.DashStyle = 1
.WrapFormat.Type = 3
End With
On Error GoTo 0
End Sub
Private Function IsFooterBreakerLine(ByVal sRaw As String, ByVal sAsc As String) As Boolean
IsFooterBreakerLine = False
Dim t As String: t = Trim(sAsc)
Dim r As String: r = Trim(sRaw)
If Len(t) = 0 Then Exit Function
If InStr(1, sRaw, vbTab) > 0 Then Exit Function
Dim fc As String: fc = Left(r, 1)
If fc = "-" Or fc = ChrW(8211) Or fc = "+" Or fc = ChrW(8226) Then Exit Function
If t Like "NOI NHAN*" Or t Like "LUU*" Or t Like "NHU TREN*" Or t Like "NHU DIEU*" Then Exit Function
If t Like "CAN CU*" Or t Like "XET DE NGHI*" Or t Like "THEO DE NGHI*" Or _
t Like "QUYET DINH:*" Or t Like "QUYET DINH :*" Or t Like "DIEU #*" Or t Like "DIEU ##*" Or _
t Like "#.*" Or t Like "##.*" Or t Like "[A-Za-z])*" Or Len(TT_SectionKeyword(t)) > 0 Then
IsFooterBreakerLine = True
Exit Function
End If
If Len(r) > 65 Then
IsFooterBreakerLine = True
Exit Function
End If
Dim lc As String: lc = Right(r, 1)
If (lc = "." Or lc = ";" Or lc = ":") And Not IsSignatureRoleOrAuth(t) Then IsFooterBreakerLine = True
End Function""")

# ---------- E2: set QPPL flag after resolving profile ----------
rep("""Call TT_Resolve(doc)
Call TT_DebugLog("Khoi: " & TT_Profile())""",
"""Call TT_Resolve(doc)
mTTQPPL = TT_DetectQPPL(doc)
Call TT_DebugLog("Khoi: " & TT_Profile() & " | VBQPPL: " & CStr(mTTQPPL))""")

# ---------- E4: footer table must not swallow body text ----------
rep("""If IsPartyDocTitle(sAsc) Or sAsc Like "PHU LUC*" Then Exit For
Set pFooterEnd = pCur""",
"""If IsPartyDocTitle(sAsc) Or sAsc Like "PHU LUC*" Then Exit For
If Not sawNoiNhan And IsFooterBreakerLine(sRaw, sAsc) Then
Set pFooterStart = Nothing
Set pFooterEnd = Nothing
sawSigAuthOrRole = False
GoTo NextFooterScan
End If
Set pFooterEnd = pCur""")
# footer start must not be a body line (lowercase sentence etc.)
rep("""If sAsc Like "NOI NHAN*" Or IsSignatureRoleOrAuth(sAsc) Or HasTabSeparatedFooterLine(pCur.Range.Text) Then
Set pFooterStart = pCur""",
"""If sAsc Like "NOI NHAN*" Or (IsSignatureRoleOrAuth(sAsc) And Not IsFooterBreakerLine(sRaw, sAsc)) Or HasTabSeparatedFooterLine(pCur.Range.Text) Then
Set pFooterStart = pCur""")

# ---------- shapes cleanup on re-run ----------
rep("""If StrComp(sItem.Name, "LineCoQuanBanHanh", vbTextCompare) = 0 Then shouldDel = True
""",
"""If StrComp(sItem.Name, "LineCoQuanBanHanh", vbTextCompare) = 0 Then shouldDel = True
If StrComp(sItem.Name, "LineTrichYeu", vbTextCompare) = 0 Then shouldDel = True
""")

# ---------- E3: body loop ----------
rep("""Dim afterTitleSep As Boolean: afterTitleSep = False
For Each p In doc.Paragraphs""",
"""Dim afterTitleSep As Boolean: afterTitleSep = False
Dim pLastSummary As Paragraph: Set pLastSummary = Nothing
Dim sawTitleSep As Boolean: sawTitleSep = False
Dim expectSectionTitle As Boolean: expectSectionTitle = False
For Each p In doc.Paragraphs""")

rep("""If Len(s) = 0 Then
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Format.SpaceBefore = 0""",
"""If Len(s) = 0 Then
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Format.FirstLineIndent = 0
p.Format.SpaceBefore = 0""")

rep("""If IsSeparatorLine(s) Then
titleState = 6
afterTitleSep = True
Call FormatCenteredSeparator(p)""",
"""If IsSeparatorLine(s) Then
titleState = 6
afterTitleSep = True
sawTitleSep = True
Call FormatCenteredSeparator(p)""")

rep("""ElseIf IsKinhGuiStart(s, sAscPara) Then
titleState = 6
ElseIf sAscPara Like "(KEM THEO*\"""",
"""ElseIf IsKinhGuiStart(s, sAscPara) Then
titleState = 6
ElseIf titleState >= 2 And IsAuthorityLine(p, s, sAscPara) Then
titleState = 6
Call FormatAuthorityLine(p)
GoTo NextParagraphMain
ElseIf sAscPara Like "(KEM THEO*\"""")

rep("""ElseIf HasSeparatorWithinNextParas(p, 4) And Not IsStrictBodyOrHeadingStart(s, sAscPara) Then
titleState = titleState + 1
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
ElseIf titleState = 1 And Not IsStrictBodyOrHeadingStart(s, sAscPara) And Len(s) <= 350 Then
titleState = 2
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
ElseIf titleState = 2 And Not IsBodyStart(s, sAscPara) And Len(s) <= 130 And Right(s, 1) <> "." And Right(s, 1) <> ":" Then
titleState = 3
Call FormatCenteredSummary(p)
GoTo NextParagraphMain""",
"""ElseIf HasSeparatorWithinNextParas(p, 4) And Not IsStrictBodyOrHeadingStart(s, sAscPara) Then
titleState = titleState + 1
Call FormatCenteredSummary(p)
Set pLastSummary = p
GoTo NextParagraphMain
ElseIf titleState = 1 And Not IsStrictBodyOrHeadingStart(s, sAscPara) And Len(s) <= 350 Then
titleState = 2
Call FormatCenteredSummary(p)
Set pLastSummary = p
GoTo NextParagraphMain
ElseIf titleState = 2 And Not IsBodyStart(s, sAscPara) And Len(s) <= 130 And Right(s, 1) <> "." And Right(s, 1) <> ":" And UCase(s) <> s Then
titleState = 3
Call FormatCenteredSummary(p)
Set pLastSummary = p
GoTo NextParagraphMain""")

rep("""Not IsStrictBodyOrHeadingStart(s, sAscPara) And Not IsKinhGuiStart(s, sAscPara) Then
Call FormatCenteredSummary(p)
GoTo NextParagraphMain""",
"""Not IsStrictBodyOrHeadingStart(s, sAscPara) And Not IsKinhGuiStart(s, sAscPara) Then
Call FormatAuthorityLine(p)
GoTo NextParagraphMain""")

# Phan / Chuong / Muc headings: keyword in normal case, title line below in caps (ND 30 Phu luc I; VBQPPL)
rep("""If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Or _
sAscPara Like "PHAN I*" Or sAscPara Like "PHAN II*" Or sAscPara Like "PHAN III*" Or sAscPara Like "PHAN IV*" Or sAscPara Like "PHAN THU *" Or _
sAscPara Like "CHUONG I*" Or sAscPara Like "CHUONG II*" Or sAscPara Like "CHUONG III*" Or sAscPara Like "CHUONG IV*" Then""",
"""If expectSectionTitle Then
expectSectionTitle = False
If Len(s) <= 200 And Not (sAscPara Like "DIEU #*") And Not (sAscPara Like "DIEU ##*") And Len(TT_SectionKeyword(sAscPara)) = 0 And Right(s, 1) <> "." And Right(s, 1) <> ";" Then
Call FormatSectionHeading(p, True)
GoTo NextParagraphMain
End If
End If
Dim secKw As String: secKw = ""
If Len(s) <= 120 Then secKw = TT_SectionKeyword(sAscPara)
If Len(secKw) > 0 Then
Dim kwLen As Long: kwLen = InStr(1, s, " ") - 1
If kwLen > 0 And Left(s, kwLen) <> TT_SectionKeywordText(secKw) Then
doc.Range(p.Range.Start, p.Range.Start + kwLen).Text = TT_SectionKeywordText(secKw)
End If
Call FormatSectionHeading(p, False)
Dim restAfterNum As String
restAfterNum = Trim(Mid(sAscPara, Len(secKw) + 2))
If InStr(1, restAfterNum, " ") = 0 Then expectSectionTitle = True
GoTo NextParagraphMain
End If
If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Then""")

# khoan "1." only bold when it is a heading line (no closing punctuation); never in VBQPPL
rep("""ElseIf (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And ((Len(s) < 350 And Right(s, 1) <> "." And Right(s, 1) <> ";") Or (Len(s) < 120 And Right(s, 1) <> ";")) Then
p.Range.Font.Bold = True
ElseIf sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU #:*" Or sAscPara Like "DIEU ##:*" Then
Dim colonOrDotPos As Long""",
"""ElseIf Not mTTQPPL And (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And Len(s) < 250 And Right(s, 1) <> "." And Right(s, 1) <> ";" And Right(s, 1) <> ":" And Right(s, 1) <> "," Then
p.Range.Font.Bold = True
ElseIf (sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU ###.*") And Len(s) <= 250 And Right(s, 1) <> "." And Right(s, 1) <> ";" And Right(s, 1) <> ":" Then
p.Range.Font.Bold = True
ElseIf sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU #:*" Or sAscPara Like "DIEU ##:*" Then
p.Range.Font.Bold = False
Dim colonOrDotPos As Long""")

# after loop: line under trich yeu; normalise blank spacer paragraphs
rep("""NextParagraphMain:
Next p
Call TT_DebugLog("BUOC 5 bat dau")""",
"""NextParagraphMain:
Next p
If Not (pLastSummary Is Nothing) And Not sawTitleSep Then Call TT_AddSummaryRule(doc, pLastSummary)
Call TT_DebugLog("BUOC 5 bat dau")""")

rep("""tblAll.Range.Font.ColorIndex = wdBlack
On Error Resume Next
Next tblAll""",
"""tblAll.Range.Font.ColorIndex = wdBlack
On Error Resume Next
Next tblAll
Dim pGap As Paragraph
For Each pGap In doc.Paragraphs
If Not pGap.Range.Information(wdWithInTable) Then
If Len(CleanParaText(pGap.Range.Text)) = 0 Then
pGap.Range.Font.Name = "Times New Roman"
pGap.Range.Font.Size = 12
pGap.Format.SpaceBefore = 0
pGap.Format.SpaceAfter = 0
pGap.Format.FirstLineIndent = 0
pGap.Format.LineSpacingRule = wdLineSpaceSingle
End If
End If
Next pGap""")

# ---------- E5: checks ----------
rep("""nextExpectedDieu = actualDieu + 1
End If""",
"""nextExpectedDieu = actualDieu + 1
nextExpectedArabic = 1
End If""")
rep("""If foundTitle = "QUYET DINH" Or (foundCode = "QD" And Len(foundTitle) = 0) Then
If Not hasQuyetDinhVerb Or totalDieuCount = 0 Then""",
"""If mTTQPPL And (InStr(1, docAscFull, "BAN HANH QUYET DINH") > 0 Or InStr(1, docAscFull, "BAN HANH NGHI QUYET") > 0) Then hasQuyetDinhVerb = True
If mTTQPPL And Len(foundSoRaw) > 0 And Len(foundDateRaw) > 0 Then
Dim qpYearNum As String, qpSlash As Long, qpYearDate As String
qpSlash = InStr(1, foundSoRaw, "/")
If qpSlash > 0 Then qpYearNum = Mid(foundSoRaw, qpSlash + 1, 4)
qpYearDate = Trim(Right(Trim(foundDateRaw), 4))
If IsNumeric(qpYearNum) And IsNumeric(qpYearDate) And qpYearNum <> qpYearDate Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So ky hieu VBQPPL] Nam trong so ky hieu (" & qpYearNum & ") khac nam ban hanh (" & qpYearDate & ")."
End If
End If
If foundTitle = "QUYET DINH" Or (foundCode = "QD" And Len(foundTitle) = 0) Then
If Not hasQuyetDinhVerb Or totalDieuCount = 0 Then""")

open(dst, 'w', encoding='utf-8').write(s)
print('patched OK', len(s.splitlines()), 'lines')

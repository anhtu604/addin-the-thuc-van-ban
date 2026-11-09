Private mTTProfile As String
Private mTTSource As String
Private mTTDetected As String
Private mTTDetectedStrong As Boolean
Private mTTHeaderNote As String
Private mTTForceProfile As String
Private mTTRibbon As Object
Private mTTQPPL As Boolean
Private mTTNotes As String
Private Const LIC_CAN_NHAP_MA As Boolean = False
Private Const TT_SETTING_APP As String = "OfficePartySuite"
Private Const TT_SETTING_SECTION As String = "TheThuc"
Private Const TT_DEFAULT_PROFILE As String = "DANG"
Public Function TT_ProfileCount() As Integer
TT_ProfileCount = 5
End Function
Public Function TT_ProfileIdAt(ByVal idx As Integer) As String
Select Case idx
Case 1: TT_ProfileIdAt = "DANG"
Case 2: TT_ProfileIdAt = "CHINH_QUYEN"
Case 3: TT_ProfileIdAt = "CONG_AN"
Case 4: TT_ProfileIdAt = "QUAN_DOI"
Case 5: TT_ProfileIdAt = "MTTQ"
Case Else: TT_ProfileIdAt = ""
End Select
End Function
Public Function TT_ProfileLabelAscii(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_ProfileLabelAscii = "Dang"
Case "CHINH_QUYEN": TT_ProfileLabelAscii = "Chinh quyen"
Case "CONG_AN": TT_ProfileLabelAscii = "Cong an"
Case "QUAN_DOI": TT_ProfileLabelAscii = "Quan doi"
Case "MTTQ": TT_ProfileLabelAscii = "MTTQ va to chuc CT-XH"
Case "AUTO": TT_ProfileLabelAscii = "Tu dong nhan dien"
Case Else: TT_ProfileLabelAscii = pid
End Select
End Function
Public Function TT_LegalBasisAscii(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_LegalBasisAscii = "Quy dinh so 399-QD/TW ngay 09/01/2026 cua Ban Bi thu; Huong dan so 05-HD/VPTW ngay 27/5/2026 cua Van phong Trung uong Dang"
Case "CHINH_QUYEN": TT_LegalBasisAscii = "Nghi dinh so 30/2020/ND-CP ngay 05/3/2020 cua Chinh phu ve cong tac van thu (Phu luc I: the thuc, ky thuat trinh bay van ban hanh chinh)"
Case "CONG_AN": TT_LegalBasisAscii = "Thong tu so 26/2019/TT-BCA ve cong tac van thu trong Cong an nhan dan (hop nhat tai Van ban hop nhat so 08/VBHN-BCA nam 2021); The thuc, ky thuat trinh bay theo Dieu 8, 9 Nghi dinh so 30/2020/ND-CP"
Case "QUAN_DOI": TT_LegalBasisAscii = "The thuc, ky thuat trinh bay theo Nghi dinh so 30/2020/ND-CP; Quy che cong tac van thu trong Bo Quoc phong (Thong tu so 68/2023/TT-BQP ngay 25/9/2023) - chua doi chieu toan van"
Case "MTTQ": TT_LegalBasisAscii = "Quyet dinh so 207/QD-MTTW-BTT ngay 29/8/2025 cua Ban Thuong truc Uy ban Trung uong MTTQ Viet Nam - chua doi chieu toan van; The thuc van ban cua tung to chuc chinh tri - xa hoi do Thuong truc to chuc do o Trung uong quy dinh"
Case Else: TT_LegalBasisAscii = ""
End Select
End Function
Private Function TT_RawN(ByVal pid As String, ByVal key As String) As Single
Select Case pid
Case "DANG"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13.5
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 14
Case "HDR_LEFT_MM": TT_RawN = 72
Case "HDR_RIGHT_MM": TT_RawN = 93
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 14
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 15
Case "SUMMARY_SIZE": TT_RawN = 15
Case "BODY_SIZE": TT_RawN = 15
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 14
Case "NN_ITEM_SIZE": TT_RawN = 12
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "CHINH_QUYEN"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "CONG_AN"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "QUAN_DOI"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "MTTQ"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
End Select
End Function
Private Function TT_RawB(ByVal pid As String, ByVal key As String) As Boolean
Select Case pid
Case "DANG"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = False
Case "CV_SUMMARY_ITALIC": TT_RawB = True
Case "LEGAL_BASIS_ITALIC": TT_RawB = False
Case "KINH_GUI_ITALIC": TT_RawB = True
Case "NN_LABEL_UNDERLINE": TT_RawB = True
Case "NN_LABEL_BOLD": TT_RawB = False
Case "NN_LABEL_ITALIC": TT_RawB = False
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = False
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = True
End Select
Case "CHINH_QUYEN"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "CONG_AN"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = True
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "QUAN_DOI"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = True
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "MTTQ"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = False
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = False
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
End Select
End Function
Private Function TT_RawS(ByVal pid As String, ByVal key As String) As String
Select Case pid
Case "DANG"
Select Case key
Case "MOTTO_KIND": TT_RawS = "DANG"
Case "RIGHT_ALIGN": TT_RawS = "RIGHT"
Case "DOC_NUMBER_STYLE": TT_RawS = "DANG"
Case "SIGN_PREFIX_STYLE": TT_RawS = "SLASH"
Case "STATUS": TT_RawS = "official"
End Select
Case "CHINH_QUYEN"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "STATUS": TT_RawS = "official"
End Select
Case "CONG_AN"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "STATUS": TT_RawS = "official"
End Select
Case "QUAN_DOI"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "STATUS": TT_RawS = "needs_review"
End Select
Case "MTTQ"
Select Case key
Case "MOTTO_KIND": TT_RawS = "LINH_HOAT"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "LINH_HOAT"
Case "SIGN_PREFIX_STYLE": TT_RawS = "GIU_NGUYEN"
Case "STATUS": TT_RawS = "experimental"
End Select
End Select
End Function
Private Function TT_RawDocCodeTitle(ByVal pid As String, ByVal codeAsc As String) As String
Select Case pid
Case "DANG"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "CV": TT_RawDocCodeTitle = "CONG VAN"
Case "DA": TT_RawDocCodeTitle = "DU AN|DE AN"
Case "DL": TT_RawDocCodeTitle = "DIEU LE"
Case "GCN": TT_RawDocCodeTitle = "GIAY CHUNG NHAN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "KL": TT_RawDocCodeTitle = "KET LUAN"
Case "LKG": TT_RawDocCodeTitle = "LOI KEU GOI"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH|QUY DINH"
Case "QH": TT_RawDocCodeTitle = "QUY HOACH"
Case "TB": TT_RawDocCodeTitle = "THONG BAO|TUYEN BO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TT": TT_RawDocCodeTitle = "THONG TRI"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "CHINH_QUYEN"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "CONG_AN"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "QUAN_DOI"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "MTTQ"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
End Select
End Function
Public Function TT_DetectFromHeaderText(ByVal sAsc As String, ByRef isStrong As Boolean) As String
Dim t As String: t = Replace(UCase(sAsc), " ", "")
isStrong = False
TT_DetectFromHeaderText = ""
If InStr(1, t, "DANGCONGSANVIETNAM") > 0 Then
isStrong = True
TT_DetectFromHeaderText = "DANG"
Exit Function
End If
Dim hasNational As Boolean: hasNational = False
If InStr(1, t, "CONGHOAXAHOICHUNGHIA") > 0 Or InStr(1, t, "DOCLAP-TUDO") > 0 Then hasNational = True
If Not hasNational Then
If InStr(1, t, "DANGUY") > 0 Or InStr(1, t, "DANGBO") > 0 Or InStr(1, t, "CHIBO") > 0 Or InStr(1, t, "TINHUY") > 0 Or InStr(1, t, "THANHUY") > 0 Or InStr(1, t, "BANTHUONGVU") > 0 Or InStr(1, t, "BANCHAPHANHDANGBO") > 0 Then
TT_DetectFromHeaderText = "DANG"
Exit Function
End If
End If
If InStr(1, t, "BOCONGAN") > 0 Or InStr(1, t, "CONGANTINH") > 0 Or InStr(1, t, "CONGANTHANHPHO") > 0 Or InStr(1, t, "CONGANXA") > 0 Or InStr(1, t, "CONGANPHUONG") > 0 Or InStr(1, t, "CONGANDACKHU") > 0 Or InStr(1, t, "CONGANNHANDAN") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "CONG_AN"
Exit Function
End If
If InStr(1, t, "BOQUOCPHONG") > 0 Or InStr(1, t, "QUANKHU") > 0 Or InStr(1, t, "QUANDOAN") > 0 Or InStr(1, t, "BOTULENH") > 0 Or InStr(1, t, "BOCHIHUYQUANSU") > 0 Or InStr(1, t, "BANCHIHUYQUANSU") > 0 Or InStr(1, t, "SUDOAN") > 0 Or InStr(1, t, "LUDOAN") > 0 Or InStr(1, t, "TRUNGDOAN") > 0 Or InStr(1, t, "BIENPHONG") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "QUAN_DOI"
Exit Function
End If
If InStr(1, t, "MATTRANTOQUOC") > 0 Or InStr(1, t, "UYBANMATTRAN") > 0 Or InStr(1, t, "LIENDOANLAODONG") > 0 Or InStr(1, t, "CONGDOAN") > 0 Or InStr(1, t, "HOINONGDAN") > 0 Or InStr(1, t, "HOILIENHIEPPHUNU") > 0 Or InStr(1, t, "HOICUUCHIENBINH") > 0 Or InStr(1, t, "DOANTNCS") > 0 Or InStr(1, t, "DOANTHANHNIEN") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "MTTQ"
Exit Function
End If
If InStr(1, t, "UYBANNHANDAN") > 0 Or InStr(1, t, "HOIDONGNHANDAN") > 0 Or InStr(1, t, "UBND") > 0 Or InStr(1, t, "HDND") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "CHINH_QUYEN"
Exit Function
End If
If hasNational Then
isStrong = True
TT_DetectFromHeaderText = "CHINH_QUYEN"
End If
End Function
Public Function TT_TextQuocHieu() As String
TT_TextQuocHieu = "C" & ChrW(7896) & "NG H" & ChrW(210) & "A X" & ChrW(195) & " H" & ChrW(7896) & "I CH" & ChrW(7910) & " NGH" & ChrW(296) & "A VI" & ChrW(7878) & "T NAM"
End Function
Public Function TT_TextTieuNgu() As String
TT_TextTieuNgu = ChrW(272) & ChrW(7897) & "c l" & ChrW(7853) & "p - T" & ChrW(7921) & " do - H" & ChrW(7841) & "nh ph" & ChrW(250) & "c"
End Function
Public Function TT_MauCount(ByVal grp As String) As Integer
Select Case grp
Case "DANG": TT_MauCount = 9
Case "NHA_NUOC": TT_MauCount = 8
Case Else: TT_MauCount = 0
End Select
End Function
Public Function TT_MauLabelAscii(ByVal grp As String, ByVal idx As Integer) As String
TT_MauLabelAscii = ""
Select Case grp
Case "DANG"
Select Case idx
Case 1: TT_MauLabelAscii = "Nghi quyet"
Case 2: TT_MauLabelAscii = "Quyet dinh"
Case 3: TT_MauLabelAscii = "Ke hoach"
Case 4: TT_MauLabelAscii = "Bao cao"
Case 5: TT_MauLabelAscii = "Thong bao"
Case 6: TT_MauLabelAscii = "To trinh"
Case 7: TT_MauLabelAscii = "Cong van"
Case 8: TT_MauLabelAscii = "Giay moi"
Case 9: TT_MauLabelAscii = "Bien ban"
End Select
Case "NHA_NUOC"
Select Case idx
Case 1: TT_MauLabelAscii = "Quyet dinh"
Case 2: TT_MauLabelAscii = "Ke hoach"
Case 3: TT_MauLabelAscii = "Bao cao"
Case 4: TT_MauLabelAscii = "Thong bao"
Case 5: TT_MauLabelAscii = "To trinh"
Case 6: TT_MauLabelAscii = "Cong van"
Case 7: TT_MauLabelAscii = "Giay moi"
Case 8: TT_MauLabelAscii = "Bien ban"
End Select
End Select
End Function
Public Function TT_MauData(ByVal grp As String, ByVal idx As Integer) As String
TT_MauData = ""
Select Case grp
Case "DANG"
Select Case idx
Case 1: TT_MauData = TT_Mau_DANG_1()
Case 2: TT_MauData = TT_Mau_DANG_2()
Case 3: TT_MauData = TT_Mau_DANG_3()
Case 4: TT_MauData = TT_Mau_DANG_4()
Case 5: TT_MauData = TT_Mau_DANG_5()
Case 6: TT_MauData = TT_Mau_DANG_6()
Case 7: TT_MauData = TT_Mau_DANG_7()
Case 8: TT_MauData = TT_Mau_DANG_8()
Case 9: TT_MauData = TT_Mau_DANG_9()
End Select
Case "NHA_NUOC"
Select Case idx
Case 1: TT_MauData = TT_Mau_NHA_NUOC_1()
Case 2: TT_MauData = TT_Mau_NHA_NUOC_2()
Case 3: TT_MauData = TT_Mau_NHA_NUOC_3()
Case 4: TT_MauData = TT_Mau_NHA_NUOC_4()
Case 5: TT_MauData = TT_Mau_NHA_NUOC_5()
Case 6: TT_MauData = TT_Mau_NHA_NUOC_6()
Case 7: TT_MauData = TT_Mau_NHA_NUOC_7()
Case 8: TT_MauData = TT_Mau_NHA_NUOC_8()
End Select
End Select
End Function
Private Function TT_Mau_DANG_1() As String
Dim s As String
s = ""
s = s & "CODE|" & "NQ" & vbLf
s = s & "L|" & "NGH" & ChrW(7882) & " QUY" & ChrW(7870) & "T" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[Ph" & ChrW(7847) & "n m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u: c" & ChrW(259) & "n c" & ChrW(7913) & ", b" & ChrW(7889) & "i c" & ChrW(7843) & "nh ban h" & ChrW(224) & "nh ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & "]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [M" & ChrW(7908) & "C TI" & ChrW(202) & "U, NHI" & ChrW(7878) & "M V" & ChrW(7908) & ", GI" & ChrW(7842) & "I PH" & ChrW(193) & "P]" & vbLf
s = s & "L|" & "1. [M" & ChrW(7909) & "c ti" & ChrW(234) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", gi" & ChrW(7843) & "i ph" & ChrW(225) & "p]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_1 = s
End Function
Private Function TT_Mau_DANG_2() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " vi" & ChrW(7879) & "c ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [" & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " " & ChrW(272) & ChrW(7843) & "ng; quy ch" & ChrW(7871) & " l" & ChrW(224) & "m vi" & ChrW(7879) & "c; v" & ChrW(259) & "n b" & ChrW(7843) & "n li" & ChrW(234) & "n quan];" & vbLf
s = s & "L|" & "- X" & ChrW(233) & "t " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_2 = s
End Function
Private Function TT_Mau_DANG_3() As String
Dim s As String
s = ""
s = s & "CODE|" & "KH" & vbLf
s = s & "L|" & "K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung k" & ChrW(7871) & " ho" & ChrW(7841) & "ch]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(259) & "n c" & ChrW(7913) & ", m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u k" & ChrW(7871) & " ho" & ChrW(7841) & "ch.]" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "1. M" & ChrW(7909) & "c " & ChrW(273) & ChrW(237) & "ch" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. Y" & ChrW(234) & "u c" & ChrW(7847) & "u" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_3 = s
End Function
Private Function TT_Mau_DANG_4() As String
Dim s As String
s = ""
s = s & "CODE|" & "BC" & vbLf
s = s & "L|" & "B" & ChrW(193) & "O C" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung b" & ChrW(225) & "o c" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[M" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u b" & ChrW(225) & "o c" & ChrW(225) & "o.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & " TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [H" & ChrW(7840) & "N CH" & ChrW(7870) & ", NGUY" & ChrW(202) & "N NH" & ChrW(194) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. [PH" & ChrW(431) & ChrW(416) & "NG H" & ChrW(431) & ChrW(7898) & "NG, NHI" & ChrW(7878) & "M V" & ChrW(7908) & " TH" & ChrW(7900) & "I GIAN T" & ChrW(7898) & "I]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "IV. [KI" & ChrW(7870) & "N NGH" & ChrW(7882) & ", " & ChrW(272) & ChrW(7872) & " XU" & ChrW(7844) & "T]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_4 = s
End Function
Private Function TT_Mau_DANG_5() As String
Dim s As String
s = ""
s = s & "CODE|" & "TB" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG B" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_5 = s
End Function
Private Function TT_Mau_DANG_6() As String
Dim s As String
s = ""
s = s & "CODE|" & "TTr" & vbLf
s = s & "L|" & "T" & ChrW(7900) & " TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n t" & ChrW(7901) & " tr" & ChrW(236) & "nh]." & vbLf
s = s & "L|" & "[L" & ChrW(253) & " do, c" & ChrW(259) & "n c" & ChrW(7913) & " tr" & ChrW(236) & "nh.]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & ".]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c] xem x" & ChrW(233) & "t, quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_6 = s
End Function
Private Function TT_Mau_DANG_7() As String
Dim s As String
s = ""
s = s & "CODE|" & "CV" & vbLf
s = s & "CV|1" & vbLf
s = s & "SUM|" & "V/v [tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n c" & ChrW(244) & "ng v" & ChrW(259) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n.]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_7 = s
End Function
Private Function TT_Mau_DANG_8() As String
Dim s As String
s = ""
s = s & "CODE|" & "GM" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y M" & ChrW(7900) & "I" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: d" & ChrW(7921) & " h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng k" & ChrW(237) & "nh m" & ChrW(7901) & "i: [" & ChrW(272) & ChrW(7841) & "i bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7901) & "i]." & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [n" & ChrW(7897) & "i dung]." & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " chu" & ChrW(7849) & "n b" & ChrW(7883) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u, li" & ChrW(234) & "n h" & ChrW(7879) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ").]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_8 = s
End Function
Private Function TT_Mau_DANG_9() As String
Dim s As String
s = ""
s = s & "CODE|" & "BB" & vbLf
s = s & "L|" & "BI" & ChrW(202) & "N B" & ChrW(7842) & "N" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "Th" & ChrW(224) & "nh ph" & ChrW(7847) & "n tham d" & ChrW(7921) & ": [th" & ChrW(224) & "nh ph" & ChrW(7847) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & ": [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(432) & " k" & ChrW(253) & " (ng" & ChrW(432) & ChrW(7901) & "i ghi bi" & ChrW(234) & "n b" & ChrW(7843) & "n): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [theo di" & ChrW(7877) & "n bi" & ChrW(7871) & "n cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "Cu" & ChrW(7897) & "c h" & ChrW(7885) & "p k" & ChrW(7871) & "t th" & ChrW(250) & "c v" & ChrW(224) & "o [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " TR" & ChrW(204) & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
s = s & "LR|" & "TH" & ChrW(431) & " K" & ChrW(221) & vbLf
s = s & "LH|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_9 = s
End Function
Private Function TT_Mau_NHA_NUOC_1() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[TH" & ChrW(7848) & "M QUY" & ChrW(7872) & "N BAN H" & ChrW(192) & "NH, V" & ChrW(205) & " D" & ChrW(7908) & ": CH" & ChrW(7910) & " T" & ChrW(7882) & "CH " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh th" & ChrW(7849) & "m quy" & ChrW(7873) & "n, ch" & ChrW(7913) & "c n" & ChrW(259) & "ng, nhi" & ChrW(7879) & "m v" & ChrW(7909) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(7897) & "i dung, c" & ChrW(417) & " s" & ChrW(7903) & " ban h" & ChrW(224) & "nh];" & vbLf
s = s & "L|" & "Theo " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c thi h" & ChrW(224) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_1 = s
End Function
Private Function TT_Mau_NHA_NUOC_2() As String
Dim s As String
s = ""
s = s & "CODE|" & "KH" & vbLf
s = s & "L|" & "K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung k" & ChrW(7871) & " ho" & ChrW(7841) & "ch]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(259) & "n c" & ChrW(7913) & ", m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u k" & ChrW(7871) & " ho" & ChrW(7841) & "ch.]" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "1. M" & ChrW(7909) & "c " & ChrW(273) & ChrW(237) & "ch" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. Y" & ChrW(234) & "u c" & ChrW(7847) & "u" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_2 = s
End Function
Private Function TT_Mau_NHA_NUOC_3() As String
Dim s As String
s = ""
s = s & "CODE|" & "BC" & vbLf
s = s & "L|" & "B" & ChrW(193) & "O C" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung b" & ChrW(225) & "o c" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[M" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u b" & ChrW(225) & "o c" & ChrW(225) & "o.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & " TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [H" & ChrW(7840) & "N CH" & ChrW(7870) & ", NGUY" & ChrW(202) & "N NH" & ChrW(194) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. [PH" & ChrW(431) & ChrW(416) & "NG H" & ChrW(431) & ChrW(7898) & "NG, NHI" & ChrW(7878) & "M V" & ChrW(7908) & " TH" & ChrW(7900) & "I GIAN T" & ChrW(7898) & "I]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "IV. [KI" & ChrW(7870) & "N NGH" & ChrW(7882) & ", " & ChrW(272) & ChrW(7872) & " XU" & ChrW(7844) & "T]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_3 = s
End Function
Private Function TT_Mau_NHA_NUOC_4() As String
Dim s As String
s = ""
s = s & "CODE|" & "TB" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG B" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_4 = s
End Function
Private Function TT_Mau_NHA_NUOC_5() As String
Dim s As String
s = ""
s = s & "CODE|" & "TTr" & vbLf
s = s & "L|" & "T" & ChrW(7900) & " TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung t" & ChrW(7901) & " tr" & ChrW(236) & "nh]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n t" & ChrW(7901) & " tr" & ChrW(236) & "nh]." & vbLf
s = s & "L|" & "[L" & ChrW(253) & " do, c" & ChrW(259) & "n c" & ChrW(7913) & " tr" & ChrW(236) & "nh.]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & ".]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c] xem x" & ChrW(233) & "t, quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_5 = s
End Function
Private Function TT_Mau_NHA_NUOC_6() As String
Dim s As String
s = ""
s = s & "CODE|" & vbLf
s = s & "CV|1" & vbLf
s = s & "SUM|" & "V/v [tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n c" & ChrW(244) & "ng v" & ChrW(259) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n.]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_6 = s
End Function
Private Function TT_Mau_NHA_NUOC_7() As String
Dim s As String
s = ""
s = s & "CODE|" & "GM" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y M" & ChrW(7900) & "I" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: d" & ChrW(7921) & " h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng k" & ChrW(237) & "nh m" & ChrW(7901) & "i: [" & ChrW(272) & ChrW(7841) & "i bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7901) & "i]." & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [n" & ChrW(7897) & "i dung]." & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " chu" & ChrW(7849) & "n b" & ChrW(7883) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u, li" & ChrW(234) & "n h" & ChrW(7879) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ").]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_7 = s
End Function
Private Function TT_Mau_NHA_NUOC_8() As String
Dim s As String
s = ""
s = s & "CODE|" & "BB" & vbLf
s = s & "L|" & "BI" & ChrW(202) & "N B" & ChrW(7842) & "N" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "Th" & ChrW(224) & "nh ph" & ChrW(7847) & "n tham d" & ChrW(7921) & ": [th" & ChrW(224) & "nh ph" & ChrW(7847) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & ": [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(432) & " k" & ChrW(253) & " (ng" & ChrW(432) & ChrW(7901) & "i ghi bi" & ChrW(234) & "n b" & ChrW(7843) & "n): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [theo di" & ChrW(7877) & "n bi" & ChrW(7871) & "n cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "Cu" & ChrW(7897) & "c h" & ChrW(7885) & "p k" & ChrW(7871) & "t th" & ChrW(250) & "c v" & ChrW(224) & "o [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " T" & ChrW(7884) & "A" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
s = s & "LR|" & "TH" & ChrW(431) & " K" & ChrW(221) & vbLf
s = s & "LH|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_8 = s
End Function
Public Function TT_TextTieuNguDang() As String
TT_TextTieuNguDang = ChrW(272) & ChrW(7842) & "NG C" & ChrW(7896) & "NG S" & ChrW(7842) & "N VI" & ChrW(7878) & "T NAM"
End Function
Public Function TT_GetSetting() As String
Dim v As String
On Error Resume Next
v = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "DonVi", "AUTO")
On Error GoTo 0
If Len(v) = 0 Then v = "AUTO"
If v <> "AUTO" And TT_IndexOfProfile(v) = 0 Then v = "AUTO"
TT_GetSetting = v
End Function
Public Sub TT_SaveSetting(ByVal pid As String)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonVi", pid
On Error GoTo 0
End Sub
Private Function TT_IndexOfProfile(ByVal pid As String) As Integer
Dim i As Integer
TT_IndexOfProfile = 0
For i = 1 To TT_ProfileCount()
If TT_ProfileIdAt(i) = pid Then
TT_IndexOfProfile = i
Exit Function
End If
Next i
End Function
Public Sub RibbonDonVi_GetSelectedIndex(control As Object, ByRef returnedVal)
Dim v As String
v = TT_GetSetting()
If v = "AUTO" Then
returnedVal = 0
Else
returnedVal = TT_IndexOfProfile(v)
End If
End Sub
Public Sub RibbonDonVi_OnAction(control As Object, selectedId As String, selectedIndex As Integer)
If selectedIndex <= 0 Then
Call TT_SaveSetting("AUTO")
Else
Call TT_SaveSetting(TT_ProfileIdAt(selectedIndex))
End If
Call TT_RibbonRefresh
End Sub
Public Sub RibbonOnLoad(ribbon As Object)
Set mTTRibbon = ribbon
Call CN_HenKiemTraKhiMo
End Sub
Private Sub TT_RibbonRefresh()
On Error Resume Next
If Not mTTRibbon Is Nothing Then mTTRibbon.Invalidate
On Error GoTo 0
End Sub
Public Sub RibbonChonKhoi(control As Object)
Call ChonKhoiDonVi
End Sub
Public Sub ChonKhoiDonVi()
Dim msg As String
Dim i As Integer
msg = "Chon khoi the thuc van ban (go so roi bam OK):" & vbCrLf & vbCrLf & "0. Tu dong nhan dien theo dong tieu ngu / co quan ban hanh"
For i = 1 To TT_ProfileCount()
msg = msg & vbCrLf & i & ". " & TT_ProfileLabelAscii(TT_ProfileIdAt(i))
Next i
msg = msg & vbCrLf & vbCrLf & "Van ban cua cap uy, to chuc dang trong Cong an, Quan doi: chon khoi Dang."
Dim cur As String
Dim curIdx As Integer
cur = TT_GetSetting()
If cur = "AUTO" Then
curIdx = 0
Else
curIdx = TT_IndexOfProfile(cur)
End If
Dim ans As String
ans = InputBox(msg, "Khoi the thuc van ban", CStr(curIdx))
If Len(Trim(ans)) = 0 Then Exit Sub
If Not IsNumeric(ans) Then Exit Sub
Dim n As Integer
n = Val(ans)
If n = 0 Then
Call TT_SaveSetting("AUTO")
ElseIf n >= 1 And n <= TT_ProfileCount() Then
Call TT_SaveSetting(TT_ProfileIdAt(n))
Else
Exit Sub
End If
MsgBox "Da chon khoi: " & TT_ProfileLabelAscii(TT_GetSetting()), vbInformation, "Khoi the thuc van ban"
End Sub
Private Function TT_HeaderAreaAscii(ByVal doc As Document) As String
Dim buf As String
Dim nonEmpty As Integer
Dim p As Paragraph
On Error Resume Next
For Each p In doc.Paragraphs
Dim sRaw As String
sRaw = CleanParaText(p.Range.Text)
If Len(sRaw) > 0 Then
Dim sA As String
sA = Trim(ToAscii(sRaw))
If nonEmpty >= 3 And Not p.Range.Information(wdWithInTable) Then
If IsPartyDocTitle(sA) Or IsKinhGuiStart(sRaw, sA) Or IsStrictBodyOrHeadingStart(sRaw, sA) Then Exit For
End If
buf = buf & " " & sA
nonEmpty = nonEmpty + 1
If nonEmpty >= 14 Then Exit For
End If
Next p
On Error GoTo 0
TT_HeaderAreaAscii = buf
End Function
Public Function TT_Resolve(ByVal doc As Document) As String
Dim setting As String
Dim strong As Boolean
setting = TT_GetSetting()
If Len(mTTForceProfile) > 0 Then setting = mTTForceProfile
mTTHeaderNote = ""
mTTDetected = TT_DetectFromHeaderText(TT_HeaderAreaAscii(doc), strong)
mTTDetectedStrong = strong
If setting = "AUTO" Then
If Len(mTTDetected) > 0 Then
mTTProfile = mTTDetected
mTTSource = "TU_NHAN_DIEN"
Else
mTTProfile = TT_DEFAULT_PROFILE
mTTSource = "MAC_DINH"
End If
ElseIf mTTDetectedStrong And Len(mTTDetected) > 0 And Len(mTTForceProfile) = 0 And _
((TT_RawS(mTTDetected, "MOTTO_KIND") = "DANG") <> (TT_RawS(setting, "MOTTO_KIND") = "DANG")) Then
mTTProfile = mTTDetected
mTTSource = "GHI_DE"
Else
mTTProfile = setting
mTTSource = "CHON"
End If
TT_Resolve = mTTProfile
End Function
Public Function TT_Profile() As String
If Len(mTTProfile) = 0 Then mTTProfile = TT_DEFAULT_PROFILE
TT_Profile = mTTProfile
End Function
Public Function TT_N(ByVal key As String) As Single
TT_N = TT_RawN(TT_Profile(), key)
End Function
Public Function TT_B(ByVal key As String) As Boolean
TT_B = TT_RawB(TT_Profile(), key)
End Function
Public Function TT_S(ByVal key As String) As String
TT_S = TT_RawS(TT_Profile(), key)
End Function
Public Function TT_IsNational() As Boolean
TT_IsNational = (TT_S("MOTTO_KIND") <> "DANG")
End Function
Public Function TT_RightColumnAlignment() As Long
If TT_S("RIGHT_ALIGN") = "CENTER" Then
TT_RightColumnAlignment = wdAlignParagraphCenter
Else
TT_RightColumnAlignment = wdAlignParagraphRight
End If
End Function
Public Function TT_DocCodeTitle(ByVal codeAsc As String) As String
TT_DocCodeTitle = TT_RawDocCodeTitle(TT_Profile(), UCase(Trim(codeAsc)))
End Function
Public Function TT_ReportHeader() As String
Dim s As String
s = "Khoi the thuc: " & TT_ProfileLabelAscii(TT_Profile())
Select Case mTTSource
Case "TU_NHAN_DIEN": s = s & " (tu dong nhan dien)"
Case "MAC_DINH": s = s & " (khong nhan dien duoc, dung khoi mac dinh)"
Case "GHI_DE": s = s & " (tu dong chuyen vi tieu ngu tren van ban khac khoi dang chon tren Ribbon)"
Case Else: s = s & " (nguoi dung chon)"
End Select
s = s & vbCrLf & "Can cu: " & TT_LegalBasisAscii(TT_Profile())
Select Case TT_S("STATUS")
Case "experimental": s = s & vbCrLf & "Luu y: bo the thuc khoi nay dang THU NGHIEM, chi kiem tra noi long."
Case "needs_review": s = s & vbCrLf & "Luu y: bo the thuc khoi nay CAN DOI CHIEU THEM van ban goc."
End Select
If mTTQPPL Then s = s & vbCrLf & "Loai: VAN BAN QUY PHAM PHAP LUAT - trinh bay Chuong/Muc/Dieu theo Luat Ban hanh VBQPPL nam 2025 va Nghi dinh huong dan (can doi chieu van ban goc)."
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
Dim pR As Paragraph
Set pR = p.Next
Dim reuse As Boolean: reuse = False
If Not pR Is Nothing Then
If Not pR.Range.Information(wdWithInTable) And Len(CleanParaText(pR.Range.Text)) = 0 Then reuse = True
End If
If Not reuse Then
p.Range.InsertParagraphAfter
Set pR = p.Next
End If
p.Format.SpaceAfter = 8
p.Format.KeepWithNext = True
With pR.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = lPos
.RightIndent = lPos
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 10
.LineSpacingRule = wdLineSpaceExactly
.LineSpacing = 2
.KeepWithNext = False
End With
pR.Range.Font.Size = 2
pR.Range.Font.Bold = False
With pR.Borders(wdBorderTop)
.LineStyle = wdLineStyleSingle
.LineWidth = wdLineWidth075pt
.Color = wdColorBlack
End With
On Error GoTo 0
End Sub
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
End Function
Public Function TT_MismatchWarning() As String
TT_MismatchWarning = ""
If mTTSource <> "CHON" Or Not mTTDetectedStrong Or Len(mTTDetected) = 0 Then Exit Function
If mTTDetected = mTTProfile Then Exit Function
If (TT_RawS(mTTDetected, "MOTTO_KIND") = "DANG") <> (TT_RawS(mTTProfile, "MOTTO_KIND") = "DANG") Then
TT_MismatchWarning = "[Khoi the thuc] Dang chon khoi '" & TT_ProfileLabelAscii(mTTProfile) & _
"' nhung dong tieu ngu tren van ban thuoc khoi '" & TT_ProfileLabelAscii(mTTDetected) & _
"'. Kiem tra lai muc Khoi tren Ribbon (hoac chon Tu dong nhan dien)."
End If
End Function
Private Function TT_EstimateTextWidth(ByVal p As Paragraph) As Single
Dim t As String
t = CleanParaText(p.Range.Text)
Dim sz As Single
sz = p.Range.Font.Size
If sz <= 0 Or sz > 40 Then sz = 13
If InStr(1, Replace(ToAscii(t), " ", ""), "DOCLAP-TUDO-HANHPHUC") > 0 Then
TT_EstimateTextWidth = CentimetersToPoints(5.95) * sz / 14
ElseIf UCase(t) = t Then
TT_EstimateTextWidth = Len(t) * sz * 0.62
Else
TT_EstimateTextWidth = Len(t) * sz * 0.48
End If
End Function
Private Function TT_ParaLines(ByVal p As Paragraph) As Long
TT_ParaLines = 1
If p Is Nothing Then Exit Function
On Error Resume Next
Dim n As Long, r As Range
n = 0
n = p.Range.ComputeStatistics(wdStatisticLines)
If n < 1 Then
Set r = p.Range.Duplicate
r.MoveEnd wdCharacter, -1
n = r.ComputeStatistics(wdStatisticLines)
End If
If n < 1 Then
Dim y1 As Single, y2 As Single
Set r = p.Range.Duplicate
r.MoveEnd wdCharacter, -1
y1 = r.Characters(1).Information(wdVerticalPositionRelativeToPage)
y2 = r.Characters(r.Characters.Count).Information(wdVerticalPositionRelativeToPage)
If y2 - y1 > p.Range.Font.Size * 0.6 Then n = 2 Else n = 1
End If
If n < 1 Then n = 1
TT_ParaLines = n
End Function
Private Function TT_OrganWraps(ByVal organParas As Collection) As Boolean
TT_OrganWraps = False
Dim o As Variant
On Error Resume Next
For Each o In organParas
If TT_ParaLines(o) > 1 Then
TT_OrganWraps = True
Exit Function
End If
Next o
End Function
Private Sub TT_WidenOrganCol(ByVal tbl As Table, ByVal organParas As Collection, ByVal pKey As Paragraph, ByVal pDate As Paragraph, ByVal totalW As Single, ByRef rightW As Single)
Dim k As Integer
On Error Resume Next
For k = 1 To 25
If Not TT_OrganWraps(organParas) Then Exit Sub
rightW = rightW - MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
If TT_ParaLines(pKey) > 1 Or TT_ParaLines(pDate) > 1 Then
rightW = rightW + MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Exit Sub
End If
Next k
End Sub
Private Function TT_TenDaiMode() As String
On Error Resume Next
TT_TenDaiMode = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "HA_CO")
End Function
Public Sub RibbonTenDai_GetSelectedIndex(control As Object, ByRef returnedVal)
If TT_TenDaiMode() = "LAN_LE" Then returnedVal = 1 Else returnedVal = 0
End Sub
Public Sub RibbonTenDai_OnAction(control As Object, selectedId As String, selectedIndex As Integer)
On Error Resume Next
If selectedIndex = 1 Then
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "LAN_LE"
Else
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "HA_CO"
End If
End Sub
Private Sub TT_LanLeTrai(ByVal tbl As Table, ByVal organParas As Collection, ByRef totalW As Single, ByVal rightW As Single, ByRef lanW As Single)
Dim k As Integer, baseIndent As Single
On Error Resume Next
baseIndent = tbl.Rows.LeftIndent + lanW
For k = 1 To 8
If Not TT_OrganWraps(organParas) Then Exit Sub
If lanW + MillimetersToPoints(2) > MillimetersToPoints(15) + 0.5 Then Exit Sub
lanW = lanW + MillimetersToPoints(2)
totalW = totalW + MillimetersToPoints(2)
tbl.Rows.LeftIndent = baseIndent - lanW
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Call TT_DebugLog("  Lan le: " & Format(PointsToMillimeters(lanW), "0") & "mm, indent " & Format(PointsToMillimeters(tbl.Rows.LeftIndent), "0.0") & "mm, cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0") & ", o " & Format(PointsToMillimeters(tbl.Cell(1, 1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Cell(1, 2).Width), "0"))
Next k
End Sub
Private Sub TT_NarrowForRight(ByVal tbl As Table, ByRef organParas As Collection, ByVal pRight As Paragraph, ByVal totalW As Single, ByRef rightW As Single)
Dim k As Integer
On Error Resume Next
For k = 1 To 15
If TT_ParaLines(pRight) <= 1 Then Exit Sub
If totalW - rightW - MillimetersToPoints(2) < MillimetersToPoints(52) Then Exit Sub
rightW = rightW + MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
If TT_OrganWraps(organParas) Then Set organParas = TT_SplitLongOrgans(organParas, False)
If TT_OrganWraps(organParas) Then
rightW = rightW - MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Exit Sub
End If
Next k
End Sub
Private Function TT_SplitLongOrgans(ByVal organParas As Collection, ByVal standardOnly As Boolean) As Collection
Dim outC As New Collection
Dim o As Variant, p As Paragraph
On Error Resume Next
For Each o In organParas
Set p = o
outC.Add p
If TT_ParaLines(p) > 1 Then
Dim t As String, a As String, pos As Long, raw As String, off As Long
raw = p.Range.Text
t = CleanParaText(raw)
a = UCase(ToAscii(t))
pos = 0
If Len(a) = Len(t) Then
If a Like "UY BAN NHAN DAN *" Then
pos = 16
ElseIf a Like "HOI DONG NHAN DAN *" Then
pos = 18
ElseIf a Like "BO CHI HUY QUAN SU *" Then
pos = 19
ElseIf Not standardOnly Then
Dim kw As Variant
For Each kw In Array(" THANH PHO ", " TINH ", " DAC KHU ", " PHUONG ", " XA ")
pos = InStr(2, a, kw)
If pos > 0 Then Exit For
Next kw
End If
End If
If pos > 1 Then
off = InStr(1, raw, t) - 1
If off < 0 Then off = 0
Dim r As Range
Set r = p.Range.Duplicate
r.SetRange p.Range.Start + off + pos - 1, p.Range.Start + off + pos
If r.Text = " " Then
Dim pStart As Long, p1 As Paragraph
pStart = p.Range.Start
r.Text = vbCr
Set p1 = p.Range.Document.Range(pStart, pStart).Paragraphs(1)
outC.Remove outC.Count
outC.Add p1
outC.Add p1.Next
End If
End If
End If
Next o
Set TT_SplitLongOrgans = outC
End Function
Private Sub TT_AddCenteredLineUnder(ByVal doc As Document, ByVal p As Paragraph, ByVal shapeName As String, ByVal ratio As Single, ByVal fallbackWidthPt As Single)
On Error Resume Next
Dim c As Cell
Set c = Nothing
Set c = p.Range.Cells(1)
If c Is Nothing Then Exit Sub
Dim lPad As Single, rPad As Single
lPad = 5.4: rPad = 5.4
lPad = c.LeftPadding
rPad = c.RightPadding
Dim textAreaW As Single
textAreaW = c.Width - lPad - rPad
If textAreaW <= 0 Or textAreaW > 600 Then textAreaW = 240
Dim textW As Single
textW = TT_EstimateTextWidth(p)
If textW <= 0 Then textW = fallbackWidthPt
If textW > textAreaW Then textW = textAreaW
Dim lineW As Single
lineW = textW * ratio
If lineW < 28 Then lineW = 28
Dim nLines As Long
nLines = 1
nLines = p.Range.ComputeStatistics(wdStatisticLines)
If nLines < 1 Then nLines = 1
Dim lPos As Single, tPos As Single
lPos = (textAreaW - lineW) / 2
If lPos < 0 Then lPos = 0
tPos = nLines * p.Range.Font.Size * 1.17 + 4
Dim newL As Shape
Set newL = doc.Shapes.AddLine(lPos, tPos, lPos + lineW, tPos, p.Range)
With newL
.Name = shapeName
.LayoutInCell = True
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
Private Function TT_FindSignerName(ByVal doc As Document) As String
TT_FindSignerName = ""
On Error Resume Next
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.Text)
If IsFooterSignatureTable(tbl, tAsc) Then
Dim c As Cell
For Each c In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(c.Range.Text))
If InStr(1, cAsc, "NOI NHAN") = 0 And HasSignatureKeywordInCell(cAsc) Then
Dim pS As Paragraph
Dim lastText As String
lastText = ""
For Each pS In c.Range.Paragraphs
If Len(CleanParaText(pS.Range.Text)) > 0 Then lastText = CleanParaText(pS.Range.Text)
Next pS
If Len(lastText) > 0 Then
If Not IsSignatureRoleOrAuth(Trim(ToAscii(lastText))) Then
TT_FindSignerName = lastText
End If
End If
Exit Function
End If
Next c
End If
Next tbl
On Error GoTo 0
End Function
Private Function TT_HasRankPrefix(ByVal nameAsc As String) As Boolean
Dim t As String
t = Trim(UCase(nameAsc))
TT_HasRankPrefix = (t Like "DAI TUONG *" Or t Like "THUONG TUONG *" Or t Like "TRUNG TUONG *" Or t Like "THIEU TUONG *" Or _
t Like "DAI TA *" Or t Like "THUONG TA *" Or t Like "TRUNG TA *" Or t Like "THIEU TA *" Or _
t Like "DAI UY *" Or t Like "THUONG UY *" Or t Like "TRUNG UY *" Or t Like "THIEU UY *" Or _
t Like "DO DOC *" Or t Like "PHO DO DOC *" Or t Like "CHUAN DO DOC *")
End Function
Private Sub TT_DebugLog(ByVal msg As String)
On Error Resume Next
Dim baseDir As String
baseDir = Environ("HOME")
If Len(baseDir) = 0 Then Exit Sub
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim f As Integer
f = FreeFile
Open baseDir & "/thethuc_debug.log" For Append As #f
Print #f, Format(Now, "hh:nn:ss") & " " & msg
Close #f
On Error GoTo 0
End Sub
Private Function TT_EncU(ByVal s As String) As String
Dim i As Long, w As Long, r As String
r = ""
For i = 1 To Len(s)
w = AscW(Mid$(s, i, 1))
If w < 0 Then w = w + 65536
r = r & Right$("0000" & Hex$(w), 4)
Next i
TT_EncU = r
End Function
Private Function TT_DecU(ByVal h As String) As String
Dim i As Long, r As String
r = ""
On Error GoTo DecFail
For i = 1 To Len(h) - 3 Step 4
r = r & ChrW(CLng("&H" & Mid$(h, i, 4)))
Next i
TT_DecU = r
Exit Function
DecFail:
TT_DecU = r
End Function
Public Function TT_UnitGet(ByVal grp As String, ByVal field As String) As String
Dim v As String
On Error Resume Next
v = GetSetting(TT_SETTING_APP, "DonVi_" & grp, field, "")
On Error GoTo 0
TT_UnitGet = TT_DecU(v)
End Function
Public Sub TT_UnitSet(ByVal grp As String, ByVal field As String, ByVal value As String)
On Error Resume Next
SaveSetting TT_SETTING_APP, "DonVi_" & grp, field, TT_EncU(Trim(value))
On Error GoTo 0
End Sub
Public Function TT_UnitGroupCurrent() As String
Dim pid As String
pid = TT_GetSetting()
If pid <> "AUTO" Then
If Len(TT_TemplateGroupOf(pid)) > 0 Then
TT_UnitGroupCurrent = TT_TemplateGroupOf(pid)
Exit Function
End If
End If
Dim g As String
On Error Resume Next
g = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", "DANG")
On Error GoTo 0
If g <> "DANG" And g <> "NHA_NUOC" Then g = "DANG"
TT_UnitGroupCurrent = g
End Function
Public Sub RibbonUnit_GetText(control As Object, ByRef returnedVal)
returnedVal = TT_UnitGet(TT_UnitGroupCurrent(), control.Tag)
End Sub
Public Sub RibbonUnit_OnChange(control As Object, text As String)
Call TT_UnitSet(TT_UnitGroupCurrent(), control.Tag, text)
End Sub
Public Sub RibbonUnit_GetLabel(control As Object, ByRef returnedVal)
If TT_UnitGroupCurrent() = "DANG" Then
returnedVal = "(" & ChrW(273) & "ang s" & ChrW(7917) & "a: kh" & ChrW(7889) & "i " & ChrW(272) & ChrW(7843) & "ng)"
Else
returnedVal = "(" & ChrW(273) & "ang s" & ChrW(7917) & "a: kh" & ChrW(7889) & "i Nh" & ChrW(224) & " n" & ChrW(432) & ChrW(7899) & "c)"
End If
End Sub
Private Sub TT_ReadSignerInfo(ByVal src As Document, ByRef auth As String, ByRef role As String, ByRef signer As String)
auth = "": role = "": signer = ""
On Error GoTo SigFail
Dim tbl As Table
For Each tbl In src.Tables
If IsFooterSignatureTable(tbl, ToAscii(tbl.Range.Text)) Then
Dim c As Cell
For Each c In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(c.Range.Text))
If InStr(1, cAsc, "NOI NHAN") = 0 And HasSignatureKeywordInCell(cAsc) Then
Dim p As Paragraph
For Each p In c.Range.Paragraphs
Dim t As String
t = CleanParaText(p.Range.Text)
If Len(t) > 0 Then
Dim a As String
a = Trim(ToAscii(t))
If (a Like "T/M *" Or a Like "TM.*" Or a Like "K/T *" Or a Like "KT.*" Or a Like "T/L *" Or a Like "TL.*" Or a Like "TUQ.*" Or a Like "Q. *") And Len(auth) = 0 And Len(role) = 0 Then
auth = t
ElseIf IsSignatureRoleOrAuth(a) And Len(role) = 0 Then
role = t
ElseIf Not IsSignatureRoleOrAuth(a) Then
signer = t
End If
End If
Next p
Exit Sub
End If
Next c
End If
Next tbl
Exit Sub
SigFail:
End Sub
Public Sub RibbonUnit_LayTuVanBan(control As Object)
Call TT_UnitLayTuVanBan(True)
End Sub
Public Function TT_UnitLayTuVanBan(ByVal showMsg As Boolean) As Boolean
TT_UnitLayTuVanBan = False
If Documents.Count = 0 Then Exit Function
Dim src As Document
Set src = ActiveDocument
Dim strong As Boolean, pid As String, grp As String
pid = TT_DetectFromHeaderText(TT_HeaderAreaAscii(src), strong)
If Len(pid) = 0 Then pid = TT_GetSetting()
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Then
If showMsg Then MsgBox "Khong nhan dien duoc khoi cua van ban dang mo (can co bang tieu de da chuan hoa).", vbExclamation, "Thong tin don vi"
Exit Function
End If
Dim organLines As String, orgAbbr As String, placeName As String
Call TT_ReadOrganInfo(src, grp, organLines, orgAbbr, placeName)
If Len(organLines) = 0 Then
If showMsg Then MsgBox "Van ban dang mo chua co bang tieu de (Ten co quan, So ky hieu). Hay Chuan hoa the thuc truoc.", vbExclamation, "Thong tin don vi"
Exit Function
End If
Dim parts() As String
parts = Split(organLines, vbCr)
Dim capTren As String, coQuan As String
coQuan = parts(UBound(parts))
Dim k As Long
For k = LBound(parts) To UBound(parts) - 1
If Left(parts(k), 1) <> "[" Then
If Len(capTren) > 0 Then capTren = capTren & " | "
capTren = capTren & parts(k)
End If
Next k
Call TT_UnitSet(grp, "CAPTREN", capTren)
Call TT_UnitSet(grp, "COQUAN", coQuan)
If Left(orgAbbr, 1) <> "[" Then Call TT_UnitSet(grp, "VIETTAT", orgAbbr)
If Left(placeName, 1) <> "[" Then Call TT_UnitSet(grp, "DIADANH", placeName)
Dim auth As String, role As String, signer As String
Call TT_ReadSignerInfo(src, auth, role, signer)
If Len(auth) > 0 Then Call TT_UnitSet(grp, "THAMQUYEN", auth)
If Len(role) > 0 Then Call TT_UnitSet(grp, "CHUCVU", role)
If Len(signer) > 0 And Left(signer, 1) <> "[" Then Call TT_UnitSet(grp, "HOTEN", signer)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", grp
On Error GoTo 0
Call TT_RibbonRefresh
TT_UnitLayTuVanBan = True
If showMsg Then MsgBox "Da luu thong tin don vi (khoi " & IIf(grp = "DANG", "Dang", "Nha nuoc") & ") tu van ban dang mo. Xem va sua o nhom 'Thong tin don vi' tren Ribbon.", vbInformation, "Thong tin don vi"
End Function
Public Sub RibbonMau_GetVisible(control As Object, ByRef returnedVal)
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Then
returnedVal = True
Else
returnedVal = (TT_TemplateGroupOf(pid) = control.Tag)
End If
End Sub
Public Sub RibbonMau_GetLabel(control As Object, ByRef returnedVal)
If TT_GetSetting() = "AUTO" Then
If control.Tag = "DANG" Then
returnedVal = "M" & ChrW(7851) & "u kh" & ChrW(7889) & "i " & ChrW(272) & ChrW(7843) & "ng"
Else
returnedVal = "M" & ChrW(7851) & "u kh" & ChrW(7889) & "i Nh" & ChrW(224) & " n" & ChrW(432) & ChrW(7899) & "c"
End If
Else
returnedVal = "T" & ChrW(7841) & "o v" & ChrW(259) & "n b" & ChrW(7843) & "n m" & ChrW(7851) & "u"
End If
End Sub
Public Sub RibbonMoDonVi(control As Object)
On Error GoTo NoActivate
If mTTRibbon Is Nothing Then GoTo NoActivate
mTTRibbon.ActivateTab "tabVanBanDang"
Exit Sub
NoActivate:
MsgBox "Thong tin don vi nam o tab 'THE THUC VAN BAN' tren thanh Ribbon (nhom 'Thong Tin Don Vi').", vbInformation, "Thong tin don vi"
End Sub
Public Sub RibbonTaoMauItem(control As Object)
Call TT_TaoMauTheoTag(control.Tag, True)
End Sub
Public Function TT_TaoMauTheoTag(ByVal tagVal As String, ByVal showMsg As Boolean) As Document
Set TT_TaoMauTheoTag = Nothing
If showMsg Then
If Not CheckLicenseActivation() Then Exit Function
End If
Dim bar As Long
bar = InStr(1, tagVal, "|")
If bar = 0 Then Exit Function
Dim grp As String, nM As Integer
grp = Left(tagVal, bar - 1)
nM = Val(Mid(tagVal, bar + 1))
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Or TT_TemplateGroupOf(pid) <> grp Then
If grp = "DANG" Then
pid = "DANG"
Else
pid = "CHINH_QUYEN"
End If
End If
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", grp
On Error GoTo 0
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Set TT_TaoMauTheoTag = TT_CreateFromTemplate(pid, nM, srcDoc, showMsg)
Call TT_RibbonRefresh
End Function
Public Sub RibbonTaoMau(control As Object)
Call TaoVanBanMau
End Sub
Private Function TT_TemplateGroupOf(ByVal pid As String) As String
Select Case TT_RawS(pid, "MOTTO_KIND")
Case "DANG": TT_TemplateGroupOf = "DANG"
Case "QUOC_HIEU": TT_TemplateGroupOf = "NHA_NUOC"
Case Else: TT_TemplateGroupOf = ""
End Select
End Function
Private Function TT_AskNumber(ByVal msg As String, ByVal title As String, ByVal defaultVal As Integer, ByVal maxVal As Integer) As Integer
TT_AskNumber = 0
Dim ans As String
ans = InputBox(msg, title, CStr(defaultVal))
If Len(Trim(ans)) = 0 Then Exit Function
If Not IsNumeric(ans) Then Exit Function
Dim n As Integer
n = Val(ans)
If n >= 1 And n <= maxVal Then TT_AskNumber = n
End Function
Private Sub TT_ReadOrganInfo(ByVal src As Document, ByVal grp As String, ByRef organLines As String, ByRef orgAbbr As String, ByRef placeName As String)
organLines = "": orgAbbr = "": placeName = ""
On Error GoTo ReadFail
If src Is Nothing Then Exit Sub
If src.Tables.Count = 0 Then Exit Sub
Dim tbl As Table
Set tbl = src.Tables(1)
Dim tAsc As String
tAsc = ToAscii(tbl.Range.Text)
Dim strong As Boolean
Dim srcPid As String
srcPid = TT_DetectFromHeaderText(tAsc, strong)
If Len(srcPid) = 0 Then Exit Sub
If TT_TemplateGroupOf(srcPid) <> grp Then Exit Sub
Dim c As Cell
Set c = tbl.Range.Cells(1)
Dim p As Paragraph
For Each p In c.Range.Paragraphs
Dim t As String
t = CleanParaText(p.Range.Text)
Dim a As String
a = Trim(ToAscii(t))
If Len(a) > 0 Then
If IsDocNumberLineAsc(a) Then
Dim slashPos As Long
slashPos = InStr(1, t, "/")
If slashPos > 0 Then
Dim tail As String
tail = Trim(Mid(t, slashPos + 1))
If DocNumberStyleOf(t) = "NHA_NUOC" Then
Dim dashPos As Long
dashPos = InStr(1, tail, "-")
If dashPos > 0 Then
If Len(TT_RawDocCodeTitle(srcPid, ToAscii(Left(tail, dashPos - 1)))) > 0 Then
orgAbbr = Trim(Mid(tail, dashPos + 1))
Else
orgAbbr = Trim(Left(tail, dashPos - 1))
End If
Else
orgAbbr = tail
End If
Else
orgAbbr = tail
End If
End If
Exit For
ElseIf a = "*" Or a = "***" Or IsSeparatorLine(t) Then
ElseIf a Like "V/V*" Then
Exit For
Else
Call AppendLineToBuffer(organLines, t)
End If
End If
Next p
If grp = "DANG" And Len(organLines) > 0 And InStr(1, organLines, vbCr) = 0 Then
Dim oneAsc As String
oneAsc = Trim(ToAscii(organLines))
If oneAsc Like "DANG UY XA *" Or oneAsc Like "DANG UY PHUONG *" Or oneAsc Like "DANG UY DAC KHU *" Or oneAsc Like "CHI BO *" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & organLines
End If
End If
For Each p In tbl.Range.Paragraphs
Dim tD As String
tD = CleanParaText(p.Range.Text)
Dim aD As String
aD = ToAscii(tD)
If InStr(1, aD, "NGAY") > 0 And InStr(1, aD, "THANG") > 0 Then
Dim commaPos As Long
commaPos = InStr(1, tD, ",")
If commaPos > 1 Then placeName = Trim(Left(tD, commaPos - 1))
Exit For
End If
Next p
Exit Sub
ReadFail:
organLines = "": orgAbbr = "": placeName = ""
End Sub
Private Function TT_PartySuperiorPlaceholder() As String
TT_PartySuperiorPlaceholder = "[" & ChrW(272) & ChrW(7842) & "NG B" & ChrW(7896) & " C" & ChrW(7844) & "P TR" & ChrW(202) & "N TR" & ChrW(7920) & "C TI" & ChrW(7870) & "P]"
End Function
Private Function TT_HighlightPlaceholders(ByVal d As Document) As Long
Dim cnt As Long
cnt = 0
On Error Resume Next
Dim p As Paragraph
For Each p In d.Paragraphs
Dim t As String
t = p.Range.Text
Dim pos As Long
pos = InStr(1, t, "[")
Dim guard As Integer
guard = 0
Do While pos > 0 And guard < 50
guard = guard + 1
Dim pe As Long
pe = InStr(pos + 1, t, "]")
If pe = 0 Then Exit Do
d.Range(p.Range.Start + pos - 1, p.Range.Start + pe).HighlightColorIndex = wdYellow
cnt = cnt + 1
pos = InStr(pe + 1, t, "[")
Loop
Next p
On Error GoTo 0
TT_HighlightPlaceholders = cnt
End Function
Public Sub TaoVanBanMau()
If Not CheckLicenseActivation() Then Exit Sub
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Then
Dim msgK As String
Dim i As Integer
msgK = "Tao van ban mau cho khoi nao? (go so roi bam OK)" & vbCrLf
For i = 1 To TT_ProfileCount()
msgK = msgK & vbCrLf & i & ". " & TT_ProfileLabelAscii(TT_ProfileIdAt(i))
Next i
Dim nK As Integer
nK = TT_AskNumber(msgK, "Tao van ban mau", 1, TT_ProfileCount())
If nK = 0 Then Exit Sub
pid = TT_ProfileIdAt(nK)
End If
Dim grp As String
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Or TT_MauCount(grp) = 0 Then
MsgBox "Khoi '" & TT_ProfileLabelAscii(pid) & "' chua co mau van ban (bo the thuc dang thu nghiem).", vbInformation, "Tao van ban mau"
Exit Sub
End If
Dim msgM As String
msgM = "Chon mau van ban - khoi " & TT_ProfileLabelAscii(pid) & " (go so roi bam OK):" & vbCrLf
For i = 1 To TT_MauCount(grp)
msgM = msgM & vbCrLf & i & ". " & TT_MauLabelAscii(grp, i)
Next i
Dim nM As Integer
nM = TT_AskNumber(msgM, "Tao van ban mau", 1, TT_MauCount(grp))
If nM = 0 Then Exit Sub
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Call TT_CreateFromTemplate(pid, nM, srcDoc, True)
End Sub
Public Function TT_CreateFromTemplate(ByVal pid As String, ByVal nM As Integer, ByVal srcDoc As Document, ByVal showMsg As Boolean) As Document
Set TT_CreateFromTemplate = Nothing
Dim grp As String
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Or nM < 1 Or nM > TT_MauCount(grp) Then Exit Function
Dim organLines As String, orgAbbr As String, placeName As String
Dim infoSource As String
infoSource = ""
Dim uCoQuan As String
uCoQuan = TT_UnitGet(grp, "COQUAN")
If Len(uCoQuan) > 0 Then
Dim uCapTren As String
uCapTren = TT_UnitGet(grp, "CAPTREN")
If Len(uCapTren) > 0 Then
organLines = Replace(Replace(uCapTren, " | ", vbCr), "|", vbCr) & vbCr & Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
ElseIf grp = "DANG" Then
organLines = Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
Dim uAsc As String
uAsc = Trim(ToAscii(organLines))
If uAsc Like "DANG UY XA *" Or uAsc Like "DANG UY PHUONG *" Or uAsc Like "DANG UY DAC KHU *" Or uAsc Like "CHI BO *" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & organLines
End If
Else
organLines = Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
End If
orgAbbr = TT_UnitGet(grp, "VIETTAT")
placeName = TT_UnitGet(grp, "DIADANH")
infoSource = "DONVI"
Else
Call TT_ReadOrganInfo(srcDoc, grp, organLines, orgAbbr, placeName)
If Len(organLines) > 0 Then infoSource = "VANBAN"
End If
Dim tookInfo As Boolean
tookInfo = (Len(infoSource) > 0)
Application.ScreenUpdating = False
Dim d As Document
Set d = Documents.Add
Call TT_BuildTemplateDoc(d, grp, TT_MauData(grp, nM), organLines, orgAbbr, placeName, _
TT_UnitGet(grp, "THAMQUYEN"), TT_UnitGet(grp, "CHUCVU"), TT_UnitGet(grp, "HOTEN"))
d.Activate
mTTForceProfile = pid
Call ChuanHoa_VanBanDang(True)
mTTForceProfile = ""
On Error Resume Next
d.Variables("TheThuc_BaoCao").Delete
On Error GoTo 0
Dim nPh As Long
nPh = TT_HighlightPlaceholders(d)
Application.ScreenUpdating = True
On Error Resume Next
Dim pSel As Paragraph
For Each pSel In d.Paragraphs
If Not pSel.Range.Information(wdWithInTable) Then
Dim posB As Long
posB = InStr(1, pSel.Range.Text, "[")
If posB > 0 Then
Dim posE As Long
posE = InStr(posB, pSel.Range.Text, "]")
If posE > posB Then d.Range(pSel.Range.Start + posB - 1, pSel.Range.Start + posE).Select
Exit For
End If
End If
Next pSel
On Error GoTo 0
Set TT_CreateFromTemplate = d
If Not showMsg Then Exit Function
Dim info As String
If infoSource = "DONVI" Then
info = "Ten co quan, ky hieu, dia danh, nguoi ky lay tu Thong tin don vi da luu (sua tren Ribbon)."
ElseIf infoSource = "VANBAN" Then
info = "Ten co quan, ky hieu, dia danh lay tu van ban dang mo."
Else
info = "Chua nhap Thong tin don vi tren Ribbon: ten co quan, ky hieu, dia danh de trong cho dien."
End If
MsgBox "Da tao mau: " & TT_MauLabelAscii(grp, nM) & " (khoi " & TT_ProfileLabelAscii(pid) & ")." & vbCrLf & _
info & vbCrLf & "Con " & nPh & " cho can dien trong dau [ ] (to vang). So va ngay thang de trong cho van thu.", _
vbInformation, "Tao van ban mau"
End Function
Public Sub TT_KiemThuDonVi()
Dim baseDir As String
baseDir = Environ("HOME")
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
If Documents.Count = 0 Then Exit Sub
Dim src As Document
Set src = ActiveDocument
Dim ok As Boolean
ok = TT_UnitLayTuVanBan(False)
Dim grp As String
grp = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", "DANG")
Call TT_DebugLog("LayTuVanBan=" & ok & " nhom=" & grp & " | CAPTREN=" & ToAscii(TT_UnitGet(grp, "CAPTREN")) & " | COQUAN=" & ToAscii(TT_UnitGet(grp, "COQUAN")) & _
" | VIETTAT=" & ToAscii(TT_UnitGet(grp, "VIETTAT")) & " | DIADANH=" & ToAscii(TT_UnitGet(grp, "DIADANH")) & _
" | THAMQUYEN=" & ToAscii(TT_UnitGet(grp, "THAMQUYEN")) & " | CHUCVU=" & ToAscii(TT_UnitGet(grp, "CHUCVU")) & " | HOTEN=" & ToAscii(TT_UnitGet(grp, "HOTEN")))
src.Close SaveChanges:=False
Dim idxCV As Integer
Dim i As Integer
For i = 1 To TT_MauCount(grp)
If TT_MauLabelAscii(grp, i) = "Cong van" Then idxCV = i
Next i
Dim d As Document
Set d = TT_TaoMauTheoTag(grp & "|" & idxCV, False)
If Not d Is Nothing Then
On Error Resume Next
d.SaveAs FileName:=baseDir & "/Documents/thethuc_test/out/donvi_" & grp & ".docx", FileFormat:=wdFormatXMLDocument
d.Close SaveChanges:=False
On Error GoTo 0
End If
End Sub
Public Sub TT_KiemThuTatCaMau()
Dim baseDir As String
baseDir = Environ("HOME")
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim outDir As String
outDir = baseDir & "/Documents/thethuc_test/out/"
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Dim tag As String
If srcDoc Is Nothing Then
tag = "trong"
Else
tag = "nguon"
End If
Dim pids(1 To 3) As String
pids(1) = "DANG": pids(2) = "CHINH_QUYEN": pids(3) = "CONG_AN"
Dim ip As Integer, im As Integer
For ip = 1 To 3
For im = 1 To TT_MauCount(TT_TemplateGroupOf(pids(ip)))
Dim d As Document
Set d = TT_CreateFromTemplate(pids(ip), im, srcDoc, False)
If Not d Is Nothing Then
On Error Resume Next
d.Variables.Add "TheThuc_BaoCao", "Mau " & pids(ip) & " " & im
d.SaveAs FileName:=outDir & "mau_" & tag & "_" & pids(ip) & "_" & Format(im, "00") & ".docx", FileFormat:=wdFormatXMLDocument
d.Close SaveChanges:=False
On Error GoTo 0
Call TT_DebugLog("Tao mau " & pids(ip) & " " & im & " xong")
End If
Next im
Next ip
If Not srcDoc Is Nothing Then srcDoc.Activate
End Sub
Private Sub TT_BuildTemplateDoc(ByVal d As Document, ByVal grp As String, ByVal data As String, _
ByVal organLines As String, ByVal orgAbbr As String, ByVal placeName As String, _
ByVal uAuth As String, ByVal uRole As String, ByVal uName As String)
On Error Resume Next
Dim code As String, isCV As Boolean, summary As String
Dim bodyText As String, recipText As String
Dim auth As String, role As String, signer As String
Dim lRole As String, lName As String, hasLeftSigner As Boolean
isCV = False: hasLeftSigner = False
Dim items() As String
items = Split(data, vbLf)
Dim k As Long
For k = LBound(items) To UBound(items)
Dim it As String
it = items(k)
Dim bar As Long
bar = InStr(1, it, "|")
If bar > 0 Then
Dim tg As String, v As String
tg = Left(it, bar - 1)
v = Mid(it, bar + 1)
Select Case tg
Case "CODE": code = v
Case "CV": isCV = True
Case "SUM": summary = v
Case "L": Call AppendLineToBuffer(bodyText, v)
Case "N": Call AppendLineToBuffer(recipText, v)
Case "A": auth = v
Case "R": role = v
Case "H": signer = v
Case "LR": lRole = v: hasLeftSigner = True
Case "LH": lName = v
End Select
End If
Next k
If Not hasLeftSigner Then
If Len(uRole) > 0 Or Len(uName) > 0 Then
auth = uAuth
ElseIf Len(uAuth) > 0 Then
auth = uAuth
End If
If Len(uRole) > 0 Then role = uRole
If Len(uName) > 0 Then signer = uName
End If
If Len(organLines) = 0 Then
If grp = "DANG" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & _
"[" & ChrW(272) & ChrW(7842) & "NG " & ChrW(7910) & "Y / CHI B" & ChrW(7896) & " BAN H" & ChrW(192) & "NH]"
Else
organLines = "[T" & ChrW(202) & "N C" & ChrW(416) & " QUAN C" & ChrW(7844) & "P TR" & ChrW(202) & "N (N" & ChrW(7870) & "U C" & ChrW(211) & ")]" & vbCr & _
"[T" & ChrW(202) & "N C" & ChrW(416) & " QUAN BAN H" & ChrW(192) & "NH]"
End If
End If
If Len(orgAbbr) = 0 Then orgAbbr = "[VI" & ChrW(7870) & "T T" & ChrW(7854) & "T]"
If Len(placeName) = 0 Then placeName = "[" & ChrW(272) & ChrW(7883) & "a danh]"
Dim soLine As String
If grp = "DANG" Then
soLine = "S" & ChrW(7889) & "   -" & code & "/" & orgAbbr
ElseIf isCV Then
soLine = "S" & ChrW(7889) & ":   /" & orgAbbr & "-[" & ChrW(272) & "VST]"
Else
soLine = "S" & ChrW(7889) & ":   /" & code & "-" & orgAbbr
End If
Dim leftText As String
leftText = organLines
If grp = "DANG" Then Call AppendLineToBuffer(leftText, "*")
Call AppendLineToBuffer(leftText, soLine)
If isCV And Len(summary) > 0 Then Call AppendLineToBuffer(leftText, summary)
Dim rightText As String
If grp = "DANG" Then
rightText = TT_TextTieuNguDang()
Else
rightText = TT_TextQuocHieu() & vbCr & TT_TextTieuNgu()
End If
Call AppendLineToBuffer(rightText, placeName & ", ng" & ChrW(224) & "y    th" & ChrW(225) & "ng    n" & ChrW(259) & "m " & Year(Date))
Dim tH As Table
Set tH = d.Tables.Add(d.Range(0, 0), 1, 2)
tH.Borders.Enable = False
tH.Cell(1, 1).Range.Text = leftText
tH.Cell(1, 2).Range.Text = rightText
Dim rB As Range
Set rB = d.Paragraphs(d.Paragraphs.Count).Range
rB.InsertBefore bodyText & vbCr
Dim leftF As String, rightF As String
If hasLeftSigner Then
leftF = lRole & vbCr & vbCr & vbCr & vbCr & vbCr & lName
ElseIf Len(recipText) > 0 Then
leftF = "N" & ChrW(417) & "i nh" & ChrW(7853) & "n:" & vbCr & recipText
End If
If Len(auth) > 0 Then rightF = Replace(Replace(auth, " | ", vbCr), "|", vbCr) & vbCr
rightF = rightF & Replace(Replace(role, " | ", vbCr), "|", vbCr) & vbCr & vbCr & vbCr & vbCr & vbCr & signer
Dim rF As Range
Set rF = d.Paragraphs(d.Paragraphs.Count).Range
Dim tF As Table
Set tF = d.Tables.Add(rF, 1, 2)
tF.Borders.Enable = False
tF.Cell(1, 1).Range.Text = leftF
tF.Cell(1, 2).Range.Text = rightF
On Error GoTo 0
End Sub
Public Sub RibbonChuanHoa(control As Object)
Call ChuanHoa_VanBanDang(False)
End Sub
Public Sub ChuanHoa_VanBanDang_Silent()
Call ChuanHoa_VanBanDang(True)
End Sub
Sub ChuanHoa_VanBanDang(Optional ByVal silentMode As Boolean = False)
If Documents.Count = 0 Then Exit Sub
If Not silentMode Then
Call CN_TuKiemTraHangTuan(True)
If Not CheckLicenseActivation() Then Exit Sub
End If
Application.ScreenUpdating = False
On Error Resume Next
Application.UndoRecord.StartCustomRecord "Chuan hoa the thuc"
Dim doc As Document
Set doc = ActiveDocument
Call TT_DebugLog("BAT DAU: " & doc.Name & " | doan van: " & doc.Paragraphs.Count)
Call TT_Resolve(doc)
mTTQPPL = TT_DetectQPPL(doc)
mTTNotes = ""
Call TT_DebugLog("Khoi: " & TT_Profile() & " | VBQPPL: " & CStr(mTTQPPL))
On Error Resume Next
doc.ConvertNumbersToText
Call TT_DebugLog("0.0 ConvertNumbersToText xong")
On Error Resume Next
Call RemoveHeaderFloatingShapes(doc, Nothing)
Call TT_DebugLog("0.1 RemoveHeaderFloatingShapes xong")
Call SafeReplaceSoftBreaksInRange(doc.Content)
Call TT_SuaLoiGo(doc)
Call TT_DebugLog("0.2 SafeReplaceSoftBreaks xong")
Call AutoCreateHeaderTable(doc)
Call TT_DebugLog("0.3a AutoCreateHeaderTable xong")
Call SplitGluedHeaderAndTitleLines(doc)
Call TT_DebugLog("0.3b SplitGlued xong")
Call AutoCreateKinhGuiTable(doc)
Call TT_DebugLog("0.3c AutoCreateKinhGuiTable xong")
Call AutoCreateFooterTable(doc)
Call TT_DebugLog("0.3d AutoCreateFooterTable xong")
Dim pIdx As Long
For pIdx = doc.Paragraphs.Count To 2 Step -1
Dim pCur As Paragraph
Set pCur = doc.Paragraphs(pIdx)
If Len(CleanParaText(pCur.Range.Text)) = 0 Then
Dim pPrev As Paragraph
Set pPrev = doc.Paragraphs(pIdx - 1)
If Len(CleanParaText(pPrev.Range.Text)) = 0 Then
If Not pCur.Range.Information(wdWithInTable) And Not pPrev.Range.Information(wdWithInTable) Then
On Error Resume Next
pCur.Range.Delete
On Error Resume Next
End If
End If
End If
Next pIdx
Call TT_DebugLog("BUOC 1 bat dau")
Dim sec As Section
For Each sec In doc.Sections
On Error Resume Next
With sec.PageSetup
.PaperSize = wdPaperA4
.TopMargin = MillimetersToPoints(TT_N("PAGE_TOP"))
.BottomMargin = MillimetersToPoints(TT_N("PAGE_BOTTOM"))
.LeftMargin = MillimetersToPoints(TT_N("PAGE_LEFT"))
.RightMargin = MillimetersToPoints(TT_N("PAGE_RIGHT"))
.HeaderDistance = MillimetersToPoints(TT_N("PAGE_HEADER"))
.FooterDistance = MillimetersToPoints(TT_N("PAGE_FOOTER"))
.MirrorMargins = False
.DifferentFirstPageHeaderFooter = True
End With
On Error Resume Next
Next sec
Call TT_DebugLog("BUOC 2 bat dau")
For Each sec In doc.Sections
On Error Resume Next
Dim hdr As HeaderFooter
Set hdr = sec.Headers(wdHeaderFooterPrimary)
hdr.Range.Delete
hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
hdr.Range.Font.Name = "Times New Roman"
hdr.Range.Font.Size = TT_N("PAGE_NUMBER_SIZE")
hdr.Range.Font.Bold = False
hdr.Range.Font.Italic = False
hdr.Range.Font.ColorIndex = wdBlack
hdr.Range.Fields.Add Range:=hdr.Range, Type:=wdFieldPage
hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
sec.Headers(wdHeaderFooterFirstPage).Range.Delete
On Error Resume Next
Next sec
Call TT_DebugLog("BUOC 3 bat dau")
Dim tFit As Table
For Each tFit In doc.Tables
On Error Resume Next
tFit.AutoFitBehavior 2
tFit.AllowAutoFit = True
On Error Resume Next
Next tFit
Dim tbl As Table
Dim tblCounter As Integer
tblCounter = 0
Dim c As Cell
Dim pCell As Paragraph
Dim pDCS As Paragraph
Dim cRight As Cell
For Each tbl In doc.Tables
tblCounter = tblCounter + 1
Dim tAsc As String
tAsc = ToAscii(tbl.Range.Text)
Dim isHeaderTbl As Boolean
isHeaderTbl = False
Dim hdrDetectStrong As Boolean
If tblCounter = 1 And (InStr(1, tAsc, "DANG CONG SAN") > 0 Or InStr(1, tAsc, "VIET NAM") > 0 Or InStr(1, tAsc, "DANG BO") > 0 Or InStr(1, tAsc, "DANG UY") > 0 Or InStr(1, tAsc, "CHI BO") > 0 Or InStr(1, tAsc, "TINH UY") > 0 Or InStr(1, tAsc, "THANH UY") > 0 Or InStr(1, tAsc, "HUYEN UY") > 0) Then
isHeaderTbl = True
ElseIf tblCounter = 1 And tbl.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(tAsc, hdrDetectStrong)) > 0 And InStr(1, tAsc, "NOI NHAN") = 0 Then
isHeaderTbl = True
ElseIf InStr(1, tAsc, "DANG CONG SAN VIET NAM") > 0 Or InStr(1, tAsc, "CONG HOA XA HOI CHU NGHIA") > 0 Then
isHeaderTbl = True
End If
If isHeaderTbl Then
Set pDCS = Nothing
Set cRight = Nothing
Dim pQuocHieu As Paragraph: Set pQuocHieu = Nothing
Dim pTieuNgu As Paragraph: Set pTieuNgu = Nothing
Dim pNgayThangHdr As Paragraph: Set pNgayThangHdr = Nothing
Dim pLastOrgan As Paragraph: Set pLastOrgan = Nothing
Dim organParas As Collection: Set organParas = New Collection
Call RemoveHeaderFloatingShapes(doc, tbl)
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim totalCells As Integer
totalCells = tbl.Range.Cells.Count
Dim cellIdx As Integer
cellIdx = 1
Dim leftColCqIdx As Integer
leftColCqIdx = 1
Dim leftColInTrichYeu As Boolean
leftColInTrichYeu = False
Dim leftColSawSo As Boolean
leftColSawSo = False
Dim sawKinhGuiRowInHeaderTbl As Boolean
sawKinhGuiRowInHeaderTbl = False
For Each c In tbl.Range.Cells
Dim cellTextRaw As String
cellTextRaw = CleanParaText(c.Range.Text)
If Len(cellTextRaw) = 0 Then GoTo NextCellHeader
Dim cAsc As String
cAsc = ToAscii(c.Range.Text)
If HasKinhGuiOrTrinh(cAsc) Then
sawKinhGuiRowInHeaderTbl = True
GoTo NextCellHeader
End If
If sawKinhGuiRowInHeaderTbl Or Left(cellTextRaw, 1) = "-" Then
GoTo NextCellHeader
End If
Dim isRightCell As Boolean
isRightCell = False
If InStr(1, cAsc, "DANG CONG SAN") > 0 Or InStr(1, cAsc, "DOC LAP - TU DO") > 0 Or InStr(1, cAsc, "CONG HOA XA HOI") > 0 Then
isRightCell = True
Set cRight = c
ElseIf InStr(1, cAsc, "NGAY") > 0 And InStr(1, cAsc, "THANG") > 0 And InStr(1, cAsc, "NAM") > 0 And InStr(1, cAsc, "V/V") = 0 Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
ElseIf InStr(1, cAsc, "DANG BO") > 0 Or InStr(1, cAsc, "DANG UY") > 0 Or InStr(1, cAsc, "CHI BO") > 0 Or _
InStr(1, cAsc, "TINH UY") > 0 Or InStr(1, cAsc, "THANH UY") > 0 Or InStr(1, cAsc, "HUYEN UY") > 0 Or _
InStr(1, cAsc, "SO ") > 0 Or InStr(1, cAsc, "SO:") > 0 Or InStr(1, cAsc, "V/V") > 0 Or InStr(1, cAsc, "VE VIEC") > 0 Then
isRightCell = False
Else
On Error Resume Next
If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
ElseIf cellIdx = totalCells And totalCells = 2 Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
End If
On Error Resume Next
End If
On Error Resume Next
c.Borders.Enable = False
c.Shading.BackgroundPatternColor = wdColorAutomatic
c.Shading.Texture = wdTextureNone
Call SafeReplaceSoftBreaksInRange(c.Range)
On Error Resume Next
Call RemoveEmptyParagraphsInCell(doc, c)
For Each pCell In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pCell)
Dim pAsc As String
pAsc = Trim(ToAscii(pCell.Range.Text))
If Len(pAsc) = 0 Then GoTo SkipEmptyLineHeader
With pCell.Format
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
If isRightCell Then
pCell.Format.Alignment = TT_RightColumnAlignment()
If InStr(1, pAsc, "CONG HOA XA HOI") > 0 Then
If Len(pAsc) <= 45 And CleanParaText(pCell.Range.Text) <> TT_TextQuocHieu() Then Call SafeSetParagraphText(doc, pCell, TT_TextQuocHieu())
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Size = TT_N("QH_SIZE")
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.Alignment = wdAlignParagraphCenter
Set pQuocHieu = pCell
ElseIf InStr(1, Replace(pAsc, " ", ""), "DOCLAP-TUDO") > 0 Then
If CleanParaText(pCell.Range.Text) <> TT_TextTieuNgu() And InStr(1, Replace(pAsc, " ", ""), "DOCLAP-TUDO-HANHPHUC") > 0 Then
Call SafeSetParagraphText(doc, pCell, TT_TextTieuNgu())
End If
pCell.Range.Font.Size = TT_N("TN_SIZE")
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Format.SpaceAfter = 8
Set pTieuNgu = pCell
ElseIf InStr(1, pAsc, "DANG CONG SAN") > 0 Or InStr(1, pAsc, "VIET NAM") > 0 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Size = 15
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.SpaceAfter = 4
Set pDCS = pCell
ElseIf InStr(1, pAsc, "NGAY") > 0 And InStr(1, pAsc, "THANG") > 0 Then
Set pNgayThangHdr = pCell
Dim origNgay As String, fixNgay As String
origNgay = CleanParaText(pCell.Range.Text)
fixNgay = ChuanHoa_NgayThang(origNgay)
If fixNgay <> origNgay And Len(fixNgay) > 0 Then
Call SafeSetParagraphText(doc, pCell, fixNgay)
End If
pCell.Format.Alignment = TT_RightColumnAlignment()
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("DATE_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = True
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
ElseIf pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
Call SafeSetParagraphText(doc, pCell, "")
ElseIf TT_S("MOTTO_KIND") = "LINH_HOAT" Then
pCell.Range.Font.Size = 14
Else
pCell.Range.Font.Size = 14
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
End If
Else
If leftColSawSo Then
If pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = 10
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
Else
leftColInTrichYeu = True
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
End If
ElseIf pAsc Like "V/V*" Or pAsc Like "V.V*" Or pAsc Like "VE VIEC*" Or InStr(1, pAsc, "V/V") > 0 Or InStr(1, pAsc, "V.V") > 0 Then
leftColInTrichYeu = True
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
ElseIf IsDocNumberLineAsc(pAsc) Then
leftColSawSo = True
Dim origSo As String, fixSo As String
origSo = CleanParaText(pCell.Range.Text)
fixSo = ChuanHoa_SoKyHieu(origSo)
If fixSo <> origSo And Len(fixSo) > 0 Then
Call SafeSetParagraphText(doc, pCell, fixSo)
End If
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("DOC_NUMBER_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
ElseIf pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
If TT_B("ORGAN_UNDERLINE") Then
Call SafeSetParagraphText(doc, pCell, "")
Else
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = 14
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
End If
Else
Set pLastOrgan = pCell
organParas.Add pCell
pCell.Format.Alignment = wdAlignParagraphCenter
Call TT_ChuanUY(pCell.Range)
pCell.Range.Font.Size = TT_N("ORGAN_SIZE")
If leftColCqIdx = 1 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
ElseIf leftColCqIdx = 2 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
Else
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
End If
leftColCqIdx = leftColCqIdx + 1
End If
End If
SkipEmptyLineHeader:
Next pCell
NextCellHeader:
cellIdx = cellIdx + 1
Next c
If TT_IsNational() And leftColCqIdx = 3 Then
On Error Resume Next
Dim orgFirstP As Paragraph
Set orgFirstP = tbl.Range.Cells(1).Range.Paragraphs(1)
Dim orgFirstAsc As String
orgFirstAsc = Trim(ToAscii(CleanParaText(orgFirstP.Range.Text)))
If orgFirstAsc = "UY BAN NHAN DAN" Or orgFirstAsc = "HOI DONG NHAN DAN" Then orgFirstP.Range.Font.Bold = True
On Error Resume Next
End If
If leftColCqIdx = 2 Then
On Error Resume Next
Dim firstLeftP As Paragraph
Set firstLeftP = tbl.Range.Cells(1).Range.Paragraphs(1)
firstLeftP.Range.Font.Bold = True
On Error Resume Next
End If
If TT_IsNational() Then
On Error Resume Next
If tbl.Columns.Count = 2 Then
Dim hdrTotalW As Single
hdrTotalW = 0
With tbl.Range.Sections(1).PageSetup
hdrTotalW = .PageWidth - .LeftMargin - .RightMargin
End With
If hdrTotalW < 300 Or hdrTotalW > 800 Then hdrTotalW = MillimetersToPoints(165)
Dim hdrRightW As Single
hdrRightW = MillimetersToPoints(TT_N("HDR_RIGHT_MM"))
Dim hdrMinLeftW As Single
hdrMinLeftW = MillimetersToPoints(TT_N("HDR_LEFT_MIN_MM"))
If hdrTotalW - hdrRightW < hdrMinLeftW Then hdrRightW = hdrTotalW - hdrMinLeftW
tbl.AllowAutoFit = False
If tbl.Rows.LeftIndent < -MillimetersToPoints(3) Then tbl.Rows.LeftIndent = 0
Dim hdrLanW As Single: hdrLanW = 0
tbl.Columns(1).Width = hdrTotalW - hdrRightW
tbl.Columns(2).Width = hdrRightW
If Not (pQuocHieu Is Nothing) Then
Dim shiftStep As Integer
Dim qhLines As Long
For shiftStep = 1 To 40
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines <= 1 Then Exit For
If hdrTotalW - hdrRightW - MillimetersToPoints(2) < hdrMinLeftW Then Exit For
hdrRightW = hdrRightW + MillimetersToPoints(2)
tbl.Columns(1).Width = hdrTotalW - hdrRightW
tbl.Columns(2).Width = hdrRightW
Next shiftStep
Call TT_DebugLog("  Cot Hdr: trai " & Format(PointsToMillimeters(hdrTotalW - hdrRightW), "0") & "mm, phai " & Format(PointsToMillimeters(hdrRightW), "0") & "mm, so lan doi " & (shiftStep - 1))
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines > 1 Then
pQuocHieu.Range.Font.Size = 12
If Not (pTieuNgu Is Nothing) Then pTieuNgu.Range.Font.Size = 13
If Not (pNgayThangHdr Is Nothing) Then pNgayThangHdr.Range.Font.Size = 13
Call TT_DebugLog("  Quoc hieu van xuong dong -> giam dong bo 12/13/13")
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines > 1 Then
mTTHeaderNote = "[Quoc hieu] Quoc hieu van bi xuong dong du da noi rong cot va giam co chu 12/13/13. Kiem tra lai do rong bang tieu de (Quoc hieu phai nam tren 1 dong)."
End If
End If
If TT_OrganWraps(organParas) Then Set organParas = TT_SplitLongOrgans(organParas, True): If organParas.Count > 0 Then Set pLastOrgan = organParas(organParas.Count)
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
Call TT_DebugLog("  Truoc buoc ten dai (" & TT_TenDaiMode() & "): cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0") & ", rot dong " & CStr(TT_OrganWraps(organParas)) & ", QH dong " & TT_ParaLines(pQuocHieu))
If TT_TenDaiMode() = "LAN_LE" Then
Call TT_LanLeTrai(tbl, organParas, hdrTotalW, hdrRightW, hdrLanW)
Else
If TT_OrganWraps(organParas) And pQuocHieu.Range.Font.Size > 12 Then
pQuocHieu.Range.Font.Size = 12
If Not (pTieuNgu Is Nothing) Then pTieuNgu.Range.Font.Size = 13
If Not (pNgayThangHdr Is Nothing) Then pNgayThangHdr.Range.Font.Size = 13
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
If TT_OrganWraps(organParas) Then
Dim oPara As Variant
For Each oPara In organParas
If oPara.Range.Font.Size > 12 Then oPara.Range.Font.Size = 12
Next oPara
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
End If
If TT_OrganWraps(organParas) Then
Set organParas = TT_SplitLongOrgans(organParas, False)
If organParas.Count > 0 Then Set pLastOrgan = organParas(organParas.Count)
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
Call TT_DebugLog("  Noi cot ten co quan: trai " & Format(PointsToMillimeters(hdrTotalW - hdrRightW), "0") & "mm")
End If
End If
If TT_B("ORGAN_UNDERLINE") And Not (pLastOrgan Is Nothing) Then
If pLastOrgan.Format.SpaceAfter < 6 Then pLastOrgan.Format.SpaceAfter = 6
End If
On Error Resume Next
ElseIf tbl.Columns.Count = 2 And Not (pDCS Is Nothing) Then
On Error Resume Next
Dim dTotalW As Single, dRightW As Single
Dim pDateDang As Paragraph, pR As Paragraph, pRAsc As String
Set pDateDang = Nothing
For Each pR In tbl.Cell(1, 2).Range.Paragraphs
pRAsc = UCase(ToAscii(CleanParaText(pR.Range.Text)))
If InStr(1, pRAsc, "NGAY") > 0 And InStr(1, pRAsc, "NAM") > 0 Then
Set pDateDang = pR
Exit For
End If
Next pR
Dim dLanW As Single: dLanW = 0
If tbl.Rows.LeftIndent < -MillimetersToPoints(3) Then
tbl.Rows.LeftIndent = 0
With tbl.Range.Sections(1).PageSetup
tbl.Columns(1).Width = .PageWidth - .LeftMargin - .RightMargin - tbl.Columns(2).Width
End With
End If
dTotalW = tbl.Columns(1).Width + tbl.Columns(2).Width
Call TT_DebugLog("  Khoi Dang - co dong ngay thang: " & CStr(Not (pDateDang Is Nothing)) & ", so dong " & TT_ParaLines(pDateDang) & ", ten co quan rot dong: " & CStr(TT_OrganWraps(organParas)) & ", cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0"))
dRightW = tbl.Columns(2).Width
If TT_OrganWraps(organParas) Or TT_ParaLines(pDateDang) > 1 Then tbl.AllowAutoFit = False
If TT_OrganWraps(organParas) Then
Call TT_WidenOrganCol(tbl, organParas, pDCS, pDateDang, dTotalW, dRightW)
If TT_TenDaiMode() = "LAN_LE" Then Call TT_LanLeTrai(tbl, organParas, dTotalW, dRightW, dLanW)
If TT_OrganWraps(organParas) Then
Set organParas = TT_SplitLongOrgans(organParas, False)
Call TT_WidenOrganCol(tbl, organParas, pDCS, pDateDang, dTotalW, dRightW)
End If
End If
Call TT_NarrowForRight(tbl, organParas, pDateDang, dTotalW, dRightW)
Call TT_DebugLog("  Khoi Dang - cot trai " & Format(PointsToMillimeters(dTotalW - dRightW), "0") & "mm")
On Error Resume Next
End If
If Not (pTieuNgu Is Nothing) Then
Call TT_AddCenteredLineUnder(doc, pTieuNgu, "LineTieuNguQuocHieu", 1, CentimetersToPoints(6))
End If
If TT_B("ORGAN_UNDERLINE") And Not (pLastOrgan Is Nothing) Then
Call TT_AddCenteredLineUnder(doc, pLastOrgan, "LineCoQuanBanHanh", 0.4, CentimetersToPoints(5))
End If
If Not (pDCS Is Nothing) And pTieuNgu Is Nothing Then
pDCS.Range.Font.Underline = wdUnderlineNone
On Error Resume Next
pDCS.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
On Error Resume Next
Dim cRightCell As Cell
Set cRightCell = Nothing
On Error Resume Next
Set cRightCell = pDCS.Range.Cells(1)
On Error Resume Next
If cRightCell Is Nothing Then Set cRightCell = cRight
Dim rPad As Single: rPad = 5.4
Dim cellW As Single: cellW = 248
If Not (cRightCell Is Nothing) Then
cellW = cRightCell.Width
On Error Resume Next
rPad = cRightCell.RightPadding
On Error Resume Next
End If
Dim lineWidth As Single: lineWidth = CentimetersToPoints(7.2)
Dim lPos As Single: lPos = cellW - lineWidth - rPad - 5.5
If lPos < 0 Then lPos = 0
On Error Resume Next
Dim newL As Shape
Set newL = doc.Shapes.AddLine(lPos, 16.5, lPos + lineWidth, 16.5, pDCS.Range)
With newL
.Name = "LineTieuNguDang"
.LayoutInCell = True
.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
.Left = lPos
.Top = 16.5
.Line.Weight = 0.75
.Line.ForeColor.RGB = RGB(0, 0, 0)
.Line.DashStyle = 1
.WrapFormat.Type = 3
End With
On Error Resume Next
End If
Exit For
End If
Next tbl
Call TT_DebugLog("BUOC 3.2 bat dau")
For Each tbl In doc.Tables
Dim tAscKG As String
tAscKG = ToAscii(tbl.Range.Text)
If IsKinhGuiTable(tbl, tAscKG) Then
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim cKG As Cell
Dim seenKinhGuiCell As Boolean
seenKinhGuiCell = False
For Each cKG In tbl.Range.Cells
Dim cellTxtKG As String
cellTxtKG = CleanParaText(cKG.Range.Text)
If Len(cellTxtKG) > 0 Then
Dim cAscKG As String
cAscKG = Trim(ToAscii(cellTxtKG))
If InStr(1, cAscKG, "DANG CONG SAN") > 0 Or InStr(1, cAscKG, "CONG HOA XA HOI") > 0 Or _
InStr(1, Replace(cAscKG, " ", ""), "DOCLAP-TUDO") > 0 Or _
(InStr(1, cAscKG, "NGAY") > 0 And InStr(1, cAscKG, "THANG") > 0 And InStr(1, cAscKG, "NAM") > 0) Or _
IsDocNumberLineAsc(cAscKG) Then
GoTo NextCellKG
End If
If HasKinhGuiOrTrinh(cAscKG) Then
seenKinhGuiCell = True
On Error Resume Next
cKG.Borders.Enable = False
cKG.Shading.BackgroundPatternColor = wdColorAutomatic
cKG.Shading.Texture = wdTextureNone
cKG.VerticalAlignment = wdCellAlignVerticalTop
On Error Resume Next
cKG.Range.Font.Bold = False
cKG.Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
Dim pKGLeft As Paragraph
For Each pKGLeft In cKG.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pKGLeft)
Dim pAscLKG As String
pAscLKG = Trim(ToAscii(pKGLeft.Range.Text))
If Len(pAscLKG) > 0 Then
Call SafeSetParagraphText(doc, pKGLeft, KinhGuiPrefixFor(pAscLKG))
With pKGLeft.Format
.Alignment = wdAlignParagraphRight
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pKGLeft.Range.Font.Name = "Times New Roman"
pKGLeft.Range.Font.Size = 14
pKGLeft.Range.Font.Bold = False
pKGLeft.Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
pKGLeft.Range.Font.ColorIndex = wdBlack
pKGLeft.Range.Font.Underline = wdUnderlineNone
End If
Next pKGLeft
ElseIf seenKinhGuiCell Then
On Error Resume Next
cKG.Borders.Enable = False
cKG.Shading.BackgroundPatternColor = wdColorAutomatic
cKG.Shading.Texture = wdTextureNone
cKG.VerticalAlignment = wdCellAlignVerticalTop
Call SafeReplaceSoftBreaksInRange(cKG.Range)
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
Dim lastLineKG As Integer: lastLineKG = 0
Dim counterKG As Integer: counterKG = 1
Dim tPkg As Paragraph
For Each tPkg In cKG.Range.Paragraphs
If Len(CleanParaText(tPkg.Range.Text)) > 0 Then lastLineKG = counterKG
counterKG = counterKG + 1
Next tPkg
Dim pKGRight As Paragraph
Dim lineKGIdx As Integer: lineKGIdx = 1
For Each pKGRight In cKG.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pKGRight)
If Len(CleanParaText(pKGRight.Range.Text)) > 0 Then
If lastLineKG > 1 Then
Dim fChr As String
fChr = pKGRight.Range.Characters(1).Text
If fChr = ChrW(8211) Or fChr = ChrW(8212) Or fChr = ChrW(8210) Or fChr = ChrW(8722) Or _
fChr = ChrW(8226) Or fChr = ChrW(8259) Or fChr = ChrW(183) Or fChr = "+" Or fChr = "*" Then
pKGRight.Range.Characters(1).Text = "-"
ElseIf fChr <> "-" Then
pKGRight.Range.InsertBefore "- "
End If
If pKGRight.Range.Characters(1).Text = "-" And pKGRight.Range.Characters.Count > 2 Then
If pKGRight.Range.Characters(2).Text <> " " Then
pKGRight.Range.Characters(1).InsertAfter " "
End If
End If
ElseIf lastLineKG = 1 Then
Dim fChr1 As String
fChr1 = pKGRight.Range.Characters(1).Text
If fChr1 = "-" Or fChr1 = ChrW(8211) Or fChr1 = ChrW(8212) Or fChr1 = ChrW(8226) Or fChr1 = "+" Or fChr1 = "*" Then
pKGRight.Range.Characters(1).Delete
Call SafeTrimLeadingSpaces(doc, pKGRight)
End If
End If
With pKGRight.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pKGRight.Range.Font.Name = "Times New Roman"
pKGRight.Range.Font.Size = 14
pKGRight.Range.Font.Bold = False
pKGRight.Range.Font.Italic = False
pKGRight.Range.Font.ColorIndex = wdBlack
pKGRight.Range.Font.Underline = wdUnderlineNone
Call NormalizeLineEndingPunctuation(doc, pKGRight, (lineKGIdx = lastLineKG))
End If
lineKGIdx = lineKGIdx + 1
Next pKGRight
End If
End If
NextCellKG:
Next cKG
End If
Next tbl
Call TT_DebugLog("BUOC 3.5 bat dau")
For Each tbl In doc.Tables
Dim tAscFooter As String
tAscFooter = ToAscii(tbl.Range.Text)
If IsFooterSignatureTable(tbl, tAscFooter) Then
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim fCellIdx As Integer: fCellIdx = 1
Dim fTotalCells As Integer: fTotalCells = tbl.Range.Cells.Count
For Each c In tbl.Range.Cells
On Error Resume Next
c.Borders.Enable = False
c.Shading.BackgroundPatternColor = wdColorAutomatic
c.Shading.Texture = wdTextureNone
Call SafeReplaceSoftBreaksInRange(c.Range)
On Error Resume Next
Dim cAscFooter As String
cAscFooter = Trim(ToAscii(c.Range.Text))
If Len(cAscFooter) = 0 Then GoTo NextFooterCell
Dim isSignatureCell As Boolean
isSignatureCell = False
If InStr(1, cAscFooter, "NOI NHAN") > 0 Or Left(CleanParaText(c.Range.Text), 1) = "-" Then
isSignatureCell = False
ElseIf HasSignatureKeywordInCell(cAscFooter) Then
isSignatureCell = True
ElseIf fCellIdx = fTotalCells And fTotalCells >= 2 Then
isSignatureCell = True
Else
On Error Resume Next
If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
isSignatureCell = True
End If
On Error Resume Next
End If
If Not isSignatureCell Then
Dim lastListItemIdx As Integer: lastListItemIdx = 0
Dim tempP As Paragraph
Dim tCounter As Integer: tCounter = 1
For Each tempP In c.Range.Paragraphs
Dim tVal As String
tVal = CleanParaText(tempP.Range.Text)
If Len(tVal) > 0 Then
If Not (LTrim(ToAscii(tVal)) Like "NOI NHAN*") Then
lastListItemIdx = tCounter
End If
End If
tCounter = tCounter + 1
Next tempP
Dim pNoiNhan As Paragraph
Dim pIndexNN As Integer: pIndexNN = 1
For Each pNoiNhan In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pNoiNhan)
Dim pTextNN As String
pTextNN = CleanParaText(pNoiNhan.Range.Text)
If Len(pTextNN) = 0 Then GoTo SkipNNLine
On Error Resume Next
pNoiNhan.Borders.Enable = False
pNoiNhan.Shading.BackgroundPatternColor = wdColorAutomatic
pNoiNhan.Shading.Texture = wdTextureNone
On Error Resume Next
With pNoiNhan.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pNoiNhan.Range.Font.Name = "Times New Roman"
pNoiNhan.Range.Font.ColorIndex = wdBlack
Dim pTextAsc As String
pTextAsc = ToAscii(pTextNN)
If LTrim(pTextAsc) Like "NOI NHAN*" Then
pNoiNhan.Range.Font.Size = TT_N("NN_LABEL_SIZE")
If TT_B("NN_LABEL_UNDERLINE") Then
pNoiNhan.Range.Font.Underline = wdUnderlineSingle
Else
pNoiNhan.Range.Font.Underline = wdUnderlineNone
End If
pNoiNhan.Range.Font.Bold = TT_B("NN_LABEL_BOLD")
pNoiNhan.Range.Font.Italic = TT_B("NN_LABEL_ITALIC")
Else
pNoiNhan.Range.Font.Size = TT_N("NN_ITEM_SIZE")
pNoiNhan.Range.Font.Underline = wdUnderlineNone
pNoiNhan.Range.Font.Bold = False
pNoiNhan.Range.Font.Italic = False
pNoiNhan.Format.LeftIndent = 0
pNoiNhan.Format.FirstLineIndent = 0
Dim firstChar As String
firstChar = Left(pTextNN, 1)
If firstChar <> "-" Then
If firstChar = ChrW(8211) Or firstChar = ChrW(8212) Or firstChar = "*" Or firstChar = "+" Or firstChar = ChrW(8226) Then
pNoiNhan.Range.Characters(1).Text = "-"
Else
pNoiNhan.Range.InsertBefore "- "
End If
End If
If pNoiNhan.Range.Characters(1).Text = "-" And pNoiNhan.Range.Characters.Count > 2 Then
If pNoiNhan.Range.Characters(2).Text <> " " Then
pNoiNhan.Range.Characters(1).InsertAfter " "
End If
End If
Call NormalizeLineEndingPunctuation(doc, pNoiNhan, (pIndexNN = lastListItemIdx))
End If
SkipNNLine:
pIndexNN = pIndexNN + 1
Next pNoiNhan
Else
On Error Resume Next
Dim textIndices(1 To 30) As Integer
Dim textCount As Integer: textCount = 0
Dim pScan As Paragraph
Dim scanIdx As Integer: scanIdx = 1
For Each pScan In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pScan)
If Len(CleanParaText(pScan.Range.Text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
Dim lastLineHasSignerName As Boolean
lastLineHasSignerName = False
If textCount >= 2 Then
Dim lastSigLineAsc As String
lastSigLineAsc = Trim(ToAscii(CleanParaText(c.Range.Paragraphs(textIndices(textCount)).Range.Text)))
If Not IsSignatureRoleOrAuth(lastSigLineAsc) Then
lastLineHasSignerName = True
End If
End If
If textCount >= 2 And lastLineHasSignerName Then
If textCount >= 3 Then
Dim delIdx As Integer
For delIdx = textIndices(textCount - 1) - 1 To textIndices(1) + 1 Step -1
If Len(CleanParaText(c.Range.Paragraphs(delIdx).Range.Text)) = 0 Then
c.Range.Paragraphs(delIdx).Range.Delete
End If
Next delIdx
textCount = 0
scanIdx = 1
For Each pScan In c.Range.Paragraphs
If Len(CleanParaText(pScan.Range.Text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
End If
Dim sigIdx As Integer: sigIdx = textIndices(textCount)
Dim prevTitleIdx As Integer: prevTitleIdx = textIndices(textCount - 1)
Dim curEmpty As Integer: curEmpty = sigIdx - prevTitleIdx - 1
If curEmpty < 4 Then
Dim nAdd As Integer
For nAdd = 1 To (4 - curEmpty)
Dim rIns As Range
Set rIns = doc.Range(c.Range.Paragraphs(sigIdx).Range.Start, c.Range.Paragraphs(sigIdx).Range.Start)
rIns.Text = vbCr
Next nAdd
ElseIf curEmpty > 4 Then
Dim nSub As Integer
For nSub = 1 To (curEmpty - 4)
c.Range.Paragraphs(prevTitleIdx + 1).Range.Delete
Next nSub
End If
ElseIf textCount >= 1 And Not lastLineHasSignerName Then
If textCount >= 2 Then
Dim delMid As Integer
For delMid = textIndices(textCount) - 1 To textIndices(1) + 1 Step -1
If Len(CleanParaText(c.Range.Paragraphs(delMid).Range.Text)) = 0 Then
c.Range.Paragraphs(delMid).Range.Delete
End If
Next delMid
textCount = 0
scanIdx = 1
For Each pScan In c.Range.Paragraphs
If Len(CleanParaText(pScan.Range.Text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
End If
Dim trailingEmpty As Integer
trailingEmpty = c.Range.Paragraphs.Count - textIndices(textCount)
If trailingEmpty < 4 Then
Dim nTrail As Integer
For nTrail = 1 To (4 - trailingEmpty)
Call SafeAppendBlankParagraphInCell(doc, c)
Next nTrail
End If
End If
On Error Resume Next
Dim pSig As Paragraph
Dim nonEmptyOrder As Integer: nonEmptyOrder = 0
Dim sawAuthLine As Boolean: sawAuthLine = False
For Each pSig In c.Range.Paragraphs
On Error Resume Next
pSig.Borders.Enable = False
pSig.Shading.BackgroundPatternColor = wdColorAutomatic
pSig.Shading.Texture = wdTextureNone
On Error Resume Next
With pSig.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pSig.Range.Font.Name = "Times New Roman"
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
pSig.Range.Font.ColorIndex = wdBlack
pSig.Range.Font.Underline = wdUnderlineNone
Dim curText As String
curText = CleanParaText(pSig.Range.Text)
If Len(curText) > 0 Then
nonEmptyOrder = nonEmptyOrder + 1
Dim uText As String
uText = Trim(ToAscii(curText))
If uText Like "T/M *" Or uText Like "TM.*" Or uText Like "TM *" Or _
uText Like "K/T *" Or uText Like "KT.*" Or uText Like "KT *" Or _
uText Like "T/L *" Or uText Like "TL.*" Or uText Like "TL *" Or _
uText Like "TUQ.*" Or uText Like "T/UQ *" Or uText Like "Q. *" Then
sawAuthLine = True
Dim fixAuth As String
fixAuth = ChuanHoa_ThamQuyen(curText)
If fixAuth <> curText And Len(fixAuth) > 0 Then
Call SafeSetParagraphText(doc, pSig, fixAuth)
End If
pSig.Format.Alignment = wdAlignParagraphCenter
pSig.Range.Font.Name = "Times New Roman"
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
ElseIf Not IsSignatureRoleOrAuth(uText) And nonEmptyOrder = textCount Then
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
ElseIf sawAuthLine Then
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = TT_B("ROLE_UNDER_AUTH_BOLD")
pSig.Range.Font.Italic = False
Else
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
End If
Else
pSig.Range.Font.Bold = False
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
End If
Next pSig
End If
NextFooterCell:
fCellIdx = fCellIdx + 1
Next c
End If
Next tbl
Call TT_DebugLog("BUOC 4 bat dau")
Dim p As Paragraph
Dim s As String
Dim titleState As Integer: titleState = 0
Dim nonTableTextParaCount As Integer: nonTableTextParaCount = 0
Dim inSalutationState As Boolean: inSalutationState = False
Dim outsideSawSo As Boolean: outsideSawSo = False
Dim bodySawDecisionVerb As Boolean: bodySawDecisionVerb = False
Dim afterTitleSep As Boolean: afterTitleSep = False
Dim pLastSummary As Paragraph: Set pLastSummary = Nothing
Dim sawTitleSep As Boolean: sawTitleSep = False
Dim expectSectionTitle As Boolean: expectSectionTitle = False
For Each p In doc.Paragraphs
If p.Range.Information(wdWithInTable) Then
GoTo NextParagraphMain
End If
Call SafeTrimLeadingSpaces(doc, p)
s = CleanParaText(p.Range.Text)
On Error Resume Next
p.Borders.Enable = False
p.Borders(wdBorderTop).LineStyle = wdLineStyleNone
p.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
p.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
p.Borders(wdBorderRight).LineStyle = wdLineStyleNone
p.Shading.BackgroundPatternColor = wdColorAutomatic
p.Shading.Texture = wdTextureNone
On Error Resume Next
If Len(s) = 0 Then
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Format.FirstLineIndent = 0
p.Format.SpaceBefore = 0
p.Format.SpaceAfter = 0
p.Format.LineSpacingRule = wdLineSpaceSingle
GoTo NextParagraphMain
End If
If InStr(1, p.Range.Text, vbTab) > 0 Or InStr(1, p.Range.Text, "  ") > 0 Then
Call SafeCollapseSpacesInParagraph(doc, p)
Call SafeTrimLeadingSpaces(doc, p)
s = CleanParaText(p.Range.Text)
End If
nonTableTextParaCount = nonTableTextParaCount + 1
p.Range.Font.Name = "Times New Roman"
p.Range.Font.ColorIndex = wdBlack
Dim sAscPara As String
sAscPara = Trim(ToAscii(s))
If IsDocNumberLineAsc(sAscPara) And nonTableTextParaCount <= 4 And titleState = 0 Then
outsideSawSo = True
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("DOC_NUMBER_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = False
GoTo NextParagraphMain
ElseIf (outsideSawSo Or sAscPara Like "V/V*" Or sAscPara Like "V.V*") And nonTableTextParaCount <= 5 And Not IsKinhGuiStart(s, sAscPara) And titleState = 0 Then
outsideSawSo = False
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 2
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
p.Range.Font.ColorIndex = wdBlack
p.Range.Font.Underline = wdUnderlineNone
titleState = 6
GoTo NextParagraphMain
End If
If titleState = 0 Then
If nonTableTextParaCount <= 4 And (sAscPara = "DU THAO" Or sAscPara = "(DU THAO)") Then
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
End With
p.Range.Font.Size = 14
p.Range.Font.Bold = True
p.Range.Font.Italic = True
GoTo NextParagraphMain
ElseIf nonTableTextParaCount <= 4 And IsPartyDocTitle(sAscPara) Then
titleState = 1
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 14
.SpaceAfter = 4
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("TITLE_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
GoTo NextParagraphMain
Else
titleState = 6
End If
ElseIf titleState >= 1 And titleState <= 5 Then
If IsSeparatorLine(s) Then
titleState = 6
afterTitleSep = True
sawTitleSep = True
Call FormatCenteredSeparator(p)
GoTo NextParagraphMain
ElseIf IsKinhGuiStart(s, sAscPara) Then
titleState = 6
ElseIf titleState >= 2 And IsAuthorityLine(p, s, sAscPara) Then
titleState = 6
Call FormatAuthorityLine(p)
GoTo NextParagraphMain
ElseIf sAscPara Like "(KEM THEO*" Or sAscPara Like "(BAN HANH KEM THEO*" Or sAscPara Like "KEM THEO*" Or sAscPara Like "BAN HANH KEM THEO*" Then
titleState = titleState + 1
Call FormatCenteredSubNote(p)
GoTo NextParagraphMain
ElseIf HasSeparatorWithinNextParas(p, 4) And Not IsStrictBodyOrHeadingStart(s, sAscPara) Then
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
GoTo NextParagraphMain
Else
titleState = 6
End If
End If
If afterTitleSep Then
afterTitleSep = False
If UCase(s) = s And Len(s) <= 150 And sAscPara Like "*[A-Z][A-Z]*" And _
Not IsStrictBodyOrHeadingStart(s, sAscPara) And Not IsKinhGuiStart(s, sAscPara) Then
Call FormatAuthorityLine(p)
GoTo NextParagraphMain
End If
End If
If expectSectionTitle Then
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
If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Then
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
p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Then bodySawDecisionVerb = True
GoTo NextParagraphMain
ElseIf IsSeparatorLine(s) Then
Call FormatCenteredSeparator(p)
GoTo NextParagraphMain
End If
If IsKinhGuiStart(s, sAscPara) Then
Call SafeTrimLeadingSpaces(doc, p)
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Underline = wdUnderlineNone
p.Range.Font.ColorIndex = wdBlack
Dim cPosColon As Long: cPosColon = InStr(1, p.Range.Text, ":")
If Not TT_B("KINH_GUI_ITALIC") Then
p.Range.Font.Italic = False
ElseIf cPosColon > 0 Then
p.Range.Font.Italic = False
doc.Range(p.Range.Start, p.Range.Start + cPosColon).Font.Italic = True
Else
p.Range.Font.Italic = True
End If
Dim hasNextRecip As Boolean: hasNextRecip = False
Dim pNextCheck As Paragraph: Set pNextCheck = p.Next
Do While Not pNextCheck Is Nothing
If pNextCheck.Range.Information(wdWithInTable) Then Exit Do
Dim sNextC As String: sNextC = CleanParaText(pNextCheck.Range.Text)
If Len(sNextC) > 0 Then
Dim sAscNextC As String: sAscNextC = Trim(ToAscii(sNextC))
If IsKinhGuiRecipientItem(sNextC, sAscNextC) Then
hasNextRecip = True
End If
Exit Do
End If
Set pNextCheck = pNextCheck.Next
Loop
If Not hasNextRecip Then
p.Format.SpaceAfter = 6
inSalutationState = False
Call NormalizeLineEndingPunctuation(doc, p, True)
Else
inSalutationState = True
If cPosColon > 0 And cPosColon < Len(CleanParaText(p.Range.Text)) Then
Call NormalizeLineEndingPunctuation(doc, p, False)
End If
End If
GoTo NextParagraphMain
ElseIf inSalutationState And IsKinhGuiRecipientItem(s, sAscPara) Then
Call SafeTrimLeadingSpaces(doc, p)
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
p.Range.Font.ColorIndex = wdBlack
Dim nextIsRecipient As Boolean: nextIsRecipient = False
Dim pCheckNext As Paragraph: Set pCheckNext = p.Next
Do While Not pCheckNext Is Nothing
If pCheckNext.Range.Information(wdWithInTable) Then Exit Do
Dim sNextChk As String: sNextChk = CleanParaText(pCheckNext.Range.Text)
If Len(sNextChk) > 0 Then
Dim sAscNextChk As String: sAscNextChk = Trim(ToAscii(sNextChk))
If IsKinhGuiRecipientItem(sNextChk, sAscNextChk) Then
nextIsRecipient = True
End If
Exit Do
End If
Set pCheckNext = pCheckNext.Next
Loop
If Not nextIsRecipient Then
p.Format.SpaceAfter = 6
inSalutationState = False
Call NormalizeLineEndingPunctuation(doc, p, True)
Else
Call NormalizeLineEndingPunctuation(doc, p, False)
End If
GoTo NextParagraphMain
Else
inSalutationState = False
End If
If sAscPara Like "NOI NHAN:*" Then
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("NN_LABEL_SIZE")
p.Range.Font.Bold = TT_B("NN_LABEL_BOLD")
p.Range.Font.Italic = TT_B("NN_LABEL_ITALIC")
If TT_B("NN_LABEL_UNDERLINE") Then
p.Range.Font.Underline = wdUnderlineSingle
Else
p.Range.Font.Underline = wdUnderlineNone
End If
GoTo NextParagraphMain
End If
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.ColorIndex = wdBlack
With p.Format
.Alignment = wdAlignParagraphJustify
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = CentimetersToPoints(TT_N("INDENT_CM"))
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
.WidowControl = True
End With
If TT_B("LEGAL_BASIS_ITALIC") And Not bodySawDecisionVerb And (sAscPara Like "CAN CU *" Or sAscPara Like "XET DE NGHI*" Or sAscPara Like "THEO DE NGHI*") Then
p.Range.Font.Italic = True
p.Range.Font.Bold = False
GoTo NextParagraphMain
End If
If sAscPara Like "I.*" Or sAscPara Like "II.*" Or sAscPara Like "III.*" Or sAscPara Like "IV.*" Or sAscPara Like "V.*" Or sAscPara Like "VI.*" Or sAscPara Like "VII.*" Or sAscPara Like "VIII.*" Or sAscPara Like "IX.*" Or sAscPara Like "X.*" Or _
sAscPara Like "I-*" Or sAscPara Like "II-*" Or sAscPara Like "III-*" Or sAscPara Like "IV-*" Or sAscPara Like "V-*" Then
p.Range.Case = wdUpperCase
p.Range.Font.Bold = True
ElseIf Not mTTQPPL And (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And Len(s) < 250 And Right(s, 1) <> "." And Right(s, 1) <> ";" And Right(s, 1) <> ":" And Right(s, 1) <> "," Then
p.Range.Font.Bold = True
ElseIf (sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU ###.*") And Len(s) <= 250 And Right(s, 1) <> "." And Right(s, 1) <> ";" And Right(s, 1) <> ":" Then
p.Range.Font.Bold = True
ElseIf sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU #:*" Or sAscPara Like "DIEU ##:*" Then
p.Range.Font.Bold = False
Dim colonOrDotPos As Long
colonOrDotPos = InStr(1, s, ".")
If colonOrDotPos = 0 Or colonOrDotPos > 10 Then colonOrDotPos = InStr(1, s, ":")
If colonOrDotPos > 0 And colonOrDotPos <= 10 Then
Dim rDieu As Range
Set rDieu = doc.Range(p.Range.Start, p.Range.Start + colonOrDotPos)
rDieu.Font.Bold = True
End If
ElseIf p.Range.Font.Bold = True Then
p.Range.Font.Bold = False
End If
NextParagraphMain:
Next p
If Not (pLastSummary Is Nothing) And Not sawTitleSep Then Call TT_AddSummaryRule(doc, pLastSummary)
Call TT_DebugLog("BUOC 5 bat dau")
Dim tblAll As Table
For Each tblAll In doc.Tables
On Error Resume Next
If tblAll.Rows.LeftIndent >= -MillimetersToPoints(3) Then
tblAll.AutoFitBehavior 2
tblAll.AllowAutoFit = True
End If
tblAll.Range.Font.Name = "Times New Roman"
tblAll.Range.Font.ColorIndex = wdBlack
On Error Resume Next
Next tblAll
Dim pGap As Paragraph
For Each pGap In doc.Paragraphs
If Not pGap.Range.Information(wdWithInTable) Then
If Len(CleanParaText(pGap.Range.Text)) = 0 And pGap.Borders(wdBorderTop).LineStyle = wdLineStyleNone Then
pGap.Range.Font.Name = "Times New Roman"
pGap.Range.Font.Size = 12
pGap.Format.SpaceBefore = 0
pGap.Format.SpaceAfter = 0
pGap.Format.FirstLineIndent = 0
pGap.Format.LineSpacingRule = wdLineSpaceSingle
End If
End If
Next pGap
Call TT_XuLyBangNoiDung(doc)
Call TT_GiuChuKyCungTrang(doc)
Call TT_EpTrang(doc)
Call TT_DebugLog("BUOC 6 bat dau")
Dim strWarnings As String
strWarnings = KiemTra_LogicTheThuc(doc)
On Error Resume Next
Application.UndoRecord.EndCustomRecord
Application.ScreenUpdating = True
On Error GoTo 0
Call TT_DebugLog("KET THUC")
If silentMode Then
On Error Resume Next
doc.Variables("TheThuc_BaoCao").Delete
doc.Variables.Add "TheThuc_BaoCao", TT_ReportHeader() & vbCrLf & "Canh bao:" & strWarnings & vbCrLf & "Tu dong:" & mTTNotes
On Error GoTo 0
End If
If Not silentMode Then
If Len(strWarnings) > 0 Then
MsgBox "DA CHUAN HOA THE THUC VAN BAN!" & vbCrLf & TT_ReportHeader() & TT_NotesText() & vbCrLf & vbCrLf & _
"Phat hien cac diem bat hop ly ve LOGIC can kiem tra truoc khi trinh ky:" & vbCrLf & _
strWarnings & vbCrLf & vbCrLf & "Muon tra lai nhu cu: bam Ctrl+Z mot lan." & vbCrLf & vbCrLf & _
"Duy tri: Vuong Dinh Anh Tu - dua tren add-in cua VPDU xa Thanh Phu", vbExclamation, "Bao Cao Ra Soat Logic & The Thuc"
Else
MsgBox "Da chuan hoa the thuc van ban!" & vbCrLf & TT_ReportHeader() & TT_NotesText() & vbCrLf & vbCrLf & _
"Van ban dat chuan, khong phat hien loi logic." & vbCrLf & "Muon tra lai nhu cu: bam Ctrl+Z mot lan." & vbCrLf & vbCrLf & _
"Duy tri: Vuong Dinh Anh Tu - dua tren add-in cua VPDU xa Thanh Phu", vbInformation, "Chuan Hoa The Thuc Thanh Cong"
End If
End If
Exit Sub
CleanUpAndExit:
On Error Resume Next
Application.ScreenUpdating = True
On Error GoTo 0
End Sub
Private Sub SafeReplaceSoftBreaksInRange(ByVal rng As Range)
On Error Resume Next
If InStr(1, rng.Text, Chr(11)) = 0 Then Exit Sub
Dim pIdx As Long
For pIdx = rng.Paragraphs.Count To 1 Step -1
Dim p As Paragraph
Set p = rng.Paragraphs(pIdx)
If InStr(1, p.Range.Text, Chr(11)) > 0 Then
Dim cIdx As Long
For cIdx = p.Range.Characters.Count To 1 Step -1
If p.Range.Characters(cIdx).Text = Chr(11) Then
p.Range.Characters(cIdx).Text = vbCr
End If
Next cIdx
End If
Next pIdx
On Error GoTo 0
End Sub
Private Sub SafeCollapseSpacesInParagraph(ByVal doc As Document, ByVal p As Paragraph)
On Error Resume Next
Dim rawTxt As String
rawTxt = p.Range.Text
If InStr(1, rawTxt, "  ") = 0 And InStr(1, rawTxt, vbTab) = 0 Then Exit Sub
Dim ascTxt As String
ascTxt = Trim(ToAscii(rawTxt))
If Left(ascTxt, 2) = "SO" And (InStr(1, ascTxt, "-") > 0 Or InStr(1, ascTxt, "/") > 0) Then Exit Sub
If InStr(1, ascTxt, "NGAY") > 0 And InStr(1, ascTxt, "THANG") > 0 And InStr(1, ascTxt, "NAM") > 0 Then Exit Sub
Dim k As Long
Dim nChars As Long
nChars = p.Range.Characters.Count
For k = nChars - 1 To 1 Step -1
Dim chCur As String
chCur = p.Range.Characters(k).Text
If chCur = vbTab Then
p.Range.Characters(k).Text = " "
chCur = " "
End If
If chCur = " " And k > 1 Then
Dim chPrev As String
chPrev = p.Range.Characters(k - 1).Text
If chPrev = " " Or chPrev = vbTab Then
p.Range.Characters(k).Delete
End If
End If
Next k
On Error GoTo 0
End Sub
Private Function CleanParaText(ByVal rawText As String) As String
Dim s As String
s = Replace(Replace(rawText, vbCr, ""), vbLf, "")
s = Replace(s, Chr(7), "")
s = Replace(s, Chr(11), "")
s = Replace(s, Chr(12), "")
s = Replace(s, ChrW(8203), "")
s = Replace(s, ChrW(65279), "")
s = Replace(s, Chr(160), " ")
s = Replace(s, vbTab, " ")
CleanParaText = Trim(s)
End Function
Private Sub RemoveEmptyParagraphsInCell(ByVal doc As Document, ByVal c As Cell)
On Error Resume Next
Dim idxP As Long
For idxP = c.Range.Paragraphs.Count To 1 Step -1
If c.Range.Paragraphs.Count <= 1 Then Exit For
Dim pCheck As Paragraph
Set pCheck = c.Range.Paragraphs(idxP)
If Len(CleanParaText(pCheck.Range.Text)) = 0 Then
If idxP < c.Range.Paragraphs.Count Then
pCheck.Range.Delete
ElseIf idxP > 1 Then
Dim pPrev As Paragraph
Set pPrev = c.Range.Paragraphs(idxP - 1)
Dim rDelCr As Range
Set rDelCr = doc.Range(pPrev.Range.End - 1, pPrev.Range.End)
If rDelCr.Text = vbCr Then rDelCr.Delete
End If
End If
Next idxP
On Error GoTo 0
End Sub
Private Sub SafeTrimLeadingSpaces(ByVal doc As Document, ByVal p As Paragraph)
On Error Resume Next
Dim rawTxt As String
rawTxt = p.Range.Text
If Len(rawTxt) <= 1 Then Exit Sub
Dim firstCode As Long
firstCode = AscW(Left$(rawTxt, 1))
If firstCode < 0 Then firstCode = firstCode + 65536
If firstCode <> 32 And firstCode <> 9 And firstCode <> 160 Then Exit Sub
Dim guard As Integer: guard = 0
Dim prevCount As Long
Do While p.Range.Characters.Count > 1 And guard < 40
guard = guard + 1
Dim chTxt As String
chTxt = p.Range.Characters(1).Text
If chTxt = " " Or chTxt = vbTab Or chTxt = Chr(160) Then
prevCount = p.Range.Characters.Count
p.Range.Characters(1).Delete
If p.Range.Characters.Count >= prevCount Then Exit Do
Else
Exit Do
End If
Loop
On Error GoTo 0
End Sub
Private Sub RemoveHeaderFloatingShapes(ByVal doc As Document, ByVal hdrTbl As Table)
On Error Resume Next
Dim sIdx As Long
For sIdx = doc.Shapes.Count To 1 Step -1
Dim sItem As Shape
Set sItem = Nothing
Set sItem = doc.Shapes(sIdx)
If Not sItem Is Nothing Then
Dim shouldDel As Boolean
shouldDel = False
If StrComp(sItem.Name, "LineTieuNguDang", vbTextCompare) = 0 Then shouldDel = True
If StrComp(sItem.Name, "LineTieuNguQuocHieu", vbTextCompare) = 0 Then shouldDel = True
If StrComp(sItem.Name, "LineCoQuanBanHanh", vbTextCompare) = 0 Then shouldDel = True
If StrComp(sItem.Name, "LineTrichYeu", vbTextCompare) = 0 Then shouldDel = True
If Not hdrTbl Is Nothing Then
If sItem.Anchor.InRange(hdrTbl.Range) Then shouldDel = True
End If
If sItem.Height < 8 And sItem.Top < 260 Then shouldDel = True
If (sItem.Type = 9 Or sItem.Type = 1) And sItem.Top < 260 And sItem.Height < 15 Then shouldDel = True
If shouldDel Then sItem.Delete
End If
Next sIdx
If Not hdrTbl Is Nothing Then
Dim iShpIdx As Long
For iShpIdx = hdrTbl.Range.InlineShapes.Count To 1 Step -1
hdrTbl.Range.InlineShapes(iShpIdx).Delete
Next iShpIdx
End If
On Error GoTo 0
End Sub
Private Sub SafeSetParagraphText(ByVal doc As Document, ByVal p As Paragraph, ByVal newText As String)
On Error Resume Next
Dim r As Range
Set r = p.Range
Do While r.End > r.Start
Dim code As Long
code = AscW(Right(r.Text, 1))
If code < 0 Then code = code + 65536
If code = 13 Or code = 10 Or code = 7 Then
r.End = r.End - 1
Else
Exit Do
End If
Loop
r.Text = newText
On Error GoTo 0
End Sub
Private Sub SafeAppendBlankParagraphInCell(ByVal doc As Document, ByVal c As Cell)
On Error Resume Next
Dim rEnd As Range
Set rEnd = c.Range
rEnd.End = rEnd.End - 1
rEnd.Collapse Direction:=wdCollapseEnd
rEnd.InsertAfter vbCr
On Error GoTo 0
End Sub
Private Sub AutoCreateHeaderTable(ByVal doc As Document)
On Error Resume Next
If doc.Tables.Count >= 1 Then
Dim firstTbl As Table
Set firstTbl = doc.Tables(1)
Dim t1Asc As String
t1Asc = ToAscii(firstTbl.Range.Text)
Dim parasBeforeTbl1 As Integer: parasBeforeTbl1 = 0
Dim pCheckBefore As Paragraph
For Each pCheckBefore In doc.Paragraphs
If pCheckBefore.Range.Start >= firstTbl.Range.Start Then Exit For
If Len(CleanParaText(pCheckBefore.Range.Text)) > 0 Then
parasBeforeTbl1 = parasBeforeTbl1 + 1
End If
Next pCheckBefore
Dim t1DetectStrong As Boolean
If parasBeforeTbl1 <= 1 And (InStr(1, t1Asc, "DANG CONG SAN") > 0 Or InStr(1, t1Asc, "DANG BO") > 0 Or InStr(1, t1Asc, "DANG UY") > 0 Or InStr(1, t1Asc, "CHI BO") > 0 Or InStr(1, t1Asc, "TINH UY") > 0 Or InStr(1, t1Asc, "THANH UY") > 0 Or _
InStr(1, t1Asc, "CONG HOA XA HOI") > 0 Or (firstTbl.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(t1Asc, t1DetectStrong)) > 0 And InStr(1, t1Asc, "NOI NHAN") = 0)) Then
Dim pAfterH As Paragraph
Set pAfterH = firstTbl.Range.Paragraphs(firstTbl.Range.Paragraphs.Count).Next
Dim scanAfterCount As Integer: scanAfterCount = 0
Do While Not pAfterH Is Nothing And scanAfterCount < 4
If pAfterH.Range.Information(wdWithInTable) Then Exit Do
Dim sAfterH As String
sAfterH = CleanParaText(pAfterH.Range.Text)
Dim pNextH As Paragraph
Set pNextH = pAfterH.Next
If Len(sAfterH) > 0 Then
scanAfterCount = scanAfterCount + 1
Dim sAfterHAsc As String
sAfterHAsc = Trim(ToAscii(sAfterH))
If IsPartyDocTitle(sAfterHAsc) Or IsKinhGuiStart(sAfterH, sAfterHAsc) Or IsGluedTitleLine(sAfterH, sAfterHAsc) Then
Exit Do
End If
If InStr(1, sAfterHAsc, "NGAY") > 0 And InStr(1, sAfterHAsc, "THANG") > 0 And InStr(1, sAfterHAsc, "NAM") > 0 And Len(sAfterH) < 95 And InStr(1, t1Asc, "NGAY") = 0 Then
If firstTbl.Columns.Count >= 2 Then
Call AppendTextToCell(doc, firstTbl.Cell(1, firstTbl.Columns.Count), sAfterH)
pAfterH.Range.Delete
End If
ElseIf IsDocNumberLineAsc(sAfterHAsc) And Len(sAfterH) < 60 And InStr(1, t1Asc, "SO ") = 0 And InStr(1, t1Asc, "SO:") = 0 Then
Call AppendTextToCell(doc, firstTbl.Cell(1, 1), sAfterH)
pAfterH.Range.Delete
End If
End If
Set pAfterH = pNextH
Loop
Exit Sub
End If
End If
Dim p As Paragraph
Dim pHStart As Paragraph: Set pHStart = Nothing
Dim pHEnd As Paragraph: Set pHEnd = Nothing
Dim nonEmptyTopCount As Integer: nonEmptyTopCount = 0
Dim sawMottoOrOrgan As Boolean: sawMottoOrOrgan = False
Dim sawMotto As Boolean: sawMotto = False
Dim sawSoOrDateOrStar As Boolean: sawSoOrDateOrStar = False
For Each p In doc.Paragraphs
If p.Range.Information(wdWithInTable) Then Exit For
Dim sRaw As String
sRaw = CleanParaText(p.Range.Text)
If Len(sRaw) > 0 Then
nonEmptyTopCount = nonEmptyTopCount + 1
If nonEmptyTopCount > 12 Then Exit For
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If IsPartyDocTitle(sAsc) Or IsGluedTitleLine(sRaw, sAsc) Or IsKinhGuiStart(sRaw, sAsc) Then
Exit For
End If
If IsStrictBodyOrHeadingStart(sRaw, sAsc) Then
Exit For
End If
Dim isHdrLine As Boolean: isHdrLine = False
If InStr(1, sAsc, "DANG CONG SAN") > 0 Or InStr(1, sAsc, "DOC LAP - TU DO") > 0 Or InStr(1, sAsc, "CONG HOA XA HOI") > 0 Then
isHdrLine = True
sawMottoOrOrgan = True
sawMotto = True
ElseIf (IsHeaderOrganLineAsc(sAsc) Or IsNationalOrganLine(sRaw, sAsc)) And Len(sRaw) < 110 Then
isHdrLine = True
sawMottoOrOrgan = True
ElseIf IsDocNumberLineAsc(sAsc) And Len(sRaw) < 95 Then
isHdrLine = True
sawSoOrDateOrStar = True
ElseIf InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 And Len(sRaw) < 110 Then
isHdrLine = True
sawSoOrDateOrStar = True
ElseIf (sAsc Like "V/V*" Or sAsc Like "V.V*" Or sAsc Like "VE VIEC*") And sawMottoOrOrgan Then
isHdrLine = True
ElseIf (sAsc = "*" Or sAsc = "***" Or IsSeparatorLine(sRaw)) And sawMottoOrOrgan Then
isHdrLine = True
sawSoOrDateOrStar = True
End If
If isHdrLine Then
If pHStart Is Nothing Then Set pHStart = p
Set pHEnd = p
ElseIf Not pHStart Is Nothing Then
Exit For
End If
End If
Next p
If pHStart Is Nothing Or pHEnd Is Nothing Then Exit Sub
If Not (sawMotto Or (sawMottoOrOrgan And sawSoOrDateOrStar)) Then Exit Sub
Dim leftLines As String: leftLines = ""
Dim rightLines As String: rightLines = ""
Dim pCurH As Paragraph
Set pCurH = pHStart
Do While Not pCurH Is Nothing
Dim rawLine As String
rawLine = Replace(Replace(Replace(pCurH.Range.Text, vbCr, ""), vbLf, ""), Chr(7), "")
rawLine = Replace(rawLine, Chr(11), vbCr)
Dim subLines() As String
subLines = Split(rawLine, vbCr)
Dim slIdx As Long
For slIdx = LBound(subLines) To UBound(subLines)
Call ClassifyAndAppendHeaderLine(subLines(slIdx), leftLines, rightLines)
Next slIdx
If pCurH.Range.Start >= pHEnd.Range.Start Then Exit Do
Set pCurH = pCurH.Next
Loop
If Len(Trim(leftLines)) = 0 And Len(Trim(rightLines)) = 0 Then Exit Sub
Dim rightAsc As String
rightAsc = ToAscii(rightLines)
Dim hasAnyMotto As Boolean
hasAnyMotto = (InStr(1, rightAsc, "DANG CONG SAN") > 0 Or InStr(1, rightAsc, "CONG HOA XA HOI") > 0)
Dim mottoStr As String
mottoStr = ""
If TT_B("AUTO_INSERT_MOTTO") Then
If TT_S("MOTTO_KIND") = "DANG" And Not hasAnyMotto Then
mottoStr = TT_TextTieuNguDang()
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" Then
If Not hasAnyMotto Then
mottoStr = TT_TextQuocHieu() & vbCr & TT_TextTieuNgu()
ElseIf InStr(1, rightAsc, "CONG HOA XA HOI") > 0 And InStr(1, Replace(rightAsc, " ", ""), "DOCLAP-TUDO") = 0 Then
Dim rParts() As String
rParts = Split(rightLines, vbCr)
Dim rIdx As Long
Dim rebuilt As String
rebuilt = ""
For rIdx = LBound(rParts) To UBound(rParts)
Call AppendLineToBuffer(rebuilt, rParts(rIdx))
If InStr(1, ToAscii(rParts(rIdx)), "CONG HOA XA HOI") > 0 Then
Call AppendLineToBuffer(rebuilt, TT_TextTieuNgu())
End If
Next rIdx
rightLines = rebuilt
End If
End If
End If
If Len(mottoStr) > 0 Then
If Len(Trim(rightLines)) > 0 Then
rightLines = mottoStr & vbCr & rightLines
Else
rightLines = mottoStr
End If
End If
Dim rngReplaceH As Range
Set rngReplaceH = doc.Range(pHStart.Range.Start, pHEnd.Range.End)
Dim newHdrTbl As Table
Set newHdrTbl = doc.Tables.Add(Range:=rngReplaceH, NumRows:=1, NumColumns:=2)
With newHdrTbl
.Borders.Enable = False
.Columns(1).Width = MillimetersToPoints(TT_N("HDR_LEFT_MM"))
.Columns(2).Width = MillimetersToPoints(TT_N("HDR_RIGHT_MM"))
.Cell(1, 1).Range.Text = leftLines
.Cell(1, 2).Range.Text = rightLines
End With
On Error GoTo 0
End Sub
Private Sub ClassifyAndAppendHeaderLine(ByVal rawSubLine As String, ByRef leftLines As String, ByRef rightLines As String)
Dim sLine As String
sLine = Trim(Replace(rawSubLine, Chr(160), " "))
If Len(sLine) = 0 Then Exit Sub
If InStr(1, sLine, vbTab) > 0 Then
Dim tabParts() As String
tabParts = Split(sLine, vbTab)
Dim tIdx As Long
For tIdx = LBound(tabParts) To UBound(tabParts)
If Len(Trim(tabParts(tIdx))) > 0 Then
Call ClassifyAndAppendHeaderLine(Trim(tabParts(tIdx)), leftLines, rightLines)
End If
Next tIdx
Exit Sub
End If
Dim sAsc As String
sAsc = Trim(ToAscii(sLine))
If Len(sAsc) = 0 Then Exit Sub
Dim posDCS As Long
posDCS = InStr(1, sAsc, "DANG CONG SAN")
If posDCS = 0 Then posDCS = InStr(1, sAsc, "CONG HOA XA HOI")
If posDCS = 0 Then posDCS = InStr(1, sAsc, "DOC LAP - TU DO")
If posDCS > 3 Then
Dim leftPart1 As String, rightPart1 As String
leftPart1 = Trim(Left(sLine, posDCS - 1))
rightPart1 = Trim(Mid(sLine, posDCS))
If Len(leftPart1) > 0 Then Call AppendLineToBuffer(leftLines, leftPart1)
If Len(rightPart1) > 0 Then Call AppendLineToBuffer(rightLines, rightPart1)
Exit Sub
End If
If InStr(1, sAsc, "DANG CONG SAN") > 0 Or sAsc = "VIET NAM" Or InStr(1, sAsc, "DOC LAP - TU DO") > 0 Or InStr(1, sAsc, "CONG HOA XA HOI") > 0 Then
Call AppendLineToBuffer(rightLines, sLine)
ElseIf InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 And InStr(1, sAsc, "V/V") = 0 Then
Call AppendLineToBuffer(rightLines, sLine)
ElseIf sAsc = "*" Or sAsc = "***" Then
If InStr(1, leftLines, "*") = 0 And Not TT_B("ORGAN_UNDERLINE") Then
Call AppendLineToBuffer(leftLines, "*")
End If
ElseIf IsSeparatorLine(sLine) Then
If InStr(1, ToAscii(leftLines), "V/V") > 0 Or InStr(1, ToAscii(leftLines), "VE VIEC") > 0 Then
Call AppendLineToBuffer(leftLines, sLine)
End If
Else
Call AppendLineToBuffer(leftLines, sLine)
End If
End Sub
Private Sub AppendLineToBuffer(ByRef buf As String, ByVal newLine As String)
Dim cleanL As String
cleanL = Trim(newLine)
If Len(cleanL) = 0 Then Exit Sub
If Len(buf) > 0 Then
buf = buf & vbCr & cleanL
Else
buf = cleanL
End If
End Sub
Private Sub AppendTextToCell(ByVal doc As Document, ByVal targetCell As Cell, ByVal extraLine As String)
On Error Resume Next
Dim existingTxt As String
existingTxt = CleanParaText(targetCell.Range.Text)
If Len(existingTxt) > 0 Then
Call SafeAppendBlankParagraphInCell(doc, targetCell)
Call SafeSetParagraphText(doc, targetCell.Range.Paragraphs(targetCell.Range.Paragraphs.Count), extraLine)
Else
targetCell.Range.Text = extraLine
End If
On Error GoTo 0
End Sub
Private Sub SplitGluedHeaderAndTitleLines(ByVal doc As Document)
On Error Resume Next
If doc.Tables.Count >= 1 Then
Dim tblH As Table
Set tblH = doc.Tables(1)
Dim tHAsc As String
tHAsc = ToAscii(tblH.Range.Text)
If InStr(1, tHAsc, "DANG CONG SAN") > 0 Or InStr(1, tHAsc, "DANG BO") > 0 Or InStr(1, tHAsc, "DANG UY") > 0 Then
Dim cH As Cell
For Each cH In tblH.Range.Cells
Dim pH As Paragraph
For Each pH In cH.Range.Paragraphs
Dim sH As String: sH = CleanParaText(pH.Range.Text)
Dim sHAsc As String: sHAsc = Trim(ToAscii(sH))
If InStr(1, sHAsc, "DANG CONG SAN VIET NAM") > 0 And InStr(1, sHAsc, "NGAY") > 0 And InStr(1, sHAsc, "THANG") > 0 Then
Dim posVN As Long
posVN = InStr(1, sHAsc, "VIET NAM")
If posVN > 0 Then
Dim cutVN As Long: cutVN = posVN + 7
Dim leftMotto As String, rightDate As String
leftMotto = Trim(Left(sH, cutVN))
rightDate = Trim(Mid(sH, cutVN + 1))
Do While Len(rightDate) > 0 And (Left(rightDate, 1) = "*" Or Left(rightDate, 1) = "-" Or Left(rightDate, 1) = " ")
rightDate = Trim(Mid(rightDate, 2))
Loop
If Len(rightDate) > 0 Then
Call SafeSetParagraphText(doc, pH, leftMotto & vbCr & rightDate)
End If
End If
ElseIf sHAsc Like "DANG BO *" Or sHAsc Like "TINH UY *" Or sHAsc Like "THANH UY *" Or sHAsc Like "HUYEN UY *" Or sHAsc Like "DANG UY *" Then
Dim splitPos As Long: splitPos = 0
If InStr(12, sHAsc, " BAN THUONG VU") > 0 Then
splitPos = InStr(12, sHAsc, " BAN THUONG VU")
ElseIf InStr(12, sHAsc, " BAN CHAP HANH") > 0 Then
splitPos = InStr(12, sHAsc, " BAN CHAP HANH")
ElseIf InStr(12, sHAsc, " BAN XAY DUNG") > 0 Then
splitPos = InStr(12, sHAsc, " BAN XAY DUNG")
ElseIf InStr(12, sHAsc, " BAN TUYEN GIAO") > 0 Then
splitPos = InStr(12, sHAsc, " BAN TUYEN GIAO")
ElseIf InStr(12, sHAsc, " BAN TO CHUC") > 0 Then
splitPos = InStr(12, sHAsc, " BAN TO CHUC")
ElseIf InStr(12, sHAsc, " VAN PHONG") > 0 Then
splitPos = InStr(12, sHAsc, " VAN PHONG")
ElseIf InStr(12, sHAsc, " UY BAN KIEM TRA") > 0 Then
splitPos = InStr(12, sHAsc, " UY BAN KIEM TRA")
ElseIf InStr(12, sHAsc, " DANG UY ") > 0 And Not (sHAsc Like "BAN CHAP HANH*") Then
splitPos = InStr(12, sHAsc, " DANG UY ")
ElseIf InStr(12, sHAsc, " CHI BO ") > 0 Then
splitPos = InStr(12, sHAsc, " CHI BO ")
End If
If splitPos > 10 Then
Dim org1 As String, org2 As String
org1 = Trim(Left(sH, splitPos - 1))
org2 = Trim(Mid(sH, splitPos + 1))
If Len(org1) > 0 And Len(org2) > 0 Then
Call SafeSetParagraphText(doc, pH, org1 & vbCr & org2)
End If
End If
End If
Next pH
Next cH
End If
End If
Dim pBody As Paragraph
Dim nonTblChecked As Integer: nonTblChecked = 0
For Each pBody In doc.Paragraphs
If Not pBody.Range.Information(wdWithInTable) Then
Dim sRawB As String
sRawB = CleanParaText(pBody.Range.Text)
If Len(sRawB) > 0 Then
nonTblChecked = nonTblChecked + 1
If nonTblChecked > 4 Then Exit For
Dim sAscB As String
sAscB = Trim(ToAscii(sRawB))
Dim titleLen As Integer
titleLen = GetGluedTitlePrefixLen(sRawB, sAscB)
If titleLen > 0 Then
Dim partTitle As String, partSummary As String
partTitle = Trim(Left(sRawB, titleLen))
partSummary = Trim(Mid(sRawB, titleLen + 1))
If Left(partSummary, 1) = ":" Or Left(partSummary, 1) = "-" Then
partSummary = Trim(Mid(partSummary, 2))
End If
If Len(partTitle) > 0 And Len(partSummary) > 0 Then
Call SafeSetParagraphText(doc, pBody, partTitle & vbCr & partSummary)
Exit For
End If
End If
End If
End If
Next pBody
On Error GoTo 0
End Sub
Private Function IsGluedTitleLine(ByVal sRaw As String, ByVal sAsc As String) As Boolean
IsGluedTitleLine = (GetGluedTitlePrefixLen(sRaw, sAsc) > 0)
End Function
Private Function GetGluedTitlePrefixLen(ByVal sRaw As String, ByVal sAsc As String) As Integer
GetGluedTitlePrefixLen = 0
If IsPartyDocTitle(sAsc) Then Exit Function
Dim titles(1 To 16) As String
titles(1) = "CHUONG TRINH HANH DONG"
titles(2) = "BAO CAO CHINH TRI"
titles(3) = "DE CUONG BAO CAO"
titles(4) = "CHUONG TRINH"
titles(5) = "NGHI QUYET"
titles(6) = "QUYET DINH"
titles(7) = "HUONG DAN"
titles(8) = "THONG BAO"
titles(9) = "THONG TRI"
titles(10) = "KE HOACH"
titles(11) = "KET LUAN"
titles(12) = "QUY DINH"
titles(13) = "TO TRINH"
titles(14) = "BIEN BAN"
titles(15) = "BAO CAO"
titles(16) = "QUY CHE"
Dim idx As Integer
For idx = 1 To 16
Dim kw As String: kw = titles(idx)
Dim kwLen As Integer: kwLen = Len(kw)
If Len(sAsc) > kwLen + 3 Then
If Left(sAsc, kwLen) = kw Then
Dim rawPrefix As String
rawPrefix = Left(sRaw, kwLen)
Dim rawRest As String
rawRest = Trim(Mid(sRaw, kwLen + 1))
If rawPrefix = UCase(rawPrefix) And Len(rawRest) > 2 And rawRest <> UCase(rawRest) Then
GetGluedTitlePrefixLen = kwLen
Exit Function
End If
End If
End If
Next idx
End Function
Private Function InsertGapIfPrecededByTable(ByVal doc As Document, ByVal pFirst As Paragraph) As Boolean
InsertGapIfPrecededByTable = False
On Error GoTo GapFail
Dim pPrev As Paragraph
Set pPrev = pFirst.Previous
If pPrev Is Nothing Then Exit Function
If Not pPrev.Range.Information(wdWithInTable) Then Exit Function
Dim posStart As Long
posStart = pFirst.Range.Start
doc.Range(posStart, posStart).InsertBefore vbCr
InsertGapIfPrecededByTable = True
Exit Function
GapFail:
InsertGapIfPrecededByTable = False
End Function
Private Sub AutoCreateKinhGuiTable(ByVal doc As Document)
On Error Resume Next
Dim prefixText As String
prefixText = "K" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.Text)
If IsKinhGuiTable(tbl, tAsc) Then
Dim kgCellLeft As Cell: Set kgCellLeft = Nothing
Dim kgCellRight As Cell: Set kgCellRight = Nothing
Dim cScanKG As Cell
For Each cScanKG In tbl.Range.Cells
Dim cScanAsc As String
cScanAsc = Trim(ToAscii(cScanKG.Range.Text))
If HasKinhGuiOrTrinh(cScanAsc) Then
Set kgCellLeft = cScanKG
ElseIf Not kgCellLeft Is Nothing And kgCellRight Is Nothing Then
Set kgCellRight = cScanKG
Exit For
End If
Next cScanKG
Dim pAfterTbl As Paragraph
Dim pAbsorbStart As Paragraph: Set pAbsorbStart = Nothing
Dim pAbsorbEnd As Paragraph: Set pAbsorbEnd = Nothing
Dim lastPInTbl As Paragraph
Set lastPInTbl = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count)
Set pAfterTbl = lastPInTbl.Next
Do While Not pAfterTbl Is Nothing
If pAfterTbl.Range.Information(wdWithInTable) Then Exit Do
Dim sAfterRaw As String
sAfterRaw = CleanParaText(pAfterTbl.Range.Text)
If Len(sAfterRaw) > 0 Then
Dim sAfterAsc As String
sAfterAsc = Trim(ToAscii(sAfterRaw))
If IsKinhGuiRecipientItem(sAfterRaw, sAfterAsc) Then
If pAbsorbStart Is Nothing Then Set pAbsorbStart = pAfterTbl
Set pAbsorbEnd = pAfterTbl
Else
Exit Do
End If
End If
Set pAfterTbl = pAfterTbl.Next
Loop
If Not pAbsorbStart Is Nothing And Not pAbsorbEnd Is Nothing Then
Dim targetCell As Cell
Set targetCell = kgCellRight
If targetCell Is Nothing Then Set targetCell = kgCellLeft
If Not targetCell Is Nothing Then
Dim allRecipText As String: allRecipText = ""
Dim pInC As Paragraph
For Each pInC In targetCell.Range.Paragraphs
Dim sCLine As String
sCLine = CleanParaText(pInC.Range.Text)
If Len(sCLine) > 0 Then
If Len(allRecipText) > 0 Then
allRecipText = allRecipText & vbCr & sCLine
Else
allRecipText = sCLine
End If
End If
Next pInC
Dim pCurAbsorb As Paragraph
Set pCurAbsorb = pAbsorbStart
Do While Not pCurAbsorb Is Nothing
Dim sAbsLine As String
sAbsLine = CleanParaText(pCurAbsorb.Range.Text)
If Len(sAbsLine) > 0 Then
If Len(allRecipText) > 0 Then
allRecipText = allRecipText & vbCr & sAbsLine
Else
allRecipText = sAbsLine
End If
End If
If pCurAbsorb.Range.Start >= pAbsorbEnd.Range.Start Then Exit Do
Set pCurAbsorb = pCurAbsorb.Next
Loop
doc.Range(pAbsorbStart.Range.Start, pAbsorbEnd.Range.End).Delete
targetCell.Range.Text = allRecipText
targetCell.Range.Font.Bold = False
targetCell.Range.Font.Italic = False
End If
End If
If Not kgCellLeft Is Nothing Then
Dim leftCellTxt As String
leftCellTxt = ToAscii(kgCellLeft.Range.Text)
Call SafeSetParagraphText(doc, kgCellLeft.Range.Paragraphs(1), KinhGuiPrefixFor(leftCellTxt))
End If
Exit Sub
End If
Next tbl
Dim p As Paragraph
Dim pStart As Paragraph: Set pStart = Nothing
Dim pEnd As Paragraph: Set pEnd = Nothing
Dim foundStart As Boolean: foundStart = False
Dim scanCount As Integer: scanCount = 0
For Each p In doc.Paragraphs
If Not p.Range.Information(wdWithInTable) Then
Dim sRaw As String
sRaw = CleanParaText(p.Range.Text)
If Len(sRaw) > 0 Then
scanCount = scanCount + 1
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If Not foundStart Then
If IsKinhGuiStart(sRaw, sAsc) Then
Set pStart = p
Set pEnd = p
foundStart = True
ElseIf scanCount > 25 Then
Exit Sub
End If
Else
If IsKinhGuiRecipientItem(sRaw, sAsc) Then
Set pEnd = p
Else
Exit For
End If
End If
End If
End If
Next p
If Not foundStart Or pStart Is Nothing Or pEnd Is Nothing Then
Exit Sub
End If
Dim recipientsText As String: recipientsText = ""
Dim currP As Paragraph
Set currP = pStart
Do While Not currP Is Nothing
Dim lineText As String
lineText = CleanParaText(currP.Range.Text)
If Len(lineText) > 0 Then
Dim lineAsc As String
lineAsc = Trim(ToAscii(lineText))
If currP.Range.Start = pStart.Range.Start Then
prefixText = KinhGuiPrefixFor(lineAsc)
Dim cPos As Long
cPos = InStr(1, lineText, ":")
If cPos > 0 And cPos < Len(lineText) Then
Dim afterColon As String
afterColon = Trim(Mid(lineText, cPos + 1))
If Len(afterColon) > 0 Then
recipientsText = afterColon
End If
ElseIf cPos = 0 And (lineAsc Like "KINH GUI *" Or lineAsc Like "DONG KINH GUI *") Then
Dim firstSpacePos As Long
If lineAsc Like "DONG KINH GUI *" Then
firstSpacePos = InStr(15, lineText, " ")
If firstSpacePos = 0 Then firstSpacePos = 14
Else
firstSpacePos = InStr(9, lineText, " ")
If firstSpacePos = 0 Then firstSpacePos = 9
End If
Dim remText As String
remText = Trim(Mid(lineText, firstSpacePos + 1))
If Len(remText) > 0 Then
recipientsText = remText
End If
End If
Else
If Len(recipientsText) > 0 Then
recipientsText = recipientsText & vbCr & lineText
Else
recipientsText = lineText
End If
End If
End If
If currP.Range.Start >= pEnd.Range.Start Then Exit Do
Set currP = currP.Next
Loop
If Len(Trim(recipientsText)) = 0 Then
Exit Sub
End If
Dim kgStartPos As Long, kgEndPos As Long
kgStartPos = pStart.Range.Start
kgEndPos = pEnd.Range.End
If InsertGapIfPrecededByTable(doc, pStart) Then
kgStartPos = kgStartPos + 1
kgEndPos = kgEndPos + 1
End If
Dim rngReplace As Range
Set rngReplace = doc.Range(kgStartPos, kgEndPos)
Dim newTbl As Table
Set newTbl = doc.Tables.Add(Range:=rngReplace, NumRows:=1, NumColumns:=2)
With newTbl
.Borders.Enable = False
On Error Resume Next
.Rows(1).LeftIndent = 0
.Rows(1).Alignment = wdAlignRowLeft
.Columns(1).Width = MillimetersToPoints(52)
.Columns(2).Width = MillimetersToPoints(117)
On Error GoTo 0
.Cell(1, 1).Range.Text = prefixText
.Cell(1, 1).Range.Font.Bold = False
.Cell(1, 1).Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
.Cell(1, 2).Range.Text = recipientsText
.Cell(1, 2).Range.Font.Bold = False
.Cell(1, 2).Range.Font.Italic = False
End With
On Error GoTo 0
End Sub
Private Sub AutoCreateFooterTable(ByVal doc As Document)
On Error Resume Next
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.Text)
If IsFooterSignatureTable(tbl, tAsc) Then Exit Sub
Next tbl
Dim totalP As Long: totalP = doc.Paragraphs.Count
If totalP < 3 Then Exit Sub
Dim startScanIdx As Long
startScanIdx = totalP - 30
If startScanIdx < 2 Then startScanIdx = 2
Dim pFooterStart As Paragraph: Set pFooterStart = Nothing
Dim pFooterEnd As Paragraph: Set pFooterEnd = Nothing
Dim sawNoiNhan As Boolean: sawNoiNhan = False
Dim sawSigAuthOrRole As Boolean: sawSigAuthOrRole = False
Dim idx As Long
For idx = startScanIdx To totalP
Dim pCur As Paragraph
Set pCur = doc.Paragraphs(idx)
If pCur.Range.Information(wdWithInTable) Then
If Not pFooterStart Is Nothing Then Exit For
GoTo NextFooterScan
End If
Dim sRaw As String
sRaw = CleanParaText(pCur.Range.Text)
If Len(sRaw) = 0 Then GoTo NextFooterScan
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If pFooterStart Is Nothing Then
If sAsc Like "NOI NHAN*" Or (IsSignatureRoleOrAuth(sAsc) And Not IsFooterBreakerLine(sRaw, sAsc)) Or HasTabSeparatedFooterLine(pCur.Range.Text) Then
Set pFooterStart = pCur
Set pFooterEnd = pCur
If InStr(1, sAsc, "NOI NHAN") > 0 Then sawNoiNhan = True
If IsSignatureRoleOrAuth(sAsc) Or HasTabSeparatedFooterLine(pCur.Range.Text) Then sawSigAuthOrRole = True
End If
Else
If IsPartyDocTitle(sAsc) Or sAsc Like "PHU LUC*" Then Exit For
If Not sawNoiNhan And IsFooterBreakerLine(sRaw, sAsc) Then
Set pFooterStart = Nothing
Set pFooterEnd = Nothing
sawSigAuthOrRole = False
GoTo NextFooterScan
End If
Set pFooterEnd = pCur
If InStr(1, sAsc, "NOI NHAN") > 0 Then sawNoiNhan = True
If IsSignatureRoleOrAuth(sAsc) Then sawSigAuthOrRole = True
End If
NextFooterScan:
Next idx
If pFooterStart Is Nothing Or pFooterEnd Is Nothing Then Exit Sub
If Not (sawNoiNhan Or sawSigAuthOrRole) Then Exit Sub
Dim leftFooter As String: leftFooter = ""
Dim rightFooter As String: rightFooter = ""
Dim inSigBlockSeq As Boolean: inSigBlockSeq = False
Dim pIter As Paragraph
Set pIter = pFooterStart
Do While Not pIter Is Nothing
Dim rawF As String
rawF = Replace(Replace(Replace(pIter.Range.Text, vbCr, ""), vbLf, ""), Chr(7), "")
rawF = Replace(rawF, Chr(11), vbCr)
Dim fSubLines() As String
fSubLines = Split(rawF, vbCr)
Dim fIdx As Long
For fIdx = LBound(fSubLines) To UBound(fSubLines)
Call ClassifyAndAppendFooterLine(fSubLines(fIdx), leftFooter, rightFooter, inSigBlockSeq)
Next fIdx
If pIter.Range.Start >= pFooterEnd.Range.Start Then Exit Do
Set pIter = pIter.Next
Loop
If Len(Trim(leftFooter)) = 0 And Len(Trim(rightFooter)) = 0 Then Exit Sub
Dim ftStartPos As Long, ftEndPos As Long
ftStartPos = pFooterStart.Range.Start
ftEndPos = pFooterEnd.Range.End
If InsertGapIfPrecededByTable(doc, pFooterStart) Then
ftStartPos = ftStartPos + 1
ftEndPos = ftEndPos + 1
End If
Dim rngFooterReplace As Range
Set rngFooterReplace = doc.Range(ftStartPos, ftEndPos)
Dim newFooterTbl As Table
Set newFooterTbl = doc.Tables.Add(Range:=rngFooterReplace, NumRows:=1, NumColumns:=2)
With newFooterTbl
.Borders.Enable = False
.Columns(1).Width = MillimetersToPoints(80)
.Columns(2).Width = MillimetersToPoints(85)
.Cell(1, 1).Range.Text = leftFooter
.Cell(1, 2).Range.Text = rightFooter
End With
On Error GoTo 0
End Sub
Private Function HasTabSeparatedFooterLine(ByVal rawParaText As String) As Boolean
HasTabSeparatedFooterLine = False
If InStr(1, rawParaText, vbTab) = 0 Then Exit Function
Dim sAsc As String
sAsc = Trim(ToAscii(CleanParaText(rawParaText)))
If InStr(1, sAsc, "NOI NHAN") > 0 Or InStr(1, sAsc, "T/M ") > 0 Or InStr(1, sAsc, "TM.") > 0 Or _
InStr(1, sAsc, "K/T ") > 0 Or InStr(1, sAsc, "KT.") > 0 Or InStr(1, sAsc, "BI THU") > 0 Or _
InStr(1, sAsc, "TRUONG BAN") > 0 Or InStr(1, sAsc, "CHANH VAN PHONG") > 0 Or _
InStr(1, sAsc, "TUQ.") > 0 Or InStr(1, sAsc, "TL.") > 0 Or InStr(1, sAsc, "T/L ") > 0 Or _
InStr(1, sAsc, "CHU TICH") > 0 Or InStr(1, sAsc, "GIAM DOC") > 0 Then
HasTabSeparatedFooterLine = True
End If
End Function
Private Sub ClassifyAndAppendFooterLine(ByVal rawLine As String, ByRef leftFooter As String, ByRef rightFooter As String, ByRef inSigBlockSeq As Boolean)
Dim sCleanCheck As String
sCleanCheck = CleanParaText(rawLine)
If Len(sCleanCheck) = 0 Then Exit Sub
If InStr(1, rawLine, vbTab) > 0 Then
Dim firstTab As Long
firstTab = InStr(1, rawLine, vbTab)
Dim lPart As String, rPart As String
lPart = CleanParaText(Left(rawLine, firstTab - 1))
rPart = CleanParaText(Mid(rawLine, firstTab + 1))
If Len(lPart) > 0 Then
Dim lAsc As String: lAsc = Trim(ToAscii(lPart))
If IsSignatureRoleOrAuth(lAsc) Then
Call AppendLineToBuffer(rightFooter, lPart)
inSigBlockSeq = True
Else
Call AppendLineToBuffer(leftFooter, lPart)
End If
End If
If Len(rPart) > 0 Then
Call AppendLineToBuffer(rightFooter, rPart)
inSigBlockSeq = True
End If
Exit Sub
End If
Dim sAsc As String
sAsc = Trim(ToAscii(sCleanCheck))
If sAsc Like "NOI NHAN*" Or Left(sCleanCheck, 1) = "-" Or Left(sCleanCheck, 1) = ChrW(8211) Or _
sAsc Like "LUU:*" Or sAsc Like "LUU VT*" Or sAsc Like "NHU TREN*" Or sAsc Like "NHU DIEU*" Then
Call AppendLineToBuffer(leftFooter, sCleanCheck)
ElseIf IsSignatureRoleOrAuth(sAsc) Then
Call AppendLineToBuffer(rightFooter, sCleanCheck)
inSigBlockSeq = True
ElseIf inSigBlockSeq And Len(sCleanCheck) < 65 Then
Call AppendLineToBuffer(rightFooter, sCleanCheck)
Else
Call AppendLineToBuffer(leftFooter, sCleanCheck)
End If
End Sub
Private Function IsDocNumberLineAsc(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If Len(t) = 0 Or Len(t) > 35 Then
IsDocNumberLineAsc = False
Exit Function
End If
If (t Like "SO *" Or t Like "SO:*" Or t Like "SO-*") And _
(InStr(1, t, "/") > 0 Or t Like "SO #*" Or t Like "SO:*#*" Or t Like "SO -*") Then
IsDocNumberLineAsc = True
ElseIf Len(t) <= 20 And InStr(1, t, " ") = 0 And _
(InStr(1, t, "-CV/") > 0 Or InStr(1, t, "-BC/") > 0 Or InStr(1, t, "-KH/") > 0 Or _
InStr(1, t, "-QD/") > 0 Or InStr(1, t, "-NQ/") > 0 Or InStr(1, t, "-TTR/") > 0 Or _
InStr(1, t, "-TB/") > 0 Or InStr(1, t, "-KL/") > 0 Or InStr(1, t, "-CT/") > 0 Or _
InStr(1, t, "-HD/") > 0 Or InStr(1, t, "-QC/") > 0 Or InStr(1, t, "-QYD/") > 0 Or _
InStr(1, t, "-CTR/") > 0) Then
IsDocNumberLineAsc = True
Else
IsDocNumberLineAsc = False
End If
End Function
Private Function IsHeaderOrganLineAsc(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If t Like "DANG BO *" Or t Like "DANG UY *" Or t Like "TINH UY *" Or _
t Like "THANH UY *" Or t Like "HUYEN UY *" Or t Like "THI UY *" Or _
t Like "QUAN UY *" Or t Like "CHI BO *" Or t Like "BAN CHAP HANH*" Or _
t Like "BAN THUONG VU*" Or t Like "BAN XAY DUNG*" Or t Like "BAN TUYEN GIAO*" Or _
t Like "BAN TO CHUC*" Or t Like "BAN DAN VAN*" Or t Like "VAN PHONG*" Or _
t Like "UY BAN KIEM TRA*" Or t Like "TIEU BAN*" Or t Like "HOI DONG*" Or _
t Like "UY BAN NHAN DAN*" Or t Like "UBND *" Or t Like "HDND *" Or _
t Like "BO CONG AN*" Or t Like "CONG AN *" Or t Like "BO QUOC PHONG*" Or _
t Like "BO CHI HUY*" Or t Like "BAN CHI HUY*" Or t Like "BO TU LENH*" Or t Like "QUAN KHU*" Or _
t Like "UY BAN MAT TRAN*" Or t Like "MAT TRAN TO QUOC*" Or t Like "BAN THUONG TRUC*" Then
IsHeaderOrganLineAsc = True
Else
IsHeaderOrganLineAsc = False
End If
End Function
Private Function IsNationalOrganLine(ByVal sRaw As String, ByVal sAsc As String) As Boolean
IsNationalOrganLine = False
If Not TT_IsNational() Then Exit Function
Dim t As String
t = Trim(sRaw)
If Len(t) < 3 Or Len(t) > 90 Then Exit Function
If UCase(t) <> t Then Exit Function
If Not (sAsc Like "*[A-Z][A-Z]*") Then Exit Function
If IsDocNumberLineAsc(sAsc) Or IsSeparatorLine(t) Then Exit Function
If InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 Then Exit Function
If IsPartyDocTitle(sAsc) Or IsKinhGuiStart(sRaw, sAsc) Then Exit Function
IsNationalOrganLine = True
End Function
Private Function IsKinhGuiTable(ByVal tbl As Table, ByVal tAscKG As String) As Boolean
IsKinhGuiTable = False
If Not HasKinhGuiOrTrinh(tAscKG) Or InStr(1, tAscKG, "NOI NHAN") > 0 Then Exit Function
If tbl.Columns.Count > 3 Or tbl.Rows.Count > 5 Then Exit Function
Dim cCheck As Cell
For Each cCheck In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(CleanParaText(cCheck.Range.Text)))
If cAsc Like "KINH GUI*" Or cAsc Like "DONG KINH GUI*" Or IsKinhTrinhAsc(cAsc) Then
IsKinhGuiTable = True
Exit Function
End If
Next cCheck
End Function
Private Function IsFooterSignatureTable(ByVal tbl As Table, ByVal tAscFooter As String) As Boolean
IsFooterSignatureTable = False
If InStr(1, tAscFooter, "DANG CONG SAN") > 0 Then Exit Function
If InStr(1, tAscFooter, "CONG HOA XA HOI") > 0 Then Exit Function
If tbl.Rows.Count > 4 Or tbl.Columns.Count > 4 Then Exit Function
Dim cCheck As Cell
For Each cCheck In tbl.Range.Cells
Dim pCheck As Paragraph
For Each pCheck In cCheck.Range.Paragraphs
Dim pAsc As String
pAsc = Trim(ToAscii(CleanParaText(pCheck.Range.Text)))
If pAsc Like "NOI NHAN*" Or IsSignatureRoleOrAuth(pAsc) Then
If InStr(1, tAscFooter, "NOI NHAN") = 0 And tbl.Rows.Count >= 3 Then
If tbl.Borders.Enable <> 0 Then Exit Function
End If
IsFooterSignatureTable = True
Exit Function
End If
Next pCheck
Next cCheck
End Function
Private Function HasSignatureKeywordInCell(ByVal cAsc As String) As Boolean
HasSignatureKeywordInCell = False
If InStr(1, cAsc, "T/M ") > 0 Or InStr(1, cAsc, "TM.") > 0 Or _
InStr(1, cAsc, "K/T ") > 0 Or InStr(1, cAsc, "KT.") > 0 Or _
InStr(1, cAsc, "T/L ") > 0 Or InStr(1, cAsc, "TL.") > 0 Or _
InStr(1, cAsc, "BI THU") > 0 Or InStr(1, cAsc, "CHANH VAN PHONG") > 0 Or _
InStr(1, cAsc, "TRUONG BAN") > 0 Or InStr(1, cAsc, "CHU NHIEM") > 0 Or _
InStr(1, cAsc, "CHU TICH") > 0 Or InStr(1, cAsc, "NGUOI LAP") > 0 Or _
InStr(1, cAsc, "TUQ.") > 0 Or InStr(1, cAsc, "GIAM DOC") > 0 Or _
InStr(1, cAsc, "TRUONG PHONG") > 0 Or InStr(1, cAsc, "THU TRUONG") > 0 Or _
InStr(1, cAsc, "TU LENH") > 0 Or InStr(1, cAsc, "CHINH UY") > 0 Or _
InStr(1, cAsc, "CHI HUY TRUONG") > 0 Or InStr(1, cAsc, "BO TRUONG") > 0 Or _
InStr(1, cAsc, "THU KY") > 0 Or InStr(1, cAsc, "CHU TOA") > 0 Or InStr(1, cAsc, "CHU TRI") > 0 Then
HasSignatureKeywordInCell = True
End If
End Function
Private Function IsSignatureRoleOrAuth(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If Len(t) = 0 Or Len(t) > 65 Then
IsSignatureRoleOrAuth = False
Exit Function
End If
If t Like "T/M *" Or t Like "TM.*" Or t Like "TM *" Or _
t Like "K/T *" Or t Like "KT.*" Or t Like "KT *" Or _
t Like "T/L *" Or t Like "TL.*" Or t Like "TL *" Or _
t = "BI THU" Or t Like "PHO BI THU*" Or t Like "QUYEN BI THU*" Or _
t Like "CHANH VAN PHONG*" Or t Like "PHO CHANH VAN PHONG*" Or _
t Like "TRUONG BAN*" Or t Like "PHO TRUONG BAN*" Or _
t Like "CHU NHIEM*" Or t Like "PHO CHU NHIEM*" Or _
t Like "UY VIEN*" Or t = "TONG BI THU" Or _
t Like "CHU TICH*" Or t Like "PHO CHU TICH*" Or _
t Like "THU TRUONG*" Or t Like "NGUOI LAP*" Or t Like "NGUOI BAO CAO*" Or t Like "XAC NHAN*" Or _
t Like "TUQ.*" Or t Like "T/UQ *" Or t Like "Q. *" Or _
t Like "GIAM DOC*" Or t Like "PHO GIAM DOC*" Or t Like "TRUONG PHONG*" Or t Like "PHO TRUONG PHONG*" Or _
t Like "BO TRUONG*" Or t Like "CUC TRUONG*" Or t Like "PHO CUC TRUONG*" Or _
t Like "TU LENH*" Or t Like "PHO TU LENH*" Or t Like "CHINH UY*" Or t Like "PHO CHINH UY*" Or _
t Like "CHI HUY TRUONG*" Or t Like "CHI HUY PHO*" Or t Like "TRUONG CONG AN*" Or t Like "PHO TRUONG CONG AN*" Or _
t Like "HIEU TRUONG*" Or t Like "PHO HIEU TRUONG*" Or t Like "VIEN TRUONG*" Or t Like "PHO VIEN TRUONG*" Or _
t = "THU KY" Or t = "CHU TOA" Or t = "CHU TRI" Then
IsSignatureRoleOrAuth = True
Else
IsSignatureRoleOrAuth = False
End If
End Function
Private Function IsKinhTrinhAsc(ByVal sAsc As String) As Boolean
Dim t As String
t = Trim(sAsc)
IsKinhTrinhAsc = (t Like "KINH TRINH:*" Or t Like "KINH TRINH :*")
End Function
Private Function HasKinhGuiOrTrinh(ByVal sAsc As String) As Boolean
HasKinhGuiOrTrinh = (InStr(1, sAsc, "KINH GUI") > 0 Or InStr(1, sAsc, "KINH TRINH:") > 0 Or InStr(1, sAsc, "KINH TRINH :") > 0)
End Function
Private Function KinhGuiPrefixFor(ByVal sAsc As String) As String
If InStr(1, sAsc, "DONG KINH GUI") > 0 Then
KinhGuiPrefixFor = ChrW(272) & ChrW(7891) & "ng k" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
ElseIf InStr(1, sAsc, "KINH TRINH") > 0 Then
KinhGuiPrefixFor = "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh:"
Else
KinhGuiPrefixFor = "K" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
End If
End Function
Private Function IsKinhGuiStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tAsc As String: tAsc = Trim(sAsc)
If tAsc Like "KINH GUI:*" Or tAsc Like "KINH GUI :*" Or _
tAsc Like "DONG KINH GUI:*" Or tAsc Like "DONG KINH GUI :*" Or _
tAsc Like "KINH GUI *" Or tAsc Like "DONG KINH GUI *" Or _
tAsc = "KINH GUI" Or tAsc = "DONG KINH GUI" Or IsKinhTrinhAsc(tAsc) Then
IsKinhGuiStart = True
Else
IsKinhGuiStart = False
End If
End Function
Private Function IsKinhGuiRecipientItem(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tRaw As String: tRaw = CleanParaText(rawText)
Dim tAsc As String: tAsc = Trim(ToAscii(tRaw))
If Len(tRaw) = 0 Or Len(tRaw) > 300 Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If IsSeparatorLine(tRaw) Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If tAsc Like "CAN CU*" Or tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
tAsc Like "QUYET DINH:*" Or tAsc Like "QUYET DINH :*" Or _
tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Or _
tAsc Like "DE *" Or tAsc Like "NHAM *" Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
tAsc Like "1.*" Or tAsc Like "2.*" Or tAsc Like "3.*" Or _
tAsc Like "DIEU #*" Or tAsc Like "DIEU ##*" Or _
IsPartyDocTitle(tAsc) Then
IsKinhGuiRecipientItem = False
Exit Function
End If
Dim firstCh As String: firstCh = Left(tRaw, 1)
If firstCh = "-" Or firstCh = ChrW(8211) Or firstCh = ChrW(8212) Or _
firstCh = ChrW(8210) Or firstCh = ChrW(8722) Or firstCh = ChrW(8226) Or _
firstCh = ChrW(8259) Or firstCh = ChrW(183) Or firstCh = "+" Or firstCh = "*" Then
IsKinhGuiRecipientItem = True
Exit Function
End If
Dim lastCh As String: lastCh = Right(tRaw, 1)
If lastCh = ";" Or lastCh = "," Then
IsKinhGuiRecipientItem = True
Exit Function
End If
If lastCh = "." And Len(tRaw) < 180 Then
If tAsc Like "*DANG UY*" Or tAsc Like "*CHI BO*" Or tAsc Like "*DANG BO*" Or _
tAsc Like "*BAN *" Or tAsc Like "*UBND*" Or tAsc Like "*HDND*" Or _
tAsc Like "*MAT TRAN*" Or tAsc Like "*MTTQ*" Or tAsc Like "*DOAN *" Or _
tAsc Like "*HOI *" Or tAsc Like "*PHONG *" Or tAsc Like "*SO *" Or _
tAsc Like "*XA *" Or tAsc Like "*HUYEN *" Or tAsc Like "*TINH *" Or _
tAsc Like "*TRUONG *" Or tAsc Like "*DON VI*" Or tAsc Like "*CO QUAN*" Or _
tAsc Like "*TRUNG TAM*" Or tAsc Like "*TRUYEN THANH*" Or tAsc Like "*VAN HOA*" Then
IsKinhGuiRecipientItem = True
Exit Function
End If
End If
IsKinhGuiRecipientItem = False
End Function
Private Sub NormalizeLineEndingPunctuation(ByVal doc As Document, ByVal p As Paragraph, ByVal isLastLine As Boolean)
On Error Resume Next
Dim rngLine As Range
Set rngLine = p.Range
Do While rngLine.End > rngLine.Start
Dim cCode As Long
cCode = AscW(Right(rngLine.Text, 1))
If cCode < 0 Then cCode = cCode + 65536
If cCode = 13 Or cCode = 10 Or cCode = 7 Then
rngLine.End = rngLine.End - 1
Else
Exit Do
End If
Loop
Dim rngEndPos As Long
rngEndPos = rngLine.End
Do While rngLine.End > rngLine.Start
Dim endChCode As Long
endChCode = AscW(Right(rngLine.Text, 1))
If endChCode < 0 Then endChCode = endChCode + 65536
If endChCode = 32 Or endChCode = 160 Or endChCode = 9 Or endChCode = 46 Or endChCode = 59 Or endChCode = 44 Then
rngLine.End = rngLine.End - 1
Else
Exit Do
End If
Loop
Dim rngTail As Range
Set rngTail = doc.Range(rngLine.End, rngEndPos)
If isLastLine Then
rngTail.Text = "."
Else
rngTail.Text = ";"
End If
On Error GoTo 0
End Sub
Private Sub FormatCenteredSummary(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub FormatCenteredSubNote(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Italic = True
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub FormatCenteredSeparator(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 14
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Function HasSeparatorWithinNextParas(ByVal pCur As Paragraph, ByVal maxLookahead As Integer) As Boolean
HasSeparatorWithinNextParas = False
On Error Resume Next
Dim pNext As Paragraph
Set pNext = pCur.Next
Dim stepCount As Integer: stepCount = 0
Do While Not pNext Is Nothing And stepCount < maxLookahead
If pNext.Range.Information(wdWithInTable) Then Exit Do
Dim sNext As String
sNext = CleanParaText(pNext.Range.Text)
If Len(sNext) > 0 Then
stepCount = stepCount + 1
If IsSeparatorLine(sNext) Then
HasSeparatorWithinNextParas = True
Exit Function
End If
Dim sNextAsc As String
sNextAsc = Trim(ToAscii(sNext))
If IsStrictBodyOrHeadingStart(sNext, sNextAsc) Or IsKinhGuiStart(sNext, sNextAsc) Then
Exit Function
End If
End If
Set pNext = pNext.Next
Loop
On Error GoTo 0
End Function
Private Function TT_CanonDocCode(ByVal code As String, ByVal isDang As Boolean) As String
Dim u As String
u = UCase(code)
TT_CanonDocCode = u
If u = "TTR" Then
TT_CanonDocCode = "TTr"
ElseIf u = "CTR" Then
TT_CanonDocCode = "CTr"
ElseIf Not isDang Then
If u = "QY" & ChrW(272) Then
TT_CanonDocCode = "Qy" & ChrW(272)
ElseIf u = "TRS" Then
TT_CanonDocCode = "TrS"
End If
End If
End Function
Private Function ChuanHoa_SoKyHieu(ByVal sRaw As String) As String
If InStr(1, sRaw, "[") > 0 Then
ChuanHoa_SoKyHieu = Trim(sRaw)
Exit Function
End If
Select Case TT_S("DOC_NUMBER_STYLE")
Case "NHA_NUOC"
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_NhaNuoc(sRaw, False)
Case "LINH_HOAT"
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_LinhHoat(sRaw)
Case Else
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_Dang(sRaw)
End Select
End Function
Private Function DocNumberStyleOf(ByVal sRaw As String) As String
DocNumberStyleOf = ""
Dim rest As String
rest = Trim(sRaw)
If UCase(Left(ToAscii(rest), 2)) <> "SO" Then Exit Function
rest = Trim(Mid(rest, 3))
If Left(rest, 1) = ":" Then rest = Trim(Mid(rest, 2))
rest = Replace(Replace(Replace(rest, " ", ""), ChrW(8211), "-"), ChrW(8212), "-")
Dim k As Long
For k = 1 To Len(rest)
Dim ch As String
ch = Mid(rest, k, 1)
If ch = "-" Then
If InStr(k + 1, rest, "/") > 0 Then DocNumberStyleOf = "DANG"
Exit Function
ElseIf ch = "/" Then
DocNumberStyleOf = "NHA_NUOC"
Exit Function
ElseIf Not (ch Like "#" Or ch = "." Or ch = ChrW(8230)) Then
Exit Function
End If
Next k
End Function
Private Function ChuanHoa_SoKyHieu_LinhHoat(ByVal sRaw As String) As String
Select Case DocNumberStyleOf(sRaw)
Case "DANG": ChuanHoa_SoKyHieu_LinhHoat = ChuanHoa_SoKyHieu_Dang(sRaw)
Case "NHA_NUOC": ChuanHoa_SoKyHieu_LinhHoat = ChuanHoa_SoKyHieu_NhaNuoc(sRaw, True)
Case Else: ChuanHoa_SoKyHieu_LinhHoat = Trim(sRaw)
End Select
End Function
Private Function ChuanHoa_SoKyHieu_NhaNuoc(ByVal sRaw As String, ByVal keepPrefix As Boolean) As String
Dim s As String
s = Trim(sRaw)
ChuanHoa_SoKyHieu_NhaNuoc = s
If UCase(Left(ToAscii(s), 2)) <> "SO" Then Exit Function
Dim rest As String
rest = Trim(Mid(s, 3))
Dim hadColon As Boolean
hadColon = (Left(rest, 1) = ":")
If hadColon Then rest = Mid(rest, 2)
rest = Replace(Trim(rest), " ", "")
rest = Replace(rest, Chr(160), "")
rest = Replace(rest, ChrW(8211), "-")
rest = Replace(rest, ChrW(8212), "-")
If Len(rest) = 0 Then Exit Function
Dim k As Long
Dim numPart As String
k = 1
Do While k <= Len(rest)
Dim ch As String
ch = Mid(rest, k, 1)
If ch Like "#" Or ch = "." Or ch = ChrW(8230) Then
numPart = numPart & ch
k = k + 1
Else
Exit Do
End If
Loop
Dim tail As String
tail = Mid(rest, k)
If Len(tail) < 2 Then Exit Function
Dim sep As String
sep = Left(tail, 1)
If sep <> "/" And sep <> "-" Then Exit Function
tail = Mid(tail, 2)
If sep = "-" Then
Dim sl As Long
sl = InStr(1, tail, "/")
If sl <= 1 Then Exit Function
If InStr(1, Left(tail, sl - 1), "-") = 0 Then
tail = Left(tail, sl - 1) & "-" & Mid(tail, sl + 1)
End If
End If
tail = UCase(tail)
Dim dashCode As Long
dashCode = InStr(1, tail, "-")
If dashCode > 1 Then
tail = TT_CanonDocCode(Left(tail, dashCode - 1), False) & Mid(tail, dashCode)
Else
tail = TT_CanonDocCode(tail, False)
End If
Dim cleanNum As String
cleanNum = Trim(Replace(Replace(numPart, ".", ""), ChrW(8230), ""))
Dim prefix As String
If keepPrefix And Not hadColon Then
prefix = "S" & ChrW(7889) & " "
Else
prefix = "S" & ChrW(7889) & ": "
End If
If Len(cleanNum) = 0 Or Not IsNumeric(cleanNum) Then
ChuanHoa_SoKyHieu_NhaNuoc = RTrim(prefix) & "   /" & tail
Else
If Len(cleanNum) = 1 And Val(cleanNum) > 0 Then cleanNum = "0" & cleanNum
ChuanHoa_SoKyHieu_NhaNuoc = prefix & cleanNum & "/" & tail
End If
End Function
Private Function ChuanHoa_SoKyHieu_Dang(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
Dim sAsc As String
sAsc = ToAscii(s)
If Left(sAsc, 3) = "SO:" Then
s = "S" & ChrW(7889) & " " & Trim(Mid(s, 4))
End If
Dim p1 As Long
p1 = InStr(1, s, " ")
If p1 > 0 Then
Dim prefix As String, rest As String
prefix = Left(s, p1)
rest = Replace(Trim(Mid(s, p1 + 1)), " ", "")
rest = Replace(rest, ChrW(8211), "-")
rest = Replace(rest, ChrW(8212), "-")
Dim dashPos As Long, slashPos As Long
dashPos = InStr(1, rest, "-")
slashPos = InStr(1, rest, "/")
Dim numPart As String, codePart As String, orgPart As String
If slashPos > 0 And dashPos > slashPos Then
numPart = Trim(Left(rest, slashPos - 1))
codePart = Trim(Mid(rest, slashPos + 1, dashPos - slashPos - 1))
orgPart = Trim(Mid(rest, dashPos + 1))
Dim cleanNum As String
cleanNum = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum = Trim(cleanNum)
If Len(cleanNum) = 0 Or Not IsNumeric(cleanNum) Then
s = "S" & ChrW(7889) & "   -" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
Else
If Val(cleanNum) < 10 And Val(cleanNum) > 0 And Len(cleanNum) = 1 Then
cleanNum = "0" & cleanNum
End If
s = "S" & ChrW(7889) & " " & cleanNum & "-" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
End If
ElseIf dashPos > 0 And slashPos > dashPos Then
numPart = Trim(Left(rest, dashPos - 1))
codePart = Trim(Mid(rest, dashPos + 1, slashPos - dashPos - 1))
orgPart = Trim(Mid(rest, slashPos + 1))
Dim cleanNum2 As String
cleanNum2 = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum2 = Trim(cleanNum2)
If Len(cleanNum2) = 0 Or Not IsNumeric(cleanNum2) Then
s = "S" & ChrW(7889) & "   -" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
Else
If Val(cleanNum2) < 10 And Val(cleanNum2) > 0 And Len(cleanNum2) = 1 Then
cleanNum2 = "0" & cleanNum2
End If
s = "S" & ChrW(7889) & " " & cleanNum2 & "-" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
End If
ElseIf dashPos > 0 Then
numPart = Trim(Left(rest, dashPos - 1))
Dim cleanNum3 As String
cleanNum3 = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum3 = Trim(cleanNum3)
If Len(cleanNum3) = 0 Or Not IsNumeric(cleanNum3) Then
s = "S" & ChrW(7889) & "   -" & Mid(rest, dashPos + 1)
Else
If Val(cleanNum3) < 10 And Val(cleanNum3) > 0 And Len(cleanNum3) = 1 Then
s = "S" & ChrW(7889) & " 0" & rest
Else
s = "S" & ChrW(7889) & " " & rest
End If
End If
End If
End If
ChuanHoa_SoKyHieu_Dang = s
End Function
Private Function ChuanHoa_NgayThang(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
Dim sAsc As String
sAsc = ToAscii(s)
Dim pNgay As Long, pThang As Long, pNam As Long
pNgay = InStr(1, sAsc, "NGAY")
pThang = InStr(1, sAsc, "THANG")
pNam = InStr(1, sAsc, "NAM")
If pNgay > 0 And pThang > pNgay And pNam > pThang Then
Dim locPart As String, dayPart As String, monthPart As String, yearPart As String
locPart = Left(s, pNgay - 1)
dayPart = Trim(Mid(s, pNgay + 4, pThang - (pNgay + 4)))
monthPart = Trim(Mid(s, pThang + 5, pNam - (pThang + 5)))
yearPart = Trim(Mid(s, pNam + 3))
Dim cleanDay As String
cleanDay = Replace(Replace(dayPart, ".", ""), ChrW(8230), "")
cleanDay = Trim(cleanDay)
Dim dayText As String
If Len(cleanDay) = 0 Or Not IsNumeric(cleanDay) Then
dayText = "   "
Else
Dim dVal As Long: dVal = Val(cleanDay)
If dVal < 10 And dVal > 0 And Len(cleanDay) = 1 Then
cleanDay = "0" & cleanDay
End If
dayText = " " & cleanDay & " "
End If
Dim cleanMonth As String
cleanMonth = Replace(Replace(monthPart, ".", ""), ChrW(8230), "")
cleanMonth = Trim(cleanMonth)
Dim monthText As String
If Len(cleanMonth) = 0 Or Not IsNumeric(cleanMonth) Then
monthText = "   "
Else
Dim mVal As Long: mVal = Val(cleanMonth)
If (mVal = 1 Or mVal = 2) And Len(cleanMonth) = 1 Then
cleanMonth = "0" & cleanMonth
ElseIf mVal >= 3 And mVal <= 9 And Len(cleanMonth) > 1 And Left(cleanMonth, 1) = "0" Then
cleanMonth = CStr(mVal)
End If
monthText = " " & cleanMonth & " "
End If
Dim strNgay As String, strThang As String, strNam As String
strNgay = "ng" & ChrW(224) & "y"
strThang = "th" & ChrW(225) & "ng"
strNam = "n" & ChrW(259) & "m"
s = locPart & strNgay & dayText & strThang & monthText & strNam & " " & yearPart
End If
ChuanHoa_NgayThang = s
End Function
Private Function ChuanHoa_ThamQuyen(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
ChuanHoa_ThamQuyen = s
Dim style As String
style = TT_S("SIGN_PREFIX_STYLE")
If style = "GIU_NGUYEN" Then Exit Function
Dim sUp As String
sUp = ToAscii(s)
Dim kind As String, restText As String
kind = ""
If Left(sUp, 4) = "T/UQ" Then
kind = "TUQ": restText = Mid(s, 5)
ElseIf Left(sUp, 4) = "TUQ." Or Left(sUp, 4) = "TUQ " Then
kind = "TUQ": restText = Mid(s, 5)
ElseIf Left(sUp, 3) = "T/M" Then
kind = "TM": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "TM." Or Left(sUp, 3) = "TM " Then
kind = "TM": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "K/T" Then
kind = "KT": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "KT." Or Left(sUp, 3) = "KT " Then
kind = "KT": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "T/L" Then
kind = "TL": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "TL." Or Left(sUp, 3) = "TL " Then
kind = "TL": restText = Mid(s, 4)
End If
If Len(kind) = 0 Then Exit Function
restText = Trim(restText)
If Len(restText) = 0 Then Exit Function
If style = "DOT" Then
ChuanHoa_ThamQuyen = kind & ". " & restText
Else
Select Case kind
Case "TM": ChuanHoa_ThamQuyen = "T/M " & restText
Case "KT": ChuanHoa_ThamQuyen = "K/T " & restText
Case "TL": ChuanHoa_ThamQuyen = "T/L " & restText
Case Else: ChuanHoa_ThamQuyen = s
End Select
End If
End Function
Private Function KiemTra_LogicTheThuc(ByVal doc As Document) As String
Dim warnings As String: warnings = ""
Dim warnCount As Integer: warnCount = 0
Dim docTextFull As String
docTextFull = doc.Content.Text
Dim docAscFull As String
docAscFull = ToAscii(docTextFull)
Dim mismatchMsg As String
mismatchMsg = TT_MismatchWarning()
If Len(mismatchMsg) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". " & mismatchMsg
End If
If Len(mTTHeaderNote) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". " & mTTHeaderNote
End If
If InStr(1, docTextFull, "...") > 0 Or InStr(1, docTextFull, ChrW(8230)) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Du lieu trong] Phat hien con dau '...' chua dien day du thong tin."
End If
Dim phCount As Long: phCount = 0
Dim phPos As Long: phPos = InStr(1, docTextFull, "[")
Do While phPos > 0 And phCount < 500
If InStr(phPos + 1, docTextFull, "]") = 0 Then Exit Do
phCount = phCount + 1
phPos = InStr(phPos + 1, docTextFull, "[")
Loop
If phCount > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Cho trong] Con " & phCount & " cho can dien trong dau [ ] (vi du: [Ho va ten])."
End If
Dim foundSoRaw As String: foundSoRaw = ""
Dim foundNumVal As String: foundNumVal = ""
Dim foundCode As String: foundCode = ""
Dim foundDateRaw As String: foundDateRaw = ""
Dim foundTitle As String: foundTitle = ""
Dim foundSuperiorOrgan As String: foundSuperiorOrgan = ""
Dim foundIssuingOrgan As String: foundIssuingOrgan = ""
Dim hasVvSummary As Boolean: hasVvSummary = False
Dim hasKinhGui As Boolean: hasKinhGui = False
If InStr(1, docAscFull, "V/V") > 0 Or InStr(1, docAscFull, "VE VIEC") > 0 Then hasVvSummary = True
If InStr(1, docAscFull, "KINH GUI") > 0 Then hasKinhGui = True
Dim hasQuyetDinhVerb As Boolean: hasQuyetDinhVerb = False
Dim hasNoiNhan As Boolean: hasNoiNhan = False
Dim hasLuuVT As Boolean: hasLuuVT = False
Dim hasLuuOther As Boolean: hasLuuOther = False
Dim hasNhuDieu As Boolean: hasNhuDieu = False
Dim foundSignAuth As String: foundSignAuth = ""
Dim foundSignRole As String: foundSignRole = ""
If doc.Tables.Count >= 1 Then
Dim tblH As Table
Set tblH = doc.Tables(1)
Dim tHAsc As String
tHAsc = ToAscii(tblH.Range.Text)
Dim tHDetectStrong As Boolean
If InStr(1, tHAsc, "DANG CONG SAN") > 0 Or InStr(1, tHAsc, "DANG UY") > 0 Or InStr(1, tHAsc, "DANG BO") > 0 Or InStr(1, tHAsc, "CHI BO") > 0 Or _
InStr(1, tHAsc, "CONG HOA XA HOI") > 0 Or (tblH.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(tHAsc, tHDetectStrong)) > 0 And InStr(1, tHAsc, "NOI NHAN") = 0) Then
Dim cH As Cell
Dim orgLinesCount As Integer: orgLinesCount = 0
For Each cH In tblH.Range.Cells
Dim pH As Paragraph
For Each pH In cH.Range.Paragraphs
Dim sH As String: sH = CleanParaText(pH.Range.Text)
Dim sHAsc As String: sHAsc = Trim(ToAscii(sH))
If Len(sHAsc) > 0 Then
If IsDocNumberLineAsc(sHAsc) Then
foundSoRaw = sHAsc
ElseIf InStr(1, sHAsc, "NGAY") > 0 And InStr(1, sHAsc, "THANG") > 0 Then
foundDateRaw = sHAsc
ElseIf sHAsc Like "V/V*" Or sHAsc Like "VE VIEC*" Then
hasVvSummary = True
ElseIf InStr(1, sHAsc, "DANG CONG SAN") = 0 And InStr(1, sHAsc, "CONG HOA XA HOI") = 0 And InStr(1, Replace(sHAsc, " ", ""), "DOCLAP-TUDO") = 0 And _
InStr(1, sHAsc, "KINH GUI") = 0 And Left(sH, 1) <> "-" And Not IsSeparatorLine(sH) And sHAsc <> "*" And sHAsc <> "***" Then
orgLinesCount = orgLinesCount + 1
If orgLinesCount = 1 Then
foundIssuingOrgan = sHAsc
ElseIf orgLinesCount = 2 Then
foundSuperiorOrgan = foundIssuingOrgan
foundIssuingOrgan = sHAsc
End If
End If
End If
Next pH
Next cH
End If
End If
Dim pCount As Long: pCount = doc.Paragraphs.Count
Dim nonTblIdx As Long: nonTblIdx = 0
Dim romanOrder(1 To 10) As String
romanOrder(1) = "I": romanOrder(2) = "II": romanOrder(3) = "III": romanOrder(4) = "IV": romanOrder(5) = "V"
romanOrder(6) = "VI": romanOrder(7) = "VII": romanOrder(8) = "VIII": romanOrder(9) = "IX": romanOrder(10) = "X"
Dim nextExpectedRoman As Integer: nextExpectedRoman = 1
Dim nextExpectedArabic As Integer: nextExpectedArabic = 1
Dim nextExpectedDieu As Integer: nextExpectedDieu = 1
Dim totalDieuCount As Integer: totalDieuCount = 0
Dim idx As Long
For idx = 1 To pCount
Dim pItem As Paragraph
Set pItem = doc.Paragraphs(idx)
Dim pText As String
pText = CleanParaText(pItem.Range.Text)
If Len(pText) = 0 Then GoTo NextParaCheck
Dim pAsc As String
pAsc = Trim(ToAscii(pText))
Dim inTbl As Boolean
inTbl = pItem.Range.Information(wdWithInTable)
If Not inTbl Then nonTblIdx = nonTblIdx + 1
If Not inTbl And nonTblIdx <= 4 And Len(foundTitle) = 0 Then
If IsPartyDocTitle(pAsc) Then
foundTitle = NormalizeDocTitleForCompare(pAsc)
GoTo NextParaCheck
End If
End If
If InStr(1, pAsc, "KINH GUI") > 0 Then hasKinhGui = True
If pAsc Like "V/V*" Or pAsc Like "VE VIEC*" Then hasVvSummary = True
If pAsc = "QUYET DINH:" Or pAsc = "QUYET DINH :" Then hasQuyetDinhVerb = True
If idx >= pCount - 35 Or inTbl Then
If InStr(1, pAsc, "NOI NHAN") > 0 Then hasNoiNhan = True
If InStr(1, pAsc, "LUU VT") > 0 Or InStr(1, pAsc, "LUU: VT") > 0 Or InStr(1, pAsc, "LUU VAN THU") > 0 Or InStr(1, pAsc, "LUU: VAN THU") > 0 Or _
(TT_S("MOTTO_KIND") = "DANG" And InStr(1, pAsc, "LUU VAN PHONG") > 0) Then
hasLuuVT = True
ElseIf InStr(1, pAsc, "LUU ") > 0 Or InStr(1, pAsc, "LUU:") > 0 Then
hasLuuOther = True
End If
If InStr(1, pAsc, "NHU DIEU") > 0 Then hasNhuDieu = True
If pAsc Like "T/M *" Or pAsc Like "TM.*" Or pAsc Like "K/T *" Or pAsc Like "KT.*" Or pAsc Like "T/L *" Or pAsc Like "TL.*" Or pAsc Like "TUQ.*" Then
foundSignAuth = pAsc
ElseIf pAsc = "BI THU" Or pAsc Like "PHO BI THU*" Or pAsc = "CHANH VAN PHONG" Or pAsc = "PHO CHANH VAN PHONG" Or pAsc Like "TRUONG BAN*" Or pAsc Like "CHU NHIEM*" Then
foundSignRole = pAsc
End If
End If
If Not inTbl Then
If pAsc Like "DIEU #.*" Or pAsc Like "DIEU ##.*" Or pAsc Like "DIEU #:*" Then
Dim dieuNumStr As String
dieuNumStr = Trim(Mid(pAsc, 5, 3))
dieuNumStr = Replace(Replace(dieuNumStr, ".", ""), ":", "")
If IsNumeric(dieuNumStr) Then
Dim actualDieu As Integer: actualDieu = Val(dieuNumStr)
totalDieuCount = totalDieuCount + 1
If actualDieu <> nextExpectedDieu Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Thu tu Dieu] Nhay coc hoac sai thu tu: Ky vong 'Dieu " & nextExpectedDieu & "' nhung gap 'Dieu " & actualDieu & "'."
End If
nextExpectedDieu = actualDieu + 1
nextExpectedArabic = 1
End If
End If
Dim rIdx As Integer
Dim matchedRoman As Boolean: matchedRoman = False
For rIdx = 1 To 10
If pAsc Like romanOrder(rIdx) & ". *" Or pAsc Like romanOrder(rIdx) & "- *" Or pAsc Like romanOrder(rIdx) & " - *" Then
matchedRoman = True
If rIdx <> nextExpectedRoman Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [De muc La Ma] Nhay coc thu tu: Ky vong muc '" & romanOrder(nextExpectedRoman) & "' nhung gap muc '" & romanOrder(rIdx) & "'."
End If
nextExpectedRoman = rIdx + 1
nextExpectedArabic = 1
Exit For
End If
Next rIdx
If Not matchedRoman Then
If (pAsc Like "#. *" Or pAsc Like "##. *") And Not (pAsc Like "#.#*") Then
Dim dotPos As Long: dotPos = InStr(1, pAsc, ".")
If dotPos > 1 Then
Dim arabStr As String: arabStr = Left(pAsc, dotPos - 1)
If IsNumeric(arabStr) Then
Dim actualArab As Integer: actualArab = Val(arabStr)
If actualArab <> nextExpectedArabic Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [De muc So] Nhay coc thu tu khoan: Ky vong muc '" & nextExpectedArabic & ".' nhung gap '" & actualArab & ".' (" & Left(pAsc, 35) & "...)."
End If
nextExpectedArabic = actualArab + 1
End If
End If
End If
End If
End If
NextParaCheck:
Next idx
Dim isStandaloneNonDispatch As Boolean
isStandaloneNonDispatch = False
If foundTitle = "PHU LUC" Or foundTitle = "DANH MUC" Or foundTitle = "BANG TONG HOP" Or _
foundTitle = "BIEU TONG HOP" Or foundTitle = "BAI THU HOACH" Or foundTitle = "BAI PHAT BIEU" Or _
foundTitle = "BAN KIEM DIEM" Or foundTitle = "BAN CAM KET" Or foundTitle = "BAN DANG KY" Then
isStandaloneNonDispatch = True
End If
Dim foundSoStyle As String: foundSoStyle = ""
If Len(foundSoRaw) > 0 Then
Dim soBody As String
soBody = Trim(Mid(foundSoRaw, 3))
If Left(soBody, 1) = ":" Then soBody = Trim(Mid(soBody, 2))
soBody = Replace(soBody, " ", "")
Dim dashIdx As Long, slashIdx As Long
dashIdx = InStr(1, soBody, "-")
slashIdx = InStr(1, soBody, "/")
If dashIdx > 0 And slashIdx > dashIdx Then
foundSoStyle = "DANG"
foundNumVal = Left(soBody, dashIdx - 1)
foundCode = Mid(soBody, dashIdx + 1, slashIdx - dashIdx - 1)
ElseIf slashIdx > 0 Then
foundSoStyle = "NHA_NUOC"
foundNumVal = Left(soBody, slashIdx - 1)
Dim afterSlash As String
afterSlash = Mid(soBody, slashIdx + 1)
If afterSlash Like "####/*" Then afterSlash = Mid(afterSlash, 6)
Dim dashAfter As Long
dashAfter = InStr(1, afterSlash, "-")
If dashAfter > 0 Then
foundCode = Left(afterSlash, dashAfter - 1)
Else
foundCode = afterSlash
End If
End If
foundNumVal = Trim(Replace(Replace(foundNumVal, ".", ""), ChrW(8230), ""))
If Len(foundNumVal) = 0 Or Not IsNumeric(foundNumVal) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So van ban] Dang de trong 3 khoang trang ('" & foundSoRaw & "') cho van thu cap so."
End If
If TT_S("DOC_NUMBER_STYLE") = "DANG" And foundSoStyle = "NHA_NUOC" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Nha nuoc ('" & foundSoRaw & "'); van ban Dang ghi dang 'So 12-QD/DU'."
ElseIf TT_S("DOC_NUMBER_STYLE") = "NHA_NUOC" And foundSoStyle = "DANG" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Dang ('" & foundSoRaw & "'); van ban hanh chinh ghi dang 'So: 12/QD-UBND' (ND 30/2020)."
End If
ElseIf Not isStandaloneNonDispatch Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So van ban] Khong tim thay dong So, ky hieu van ban o phan Tieu de."
End If
If Len(foundDateRaw) > 0 Then
Dim pNg As Long, pTh As Long, pNm As Long
pNg = InStr(1, foundDateRaw, "NGAY")
pTh = InStr(1, foundDateRaw, "THANG")
pNm = InStr(1, foundDateRaw, "NAM")
If pNg > 0 And pTh > pNg And pNm > pTh Then
Dim sDay As String, sMonth As String, sYear As String
sDay = Trim(Mid(foundDateRaw, pNg + 4, pTh - (pNg + 4)))
sMonth = Trim(Mid(foundDateRaw, pTh + 5, pNm - (pTh + 5)))
sYear = Trim(Left(Trim(Mid(foundDateRaw, pNm + 3)), 4))
If Not IsNumeric(sDay) Or Not IsNumeric(sMonth) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Dang de trong ngay/thang ban hanh ('" & foundDateRaw & "') cho ngay ky duyet."
Else
Dim dInt As Integer, mInt As Integer, yInt As Integer
dInt = Val(sDay): mInt = Val(sMonth): yInt = Val(sYear)
If mInt < 1 Or mInt > 12 Or dInt < 1 Or dInt > 31 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Ngay thang khong hop le: ngay " & dInt & " thang " & mInt & "."
ElseIf (mInt = 4 Or mInt = 6 Or mInt = 9 Or mInt = 11) And dInt > 30 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang " & mInt & " chi co toi da 30 ngay (dang ghi ngay " & dInt & ")."
ElseIf mInt = 2 And dInt > 29 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang 02 khong the co ngay " & dInt & "."
End If
End If
End If
End If
If foundCode = "SY" Or foundCode = "TRS" Or foundCode = "SL" Then foundCode = ""
If Len(foundCode) > 0 Then
Dim expectTitle As String
expectTitle = TT_DocCodeTitle(foundCode)
Select Case foundCode
Case "TTR": foundCode = "TTr"
Case "CTR": foundCode = "CTr"
Case "QYD": foundCode = "QyD"
End Select
If Len(expectTitle) = 0 And foundSoStyle = "NHA_NUOC" And Len(foundTitle) = 0 And hasVvSummary Then expectTitle = "CONG VAN"
If expectTitle = "CONG VAN" Then
If Len(foundTitle) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Loai van ban] Ky hieu la Cong van ('CV') nhung lai co tieu de '" & foundTitle & "' o giua bai."
End If
If Not hasVvSummary Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Trich yeu Cong van] Cong van dang thieu dong trich yeu 'V/v...' duoi so van ban."
End If
If Not hasKinhGui Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Kinh gui] Cong van dang thieu muc 'Kinh gui:'."
End If
ElseIf Len(expectTitle) > 0 Then
If Len(foundTitle) = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Thieu tieu de] Ky hieu van ban la '" & foundCode & "' nhung thieu dong ten loai van ban ('" & Replace(expectTitle, "|", " / ") & "') o giua trang."
ElseIf InStr(1, "|" & expectTitle & "|", "|" & foundTitle & "|") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Mau thuan loai VB] Ky hieu ghi la '" & foundCode & "' nhung tieu de van ban lai ghi la '" & foundTitle & "'."
End If
End If
End If
If mTTQPPL And (InStr(1, docAscFull, "BAN HANH QUYET DINH") > 0 Or InStr(1, docAscFull, "BAN HANH NGHI QUYET") > 0) Then hasQuyetDinhVerb = True
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
If Not hasQuyetDinhVerb Or totalDieuCount = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Cau truc Quyet dinh] Van ban Quyet dinh nhung thieu dong 'QUYET DINH:' hoac thieu cac 'Dieu 1.', 'Dieu 2.'."
End If
If totalDieuCount > 0 And hasNoiNhan And Not hasNhuDieu Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Logic Noi nhan] Van ban Quyet dinh co Dieu khoan thi hanh nhung phan Noi nhan thieu dong '- Nhu Dieu ...;'."
End If
End If
If TT_B("CHECK_COLLECTIVE") And Len(foundIssuingOrgan) > 0 And (Len(foundSignAuth) > 0 Or Len(foundSignRole) > 0) Then
Dim isExecutive As Boolean: isExecutive = False
If foundIssuingOrgan Like "VAN PHONG*" Or foundIssuingOrgan Like "BAN TO CHUC*" Or _
foundIssuingOrgan Like "BAN TUYEN GIAO*" Or foundIssuingOrgan Like "BAN XAY DUNG*" Or _
foundIssuingOrgan Like "BAN DAN VAN*" Or foundIssuingOrgan Like "TRUNG TAM*" Then
isExecutive = True
End If
Dim isCollective As Boolean: isCollective = False
If Not isExecutive And (InStr(1, foundIssuingOrgan, "DANG UY") > 0 Or InStr(1, foundIssuingOrgan, "BAN CHAP HANH") > 0 Or InStr(1, foundIssuingOrgan, "BAN THUONG VU") > 0 Or InStr(1, foundIssuingOrgan, "UY BAN KIEM TRA") > 0 Or InStr(1, foundIssuingOrgan, "CHI BO") > 0) Then
isCollective = True
End If
If isCollective And Not (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la tap the ('" & foundIssuingOrgan & "') nhung chu ky thieu tham quyen 'T/M ...'."
ElseIf isExecutive And (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la '" & foundIssuingOrgan & "' (che do thu truong) nen khong duoc ky 'T/M' (" & foundSignAuth & ")."
End If
End If
If hasNoiNhan Then
If Not hasLuuVT And Not hasLuuOther Then
warnCount = warnCount + 1
If TT_S("MOTTO_KIND") = "DANG" Then
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong luu o cuoi (vi du: '- Luu Van phong Tinh uy.')."
Else
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong '- Luu: VT...' o cuoi."
End If
ElseIf Not hasLuuVT And hasLuuOther Then
warnCount = warnCount + 1
If TT_S("MOTTO_KIND") = "DANG" Then
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen ghi noi luu (vi du theo HD 05-HD/VPTW: '- Luu Van phong Tinh uy.')."
Else
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen bo sung ky hieu 'VT' (Van thu) (vi du: '- Luu: VT, ...')."
End If
End If
End If
If TT_S("MOTTO_KIND") = "DANG" And Len(foundIssuingOrgan) > 0 Then
Dim issIsLocalCommittee As Boolean
issIsLocalCommittee = (foundIssuingOrgan Like "DANG UY XA *" Or foundIssuingOrgan Like "DANG UY PHUONG *" Or _
foundIssuingOrgan Like "DANG UY DAC KHU *" Or foundIssuingOrgan Like "CHI BO *")
Dim issIsStaffOrgan As Boolean
issIsStaffOrgan = (foundIssuingOrgan Like "VAN PHONG*" Or foundIssuingOrgan Like "BAN XAY DUNG DANG*" Or _
foundIssuingOrgan Like "BAN TO CHUC*" Or foundIssuingOrgan Like "BAN TUYEN GIAO*" Or _
foundIssuingOrgan Like "BAN DAN VAN*" Or foundIssuingOrgan Like "CO QUAN UY BAN KIEM TRA*")
If (issIsLocalCommittee Or issIsStaffOrgan) And Len(foundSuperiorOrgan) = 0 Then
warnCount = warnCount + 1
If issIsStaffOrgan Then
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] '" & foundIssuingOrgan & "' la co quan tham muu, giup viec: dong tren phai ghi ten cap uy ma co quan do truc thuoc (HD 05-HD/VPTW muc 2.1.3), vi du 'DANG UY XA ...'."
Else
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Van ban cua '" & foundIssuingOrgan & "' phai ghi ten dang bo cap tren truc tiep o dong tren (HD 05-HD/VPTW muc 2.1.2), vi du 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."
End If
ElseIf issIsLocalCommittee And Not (foundIssuingOrgan Like "CHI BO *") Then
If foundSuperiorOrgan Like "HUYEN UY*" Or foundSuperiorOrgan Like "QUAN UY*" Or foundSuperiorOrgan Like "THI UY*" Or _
foundSuperiorOrgan Like "DANG BO HUYEN*" Or foundSuperiorOrgan Like "DANG BO QUAN*" Or foundSuperiorOrgan Like "DANG BO THI XA*" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren dang ghi '" & foundSuperiorOrgan & "' (mo hinh cap huyen cu). Dang bo xa, phuong, dac khu truc thuoc truc tiep Tinh uy, Thanh uy: ghi 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."
ElseIf foundSuperiorOrgan Like "TINH UY*" Or foundSuperiorOrgan Like "THANH UY*" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren ghi ten DANG BO cap tren truc tiep (vi du 'DANG BO TINH ...'), khong ghi '" & foundSuperiorOrgan & "' (HD 05-HD/VPTW muc 2.1.2)."
End If
End If
End If
If Not isStandaloneNonDispatch Then
If TT_S("MOTTO_KIND") = "DANG" And InStr(1, docAscFull, "DANG CONG SAN VIET NAM") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tieu ngu] Van ban Dang chua co dong 'DANG CONG SAN VIET NAM' o dau trang."
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" And InStr(1, docAscFull, "CONG HOA XA HOI CHU NGHIA") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Van ban hanh chinh chua co Quoc hieu 'CONG HOA XA HOI CHU NGHIA VIET NAM' va Tieu ngu."
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" And InStr(1, Replace(docAscFull, " ", ""), "DOCLAP-TUDO-HANHPHUC") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Thieu dong Tieu ngu 'Doc lap - Tu do - Hanh phuc' duoi Quoc hieu."
End If
End If
If TT_B("RANK_BEFORE_NAME") Then
Dim signerName As String
signerName = TT_FindSignerName(doc)
If Len(signerName) > 0 Then
If Not TT_HasRankPrefix(ToAscii(signerName)) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Nguoi ky] Ho ten nguoi ky ('" & ToAscii(signerName) & "') chua ghi cap bac ham phia truoc (vi du: 'Dai ta Nguyen Van A'). Kiem tra lai neu don vi co quy dinh ghi cap bac."
End If
End If
End If
KiemTra_LogicTheThuc = warnings
End Function
Private Function ToAscii(ByVal s As String) As String
Dim n As Long
n = Len(s)
If n = 0 Then
ToAscii = ""
Exit Function
End If
Dim res As String
res = Space$(n)
Dim k As Long
Dim w As Long
For k = 1 To n
w = AscW(Mid$(s, k, 1))
If w < 0 Then w = w + 65536
Select Case w
Case 272, 273
Mid$(res, k, 1) = "D"
Case 192 To 195, 224 To 227, 258, 259, 7840 To 7863
Mid$(res, k, 1) = "A"
Case 200 To 202, 232 To 234, 278, 279, 7864 To 7879
Mid$(res, k, 1) = "E"
Case 204, 205, 236, 237, 296, 297, 7880 To 7883
Mid$(res, k, 1) = "I"
Case 210 To 214, 242 To 246, 416, 417, 7884 To 7907
Mid$(res, k, 1) = "O"
Case 217, 218, 249, 250, 360, 361, 431, 432, 7908 To 7921
Mid$(res, k, 1) = "U"
Case 221, 253, 7922 To 7929
Mid$(res, k, 1) = "Y"
Case 8211, 8212
Mid$(res, k, 1) = "-"
Case Else
If w >= 97 And w <= 122 Then
Mid$(res, k, 1) = Chr$(w - 32)
ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or _
w = 32 Or w = 40 Or w = 41 Or w = 42 Or w = 45 Or w = 46 Or w = 47 Or w = 58 Or w = 59 Then
Mid$(res, k, 1) = Chr$(w)
Else
Mid$(res, k, 1) = " "
End If
End Select
Next k
ToAscii = res
End Function
Private Function IsPartyDocTitle(ByVal sAsc As String) As Boolean
Dim t As String
t = Trim(sAsc)
If Left(t, 10) = "(DU THAO) " Then t = Trim(Mid(t, 11))
If Left(t, 8) = "DU THAO " Then t = Trim(Mid(t, 9))
If Len(t) = 0 Or Len(t) > 45 Then
IsPartyDocTitle = False
Exit Function
End If
Select Case t
Case "NGHI QUYET", "QUYET DINH", "CHI THI", "KET LUAN", "QUY CHE", _
"QUY DINH", "THONG TRI", "HUONG DAN", "THONG BAO", "KE HOACH", _
"CHUONG TRINH", "BAO CAO", "BIEN BAN", "TO TRINH", "CONG DIEN", _
"DE AN", "PHUONG AN", "GIAY MOI", "GIAY GIOI THIEU", _
"BAN CAM KET", "BAN DANG KY", "DIEU LE", "CHUONG TRINH HANH DONG", _
"BAO CAO CHINH TRI", "BAO CAO TOM TAT", "BAO CAO KIEM DIEM", _
"BAO CAO TONG KET", "BAO CAO SO KET", "DE CUONG", "DE CUONG BAO CAO", _
"DE CUONG TUYEN TRUYEN", "QUY CHE LAM VIEC", "QUY CHE HOAT DONG", _
"QUY CHE PHOI HOP", "CHUONG TRINH LAM VIEC", "CHUONG TRINH CONG TAC", _
"CHUONG TRINH KIEM TRA", "PHU LUC", "DANH MUC", "BANG TONG HOP", _
"BIEU TONG HOP", "BAI THU HOACH", "BAI PHAT BIEU", "BAN KIEM DIEM", _
"THU MOI", "THU TRIEU TAP", "LICH LAM VIEC", "LICH CONG TAC", _
"THONG CAO", "TUYEN BO", "LOI KEU GOI", "DU AN", "HOP DONG", _
"BAN GHI NHO", "BAN THOA THUAN", "GIAY UY QUYEN", "GIAY NGHI PHEP", _
"PHIEU GUI", "PHIEU CHUYEN", "PHIEU BAO", "QUY HOACH", "GIAY CHUNG NHAN"
IsPartyDocTitle = True
Case Else
If t Like "PHU LUC *" And Len(t) <= 20 Then
IsPartyDocTitle = True
Else
IsPartyDocTitle = False
End If
End Select
End Function
Private Function NormalizeDocTitleForCompare(ByVal sAsc As String) As String
Dim t As String
t = Trim(sAsc)
If Left(t, 10) = "(DU THAO) " Then t = Trim(Mid(t, 11))
If Left(t, 8) = "DU THAO " Then t = Trim(Mid(t, 9))
If t Like "BAO CAO *" Then t = "BAO CAO"
If t Like "CHUONG TRINH *" Then t = "CHUONG TRINH"
If t Like "QUY CHE *" Then t = "QUY CHE"
If t Like "DE CUONG *" Then t = "DE CUONG"
If t Like "PHU LUC *" Then t = "PHU LUC"
NormalizeDocTitleForCompare = t
End Function
Private Function IsSeparatorLine(ByVal rawText As String) As Boolean
Dim cleanStr As String
cleanStr = Trim(Replace(Replace(Replace(rawText, vbCr, ""), vbLf, ""), " ", ""))
cleanStr = Replace(cleanStr, Chr(160), "")
cleanStr = Replace(cleanStr, vbTab, "")
If Len(cleanStr) < 2 Or Len(cleanStr) > 45 Then
IsSeparatorLine = False
Exit Function
End If
Dim isAllDash As Boolean: isAllDash = True
Dim i As Long
For i = 1 To Len(cleanStr)
Dim wCode As Long
wCode = AscW(Mid(cleanStr, i, 1))
If wCode < 0 Then wCode = wCode + 65536
If wCode <> 45 And wCode <> 95 And wCode <> 42 And _
wCode <> 8211 And wCode <> 8212 And wCode <> 8213 Then
isAllDash = False
Exit For
End If
Next i
IsSeparatorLine = isAllDash
End Function
Private Function IsStrictBodyOrHeadingStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tAsc As String: tAsc = Trim(sAsc)
Dim tRaw As String: tRaw = Trim(rawText)
If Len(tAsc) = 0 Then
IsStrictBodyOrHeadingStart = False
Exit Function
End If
If tAsc Like "CAN CU *" Or tAsc Like "- CAN CU *" Or _
tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
tAsc = "QUYET DINH:" Or tAsc = "QUYET DINH :" Or _
tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
tAsc Like "I-*" Or tAsc Like "II-*" Or tAsc Like "III-*" Or _
tAsc Like "PHAN I*" Or tAsc Like "PHAN 1*" Or tAsc Like "PHAN THU*" Or _
tAsc Like "MUC I*" Or tAsc Like "MUC 1*" Or _
tAsc Like "CHUONG I*" Or tAsc Like "CHUONG 1*" Or _
tAsc Like "DIEU 1*" Or tAsc Like "DIEU 2*" Or _
tAsc Like "1. *" Or tAsc Like "2. *" Then
IsStrictBodyOrHeadingStart = True
ElseIf Left(tRaw, 1) = "-" Or Left(tRaw, 1) = "+" Then
IsStrictBodyOrHeadingStart = True
Else
IsStrictBodyOrHeadingStart = False
End If
End Function
Private Function IsBodyStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
If IsStrictBodyOrHeadingStart(rawText, sAsc) Then
IsBodyStart = True
Exit Function
End If
Dim tAsc As String: tAsc = Trim(sAsc)
Dim tRaw As String: tRaw = Trim(rawText)
If tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
tAsc Like "BAN THUONG VU*" Or tAsc Like "BAN CHAP HANH*" Or _
tAsc Like "DANG UY*" Or tAsc Like "CHI BO*" Or _
tAsc Like "HOI NGHI*" Or tAsc Like "DAI HOI*" Or _
tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Then
IsBodyStart = True
Exit Function
End If
Dim lastCh As String: lastCh = Right(tRaw, 1)
If lastCh = ";" Or lastCh = ":" Or Len(tRaw) > 130 Then
IsBodyStart = True
Else
IsBodyStart = False
End If
End Function
Private Function CheckLicenseActivation() As Boolean
CheckLicenseActivation = True
End Function
Private Function CallHttpApi(ByVal url As String) As String
Dim res As String
res = ""
#If Mac Then
On Error Resume Next
Dim sc As String
sc = "do shell script ""curl -s -L --max-time 3 '" & url & "'"""
res = MacScript(sc)
On Error GoTo 0
#Else
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
End If
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 2000, 3000, 3000, 3000
http.Send
If Err.Number = 0 Then
res = http.responseText
End If
On Error GoTo 0
#End If
CallHttpApi = Trim(res)
End Function
Private Function GetMachineName() As String
Dim mName As String
#If Mac Then
On Error Resume Next
mName = MacScript("do shell script ""scutil --get ComputerName || hostname -s || whoami""")
mName = Trim(Replace(mName, vbCrLf, ""))
mName = Trim(Replace(mName, vbLf, ""))
If Len(mName) = 0 Then mName = Environ("USER")
On Error GoTo 0
#Else
mName = Environ("COMPUTERNAME")
If Len(mName) = 0 Then mName = Environ("USERNAME")
#End If
If Len(mName) = 0 Then mName = "UNKNOWN_PC"
GetMachineName = mName
End Function
Private Function ParseJsonField(ByVal json As String, ByVal fieldName As String) As String
Dim pos As Long, startPos As Long, endPos As Long
pos = InStr(json, """" & fieldName & """")
If pos > 0 Then
startPos = InStr(pos + Len(fieldName) + 2, json, """")
If startPos > 0 Then
endPos = InStr(startPos + 1, json, """")
If endPos > startPos Then
ParseJsonField = Mid(json, startPos + 1, endPos - startPos - 1)
End If
End If
End If
End Function
Private Function EncodeUrl(ByVal s As String) As String
Dim i As Long, ch As String, code As Long
Dim res As String
res = ""
For i = 1 To Len(s)
ch = Mid(s, i, 1)
code = AscW(ch) And &HFFFF&
If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
res = res & ch
ElseIf code < 128 Then
res = res & "%" & Right("0" & Hex(code), 2)
Else
res = res & "_"
End If
Next i
EncodeUrl = res
End Function
Public Sub AutoExec()
Call CN_HenKiemTraKhiMo
End Sub
Public Function CN_PhienBan() As String
CN_PhienBan = "2026.11.09.5"
End Function
Private Function CN_ManifestUrl() As String
CN_ManifestUrl = "https://raw.githubusercontent.com/anhtu604/addin-the-thuc-van-ban/main/version.json"
End Function
Private Function CN_SoSanh(ByVal a As String, ByVal b As String) As Integer
Dim pa() As String, pb() As String
pa = Split(a, ".")
pb = Split(b, ".")
Dim i As Integer, va As Long, vb As Long
For i = 0 To 5
va = 0: vb = 0
If i <= UBound(pa) Then va = Val(pa(i))
If i <= UBound(pb) Then vb = Val(pb(i))
If va > vb Then
CN_SoSanh = 1
Exit Function
ElseIf va < vb Then
CN_SoSanh = -1
Exit Function
End If
Next i
CN_SoSanh = 0
End Function
Private Function CN_FileTuChay() As String
On Error Resume Next
CN_FileTuChay = ThisDocument.FullName
On Error GoTo 0
End Function
Private Function CN_DungFileAddIn() As Boolean
Dim n As String
n = LCase(Mid(CN_FileTuChay(), InStrRev(Replace(CN_FileTuChay(), "\", "/"), "/") + 1))
CN_DungFileAddIn = (n = "chuanhoa_vanban_rieng.dotm")
End Function
Private Function CN_ShQ(ByVal s As String) As String
CN_ShQ = "'" & Replace(s, "'", "'\''") & "'"
End Function
Private Function CN_TaiManifest(Optional ByVal nhanh As Boolean = False) As String
Dim i As Integer, res As String, soLan As Integer, cho As Integer
soLan = IIf(nhanh, 1, 2)
cho = IIf(nhanh, 3, 15)
For i = 1 To soLan
res = ""
#If Mac Then
On Error Resume Next
res = MacScript("do shell script ""curl -s -L --max-time " & cho & " " & CN_ShQ(CN_ManifestUrl()) & """")
On Error GoTo 0
#Else
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", CN_ManifestUrl(), False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 5000, cho * 1000, cho * 1000
http.Send
If Err.Number = 0 Then
If http.Status = 200 Then res = http.responseText
End If
On Error GoTo 0
#End If
res = Trim(res)
If Len(ParseJsonField(res, "version")) > 0 Then Exit For
Next i
CN_TaiManifest = res
End Function
Public Sub RibbonCapNhat(control As Object)
Call CN_KiemTraCapNhat(True)
End Sub
Public Sub CN_HenKiemTraKhiMo()
Static daHen As Boolean
If daHen Then Exit Sub
daHen = True
#If Mac Then
Call CN_TuKiemTraHangTuan(True)
#Else
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 20), "CN_KiemTraNhanhKhiMo"
#End If
End Sub
Public Sub CN_KiemTraNhanhKhiMo()
Call CN_TuKiemTraHangTuan(True)
End Sub
Public Sub CN_TuKiemTraHangTuan(Optional ByVal nhanh As Boolean = False)
On Error Resume Next
Dim lastChk As String, daDocDuoc As Boolean
lastChk = GetSetting("OfficePartySuite", "CapNhat", "LanKiemTra_WORD", "")
If Len(lastChk) > 0 Then
If IsDate(lastChk) Then
If DateDiff("d", CDate(lastChk), Date) < 1 Then Exit Sub
End If
End If
Dim lanThu As String
lanThu = GetSetting("OfficePartySuite", "CapNhat", "LanThu_WORD", "")
If Len(lanThu) > 0 Then
If IsDate(lanThu) Then
If DateDiff("n", CDate(lanThu), Now) < 60 Then Exit Sub
End If
End If
SaveSetting "OfficePartySuite", "CapNhat", "LanThu_WORD", Format(Now, "yyyy-mm-dd hh:nn:ss")
On Error GoTo 0
Call CN_KiemTraCapNhat(False, nhanh, daDocDuoc)
If daDocDuoc Then SaveSetting "OfficePartySuite", "CapNhat", "LanKiemTra_WORD", Format(Date, "yyyy-mm-dd")
End Sub
Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
If Len(CN_ManifestUrl()) = 0 Then
daDocDuoc = True
If hoiKhiKhongCo Then MsgBox "Ban quan ly rieng - phien ban " & CN_PhienBan() & "." & vbCrLf & "Ban nay KHONG tu cap nhat tu kenh cua ban goc. Khi co ban sua moi, cai lai bang bo cai dat.", vbInformation, "Phien ban Add-in"
Exit Sub
End If
Dim js As String
js = CN_TaiManifest(nhanh)
Dim ver As String, url As String, sha As String, notes As String
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "word_url")
sha = LCase(ParseJsonField(js, "word_sha256"))
notes = ParseJsonField(js, "notes_ascii")
If Len(ver) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Then
If hoiKhiKhongCo Then MsgBox "Khong ket noi duoc kenh cap nhat (GitHub " & "anhtu604/addin-the-thuc-van-ban" & "). Kiem tra mang roi thu lai." & vbCrLf & "Phien ban dang dung: " & CN_PhienBan(), vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
daDocDuoc = True
If CN_SoSanh(ver, CN_PhienBan()) <= 0 Then
If hoiKhiKhongCo Then MsgBox "Ban dang dung phien ban moi nhat: " & CN_PhienBan(), vbInformation, "Cap nhat Add-in"
Exit Sub
End If
If Not CN_DungFileAddIn() Then
MsgBox "Co phien ban moi " & ver & " nhung macro dang chay tu '" & CN_FileTuChay() & "', khong phai file add-in da cai." & vbCrLf & _
"Hay cai lai bang bo cai dat.", vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
Dim ans As VbMsgBoxResult
ans = MsgBox("Co phien ban moi: " & ver & " (dang dung " & CN_PhienBan() & ")." & vbCrLf & vbCrLf & notes & vbCrLf & vbCrLf & "Cap nhat ngay?", vbYesNo + vbQuestion, "Cap nhat Add-in")
If ans <> vbYes Then Exit Sub
Dim msg As String
msg = CN_TaiVaCai(url, sha)
If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cap nhat Add-in"
End Sub
Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String) As String
Dim target As String
target = CN_FileTuChay()
#If Mac Then
Dim tmp As String
tmp = Environ("HOME") & "/cn_capnhat_WORD.tmp"
Dim got As String
On Error Resume Next
got = MacScript("do shell script """ & "curl -s -L --fail --max-time 180 -o " & CN_ShQ(tmp) & " " & CN_ShQ(url) & _
" && shasum -a 256 " & CN_ShQ(tmp) & " | cut -d' ' -f1" & """")
On Error GoTo 0
If LCase(Trim(got)) <> sha Then
On Error Resume Next
MacScript "do shell script """ & "rm -f " & CN_ShQ(tmp) & """"
On Error GoTo 0
CN_TaiVaCai = "Tai ban cap nhat that bai hoac file khong khop ma kiem tra (SHA-256). Chua thay doi gi."
Exit Function
End If
Dim okCp As String
On Error Resume Next
okCp = MacScript("do shell script """ & "cp -f " & CN_ShQ(tmp) & " " & CN_ShQ(target) & " && xattr -c " & CN_ShQ(target) & _
" ; rm -f " & CN_ShQ(tmp) & " ; echo OK" & """")
On Error GoTo 0
If InStr(1, okCp, "OK") = 0 Then
CN_TaiVaCai = "Khong ghi duoc file add-in moi. Hay cai lai bang bo cai dat."
Exit Function
End If
CN_TaiVaCai = "Da cap nhat xong. Hay THOAT han Word (Cmd+Q) roi mo lai de dung phien ban moi."
#Else
On Error GoTo WinFail
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 10000, 15000, 60000, 180000
http.Send
If http.Status <> 200 Then GoTo WinFail
Dim tmpW As String
tmpW = Environ("TEMP") & "\cn_capnhat_WORD.tmp"
Dim stm As Object
Set stm = CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.Write http.responseBody
stm.SaveToFile tmpW, 2
stm.Close
Dim sh As Object, ex As Object, outTxt As String
Set sh = CreateObject("WScript.Shell")
Set ex = sh.Exec("certutil -hashfile """ & tmpW & """ SHA256")
outTxt = ex.StdOut.ReadAll
Dim lines() As String, hashW As String
lines = Split(Replace(outTxt, vbCr, ""), vbLf)
If UBound(lines) >= 1 Then hashW = LCase(Replace(Trim(lines(1)), " ", ""))
If hashW <> sha Then
Kill tmpW
CN_TaiVaCai = "File tai ve khong khop ma kiem tra (SHA-256). Chua thay doi gi."
Exit Function
End If
Dim tgtW As String
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
WinFail:
CN_TaiVaCai = "Tai ban cap nhat that bai (kiem tra mang). Chua thay doi gi."
#End If
End Function
Public Sub CN_KiemThuCapNhat()
Dim h As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim js As String
js = CN_TaiManifest()
Dim ver As String, url As String, sha As String, res As String
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "word_url")
sha = LCase(ParseJsonField(js, "word_sha256"))
res = "file=" & CN_FileTuChay() & " | dungFile=" & CN_DungFileAddIn() & " | dangDung=" & CN_PhienBan() & " | banMoi=" & ver
If Len(ver) > 0 And CN_SoSanh(ver, CN_PhienBan()) > 0 And CN_DungFileAddIn() Then res = res & " | " & CN_TaiVaCai(url, sha)
Dim f As Integer
f = FreeFile
Open h & "/cn_test_WORD.log" For Output As #f
Print #f, res
Close #f
End Sub

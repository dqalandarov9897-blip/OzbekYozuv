import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import re
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"gen2.py")).read()
pre = src.split("# ---------------- WORD ----------------")[0]
ns = {}
exec(pre, ns)
COMMON = ns["COMMON"]
COMMON = COMMON[:COMMON.index("' Istalgan matnni yangi alifboga o'giradi")]
COMMON = COMMON.replace("        Case 32, 9, 10, 11, 13, 160: IsSep = True", "        Case 32, 9, 10, 11, 13, 160, 7, 12, 1, 14, 30, 31: IsSep = True")
COMMON = COMMON.replace('''Private Function TokenOgir(ByVal t As String) As String
    If HasCyr(t) Then t = KirillToken(t)
    If Not Himoya(t) Then t = EskiToken(t)
    TokenOgir = t
End Function
''', "")
assert "TokenOgir" not in COMMON

COMMON += r'''
' ===== Lotin 2026 -> Lotin 1995 / Kirill =====
Private Function IsLatVowel(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsLatVowel = (InStr("aeiouAEIOUöÖ", ch) > 0)
End Function

Private Function LowX(ByVal ch As String) As String
    Select Case ch
        Case "Ö": LowX = "ö"
        Case "Ğ": LowX = "ğ"
        Case "Ş": LowX = "ş"
        Case "Ç": LowX = "ç"
        Case Else: LowX = LCase$(ch)
    End Select
End Function

Private Function To1995(ByVal t As String) As String
    Dim i As Long, n As Long, c As String, nx As String, pv As String, res As String, kat As Boolean
    n = Len(t)
    For i = 1 To n
        c = Mid$(t, i, 1)
        If i < n Then nx = Mid$(t, i + 1, 1) Else nx = ""
        If i > 1 Then pv = Mid$(t, i - 1, 1) Else pv = ""
        kat = IsUp(nx) Or (Not IsLet(nx) And IsUp(pv))
        Select Case c
            Case "ö": res = res & "oʻ"
            Case "Ö": res = res & "Oʻ"
            Case "ğ": res = res & "gʻ"
            Case "Ğ": res = res & "Gʻ"
            Case "ş": res = res & "sh"
            Case "Ş": If kat Then res = res & "SH" Else res = res & "Sh"
            Case "ç": res = res & "ch"
            Case "Ç": If kat Then res = res & "CH" Else res = res & "Ch"
            Case Else: res = res & c
        End Select
    Next i
    To1995 = res
End Function

Private Function UpCyr(ByVal k As Long) As Long
    If k >= &H430 And k <= &H44F Then UpCyr = k - &H20: Exit Function
    Select Case k
        Case &H451: UpCyr = &H401
        Case &H45E: UpCyr = &H40E
        Case &H493: UpCyr = &H492
        Case &H49B: UpCyr = &H49A
        Case &H4B3: UpCyr = &H4B2
        Case Else: UpCyr = k
    End Select
End Function

Private Function Cy(ByVal k As Long, ByVal up As Boolean) As String
    If up Then Cy = ChrW(UpCyr(k)) Else Cy = ChrW(k)
End Function

Private Function LatCyr(ByVal lc As String) As Long
    Select Case lc
__LATCYR__
        Case Else: LatCyr = 0
    End Select
End Function

Private Function ToKirill(ByVal t As String) As String
    Dim i As Long, n As Long, c As String, lc As String, nx As String, ln As String
    Dim pv As String, n2 As String, res As String, k As Long, up As Boolean
    n = Len(t): i = 1
    Do While i <= n
        c = Mid$(t, i, 1): lc = LowX(c): up = IsUp(c)
        nx = Mid$(t, i + 1, 1): ln = LowX(nx)
        n2 = Mid$(t, i + 2, 1)
        If i > 1 Then pv = Mid$(t, i - 1, 1) Else pv = ""
        If lc = "y" And (ln = "e" Or ln = "o" Or ln = "u" Or ln = "a") Then
            Select Case ln
                Case "e": k = &H435
                Case "o": k = &H451
                Case "u": k = &H44E
                Case Else: k = &H44F
            End Select
            res = res & Cy(k, up): i = i + 2
        ElseIf lc = "s" And nx = "ʼ" And LowX(n2) = "h" Then
            res = res & Cy(&H441, up) & Cy(&H4B3, IsUp(n2)): i = i + 3
        ElseIf lc = "e" Then
            If Not IsLet(pv) Or IsLatVowel(pv) Then k = &H44D Else k = &H435
            res = res & Cy(k, up): i = i + 1
        ElseIf c = "ʼ" Then
            If IsLet(pv) And IsLet(nx) Then res = res & Cy(&H44A, IsUp(pv) And IsUp(nx)) Else res = res & c
            i = i + 1
        Else
            If lc = "c" Then
                If ln = "e" Or ln = "i" Or ln = "y" Then k = &H441 Else k = &H43A
            ElseIf lc = "w" Then
                k = &H432
            Else
                k = LatCyr(lc)
            End If
            If k = 0 Then res = res & c Else res = res & Cy(k, up)
            i = i + 1
        End If
    Loop
    ToKirill = res
End Function

' m: 1 = Lotin 2026, 2 = Lotin 1995, 3 = Kirill
Private Function TokenOgir(ByVal t As String, ByVal m As Long) As String
    Dim cyr As Boolean
    cyr = HasCyr(t)
    If m = 3 And cyr Then TokenOgir = t: Exit Function
    If cyr Then
        t = KirillToken(t)
    Else
        t = EskiToken(t)
    End If
    Select Case m
        Case 2: t = To1995(t)
        Case 3: t = ToKirill(t)
    End Select
    TokenOgir = t
End Function

' Istalgan matnni tanlangan yozuvga o'giradi
Public Function OzbekOgir(ByVal matn As String, ByVal m As Long) As String
    Dim i As Long, ch As String, tok As String, res As String
    For i = 1 To Len(matn)
        ch = Mid$(matn, i, 1)
        If IsSep(ch) Then
            If Len(tok) > 0 Then
                res = res & TokenOgir(tok, m)
                tok = ""
            End If
            res = res & ch
        Else
            tok = tok & ch
        End If
    Next i
    If Len(tok) > 0 Then res = res & TokenOgir(tok, m)
    OzbekOgir = res
End Function
'''
LC = {"a":0x430,"b":0x431,"d":0x434,"f":0x444,"g":0x433,"h":0x4B3,"i":0x438,"j":0x436,"k":0x43A,"l":0x43B,"m":0x43C,"n":0x43D,"o":0x43E,"p":0x43F,"q":0x49B,"r":0x440,"s":0x441,"t":0x442,"u":0x443,"v":0x432,"x":0x445,"y":0x439,"z":0x437,"ö":0x45E,"ğ":0x493,"ş":0x448,"ç":0x447}
COMMON = COMMON.replace("__LATCYR__", "\n".join(f'        Case "{a}": LatCyr = &H{b:X}' for a, b in LC.items()))

# ---------------- WORD ----------------
kit = src[src.index("' ---------------------------------------------------------------------\n'  2. KITOBCHA"):src.index("''' + COMMON")]
WORD = r'''Attribute VB_Name = "OzbekYozuv"
' =====================================================================
'  OzbekYozuv 1.0 - Word uchun: Lotin 2026 / Lotin 1995 / Kirill va Kitobcha
'  Tugmalar: Lotin 2026 | Lotin 1995 | Kirill | Kitobcha
'  ORNATISH.vbs bu modulni Word STARTUP papkasiga OzbekYozuv.dotm qilib joylaydi
'  Word 2010 - 2021 va Microsoft 365, Windows 7 - 11
' =====================================================================
Option Explicit

Public Sub TugmalarniOrnatish()
    Dim cb As CommandBar, sv As Boolean
    sv = NormalTemplate.Saved
    CustomizationContext = NormalTemplate
    On Error Resume Next
    CommandBars("Ozbek alifbo").Delete
    CommandBars("Ozbek yozuvlari").Delete
    On Error GoTo 0
    Set cb = CommandBars.Add(Name:="Ozbek yozuvlari", Position:=msoBarTop, Temporary:=True)
    Tugma cb, "Lotin 2026", "Lotin2026ga", 2059, "Matnni lotin 2026 alifbosiga o'girish (Ö, Ğ, Ş, Ç)"
    Tugma cb, "Lotin 1995", "Lotin1995ga", 2059, "Matnni lotin 1995 alifbosiga o'girish (Oʻ, Gʻ, Sh, Ch)"
    Tugma cb, "Kirill", "Kirillga", 2059, "Matnni kirill alifbosiga o'girish"
    Tugma cb, "Kitobcha", "Kitobcha", 4, "Hujjatning kitobcha nusxasini tayyorlash va chop etish"
    cb.Visible = True
    NormalTemplate.Saved = sv
End Sub

Public Sub AutoExec()
    On Error Resume Next
    TugmalarniOrnatish
End Sub

Private Sub Tugma(cb As CommandBar, nom As String, mac As String, ico As Long, izoh As String)
    With cb.Controls.Add(Type:=msoControlButton)
        .Caption = nom
        .OnAction = mac
        .FaceId = ico
        .Style = msoButtonIconAndCaption
        .TooltipText = izoh
    End With
End Sub

' ---------------------------------------------------------------------
'  1. YOZUVNI O'GIRISH
'  Belgilangan matn bo'lsa - faqat o'sha qism, bo'lmasa - butun hujjat.
'  Kirish matni istalgan yozuvda bo'lishi mumkin. Hammasi bitta Ctrl+Z bilan qaytariladi.
'  Cheklovsiz: barcha so'zlar o'giriladi.
' ---------------------------------------------------------------------
Public Sub Lotin2026ga()
    Ogir 1, "Lotin 2026"
End Sub

Public Sub Lotin1995ga()
    Ogir 2, "Lotin 1995"
End Sub

Public Sub Kirillga()
    Ogir 3, "Kirill"
End Sub

Public Sub YangiAlifbogaOgirish()
    Ogir 1, "Lotin 2026"
End Sub

Private Sub Ogir(ByVal m As Long, ByVal nom As String)
    Dim st As Range, ur As Object, xato As String
    If Documents.Count = 0 Then Exit Sub
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Hujjat tahrirdan himoyalangan - avval himoyani olib tashlang.", vbExclamation, "Ozbek yozuvlari"
        Exit Sub
    End If
    On Error Resume Next
    Set ur = Application.UndoRecord
    ur.StartCustomRecord nom & " yozuviga o'girish"
    Err.Clear
    On Error GoTo Xatolik
    Application.ScreenUpdating = False
    If Selection.End > Selection.Start Then
        OgirRange Selection.Range, m
    Else
        For Each st In ActiveDocument.StoryRanges
            Do While Not st Is Nothing
                OgirRange st, m
                Set st = st.NextStoryRange
            Loop
        Next st
    End If
Tugash:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ur Is Nothing Then ur.EndCustomRecord
    On Error GoTo 0
    If Len(xato) > 0 Then
        MsgBox "Matnning bir qismini o'girib bo'lmadi: " & xato & vbCrLf & "O'girilgan qismini Ctrl+Z bilan qaytarish mumkin.", vbExclamation, "Ozbek yozuvlari"
    Else
        Application.StatusBar = nom & " yozuviga o'girildi. Qaytarish - Ctrl+Z"
    End If
    Exit Sub
Xatolik:
    xato = Err.Description
    Resume Tugash
End Sub

Private Sub OgirRange(rng As Range, ByVal m As Long)
    Dim p As Paragraph, q As Range
    For Each p In rng.Paragraphs
        Set q = p.Range.Duplicate
        If q.Start < rng.Start Then q.Start = rng.Start
        If q.End > rng.End Then q.End = rng.End
        If q.End > q.Start Then OgirQism q, m
    Next p
End Sub

Private Sub OgirQism(q As Range, ByVal m As Long)
    Dim s As String, i As Long, w As Range
    s = KatakBelgisiz(q.Text, q.End - q.Start)
    If Len(s) = q.End - q.Start Then
        OgirMatn q, s, m
    Else
        ' Maydonlar (havola, sahifa raqami) bo'lsa - so'zma-so'z
        For i = q.Words.Count To 1 Step -1
            Set w = q.Words(i)
            s = KatakBelgisiz(w.Text, w.End - w.Start)
            If Len(s) = w.End - w.Start And Len(s) > 0 Then OgirMatn w, s, m
        Next i
    End If
End Sub

' Jadval katagi oxiridagi belgilar (Chr(13) & Chr(7)) matnda 2 ta, hujjatda 1 ta joy egallaydi
Private Function KatakBelgisiz(ByVal s As String, ByVal uz As Long) As String
    Do While Len(s) > uz And Len(s) > 0
        Select Case Right$(s, 1)
            Case Chr(7), Chr(13)
                s = Left$(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    KatakBelgisiz = s
End Function

Private Sub OgirMatn(q As Range, ByVal s As String, ByVal m As Long)
    Dim i As Long, j As Long, tok As String, nv As String, r As Range, base As Long
    base = q.Start
    i = Len(s)
    Do While i >= 1
        If IsSep(Mid$(s, i, 1)) Then
            i = i - 1
        Else
            j = i
            Do While j > 1
                If IsSep(Mid$(s, j - 1, 1)) Then Exit Do
                j = j - 1
            Loop
            tok = Mid$(s, j, i - j + 1)
            nv = TokenOgir(tok, m)
            If nv <> tok Then
                Set r = q.Duplicate
                r.SetRange base + j - 1, base + i
                r.Text = nv
            End If
            i = j - 1
        End If
    Loop
End Sub

''' + kit + COMMON

# ---------------- EXCEL ----------------
EXCEL = r'''Attribute VB_Name = "OzbekYozuvExcel"
' =====================================================================
'  OzbekYozuv 1.0 - Excel uchun: Lotin 2026 / Lotin 1995 / Kirill
'  Tugmalar: Lotin 2026 | Lotin 1995 | Kirill
'  Formulalar: =LOTIN2026(A1)  =LOTIN1995(A1)  =KIRILL(A1)
'  ORNATISH.vbs bu modulni Excel XLSTART papkasiga OzbekYozuv.xlam qilib joylaydi
'  Excel 2010 - 2021 va Microsoft 365, Windows 7 - 11
' =====================================================================
Option Explicit

Private eski As Collection

Public Sub Auto_Open()
    Dim cb As CommandBar
    On Error Resume Next
    Application.CommandBars("Ozbek alifbo").Delete
    Application.CommandBars("Ozbek yozuvlari").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add(Name:="Ozbek yozuvlari", Position:=msoBarTop, Temporary:=True)
    Tugma cb, "Lotin 2026", "Lotin2026ga", "Lotin 2026 alifbosiga o'girish (Ö, Ğ, Ş, Ç)"
    Tugma cb, "Lotin 1995", "Lotin1995ga", "Lotin 1995 alifbosiga o'girish (Oʻ, Gʻ, Sh, Ch)"
    Tugma cb, "Kirill", "Kirillga", "Kirill alifbosiga o'girish"
    cb.Visible = True
End Sub

Private Sub Tugma(cb As CommandBar, nom As String, mac As String, izoh As String)
    With cb.Controls.Add(Type:=msoControlButton)
        .Caption = nom
        .OnAction = "'" & ThisWorkbook.Name & "'!" & mac
        .FaceId = 2059
        .Style = msoButtonIconAndCaption
        .TooltipText = izoh
    End With
End Sub

Public Sub Lotin2026ga()
    OgirExcel 1, "Lotin 2026"
End Sub

Public Sub Lotin1995ga()
    OgirExcel 2, "Lotin 1995"
End Sub

Public Sub Kirillga()
    OgirExcel 3, "Kirill"
End Sub

Public Sub YangiAlifbogaOgirishExcel()
    OgirExcel 1, "Lotin 2026"
End Sub

' Belgilangan kataklar (bitta katak tanlansa - butun varaq). Formulalar o'zgarmaydi. Qaytarish: Ctrl+Z
Private Sub OgirExcel(ByVal m As Long, ByVal nom As String)
    Dim rng As Range, c As Range, v As String, nv As String, otkaz As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge = 1 Then
        Set rng = ActiveSheet.UsedRange
    Else
        Set rng = Selection
    End If
    On Error Resume Next
    Set rng = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "O'giriladigan matn topilmadi.", vbInformation, "Ozbek yozuvlari"
        Exit Sub
    End If
    On Error GoTo 0
    Set eski = New Collection
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        v = CStr(c.Value)
        nv = OzbekOgir(v, m)
        If nv <> v Then
            On Error Resume Next
            c.Value = nv
            If Err.Number = 0 Then
                eski.Add Array(c.Worksheet.Parent.Name, c.Worksheet.Name, c.Address, v)
            Else
                otkaz = otkaz + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next c
    Application.ScreenUpdating = True
    If otkaz > 0 Then MsgBox otkaz & " ta katakni o'zgartirib bo'lmadi (varaq himoyalangan bo'lishi mumkin).", vbExclamation, "Ozbek yozuvlari"
    If eski.Count > 0 Then
        Application.OnUndo nom & " yozuviga o'girishni qaytarish", "'" & ThisWorkbook.Name & "'!QaytarishExcel"
    End If
    Application.StatusBar = nom & ": " & eski.Count & " ta katak o'girildi. Qaytarish - Ctrl+Z"
End Sub

Public Sub QaytarishExcel()
    Dim a As Variant
    If eski Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For Each a In eski
        Workbooks(a(0)).Worksheets(a(1)).Range(a(2)).Value = a(3)
    Next a
    Application.ScreenUpdating = True
    Set eski = Nothing
    Application.StatusBar = False
End Sub

' ----- Formulalar -----
Public Function LOTIN2026(ByVal matn As String) As String
    LOTIN2026 = OzbekOgir(matn, 1)
End Function

Public Function LOTIN1995(ByVal matn As String) As String
    LOTIN1995 = OzbekOgir(matn, 2)
End Function

Public Function KIRILL(ByVal matn As String) As String
    KIRILL = OzbekOgir(matn, 3)
End Function

Public Function YANGIALIFBO(ByVal matn As String) As String
    YANGIALIFBO = OzbekOgir(matn, 1)
End Function
''' + COMMON

def enc(m):
    s = m.group(0)[1:-1]
    parts = []; buf = ""
    for ch in s:
        if ord(ch) < 128: buf += ch
        else:
            if buf: parts.append('"' + buf + '"'); buf = ""
            parts.append(f"ChrW(&H{ord(ch):04X})")
    if buf: parts.append('"' + buf + '"')
    return " & ".join(parts) if parts else '""'

def to_ascii(src):
    out = []
    for line in src.split("\n"):
        if any(ord(c) > 127 for c in line):
            assert not line.lstrip().startswith("'"), line
            line = re.sub(r'"(?:[^"]|"")*"', enc, line)
        assert all(ord(c) < 128 for c in line), line
        out.append(line)
    return "\n".join(out)

for name, s in [(os.path.join(ROOT,"src","word","OzbekYozuv_Word.bas"), WORD), (os.path.join(ROOT,"src","excel","OzbekYozuv_Excel.bas"), EXCEL)]:
    a = to_ascii(s)
    for ln in a.split("\n"):
        assert len(ln) < 1000
    open(name, "w", encoding="ascii", newline="\r\n").write(a)
    print(name, len(a))

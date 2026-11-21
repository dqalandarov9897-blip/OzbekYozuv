import re

AP = "'`‘’ʻʼ´"
LAT_L = "A-Za-zÖöĞğŞşÇç"
CYR_UP = "АБВГҒДЕЁЖЗИЙКҚЛМНОПРСТУЎФХҲЦЧШЩЪЫЬЭЮЯ"
VOW = "АаЕеЁёИиОоУуЎўЭэЮюЯя"
UPALL = CYR_UP + "A-ZÖĞŞÇ"
simple = [("А","A"),("Б","B"),("В","V"),("Г","G"),("Ғ","Ğ"),("Д","D"),("Ж","J"),("З","Z"),("И","I"),("Й","Y"),
 ("К","K"),("Қ","Q"),("Л","L"),("М","M"),("Н","N"),("О","O"),("П","P"),("Р","R"),("С","S"),("Т","T"),("У","U"),
 ("Ў","Ö"),("Ф","F"),("Х","X"),("Ҳ","H"),("Ч","Ç"),("Ш","Ş"),("Щ","Ş"),("Ы","I"),("Э","E")]

# Inglizcha/texnik so'zlar: o'girilmaydi
ISTISNO = ("chrome|chat|check|chip|cheese|chess|chicken|chief|child|choice|church|chart|channel|chapter|"
           "shop|show|shift|share|shell|shot|ship|shirt|shoe|short|should|shut|sheet|shelf|shine|"
           "flash|push|cash|crash|fresh|english|wash|finish|match|search|touch|teacher|school|machine|"
           "mechanic|technology|orchestra|chemistry|character|schedule|shampoo|chips|ketchup")

# ---------------- Umumiy satr konverteri (Excel va Word uchun) ----------------
COMMON = r'''
' ===== Satr konverteri: kirill va eski lotin -> yangi lotin (2026) =====
Private Function IsSep(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 32, 9, 10, 11, 13, 160: IsSep = True
    End Select
End Function

Private Function IsAp(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Select Case AscW(ch)
        Case 39, 96, &H2018, &H2019, &H2BB, &H2BC, &HB4: IsAp = True
    End Select
End Function

Private Function IsLet(ByVal ch As String) As Boolean
    Dim k As Long
    If Len(ch) = 0 Then Exit Function
    k = AscW(ch)
    If k < 0 Then k = k + 65536
    If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Then IsLet = True: Exit Function
    If (k >= &HC0 And k <= &H24F) Or (k >= &H400 And k <= &H4FF) Then IsLet = True
End Function

Private Function LowK(ByVal k As Long) As Long
    If k >= &H410 And k <= &H42F Then
        LowK = k + &H20
    Else
        Select Case k
            Case &H401: LowK = &H451
            Case &H40E: LowK = &H45E
            Case &H492: LowK = &H493
            Case &H49A: LowK = &H49B
            Case &H4B2: LowK = &H4B3
            Case Else: LowK = k
        End Select
    End If
End Function

Private Function IsUp(ByVal ch As String) As Boolean
    Dim k As Long
    If Len(ch) = 0 Then Exit Function
    k = AscW(ch)
    If k >= 65 And k <= 90 Then IsUp = True: Exit Function
    If k >= &H410 And k <= &H42F Then IsUp = True: Exit Function
    Select Case k
        Case &H401, &H40E, &H492, &H49A, &H4B2, &HD6, &H11E, &H15E, &HC7: IsUp = True
    End Select
End Function

Private Function IsCyrVowel(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Select Case LowK(AscW(ch))
        Case &H430, &H435, &H451, &H438, &H43E, &H443, &H45E, &H44D, &H44E, &H44F: IsCyrVowel = True
    End Select
End Function

Private Function UpX(ByVal s As String) As String
    s = UCase$(s)
    s = Replace(s, "ş", "Ş"): s = Replace(s, "ç", "Ç")
    s = Replace(s, "ö", "Ö"): s = Replace(s, "ğ", "Ğ")
    UpX = s
End Function

Private Function Cap(ByVal s As String, ByVal up As Boolean, ByVal nxUp As Boolean) As String
    If Not up Or Len(s) = 0 Then
        Cap = s
    ElseIf nxUp Or Len(s) = 1 Then
        Cap = UpX(s)
    Else
        Cap = UpX(Left$(s, 1)) & Mid$(s, 2)
    End If
End Function

Private Function SimpleMap(ByVal k As Long) As String
    Select Case k
__SIMPLE__
        Case Else: SimpleMap = ""
    End Select
End Function

Private Function HasCyr(ByVal t As String) As Boolean
    Dim i As Long, k As Long
    For i = 1 To Len(t)
        k = AscW(Mid$(t, i, 1))
        If k >= &H400 And k <= &H4FF Then HasCyr = True: Exit Function
    Next i
End Function

Private Function KirillToken(ByVal t As String) As String
    Dim i As Long, n As Long, c As String, pv As String, nx As String
    Dim up As Boolean, nxUp As Boolean, o As String, res As String, k As Long
    n = Len(t)
    For i = 1 To n
        c = Mid$(t, i, 1)
        If i > 1 Then pv = Mid$(t, i - 1, 1) Else pv = ""
        If i < n Then nx = Mid$(t, i + 1, 1) Else nx = ""
        up = IsUp(c)
        nxUp = IsUp(nx)
        k = LowK(AscW(c))
        Select Case k
            Case &H435
                If Not IsLet(pv) Or IsCyrVowel(pv) Or LowK(AscW(pv & " ")) = &H44A Or LowK(AscW(pv & " ")) = &H44C Then
                    o = Cap("ye", up, nxUp)
                Else
                    o = Cap("e", up, nxUp)
                End If
            Case &H451: o = Cap("yo", up, nxUp)
            Case &H44E: o = Cap("yu", up, nxUp)
            Case &H44F: o = Cap("ya", up, nxUp)
            Case &H446
                If IsCyrVowel(pv) Then o = Cap("ts", up, nxUp) Else o = Cap("s", up, nxUp)
            Case &H44A: o = "ʼ"
            Case &H44C
                If LowK(AscW(nx & " ")) = &H43E Then o = Cap("y", up, nxUp) Else o = ""
            Case &H441
                If LowK(AscW(nx & " ")) = &H4B3 Then o = Cap("s", up, nxUp) & "ʼ" Else o = Cap("s", up, nxUp)
            Case Else
                o = SimpleMap(k)
                If Len(o) = 0 Then o = c Else o = Cap(o, up, nxUp)
        End Select
        res = res & o
    Next i
    KirillToken = res
End Function

Private Function Himoya(ByVal t As String) As Boolean
    Dim L As String, w As String, i As Long, ch As String
    L = LCase$(t)
    If InStr(L, "@") > 0 Or InStr(L, "://") > 0 Or InStr(L, "www.") > 0 Then Himoya = True: Exit Function
    If InStr(L, "w") > 0 Then Himoya = True: Exit Function
    If SaytNomi(L) Then Himoya = True: Exit Function
    For i = 1 To Len(L)
        ch = Mid$(L, i, 1)
        If ch = "c" Then
            If Mid$(L, i + 1, 1) <> "h" Then Himoya = True: Exit Function
        End If
        If (ch >= "a" And ch <= "z") Then w = w & ch
    Next i
    If Len(w) > 0 Then
        If InStr("|__ISTISNO__|", "|" & w & "|") > 0 Then Himoya = True
    End If
End Function

' kun.uz, mail.ru kabi sayt nomlari (Sh.Mirziyoyev kabi bosh harflar emas)
Private Function SaytNomi(ByVal L As String) As Boolean
    Dim d As Long, tl As String
    Do While Len(L) > 0
        If Right$(L, 1) Like "[a-z]" Then Exit Do
        L = Left$(L, Len(L) - 1)
    Loop
    d = InStrRev(L, ".")
    If d < 2 Or d >= Len(L) Then Exit Function
    If Not (Mid$(L, d - 1, 1) Like "[a-z0-9]") Then Exit Function
    tl = Mid$(L, d + 1)
    If InStr("|uz|com|ru|org|net|info|gov|edu|io|kz|tj|kg|tr|de|me|app|ai|", "|" & tl & "|") > 0 Then SaytNomi = True
End Function

Private Function EskiToken(ByVal t As String) As String
    Dim i As Long, n As Long, c As String, nx As String, res As String
    n = Len(t): i = 1
    Do While i <= n
        c = Mid$(t, i, 1)
        If i < n Then nx = Mid$(t, i + 1, 1) Else nx = ""
        If IsAp(nx) And (c = "o" Or c = "O" Or c = "g" Or c = "G") Then
            Select Case c
                Case "o": res = res & "ö"
                Case "O": res = res & "Ö"
                Case "g": res = res & "ğ"
                Case "G": res = res & "Ğ"
            End Select
            i = i + 2
        ElseIf (nx = "h" Or nx = "H") And (c = "s" Or c = "S" Or c = "c" Or c = "C") Then
            Select Case c
                Case "s": res = res & "ş"
                Case "S": res = res & "Ş"
                Case "c": res = res & "ç"
                Case "C": res = res & "Ç"
            End Select
            i = i + 2
        ElseIf IsAp(c) And i > 1 And i < n Then
            If IsLet(Mid$(t, i - 1, 1)) And IsLet(nx) Then res = res & "ʼ" Else res = res & c
            i = i + 1
        Else
            res = res & c
            i = i + 1
        End If
    Loop
    EskiToken = res
End Function

Private Function TokenOgir(ByVal t As String) As String
    If HasCyr(t) Then t = KirillToken(t)
    If Not Himoya(t) Then t = EskiToken(t)
    TokenOgir = t
End Function

' Istalgan matnni yangi alifboga o'giradi
Public Function YANGIALIFBO(ByVal matn As String) As String
    Dim i As Long, ch As String, tok As String, res As String
    For i = 1 To Len(matn)
        ch = Mid$(matn, i, 1)
        If IsSep(ch) Then
            If Len(tok) > 0 Then
                res = res & TokenOgir(tok)
                tok = ""
            End If
            res = res & ch
        Else
            tok = tok & ch
        End If
    Next i
    If Len(tok) > 0 Then res = res & TokenOgir(tok)
    YANGIALIFBO = res
End Function
'''

sm = []
for a, b in simple:
    lo = ord(a.lower())
    sm.append(f'        Case &H{lo:X}: SimpleMap = "{b.lower()}"')
COMMON = COMMON.replace("__SIMPLE__", "\n".join(sm)).replace("__ISTISNO__", ISTISNO)

# ---------------- WORD ----------------
K = []
def R(f, t, w): K.append(f'    R rng, "{f}", "{t}", {"True" if w else "False"}')
R("сҳ", "sʼh", False); R("Сҳ", "Sʼh", False); R("СҲ", "SʼH", False)
R(f"<Е([{UPALL}])", "YE\\1", True); R("<Е", "Ye", True); R("<е", "ye", True)
R(f"([{VOW}ЪъЬь])Е", "\\1YE", True); R(f"([{VOW}ЪъЬь])е", "\\1ye", True)
R("Е", "E", False); R("е", "e", False)
for c, d in [("Ё", "O"), ("Ю", "U"), ("Я", "A")]:
    R(f"{c}([{UPALL}])", f"Y{d}\\1", True); R(c, f"Y{d.lower()}", False); R(c.lower(), f"y{d.lower()}", False)
R("<Ц", "S", True); R("<ц", "s", True)
R(f"([{VOW}])Ц", "\\1TS", True); R(f"([{VOW}])ц", "\\1ts", True)
R("Ц", "S", False); R("ц", "s", False)
R("ьо", "yo", False); R("ЬО", "YO", False); R("Ъ", "ʼ", False); R("ъ", "ʼ", False); R("Ь", "", False); R("ь", "", False)
for a, b in simple:
    R(a, b, False); R(a.lower(), b.lower(), False)
KIRILL_WORD = "Private Sub Kirill(rng As Range)\n" + "\n".join(K) + "\nEnd Sub\n"

E = []
def A(f, t, w, b=0, u=0): E.append(f'    Almash rng, "{f}", "{t}", {"True" if w else "False"}, {b}, {u}')
for a, b in [("o", "ö"), ("O", "Ö"), ("g", "ğ"), ("G", "Ğ")]: A(f"{a}[{AP}]", b, True)
A(f"[{LAT_L}][{AP}][{LAT_L}]", "ʼ", True, 1, 1)
for a, b in [("SH", "Ş"), ("Sh", "Ş"), ("sH", "ş"), ("sh", "ş"), ("CH", "Ç"), ("Ch", "Ç"), ("cH", "ç"), ("ch", "ç")]: A(a, b, False)
ESKI_WORD = "Private Sub EskiLotin(rng As Range)\n" + "\n".join(E) + "\nEnd Sub\n"

WORD = r'''Attribute VB_Name = "OzbekAlifbo"
' =====================================================================
'  Ozbek lotin 2026 - Word uchun yordamchi (v3, STARTUP qo'shimchasi)
'  Tugmalar: "Yangi alifboga o'girish" va "Kitobcha"
'  ornatish.vbs bu modulni Word STARTUP papkasiga OzbekAlifbo.dotm qilib joylaydi
'  Word 2010 - 2021 va Microsoft 365, Windows 7 - 11
' =====================================================================
Option Explicit

Public Sub TugmalarniOrnatish()
    Dim cb As CommandBar, sv As Boolean
    sv = NormalTemplate.Saved
    CustomizationContext = NormalTemplate
    On Error Resume Next
    CommandBars("Ozbek alifbo").Delete
    On Error GoTo 0
    Set cb = CommandBars.Add(Name:="Ozbek alifbo", Position:=msoBarTop, Temporary:=True)
    Tugma cb, "Yangi alifboga o'girish", "YangiAlifbogaOgirish", 2059
    Tugma cb, "Kitobcha", "Kitobcha", 4
    cb.Visible = True
    NormalTemplate.Saved = sv
End Sub

' Word ishga tushganda (STARTUP papkasidagi qo'shimcha sifatida) tugmalarni qo'shadi
Public Sub AutoExec()
    On Error Resume Next
    TugmalarniOrnatish
End Sub

Private Sub Tugma(cb As CommandBar, nom As String, mac As String, ico As Long)
    With cb.Controls.Add(Type:=msoControlButton)
        .Caption = nom
        .OnAction = mac
        .FaceId = ico
        .Style = msoButtonIconAndCaption
    End With
End Sub

' ---------------------------------------------------------------------
'  1. YANGI ALIFBOGA O'GIRISH
'  Belgilangan matn bo'lsa - faqat o'sha qism, bo'lmasa - butun hujjat.
'  Format saqlanadi. Hammasi bitta Ctrl+Z bilan qaytariladi.
'  Inglizcha so'zlar, manzillar (www, @) va ingliz tili belgilangan matn o'zgarmaydi.
' ---------------------------------------------------------------------
Public Sub YangiAlifbogaOgirish()
    Dim st As Range, ur As Object
    If Documents.Count = 0 Then Exit Sub
    On Error Resume Next
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Yangi alifboga o'girish"
    On Error GoTo 0
    Application.ScreenUpdating = False
    If Selection.Type = wdSelectionNormal And Len(Selection.Range.Text) > 1 Then
        OgirRange Selection.Range
    Else
        For Each st In ActiveDocument.StoryRanges
            Do While Not st Is Nothing
                OgirRange st
                Set st = st.NextStoryRange
            Loop
        Next st
    End If
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ur Is Nothing Then ur.EndCustomRecord
    On Error GoTo 0
    Application.StatusBar = "Yangi alifbo: tayyor. Qaytarish - Ctrl+Z"
End Sub

Private Sub OgirRange(rng As Range)
    EskiLotin rng
    Kirill rng
End Sub

' Hammasini birdan almashtirish (kirill uchun)
Private Sub R(rng As Range, f As String, t As String, w As Boolean)
    Dim x As Range
    Set x = rng.Duplicate
    With x.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = f
        .Replacement.Text = t
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = w
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' Birma-bir almashtirish, himoyalangan so'zlarni o'tkazib yuborib (eski lotin uchun)
Private Sub Almash(rng As Range, f As String, t As String, w As Boolean, bosh As Long, uz As Long)
    Dim x As Range, y As Range
    Set x = rng.Duplicate
    Do While x.Find.Execute(FindText:=f, MatchCase:=True, MatchWholeWord:=False, _
            MatchWildcards:=w, MatchSoundsLike:=False, MatchAllWordForms:=False, _
            Forward:=True, Wrap:=wdFindStop, Format:=False)
        If x.Start >= rng.End Or x.End > rng.End Then Exit Do
        If SozHimoyada(x) Then
            x.Collapse wdCollapseEnd
        Else
            Set y = x.Duplicate
            If uz > 0 Then y.SetRange x.Start + bosh, x.Start + bosh + uz
            y.Text = t
            x.SetRange y.End, y.End
        End If
    Loop
End Sub

Private Function SozHimoyada(x As Range) As Boolean
    Dim t As Range, sep As String
    Select Case x.LanguageID
        Case 1033, 2057, 3081, 4105, 5129, 6153
            SozHimoyada = True: Exit Function
    End Select
    sep = " " & vbTab & vbCr & Chr(11) & Chr(12) & ChrW(160) & "(" & ")" & """"
    Set t = x.Duplicate
    t.MoveStartUntil Cset:=sep, Count:=wdBackward
    t.MoveEndUntil Cset:=sep, Count:=wdForward
    SozHimoyada = Himoya(t.Text)
End Function

__ESKI__
__KIRILL__
' ---------------------------------------------------------------------
'  2. KITOBCHA
'  Asl hujjat o'zgarmaydi: kitobcha uchun alohida nusxa ochiladi.
'  A4 varaq o'rtasidan buklanadi, sahifalar A5, sahifa raqamlari qo'yiladi.
' ---------------------------------------------------------------------
Public Sub Kitobcha()
    Dim src As Document, nd As Document, s As Section, j As VbMsgBoxResult
    If Documents.Count = 0 Then Exit Sub
    Set src = ActiveDocument
    Application.ScreenUpdating = False
    If src.Path <> "" And src.Saved Then
        Set nd = Documents.Add(Template:=src.FullName)
    Else
        Set nd = Documents.Add
        nd.Content.FormattedText = src.Content.FormattedText
    End If
    For Each s In nd.Sections
        With s.PageSetup
            .PaperSize = wdPaperA4
            .Orientation = wdOrientLandscape
            .BookFoldPrinting = True
            .TopMargin = CentimetersToPoints(1.5)
            .BottomMargin = CentimetersToPoints(1.5)
            .LeftMargin = CentimetersToPoints(1.8)
            .RightMargin = CentimetersToPoints(1.2)
            .Gutter = 0
        End With
    Next s
    With nd.Sections(1).Footers(wdHeaderFooterPrimary)
        If .PageNumbers.Count = 0 And Len(.Range.Text) <= 1 Then
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
        End If
    End With
    Application.ScreenUpdating = True
    nd.Activate
    j = MsgBox("Kitobcha nusxasi tayyor. Asl hujjat o'zgarmadi." & vbCr & vbCr & _
        "Hozir chop etilsinmi?" & vbCr & vbCr & _
        "HA - printer qog'ozning ikki tomoniga o'zi chiqaradi" & vbCr & _
        "     (sozlamada: qisqa tomon bo'yicha aylantirish)" & vbCr & _
        "YO'Q - printer bir tomonlama: avval bir tomoni chiqadi," & vbCr & _
        "     keyin Word qog'ozni qaytarib qo'yishni so'raydi" & vbCr & _
        "BEKOR - keyinroq chop etaman", vbYesNoCancel + vbQuestion, "Kitobcha")
    If j = vbYes Then
        Dialogs(wdDialogFilePrint).Show
    ElseIf j = vbNo Then
        nd.PrintOut ManualDuplexPrint:=True
    End If
End Sub
''' + COMMON

WORD = WORD.replace("__ESKI__", ESKI_WORD).replace("__KIRILL__", KIRILL_WORD)

# ---------------- EXCEL ----------------
EXCEL = r'''Attribute VB_Name = "OzbekAlifboExcel"
' =====================================================================
'  Ozbek lotin 2026 - Excel uchun yordamchi (v2)
'  Tugma: "Yangi alifboga o'girish"   Formula: =YANGIALIFBO(A1)
'  Excel 2010 - 2021 va Microsoft 365, Windows 7 - 11
'  PERSONAL.XLSB ga import qiling - barcha kitoblarda ishlaydi
' =====================================================================
Option Explicit

Private eski As Collection

' Excel ochilganda tugmani avtomatik qo'shadi
Public Sub Auto_Open()
    Dim cb As CommandBar
    On Error Resume Next
    Application.CommandBars("Ozbek alifbo").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add(Name:="Ozbek alifbo", Position:=msoBarTop, Temporary:=True)
    With cb.Controls.Add(Type:=msoControlButton)
        .Caption = "Yangi alifboga o'girish"
        .OnAction = "'" & ThisWorkbook.Name & "'!YangiAlifbogaOgirishExcel"
        .FaceId = 2059
        .Style = msoButtonIconAndCaption
    End With
    cb.Visible = True
End Sub

' Belgilangan kataklar (bitta katak tanlansa - butun varaq) yangi alifboga o'giriladi.
' Formulalar o'zgarmaydi. Qaytarish: Ctrl+Z
Public Sub YangiAlifbogaOgirishExcel()
    Dim rng As Range, c As Range, v As String, nv As String
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
        MsgBox "O'giriladigan matn topilmadi.", vbInformation, "Yangi alifbo"
        Exit Sub
    End If
    On Error GoTo 0
    Set eski = New Collection
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        v = CStr(c.Value)
        nv = YANGIALIFBO(v)
        If nv <> v Then
            eski.Add Array(c.Worksheet.Parent.Name, c.Worksheet.Name, c.Address, c.Formula)
            c.Value = nv
        End If
    Next c
    Application.ScreenUpdating = True
    If eski.Count > 0 Then
        Application.OnUndo "Yangi alifboga o'girishni qaytarish", "'" & ThisWorkbook.Name & "'!QaytarishExcel"
    End If
    Application.StatusBar = "Yangi alifbo: " & eski.Count & " ta katak o'girildi. Qaytarish - Ctrl+Z"
End Sub

Public Sub QaytarishExcel()
    Dim a As Variant
    If eski Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For Each a In eski
        Workbooks(a(0)).Worksheets(a(1)).Range(a(2)).Formula = a(3)
    Next a
    Application.ScreenUpdating = True
    Set eski = Nothing
    Application.StatusBar = False
End Sub
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

for name, src in [("OzbekAlifbo_Word.bas", WORD), ("OzbekAlifbo_Excel.bas", EXCEL)]:
    a = to_ascii(src)
    for ln in a.split("\n"):
        assert len(ln) < 1000, ln[:80]
    open(name, "w", encoding="ascii", newline="\r\n").write(a)
    print(name, len(a))

Attribute VB_Name = "OzbekYozuv"
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
    Tugma cb, "Lotin 2026", "Lotin2026ga", 2059, "Matnni lotin 2026 alifbosiga o'girish (" & ChrW(&H00D6) & ", " & ChrW(&H011E) & ", " & ChrW(&H015E) & ", " & ChrW(&H00C7) & ")"
    Tugma cb, "Lotin 1995", "Lotin1995ga", 2059, "Matnni lotin 1995 alifbosiga o'girish (O" & ChrW(&H02BB) & ", G" & ChrW(&H02BB) & ", Sh, Ch)"
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

' ===== Satr konverteri: kirill va eski lotin -> yangi lotin (2026) =====
Private Function IsSep(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 32, 9, 10, 11, 13, 160, 7, 12, 1, 14, 30, 31: IsSep = True
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
    s = Replace(s, ChrW(&H015F), ChrW(&H015E)): s = Replace(s, ChrW(&H00E7), ChrW(&H00C7))
    s = Replace(s, ChrW(&H00F6), ChrW(&H00D6)): s = Replace(s, ChrW(&H011F), ChrW(&H011E))
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
        Case &H430: SimpleMap = "a"
        Case &H431: SimpleMap = "b"
        Case &H432: SimpleMap = "v"
        Case &H433: SimpleMap = "g"
        Case &H493: SimpleMap = ChrW(&H011F)
        Case &H434: SimpleMap = "d"
        Case &H436: SimpleMap = "j"
        Case &H437: SimpleMap = "z"
        Case &H438: SimpleMap = "i"
        Case &H439: SimpleMap = "y"
        Case &H43A: SimpleMap = "k"
        Case &H49B: SimpleMap = "q"
        Case &H43B: SimpleMap = "l"
        Case &H43C: SimpleMap = "m"
        Case &H43D: SimpleMap = "n"
        Case &H43E: SimpleMap = "o"
        Case &H43F: SimpleMap = "p"
        Case &H440: SimpleMap = "r"
        Case &H441: SimpleMap = "s"
        Case &H442: SimpleMap = "t"
        Case &H443: SimpleMap = "u"
        Case &H45E: SimpleMap = ChrW(&H00F6)
        Case &H444: SimpleMap = "f"
        Case &H445: SimpleMap = "x"
        Case &H4B3: SimpleMap = "h"
        Case &H447: SimpleMap = ChrW(&H00E7)
        Case &H448: SimpleMap = ChrW(&H015F)
        Case &H449: SimpleMap = ChrW(&H015F)
        Case &H44B: SimpleMap = "i"
        Case &H44D: SimpleMap = "e"
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
            Case &H44A: o = ChrW(&H02BC)
            Case &H44C
                If LowK(AscW(nx & " ")) = &H43E Then o = Cap("y", up, nxUp) Else o = ""
            Case &H441
                If LowK(AscW(nx & " ")) = &H4B3 Then o = Cap("s", up, nxUp) & ChrW(&H02BC) Else o = Cap("s", up, nxUp)
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
        If InStr("|chrome|chat|check|chip|cheese|chess|chicken|chief|child|choice|church|chart|channel|chapter|shop|show|shift|share|shell|shot|ship|shirt|shoe|short|should|shut|sheet|shelf|shine|flash|push|cash|crash|fresh|english|wash|finish|match|search|touch|teacher|school|machine|mechanic|technology|orchestra|chemistry|character|schedule|shampoo|chips|ketchup|", "|" & w & "|") > 0 Then Himoya = True
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
                Case "o": res = res & ChrW(&H00F6)
                Case "O": res = res & ChrW(&H00D6)
                Case "g": res = res & ChrW(&H011F)
                Case "G": res = res & ChrW(&H011E)
            End Select
            i = i + 2
        ElseIf (nx = "h" Or nx = "H") And (c = "s" Or c = "S" Or c = "c" Or c = "C") Then
            Select Case c
                Case "s": res = res & ChrW(&H015F)
                Case "S": res = res & ChrW(&H015E)
                Case "c": res = res & ChrW(&H00E7)
                Case "C": res = res & ChrW(&H00C7)
            End Select
            i = i + 2
        ElseIf IsAp(c) And i > 1 And i < n Then
            If IsLet(Mid$(t, i - 1, 1)) And IsLet(nx) Then res = res & ChrW(&H02BC) Else res = res & c
            i = i + 1
        Else
            res = res & c
            i = i + 1
        End If
    Loop
    EskiToken = res
End Function



' ===== Lotin 2026 -> Lotin 1995 / Kirill =====
Private Function IsLatVowel(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsLatVowel = (InStr("aeiouAEIOU" & ChrW(&H00F6) & ChrW(&H00D6), ch) > 0)
End Function

Private Function LowX(ByVal ch As String) As String
    Select Case ch
        Case ChrW(&H00D6): LowX = ChrW(&H00F6)
        Case ChrW(&H011E): LowX = ChrW(&H011F)
        Case ChrW(&H015E): LowX = ChrW(&H015F)
        Case ChrW(&H00C7): LowX = ChrW(&H00E7)
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
            Case ChrW(&H00F6): res = res & "o" & ChrW(&H02BB)
            Case ChrW(&H00D6): res = res & "O" & ChrW(&H02BB)
            Case ChrW(&H011F): res = res & "g" & ChrW(&H02BB)
            Case ChrW(&H011E): res = res & "G" & ChrW(&H02BB)
            Case ChrW(&H015F): res = res & "sh"
            Case ChrW(&H015E): If kat Then res = res & "SH" Else res = res & "Sh"
            Case ChrW(&H00E7): res = res & "ch"
            Case ChrW(&H00C7): If kat Then res = res & "CH" Else res = res & "Ch"
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
        Case "a": LatCyr = &H430
        Case "b": LatCyr = &H431
        Case "d": LatCyr = &H434
        Case "f": LatCyr = &H444
        Case "g": LatCyr = &H433
        Case "h": LatCyr = &H4B3
        Case "i": LatCyr = &H438
        Case "j": LatCyr = &H436
        Case "k": LatCyr = &H43A
        Case "l": LatCyr = &H43B
        Case "m": LatCyr = &H43C
        Case "n": LatCyr = &H43D
        Case "o": LatCyr = &H43E
        Case "p": LatCyr = &H43F
        Case "q": LatCyr = &H49B
        Case "r": LatCyr = &H440
        Case "s": LatCyr = &H441
        Case "t": LatCyr = &H442
        Case "u": LatCyr = &H443
        Case "v": LatCyr = &H432
        Case "x": LatCyr = &H445
        Case "y": LatCyr = &H439
        Case "z": LatCyr = &H437
        Case ChrW(&H00F6): LatCyr = &H45E
        Case ChrW(&H011F): LatCyr = &H493
        Case ChrW(&H015F): LatCyr = &H448
        Case ChrW(&H00E7): LatCyr = &H447
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
        ElseIf lc = "s" And nx = ChrW(&H02BC) And LowX(n2) = "h" Then
            res = res & Cy(&H441, up) & Cy(&H4B3, IsUp(n2)): i = i + 3
        ElseIf lc = "e" Then
            If Not IsLet(pv) Or IsLatVowel(pv) Then k = &H44D Else k = &H435
            res = res & Cy(k, up): i = i + 1
        ElseIf c = ChrW(&H02BC) Then
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

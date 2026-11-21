' =====================================================================
'  OzbekYozuv 1.0 - O'RNATUVCHI
'  Word va Excel tugmalari (Lotin 2026 / Lotin 1995 / Kirill, Kitobcha),
'  konverter oynasi va UZB klaviaturasi.
'  Administrator huquqi va internet kerak emas. Windows 7, 8, 10, 11.
'  Ishlatish: Word va Excel'ni yoping, shu faylga ikki marta bosing.
' =====================================================================
Option Explicit

Const SARLAVHA = "OzbekYozuv 1.0"

Dim sh, fso, papka, manba, ad, natija, jurnal
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
papka = fso.GetParentFolderName(WScript.ScriptFullName)
manba = papka & "\fayllar"
ad = sh.ExpandEnvironmentStrings("%APPDATA%") & "\OzbekYozuv"
jurnal = sh.ExpandEnvironmentStrings("%TEMP%") & "\OzbekYozuv_ornatish.txt"
natija = ""

' --- 1. Tekshiruvlar -------------------------------------------------
Dim kerakli, f
kerakli = Array("OzbekYozuv_Word.bas", "OzbekYozuv_Excel.bas", "uzb_klaviatura.ps1", "uzb_ishga.vbs", "OzbekYozuv_Konverter.hta")
For Each f In kerakli
    If Not fso.FileExists(manba & "\" & f) Then
        MsgBox "Kerakli fayl topilmadi: fayllar\" & f & vbCrLf & vbCrLf & _
            "Arxivni avval papkaga chiqaring (o'ng tugma -> Izvlech vse)," & vbCrLf & _
            "keyin papka ichidagi ORNATISH.vbs ga bosing.", vbExclamation, SARLAVHA
        WScript.Quit
    End If
Next

If Ishlayapti("Word.Application") Or Ishlayapti("Excel.Application") Then
    MsgBox "Avval Word va Excel'ni to'liq yoping (ishingizni saqlab qo'ying)," & vbCrLf & _
        "keyin ORNATISH.vbs ga qayta bosing.", vbExclamation, SARLAVHA
    WScript.Quit
End If

If MsgBox("Quyidagilar o'rnatiladi (administrator paroli kerak emas):" & vbCrLf & vbCrLf & _
    "  1. Word: Lotin 2026, Lotin 1995, Kirill, Kitobcha tugmalari" & vbCrLf & _
    "  2. Excel: Lotin 2026, Lotin 1995, Kirill tugmalari va formulalari" & vbCrLf & _
    "  3. Ish stolida: OzbekYozuv_Konverter oynasi" & vbCrLf & _
    "  4. UZB klaviaturasi" & vbCrLf & vbCrLf & _
    "Davom etilsinmi?", vbOKCancel + vbQuestion, SARLAVHA) <> vbOK Then WScript.Quit

' --- 2. Office makros loyihasiga vaqtincha ruxsat ---------------------
Dim kalitlar(7), eski(7), versiyalar, ilovalar, ver, ilova, idx
versiyalar = Array("12.0", "14.0", "15.0", "16.0")
ilovalar = Array("Word", "Excel")
idx = 0
On Error Resume Next
For Each ver In versiyalar
    For Each ilova In ilovalar
        kalitlar(idx) = "HKCU\Software\Microsoft\Office\" & ver & "\" & ilova & "\Security\AccessVBOM"
        eski(idx) = sh.RegRead(kalitlar(idx))
        If Err.Number <> 0 Then eski(idx) = Empty
        Err.Clear
        sh.RegWrite kalitlar(idx), 1, "REG_DWORD"
        Err.Clear
        idx = idx + 1
    Next
Next
On Error GoTo 0

' --- 3. O'rnatish ----------------------------------------------------
Yoz "Word:       " & WordOrnat()
Yoz "Excel:      " & ExcelOrnat()
Yoz "Konverter:  " & KonverterOrnat()
Yoz "Klaviatura: " & KlaviaturaOrnat()

' --- 4. Ruxsatni avvalgi holiga qaytarish -----------------------------
On Error Resume Next
For idx = 0 To 7
    If IsEmpty(eski(idx)) Then
        sh.RegDelete kalitlar(idx)
    Else
        sh.RegWrite kalitlar(idx), eski(idx), "REG_DWORD"
    End If
    Err.Clear
Next
On Error GoTo 0

JurnalYoz
MsgBox natija & vbCrLf & _
    "Word va Excel: tugmalar lentadagi NADSTROYKI (Add-ins) yorlig'ida." & vbCrLf & _
    "Excel formulalari: =LOTIN2026(A1)  =LOTIN1995(A1)  =KIRILL(A1)" & vbCrLf & _
    "Klaviatura: Alt + Shift bilan UZB ni tanlang." & vbCrLf & vbCrLf & _
    "Batafsil: " & jurnal, vbInformation, SARLAVHA
WScript.Quit

' =====================================================================
Sub Yoz(t)
    natija = natija & t & vbCrLf
End Sub

Sub JurnalYoz()
    Dim j
    On Error Resume Next
    Set j = fso.CreateTextFile(jurnal, True)
    j.Write SARLAVHA & " - " & Now & vbCrLf & natija
    j.Close
    On Error GoTo 0
End Sub

Function Ishlayapti(prog)
    Dim o
    On Error Resume Next
    Set o = GetObject(, prog)
    Ishlayapti = (Err.Number = 0)
    Err.Clear
End Function

Sub PapkaYarat(p)
    On Error Resume Next
    If Not fso.FolderExists(fso.GetParentFolderName(p)) Then PapkaYarat fso.GetParentFolderName(p)
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    Err.Clear
End Sub

' Nomi shu so'z bilan boshlanadigan barcha modullarni olib tashlaydi (OzbekAlifbo, OzbekAlifbo1 ...)
Function KomponentOchir(vbp, nom)
    Dim c, royxat, i
    KomponentOchir = False
    royxat = ""
    For Each c In vbp.VBComponents
        If c.Type = 1 And Left(LCase(c.Name), Len(nom)) = LCase(nom) Then royxat = royxat & c.Name & "|"
    Next
    If Len(royxat) = 0 Then Exit Function
    royxat = Split(Left(royxat, Len(royxat) - 1), "|")
    For i = 0 To UBound(royxat)
        vbp.VBComponents.Remove vbp.VBComponents(royxat(i))
        KomponentOchir = True
    Next
End Function

' --- Word: STARTUP papkasiga OzbekYozuv.dotm -------------------------
Function WordOrnat()
    Dim w, vbp, d, sp, nom, a, o
    On Error Resume Next
    Set w = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        WordOrnat = "Word topilmadi - o'tkazib yuborildi"
        Err.Clear: Exit Function
    End If
    w.DisplayAlerts = 0

    ' Eski versiyalar qoldiqlarini tozalash (Normal shablonidagi modul va tugmalar)
    Set vbp = w.NormalTemplate.VBProject
    o = False
    If Err.Number = 0 Then
        If KomponentOchir(vbp, "OzbekAlifbo") Then o = True
        If KomponentOchir(vbp, "OzbekYozuv") Then o = True
    End If
    Err.Clear
    w.CustomizationContext = w.NormalTemplate
    w.CommandBars("Ozbek alifbo").Delete
    Err.Clear
    w.CommandBars("Ozbek yozuvlari").Delete
    Err.Clear
    ' Modul olib tashlanganda Word "Saved" belgisini o'zgartirmasligi mumkin - shuning uchun majburan saqlaymiz
    If o Or Not w.NormalTemplate.Saved Then w.NormalTemplate.Save
    Err.Clear

    sp = w.StartupPath
    If Len(sp) = 0 Then sp = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word\STARTUP"
    PapkaYarat sp
    For Each a In w.AddIns
        nom = LCase(a.Name)
        If nom = "ozbekalifbo.dotm" Or nom = "ozbekyozuv.dotm" Then a.Installed = False
    Next
    Err.Clear
    For Each nom In Array("OzbekAlifbo.dotm", "OzbekYozuv.dotm")
        If fso.FileExists(sp & "\" & nom) Then fso.DeleteFile sp & "\" & nom, True
    Next
    If fso.FileExists(sp & "\OzbekYozuv.dotm") Then
        WordOrnat = "XATO - eski OzbekYozuv.dotm faylini almashtirib bo'lmadi"
        Err.Clear: w.Quit 0: Exit Function
    End If
    Err.Clear

    Set d = w.Documents.Add
    Set vbp = d.VBProject
    o = vbp.VBComponents.Count
    If Err.Number <> 0 Then
        WordOrnat = "XATO - Word makroslarga ruxsat bermadi (" & Err.Description & ")"
        Err.Clear: d.Close 0: w.Quit 0: Exit Function
    End If
    vbp.VBComponents.Import manba & "\OzbekYozuv_Word.bas"
    If Err.Number <> 0 Then
        WordOrnat = "XATO - modul qo'shilmadi (" & Err.Description & ")"
        Err.Clear: d.Close 0: w.Quit 0: Exit Function
    End If
    d.SaveAs sp & "\OzbekYozuv.dotm", 15
    If Err.Number <> 0 Then
        WordOrnat = "XATO - saqlab bo'lmadi (" & Err.Description & ")"
    Else
        WordOrnat = "o'rnatildi"
    End If
    Err.Clear
    d.Close 0
    w.Quit 0
    Err.Clear
End Function

' --- Excel: XLSTART papkasiga OzbekYozuv.xlam ------------------------
Function ExcelOrnat()
    Dim x, wb, xs, vbp, o, kerak
    On Error Resume Next
    Set x = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        ExcelOrnat = "Excel topilmadi - o'tkazib yuborildi"
        Err.Clear: Exit Function
    End If
    x.DisplayAlerts = False
    x.EnableEvents = False
    xs = x.StartupPath
    If Len(xs) = 0 Then xs = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Excel\XLSTART"
    PapkaYarat xs

    ' Eski versiyalar: PERSONAL.XLSB ichidagi modulni olib tashlash
    If fso.FileExists(xs & "\PERSONAL.XLSB") Then
        Set wb = x.Workbooks.Open(xs & "\PERSONAL.XLSB")
        If Err.Number = 0 Then
            Set vbp = wb.VBProject
            kerak = False
            If Err.Number = 0 Then
                If KomponentOchir(vbp, "OzbekAlifboExcel") Then kerak = True
                If KomponentOchir(vbp, "OzbekYozuvExcel") Then kerak = True
            End If
            If kerak Then wb.Save
            wb.Close False
        End If
        Err.Clear
    End If
    If fso.FileExists(xs & "\OzbekYozuv.xlam") Then fso.DeleteFile xs & "\OzbekYozuv.xlam", True
    Err.Clear

    Set wb = x.Workbooks.Add
    Set vbp = wb.VBProject
    o = vbp.VBComponents.Count
    If Err.Number <> 0 Then
        ExcelOrnat = "XATO - Excel makroslarga ruxsat bermadi (" & Err.Description & ")"
        Err.Clear: wb.Close False: x.Quit: Exit Function
    End If
    vbp.VBComponents.Import manba & "\OzbekYozuv_Excel.bas"
    If Err.Number <> 0 Then
        ExcelOrnat = "XATO - modul qo'shilmadi (" & Err.Description & ")"
        Err.Clear: wb.Close False: x.Quit: Exit Function
    End If
    wb.SaveAs xs & "\OzbekYozuv.xlam", 55
    If Err.Number <> 0 Then
        ExcelOrnat = "XATO - saqlab bo'lmadi (" & Err.Description & ")"
    Else
        ExcelOrnat = "o'rnatildi"
    End If
    Err.Clear
    wb.Close False
    x.Quit
    Err.Clear
End Function

' --- Konverter: ish stoliga ------------------------------------------
Function KonverterOrnat()
    Dim ist
    On Error Resume Next
    ist = sh.SpecialFolders("Desktop")
    fso.CopyFile manba & "\OzbekYozuv_Konverter.hta", ist & "\", True
    If Err.Number = 0 Then
        KonverterOrnat = "ish stoliga qo'yildi (OzbekYozuv_Konverter)"
    Else
        KonverterOrnat = "XATO - ish stoliga qo'yib bo'lmadi (" & Err.Description & ")"
    End If
    Err.Clear
End Function

' --- UZB klaviaturasi (faqat shu foydalanuvchi uchun) ----------------
Function KlaviaturaOrnat()
    Dim wmi, pr, winVer, win7, k, pv, bor, rc, cmd
    On Error Resume Next

    ' Eski nusxa ishlayotgan bo'lsa to'xtatish
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each pr In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
        If InStr(LCase(pr.CommandLine), "uzb_klaviatura") > 0 Then pr.Terminate
    Next
    Err.Clear

    PapkaYarat ad
    fso.CopyFile manba & "\uzb_klaviatura.ps1", ad & "\", True
    fso.CopyFile manba & "\uzb_ishga.vbs", ad & "\", True
    If Err.Number <> 0 Then
        KlaviaturaOrnat = "XATO - fayllarni nusxalab bo'lmadi (" & Err.Description & ")"
        Err.Clear: Exit Function
    End If

    ' Windows ochilganda avtomatik ishga tushish
    sh.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\OzbekUZB", "wscript.exe """ & ad & "\uzb_ishga.vbs""", "REG_SZ"
    Err.Clear

    ' Til paneliga UZB (o'zbek lotin) qo'shish
    winVer = sh.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    Err.Clear
    win7 = (winVer = "6.1" Or winVer = "6.0")
    If win7 Then
        bor = False
        For k = 1 To 50
            pv = sh.RegRead("HKCU\Keyboard Layout\Preload\" & k)
            If Err.Number <> 0 Then
                Err.Clear
                If Not bor Then sh.RegWrite "HKCU\Keyboard Layout\Preload\" & k, "00000443", "REG_SZ"
                Exit For
            End If
            If LCase(pv) = "00000443" Then bor = True
        Next
        sh.RegWrite "HKCU\Keyboard Layout\Substitutes\00000443", "00000409", "REG_SZ"
        If Err.Number = 0 Then rc = 0 Else rc = 1
        Err.Clear
    Else
        cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""$l=Get-WinUserLanguageList; " & _
              "$u=$l | Where-Object {$_.LanguageTag -like 'uz-Latn*'} | Select-Object -First 1; " & _
              "if(-not $u){$l.Add('uz-Latn-UZ'); $u=$l | Where-Object {$_.LanguageTag -like 'uz-Latn*'} | Select-Object -First 1}; " & _
              "$u.InputMethodTips.Clear(); $u.InputMethodTips.Add('0443:00000409'); Set-WinUserLanguageList $l -Force"""
        rc = sh.Run(cmd, 0, True)
        If Err.Number <> 0 Then rc = 1
        Err.Clear
    End If

    ' Hozir ishga tushirish
    sh.Run "wscript.exe """ & ad & "\uzb_ishga.vbs""", 0, False
    Err.Clear

    If rc <> 0 Then
        KlaviaturaOrnat = "dastur o'rnatildi, lekin UZB tilini qo'shib bo'lmadi"
    ElseIf win7 Then
        KlaviaturaOrnat = "o'rnatildi (Windows'dan chiqib qayta kiring)"
    Else
        KlaviaturaOrnat = "o'rnatildi"
    End If
End Function

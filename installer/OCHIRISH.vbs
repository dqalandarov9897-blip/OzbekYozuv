' =====================================================================
'  OzbekYozuv 1.0 - OLIB TASHLASH
'  Word/Excel tugmalari, konverter va UZB klaviatura dasturini o'chiradi.
'  Ishlatish: Word va Excel'ni yoping, shu faylga ikki marta bosing.
' =====================================================================
Option Explicit

Const SARLAVHA = "OzbekYozuv 1.0 - olib tashlash"

Dim sh, fso, ad, natija, ap
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ap = sh.ExpandEnvironmentStrings("%APPDATA%")
ad = ap & "\OzbekYozuv"
natija = ""

If Ishlayapti("Word.Application") Or Ishlayapti("Excel.Application") Then
    MsgBox "Avval Word va Excel'ni to'liq yoping, keyin qayta urinib ko'ring.", vbExclamation, SARLAVHA
    WScript.Quit
End If

If MsgBox("OzbekYozuv olib tashlansinmi?" & vbCrLf & vbCrLf & _
    "  - Word va Excel tugmalari" & vbCrLf & _
    "  - ish stolidagi konverter" & vbCrLf & _
    "  - UZB klaviatura dasturi", vbOKCancel + vbQuestion, SARLAVHA) <> vbOK Then WScript.Quit

Dim wmi, pr, nom
On Error Resume Next

' Klaviatura dasturini to'xtatish va avtoyuklashdan olib tashlash
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
For Each pr In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    If InStr(LCase(pr.CommandLine), "uzb_klaviatura") > 0 Then pr.Terminate
Next
Err.Clear
sh.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\OzbekUZB"
Err.Clear
If fso.FolderExists(ad) Then fso.DeleteFolder ad, True
Qayd "Klaviatura dasturi"

' Word qo'shimchasi
For Each nom In Array("OzbekYozuv.dotm", "OzbekAlifbo.dotm")
    If fso.FileExists(ap & "\Microsoft\Word\STARTUP\" & nom) Then fso.DeleteFile ap & "\Microsoft\Word\STARTUP\" & nom, True
Next
Qayd "Word tugmalari"

' Excel qo'shimchasi
If fso.FileExists(ap & "\Microsoft\Excel\XLSTART\OzbekYozuv.xlam") Then fso.DeleteFile ap & "\Microsoft\Excel\XLSTART\OzbekYozuv.xlam", True
Qayd "Excel tugmalari"

' Konverter
If fso.FileExists(sh.SpecialFolders("Desktop") & "\OzbekYozuv_Konverter.hta") Then fso.DeleteFile sh.SpecialFolders("Desktop") & "\OzbekYozuv_Konverter.hta", True
Qayd "Konverter"
On Error GoTo 0

' UZB tilini til panelidan olib tashlash (ixtiyoriy)
If MsgBox("Til panelidagi UZB tilini ham olib tashlaymi?", vbYesNo + vbQuestion, SARLAVHA) = vbYes Then
    Dim winVer, k, pv, rc
    On Error Resume Next
    winVer = sh.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    Err.Clear
    If winVer = "6.1" Or winVer = "6.0" Then
        For k = 1 To 50
            pv = sh.RegRead("HKCU\Keyboard Layout\Preload\" & k)
            If Err.Number <> 0 Then Err.Clear: Exit For
            If LCase(pv) = "00000443" Then sh.RegDelete "HKCU\Keyboard Layout\Preload\" & k
        Next
        sh.RegDelete "HKCU\Keyboard Layout\Substitutes\00000443"
        Err.Clear
        natija = natija & "UZB tili: olib tashlandi (Windows'dan chiqib qayta kiring)" & vbCrLf
    Else
        rc = sh.Run("powershell -NoProfile -ExecutionPolicy Bypass -Command ""$l=Get-WinUserLanguageList; " & _
            "$u=@($l | Where-Object {$_.LanguageTag -like 'uz-Latn*'}); foreach($x in $u){[void]$l.Remove($x)}; " & _
            "if($l.Count -gt 0){Set-WinUserLanguageList $l -Force}""", 0, True)
        If Err.Number = 0 And rc = 0 Then
            natija = natija & "UZB tili: olib tashlandi" & vbCrLf
        Else
            natija = natija & "UZB tili: olib tashlab bo'lmadi - Parametry -> Vremya i yazyk -> Yazyk" & vbCrLf
        End If
        Err.Clear
    End If
    On Error GoTo 0
End If

MsgBox natija, vbInformation, SARLAVHA
WScript.Quit

Sub Qayd(t)
    If Err.Number = 0 Then
        natija = natija & t & ": olib tashlandi" & vbCrLf
    Else
        natija = natija & t & ": XATO - " & Err.Description & vbCrLf
    End If
    Err.Clear
End Sub

Function Ishlayapti(prog)
    Dim o
    On Error Resume Next
    Set o = GetObject(, prog)
    Ishlayapti = (Err.Number = 0)
    Err.Clear
End Function

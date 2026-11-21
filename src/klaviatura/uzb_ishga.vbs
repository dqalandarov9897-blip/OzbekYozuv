' OzbekYozuv 1.0 - UZB klaviatura dasturini yashirin ishga tushiradi (Windows ochilganda)
Dim sh
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""& { iex ([IO.File]::ReadAllText($env:APPDATA + '\OzbekYozuv\uzb_klaviatura.ps1')) }""", 0, False

"""Release zip fayllarini yasaydi: python tools/paket.py  ->  dist/"""
import os, shutil, zipfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V = "1.0"
J = lambda *p: os.path.join(ROOT, *p)
DIST = J("dist")
shutil.rmtree(DIST, ignore_errors=True)

def zipla(nom, fayllar):
    z = os.path.join(DIST, nom + ".zip")
    os.makedirs(DIST, exist_ok=True)
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
        for manba, ichki in fayllar:
            zf.write(manba, nom + "/" + ichki)
    print(z, os.path.getsize(z), "bayt")

zipla("OzbekYozuv-" + V, [
    (J("installer", "ORNATISH.vbs"), "ORNATISH.vbs"),
    (J("installer", "OCHIRISH.vbs"), "OCHIRISH.vbs"),
    (J("installer", "OQING.txt"), "OQING.txt"),
    (J("installer", "README.md"), "README.md"),
    (J("LICENSE"), "LICENSE.txt"),
    (J("docs", "Yoriqnoma_A5.pdf"), "Yoriqnoma_A5.pdf"),
    (J("src", "word", "OzbekYozuv_Word.bas"), "fayllar/OzbekYozuv_Word.bas"),
    (J("src", "excel", "OzbekYozuv_Excel.bas"), "fayllar/OzbekYozuv_Excel.bas"),
    (J("src", "klaviatura", "uzb_klaviatura.ps1"), "fayllar/uzb_klaviatura.ps1"),
    (J("src", "klaviatura", "uzb_ishga.vbs"), "fayllar/uzb_ishga.vbs"),
    (J("src", "konverter", "OzbekYozuv_Konverter.hta"), "fayllar/OzbekYozuv_Konverter.hta"),
])
zipla("OzbekYozuv-Konverter-" + V, [
    (J("src", "konverter", "OzbekYozuv_Konverter.hta"), "OzbekYozuv_Konverter.hta"),
    (J("src", "konverter", "OQING.txt"), "OQING.txt"),
    (J("LICENSE"), "LICENSE.txt"),
])

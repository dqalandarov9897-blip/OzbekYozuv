# OzbekYozuv — Lotin 2026 · Lotin 1995 · Kirill

**Özbek tilining uç yozuvi bilan işlaş uçun bepul va oçiq kodli töplam.**
Matnni bir bosişda yangi alifboga (Ö, Ğ, Ş, Ç, ʼ), eski lotinga yoki kirillga öğiradi.

> Ўзбек тилининг уч ёзуви учун бепул конвертер ва клавиатура.
> Free converter and keyboard for the three Uzbek scripts: Latin 2026, Latin 1995 and Cyrillic.

## 🌐 Onlayn konverter

**[Konverterni oçiş →](https://dqalandarov9897-blip.github.io/OzbekYozuv/)** — telefon, planşet va istalgan kompyuterda işlaydi, örnatiş kerak emas. Matningiz heç qayerga yuborilmaydi.

## 💾 Windows uçun yuklab oliş

| Fayl | Kimga |
|---|---|
| [**OzbekYozuv-1.0.zip**](../../releases/latest/download/OzbekYozuv-1.0.zip) | Töliq paket: Word, Excel, UZB klaviaturasi, konverter |
| [**OzbekYozuv-Konverter-1.0.zip**](../../releases/latest/download/OzbekYozuv-Konverter-1.0.zip) | Faqat konverter — örnatiş taqiqlangan iş kompyuterlari uçun |

Barça versiyalar: [Releases](../../releases)

## Imkoniyatlar

| Qism | Nima qiladi |
|---|---|
| **Word** | Lotin 2026 · Lotin 1995 · Kirill tugmalari; **Kitobça** — hujjatni A5 kitobça qilib çop etiş. Ctrl+Z bilan bir marta qaytariladi |
| **Excel** | Tugmalar va formulalar: `=LOTIN2026(A1)`, `=LOTIN1995(A1)`, `=KIRILL(A1)` |
| **UZB klaviaturasi** | W = ş, C = ç, [ = ğ, ] = ö, ` = ʼ — bitta tugmada |
| **Konverter oynasi** | Telegram, brauzer va boşqa dasturlar uçun: Ctrl+C → tugma → Ctrl+V |

- Administrator huquqi **kerak emas** — faqat joriy foydalanuvçi uçun örnatiladi.
- Internet **kerak emas**.
- Windows 7, 8, 10, 11 · Microsoft Office 2010–2021 va Microsoft 365.

## Örnatiş

1. Arxivni papkaga çiqaring (öng tugma → **Извлечь все / Extract all**).
2. Word va Excelni yoping.
3. **ORNATISH.vbs** ga ikki marta bosing → **OK**.
4. Windows 7 bölsa — tizimdan çiqib qayta kiring.

Olib taşlaş: **OCHIRISH.vbs**.

## UZB klaviaturasi

**Alt + Shift** bilan **UZB** ni tanlang.

| Tugma | Oddiy | Shift bilan | Öng Alt bilan |
|---|---|---|---|
| W | ş | Ş | w / W |
| C | ç | Ç | c / C |
| [ | ğ | Ğ | [ / { |
| ] | ö | Ö | ] / } |
| ` | ʼ | ' | ` / ~ |

## Öğiriş qoidalari

- **Kirill → lotin:** е → ye/e, ё → yo, ю → yu, я → ya, ц → s/ts, ъ → ʼ, сҳ → sʼh.
- **Lotin 1995 → 2026:** oʻ (o‘, o', o\`) → ö, gʻ → ğ, sh → ş, ch → ç, apostrof → ʼ.
- **Cheklovsiz:** barça sözlar öğiriladi, jumladan inglizça sözlar ham. Aralaş matnda faqat özbekça qismni tanlang.
- **Lotin → kirill:** «ts» doim «тс» böladi; rusça sözlarni qölda tekşiring.

## Xavfsizlik

- Barça kod oçiq va öqiladigan matn: `.bas`, `.vbs`, `.ps1`, `.hta`.
- Örnatuvçi Office'ning «VBA loyihasiga dastur orqali kiriş» ruxsatini faqat örnatiş vaqtida yoqadi va keyin avvalgi holatiga qaytaradi.
- UZB klaviaturasi — fonda işlaydigan kiçik PowerShell dasturi. U faqat UZB tanlanganda beş tugmani almaştiradi, heç narsani saqlamaydi va heç qayerga yubormaydi. Şu sabab ayrim antiviruslar uni «klaviatura kuzatuvçisi» deb ogohlantirişi mumkin — kodni [src/klaviatura/uzb_klaviatura.ps1](src/klaviatura/uzb_klaviatura.ps1) da tekşiring.

## Loyiha tuzilmasi

```
docs/        onlayn konverter (GitHub Pages) va A5 yöriqnoma
src/word/    Word moduli (VBA)
src/excel/   Excel moduli (VBA)
src/klaviatura/  UZB klaviaturasi (PowerShell)
src/konverter/   konverter (JavaScript + HTA)
installer/   ORNATISH.vbs, OCHIRISH.vbs
src/web/     onlayn sahifa andozasi
tools/       modullar, sahifa va release zip fayllarini yasovçi skriptlar
test/        30 ta öğiriş testi (VBA va JS uçun bir xil)
```

Testlar: `python test/run.py` · Sahifa: `python tools/web.py` · Release: `python tools/paket.py`

## Xato topdingizmi?

[Issues](../../issues) böliminda yozing: qaysi matn, qaysi yozuvga, nima çiqdi va nima çiqişi kerak edi.

## Litsenziya

[MIT](LICENSE) — bepul foydalaniş, özgartiriş va tarqatiş mumkin.
© 2026 Dilshod Qalandarov

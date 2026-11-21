# OzbekYozuv 1.0

O‘zbek tilining uch yozuvi — **Lotin 2026**, **Lotin 1995** va **Kirill** — bilan ishlash uchun bepul to‘plam:

| Qism | Nima qiladi |
|---|---|
| **Word tugmalari** | Hujjatni yoki belgilangan qismni istalgan yozuvga o‘giradi. **Kitobcha** — hujjatni A5 kitobcha qilib chop etadi |
| **Excel tugmalari va formulalari** | Kataklarni o‘giradi; `=LOTIN2026(A1)`, `=LOTIN1995(A1)`, `=KIRILL(A1)` |
| **Konverter oynasi** | Telegram, brauzer va boshqa dasturlar uchun: Ctrl+C → tugma → Ctrl+V |
| **UZB klaviaturasi** | Yangi harflar bitta tugmada: W = ş, C = ç, [ = ğ, ] = ö, ` = ʼ |

- Administrator huquqi **kerak emas** — hamma narsa faqat joriy foydalanuvchi uchun o‘rnatiladi.
- Internet **kerak emas**.
- Windows 7, 8, 10, 11; Microsoft Office 2010–2021 va Microsoft 365.

## O‘rnatish

1. Arxivni papkaga chiqaring (o‘ng tugma → **Извлечь все**).
2. Word va Excel’ni yoping.
3. **ORNATISH.vbs** ga ikki marta bosing → **OK**.
4. Windows 7 bo‘lsa — tizimdan chiqib qayta kiring.

Natija oynada ko‘rsatiladi va `%TEMP%\OzbekYozuv_ornatish.txt` fayliga yoziladi.

## Ishlatish

**Word va Excel.** Tugmalar lentadagi **НАДСТРОЙКИ (Add-ins)** yorlig‘ida: `Lotin 2026 | Lotin 1995 | Kirill | Kitobcha`. Matnni belgilang (belgilamasangiz — butun hujjat yoki varaq) va tugmani bosing. Qaytarish — **Ctrl+Z**.

**Konverter.** Ish stolidagi **OzbekYozuv_Konverter**: matnni belgilab **Ctrl+C** → konverterda yozuv tugmasi → **Ctrl+V**. Matnni oynaning chap tomoniga yozsangiz ham bo‘ladi.

**UZB klaviaturasi.** **Alt + Shift** bilan **UZB** ni tanlang:

| Tugma | Oddiy | Shift bilan | O‘ng Alt bilan |
|---|---|---|---|
| W | ş | Ş | w / W |
| C | ç | Ç | c / C |
| [ | ğ | Ğ | [ / { |
| ] | ö | Ö | ] / } |
| ` | ʼ (tutuq) | ' | ` / ~ |
| - | - | _ | — / – |

Ctrl+C, Ctrl+W kabi buyruqlar o‘zgarmaydi. Boshqa tillarda (ENG, RUS) klaviatura odatdagidek ishlaydi.

## O‘girish qoidalari

- Kirill → lotin: rasmiy qoidalar (е → ye/e, ё → yo, ц → s/ts, ъ → ʼ, сҳ → sʼh).
- Lotin 1995 → 2026: oʻ/o‘/o'/o\` → ö, gʻ → ğ, sh → ş, ch → ç, tutuq → ʼ.
- **Cheklovsiz:** barcha so‘zlar o‘giriladi, jumladan inglizcha so‘zlar ham (Lotin 2026 da *show* → *şow*). Aralash matnda faqat o‘zbekcha qismni belgilang.
- Lotindan kirillga: «ts» doim «тс» bo‘ladi (*militsiya* → *милитсия*); ruscha so‘zlarni qo‘lda tekshiring.

## Xavfsizlik haqida

- Barcha kod ochiq: `fayllar` papkasidagi `.bas`, `.ps1`, `.vbs`, `.hta` fayllarini Bloknot’da o‘qish mumkin.
- O‘rnatuvchi Office’ning «VBA loyihasiga dastur orqali kirish» ruxsatini faqat o‘rnatish vaqtida yoqadi va keyin avvalgi holatiga qaytaradi.
- UZB klaviaturasi fonda ishlaydigan kichik PowerShell dasturi: u faqat **UZB tanlanganda** W, C, [, ], ` va o‘ng Alt tugmalarini almashtiradi. Hech narsani saqlamaydi va hech qayerga yubormaydi. Shu sababli ayrim antiviruslar uni «klaviatura kuzatuvchisi» deb ogohlantirishi mumkin.
- Administrator huquqi bo‘lsa, alternativ sifatida haqiqiy klaviatura tartibini o‘rnatish mumkin: [doston-usmonov/uzbek-keyboard](https://github.com/doston-usmonov/uzbek-keyboard).

## Olib tashlash

Word va Excel’ni yoping, **OCHIRISH.vbs** ga bosing. UZB tilini ham olib tashlash-tashlamaslikni so‘raydi.

## Litsenziya

MIT — `LICENSE.txt` faylida. Bepul foydalanish, o‘zgartirish va tarqatish mumkin.

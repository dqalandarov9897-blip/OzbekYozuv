# (kirish, maqsad, kutilgan)  maqsad: 1=Lotin 2026, 2=Lotin 1995, 3=Kirill
C = [
 # Lotin 1995 -> 2026
 ("O'zbekiston Respublikasi", 1, "Özbekiston Respublikasi"),
 ("Oʻzbekiston, gʻalaba, shahar, choy", 1, "Özbekiston, ğalaba, şahar, çoy"),
 ("O‘ZBEKISTON SHAHRI CHOYXONA", 1, "ÖZBEKISTON ŞAHRI ÇOYXONA"),
 ("O`g`il, g'ishtin", 1, "Öğil, ğiştin"),
 ("ma'lumot, ta'lim, san'at, Is'hoq", 1, "maʼlumot, taʼlim, sanʼat, Isʼhoq"),
 ("UMUMTA’LIM", 1, "UMUMTAʼLIM"),
 ("'Toshkent' shahri", 1, "'Toşkent' şahri"),
 ("O.Turg'unboyev, Sh.Mirziyoyev", 1, "O.Turğunboyev, Ş.Mirziyoyev"),
 ("bo'lim boshlig'i", 1, "bölim boşliği"),
 # Kirill -> 2026
 ("Ўзбекистон Республикаси", 1, "Özbekiston Respublikasi"),
 ("ер, Европа, поезд, еттинчи", 1, "yer, Yevropa, poyezd, yettinchi".replace("ch","ç")),
 ("ёз, юлдуз, янги, ЁШЛАР, Ёшлар", 1, "yoz, yulduz, yangi, YOŞLAR, Yoşlar"),
 ("цирк, милиция, концерт", 1, "sirk, militsiya, konsert"),
 ("маъно, Исҳоқ, мебель, батальон", 1, "maʼno, Isʼhoq, mebel, batalyon"),
 ("Қўқон, ғалаба, чой, шаҳар, ҳаёт", 1, "Qöqon, ğalaba, çoy, şahar, hayot"),
 ("УМУМТАЪЛИМ МАКТАБИ", 1, "UMUMTAʼLIM MAKTABI"),
 # -> 1995
 ("Özbekiston, şahar, çoy, ğalaba, maʼlumot", 2, "Oʻzbekiston, shahar, choy, gʻalaba, maʼlumot"),
 ("ŞAHAR, Şahar, ÇOY, Ç", 2, "SHAHAR, Shahar, CHOY, Ch"),
 ("Ўзбекистон, шаҳар, ЧОЙ", 2, "Oʻzbekiston, shahar, CHOY"),
 ("O'zbekiston", 2, "Oʻzbekiston"),
 # -> Kirill
 ("O'zbekiston Respublikasi, Xonqa tumani", 3, "Ўзбекистон Республикаси, Хонқа тумани"),
 ("yer, Yevropa, poyezd, yo'l, kuyov, Abdullayev", 3, "ер, Европа, поезд, йўл, куёв, Абдуллаев"),
 ("ega, poeziya, Is'hoq, ma'no", 3, "эга, поэзия, Исҳоқ, маъно"),
 ("UMUMTA'LIM, YOSHLAR", 3, "УМУМТАЪЛИМ, ЁШЛАР"),
 ("Microsoft Word", 3, "Микрософт Ворд"),
 ("Ўзбекистон", 3, "Ўзбекистон"),
 ("O.Turg'unboyev", 3, "О.Турғунбоев"),
 # boshqa belgilar o'zgarmaydi
 ("2026-yil 30-oktyabr, № 15", 1, "2026-yil 30-oktyabr, № 15"),
 ("", 1, ""),
 ("  ikki  bo'shliq\tva\nqator ", 1, "  ikki  böşliq\tva\nqator "),
]

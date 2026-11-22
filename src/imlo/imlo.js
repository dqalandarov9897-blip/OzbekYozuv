// OzbekYozuv imlo tekshiruvchi yadrosi.
// Joriy lotin (1995) va yangi lotin (2026) yozuvidagi matnni tekshiradi.
// Lug'at: Hunspell .aff/.dic fayllari (joriy 1995 yozuvida, masalan uzbek-net/uz-hunspell).
// Yangi yozuvdagi so'z avval 1995 shakliga o'tkaziladi, lug'atdan tekshiriladi,
// takliflar esa yana 2026 shakliga qaytariladi. Konverter (konv.js, OZB) kerak.
// ES3 uslubida yozilgan: brauzer, HTA va Node'da ishlaydi.
var IMLO = (function () {
  var AP = "'\x60‘’ʻʼ´";
  var OK = "ʻ";   // ʻ  (oʻ, gʻ)
  var TQ = "ʼ";   // ʼ  (tutuq belgisi)
  var YANGI = "öğşçÖĞŞÇ"; // ö ğ ş ç
  var HARF = "a-zA-ZÀ-ɏЀ-ӿ";
  var KIRLAT = { "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x",
    "у": "y", "к": "k", "м": "m", "т": "t", "н": "n",
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "Х": "X", "У": "Y" };

  var lug = null;

  function isAp(c) { return c !== "" && AP.indexOf(c) >= 0; }
  function isKir(c) { var k = c.charCodeAt(0); return k >= 0x400 && k <= 0x4FF; }
  function isLat(c) { var k = c.charCodeAt(0); return (k >= 65 && k <= 90) || (k >= 97 && k <= 122) || (k >= 0xC0 && k <= 0x24F); }
  function isUp(c) { return c !== c.toLowerCase() && c === c.toUpperCase(); }
  function kichik(s) { return s.toLowerCase(); }
  function bosh(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  // ---------- Hunspell lug'atini o'qish ----------
  function shartRe(c, kesh) {
    if (c === "." || c === "") return null;
    if (kesh["$" + c] !== undefined) return kesh["$" + c];
    var r = "", ichida = false;
    for (var i = 0; i < c.length; i++) {
      var h = c.charAt(i);
      if (h === "[") { ichida = true; r += h; }
      else if (h === "]") { ichida = false; r += h; }
      else if (h === "^" && ichida && c.charAt(i - 1) === "[") r += h;
      else if (h === "." && !ichida) r += "[\\s\\S]";
      else if (/[\\^$.*+?()[\]{}|\-\/]/.test(h)) r += "\\" + h;
      else r += h;
    }
    var re = new RegExp("(?:" + r + ")$");
    kesh["$" + c] = re;
    return re;
  }

  function yukla(affMatn, dicMatn) {
    var L = { soz: {}, sfx: {}, maxQ: 0, rep: [], tri: "", nosug: "", davom: {}, sozSoni: 0 };
    var kesh = {}, qator = affMatn.split(/\r?\n/), i, p;
    for (i = 0; i < qator.length; i++) {
      var q = qator[i].replace(/\r$/, "");
      if (!q || q.charAt(0) === "#") continue;
      p = q.split(/\s+/);
      if (p[0] === "TRY") L.tri = p[1] || "";
      else if (p[0] === "NOSUGGEST") L.nosug = p[1] || "";
      else if (p[0] === "REP" && p.length >= 3) L.rep.push([p[1].replace(/_/g, " "), p[2].replace(/_/g, " ")]);
      else if (p[0] === "SFX" && p.length >= 5) {
        var bayroq = p[1], olib = p[2] === "0" ? "" : p[2], qosh = p[3], dv = "";
        var k = qosh.indexOf("/");
        if (k >= 0) { dv = qosh.slice(k + 1); qosh = qosh.slice(0, k); }
        if (qosh === "0") qosh = "";
        var e = { f: bayroq, olib: olib, qosh: qosh, dv: dv, re: shartRe(p[4], kesh) };
        var kalit = "$" + qosh;
        (L.sfx[kalit] || (L.sfx[kalit] = [])).push(e);
        if (qosh.length > L.maxQ) L.maxQ = qosh.length;
        for (var j = 0; j < dv.length; j++) L.davom[dv.charAt(j)] = true;
      }
    }
    var d = dicMatn.split(/\r?\n/);
    for (i = 1; i < d.length; i++) {
      var s = d[i].replace(/\r$/, "");
      if (!s) continue;
      var t = s.indexOf("\t"); if (t >= 0) s = s.slice(0, t);
      var b = "", sl = s.indexOf("/");
      if (sl > 0) { b = s.slice(sl + 1); s = s.slice(0, sl); }
      var key = "$" + s;
      if (L.soz[key] === undefined) { L.soz[key] = b; L.sozSoni++; }
      else L.soz[key] += b;
    }
    lug = L;
    return L.sozSoni;
  }

  // ---------- So'zni tekshirish (1995 shaklida) ----------
  function ildiz(s, f) {
    var b = lug.soz["$" + s];
    return b !== undefined && b.indexOf(f) >= 0;
  }
  function bor1(w) {
    var L = lug, n = w.length;
    if (L.soz["$" + w] !== undefined) return true;
    for (var q = 0; q <= L.maxQ && q < n; q++) {
      var ro = L.sfx["$" + w.slice(n - q)];
      if (!ro) continue;
      var asos = w.slice(0, n - q);
      for (var i = 0; i < ro.length; i++) {
        var e = ro[i], st = asos + e.olib;
        if (e.re && !e.re.test(st)) continue;
        if (ildiz(st, e.f)) return true;
        if (L.davom[e.f] && ikki(st, e.f)) return true;
      }
    }
    return false;
  }
  function ikki(st, tashqi) {
    var L = lug, n = st.length;
    for (var q = 0; q <= L.maxQ && q < n; q++) {
      var ro = L.sfx["$" + st.slice(n - q)];
      if (!ro) continue;
      var asos = st.slice(0, n - q);
      for (var i = 0; i < ro.length; i++) {
        var e = ro[i];
        if (e.dv.indexOf(tashqi) < 0) continue;
        var r = asos + e.olib;
        if (e.re && !e.re.test(r)) continue;
        if (ildiz(r, e.f)) return true;
      }
    }
    return false;
  }
  function bor(w) {
    if (bor1(w)) return true;
    var k = kichik(w);
    if (k !== w) {
      var ikkinchi = w.slice(1);
      if (ikkinchi === kichik(ikkinchi) && bor1(k)) return true;       // Gap boshidagi so'z
      if (w === w.toUpperCase() && (bor1(k) || bor1(bosh(k)))) return true; // HAMMASI KATTA
    }
    if (w.indexOf("-") > 0) {
      var qism = w.split("-"), hammasi = true;
      for (var i = 0; i < qism.length; i++) if (!qism[i] || !bor(qism[i])) { hammasi = false; break; }
      if (hammasi) return true;
    }
    return false;
  }

  // ---------- Alifbo ----------
  function belgi1995(w) {
    // Har xil tutuq/apostrof belgilarini 1995 me'yoriga keltiradi: oʻ gʻ va ʼ
    var r = "";
    for (var i = 0; i < w.length; i++) {
      var c = w.charAt(i);
      if (isAp(c)) { var p = w.charAt(i - 1).toLowerCase(); r += (p === "o" || p === "g") ? OK : TQ; }
      else r += c;
    }
    return r;
  }
  function apsiz(s) { var r = ""; for (var i = 0; i < s.length; i++) if (!isAp(s.charAt(i))) r += s.charAt(i); return r; }
  function shakl(w, yozuv) { return OZB.ogir3(w, yozuv === "2026" ? 1 : 2); }
  function aniqla(matn) {
    var y = (matn.match(/[öğşçÖĞŞÇ]/g) || []).length;
    var e = (matn.match(/[oOgG]['\x60‘’ʻʼ´]|sh|ch|Sh|Ch|SH|CH/g) || []).length;
    return y > e ? "2026" : "1995";
  }

  // ---------- Takliflar ----------
  function qoliplash(namuna, s) {
    if (namuna.length > 1 && namuna === namuna.toUpperCase()) return s.toUpperCase();
    if (isUp(namuna.charAt(0))) return bosh(s);
    return s;
  }
  function taklif1995(w, chegara) {
    var L = lug, nat = [], kor = {}, k = kichik(w), i, j;
    chegara = chegara || 5;
    function qosh(s) {
      if (nat.length >= chegara || !s || kor["$" + s]) return;
      kor["$" + s] = 1;
      if (s !== k && bor(s)) {
        var b = L.soz["$" + s];
        if (b !== undefined && L.nosug && b.indexOf(L.nosug) >= 0) return;
        nat.push(s);
      }
    }
    if (bor(bosh(k)) && !bor(k)) qosh(bosh(k));
    for (i = 0; i < L.rep.length && nat.length < chegara; i++) {
      var a = L.rep[i][0], b2 = L.rep[i][1], x = k.indexOf(a);
      while (x >= 0) { qosh(k.slice(0, x) + b2 + k.slice(x + a.length)); x = k.indexOf(a, x + 1); }
    }
    var harf = "abdefghijklmnopqrstuvxyz" + OK + TQ;
    var n = k.length, birinchi = [];
    for (i = 0; i < n; i++) birinchi.push(k.slice(0, i) + k.slice(i + 1));
    for (i = 0; i < n - 1; i++) birinchi.push(k.slice(0, i) + k.charAt(i + 1) + k.charAt(i) + k.slice(i + 2));
    for (i = 0; i < n; i++) for (j = 0; j < harf.length; j++) if (harf.charAt(j) !== k.charAt(i)) birinchi.push(k.slice(0, i) + harf.charAt(j) + k.slice(i + 1));
    for (i = 0; i <= n; i++) for (j = 0; j < harf.length; j++) birinchi.push(k.slice(0, i) + harf.charAt(j) + k.slice(i));
    for (i = 0; i < birinchi.length && nat.length < chegara; i++) qosh(birinchi[i]);
    for (i = 0; i < nat.length; i++) nat[i] = qoliplash(w, nat[i]);
    return nat;
  }

  // ---------- Matnni tekshirish ----------
  var SOZ_RE = new RegExp("[" + HARF + "]+(?:[-'\x60‘’ʻʼ´][" + HARF + "]+)*", "g");
  var HAVOLA_RE = /(?:https?:\/\/|www\.)\S+|\S+@\S+\.\S+/g;

  function tekshir(matn, sozlama) {
    sozlama = sozlama || {};
    var yozuv = sozlama.yozuv && sozlama.yozuv !== "auto" ? sozlama.yozuv : aniqla(matn);
    var xatolar = [], qochir = [], m;
    HAVOLA_RE.lastIndex = 0;
    while ((m = HAVOLA_RE.exec(matn))) qochir.push([m.index, m.index + m[0].length]);
    SOZ_RE.lastIndex = 0;
    while ((m = SOZ_RE.exec(matn))) {
      var w = m[0], b = m.index, o = b + w.length;
      // so'z oxiridagi oʻ / gʻ belgisi (tog', bog')
      if (o < matn.length && isAp(matn.charAt(o)) && /[oOgG]$/.test(w)) { w += matn.charAt(o); o++; SOZ_RE.lastIndex = o; }
      var ich = false;
      for (var z = 0; z < qochir.length; z++) if (b >= qochir[z][0] && b < qochir[z][1]) { ich = true; break; }
      if (ich || w.length < 2) continue;
      if (/[0-9_]/.test(matn.charAt(b - 1) || "") || /[0-9_]/.test(matn.charAt(o) || "")) continue;
      if (/[0-9][ \xA0]$/.test(matn.slice(Math.max(0, b - 2), b))) continue; // 300 dan, 5 ta
      var x = sozTekshir(w, yozuv, sozlama);
      if (x) { x.bosh = b; x.oxir = o; x.soz = w; xatolar.push(x); }
    }
    return { yozuv: yozuv, xatolar: xatolar };
  }

  function sozTekshir(w, yozuv, sozlama) {
    var lat = 0, kir = 0, i, c;
    for (i = 0; i < w.length; i++) { c = w.charAt(i); if (isKir(c)) kir++; else if (isLat(c)) lat++; }
    if (kir && !lat) return null;                         // kirillcha so'z - tekshirilmaydi
    if (kir && lat) {                                     // aralash yozuv
      var t = "";
      for (i = 0; i < w.length; i++) { c = w.charAt(i); t += KIRLAT[c] || c; }
      return { tur: "kirill", izoh: "Lotin soʻzi ichida kirill harfi bor", takliflar: t !== w ? [t] : [] };
    }
    if (w.length <= 5 && w === w.toUpperCase() && !/[ʻʼ'\x60]/.test(w)) return null; // qisqartma: BMT, OAV
    var bir;   // so'zning 1995 shakli (lug'at uchun)
    if (yozuv === "2026") {
      var yangi = shakl(w, "2026");
      if (yangi !== w) {
        if (apsiz(yangi) !== apsiz(w)) return { tur: "alifbo", izoh: "Eski yozuv harfi: yangi alifboda " + yangi, takliflar: [yangi] };
        if (sozlama.belgilar) return { tur: "belgi", izoh: "Tutuq belgisi ʼ boʻlishi kerak", takliflar: [yangi] };
      }
      bir = shakl(yangi, "1995");
    } else {
      var n = belgi1995(w);
      if (/[öğşçÖĞŞÇ]/.test(w)) {
        var eski = shakl(w, "1995");
        return { tur: "alifbo", izoh: "Yangi yozuv harfi: joriy alifboda " + eski, takliflar: [eski] };
      }
      if (n !== w && sozlama.belgilar) return { tur: "belgi", izoh: "Toʻgʻri belgi: " + n, takliflar: [n] };
      bir = n;
    }
    if (!lug || bor(bir)) return null;
    var x = { tur: "imlo", izoh: "Lugʻatda topilmadi", takliflar: null, _bir: bir, _yozuv: yozuv };
    if (sozlama.takliflar !== false) x.takliflar = takliflar(x);
    return x;
  }

  function takliflar(x) {
    if (x.takliflar) return x.takliflar;
    var t = taklif1995(x._bir, 5);
    if (x._yozuv === "2026") for (var i = 0; i < t.length; i++) t[i] = shakl(t[i], "2026");
    x.takliflar = t;
    return t;
  }

  return {
    yukla: yukla,
    tekshir: tekshir,
    takliflar: takliflar,
    bor: function (w, yozuv) { return bor(yozuv === "2026" ? shakl(w, "1995") : belgi1995(w)); },
    aniqla: aniqla,
    // Shaxsiy lug'atga so'z qo'shish (istalgan yozuvda berilishi mumkin)
    qosh: function (w) {
      if (!lug) return false;
      var b = /[öğşçÖĞŞÇ]/.test(w) ? shakl(w, "1995") : belgi1995(w);
      if (lug.soz["$" + b] === undefined) { lug.soz["$" + b] = ""; lug.sozSoni++; }
      return true;
    },
    tayyor: function () { return !!lug; }
  };
})();
if (typeof module !== "undefined") module.exports = IMLO;

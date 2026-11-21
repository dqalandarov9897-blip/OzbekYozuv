// Ozbek alifbo 2026 konverteri (ES3 - Internet Explorer / HTA uchun ham ishlaydi)
var OZB = (function () {
  var AP = "'`‘’ʻʼ´";
  var TUTUQ = "ʼ";
  var ISTISNO = "|chrome|chat|check|chip|cheese|chess|chicken|chief|child|choice|church|chart|channel|chapter|" +
    "shop|show|shift|share|shell|shot|ship|shirt|shoe|short|should|shut|sheet|shelf|shine|" +
    "flash|push|cash|crash|fresh|english|wash|finish|match|search|touch|teacher|school|machine|" +
    "mechanic|technology|orchestra|chemistry|character|schedule|shampoo|chips|ketchup|";
  var MAP = {};
  var pairs = [
    [0x430, "a"], [0x431, "b"], [0x432, "v"], [0x433, "g"], [0x493, "ğ"], [0x434, "d"], [0x436, "j"],
    [0x437, "z"], [0x438, "i"], [0x439, "y"], [0x43A, "k"], [0x49B, "q"], [0x43B, "l"], [0x43C, "m"],
    [0x43D, "n"], [0x43E, "o"], [0x43F, "p"], [0x440, "r"], [0x442, "t"], [0x443, "u"], [0x45E, "ö"],
    [0x444, "f"], [0x445, "x"], [0x4B3, "h"], [0x447, "ç"], [0x448, "ş"], [0x449, "ş"],
    [0x44B, "i"], [0x44D, "e"]
  ];
  for (var p = 0; p < pairs.length; p++) MAP[pairs[p][0]] = pairs[p][1];

  function isAp(c) { return c !== "" && AP.indexOf(c) >= 0; }
  function isLet(c) {
    if (!c) return false;
    var k = c.charCodeAt(0);
    return (k >= 65 && k <= 90) || (k >= 97 && k <= 122) || (k >= 0xC0 && k <= 0x24F) || (k >= 0x400 && k <= 0x4FF);
  }
  function lowK(k) {
    if (k >= 0x410 && k <= 0x42F) return k + 0x20;
    if (k === 0x401) return 0x451;
    if (k === 0x40E) return 0x45E;
    if (k === 0x492) return 0x493;
    if (k === 0x49A) return 0x49B;
    if (k === 0x4B2) return 0x4B3;
    return k;
  }
  function isUp(c) {
    if (!c) return false;
    var k = c.charCodeAt(0);
    if (k >= 65 && k <= 90) return true;
    if (k >= 0x410 && k <= 0x42F) return true;
    return k === 0x401 || k === 0x40E || k === 0x492 || k === 0x49A || k === 0x4B2 ||
      k === 0xD6 || k === 0x11E || k === 0x15E || k === 0xC7;
  }
  function isVow(c) {
    if (!c) return false;
    var k = lowK(c.charCodeAt(0));
    return k === 0x430 || k === 0x435 || k === 0x451 || k === 0x438 || k === 0x43E ||
      k === 0x443 || k === 0x45E || k === 0x44D || k === 0x44E || k === 0x44F;
  }
  function lk(c) { return c ? lowK(c.charCodeAt(0)) : 0; }
  function cap(s, up, nu) {
    if (!up || !s) return s;
    if (nu || s.length === 1) return s.toUpperCase();
    return s.charAt(0).toUpperCase() + s.substring(1);
  }
  function hasCyr(t) {
    for (var i = 0; i < t.length; i++) { var k = t.charCodeAt(i); if (k >= 0x400 && k <= 0x4FF) return true; }
    return false;
  }
  function kirill(t) {
    var r = "", n = t.length;
    for (var i = 0; i < n; i++) {
      var c = t.charAt(i), pv = i > 0 ? t.charAt(i - 1) : "", nx = i < n - 1 ? t.charAt(i + 1) : "";
      var up = isUp(c), nu = isUp(nx), k = lowK(c.charCodeAt(0)), o;
      if (k === 0x435) o = cap((!isLet(pv) || isVow(pv) || lk(pv) === 0x44A || lk(pv) === 0x44C) ? "ye" : "e", up, nu);
      else if (k === 0x451) o = cap("yo", up, nu);
      else if (k === 0x44E) o = cap("yu", up, nu);
      else if (k === 0x44F) o = cap("ya", up, nu);
      else if (k === 0x446) o = cap(isVow(pv) ? "ts" : "s", up, nu);
      else if (k === 0x44A) o = TUTUQ;
      else if (k === 0x44C) o = lk(nx) === 0x43E ? cap("y", up, nu) : "";
      else if (k === 0x441) o = cap("s", up, nu) + (lk(nx) === 0x4B3 ? TUTUQ : "");
      else { o = MAP[k]; o = o ? cap(o, up, nu) : c; }
      r += o;
    }
    return r;
  }
  function himoya(t) {
    var L = t.toLowerCase(), w = "";
    if (L.indexOf("@") >= 0 || L.indexOf("://") >= 0 || L.indexOf("www.") >= 0 || L.indexOf("w") >= 0) return true;
    for (var i = 0; i < L.length; i++) {
      var ch = L.charAt(i), nx = L.charAt(i + 1);
      if (ch === "c" && nx !== "h") return true;
      if (ch === "." && i > 0 && nx >= "a" && nx <= "z") return true;
      if (ch >= "a" && ch <= "z") w += ch;
    }
    return w !== "" && ISTISNO.indexOf("|" + w + "|") >= 0;
  }
  var OG = { "o": "ö", "O": "Ö", "g": "ğ", "G": "Ğ" };
  var SC = { "s": "ş", "S": "Ş", "c": "ç", "C": "Ç" };
  function eski(t) {
    var r = "", i = 0, n = t.length;
    while (i < n) {
      var c = t.charAt(i), nx = i < n - 1 ? t.charAt(i + 1) : "";
      if (isAp(nx) && OG[c]) { r += OG[c]; i += 2; }
      else if ((nx === "h" || nx === "H") && SC[c]) { r += SC[c]; i += 2; }
      else if (isAp(c) && i > 0 && i < n - 1) { r += (isLet(t.charAt(i - 1)) && isLet(nx)) ? TUTUQ : c; i++; }
      else { r += c; i++; }
    }
    return r;
  }
  function token(t) {
    if (hasCyr(t)) t = kirill(t);
    if (!himoya(t)) t = eski(t);
    return t;
  }
  function ogir(s) {
    var res = "", tok = "";
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === " " || ch === "\u000B") {
        if (tok) { res += token(tok); tok = ""; }
        res += ch;
      } else tok += ch;
    }
    if (tok) res += token(tok);
    return res;
  }

  var LATCYR = {"a":0x430,"b":0x431,"d":0x434,"f":0x444,"g":0x433,"h":0x4B3,"i":0x438,"j":0x436,"k":0x43A,"l":0x43B,
    "m":0x43C,"n":0x43D,"o":0x43E,"p":0x43F,"q":0x49B,"r":0x440,"s":0x441,"t":0x442,"u":0x443,"v":0x432,"x":0x445,
    "y":0x439,"z":0x437,"ö":0x45E,"ğ":0x493,"ş":0x448,"ç":0x447};
  function upCyr(k) {
    if (k >= 0x430 && k <= 0x44F) return k - 0x20;
    if (k === 0x451) return 0x401; if (k === 0x45E) return 0x40E; if (k === 0x493) return 0x492;
    if (k === 0x49B) return 0x49A; if (k === 0x4B3) return 0x4B2; return k;
  }
  function cy(k, up) { return String.fromCharCode(up ? upCyr(k) : k); }
  function latVow(c) { return c !== "" && "aeiouAEIOUöÖ".indexOf(c) >= 0; }
  function to1995(t) {
    var r = "", n = t.length;
    for (var i = 0; i < n; i++) {
      var c = t.charAt(i), nx = i < n - 1 ? t.charAt(i + 1) : "", pv = i > 0 ? t.charAt(i - 1) : "";
      var kat = isUp(nx) || (!isLet(nx) && isUp(pv));
      if (c === "ö") r += "oʻ"; else if (c === "Ö") r += "Oʻ";
      else if (c === "ğ") r += "gʻ"; else if (c === "Ğ") r += "Gʻ";
      else if (c === "ş") r += "sh"; else if (c === "Ş") r += kat ? "SH" : "Sh";
      else if (c === "ç") r += "ch"; else if (c === "Ç") r += kat ? "CH" : "Ch";
      else r += c;
    }
    return r;
  }
  function toKirill(t) {
    var r = "", i = 0, n = t.length;
    while (i < n) {
      var c = t.charAt(i), lc = c.toLowerCase(), up = isUp(c), nx = t.charAt(i + 1), ln = nx.toLowerCase();
      var n2 = t.charAt(i + 2), pv = i > 0 ? t.charAt(i - 1) : "", k;
      if (lc === "y" && (ln === "e" || ln === "o" || ln === "u" || ln === "a")) {
        k = ln === "e" ? 0x435 : ln === "o" ? 0x451 : ln === "u" ? 0x44E : 0x44F;
        r += cy(k, up); i += 2;
      } else if (lc === "s" && nx === TUTUQ && n2.toLowerCase() === "h") {
        r += cy(0x441, up) + cy(0x4B3, isUp(n2)); i += 3;
      } else if (lc === "e") {
        r += cy((!isLet(pv) || latVow(pv)) ? 0x44D : 0x435, up); i++;
      } else if (c === TUTUQ) {
        r += (isLet(pv) && isLet(nx)) ? cy(0x44A, isUp(pv) && isUp(nx)) : c; i++;
      } else {
        if (lc === "c") k = (ln === "e" || ln === "i" || ln === "y") ? 0x441 : 0x43A;
        else if (lc === "w") k = 0x432;
        else k = LATCYR[lc];
        r += k ? cy(k, up) : c; i++;
      }
    }
    return r;
  }
  function token3(t, m) {
    var cyr = hasCyr(t);
    if (m === 3 && cyr) return t;
    t = cyr ? kirill(t) : eski(t);
    return m === 2 ? to1995(t) : (m === 3 ? toKirill(t) : t);
  }
  function ogir3(s, m) {
    var res = "", tok = "";
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === " " || ch === "\u000B") {
        if (tok) { res += token3(tok, m); tok = ""; }
        res += ch;
      } else tok += ch;
    }
    if (tok) res += token3(tok, m);
    return res;
  }
  return { ogir: function (s) { return ogir3(s, 1); }, ogir3: ogir3 };
})();
if (typeof module !== "undefined") module.exports = OZB;

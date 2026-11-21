import os
D = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src", "konverter")
js=open(os.path.join(D,"konv.js"),encoding="utf-8").read().replace('if (typeof module !== "undefined") module.exports = OZB;\n','')
html='''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="x-ua-compatible" content="ie=edge">
<title>Özbek yozuvlari — konverter</title>
<hta:application id="ozb" applicationname="OzbekAlifbo2026" border="thick" maximizebutton="yes" singleinstance="yes" scroll="no" icon="mshta.exe">
<style>
body{margin:0;padding:14px 18px;font-family:"Segoe UI",Tahoma,Arial,sans-serif;font-size:14px;background:#f4f4f4;color:#1a1a1a}
h1{margin:0 0 4px 0;font-size:20px}
.sub{margin:0 0 12px 0;color:#555}
.row{width:100%}
.col{width:49%;float:left}
.col2{width:49%;float:right}
label{display:block;font-weight:bold;margin-bottom:4px}
textarea{width:100%;height:300px;box-sizing:border-box;padding:8px;font-size:15px;font-family:"Segoe UI",Tahoma,Arial,sans-serif;border:1px solid #8a8a8a}
#chiq{background:#fff7ef;border-color:#C2410C}
.btns{clear:both;padding-top:12px}
button{font-size:14px;padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #8a8a8a;background:#fff}
.tanlangan{background:#C2410C;color:#fff;border-color:#C2410C;font-weight:bold}
#holat{margin-top:10px;color:#1D4ED8;font-weight:bold;height:20px}
.yord{margin-top:10px;padding:10px 12px;background:#fff;border:1px solid #c9c9c9;font-size:13px;line-height:1.45}
.yord b{color:#1a1a1a}
</style>
<script type="text/javascript">
''' + js + '''
function $(id){return document.getElementById(id);}
function holat(t){$("holat").innerText=t;}
var MAQSAD=1, NOM={1:"Lotin 2026",2:"Lotin 1995",3:"Kirill"};
function ogirish(){ $("chiq").value = OZB.ogir3($("kir").value, MAQSAD); $("chiqnom").innerText = NOM[MAQSAD]; }
function tanla(m){ MAQSAD=m; for(var i=1;i<=3;i++){ $("b"+i).className = (i===m) ? "tanlangan" : ""; } ogirish(); }
function kechiktir(){ window.setTimeout(ogirish, 30); }
function buferdan(m){
  tanla(m);
  var t = "";
  try { t = window.clipboardData.getData("Text") || ""; } catch(e) {}
  if (!t) { holat("Buferda matn yo'q. Avval matnni belgilab Ctrl+C bosing."); return; }
  var r = OZB.ogir3(t, MAQSAD);
  $("kir").value = t; $("chiq").value = r;
  try { window.clipboardData.setData("Text", r); holat("Tayyor! Endi kerakli joyga Ctrl+V bilan qo'ying."); }
  catch(e) { holat("O'girildi. O'ng tomondagi matnni nusxalang."); }
}
function nusxa(){
  try { window.clipboardData.setData("Text", $("chiq").value); holat("Nusxalandi. Ctrl+V bilan qo'ying."); }
  catch(e) { $("chiq").select(); holat("Matn belgilandi - Ctrl+C bosing."); }
}
function tozala(){ $("kir").value=""; $("chiq").value=""; holat(""); $("kir").focus(); }
window.onload=function(){ try{ window.resizeTo(1000,720); window.moveTo(60,30);}catch(e){} $("kir").focus(); };
</script>
</head>
<body>
<h1>Özbek yozuvlari — konverter</h1>
<p class="sub">Istalgan o'zbekcha matnni Lotin 2026, Lotin 1995 yoki Kirill yozuviga o'giradi. Internet kerak emas.</p>
<div class="row">
<div class="col"><label for="kir">Matn (istalgan yozuvda)</label>
<textarea id="kir" onkeyup="kechiktir()" onpaste="kechiktir()" oninput="kechiktir()" onpropertychange="kechiktir()"></textarea></div>
<div class="col2"><label for="chiq">Natija: <span id="chiqnom">Lotin 2026</span></label>
<textarea id="chiq" readonly="readonly"></textarea></div>
</div>
<div class="btns">
<span style="font-weight:bold;margin-right:8px">Buferdagi matnni o'girish:</span>
<button id="b1" class="tanlangan" onclick="buferdan(1)">Lotin 2026</button>
<button id="b2" onclick="buferdan(2)">Lotin 1995</button>
<button id="b3" onclick="buferdan(3)">Kirill</button>
<span style="display:inline-block;width:18px"></span>
<button onclick="nusxa()">Natijani nusxalash</button>
<button onclick="tozala()">Tozalash</button>
</div>
<div id="holat"></div>
<div class="yord">
<b>Eng tez usul (Word, Excel, Telegram — hammasida):</b> matnni belgilang → <b>Ctrl+C</b> → shu oynada kerakli yozuv tugmasi (<b>Lotin 2026</b>, <b>Lotin 1995</b> yoki <b>Kirill</b>) → o'sha joyga qaytib <b>Ctrl+V</b>.<br>
<b>Kitobcha (Word):</b> Разметка страницы → Поля → Настраиваемые поля → «Несколько страниц»: <b>Брошюра</b> → OK. Chop etishda: <b>Печать вручную на обеих сторонах</b>.
</div>
</body>
</html>
'''
open(os.path.join(D,"OzbekYozuv_Konverter.hta"),"w",encoding="utf-8-sig",newline="\r\n").write(html)

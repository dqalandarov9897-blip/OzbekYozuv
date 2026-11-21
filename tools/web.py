"""Onlayn konverter sahifasini yasaydi: src/web/sahifa.html + src/konverter/konv.js -> docs/index.html"""
import os
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
t = open(os.path.join(R, "src", "web", "sahifa.html"), encoding="utf-8").read()
k = open(os.path.join(R, "src", "konverter", "konv.js"), encoding="utf-8").read()
open(os.path.join(R, "docs", "index.html"), "w", encoding="utf-8").write(t.replace("/*KONV*/", k))
print("docs/index.html tayyor")

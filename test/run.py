import json, os, subprocess, sys
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from cases import C
from vbaport import OzbekOgir
js = """const O=require(process.argv[1]);const C=JSON.parse(require('fs').readFileSync(0,'utf8'));
console.log(JSON.stringify(C.map(c=>O.ogir3(c[0],c[1]))));"""
out = json.loads(subprocess.run(["node","-e",js,os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","src","konverter","konv.js")],input=json.dumps(C),capture_output=True,text=True).stdout)
bad=0
for (inp,m,exp),j in zip(C,out):
    v=OzbekOgir(inp,m)
    if v!=exp or j!=exp:
        bad+=1; print(f"XATO m={m}\n  kirish : {inp!r}\n  kutil. : {exp!r}\n  VBA    : {v!r}\n  JS     : {j!r}")
print(f"{len(C)-bad}/{len(C)} o'tdi")

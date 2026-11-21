# VBA COMMON kodining satrma-satr Python nusxasi (sinov uchun)
AP = "'`‘’ʻʼ´"
def IsSep(ch): return ord(ch) in (32,9,10,11,13,160,7,12,1,14,30,31)
def IsAp(ch): return ch!="" and ord(ch) in (39,96,0x2018,0x2019,0x2BB,0x2BC,0xB4)
def IsLet(ch):
    if ch=="": return False
    k=ord(ch); return (65<=k<=90) or (97<=k<=122) or (0xC0<=k<=0x24F) or (0x400<=k<=0x4FF)
def LowK(k):
    if 0x410<=k<=0x42F: return k+0x20
    return {0x401:0x451,0x40E:0x45E,0x492:0x493,0x49A:0x49B,0x4B2:0x4B3}.get(k,k)
def IsUp(ch):
    if ch=="": return False
    k=ord(ch)
    return 65<=k<=90 or 0x410<=k<=0x42F or k in (0x401,0x40E,0x492,0x49A,0x4B2,0xD6,0x11E,0x15E,0xC7)
def IsCyrVowel(ch): return ch!="" and LowK(ord(ch)) in (0x430,0x435,0x451,0x438,0x43E,0x443,0x45E,0x44D,0x44E,0x44F)
def UpX(s): return s.upper()
def Cap(s,up,nu):
    if not up or s=="": return s
    if nu or len(s)==1: return UpX(s)
    return UpX(s[0])+s[1:]
SM={}
for a,b in [("А","A"),("Б","B"),("В","V"),("Г","G"),("Ғ","Ğ"),("Д","D"),("Ж","J"),("З","Z"),("И","I"),("Й","Y"),("К","K"),("Қ","Q"),("Л","L"),("М","M"),("Н","N"),("О","O"),("П","P"),("Р","R"),("С","S"),("Т","T"),("У","U"),("Ў","Ö"),("Ф","F"),("Х","X"),("Ҳ","H"),("Ч","Ç"),("Ш","Ş"),("Щ","Ş"),("Ы","I"),("Э","E")]:
    SM[ord(a.lower())]=b.lower()
def HasCyr(t): return any(0x400<=ord(c)<=0x4FF for c in t)
def LK(s): return LowK(ord((s+" ")[0]))
def KirillToken(t):
    n=len(t); r=""
    for i in range(n):
        c=t[i]; pv=t[i-1] if i>0 else ""; nx=t[i+1] if i<n-1 else ""
        up=IsUp(c); nu=IsUp(nx); k=LowK(ord(c))
        if k==0x435: o=Cap("ye" if (not IsLet(pv) or IsCyrVowel(pv) or LK(pv)==0x44A or LK(pv)==0x44C) else "e",up,nu)
        elif k==0x451: o=Cap("yo",up,nu)
        elif k==0x44E: o=Cap("yu",up,nu)
        elif k==0x44F: o=Cap("ya",up,nu)
        elif k==0x446: o=Cap("ts" if IsCyrVowel(pv) else "s",up,nu)
        elif k==0x44A: o="ʼ"
        elif k==0x44C: o=Cap("y",up,nu) if LK(nx)==0x43E else ""
        elif k==0x441: o=Cap("s",up,nu)+("ʼ" if LK(nx)==0x4B3 else "")
        else:
            o=SM.get(k,""); o=c if o=="" else Cap(o,up,nu)
        r+=o
    return r
def EskiToken(t):
    r="";i=0;n=len(t)
    while i<n:
        c=t[i]; nx=t[i+1] if i<n-1 else ""
        if IsAp(nx) and c in ("o","O","g","G"): r+={"o":"ö","O":"Ö","g":"ğ","G":"Ğ"}[c]; i+=2
        elif nx in ("h","H") and nx!="" and c in ("s","S","c","C"): r+={"s":"ş","S":"Ş","c":"ç","C":"Ç"}[c]; i+=2
        elif IsAp(c) and i>0 and i<n-1:
            r+= "ʼ" if (IsLet(t[i-1]) and IsLet(nx)) else c; i+=1
        else: r+=c; i+=1
    return r
def IsLatVowel(ch): return ch!="" and ch in "aeiouAEIOUöÖ"
def LowX(ch): return {"Ö":"ö","Ğ":"ğ","Ş":"ş","Ç":"ç"}.get(ch, ch.lower())
def To1995(t):
    n=len(t); r=""
    for i in range(n):
        c=t[i]; nx=t[i+1] if i<n-1 else ""; pv=t[i-1] if i>0 else ""
        kat= IsUp(nx) or (not IsLet(nx) and IsUp(pv))
        r+={"ö":"oʻ","Ö":"Oʻ","ğ":"gʻ","Ğ":"Gʻ","ş":"sh","ç":"ch"}.get(c) or ({"Ş":"SH" if kat else "Sh","Ç":"CH" if kat else "Ch"}.get(c) or c)
    return r
def UpCyr(k):
    if 0x430<=k<=0x44F: return k-0x20
    return {0x451:0x401,0x45E:0x40E,0x493:0x492,0x49B:0x49A,0x4B3:0x4B2}.get(k,k)
def Cy(k,up): return chr(UpCyr(k) if up else k)
LC={"a":0x430,"b":0x431,"d":0x434,"f":0x444,"g":0x433,"h":0x4B3,"i":0x438,"j":0x436,"k":0x43A,"l":0x43B,"m":0x43C,"n":0x43D,"o":0x43E,"p":0x43F,"q":0x49B,"r":0x440,"s":0x441,"t":0x442,"u":0x443,"v":0x432,"x":0x445,"y":0x439,"z":0x437,"ö":0x45E,"ğ":0x493,"ş":0x448,"ç":0x447}
def ToKirill(t):
    n=len(t); i=0; r=""
    while i<n:
        c=t[i]; lc=LowX(c); up=IsUp(c)
        nx=t[i+1] if i+1<n else ""; ln=LowX(nx) if nx else ""
        n2=t[i+2] if i+2<n else ""
        pv=t[i-1] if i>0 else ""
        if lc=="y" and ln in ("e","o","u","a") and ln!="":
            r+=Cy({"e":0x435,"o":0x451,"u":0x44E,"a":0x44F}[ln],up); i+=2
        elif lc=="s" and nx=="ʼ" and LowX(n2)=="h" and n2!="":
            r+=Cy(0x441,up)+Cy(0x4B3,IsUp(n2)); i+=3
        elif lc=="e":
            r+=Cy(0x44D if (not IsLet(pv) or IsLatVowel(pv)) else 0x435,up); i+=1
        elif c=="ʼ":
            r+= Cy(0x44A, IsUp(pv) and IsUp(nx)) if (IsLet(pv) and IsLet(nx)) else c; i+=1
        else:
            if lc=="c": k=0x441 if ln in ("e","i","y") and ln!="" else 0x43A
            elif lc=="w": k=0x432
            else: k=LC.get(lc,0)
            r+= c if k==0 else Cy(k,up); i+=1
    return r
def TokenOgir(t,m):
    cyr=HasCyr(t)
    if m==3 and cyr: return t
    t=KirillToken(t) if cyr else EskiToken(t)
    if m==2: t=To1995(t)
    elif m==3: t=ToKirill(t)
    return t
def OzbekOgir(s,m):
    res="";tok=""
    for ch in s:
        if IsSep(ch):
            if tok: res+=TokenOgir(tok,m); tok=""
            res+=ch
        else: tok+=ch
    if tok: res+=TokenOgir(tok,m)
    return res

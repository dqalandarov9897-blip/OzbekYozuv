# OzbekYozuv 1.0 - UZB klaviaturasi (administratorsiz). Windows 7 - 11.
# Til panelida UZB tanlanganda tugmalarni almashtiradi:
#   W = sh-harf (U+015F)   C = ch-harf (U+00E7)   [ = g-harf (U+011F)   ] = o-harf (U+00F6)   ` = tutuq (U+02BC)
#   O'ng Alt bilan: w c [ { ] } ` ~ , o'ng Alt + "-" = uzun tire
$mutex = New-Object System.Threading.Mutex($false, 'Local\OzbekUZBKlaviatura')
if (-not $mutex.WaitOne(0)) { exit }

$code = @'
using System;
using System.Runtime.InteropServices;

public static class UzbKlav
{
    delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    struct KBDLLHOOKSTRUCT { public uint vkCode; public uint scanCode; public uint flags; public uint time; public IntPtr extra; }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }

    [StructLayout(LayoutKind.Explicit)]
    struct InputUnion { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT { public uint type; public InputUnion u; }

    [StructLayout(LayoutKind.Sequential)]
    struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int x; public int y; }

    [DllImport("user32.dll", SetLastError = true)] static extern IntPtr SetWindowsHookEx(int idHook, HookProc fn, IntPtr hMod, uint threadId);
    [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr pid);
    [DllImport("user32.dll")] static extern IntPtr GetKeyboardLayout(uint threadId);
    [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vk);
    [DllImport("user32.dll")] static extern short GetKeyState(int vk);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint n, INPUT[] inputs, int size);
    [DllImport("kernel32.dll")] static extern IntPtr GetModuleHandle(string name);
    [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")] static extern IntPtr SetTimer(IntPtr hWnd, IntPtr id, uint ms, IntPtr fn);

    static HookProc proc;
    static IntPtr hook;
    static bool[] yutildi = new bool[256];
    static bool ongAlt = false;

    static bool Bosilgan(int vk) { return (GetAsyncKeyState(vk) & 0x8000) != 0; }

    static bool Yoz(string s)
    {
        INPUT[] arr = new INPUT[s.Length * 2];
        for (int i = 0; i < s.Length; i++)
        {
            arr[2 * i].type = 1;
            arr[2 * i].u.ki.wScan = s[i];
            arr[2 * i].u.ki.dwFlags = 4;
            arr[2 * i + 1].type = 1;
            arr[2 * i + 1].u.ki.wScan = s[i];
            arr[2 * i + 1].u.ki.dwFlags = 4 | 2;
        }
        return SendInput((uint)arr.Length, arr, Marshal.SizeOf(typeof(INPUT))) == (uint)arr.Length;
    }

    static string Belgi(uint vk, bool shift, bool caps, bool altgr)
    {
        bool katta = shift ^ caps;
        switch (vk)
        {
            case 0x57: return altgr ? (katta ? "W" : "w") : (katta ? "\u015E" : "\u015F");
            case 0x43: return altgr ? (katta ? "C" : "c") : (katta ? "\u00C7" : "\u00E7");
            case 0xDB: return altgr ? (shift ? "{" : "[") : (katta ? "\u011E" : "\u011F");
            case 0xDD: return altgr ? (shift ? "}" : "]") : (katta ? "\u00D6" : "\u00F6");
            case 0xC0: return altgr ? (shift ? "~" : "`") : (shift ? "'" : "\u02BC");
            case 0xBD: return altgr ? (shift ? "\u2013" : "\u2014") : null;
        }
        return null;
    }

    static bool UzbFaol()
    {
        uint tid = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
        long hkl = GetKeyboardLayout(tid).ToInt64();
        return (hkl & 0xFFFF) == 0x0443;
    }

    static IntPtr Ishlov(int nCode, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            if (nCode >= 0)
            {
                KBDLLHOOKSTRUCT k = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                int msg = wParam.ToInt32();
                bool tushish = (msg == 0x101 || msg == 0x105);
                bool bosish = (msg == 0x100 || msg == 0x104);
                if (k.vkCode != 0xE7 && k.vkCode < 256)   // 0xE7 = o'zimiz yuborgan harflar (VK_PACKET)
                {
                    if (k.vkCode == 0xA5)
                    {
                        // O'ng Alt: UZB faol bo'lsa o'zimiz boshqaramiz (tizimga yubormaymiz)
                        if (bosish && UzbFaol()) { ongAlt = true; yutildi[0xA5] = true; return (IntPtr)1; }
                        if (tushish && yutildi[0xA5]) { ongAlt = false; yutildi[0xA5] = false; return (IntPtr)1; }
                    }
                    else if (tushish)
                    {
                        if (yutildi[k.vkCode]) { yutildi[k.vkCode] = false; return (IntPtr)1; }
                    }
                    else if (bosish && UzbFaol())
                    {
                        bool ctrl = Bosilgan(0x11), alt = Bosilgan(0x12);
                        if (ongAlt || !(ctrl || alt))
                        {
                            string s = Belgi(k.vkCode, Bosilgan(0x10), (GetKeyState(0x14) & 1) != 0, ongAlt);
                            if (s != null && Yoz(s)) { yutildi[k.vkCode] = true; return (IntPtr)1; }
                        }
                    }
                }
            }
        }
        catch { }
        return CallNextHookEx(hook, nCode, wParam, lParam);
    }

    static void Ulash()
    {
        if (hook != IntPtr.Zero) UnhookWindowsHookEx(hook);
        hook = SetWindowsHookEx(13, proc, GetModuleHandle(null), 0);
    }

    public static void Ishga()
    {
        proc = new HookProc(Ishlov);
        Ulash();
        SetTimer(IntPtr.Zero, IntPtr.Zero, 60000, IntPtr.Zero);
        MSG m;
        while (GetMessage(out m, IntPtr.Zero, 0, 0) > 0)
        {
            if (m.message == 0x0113) Ulash();   // WM_TIMER: har daqiqada qayta ulash
        }
    }
}
'@

Add-Type -TypeDefinition $code
[UzbKlav]::Ishga()

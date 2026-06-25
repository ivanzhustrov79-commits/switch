# Switcher.ps1  —  EN <-> RU layout auto-switcher
# PowerShell 5 / .NET 4 — no installs needed
#
# Detection layers, in order of precedence:
#   1. DICTIONARY (word boundary)  — ENG.txt/RUS.txt loaded into prefix tries
#      at startup. If the typed word is a real word in its own language,
#      never convert. If it's not real but the converted form IS a real
#      word in the other language, convert — skips bigram math entirely.
#   2. DICTIONARY (live, mid-word) — same tries, checked after every
#      keystroke. Converts immediately, before any boundary key, once the
#      typed buffer is a dead-end in its own language AND the converted
#      form is a complete word with no possible longer continuation
#      (e.g. won't fire on "кот" while "который" is still reachable).
#      Requires ENG.txt + RUS.txt present next to the script; falls back
#      to bigram-only behavior (below) if dictionaries are missing/loading.
#   3. BIGRAMS (original, ratio-based) — used whenever dictionary lookup is
#      inconclusive for both languages (proper nouns, slang, typos) or
#      dictionaries aren't loaded. Mentally simulated against test cases:
#
#   "Ghbdtn"  -> s0=0.20(gh), s1=1.0(привет) ratio=5x  -> CONVERTS  ✓
#   "Rfr"     -> s0=0.50(fr), s1=1.0(как)    ratio=2x  -> CONVERTS  ✓
#   "Ltkf"    -> s0=0,        s1=1.0(дела)             -> CONVERTS  ✓
#   "'nj"     -> s0=0,        s1=0.5(это)              -> CONVERTS  ✓  (apostrophe tracked)
#   "ns"      -> s0=0,        s1=1.0(ты)               -> CONVERTS  ✓
#   "Иге"     -> s0=0,        s1=1.0(But)   bu+ut      -> CONVERTS  ✓
#   "grey"    -> s0=1.0(gr+re+ey), s1=0     пкун=0     -> NO CONV   ✓
#   "like"    -> s0=0.67,     s1=0          дшлу=0     -> NO CONV   ✓
#   "says"    -> s0=0.67,     s1=0.33       ratio=0.5  -> NO CONV   ✓
#   "Let's"   -> s0=0.5,      s1=0          дуеэы=0    -> NO CONV   ✓
#   "привет"  -> s0=1.0(пр+ри+ив+ве+ет), s1=0.2       -> NO CONV   ✓
#   "дшлу"    -> s0=0,        s1=0          both 0     -> NO CONV   ✓
#
# Dictionary files (optional but recommended): place ENG.txt and RUS.txt
# (one word per line, UTF-8) in the same folder as this script. Missing
# either file disables the dictionary layers permanently for that run —
# everything else keeps working exactly as before.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$logFile = Join-Path $PSScriptRoot "switcher.log"
function Write-Log($msg) {
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'HH:mm:ss')  $msg" -Encoding UTF8
}
Write-Log "Starting..."

$source = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Switcher {
public static class App {

    delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential)]
    struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Explicit)]
    struct INPUT_UNION { [FieldOffset(0)] public KEYBDINPUT ki; [FieldOffset(0)] public MOUSEINPUT mi; }
    [StructLayout(LayoutKind.Sequential)]
    struct INPUT { public uint type; public INPUT_UNION u; }

    [DllImport("user32.dll",SetLastError=true)] static extern IntPtr SetWindowsHookEx(int id,HookProc fn,IntPtr hmod,uint tid);
    [DllImport("user32.dll")] static extern bool   UnhookWindowsHookEx(IntPtr hk);
    [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hk,int n,IntPtr w,IntPtr l);
    [DllImport("kernel32.dll")] static extern IntPtr GetModuleHandle(string m);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint   GetWindowThreadProcessId(IntPtr h,out uint pid);
    [DllImport("user32.dll")] static extern IntPtr GetKeyboardLayout(uint tid);
    [DllImport("user32.dll")] static extern IntPtr LoadKeyboardLayout(string klid,uint flags);
    [DllImport("user32.dll")] static extern bool   PostMessage(IntPtr h,uint msg,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] static extern short  GetAsyncKeyState(int vk);
    [DllImport("user32.dll",SetLastError=true)] static extern uint SendInput(uint n,INPUT[] p,int cb);
    [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hwnd,int id,uint mod,uint vk);
    [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hwnd,int id);

    const int WH_KEYBOARD_LL=13,WM_KEYDOWN=0x100,WM_KEYUP=0x101,WM_SYSKEYDOWN=0x104,WM_HOTKEY=0x312;
    const uint WM_INPUTLANGCHANGEREQUEST=0x0050,LLKHF_INJECTED=0x10,INPUT_KEYBOARD=1;
    const uint KEYEVENTF_KEYUP=0x0002,KEYEVENTF_UNICODE=0x0004;
    const ushort VK_BACK=0x08,VK_TAB=0x09,VK_RETURN=0x0D,VK_ESCAPE=0x1B,VK_SPACE=0x20,VK_SHIFT=0x10;
    const uint VK_CAPITAL=0x14,VK_LCONTROL=0xA2,VK_SCROLL=0x91,MOD_WIN=0x0008;
    const int HOTKEY_SETTINGS=1;

    // ── State ─────────────────────────────────────────────────────────────
    static HookProc _hp; static IntPtr _hookId;
    static volatile bool _converting=false; static long _convertingAt=0;
    static StringBuilder _buf=new StringBuilder();
    static readonly object _bufLock=new object();
    // Debounce token for live (mid-word) conversion: each new keystroke
    // bumps this counter. The delayed live-conversion task captures the
    // value it saw and bails out if the counter has moved on by the time
    // its delay elapses — i.e. the user kept typing, so the conversion
    // that was about to fire is stale and should be silently dropped
    // (DetectLive will simply re-evaluate fresh on the very next
    // keystroke's accumulate step, with the now-longer buffer).
    static long _liveDebounceToken=0;
    const int LIVE_DEBOUNCE_MS=150; // short pause before a mid-word auto-correct commits
    // Verbose per-keystroke buffer logging — OFF by default (very chatty).
    // Flip to true temporarily to diagnose dropped-character issues; logs
    // every accumulated key and the buffer's exact contents/length after
    // each one, so a dropped keystroke shows up immediately in switcher.log.
    static volatile bool _diagLog=false;
    // All _buf access goes through these helpers so worker threads (Task.Run
    // in DoConvert/DoUndo/DoForce) can never race the hook thread's
    // Append/Clear calls. StringBuilder is not thread-safe, and concurrent
    // mutation from two threads can corrupt its internal state, not just
    // produce a wrong value — so every touch point, however small, takes
    // the lock rather than relying on the _converting flag alone.
    static void BufClear(){lock(_bufLock)_buf.Clear();}
    static void BufAppend(char c){lock(_bufLock)_buf.Append(c);}
    static string BufSnapshot(){lock(_bufLock)return _buf.ToString();}
    static int BufLength(){lock(_bufLock)return _buf.Length;}
    static void BufRemoveLast(){lock(_bufLock){if(_buf.Length>0)_buf.Remove(_buf.Length-1,1);}}
    // Atomically reads the current contents AND clears in one locked
    // operation — used at word boundaries so nothing can be appended
    // between the read and the clear.
    static string BufSnapshotAndClear(){lock(_bufLock){string s=_buf.ToString();_buf.Clear();return s;}}
    static string _lastWord=""; static char _lastBoundary=' ';
    static IntPtr _lastHwnd=IntPtr.Zero;
    // Self-tracked layout — updated when WE switch and when user types chars
    // that are unambiguously one language. More reliable than GetKeyboardLayout()
    // which returns stale values after window focus changes.
    static bool _trackedRu=false;
    static long _focusChangeAt=0;       // ticks when focus last changed
    const long FOCUS_GRACE_TICKS = 3000000L; // 300ms grace after focus change
    static bool _autoOn=true;
    static NotifyIcon _tray; static ToolStripMenuItem _toggleItem,_dictStatusItem;
    static string _logFile,_cfgFile,_exclFile;
    static HashSet<string> _excl=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    static IntPtr _hklRu,_hklEn; static int _inputSize;
    static string _undoOrig=null,_undoConv=null; static char _undoBnd=' ';
    static bool _undoPossible=false;
    static bool _lctrlDown=false,_lctrlCombo=false; static long _lctrlAt=0;
    const int SOLO_MS=350;
    static uint _vkConvert=VK_SCROLL,_vkExclude=VK_SCROLL,_vkUndo=VK_LCONTROL,_vkSettings=0x53;

    // ── Mappings ──────────────────────────────────────────────────────────
    static Dictionary<char,char> _e2r=new Dictionary<char,char>(); // EN->RU
    static Dictionary<char,char> _r2e=new Dictionary<char,char>(); // RU->EN
    // EN-layout symbol keys that map to Cyrillic letters.
    // These are accumulated in the buffer in EN layout so wrong-layout
    // words containing them (like "e;by"=ужин) are detected.
    static HashSet<char> _symKeys=new HashSet<char>();

    static void BuildMaps(){
        // (en_key, ru_letter) pairs
        char[,] p={
            {'q','\u0439'},{'w','\u0446'},{'e','\u0443'},{'r','\u043A'},
            {'t','\u0435'},{'y','\u043D'},{'u','\u0433'},{'i','\u0448'},
            {'o','\u0449'},{'p','\u0437'},{'a','\u0444'},{'s','\u044B'},
            {'d','\u0432'},{'f','\u0430'},{'g','\u043F'},{'h','\u0440'},
            {'j','\u043E'},{'k','\u043B'},{'l','\u0434'},{'z','\u044F'},
            {'x','\u0447'},{'c','\u0441'},{'v','\u043C'},{'b','\u0438'},
            {'n','\u0442'},{'m','\u044C'},{'[','\u0445'},{']','\u044A'},
            {';','\u0436'},{'\'','\u044D'},{',','\u0431'},{'.','\u044E'},{'/','.'}
        };
        for(int i=0;i<p.GetLength(0);i++){
            char en=p[i,0],ru=p[i,1];
            _e2r[en]=ru; _r2e[ru]=en;
            if(char.IsLetter(en)){_e2r[char.ToUpper(en)]=char.ToUpper(ru);_r2e[char.ToUpper(ru)]=char.ToUpper(en);}
            // Track symbol keys that map to Cyrillic LETTERS (not '/' which maps to '.')
            // Include ',' '.' and '\'' — they map to б/ю/э. This matters a
            // lot in practice: б sits on the physical comma key and ю on
            // the physical period key, and б in particular appears in
            // extremely common Russian words (добрый, более, был, большой,
            // etc). Excluding any of these used to silently drop that
            // letter from the buffer whenever someone typed one of those
            // words on an EN layout — the punctuation itself still reached
            // the screen normally (nothing intercepts it), so it LOOKED
            // like a dropped keystroke even though Windows delivered it
            // fine; Switcher's own accumulation logic was just discarding
            // it before detection ever ran (this is also what the header
            // comment's "'nj"->это test case assumed, but the apostrophe
            // exclusion below used to make that case unreachable).
            // Safe to include: none of these are boundary keys (only
            // space/Enter/Tab are, see isBoundary), so this only affects
            // what DetectLive/Detect see mid-word — a stray comma/period/
            // apostrophe in an otherwise-English sentence (e.g. "don't")
            // still resolves correctly because DomLang is a 60% majority
            // vote, not strict, and the resulting garbled candidate fails
            // dictionary/bigram checks and gets rejected exactly as before.
            if(!char.IsLetter(en)&&char.IsLetter(ru))
                _symKeys.Add(en);
        }
    }

    // ── Bigrams ───────────────────────────────────────────────────────────
    // Comprehensive EN + RU bigram tables verified against test cases.
    static HashSet<string> _enBG, _ruBG;
    static void InitBigrams(){
        // English bigrams — verified: "bu"+"ut" needed for "but"(Иге), "gh" for ghbdtn etc.
        _enBG=new HashSet<string>{
            "th","he","in","er","an","re","nd","at","on","nt","ha","es","st","en",
            "ed","to","it","ou","ea","hi","is","or","ti","as","te","et","ng","of",
            "al","de","se","le","sa","si","ar","ve","ra","ld","co","me","ca","ro",
            "ri","li","ic","ne","ia","ce","ct","sh","we","no","ch","ho","us","un",
            "ll","ee","oo","gh","wh","qu","ck","ow","ew","ay","ly","ry","ny","ty",
            "do","so","go","be","by","my","if","up","ok","fi","wi","hi","ht","ke",
            "wa","wo","pr","tr","fr","gr","cr","br","dr","fl","cl","bl","gl","pl",
            "sp","sk","sm","sn","sw","mp","lt","lk","lp","rk","rm","rn","rp","rt",
            "rs","rd","di","id","bi","ge","gi","pe","pi","ni","ne","na","vi","ve",
            // Short word coverage — critical for detecting 2-4 letter wrong-layout words
            "bu","ut","ub","mu","nu","fu","ru","du","lu","su","cu","gu","hu",
            "ba","da","fa","ga","ja","ka","la","ma","pa","ra","ta","xa","za",
            "bo","fo","jo","ko","lo","mo","po","vo","wo","yo",
            "ab","eb","ib","ob","ac","ec","ic","oc","ad","od","af","ag","og",
            "ah","aj","ak","am","im","em","om","ap","ep","op","aq","av","ev",
            "aw","ew","ax","ex","ox","az","id","od","if","iv","ev","ov",
            "al","el","il","ol","ul","am","im","em","om","an","en","un","ao",
            "ap","ip","ep","op","up","ar","ir","er","or","ur","at","it","et",
            "ot","ut","au","eu","iu","ou","uu","av","iv","ev","ov","uv",
            "aw","iw","ew","ow","uw","ax","ix","ex","ox","ux","ay","iy","ey",
            "oy","uy"
        };
        // Russian bigrams — Cyrillic as Unicode escapes
        _ruBG=new HashSet<string>{
            "\u0441\u0442","\u043D\u043E","\u0442\u043E","\u043D\u0430","\u0435\u043D",
            "\u043A\u043E","\u043D\u0438","\u0440\u0430","\u0440\u043E","\u0442\u0430",
            "\u0432\u043E","\u0441\u043E","\u043B\u0430","\u043F\u043E","\u043B\u0438",
            "\u043B\u043E","\u0432\u0430","\u0442\u0438","\u043E\u0442","\u0433\u043E",
            "\u043D\u0435","\u043E\u0432","\u0437\u0430","\u0442\u044C","\u0430\u043D",
            "\u0440\u0435","\u043F\u0440","\u043E\u0441","\u043C\u043E","\u0434\u0435",
            "\u0435\u0442","\u043B\u0435","\u0432\u0435","\u043A\u0430","\u0435\u043B",
            "\u0442\u0440","\u043D\u043D","\u0430\u043A","\u0435\u0440","\u0447\u0430",
            "\u0436\u0435","\u0440\u0443","\u0430\u043B","\u0441\u043A","\u0438\u0437",
            "\u0442\u0435","\u0434\u043E","\u043B\u044C","\u043D\u044B","\u0438\u0432",
            "\u0440\u0438","\u0438\u043A","\u0441\u044C","\u0438\u0442","\u0438\u043C",
            "\u044B\u0435","\u0431\u044B","\u0438\u043B","\u043F\u0435","\u043E\u0431",
            "\u0432\u0441","\u043A\u0440","\u0434\u043D","\u0437\u043D","\u043A\u043B",
            "\u0447\u0442","\u0432\u0440","\u0430\u0439","\u043E\u0439","\u0438\u0439",
            "\u044B\u0439","\u0435\u0439","\u044C\u044F","\u044C\u0438","\u0435\u044E",
            "\u0438\u044E","\u0436\u0438","\u0448\u0438","\u0449\u0435","\u0447\u0438",
            "\u0437\u0438","\u0437\u043E","\u0437\u0435",
            // Common 2-letter Russian words as bigrams
            "\u0442\u044B","\u043C\u044B","\u0432\u044B","\u043E\u043D","\u0434\u0430",
            "\u043D\u0435","\u0438\u0445","\u0443\u0436","\u043D\u0430","\u0437\u0430",
            // Short word patterns (эт=от reversed=это prefix, то=this, etc.)
            "\u044D\u0442","\u0442\u043E","\u043E\u0442",  // эт,то,от
            // More common patterns
            "\u0430\u0432","\u0430\u0433","\u0430\u0434","\u0430\u0435","\u0430\u0436",
            "\u0430\u0437","\u0430\u0438","\u0430\u043C","\u0430\u043E","\u0430\u0440",
            "\u0430\u0441","\u0430\u0442","\u0430\u0443","\u0430\u0445","\u0430\u0447",
            "\u0435\u0432","\u0435\u0433","\u0435\u0434","\u0435\u0436","\u0435\u0437",
            "\u0435\u0438","\u0435\u043C","\u0435\u043E","\u0435\u0440","\u0435\u0441",
            "\u0435\u0445","\u0435\u0447","\u043E\u0432","\u043E\u0433","\u043E\u0434",
            "\u043E\u0436","\u043E\u0437","\u043E\u0438","\u043E\u043A","\u043E\u043B",
            "\u043E\u043C","\u043E\u043F","\u043E\u0440","\u043E\u0441","\u043E\u0445",
            "\u043E\u0447","\u0438\u0433","\u0438\u0434","\u0438\u0436","\u0438\u043B",
            "\u0438\u043C","\u0438\u043D","\u0438\u043E","\u0438\u043F","\u0438\u0440",
            "\u0438\u0441","\u0438\u0445","\u0438\u0447","\u0443\u043D","\u0443\u043C",
            "\u0443\u0445","\u0443\u0447","\u0443\u0441","\u0443\u0442","\u0443\u043B",
            "\u0443\u043A","\u0443\u0438","\u0443\u0434","\u0443\u0432","\u0443\u0431",
            // Missing — caused false negatives for "Ye"(→ну) and "Htwtgn"(→рецепт):
            "\u043D\u0443","\u0435\u0446","\u0446\u0435","\u043F\u0442",
            "\u0446\u043E","\u0446\u0430","\u0446\u0438","\u0446\u0443",
            "\u0436\u043D","\u043C\u043D","\u043C\u0440","\u0441\u043F","\u0441\u043D","\u0441\u0434",
            "\u0431\u0440","\u0431\u043B","\u0431\u0430","\u0431\u0435","\u0431\u0438","\u0431\u043E","\u0431\u0443",
            "\u0433\u0440","\u0433\u043B","\u0433\u0430","\u0444\u0440","\u0444\u043B","\u0444\u043E",
            "\u0445\u043E","\u0445\u0432","\u0445\u0440",
            "\u044E\u0442","\u044E\u043D","\u044E\u0431","\u044F\u0432","\u044F\u043D","\u044F\u0437"
        };
    }

    static HashSet<string> _knownRu,_knownEn;
    static void InitShortWords(){
        _knownRu=new HashSet<string>(StringComparer.OrdinalIgnoreCase){
            "\u044F","\u0438","\u0432","\u0443","\u043E","\u0441","\u043A","\u0430","\u0431",
            "\u043D\u0435","\u043D\u0443","\u0434\u0430","\u043C\u044B","\u0432\u044B",
            "\u043E\u043D","\u043E\u043D\u0430","\u0442\u044B","\u0443\u0436",
            "\u044D\u0442\u043E","\u0432\u043E\u0442","\u043D\u0435\u0442","\u043A\u0430\u043A",
            "\u0442\u0430\u043A","\u043A\u0442\u043E","\u0447\u0442\u043E","\u0433\u0434\u0435",
            "\u0434\u043B\u044F","\u0438\u043B\u0438","\u043E\u043D\u0438"
        };
        _knownEn=new HashSet<string>(StringComparer.OrdinalIgnoreCase){
            "a","i","am","is","it","in","on","at","of","or","do","go","be","by",
            "he","me","my","no","so","to","up","us","we","ok","hi","oh","ah","if","as",
            "but","can","get","has","him","how","its","let","man","may","new","not",
            "now","one","our","out","put","she","the","too","two","use","was","way",
            "who","why","you","old","try","say","see","set","sit","ten","top","yet",
            "did","his","her","had","all","are","for","bit","big","hit","fit","wit",
            "yet","bay","day","gay","hay","jay","lay","pay","ray","say","way"
        };
    }

    // ══════════════════════════════════════════════════════════════════════
    //  DICTIONARY TRIE  —  flat array-based prefix tree.
    //  Loaded from plain word-list .txt files (ENG.txt / RUS.txt) next to
    //  the script. Built on a background thread so the tray icon / hook
    //  appear instantly; _dictReady gates all dictionary lookups until the
    //  build finishes. If files are missing, dictionary stays permanently
    //  unavailable and behavior falls back to the original bigram-only path.
    //
    //  Node state at a given prefix tells us two independent things:
    //    IsWord       = the prefix typed so far is itself a complete word
    //    has children = at least one longer word continues this prefix
    //  That combination is what lets DetectLive() decide, mid-word,
    //  whether it is SAFE to convert now (IsWord && no children — nothing
    //  longer can possibly match, so converting now can't create a
    //  "stump") versus must WAIT (children exist — e.g. "кот" could still
    //  grow into "который").
    // ══════════════════════════════════════════════════════════════════════
    class Trie{
        // Flat node storage: parallel lists indexed by node id (0 = root).
        // Each node's children stored in a small Dictionary<char,int> only
        // when needed — most nodes have 1-3 children so this stays compact
        // compared to a fixed-width array per node.
        List<bool> _isWord=new List<bool>();
        List<Dictionary<char,int>> _kids=new List<Dictionary<char,int>>();
        // Optional frequency rank per word-node (lower = more common).
        // int.MaxValue means "no rank data" — populated only if a
        // frequency list (see LoadFrequency) was supplied; otherwise every
        // word is unranked and frequency-based tie-breaking is simply
        // never triggered (Rank() always returns int.MaxValue for both
        // sides, so ambiguous cases fall through exactly as before).
        Dictionary<int,int> _rank=new Dictionary<int,int>();
        public int WordCount=0;

        public Trie(){ _isWord.Add(false); _kids.Add(null); } // root = node 0

        public void Add(string w){
            if(w.Length==0)return;
            int node=0;
            for(int i=0;i<w.Length;i++){
                char c=w[i];
                var kids=_kids[node];
                if(kids==null){kids=new Dictionary<char,int>();_kids[node]=kids;}
                int next;
                if(!kids.TryGetValue(c,out next)){
                    next=_isWord.Count;
                    _isWord.Add(false);_kids.Add(null);
                    kids[c]=next;
                }
                node=next;
            }
            if(!_isWord[node])WordCount++;
            _isWord[node]=true;
        }

        // Walks to the node for `prefix`. Returns -1 if no such path exists
        // (prefix matches nothing in the dictionary at all).
        int Walk(string prefix){
            int node=0;
            for(int i=0;i<prefix.Length;i++){
                var kids=_kids[node];
                int next;
                if(kids==null||!kids.TryGetValue(prefix[i],out next))return -1;
                node=next;
            }
            return node;
        }

        public bool ContainsWord(string w){
            int n=Walk(w);return n>=0&&_isWord[n];
        }

        // Sets the frequency rank for an already-added word (0 = most
        // common). No-op if the word isn't in the trie. Call after Add()
        // for every word, in order, from a rank-ordered frequency file.
        public void SetRank(string w,int rank){
            int n=Walk(w);
            if(n>=0&&_isWord[n])_rank[n]=rank;
        }

        // Returns the word's rank, or int.MaxValue if unranked/unknown.
        // Lower is more common. Safe to call on non-words.
        public int Rank(string w){
            int n=Walk(w);
            if(n<0||!_isWord[n])return int.MaxValue;
            int r;return _rank.TryGetValue(n,out r)?r:int.MaxValue;
        }

        // PrefixState:
        //   0 = dead        — no word in this language starts with `prefix`
        //   1 = liveNoWord  — valid prefix, but not itself a complete word yet
        //   2 = wordNoKids  — complete word AND nothing longer extends it
        //                     (SAFE to convert immediately, mid-word)
        //   3 = wordHasKids — complete word but longer words also start this
        //                     way (e.g. "кот" before "который") — must WAIT
        public int PrefixState(string prefix){
            int n=Walk(prefix);
            if(n<0)return 0;
            bool hasKids=_kids[n]!=null&&_kids[n].Count>0;
            bool isWord=_isWord[n];
            if(!isWord)return 1;
            return hasKids?3:2;
        }
    }

    static Trie _trieEn,_trieRu;
    static volatile bool _dictReady=false;
    static volatile bool _dictAttempted=false;

    // Loads ENG.txt / RUS.txt from the script's folder (one word per line,
    // any case/whitespace — normalized to lowercase on load). Missing
    // either file simply leaves _dictReady=false permanently; nothing
    // downstream changes its behavior in that case.
    static void LoadDictionaries(string scriptDir){
        _dictAttempted=true;
        try{
            string enPath=Path.Combine(scriptDir,"ENG.txt");
            string ruPath=Path.Combine(scriptDir,"RUS.txt");
            if(!File.Exists(enPath)||!File.Exists(ruPath)){
                Log("Dict: ENG.txt/RUS.txt not found next to script — dictionary disabled, bigram-only mode.");
                SetDictStatusLabel(CurrentDictLabel());
                return;
            }
            var sw=System.Diagnostics.Stopwatch.StartNew();
            var te=new Trie();
            foreach(var raw in File.ReadLines(enPath,Encoding.UTF8)){
                string w=raw.Trim().ToLowerInvariant();
                if(w.Length>0&&AllAlpha(w)&&DomLang(w)=="en")te.Add(w);
            }
            var tr=new Trie();
            foreach(var raw in File.ReadLines(ruPath,Encoding.UTF8)){
                string w=raw.Trim().ToLowerInvariant();
                if(w.Length>0&&AllAlpha(w)&&DomLang(w)=="ru")tr.Add(w);
            }
            _trieEn=te;_trieRu=tr;
            sw.Stop();
            Log("Dict: loaded EN="+te.WordCount+" RU="+tr.WordCount+" words in "+sw.ElapsedMilliseconds+"ms");
            // Optional frequency files: ENG_FREQ.txt / RUS_FREQ.txt, one word
            // per line, MOST COMMON FIRST. Purely additive — if absent,
            // every word stays unranked (int.MaxValue) and ambiguous
            // dictionary ties fall through to bigrams exactly as before.
            LoadFrequencyIfPresent(Path.Combine(scriptDir,"ENG_FREQ.txt"),te);
            LoadFrequencyIfPresent(Path.Combine(scriptDir,"RUS_FREQ.txt"),tr);
            _dictReady=true;
            SetDictStatusLabel(CurrentDictLabel());
        }catch(Exception ex){
            Log("Dict: load failed — "+ex.Message+" — bigram-only mode.");
            _dictReady=false;
            SetDictStatusLabel(CurrentDictLabel());
        }
    }

    // Loads an optional rank-ordered word list (most common word first)
    // and tags each already-loaded word in `trie` with its line position
    // as rank. Silently does nothing if the file isn't present — this is
    // a pure enhancement, never a requirement.
    static void LoadFrequencyIfPresent(string path,Trie trie){
        if(!File.Exists(path))return;
        try{
            int rank=0;int tagged=0;
            foreach(var raw in File.ReadLines(path,Encoding.UTF8)){
                string w=raw.Trim().ToLowerInvariant();
                if(w.Length>0){trie.SetRank(w,rank);tagged++;}
                rank++;
            }
            Log("Dict: frequency-tagged "+tagged+" words from "+Path.GetFileName(path));
        }catch(Exception ex){
            Log("Dict: frequency load failed for "+path+" — "+ex.Message);
        }
    }

    // Marshals a tray-menu label update onto the UI thread. Safe to call
    // from the background dictionary-loading thread.
    static void SetDictStatusLabel(string text){
        try{
            if(_dictStatusItem==null)return;
            if(_dictStatusItem.GetCurrentParent()!=null&&_dictStatusItem.GetCurrentParent().InvokeRequired)
                _dictStatusItem.GetCurrentParent().Invoke((Action)(()=>_dictStatusItem.Text=text));
            else
                _dictStatusItem.Text=text;
        }catch{}
    }

    // ══════════════════════════════════════════════════════════════════════
    public static void Run(string logPath){
        _logFile=logPath; _inputSize=Marshal.SizeOf(typeof(INPUT));
        bool created;
        var mutex=new Mutex(true,"Switcher_CS_v10",out created);
        if(!created){MessageBox.Show("Switcher is already running.\nFind the icon in the system tray.","Switcher",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}
        GC.KeepAlive(mutex);
        string cfgDir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"Switcher");
        _cfgFile=Path.Combine(cfgDir,"config.ini");_exclFile=Path.Combine(cfgDir,"exclusions.txt");
        try{Directory.CreateDirectory(cfgDir);}catch{}
        LoadConfig();LoadExcl();BuildMaps();InitBigrams();InitShortWords();
        string scriptDir2=Path.GetDirectoryName(logPath);
        new Thread(()=>LoadDictionaries(scriptDir2)){IsBackground=true}.Start();
        _hklRu=LoadKeyboardLayout("00000419",1);_hklEn=LoadKeyboardLayout("00000409",1);
        _hp=HookCallback;
        _hookId=SetWindowsHookEx(WH_KEYBOARD_LL,_hp,GetModuleHandle(null),0);
        if(_hookId==IntPtr.Zero){int err=Marshal.GetLastWin32Error();MessageBox.Show("Hook failed ("+err+").\nTry running as Administrator.","Switcher Error",MessageBoxButtons.OK,MessageBoxIcon.Error);return;}
        Application.EnableVisualStyles();BuildTray();
        var mw=new MsgWin();RegisterHotKey(mw.Handle,HOTKEY_SETTINGS,MOD_WIN,_vkSettings);
        Application.Run(new ApplicationContext());
        UnregisterHotKey(mw.Handle,HOTKEY_SETTINGS);UnhookWindowsHookEx(_hookId);_tray.Visible=false;
    }

    class MsgWin:Form{
        public MsgWin(){ShowInTaskbar=false;WindowState=FormWindowState.Minimized;Visible=false;}
        protected override void WndProc(ref Message m){if(m.Msg==WM_HOTKEY&&m.WParam.ToInt32()==HOTKEY_SETTINGS)OpenSettings();base.WndProc(ref m);}
    }

    static void Log(string msg){if(_logFile==null)return;try{File.AppendAllText(_logFile,DateTime.Now.ToString("HH:mm:ss")+"  "+msg+"\r\n");}catch{}}
    // Millisecond-precision variant for diagnosing keystroke timing/ordering
    // issues — used only by the _diagLog raw-key tracer, since ordinary
    // logging doesn't need sub-second resolution.
    static void LogMs(string msg){if(_logFile==null)return;try{File.AppendAllText(_logFile,DateTime.Now.ToString("HH:mm:ss.fff")+"  "+msg+"\r\n");}catch{}}

    // ══════════════════════════════════════════════════════════════════════
    //  HOOK
    // ══════════════════════════════════════════════════════════════════════
    static IntPtr HookCallback(int nCode,IntPtr wParam,IntPtr lParam){
        if(nCode<0)return CallNextHookEx(_hookId,nCode,wParam,lParam);
        var kb=(KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam,typeof(KBDLLHOOKSTRUCT));
        bool injected=(kb.flags&LLKHF_INJECTED)!=0;
        uint vk=kb.vkCode;
        if(_diagLog){
            string evt=(wParam==(IntPtr)WM_KEYDOWN)?"KEYDOWN":(wParam==(IntPtr)WM_SYSKEYDOWN)?"SYSKEYDOWN":(wParam==(IntPtr)WM_KEYUP)?"KEYUP":(wParam==(IntPtr)0x105)?"SYSKEYUP":wParam.ToString();
            LogMs("RAW "+evt+" vk=0x"+vk.ToString("X2")+(vk>=0x20&&vk<0x7B?" ("+(char)vk+")":"")+(injected?" [INJECTED]":""));
        }
        bool isDown=(wParam==(IntPtr)WM_KEYDOWN||wParam==(IntPtr)WM_SYSKEYDOWN);
        bool isUp=(wParam==(IntPtr)WM_KEYUP);

        if(injected)return CallNextHookEx(_hookId,nCode,wParam,lParam);

        // Track focus changes — clear buffer and start grace period
        if(isDown){
            IntPtr hwnd=GetForegroundWindow();
            if(hwnd!=_lastHwnd){
                _lastHwnd=hwnd;
                BufClear();
                _undoPossible=false;
                _focusChangeAt=DateTime.UtcNow.Ticks;
                // Sync tracked layout from actual HKL of new window
                try{
                    uint pid2; uint tid2=GetWindowThreadProcessId(hwnd,out pid2);
                    _trackedRu=(GetKeyboardLayout(tid2).ToInt64()&0xFFFF)==0x0419;
                }catch{}
                Log("FocusChange: buffer cleared, trackedRu="+_trackedRu);
            }
        }

        // Safety: if _converting stuck >2s reset
        if(_converting){
            long ms=(DateTime.UtcNow.Ticks-_convertingAt)/TimeSpan.TicksPerMillisecond;
            if(ms>2000){Log("WATCHDOG reset");_converting=false;BufClear();}
        }
        if(_converting)return (IntPtr)1;

        // Left Ctrl solo tap
        if(vk==_vkUndo){
            if(isDown){_lctrlDown=true;_lctrlCombo=false;_lctrlAt=DateTime.UtcNow.Ticks;}
            else if(isUp){
                if(_lctrlDown&&!_lctrlCombo){
                    long ms=(DateTime.UtcNow.Ticks-_lctrlAt)/TimeSpan.TicksPerMillisecond;
                    if(ms<=SOLO_MS){if(_undoPossible)DoUndo();else DoForce();}
                }
                _lctrlDown=false;
            }
            return CallNextHookEx(_hookId,nCode,wParam,lParam);
        }

        if(!isDown)return CallNextHookEx(_hookId,nCode,wParam,lParam);
        if(_lctrlDown)_lctrlCombo=true;

        if(vk==_vkConvert){bool ctrl=(GetAsyncKeyState(0x11)&0x8000)!=0;if(ctrl)DoExclude();else DoForce();return (IntPtr)1;}

        if(vk==VK_BACK){BufRemoveLast();_undoPossible=false;return CallNextHookEx(_hookId,nCode,wParam,lParam);}
        if(vk==VK_ESCAPE){BufClear();_undoPossible=false;return CallNextHookEx(_hookId,nCode,wParam,lParam);}

        bool shiftHeld=(GetAsyncKeyState(VK_SHIFT)&0x8000)!=0;
        bool capsLock=(GetAsyncKeyState((int)VK_CAPITAL)&0x0001)!=0;
        bool upper=shiftHeld^capsLock;

        // Use self-tracked layout state — immune to stale GetKeyboardLayout() values.
        // During the grace period after focus change, trust _trackedRu only.
        // After grace, also update _trackedRu from A-Z key chars for self-correction.
        bool inGrace=(DateTime.UtcNow.Ticks-_focusChangeAt)<FOCUS_GRACE_TICKS;
        bool isRu=_trackedRu;
        if(!inGrace){
            // Self-correct: if A-Z key produces a Cyrillic char via OS HKL, it's RU
            bool osRu=IsRuLayout();
            if(osRu!=_trackedRu&&(vk>=0x41&&vk<=0x5A)){
                // OS disagrees — trust OS (user may have switched manually via Alt+Shift)
                _trackedRu=osRu; isRu=osRu;
            }
        }

        // ── Space / Enter / Tab are the ONLY word boundaries ───────────────
        // Comma and period are deliberately NOT boundaries: on RU layout
        // they ARE Cyrillic letters (б and ю), and б in particular shows up
        // mid-word in extremely common words (спасибо, добрый, более...).
        // Treating them as boundaries previously caused a real bug: typing
        // "спасибо" would get split into "спаси" + "о" right at the comma
        // key, because the hook can't know at keystroke time whether comma
        // is "the letter б, more word coming" or "actual punctuation, word
        // is done" — that's genuinely ambiguous in the moment. The correct
        // place to resolve it is at detection time (see the trailing-
        // punctuation strip in Detect()), not by guessing here.
        bool isBoundary=(vk==VK_SPACE||vk==VK_RETURN||vk==VK_TAB);

        if(isBoundary){
            string word=BufSnapshotAndClear();_undoPossible=false;
            char bChar=(vk==VK_RETURN)?'\r':(vk==VK_TAB)?'\t':' ';
            if(word.Length>=1&&_autoOn){
                _lastWord=word;_lastBoundary=bChar;
                string conv;string to;
                if(Detect(word,out conv,out to)){
                    _convertingAt=DateTime.UtcNow.Ticks;_converting=true;
                    string w2=word;
                    Task.Run(()=>DoConvert(w2,conv,to,bChar));
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookId,nCode,wParam,lParam);
        }

        // ── Accumulate character ──────────────────────────────────────────
        char bufChar='\0';
        if(vk>=0x41&&vk<=0x5A){
            char bl=(char)(vk+32);char r;
            if(isRu&&_e2r.TryGetValue(bl,out r)){
                bufChar=upper?char.ToUpper(r):r;
                _trackedRu=true;   // confirmed RU layout from Cyrillic output
            } else {
                bufChar=upper?char.ToUpper(bl):bl;
                _trackedRu=false;  // confirmed EN layout from Latin output
            }
        } else {
            string sym=OemBase(vk);
            if(sym!=null){
                char bl=sym[0];char r;
                if(isRu&&_e2r.TryGetValue(bl,out r))bufChar=upper?char.ToUpper(r):r;
                // EN layout: accumulate only strong wrong-layout indicators (;→ж [→х ]→ъ)
                else if(!isRu&&_symKeys.Contains(bl))bufChar=bl;
            }
        }

        if(bufChar!='\0'){
            BufAppend(bufChar);_undoPossible=false;
            if(_diagLog){string snap=BufSnapshot();LogMs("Accum: key='"+bufChar+"' buf=["+snap+"] len="+snap.Length);}
            // Bump on EVERY accumulated keystroke (not just when a live
            // conversion is about to be scheduled) — this is what lets a
            // pending debounced attempt recognize "the user kept typing
            // past this point" and bail out as stale, even if the newer
            // keystrokes themselves didn't trigger a fresh DetectLive hit.
            long thisKeyToken=Interlocked.Increment(ref _liveDebounceToken);

            // ── Live mid-word check (dictionary-only, see DetectLive) ──────
            // Debounced: only actually convert after LIVE_DEBOUNCE_MS of no
            // further keystrokes. This does NOT set _converting and does
            // NOT block the hook — the user can keep typing right through
            // the debounce window exactly as if live-detect didn't exist;
            // any later keystroke bumps the token above and invalidates
            // this pending attempt (see token check below).
            if(_autoOn&&!_converting&&_dictReady){
                string liveWord=BufSnapshot();
                string lconv,lto;
                if(DetectLive(liveWord,out lconv,out lto)){
                    string w2=liveWord;string c2=lconv;string t2=lto;
                    Task.Run(async ()=>{
                        await Task.Delay(LIVE_DEBOUNCE_MS);
                        // Stale if more keys arrived (token moved on) or a
                        // conversion is already in flight from elsewhere.
                        if(Interlocked.Read(ref _liveDebounceToken)!=thisKeyToken)return;
                        if(_converting)return;
                        // Buffer must still hold exactly what we snapshotted —
                        // otherwise backspace/escape/focus-change happened
                        // and this attempt is stale too.
                        if(BufSnapshot()!=w2)return;
                        _convertingAt=DateTime.UtcNow.Ticks;_converting=true;
                        DoConvert(w2,c2,t2,'\0');
                    });
                }
            }
        }
        return CallNextHookEx(_hookId,nCode,wParam,lParam);
    }

    static string OemBase(uint vk){
        switch(vk){
            case 0xBA:return";"; case 0xBC:return","; case 0xBE:return".";
            case 0xBF:return"/"; case 0xC0:return"`"; case 0xDB:return"[";
            case 0xDC:return"\\"; case 0xDD:return"]"; case 0xDE:return"'";
            default:return null;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  LIVE (MID-WORD) DETECTION  —  fires after every keystroke.
    //  Trie-only (no bigrams): bigram ratios need a finished word to be
    //  meaningful, but trie prefix-state is well-defined at any partial
    //  length, which is exactly what makes it safe to use here.
    //
    //  Convert immediately only when BOTH hold:
    //    • current buffer is a DEAD prefix in its own (typed) language
    //      — i.e. no possible completion in that language exists, so the
    //      user could not be about to finish a same-language word
    //    • current buffer's CONVERTED form is a complete word in the
    //      other language with NO further children — i.e. nothing longer
    //      could match, so this can't later turn into a wrong "stump"
    //      (the "кот" vs "который" case: "кот" has children → state 3 →
    //      we wait; only fires once a continuation makes it state 0/2)
    //
    //  Returns false (and leaves buf untouched) in every other case,
    //  including while dictionaries are still loading (_dictReady==false).
    // ══════════════════════════════════════════════════════════════════════
    static bool DetectLive(string buf,out string conv,out string to){
        conv=null;to=null;
        if(!_dictReady)return false;
        if(buf.Length<2)return false; // too short to be meaningfully unambiguous
        if(_excl.Contains(buf.ToLower()))return false;

        string lang=DomLang(buf);
        if(lang==null)return false;

        string candidate=Convert1(buf,lang=="en"?"ru":"en");
        if(!AllAlpha(candidate))return false;

        Trie srcTrie=(lang=="en")?_trieEn:_trieRu;
        Trie dstTrie=(lang=="en")?_trieRu:_trieEn;

        int srcState=srcTrie.PrefixState(buf.ToLowerInvariant());
        // If the typed prefix is still alive (could complete to a real word
        // in its OWN language) we must wait — converting now risks cutting
        // off a legitimate same-language word.
        if(srcState!=0)return false;

        int dstState=dstTrie.PrefixState(candidate.ToLowerInvariant());
        // Only convert when the candidate is a complete word with no
        // possible longer continuation (state 2). State 3 ("кот") must
        // wait for more letters; state 0/1 means it's not a real word
        // (yet, or at all) so there is nothing to convert to.
        if(dstState!=2)return false;

        conv=candidate;to=(lang=="en")?"ru":"en";
        Log("DetectLive: ["+buf+"]->["+candidate+"] dead-end src, clean dst — CONVERT mid-word");
        return true;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  DETECTION  —  ratio-based, verified against simulation table above
    //
    //  Algorithm:
    //    1. Convert word to opposite language
    //    2. If converted result has NO bigrams in target language → reject
    //       (prevents "дшлу", "grey"->пкун, "Let's"->дуеэы)
    //    3. s0==0 (source has no pattern) AND s1>0 → convert
    //       (handles "ltkf", "'nj", "ns", "Иге")
    //    4. s0>0 AND s1/s0 >= 1.3 → convert (was 2.0 — too strict for words like "руку","кудутедуыы")
    //       (handles "Ghbdtn" ratio=5x, "Rfr" ratio=2x exactly)
    //    5. Short word lookup table as fallback for 1-char words
    // ══════════════════════════════════════════════════════════════════════
    static bool Detect(string word,out string conv,out string to){
        conv=null;to=null;

        // Trailing comma/period: by the time we're here, a real boundary
        // (space/Enter/Tab) just fired, so any comma/period sitting at the
        // very END of the buffer is unambiguously punctuation, not the
        // letter б/ю — there's nothing after it for it to be "mid-word"
        // anymore. Peel it off before detection so Convert1 doesn't wrongly
        // re-map it (it otherwise would: ','->'б' and '.'->'ю' in the
        // conversion table, since those mappings are also needed for the
        // genuinely mid-word case, e.g. "добрый" typed with the comma key
        // as the 3rd letter). Re-attach it untouched to whatever gets sent
        // either way, so the on-screen text after Send() still matches.
        string trailingPunct="";
        string core=word;
        while(core.Length>0&&(core[core.Length-1]==','||core[core.Length-1]=='.')){
            trailingPunct=core[core.Length-1]+trailingPunct;
            core=core.Substring(0,core.Length-1);
        }
        if(core.Length==0)return false; // was nothing but punctuation
        if(_excl.Contains(core.ToLower()))return false;

        // Determine dominant language (ignoring leading/trailing apostrophes)
        string stripped=core.Trim('\'');
        if(stripped.Length==0)return false;
        string lang=DomLang(stripped);
        if(lang==null)return false;

        // Single char: use known-word table
        if(core.Length==1){
            if(lang=="en"){string c1=Convert1(core,"ru");if(_knownRu.Contains(c1.ToLower())&&!_knownEn.Contains(core.ToLower())){conv=c1+trailingPunct;to="ru";return true;}}
            else{string c1=Convert1(core,"en");if(_knownEn.Contains(c1.ToLower())&&!_knownRu.Contains(core.ToLower())){conv=c1+trailingPunct;to="en";return true;}}
            return false;
        }

        word=core; // rest of the function works on the punctuation-free core

        // Build candidate
        string candidate;double s0,s1;
        if(lang=="en"){
            candidate=Convert1(word,"ru");
            if(!AllAlpha(candidate)||DomLang(candidate)!="ru")return false;
        } else {
            candidate=Convert1(word,"en");
            if(!AllAlpha(candidate)||DomLang(candidate)!="en")return false;
        }
        // candidate stays PUNCTUATION-FREE for all lookups below (dictionary
        // ContainsWord/Rank, bigram scoring) — trailingPunct only gets
        // reattached at each point where conv is actually assigned for
        // output, so "дорый" matches the dictionary correctly and
        // "дорый," never gets looked up as if the comma were part of it.

        // ── Dictionary check first (authoritative when available) ─────────
        // If the word as TYPED is already a real word in its own language,
        // never convert — this is the strongest possible signal and skips
        // bigram math entirely (handles real words bigrams might otherwise
        // misjudge, e.g. short or unusual-but-valid words).
        // If the typed word is NOT a real word but the converted candidate
        // IS a real word in the target language, convert immediately —
        // again skipping bigram math, since dictionary membership is more
        // reliable than statistical letter-pair scoring.
        // If dictionary is inconclusive for both (proper nouns, slang,
        // typos, or dictionaries not loaded yet) fall through unchanged to
        // the original bigram/ratio logic below.
        if(_dictReady){
            string wLower=word.ToLowerInvariant();
            Trie srcTrie=(lang=="en")?_trieEn:_trieRu;
            Trie dstTrie=(lang=="en")?_trieRu:_trieEn;
            bool srcReal=srcTrie.ContainsWord(wLower);
            bool dstReal=dstTrie.ContainsWord(candidate.ToLowerInvariant());
            if(srcReal&&!dstReal){Log("Dict: ["+word+"] real in "+lang+" — no convert");return false;}
            if(dstReal&&!srcReal){
                to=(lang=="en")?"ru":"en";
                conv=candidate+trailingPunct;
                Log("Dict: ["+word+"]->["+candidate+"] real in "+to+" — CONVERT");
                return true;
            }
            if(srcReal&&dstReal){
                // Both are real words — genuinely ambiguous about intent.
                // Only override the conservative "no convert" default when
                // frequency data exists AND shows a large enough gap that
                // the typed word is implausible compared to the candidate
                // (e.g. typed word is rare/unranked, candidate is a top
                // few hundred common word). A modest gap is not enough —
                // this only fires for lopsided cases, and is a complete
                // no-op whenever no frequency files were supplied (every
                // word is then int.MaxValue, so the gap check never
                // passes and behavior is identical to before).
                int srcRank=srcTrie.Rank(wLower);
                int dstRank=dstTrie.Rank(candidate.ToLowerInvariant());
                const int FREQ_GAP=2000; // src must be at least this much rarer than dst
                const int FREQ_DST_MAX=500; // ...and dst itself must be a top-500 common word
                if(dstRank<FREQ_DST_MAX&&srcRank>dstRank+FREQ_GAP){
                    to=(lang=="en")?"ru":"en";
                    conv=candidate+trailingPunct;
                    Log("Dict: ["+word+"]->["+candidate+"] both real, freq override (src="+srcRank+" dst="+dstRank+") — CONVERT");
                    return true;
                }
                Log("Dict: ["+word+"] both real (src="+srcRank+" dst="+dstRank+") — no convert");
                return false;
            }
            // both unreal: ambiguous — fall through to bigrams
        }

        if(lang=="en"){
            s0=Bgram(word.ToLower(),_enBG); s1=Bgram(candidate.ToLower(),_ruBG); to="ru";
        } else {
            s0=Bgram(word.ToLower(),_ruBG); s1=Bgram(candidate.ToLower(),_enBG); to="en";
        }

        Log("Detect ["+word+"] s0="+s0.ToString("F2")+" ["+candidate+"] s1="+s1.ToString("F2"));

        // Reject if target has zero recognizable patterns
        if(s1<=0.0){Log("  s1=0 reject");return false;}

        // s0==0: source has no pattern at all in current language → convert
        if(s0<=0.0&&s1>=0.10){Log("  s0=0 CONVERT");conv=candidate+trailingPunct;return true;}

        // Ratio: target must score at least 2x better
        if(s0>0.0&&s1/s0>=1.3){Log("  ratio="+(s1/s0).ToString("F1")+" CONVERT");conv=candidate+trailingPunct;return true;}

        Log("  no conversion");return false;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  SENDINPUT & CONVERSION
    // ══════════════════════════════════════════════════════════════════════
    static INPUT IVkD(ushort vk){var i=new INPUT();i.type=INPUT_KEYBOARD;i.u.ki.wVk=vk;return i;}
    static INPUT IVkU(ushort vk){var i=new INPUT();i.type=INPUT_KEYBOARD;i.u.ki.wVk=vk;i.u.ki.dwFlags=KEYEVENTF_KEYUP;return i;}
    static INPUT IUcD(char c){var i=new INPUT();i.type=INPUT_KEYBOARD;i.u.ki.wScan=(ushort)c;i.u.ki.dwFlags=KEYEVENTF_UNICODE;return i;}
    static INPUT IUcU(char c){var i=new INPUT();i.type=INPUT_KEYBOARD;i.u.ki.wScan=(ushort)c;i.u.ki.dwFlags=KEYEVENTF_UNICODE|KEYEVENTF_KEYUP;return i;}

    static void Send(int del,string text,char bnd){
        bool hasBnd=(bnd!='\0');
        bool vkBnd=(bnd==' '||bnd=='\r'||bnd=='\t');
        ushort bvk=(bnd=='\r')?VK_RETURN:(bnd=='\t')?VK_TAB:VK_SPACE;
        int tot=del*2+text.Length*2+(hasBnd?2:0);
        var inp=new INPUT[tot];int idx=0;
        for(int i=0;i<del;i++){inp[idx++]=IVkD(VK_BACK);inp[idx++]=IVkU(VK_BACK);}
        foreach(char c in text){inp[idx++]=IUcD(c);inp[idx++]=IUcU(c);}
        if(hasBnd){if(vkBnd){inp[idx++]=IVkD(bvk);inp[idx++]=IVkU(bvk);}else{inp[idx++]=IUcD(bnd);inp[idx++]=IUcU(bnd);}}
        uint sent=SendInput((uint)tot,inp,_inputSize);
        Log("Send "+sent+"/"+tot+" del="+del+" bnd='"+(bnd=='\0'?'0':bnd)+"'");
    }

    static void DoConvert(string word,string converted,string to,char bChar){
        Log("DoConvert ["+word+"]->["+converted+"]");
        try{Thread.Sleep(30);Send(word.Length,converted,bChar);SwitchLang(to=="ru");
            _undoOrig=word;_undoConv=converted;_undoBnd=bChar;_undoPossible=true;Log("Done.");}
        catch(Exception ex){Log("DoConvert err:"+ex.Message);}
        finally{BufClear();_converting=false;}
    }

    static void DoUndo(){
        if(!_undoPossible||_undoOrig==null)return;
        string orig=_undoOrig,conv=_undoConv;char bChar=_undoBnd;
        string origLang=DomLang(orig);
        _undoPossible=false;_undoOrig=null;_undoConv=null;
        _convertingAt=DateTime.UtcNow.Ticks;_converting=true;
        Log("Undo ["+conv+"]->["+orig+"]");
        Task.Run(()=>{
            try{Thread.Sleep(30);bool hasBnd=(bChar!='\0');Send(hasBnd?conv.Length+1:conv.Length,orig,bChar);SwitchLang(origLang=="ru");}
            catch(Exception ex){Log("Undo err:"+ex.Message);}
            finally{BufClear();_converting=false;}
        });
    }

    static void DoForce(){
        if(_converting)return;
        string rawWord;char bChar;bool hasBuf=(BufLength()>0);
        if(hasBuf){rawWord=BufSnapshot();bChar='\0';}
        else if(_lastWord.Length>0){rawWord=_lastWord;bChar=_lastBoundary;}
        else return;

        // If the buffer is mixed (e.g. "Мvlookup" = leftover Cyrillic + new Latin),
        // extract only the trailing run of consistent language for conversion.
        // We still delete the FULL rawWord length to clean up the screen completely.
        string word=TailRun(rawWord);
        if(word.Length==0)return;

        string lang=DomLang(word);if(lang==null)return;
        string to=(lang=="ru")?"en":"ru";
        string conv=Convert1(word,to);if(conv==word)return;
        Log("Force raw=["+rawWord+"] word=["+word+"]->["+conv+"]");
        BufClear();_lastWord=rawWord;_convertingAt=DateTime.UtcNow.Ticks;_converting=true;
        int del=hasBuf?rawWord.Length:rawWord.Length+1;char sndBnd=hasBuf?'\0':bChar;
        Task.Run(()=>{
            try{Thread.Sleep(30);Send(del,conv,sndBnd);SwitchLang(to=="ru");
                _undoOrig=rawWord;_undoConv=conv;_undoBnd=sndBnd;_undoPossible=true;}
            catch(Exception ex){Log("Force err:"+ex.Message);}
            finally{BufClear();_converting=false;}
        });
    }

    // Returns the longest trailing substring where all alpha chars are the same language.
    // "Мvlookup" → "vlookup"   "мдщщлгз" → "мдщщлгз"   "hello" → "hello"
    static string TailRun(string s){
        if(s.Length==0)return s;
        // Determine language of the last alpha char
        bool tailRu=false,tailEn=false;
        for(int i=s.Length-1;i>=0;i--){
            char c=s[i];
            if(IsRu(c)){tailRu=true;break;}
            if(IsEn(c)){tailEn=true;break;}
        }
        if(!tailRu&&!tailEn)return s; // no alpha at all
        // Walk back to find where language changes
        int start=0;
        for(int i=0;i<s.Length;i++){
            char c=s[i];
            if(tailRu&&IsEn(c)){start=i+1;}  // EN char found in RU tail → reset start
            if(tailEn&&IsRu(c)){start=i+1;}  // RU char found in EN tail → reset start
        }
        return s.Substring(start);
    }

    static void DoExclude(){
        string word=(BufLength()>0)?BufSnapshot():_lastWord;
        if(word.Length==0)return;
        _excl.Add(word.ToLower());SaveExcl();Log("Excluded: "+word);
    }

    // ── Language helpers ──────────────────────────────────────────────────
    static bool IsRu(char c){return(c>='\u0410'&&c<='\u044F')||c=='\u0451'||c=='\u0401';}
    static bool IsEn(char c){return(c>='a'&&c<='z')||(c>='A'&&c<='Z');}
    static bool AllAlpha(string s){
        if(s.Length==0)return false;
        foreach(char c in s)if(!IsRu(c)&&!IsEn(c))return false;
        return true;
    }
    static string DomLang(string w){
        int ru=0,en=0;
        foreach(char c in w){if(IsRu(c))ru++;else if(IsEn(c))en++;}
        int t=ru+en;if(t==0)return null;
        if((double)ru/t>=0.6)return"ru";if((double)en/t>=0.6)return"en";return null;
    }
    static double Bgram(string w,HashSet<string> bg){
        // w should already be lowercase
        int n=w.Length-1;if(n<1)return 0;
        int h=0;for(int i=0;i<n;i++)if(bg.Contains(w.Substring(i,2)))h++;
        return(double)h/n;
    }
    static string Convert1(string word,string to){
        var sb=new StringBuilder(word.Length);
        foreach(char c in word){
            char lc=char.ToLower(c);bool up=char.IsUpper(c);char m;
            if(to=="ru"&&_e2r.TryGetValue(lc,out m))sb.Append(up?char.ToUpper(m):m);
            else if(to=="en"&&_r2e.TryGetValue(lc,out m))sb.Append(up?char.ToUpper(m):m);
            else sb.Append(c);
        }
        return sb.ToString();
    }

    static bool IsRuLayout(){
        try{IntPtr hwnd=GetForegroundWindow();uint pid;uint tid=GetWindowThreadProcessId(hwnd,out pid);
            return(GetKeyboardLayout(tid).ToInt64()&0xFFFF)==0x0419;}catch{return false;}
    }
    static readonly System.Collections.Concurrent.ConcurrentDictionary<System.Threading.Timer,byte> _pendingLangTimers=new System.Collections.Concurrent.ConcurrentDictionary<System.Threading.Timer,byte>();
    static void SwitchLang(bool toRu){
        _trackedRu=toRu;
        // Fire on a short timer instead of blocking the calling thread with
        // Thread.Sleep — this makes the OS layout indicator flip feel
        // instant to the user (the calling thread returns immediately)
        // while still giving slower target apps the same ~30ms window to
        // finish processing the injected keystrokes before the layout
        // actually changes underneath them.
        // Each timer roots itself in _pendingLangTimers until it fires, so
        // rapid successive conversions (e.g. live mid-word triggers) can't
        // have an earlier still-pending timer collected before it runs.
        IntPtr hwnd=GetForegroundWindow();
        IntPtr hkl=toRu?_hklRu:_hklEn;
        System.Threading.Timer timer=null;
        timer=new System.Threading.Timer(_=>{
            try{PostMessage(hwnd,WM_INPUTLANGCHANGEREQUEST,IntPtr.Zero,hkl);}catch{}
            byte b;_pendingLangTimers.TryRemove(timer,out b);
            timer.Dispose();
        },null,30,System.Threading.Timeout.Infinite);
        _pendingLangTimers[timer]=0;
    }

    // ── Config/Exclusions ─────────────────────────────────────────────────
    static void LoadConfig(){
        if(!File.Exists(_cfgFile))return;
        try{foreach(var line in File.ReadAllLines(_cfgFile)){
            string[]p=line.Split('=');if(p.Length!=2)continue;
            uint v;if(!uint.TryParse(p[1].Trim(),out v))continue;
            switch(p[0].Trim()){case"VkConvert":_vkConvert=v;_vkExclude=v;break;case"VkUndo":_vkUndo=v;break;case"VkSettings":_vkSettings=v;break;case"AutoOn":_autoOn=(v==1);break;}
        }}catch{}
    }
    static void SaveConfig(){try{File.WriteAllText(_cfgFile,"VkConvert="+_vkConvert+"\r\nVkUndo="+_vkUndo+"\r\nVkSettings="+_vkSettings+"\r\nAutoOn="+(_autoOn?1:0)+"\r\n");}catch{}}
    static void LoadExcl(){if(!File.Exists(_exclFile))return;foreach(var l in File.ReadAllLines(_exclFile,Encoding.UTF8))if(!string.IsNullOrWhiteSpace(l))_excl.Add(l.Trim().ToLower());}
    static void SaveExcl(){try{File.WriteAllLines(_exclFile,_excl,Encoding.UTF8);}catch{}}

    // Reflects whatever dictionary-load state has been reached so far —
    // called both when the tray menu is first built and after the
    // background loader finishes, so neither ordering can leave the menu
    // stuck on a stale "loading..." label.
    static string CurrentDictLabel(){
        if(_dictReady)return "Dictionary: ON  ("+_trieEn.WordCount+" EN / "+_trieRu.WordCount+" RU)";
        if(_dictAttempted)return "Dictionary: OFF (not found / load error)";
        return "Dictionary: loading...";
    }

    // ── Tray ──────────────────────────────────────────────────────────────
    static void BuildTray(){
        _tray=new NotifyIcon();_tray.Icon=MakeIcon(true);_tray.Text="Switcher \u2014 EN\u2194RU";_tray.Visible=true;
        var menu=new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem("Switcher  EN \u2194 RU"){Enabled=false});
        menu.Items.Add(new ToolStripSeparator());
        _toggleItem=new ToolStripMenuItem("Auto-detect:  ON  \u2713");
        _toggleItem.Click+=(s,e)=>{_autoOn=!_autoOn;_toggleItem.Text=_autoOn?"Auto-detect:  ON  \u2713":"Auto-detect:  OFF";_tray.Icon=MakeIcon(_autoOn);SaveConfig();};
        menu.Items.Add(_toggleItem);
        _dictStatusItem=new ToolStripMenuItem(CurrentDictLabel());_dictStatusItem.Enabled=false;
        menu.Items.Add(_dictStatusItem);
        menu.Items.Add(new ToolStripSeparator());
        var si=new ToolStripMenuItem("Settings / Hotkeys...  (Win+S)");si.Click+=(s,e)=>OpenSettings();menu.Items.Add(si);
        var ti=new ToolStripMenuItem("Test Converter...");ti.Click+=(s,e)=>OpenTestConverter();menu.Items.Add(ti);
        menu.Items.Add(new ToolStripSeparator());
        var exm=new ToolStripMenuItem("Exclusions");
        var ae=new ToolStripMenuItem("Add Exclusion...");ae.Click+=(s,e)=>OpenAddExcl();exm.DropDownItems.Add(ae);
        var re=new ToolStripMenuItem("Review / Edit Exclusions...");re.Click+=(s,e)=>OpenRevExcl();exm.DropDownItems.Add(re);
        menu.Items.Add(exm);
        var ol=new ToolStripMenuItem("Open Log");ol.Click+=(s,e)=>{try{System.Diagnostics.Process.Start(_logFile);}catch{}};menu.Items.Add(ol);
        menu.Items.Add(new ToolStripSeparator());
        var qi=new ToolStripMenuItem("Quit");qi.Click+=(s,e)=>{UnhookWindowsHookEx(_hookId);_tray.Visible=false;Application.Exit();};menu.Items.Add(qi);
        _tray.ContextMenuStrip=menu;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  TEST CONVERTER
    //  Paste text → see word-by-word what the detection engine does.
    //  Green  = would convert (shows what it becomes)
    //  Red    = would NOT convert (shows why: s0, s1, ratio)
    //  Grey   = skipped (punctuation / too short / excluded)
    // ══════════════════════════════════════════════════════════════════════
    static void OpenTestConverter(){
        var f=new Form{
            Text="Switcher — Test Converter",
            Size=new Size(820,600),
            MinimumSize=new Size(600,400),
            StartPosition=FormStartPosition.CenterScreen,
            Font=new Font("Segoe UI",9f)
        };

        // ── Input panel (top) ─────────────────────────────────────────────
        var lblIn=new Label{Text="Paste text to test (as if typed in WRONG layout):",
            Location=new Point(10,10),Size=new Size(500,18)};
        f.Controls.Add(lblIn);

        var tbIn=new TextBox{
            Multiline=true,ScrollBars=ScrollBars.Vertical,
            Location=new Point(10,32),Size=new Size(f.ClientSize.Width-20,100),
            Font=new Font("Consolas",10f),
            Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right
        };
        f.Controls.Add(tbIn);

        var btnRun=new Button{
            Text="▶  Analyse",Location=new Point(10,140),Size=new Size(110,30),
            BackColor=Color.FromArgb(34,150,80),ForeColor=Color.White,
            FlatStyle=FlatStyle.Flat,Font=new Font("Segoe UI",9f,FontStyle.Bold)
        };
        f.Controls.Add(btnRun);

        var btnClear=new Button{
            Text="Clear",Location=new Point(130,140),Size=new Size(70,30)
        };
        f.Controls.Add(btnClear);

        var lblStats=new Label{
            Text="",Location=new Point(215,147),
            Size=new Size(f.ClientSize.Width-225,20),
            ForeColor=Color.Gray,
            Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right
        };
        f.Controls.Add(lblStats);

        // ── Results grid ──────────────────────────────────────────────────
        var grid=new DataGridView{
            Location=new Point(10,178),
            Size=new Size(f.ClientSize.Width-20,f.ClientSize.Height-188),
            Anchor=AnchorStyles.Top|AnchorStyles.Bottom|AnchorStyles.Left|AnchorStyles.Right,
            ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,
            RowHeadersVisible=false,SelectionMode=DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor=Color.White,BorderStyle=BorderStyle.None,
            Font=new Font("Consolas",10f),AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill
        };
        grid.Columns.Add("word",    "Input word");
        grid.Columns.Add("result",  "Result");
        grid.Columns.Add("action",  "Action");
        grid.Columns.Add("s0",      "s0 (src)");
        grid.Columns.Add("s1",      "s1 (tgt)");
        grid.Columns.Add("ratio",   "ratio");
        grid.Columns["word"].FillWeight=18;
        grid.Columns["result"].FillWeight=18;
        grid.Columns["action"].FillWeight=28;
        grid.Columns["s0"].FillWeight=9;
        grid.Columns["s1"].FillWeight=9;
        grid.Columns["ratio"].FillWeight=9;
        grid.ColumnHeadersDefaultCellStyle.Font=new Font("Segoe UI",9f,FontStyle.Bold);
        grid.DefaultCellStyle.Padding=new Padding(4,2,4,2);
        grid.RowTemplate.Height=22;
        grid.CellFormatting+=(s,e)=>{
            if(e.RowIndex<0)return;
            string action=grid.Rows[e.RowIndex].Cells["action"].Value as string;
            if(action==null)return;
            if(action.StartsWith("CONVERT")){
                e.CellStyle.BackColor=Color.FromArgb(220,255,220);
                e.CellStyle.ForeColor=Color.FromArgb(0,100,0);
            } else if(action=="SKIP"||action=="TOO SHORT"||action=="EXCLUDED"){
                e.CellStyle.BackColor=Color.FromArgb(240,240,240);
                e.CellStyle.ForeColor=Color.Gray;
            } else {
                e.CellStyle.BackColor=Color.FromArgb(255,230,230);
                e.CellStyle.ForeColor=Color.FromArgb(150,0,0);
            }
        };
        f.Controls.Add(grid);

        // ── Logic ─────────────────────────────────────────────────────────
        btnClear.Click+=(s,e)=>{tbIn.Clear();grid.Rows.Clear();lblStats.Text="";};

        btnRun.Click+=(s,e)=>{
            grid.Rows.Clear();
            string text=tbIn.Text.Trim();
            if(text.Length==0)return;

            // Split on whitespace — preserve each token
            string[] tokens=text.Split(new char[]{' ','\t','\r','\n'},StringSplitOptions.RemoveEmptyEntries);

            int total=0,converted=0,rejected=0,skipped=0;
            foreach(string raw in tokens){
                // Strip surrounding punctuation to get the testable word
                // Strip surrounding punctuation — but NOT ';' or ':' which map to Cyrillic letters (ж)
                string word=raw.Trim('.', ',', '!', '?', '(', ')', '[', ']', '"', '\'', '\u2014', '\u2013', '-', '\u2026');
                if(word.Length==0){skipped++;AddRow(grid,raw,"\u2014","SKIP (punct only)","\u2014","\u2014","\u2014");continue;}

                total++;
                bool hasAlpha=false;
                foreach(char c in word)if(IsRu(c)||IsEn(c)){hasAlpha=true;break;}
                if(!hasAlpha){skipped++;AddRow(grid,raw,"\u2014","SKIP (no alpha)","\u2014","\u2014","\u2014");continue;}
                if(word.Length<2){skipped++;AddRow(grid,raw,"\u2014","TOO SHORT","\u2014","\u2014","\u2014");continue;}
                if(_excl.Contains(word.ToLower())){skipped++;AddRow(grid,raw,"\u2014","EXCLUDED","\u2014","\u2014","\u2014");continue;}

                string stripped=word.Trim('\'');
                string lang=DomLang(stripped);
                if(lang==null){skipped++;AddRow(grid,raw,"\u2014","SKIP (mixed)","\u2014","\u2014","\u2014");continue;}

                string cand; string toLang;
                if(lang=="en"){cand=Convert1(word,"ru");toLang="ru";}
                else{cand=Convert1(word,"en");toLang="en";}

                if(!AllAlpha(cand)||DomLang(cand)!=toLang){
                    rejected++;
                    AddRow(grid,raw,"\u2014","REJECT (bad conversion)","\u2014","\u2014","\u2014");
                    continue;
                }

                // ── Layer 1: layout mismatch simulation ──────────────────
                // In the live app, typing Cyrillic in EN layout (or Latin in RU layout)
                // is caught directly without scoring. The test simulates both cases.
                bool layer1En = (lang=="en"); // Latin chars → would convert to RU if in RU layout
                bool layer1Ru = (lang=="ru"); // Cyrillic chars → would convert to EN if in EN layout
                // Layer 1 fires when lang mismatches what the layout would produce.
                // We show this as LAYER1 to distinguish from bigram-only conversion.

                double s0=lang=="en"?Bgram(word.ToLower(),_enBG):Bgram(word.ToLower(),_ruBG);
                double s1=toLang=="ru"?Bgram(cand.ToLower(),_ruBG):Bgram(cand.ToLower(),_enBG);

                string action; string resultWord;

                // Layer 2: mapped symbols (;,[,] in EN word = strong wrong-layout signal)
                bool hasSym=false;
                if(lang=="en"){foreach(char c in word)if(_symKeys.Contains(c)){hasSym=true;break;}}

                if(hasSym){
                    converted++;action="CONVERT (sym key)";resultWord=cand;
                } else if(s1<=0.0){
                    // No target bigrams → candidate is not a real word in target lang.
                    // BUT Layer 1 still fires in live app for layout-mismatched input!
                    if(layer1En||layer1Ru){
                        converted++;action="CONVERT (L1 mismatch, s1=0)";resultWord=cand;
                    } else {
                        rejected++;action="REJECT  s1=0";resultWord="\u2014";
                    }
                } else if(s0<=0.0&&s1>=0.10){
                    converted++;action="CONVERT  s0=0 s1="+s1.ToString("F2");resultWord=cand;
                } else if(s0>0.0&&s1/s0>=1.3){
                    converted++;action="CONVERT  ratio="+(s1/s0).ToString("F1")+"x";resultWord=cand;
                } else {
                    // Layer 1 still fires in live app even when bigrams don't confirm
                    if(layer1En||layer1Ru){
                        converted++;action="CONVERT (L1 mismatch, ratio="+(s0>0?(s1/s0).ToString("F1")+"x":"n/a")+")";resultWord=cand;
                    } else {
                        rejected++;action="REJECT  ratio="+(s0>0?(s1/s0).ToString("F1")+"x":"\u2014");resultWord="\u2014";
                    }
                }

                AddRow(grid,raw,resultWord,action,s0.ToString("F2"),s1.ToString("F2"),s0>0?(s1/s0).ToString("F1"):"\u221e");
            }
            lblStats.Text="Total: "+total+"  |  Would convert: "+converted+"  |  Bigram-only rejects: "+rejected+"  |  Skipped: "+skipped;
        };

        f.Resize+=(s,e)=>{
            tbIn.Width=f.ClientSize.Width-20;
            lblStats.Width=f.ClientSize.Width-225;
            grid.Size=new Size(f.ClientSize.Width-20,f.ClientSize.Height-188);
        };

        f.Show(); // non-modal — stays open while you test
    }

    static void AddRow(DataGridView g,string word,string result,string action,string s0,string s1,string ratio){
        int i=g.Rows.Add();
        g.Rows[i].Cells["word"].Value=word;
        g.Rows[i].Cells["result"].Value=result;
        g.Rows[i].Cells["action"].Value=action;
        g.Rows[i].Cells["s0"].Value=s0;
        g.Rows[i].Cells["s1"].Value=s1;
        g.Rows[i].Cells["ratio"].Value=ratio;
    }

    static void OpenAddExcl(){
        var f=new Form{Text="Add Exclusion",Size=new Size(360,150),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false,StartPosition=FormStartPosition.CenterScreen,Font=new Font("Segoe UI",10f)};
        f.Controls.Add(new Label{Text="Word to exclude from auto-conversion:",Location=new Point(16,16),Size=new Size(316,20)});
        var tb=new TextBox{Location=new Point(16,42),Size=new Size(316,26),Font=new Font("Segoe UI",11f)};f.Controls.Add(tb);tb.Select();
        var ba=new Button{Text="Add",Location=new Point(156,80),Size=new Size(80,28)};
        ba.Click+=(s,e)=>{string w=tb.Text.Trim().ToLower();if(w.Length>0){_excl.Add(w);SaveExcl();}f.Close();};f.Controls.Add(ba);f.AcceptButton=ba;
        var bc=new Button{Text="Cancel",Location=new Point(246,80),Size=new Size(80,28)};bc.Click+=(s,e)=>f.Close();f.Controls.Add(bc);f.CancelButton=bc;f.ShowDialog();
    }

    static void OpenRevExcl(){
        var f=new Form{Text="Exclusions \u2014 one word per line",Size=new Size(380,400),FormBorderStyle=FormBorderStyle.Sizable,MinimumSize=new Size(280,260),StartPosition=FormStartPosition.CenterScreen,Font=new Font("Segoe UI",9f)};
        f.Controls.Add(new Label{Text="One word per line. Save to apply.",Location=new Point(12,10),Size=new Size(340,18),ForeColor=Color.Gray});
        var tb=new TextBox{Multiline=true,ScrollBars=ScrollBars.Vertical,Location=new Point(12,34),Size=new Size(340,300),Font=new Font("Consolas",10f),Anchor=AnchorStyles.Top|AnchorStyles.Bottom|AnchorStyles.Left|AnchorStyles.Right};
        var sl=new List<string>(_excl);sl.Sort();tb.Text=string.Join("\r\n",sl.ToArray());f.Controls.Add(tb);
        var bs=new Button{Text="Save",Size=new Size(80,28),Anchor=AnchorStyles.Bottom|AnchorStyles.Right};
        bs.Location=new Point(f.ClientSize.Width-92,f.ClientSize.Height-36);
        bs.Click+=(s,e)=>{_excl.Clear();foreach(var line in tb.Text.Split('\n')){string w=line.Trim().ToLower();if(w.Length>0)_excl.Add(w);}SaveExcl();f.Close();};f.Controls.Add(bs);f.AcceptButton=bs;
        var bc=new Button{Text="Cancel",Size=new Size(80,28),Anchor=AnchorStyles.Bottom|AnchorStyles.Right};
        bc.Location=new Point(f.ClientSize.Width-182,f.ClientSize.Height-36);bc.Click+=(s,e)=>f.Close();f.Controls.Add(bc);f.CancelButton=bc;
        f.Resize+=(s,e)=>{tb.Size=new Size(f.ClientSize.Width-24,f.ClientSize.Height-76);bs.Location=new Point(f.ClientSize.Width-92,f.ClientSize.Height-36);bc.Location=new Point(f.ClientSize.Width-182,f.ClientSize.Height-36);};
        f.ShowDialog();
    }

    static void OpenSettings(){
        var f=new Form{Text="Switcher \u2014 Settings",Size=new Size(430,265),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,StartPosition=FormStartPosition.CenterScreen,Font=new Font("Segoe UI",9f)};
        int y=16;f.Controls.Add(new Label{Text="Click a box, then press the key you want:",Location=new Point(16,y),Size=new Size(390,18),ForeColor=Color.Gray});y+=28;
        TextBox tbC=null,tbU=null,tbS=null;
        Action<string,uint,Action<TextBox>> addRow=(label,cur,setter)=>{
            f.Controls.Add(new Label{Text=label,Location=new Point(16,y+3),Size=new Size(210,20)});
            var tb=new TextBox{Text=VkName(cur),Location=new Point(228,y),Size=new Size(168,22),ReadOnly=true,Tag=(object)cur};
            tb.KeyDown+=(s,e)=>{e.SuppressKeyPress=true;uint vk=(uint)e.KeyCode;tb.Text=VkName(vk);tb.Tag=(object)vk;};
            tb.GotFocus+=(s,e)=>tb.BackColor=Color.LightYellow;tb.LostFocus+=(s,e)=>tb.BackColor=Color.White;
            f.Controls.Add(tb);setter(tb);y+=34;
        };
        addRow("Convert / Force  (Ctrl = exclude):",_vkConvert,tb=>tbC=tb);
        addRow("Undo / Force  (Left Ctrl, solo tap):",_vkUndo,tb=>tbU=tb);
        addRow("Open settings  (Win + key):",_vkSettings,tb=>tbS=tb);
        y+=8;f.Controls.Add(new Label{Text="Left Ctrl: 1st tap = force-convert, 2nd tap = undo.",Location=new Point(16,y),Size=new Size(390,18),ForeColor=Color.Gray});y+=28;
        var bs=new Button{Text="Save",Location=new Point(228,y),Size=new Size(80,28)};
        bs.Click+=(s,e)=>{_vkConvert=(uint)tbC.Tag;_vkExclude=_vkConvert;_vkUndo=(uint)tbU.Tag;_vkSettings=(uint)tbS.Tag;SaveConfig();f.Close();};f.Controls.Add(bs);
        var bc=new Button{Text="Cancel",Location=new Point(318,y),Size=new Size(80,28)};bc.Click+=(s,e)=>f.Close();f.Controls.Add(bc);f.ShowDialog();
    }

    static string VkName(uint vk){
        switch(vk){case 0x91:return"Scroll Lock";case 0xA2:return"Left Ctrl";case 0xA0:return"Left Shift";case 0xA4:return"Left Alt";case 0x14:return"Caps Lock";
            default:if(vk>=0x70&&vk<=0x87)return"F"+(vk-0x6F);if(vk>=0x30&&vk<=0x39)return((char)vk).ToString();if(vk>=0x41&&vk<=0x5A)return((char)vk).ToString();return"Key("+vk+")";}
    }

    static Icon MakeIcon(bool active){
        int sz=64;var bmp=new Bitmap(sz,sz);
        using(var g=Graphics.FromImage(bmp)){
            g.SmoothingMode=SmoothingMode.AntiAlias;g.Clear(Color.Transparent);
            Color bg=active?Color.FromArgb(255,224,0):Color.FromArgb(150,150,150);
            using(var path=RR(new Rectangle(2,2,sz-4,sz-4),12))using(var br=new SolidBrush(bg))g.FillPath(br,path);
            using(var path=RR(new Rectangle(2,2,sz-4,sz-4),12))using(var pen=new Pen(Color.FromArgb(60,255,255,255),2f))g.DrawPath(pen,path);
            Color pc=active?Color.FromArgb(28,155,70):Color.FromArgb(80,80,80);
            float cx=sz/2f,cy=sz/2f+4f,r=sz*0.27f,pw=sz*0.09f;
            using(var pen=new Pen(pc,pw)){pen.StartCap=LineCap.Round;pen.EndCap=LineCap.Round;
                g.DrawArc(pen,cx-r,cy-r,r*2,r*2,130f,280f);g.DrawLine(pen,cx,cy-r*0.35f,cx,cy-r*1.35f);}
        }
        return Icon.FromHandle(bmp.GetHicon());
    }
    static GraphicsPath RR(Rectangle r,int rad){
        var p=new GraphicsPath();
        p.AddArc(r.X,r.Y,rad*2,rad*2,180,90);p.AddArc(r.Right-rad*2,r.Y,rad*2,rad*2,270,90);
        p.AddArc(r.Right-rad*2,r.Bottom-rad*2,rad*2,rad*2,0,90);p.AddArc(r.X,r.Bottom-rad*2,rad*2,rad*2,90,90);
        p.CloseFigure();return p;
    }
}
} // namespace
'@

try {
    Add-Type -TypeDefinition $source `
        -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' `
        -Language CSharp `
        -ErrorAction Stop
    Write-Log "Compiled OK."
    [Switcher.App]::Run($logFile)
}
catch {
    $err = $_.Exception.Message
    Write-Log "COMPILE ERROR: $err"
    [System.Windows.Forms.MessageBox]::Show(
        "Switcher failed to start:`n`n$err`n`nLog: $logFile",
        "Switcher Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

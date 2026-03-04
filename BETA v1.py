"""
VALORANT Sensitivity Tracker v2
・ホットキー → スクショ → OCR → 感度自動記録
・ホットキーはアプリ内で自由に変更可能
・データはアプリを閉じても保持（SQLite）
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3, os, re, threading, webbrowser, json, ctypes, io, base64, math
from datetime import datetime
from pathlib import Path
from collections import Counter

# ── オプション依存 ───────────────────────────────
try:
    from PIL import Image, ImageGrab, ImageEnhance, ImageFilter
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pytesseract
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import pynput.keyboard as _pynput_kb
    HAS_PYNPUT = True
except ImportError:
    HAS_PYNPUT = False

try:
    import urllib.request
    HAS_URLLIB = True
except ImportError:
    HAS_URLLIB = False


# ─────────────────────────────────────────────
#  パス・設定
# ─────────────────────────────────────────────
APP_DIR = Path(os.path.expanduser("~")) / ".valo_sens_tracker"
APP_DIR.mkdir(exist_ok=True)
DB_PATH  = APP_DIR / "data.db"
CFG_PATH = APP_DIR / "config.json"


DEFAULT_CFG = {
    "hotkey_parts": ["ctrl", "shift", "s"],   # キーのリスト
    "dpi": 800,
    "tesseract_path": "",
    "henrik_api_key": "",
    "player_name": "",
    "player_tag":  "",
    "region":      "ap",
    "setup_done":  False,  
}

def load_cfg():
    if CFG_PATH.exists():
        try:
            d = json.loads(CFG_PATH.read_text(encoding="utf-8"))
            cfg = DEFAULT_CFG.copy(); cfg.update(d); return cfg
        except: pass
    return DEFAULT_CFG.copy()

def save_cfg(cfg):
    CFG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

def hotkey_display(parts):
    """["ctrl","shift","s"] → "Ctrl+Shift+S" """
    return "+".join(p.upper() if len(p)==1 else p.capitalize() for p in parts)


# ─────────────────────────────────────────────
#  DB
# ─────────────────────────────────────────────
def init_db():
    c = sqlite3.connect(DB_PATH)
    c.executescript("""
        CREATE TABLE IF NOT EXISTS sensitivity_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            sensitivity REAL    NOT NULL,
            dpi         INTEGER,
            edpi        REAL,
            note        TEXT,
            screenshot  TEXT,
            method      TEXT DEFAULT '手動',
            changed_at  TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS match_stats (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            sensitivity_id INTEGER NOT NULL,
            match_id       TEXT,
            match_date     TEXT,
            agent          TEXT,
            map            TEXT,
            kills          INTEGER,
            deaths         INTEGER,
            assists        INTEGER,
            acs            REAL,
            hs_percent     REAL,
            damage         REAL,
            won            INTEGER,
            tracker_url    TEXT,
            note           TEXT,
            added_at       TEXT NOT NULL,
            FOREIGN KEY (sensitivity_id) REFERENCES sensitivity_log(id)
        );
        CREATE TABLE IF NOT EXISTS round_stats (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            match_stats_id INTEGER NOT NULL,
            round_num      INTEGER,
            kills          INTEGER,
            score          INTEGER,
            damage         INTEGER,
            hs_hits        INTEGER,
            total_hits     INTEGER,
            FOREIGN KEY (match_stats_id) REFERENCES match_stats(id)
        );
    """)
    c.commit(); c.close()

def migrate_db():
    """既存DBに新しい列・テーブルを追加（データは消えない）"""
    c = sqlite3.connect(DB_PATH)
    # damage列を追加（なければ）
    try:
        c.execute("ALTER TABLE match_stats ADD COLUMN damage REAL")
    except: pass
    # match_id列を追加（なければ）
    try:
        c.execute("ALTER TABLE match_stats ADD COLUMN match_id TEXT")
    except: pass
    # round_statsテーブルを追加（なければ）
    c.executescript("""
        CREATE TABLE IF NOT EXISTS round_stats (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            match_stats_id INTEGER NOT NULL,
            round_num      INTEGER,
            kills          INTEGER,
            score          INTEGER,
            damage         INTEGER,
            hs_hits        INTEGER,
            total_hits     INTEGER,
            FOREIGN KEY (match_stats_id) REFERENCES match_stats(id)
        );
    """)
    try:
        c.execute("ALTER TABLE round_stats ADD COLUMN score INTEGER")
    except: pass
    c.commit(); c.close()

def db():
    return sqlite3.connect(DB_PATH)


# ─────────────────────────────────────────────
#  OCR エンジン
# ─────────────────────────────────────────────
class SensOCR:
    def __init__(self, cfg):
        self.cfg = cfg
        if HAS_OCR and cfg.get("tesseract_path"):
            try: pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_path"]
            except: pass

    def capture(self):
        if not HAS_PIL: raise RuntimeError("Pillow が必要です: pip install pillow")
        return ImageGrab.grab()

    def extract(self, img):
        """PIL Image → (感度値 or None, OCRテキスト)"""
        if not HAS_OCR: raise RuntimeError("pytesseract が必要です: pip install pytesseract")
        texts = []
        for scale in [2, 3]:
            g = img.convert("L")
            g = g.resize((g.width*scale, g.height*scale), Image.LANCZOS)
            g = ImageEnhance.Contrast(g).enhance(3.0)
            g = g.point(lambda x: 0 if x < 130 else 255)
            texts.append(pytesseract.image_to_string(
                g, config="--psm 6 -c tessedit_char_whitelist=0123456789.SensitivityMousemouse "))
        texts.append(pytesseract.image_to_string(img, config="--psm 6"))
        full = "\n".join(texts)
        return self._parse(full), full

    def _parse(self, text):
        # "Mouse Sensitivity" 周辺を優先探索
        for pat in [
            r'[Mm]ouse\s*[Ss]ensitivity[^\d]*([\d]+\.[\d]+)',
            r'[Ss]ensitivity[^\d]*([\d]+\.[\d]+)',
        ]:
            m = re.search(pat, text)
            if m:
                v = float(m.group(1))
                if 0.01 <= v <= 10.0: return v
        # 小数を全て拾って範囲フィルタ
        cands = [float(x) for x in re.findall(r'\b(\d+\.\d{1,3})\b', text)
                 if 0.01 <= float(x) <= 10.0]
        if not cands: return None
        preferred = [v for v in cands if 0.05 <= v <= 3.0]
        pool = preferred or cands
        cnt = Counter(pool)
        return cnt.most_common(1)[0][0]


# ─────────────────────────────────────────────
#  グローバルホットキー（pynput）
# ─────────────────────────────────────────────
class HotkeyManager:
    # 修飾キーの正規化マップ
    _MOD_NORMALIZE = {}  # populated after pynput import

    def __init__(self, callback):
        self.callback = callback
        self._listener = None
        self._held     = set()
        self._target   = set()
        if HAS_PYNPUT:
            K = _pynput_kb.Key
            self._MOD_NORMALIZE = {
                K.ctrl_l:  "ctrl", K.ctrl_r:  "ctrl",
                K.shift_l: "shift", K.shift_r: "shift",
                K.alt_l:   "alt",  K.alt_r:   "alt",  K.alt_gr: "alt",
                K.cmd_l:   "win",  K.cmd_r:   "win",
            }
            # ファンクションキー名
            self._FKEY = {getattr(K, f"f{i}", None): f"f{i}" for i in range(1, 13)}

    def start(self, parts):
        """parts: ["ctrl","shift","s"] のリスト"""
        self.stop()
        if not HAS_PYNPUT: return False
        self._target = set(p.lower() for p in parts)
        self._held   = set()
        self._listener = _pynput_kb.Listener(
            on_press=self._press, on_release=self._release)
        self._listener.daemon = True
        self._listener.start()
        return True

    def stop(self):
        if self._listener:
            try: self._listener.stop()
            except: pass
            self._listener = None

    def _key_name(self, key):
        # 修飾キー
        n = self._MOD_NORMALIZE.get(key)
        if n: return n
        # ファンクションキー
        n = self._FKEY.get(key)
        if n: return n
        # 通常文字
        try: return key.char.lower()
        except: pass
        # その他（key.name）
        try: return key.name.lower()
        except: return str(key).lower()

    def _press(self, key):
        self._held.add(self._key_name(key))
        if self._target and self._target.issubset(self._held):
            self._held.clear()
            self.callback()

    def _release(self, key):
        self._held.discard(self._key_name(key))


# ─────────────────────────────────────────────
#  カラーテーマ
# ─────────────────────────────────────────────
BG="# 0f1923";BG="#0f1923";BG2="#1a2535";ACCENT="#ff4655";ACCENT2="#00d4ff"
TEXT="#ecf0f1";DIM="#7f8c8d";GREEN="#2ecc71";YELLOW="#f1c40f";CARD="#16202e";ORANGE="#e67e22"

FT = ("NotoSansJP")
FH = ("NotoSansJP", 13, "bold")
FB = ("NotoSansJP", 10)
FS = ("NotoSansJP",  9)

# ─────────────────────────────────────────────
#  初回セットアップウィザード
# ─────────────────────────────────────────────
class SetupWizard(tk.Toplevel):
    """
    4ステップの初回セットアップ画面
    Step1: Tesseractインストール確認
    Step2: ホットキー設定
    Step3: HenrikDev APIキー入力
    Step4: VALORANTプレイヤーID入力
    """
    STEPS = ["Tesseract OCR", "ホットキー", "APIキー", "プレイヤーID"]

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("初回セットアップ")
        self.configure(bg=BG)
        self.geometry("620x520")
        self.resizable(False, False)
        self.grab_set()
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._skip)   # ×で閉じたらスキップ扱い
        self._step = 0
        self._build()
        self._show_step(0)

    def _build(self):
        # ── タイトル ──
        tk.Label(self, text="🎯  初回セットアップ", bg=BG, fg=ACCENT2,
                 font=("NotoSansJP", 16, "bold")).pack(pady=(18, 4))
        # ── ステップインジケーター ──
        self._step_frame = tk.Frame(self, bg=BG)
        self._step_frame.pack(pady=(0, 10))
        self._step_labels = []
        for i, name in enumerate(self.STEPS):
            f = tk.Frame(self._step_frame, bg=BG)
            f.pack(side="left", padx=6)
            num = tk.Label(f, text=str(i+1), bg=DIM, fg="white",
                           font=("NotoSansJP", 10, "bold"), width=2)
            num.pack()
            lbl = tk.Label(f, text=name, bg=BG, fg=DIM, font=("NotoSansJP", 8))
            lbl.pack()
            self._step_labels.append((num, lbl))

        # 区切り線
        tk.Frame(self, height=1, bg=ACCENT).pack(fill="x", padx=20)

        # ── コンテンツエリア（ステップごとに中身を差し替え）──
        self._content = tk.Frame(self, bg=BG)
        self._content.pack(fill="both", expand=True, padx=30, pady=16)

        # ── 下部ボタン ──
        btn_f = tk.Frame(self, bg=BG)
        btn_f.pack(pady=(0, 18))
        self._btn_skip = ttk.Button(btn_f, text="スキップ",
                                    style="Sub.TButton", command=self._skip)
        self._btn_skip.pack(side="left", padx=8)
        self._btn_next = ttk.Button(btn_f, text="次へ ▶",
                                    command=self._next)
        self._btn_next.pack(side="left", padx=8)

    def _show_step(self, step):
        # インジケーター更新
        for i, (num, lbl) in enumerate(self._step_labels):
            if i < step:
                num.config(bg=GREEN);  lbl.config(fg=GREEN)    # 完了
            elif i == step:
                num.config(bg=ACCENT); lbl.config(fg=TEXT)     # 現在
            else:
                num.config(bg=DIM);    lbl.config(fg=DIM)      # 未到達

        # コンテンツをクリア
        for w in self._content.winfo_children():
            w.destroy()

        # ステップ別に中身を描画
        [self._step_tesseract,
         self._step_hotkey,
         self._step_apikey,
         self._step_playerid][step]()

        # 最終ステップはボタンラベル変更
        self._btn_next.config(
            text="完了 ✔" if step == len(self.STEPS) - 1 else "次へ ▶")

    # ── Step 1: Tesseract ───────────────────────
    def _step_tesseract(self):
        c = self._content
        tk.Label(c, text="Step 1 ─ Tesseract OCR のインストール",
                 bg=BG, fg=TEXT, font=("NotoSansJP", 12, "bold")).pack(anchor="w", pady=(0,10))
        tk.Label(c,
            text="ホットキーでスクリーンショットから感度を自動取得するには\n"
                 "Tesseract OCR が必要です。未インストールの場合は下のリンクから。",
            bg=BG, fg=TEXT, font=("NotoSansJP", 10), justify="left").pack(anchor="w")

        # インストール状況
        status = "✔ インストール済み" if HAS_OCR else "✘ 未インストール"
        color  = GREEN if HAS_OCR else ACCENT
        tk.Label(c, text=f"現在の状態: {status}", bg=BG, fg=color,
                 font=("NotoSansJP", 11, "bold")).pack(anchor="w", pady=(12, 4))

        # ダウンロードリンク
        link = tk.Label(c,
            text="📥  https://github.com/tesseract-ocr/tesseract/releases",
            bg=BG, fg=ACCENT2, font=("NotoSansJP", 9), cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda _: webbrowser.open(
            "https://github.com/tesseract-ocr/tesseract/releases"))

        tk.Label(c, text="インストール後、以下にexeのパスを入力してください:",
                 bg=BG, fg=DIM, font=("NotoSansJP", 9)).pack(anchor="w", pady=(14, 3))

        path_f = tk.Frame(c, bg=BG); path_f.pack(fill="x")
        self._tess_entry = ttk.Entry(path_f, width=46)
        self._tess_entry.insert(0, self.parent.cfg.get("tesseract_path", ""))
        self._tess_entry.pack(side="left")
        ttk.Button(path_f, text="保存", style="Sub.TButton",
                   command=self._save_tess).pack(side="left", padx=6)

        tk.Label(c,
            text="※ デフォルトのインストール先:\n"
                 r"   C:\Program Files\Tesseract-OCR\tesseract.exe",
            bg=BG, fg=DIM, font=("NotoSansJP", 8), justify="left").pack(anchor="w", pady=(6,0))

    def _save_tess(self):
        p = self._tess_entry.get().strip()
        self.parent.cfg["tesseract_path"] = p
        save_cfg(self.parent.cfg)
        if HAS_OCR and p:
            try: pytesseract.pytesseract.tesseract_cmd = p
            except: pass
        messagebox.showinfo("保存", "Tesseractパスを保存しました", parent=self)

    # ── Step 2: ホットキー ──────────────────────
    def _step_hotkey(self):
        c = self._content
        tk.Label(c, text="Step 2 ─ ホットキーの設定",
                 bg=BG, fg=TEXT, font=("NotoSansJP", 12, "bold")).pack(anchor="w", pady=(0,10))
        tk.Label(c,
            text="VALORANTの設定画面を開いた状態でこのキーを押すと\n"
                 "感度をスクショ→OCRで自動取得します。",
            bg=BG, fg=TEXT, font=("NotoSansJP", 10), justify="left").pack(anchor="w")

        current = hotkey_display(self.parent.cfg.get("hotkey_parts", ["ctrl","shift","s"]))
        tk.Label(c, text=f"現在のホットキー: {current}",
                 bg=BG, fg=ACCENT, font=("NotoSansJP", 12, "bold")).pack(anchor="w", pady=(12,4))

        tk.Label(c, text="変更する場合は設定タブの「ホットキー設定」から変更できます。",
                 bg=BG, fg=DIM, font=("NotoSansJP", 9)).pack(anchor="w")
        tk.Label(c, text=f"このまま  {current}  を使う場合は「次へ」を押してください。",
                 bg=BG, fg=GREEN, font=("NotoSansJP", 10)).pack(anchor="w", pady=(8,0))

    # ── Step 3: APIキー ─────────────────────────
    def _step_apikey(self):
        c = self._content
        tk.Label(c, text="Step 3 ─ HenrikDev APIキーの取得・入力",
                 bg=BG, fg=TEXT, font=("NotoSansJP", 12, "bold")).pack(anchor="w", pady=(0,10))
        tk.Label(c,
            text="試合スタッツの自動取得に使います。無料で取得できます。",
            bg=BG, fg=TEXT, font=("NotoSansJP", 10)).pack(anchor="w")

        # リンク
        link = tk.Label(c, text="📥  https://docs.henrikdev.xyz  （無料登録→APIキー発行）",
                        bg=BG, fg=ACCENT2, font=("NotoSansJP", 9), cursor="hand2")
        link.pack(anchor="w", pady=(4, 12))
        link.bind("<Button-1>", lambda _: webbrowser.open("https://docs.henrikdev.xyz"))

        tk.Label(c, text="取得したAPIキー:", bg=BG, fg=DIM,
                 font=("NotoSansJP", 9)).pack(anchor="w")
        self._api_entry = ttk.Entry(c, width=50)
        self._api_entry.insert(0, self.parent.cfg.get("henrik_api_key", ""))
        self._api_entry.pack(anchor="w", pady=(3, 6))

        ttk.Button(c, text="保存", style="Sub.TButton",
                   command=self._save_api).pack(anchor="w")
        tk.Label(c, text="※ スキップしても後で設定タブから入力できます。",
                 bg=BG, fg=DIM, font=("NotoSansJP", 8)).pack(anchor="w", pady=(8,0))

    def _save_api(self):
        key = self._api_entry.get().strip()
        self.parent.cfg["henrik_api_key"] = key
        save_cfg(self.parent.cfg)
        messagebox.showinfo("保存", "APIキーを保存しました", parent=self)

    # ── Step 4: プレイヤーID ────────────────────
    def _step_playerid(self):
        c = self._content
        tk.Label(c, text="Step 4 ─ VALORANTプレイヤーIDの入力",
                 bg=BG, fg=TEXT, font=("NotoSansJP", 12, "bold")).pack(anchor="w", pady=(0,10))
        tk.Label(c,
            text="tracker.ggのプロフィールURLからスタッツを自動取得する際に使います。",
            bg=BG, fg=TEXT, font=("NotoSansJP", 10)).pack(anchor="w")

        link = tk.Label(c, text="📊  https://tracker.gg/valorant  （自分のページで名前・タグを確認）",
                        bg=BG, fg=ACCENT2, font=("NotoSansJP", 9), cursor="hand2")
        link.pack(anchor="w", pady=(4, 12))
        link.bind("<Button-1>", lambda _: webbrowser.open("https://tracker.gg/valorant"))

        f1 = tk.Frame(c, bg=BG); f1.pack(anchor="w", pady=3)
        tk.Label(f1, text="プレイヤー名:", bg=BG, fg=TEXT,
                 font=("NotoSansJP", 10), width=14, anchor="w").pack(side="left")
        self._name_entry = ttk.Entry(f1, width=24)
        self._name_entry.insert(0, self.parent.cfg.get("player_name", ""))
        self._name_entry.pack(side="left")

        f2 = tk.Frame(c, bg=BG); f2.pack(anchor="w", pady=3)
        tk.Label(f2, text="タグ (#JP1など):", bg=BG, fg=TEXT,
                 font=("NotoSansJP", 10), width=14, anchor="w").pack(side="left")
        self._tag_entry = ttk.Entry(f2, width=14)
        self._tag_entry.insert(0, self.parent.cfg.get("player_tag", ""))
        self._tag_entry.pack(side="left")

        ttk.Button(c, text="保存", style="Sub.TButton",
                   command=self._save_player).pack(anchor="w", pady=(10,0))
        tk.Label(c, text="※ スキップしても後で設定タブから入力できます。",
                 bg=BG, fg=DIM, font=("NotoSansJP", 8)).pack(anchor="w", pady=(6,0))

    def _save_player(self):
        self.parent.cfg["player_name"] = self._name_entry.get().strip()
        self.parent.cfg["player_tag"]  = self._tag_entry.get().strip().lstrip("#")
        save_cfg(self.parent.cfg)
        messagebox.showinfo("保存", "プレイヤーIDを保存しました", parent=self)

    # ── ナビゲーション ──────────────────────────
    def _next(self):
        if self._step < len(self.STEPS) - 1:
            self._step += 1
            self._show_step(self._step)
        else:
            self._finish()

    def _finish(self):
        self.parent.cfg["setup_done"] = True
        save_cfg(self.parent.cfg)
        self.parent.refresh_all()
        self.destroy()

    def _skip(self):
        """×ボタンまたはスキップ → setup_doneをTrueにして二度と出さない"""
        self.parent.cfg["setup_done"] = True
        save_cfg(self.parent.cfg)
        self.destroy()


# ─────────────────────────────────────────────
#  メインアプリ
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        init_db()
        migrate_db()
        self.cfg    = load_cfg()
        self.ocr    = SensOCR(self.cfg)
        self.hotkey = HotkeyManager(self._on_hotkey)
        self._overlay_win = None

        self.title("VALORANT Sensitivity Tracker")
        self.geometry("1150x760")
        self.minsize(960, 640)
        self.configure(bg=BG)
        self.protocol("WM_DELETE_WINDOW", self._quit)
        self._style()
        self._ui()
        self.refresh_all()
        self._hotkey_start()

         # ── 初回起動時のみウィザードを表示 ──  ← これを追加
        if not self.cfg.get("setup_done", False):
            self.after(300, lambda: SetupWizard(self))

    def _quit(self):
        self.hotkey.stop()
        self.destroy()

    # ── ホットキー起動 ──────────────────────
    def _hotkey_start(self):
        ok = self.hotkey.start(self.cfg["hotkey_parts"])
        self._update_status_bar(ok)

    def _update_status_bar(self, active=True):
        if not hasattr(self, "_status_lbl"): return
        disp = hotkey_display(self.cfg["hotkey_parts"])
        if not HAS_PYNPUT:
            txt = "⚠  pynput未インストール (pip install pynput)  ホットキー無効"
            col = ORANGE
        elif active:
            txt = f"⌨  ホットキー監視中:  {disp}   |   VALORANTの設定画面でこのキーを押すと感度を自動取得します"
            col = GREEN
        else:
            txt = "✘  ホットキー無効"; col = ACCENT
        self._status_lbl.config(text=txt, fg=col)

    # ── ホットキー発火 ──────────────────────
    def _on_hotkey(self):
        """別スレッドから呼ばれる → after でメインスレッドへ"""
        self.after(0, self._do_capture)

    def _do_capture(self):
        if not HAS_PIL:
            messagebox.showerror("エラー", "Pillow が必要です\npip install pillow")
            return
        self._show_overlay("📸  スクリーンショット撮影中…")
        threading.Thread(target=self._capture_thread, daemon=True).start()

    def _capture_thread(self):
        try:
            img  = self.ocr.capture()
            if HAS_OCR:
                val, raw = self.ocr.extract(img)
            else:
                val, raw = None, "(OCR無効)"
            # PIL Image は pickle できないので bytes で渡す
            buf = io.BytesIO(); img.save(buf, format="PNG"); img_bytes = buf.getvalue()
            self.after(0, lambda: self._capture_done(val, img_bytes, None, raw))
        except Exception as e:
            self.after(0, lambda: (self._hide_overlay(),
                messagebox.showerror("エラー", f"スクショ失敗:\n{e}")))

    def _capture_done(self, val, img_bytes, path, raw):
        self._hide_overlay()
        img = Image.open(io.BytesIO(img_bytes))
        SensConfirmDialog(self, detected=val, img=img, ss_path=path, raw=raw)

    # ── オーバーレイ ────────────────────────
    def _show_overlay(self, msg):
        self._hide_overlay()
        ov = tk.Toplevel(self)
        ov.overrideredirect(True)
        ov.attributes("-topmost", True)
        ov.attributes("-alpha", 0.88)
        W, H = 380, 72
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ov.geometry(f"{W}x{H}+{(sw-W)//2}+{sh-H-100}")
        ov.configure(bg=CARD)
        tk.Label(ov, text=msg, bg=CARD, fg=ACCENT2,
                 font=("NotoSansJP",15,"bold")).pack(expand=True)
        self._overlay_win = ov
        self.after(6000, self._hide_overlay)

    def _hide_overlay(self):
        if self._overlay_win:
            try: self._overlay_win.destroy()
            except: pass
            self._overlay_win = None

    # ─────────────────────────────────────
    #  スタイル
    # ─────────────────────────────────────
    def _style(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure(".",         background=BG,   foreground=TEXT, font=FB)
        s.configure("TFrame",    background=BG)
        s.configure("TLabel",    background=BG,   foreground=TEXT)
        s.configure("Dim.TLabel",background=BG,   foreground=DIM)
        s.configure("TButton",   background=ACCENT,foreground="white",
                    font=("NotoSansJP",10,"bold"), borderwidth=0,
                    focusthickness=0, padding=(12,6))
        s.map("TButton", background=[("active","#cc3340"),("pressed","#aa2030")])
        s.configure("Sub.TButton", background=BG2, foreground=TEXT,
                    font=FB, padding=(10,5), borderwidth=0)
        s.map("Sub.TButton", background=[("active","#253545")])
        s.configure("Treeview",  background=BG2, foreground=TEXT,
                    fieldbackground=BG2, rowheight=26, borderwidth=0, font=FB)
        s.configure("Treeview.Heading", background=BG, foreground=ACCENT2,
                    font=("NotoSansJP",10,"bold"), borderwidth=0)
        s.map("Treeview", background=[("selected","#253f5c")],
              foreground=[("selected",TEXT)])
        s.configure("TEntry", fieldbackground="white", foreground="#111111",
                    insertcolor="#111111", borderwidth=0, padding=5)
        s.configure("TCombobox", fieldbackground="white", foreground="#111111",
                    background="white", selectbackground="#ddeeff",
                    selectforeground="#111111", arrowcolor=TEXT, borderwidth=0)
        s.map("TCombobox",
              fieldbackground=[("readonly","white")],
              foreground=[("readonly","#111111")],
              selectbackground=[("readonly","#ddeeff")])
        s.configure("TNotebook",background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", background=BG2, foreground=DIM,
                    padding=(16,8), borderwidth=0)
        s.map("TNotebook.Tab", background=[("selected",BG)],
              foreground=[("selected",ACCENT)])

    # ─────────────────────────────────────
    #  UI 構築
    # ─────────────────────────────────────
    def _ui(self):
        # ヘッダー
        hdr = ttk.Frame(self); hdr.pack(fill="x", padx=20, pady=(14,0))
        tk.Label(hdr, text="🎯  VALORANT Sensitivity Tracker",
                 bg=BG, fg=TEXT, font=FT).pack(side="left")
        sep = tk.Frame(self, height=2, bg=ACCENT)
        sep.pack(fill="x", padx=20, pady=(6,0))

        # ステータスバー（ホットキー表示）
        self._status_lbl = tk.Label(self, text="", bg="#0a1218", fg=GREEN,
                                    font=FS, anchor="w")
        self._status_lbl.pack(fill="x", padx=0, pady=0, ipady=5)

        # タブ
        self.nb = ttk.Notebook(self); self.nb.pack(fill="both", expand=True, padx=10, pady=6)
        tabs = [("  ホーム  ","tab_home"),("  感度管理  ","tab_sens"),
                ("  試合記録  ","tab_matches"),("  分析  ","tab_analysis"),
                ("  設定  ","tab_settings")]
        for label, attr in tabs:
            f = ttk.Frame(self.nb); self.nb.add(f, text=label)
            setattr(self, attr, f)

        self._build_home()
        self._build_sens()
        self._build_matches()
        self._build_analysis()
        self._build_settings()

    # ───────────────────────────────────
    #  ホームタブ
    # ───────────────────────────────────
    def _build_home(self):
        f = self.tab_home
        top = ttk.Frame(f); top.pack(fill="x", padx=20, pady=16)

        # 現在の感度カード
        cl = tk.Frame(top, bg=CARD); cl.pack(side="left",fill="both",expand=True,padx=(0,8))
        tk.Label(cl, text="現在の感度", bg=CARD, fg=DIM, font=FS).pack(anchor="w",padx=16,pady=(12,0))
        self._lbl_sens  = tk.Label(cl, text="--", bg=CARD, fg=ACCENT, font=("NotoSansJP",52,"bold"))
        self._lbl_sens.pack(anchor="w", padx=16)
        self._lbl_edpi  = tk.Label(cl, text="eDPI: --", bg=CARD, fg=DIM, font=FB)
        self._lbl_edpi.pack(anchor="w", padx=16)
        self._lbl_since = tk.Label(cl, text="変更日: --", bg=CARD, fg=DIM, font=FS)
        self._lbl_since.pack(anchor="w", padx=16, pady=(2,12))

        # ホットキー案内カード
        cr = tk.Frame(top, bg=CARD); cr.pack(side="left",fill="both",expand=True)
        tk.Label(cr, text="感度の自動記録", bg=CARD, fg=DIM, font=FS).pack(anchor="w",padx=16,pady=(12,4))
        tk.Label(cr, text="VALORANTの設定画面を開いた状態でホットキーを押すと\n感度をスクショ→OCRで自動取得して記録します。",
                 bg=CARD, fg=TEXT, font=FB, justify="left").pack(anchor="w", padx=16)
        self._lbl_hk_home = tk.Label(cr, text="", bg="#0e2218", fg=GREEN,
                                     font=("NotoSansJP",13,"bold"))
        self._lbl_hk_home.pack(anchor="w", padx=16, pady=(8,4))
        tk.Frame(cr, bg=BG2, height=1).pack(fill="x", padx=16, pady=6)
        ttk.Button(cr, text="＋ 感度を手動登録",  command=self._open_add_sens).pack(fill="x",padx=16,pady=2)
        ttk.Button(cr, text="＋ 試合を記録",       command=self._open_add_match).pack(fill="x",padx=16,pady=2)
        ttk.Button(cr, text="📊 分析を見る", style="Sub.TButton",
                   command=lambda: self.nb.select(self.tab_analysis)).pack(fill="x",padx=16,pady=(2,12))

        # 最近の試合テーブル
        tk.Label(f, text="最近の試合", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(4,4))
        cols = ("日時","エージェント","マップ","K/D/A","ACS","HS%","勝敗","感度")
        self._tree_recent = ttk.Treeview(f, columns=cols, show="headings", height=8)
        for c,w in zip(cols,[135,92,92,82,62,62,56,66]):
            self._tree_recent.heading(c,text=c); self._tree_recent.column(c,width=w,anchor="center")
        self._tree_recent.pack(fill="x", padx=20)

    # ───────────────────────────────────
    #  感度管理タブ
    # ───────────────────────────────────
    def _build_sens(self):
        f = self.tab_sens
        # バナー
        bn = tk.Frame(f, bg="#0d2a1a"); bn.pack(fill="x",padx=20,pady=(14,4))
        tk.Label(bn, text="💡  VALORANTの設定画面でホットキーを押すと自動登録されます。手動登録は下のフォームを使ってください。",
                 bg="#0d2a1a", fg=GREEN, font=FS).pack(padx=12, pady=8, anchor="w")

        # 手動フォーム
        fm = ttk.LabelFrame(f, text=" 手動で感度を記録 ", padding=14)
        fm.pack(fill="x", padx=20, pady=4)
        row = ttk.Frame(fm); row.pack(fill="x")
        self._e_sv = self._field(row, "感度 (例: 0.35)", 12)
        self._e_sd = self._field(row, "DPI (例: 800)", 10)
        self._e_se = self._field(row, "eDPI (自動)", 10, ro=True)
        self._e_sn = self._field(row, "メモ", 28)
        self._e_sv.bind("<KeyRelease>", self._calc_edpi)
        self._e_sd.bind("<KeyRelease>", self._calc_edpi)
        ttk.Button(fm, text="✔ 保存", command=self._save_sens_manual).pack(anchor="e",pady=(8,0))

        # 履歴テーブル
        tk.Label(f, text="感度変更履歴", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(10,4))
        cols = ("ID","感度","DPI","eDPI","試合数","取得方法","変更日時","メモ")
        self._tree_sens = ttk.Treeview(f, columns=cols, show="headings", height=11)
        for c,w in zip(cols,[40,75,65,72,65,70,158,180]):
            self._tree_sens.heading(c,text=c); self._tree_sens.column(c,width=w,anchor="center")
        sb = ttk.Scrollbar(f, orient="vertical", command=self._tree_sens.yview)
        self._tree_sens.configure(yscrollcommand=sb.set)
        self._tree_sens.pack(fill="both", expand=True, padx=20)
        self._tree_sens.bind("<Double-1>", self._sens_dbl)
        br = ttk.Frame(f); br.pack(fill="x",padx=20,pady=4)
        ttk.Button(br, text="🗑 選択した感度を削除", style="Sub.TButton",
                   command=self._del_sens).pack(side="right")
        ttk.Button(br, text="✔ 選択した感度を現在の感度に設定", style="Sub.TButton",
           command=self._set_as_current).pack(side="right", padx=(0, 8))
        ttk.Label(br, text="※ ダブルクリックでその感度の試合一覧へ",
                  style="Dim.TLabel", font=FS).pack(side="left")

    # ───────────────────────────────────
    #  試合記録タブ
    # ───────────────────────────────────
    def _build_matches(self):
        f = self.tab_matches

        # ── tracker.gg 一括取得バナー ──────────────
        bulk_card = tk.Frame(f, bg="#0d1f2d"); bulk_card.pack(fill="x", padx=20, pady=(14,4))
        bulk_inner = tk.Frame(bulk_card, bg="#0d1f2d"); bulk_inner.pack(fill="x", padx=14, pady=10)
        tk.Label(bulk_inner,
                 text="🔄  Henrikdev API から試合を一括取得",
                 bg="#0d1f2d", fg=ACCENT2, font=FH).pack(anchor="w")
        tk.Label(bulk_inner,
                 text="感度を選択して「取得」を押すと、その感度を使っていた期間のコンペティティブ試合（ラウンド別KD/DMG/HS含む）を自動取得します。",
                 bg="#0d1f2d", fg=TEXT, font=FS, wraplength=900, justify="left").pack(anchor="w", pady=(2,6))
        tk.Label(bulk_inner,
                 text="※ プレイヤー名・タグ・APIキーは設定タブで設定してください",
                 bg="#0d1f2d", fg=DIM, font=FS).pack(anchor="w", pady=(0,6))
        ttk.Button(bulk_inner, text="🚀 この感度期間の試合を取得",
                   command=self._bulk_fetch).pack(anchor="w")
        self._lbl_bulk_status = tk.Label(bulk_inner, text="", bg="#0d1f2d", fg=DIM, font=FS)
        self._lbl_bulk_status.pack(anchor="w", pady=(4,0))

        # ── 感度選択・手動追加 ──────────────────────
        sr = ttk.Frame(f); sr.pack(fill="x",padx=20,pady=(6,4))
        tk.Label(sr, text="対象感度:", bg=BG, fg=TEXT, font=FB).pack(side="left")
        self._combo_msens = ttk.Combobox(sr, width=30, state="readonly")
        self._combo_msens.pack(side="left", padx=8)
        ttk.Button(sr, text="＋ 試合を手動追加", command=self._open_add_match).pack(side="left",padx=4)

        self._combo_msens.bind("<<ComboboxSelected>>", self._filter_matches)

        cols = ("日時","エージェント","マップ","K","D","A","ACS","DMG","HS%","勝敗","メモ")
        self._tree_matches = ttk.Treeview(f, columns=cols, show="headings", height=12)
        for c,w in zip(cols,[132,92,92,42,42,42,62,68,62,56,160]):
            self._tree_matches.heading(c,text=c); self._tree_matches.column(c,width=w,anchor="center")
        sb = ttk.Scrollbar(f, orient="vertical", command=self._tree_matches.yview)
        self._tree_matches.configure(yscrollcommand=sb.set)
        self._tree_matches.pack(fill="both", expand=True, padx=20)
        br = ttk.Frame(f); br.pack(fill="x",padx=20,pady=4)
        ttk.Button(br, text="🗑 選択した試合を削除", style="Sub.TButton",
                   command=self._del_match).pack(side="right")

    def _bulk_fetch(self):
        sel = self._combo_msens.get()
        if not sel:
            messagebox.showerror("エラー","対象の感度を選択してください"); return
        sid = int(re.match(r"\[(\d+)\]", sel).group(1))

        name    = self.cfg.get("player_name","").strip()
        tag     = self.cfg.get("player_tag","").strip()
        api_key = self.cfg.get("henrik_api_key","").strip()
        region  = self.cfg.get("region","ap").strip()

        if not name or not tag:
            messagebox.showerror("エラー","設定タブでプレイヤー名とタグを設定してください"); return
        if not api_key:
            messagebox.showerror("エラー","設定タブでHenrikdev APIキーを設定してください"); return

        # この感度の期間（開始〜次の感度変更）を取得
        c = db()
        cur_row = c.execute(
            "SELECT changed_at FROM sensitivity_log WHERE id=?", (sid,)).fetchone()
        next_row = c.execute(
            "SELECT changed_at FROM sensitivity_log WHERE id>? ORDER BY id ASC LIMIT 1", (sid,)).fetchone()
        c.close()

        if not cur_row:
            messagebox.showerror("エラー","感度データが見つかりません"); return

        since_str = cur_row[0]
        until_str = next_row[0] if next_row else None

        self._lbl_bulk_status.config(
            text=f"取得中… 期間: {since_str} 〜 {until_str or '現在'}", fg=YELLOW)
        my_puuid = self.cfg.get("player_puuid","")
        threading.Thread(
            target=self._bulk_fetch_thread,
            args=(name, tag, api_key, region, sid, since_str, until_str, my_puuid),
            daemon=True).start()

    def _bulk_fetch_thread(self, name, tag, api_key, region, sid, since_str, until_str, my_puuid=""):
        import urllib.parse
        try:
            fmt = "%Y-%m-%d %H:%M"
            since_dt = datetime.strptime(since_str, fmt)
            until_dt = datetime.strptime(until_str, fmt) if until_str else None

            headers = {"Authorization": api_key}
            all_matches = []

            enc_name = urllib.parse.quote(name, safe='')
            enc_tag  = urllib.parse.quote(tag,  safe='')

            # v4 API: /valorant/v4/matches/{region}/pc/{name}/{tag}
            # ページネーションは ?start=INDEX
            start_idx = 0
            while True:
                api = (f"https://api.henrikdev.xyz/valorant/v4/matches/{region}/pc/"
                       f"{enc_name}/{enc_tag}?mode=competitive&size=20&start={start_idx}")

                self.after(0, lambda n=len(all_matches): self._lbl_bulk_status.config(
                    text=f"取得中… ({n}試合取得済み)", fg=YELLOW))

                req = urllib.request.Request(api, headers=headers)
                with urllib.request.urlopen(req, timeout=20) as r:
                    data = json.loads(r.read().decode("utf-8","replace"))

                status = data.get("status", 0)
                if status == 429:
                    raise ValueError("APIレート制限に達しました。少し待ってから再試行してください")
                if status not in (200,) and data.get("errors"):
                    raise ValueError(f"APIエラー: {data.get('errors')}")

                matches = data.get("data", [])
                if not matches:
                    break

                stop = False
                for match in matches:
                    meta = match.get("metadata", {})
                    # v4のstarted_atはISO文字列 "2025-03-01T08:45:18.755Z"
                    ts = meta.get("started_at","")
                    try:
                        from datetime import timedelta
                        ts_clean = re.sub(r"\.\d+","", str(ts)).replace("Z","")
                        match_dt_utc = datetime.fromisoformat(ts_clean)
                        match_dt = match_dt_utc + timedelta(hours=9)  # UTC→JST
                    except Exception as te:
                        continue

                    if match_dt < since_dt:
                        stop = True; break
                    if until_dt and match_dt >= until_dt:
                        continue

                    all_matches.append((match, match_dt))

                if stop or len(matches) < 20:
                    break
                start_idx += 20

            if not all_matches:
                self.after(0, lambda: self._lbl_bulk_status.config(
                    text=f"該当する試合が見つかりませんでした（期間: {since_str} 〜 {until_str or '現在'}）",
                    fg=YELLOW))
                return

            self.after(0, lambda: BulkImportDialog(
                self, sid=sid, matches=all_matches,
                since_str=since_str, until_str=until_str,
                name=name, api_key=api_key, region=region,
                my_puuid=my_puuid))

        except Exception as e:
            self.after(0, lambda: self._lbl_bulk_status.config(
                text=f"取得失敗: {e}", fg=ACCENT))



  # ───────────────────────────────────
    #  分析タブ
    # ───────────────────────────────────
    def _build_analysis(self):
        f = self.tab_analysis

        # 左右2列のメインフレーム
        columns = tk.Frame(f, bg=BG); columns.pack(fill="both", expand=True, padx=20, pady=(14,4))

        # ── 左列：自分の感度別分析 ──────────────
        left = tk.Frame(columns, bg=BG); left.pack(side="left", fill="both", expand=True, padx=(0,8))
        tk.Label(left, text="感度別スタッツ", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w", pady=(0,4))
        wrap = tk.Frame(left, bg=BG); wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        self._analysis_frame = ttk.Frame(canvas)
        self._analysis_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self._analysis_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ttk.Button(left, text="🔄 更新", command=self.refresh_analysis).pack(anchor="se", pady=4)

        # ── 右列：プロ参照データ ──────────────
        right = tk.Frame(columns, bg=BG, width=420); right.pack(side="left", fill="both", padx=(8,0))
        right.pack_propagate(False)
        tk.Label(right, text="プロ参照データ", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w", pady=(0,4))

        pro_card = tk.Frame(right, bg=CARD); pro_card.pack(fill="x")
        pro_inner = tk.Frame(pro_card, bg=CARD); pro_inner.pack(fill="x", padx=12, pady=10)
        ttk.Button(pro_inner, text="🏆 プロデータを取得", command=self._fetch_pros).pack(anchor="w")
        self._lbl_pro_status = tk.Label(pro_inner, text="", bg=CARD, fg=DIM, font=FS)
        self._lbl_pro_status.pack(anchor="w", pady=(6,0))

        pro_wrap = tk.Frame(right, bg=BG); pro_wrap.pack(fill="both", expand=True, pady=(8,0))
        pro_canvas = tk.Canvas(pro_wrap, bg=BG, highlightthickness=0)
        pro_sb = ttk.Scrollbar(pro_wrap, orient="vertical", command=pro_canvas.yview)
        self._pro_result_frame = tk.Frame(pro_canvas, bg=BG)
        self._pro_result_frame.bind("<Configure>",
            lambda e: pro_canvas.configure(scrollregion=pro_canvas.bbox("all")))
        pro_canvas.create_window((0,0), window=self._pro_result_frame, anchor="nw")
        pro_canvas.configure(yscrollcommand=pro_sb.set)
        pro_canvas.pack(side="left", fill="both", expand=True)
        pro_sb.pack(side="right", fill="y")


    def _fetch_pros(self):
        self._lbl_pro_status.config(text="取得中…", fg=YELLOW)
        threading.Thread(target=self._fetch_pros_thread, daemon=True).start()

    def _fetch_pros_thread(self):
        try:
            url = "https://raw.githubusercontent.com/lap1sss-dev/VALORANT-Sensitivity-Tracker/main/pros.json"
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req, timeout=15) as r:
                pros = json.loads(r.read().decode("utf-8"))
            self.after(0, lambda: self._show_pros(pros))
        except Exception as e:
            self.after(0, lambda: self._lbl_pro_status.config(
                text=f"取得失敗: {e}", fg=ACCENT))

    def _show_pros(self, pros):
        for w in self._pro_result_frame.winfo_children(): w.destroy()
        self._lbl_pro_status.config(text=f"✔ {len(pros)}人のデータを取得しました", fg=GREEN)

        my = db().execute("""
            SELECT AVG(CAST(kills AS REAL)/NULLIF(deaths,0)),
                   AVG(acs), AVG(damage), AVG(hs_percent)
            FROM match_stats""").fetchone()
        my_kd, my_acs, my_dmg, my_hs = my

        for pro in pros:
            card = tk.Frame(self._pro_result_frame, bg=CARD)
            card.pack(fill="x", pady=3, padx=2)

            hdr = tk.Frame(card, bg=CARD); hdr.pack(fill="x", padx=8, pady=(8,4))
            name_txt = f"{pro.get('name','')}  {pro.get('team','')}"
            tk.Label(hdr, text=name_txt, bg=CARD, fg=ACCENT2,
                     font=("NotoSansJP",11,"bold")).pack(side="left")
            if pro.get("note"):
                tk.Label(hdr, text=pro["note"], bg=CARD, fg=DIM, font=FS).pack(side="right")

            def row(title, val, my_val, fmt, _card=card):
                r = tk.Frame(_card, bg=CARD); r.pack(fill="x", padx=12, pady=2)
                tk.Label(r, text=title, bg=CARD, fg=DIM, font=FS, width=10, anchor="w").pack(side="left")
                txt = fmt.format(val) if val is not None else "--"
                if val is not None and my_val is not None:
                    fg = GREEN if val >= my_val else "#e74c3c"
                else:
                    fg = TEXT
                tk.Label(r, text=txt, bg=CARD, fg=fg,
                         font=("NotoSansJP",11,"bold")).pack(side="left")

            row("KD",      pro.get("kd"),        my_kd,  "{:.2f}")
            row("ACS",     pro.get("acs"),        my_acs, "{:.0f}")
            row("安定度",  pro.get("stability"),  None,   "{:.2f}")
            row("平均DMG", pro.get("avg_dmg"),    my_dmg, "{:.0f}")
            row("HS%",     pro.get("hs"),         my_hs,  "{:.1f}%")


    # ───────────────────────────────────
    #  設定タブ
    # ───────────────────────────────────
    def _build_settings(self):
        f = self.tab_settings

        # ── ホットキー設定 ──────────────────────
        tk.Label(f, text="ホットキー設定", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(18,6))
        hk_card = tk.Frame(f, bg=CARD); hk_card.pack(fill="x", padx=20)
        inner = tk.Frame(hk_card, bg=CARD); inner.pack(fill="x", padx=20, pady=16)

        # 現在のホットキー表示
        r0 = tk.Frame(inner, bg=CARD); r0.pack(fill="x", pady=4)
        tk.Label(r0, text="現在のホットキー:", bg=CARD, fg=DIM, font=FS, width=18, anchor="w").pack(side="left")
        self._lbl_hk_cur = tk.Label(r0, bg=CARD, fg=ACCENT, font=("NotoSansJP",14,"bold"))
        self._lbl_hk_cur.pack(side="left")

        # ── キー入力ウィジェット ──────────────
        tk.Label(inner, text="ホットキーを入力してください:", bg=CARD, fg=DIM, font=FS).pack(anchor="w", pady=(12,4))

        # 修飾キーチェックボックス
        mod_frame = tk.Frame(inner, bg=CARD); mod_frame.pack(anchor="w", pady=4)
        tk.Label(mod_frame, text="修飾キー:", bg=CARD, fg=TEXT, font=FB).pack(side="left", padx=(0,10))
        self._mod_vars = {}
        for mod in ["ctrl","shift","alt","win"]:
            v = tk.BooleanVar()
            self._mod_vars[mod] = v
            cb = tk.Checkbutton(mod_frame, text=mod.capitalize(), variable=v,
                                bg=CARD, fg=TEXT, selectcolor=BG2,
                                activebackground=CARD, activeforeground=TEXT,
                                font=FB)
            cb.pack(side="left", padx=6)

        # メインキー入力
        key_frame = tk.Frame(inner, bg=CARD); key_frame.pack(anchor="w", pady=4)
        tk.Label(key_frame, text="キー:", bg=CARD, fg=TEXT, font=FB).pack(side="left", padx=(0,10))
        self._entry_mainkey = ttk.Entry(key_frame, width=12)
        self._entry_mainkey.pack(side="left")
        tk.Label(key_frame, text="（例: s / f6 / f12）", bg=CARD, fg=DIM, font=FS).pack(side="left", padx=8)

        # または F キーコンボ
        fkey_frame = tk.Frame(inner, bg=CARD); fkey_frame.pack(anchor="w", pady=4)
        tk.Label(fkey_frame, text="またはFキーを選択:", bg=CARD, fg=TEXT, font=FB).pack(side="left", padx=(0,10))
        self._combo_fkey = ttk.Combobox(fkey_frame, width=8, state="readonly",
                                        values=["(なし)","F1","F2","F3","F4","F5","F6",
                                                "F7","F8","F9","F10","F11","F12"])
        self._combo_fkey.current(0)
        self._combo_fkey.pack(side="left")

        # プレビュー
        prev_frame = tk.Frame(inner, bg=CARD); prev_frame.pack(anchor="w", pady=(10,0))
        tk.Label(prev_frame, text="プレビュー:", bg=CARD, fg=DIM, font=FS).pack(side="left")
        self._lbl_hk_preview = tk.Label(prev_frame, text="", bg=CARD, fg=YELLOW,
                                        font=("NotoSansJP",12,"bold"))
        self._lbl_hk_preview.pack(side="left", padx=8)

        # プレビュー更新バインド
        for v in self._mod_vars.values():
            v.trace_add("write", self._preview_hotkey)
        self._entry_mainkey.bind("<KeyRelease>", self._preview_hotkey)
        self._combo_fkey.bind("<<ComboboxSelected>>", self._preview_hotkey)

        # 適用ボタン
        ttk.Button(inner, text="✔ このホットキーを適用", command=self._apply_hotkey).pack(anchor="w", pady=(12,0))

        # 使い方ヒント
        hint = ("使い方:\n"
                "  修飾キー（Ctrl / Shift / Alt）をチェックしてから、メインキーを入力してください。\n"
                "  例1: Ctrl + Shift をチェック → キーに「s」→ Ctrl+Shift+S\n"
                "  例2: Fキーから「F6」を選択するだけ → F6 単独\n"
                "  例3: Ctrl をチェック → Fキーから「F9」 → Ctrl+F9")
        tk.Label(inner, text=hint, bg=CARD, fg=DIM, font=FS, justify="left").pack(anchor="w", pady=(10,0))

        # ── Henrikdev API設定 ──────────────────
        tk.Label(f, text="Henrikdev API 設定", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(18,6))
        hc = tk.Frame(f, bg=CARD); hc.pack(fill="x", padx=20)
        hi = tk.Frame(hc, bg=CARD); hi.pack(fill="x", padx=20, pady=14)

        # プレイヤー名
        nr = tk.Frame(hi, bg=CARD); nr.pack(anchor="w", pady=3)
        tk.Label(nr, text="プレイヤー名:", bg=CARD, fg=TEXT, font=FB, width=14, anchor="w").pack(side="left")
        self._entry_pname = ttk.Entry(nr, width=24)
        self._entry_pname.insert(0, self.cfg.get("player_name",""))
        self._entry_pname.pack(side="left", padx=4)
        tk.Label(nr, text="タグ (#JP1 など):", bg=CARD, fg=TEXT, font=FB).pack(side="left", padx=(12,0))
        self._entry_ptag = ttk.Entry(nr, width=12)
        self._entry_ptag.insert(0, self.cfg.get("player_tag",""))
        self._entry_ptag.pack(side="left", padx=4)

        # リージョン
        rr2 = tk.Frame(hi, bg=CARD); rr2.pack(anchor="w", pady=3)
        tk.Label(rr2, text="リージョン:", bg=CARD, fg=TEXT, font=FB, width=14, anchor="w").pack(side="left")
        self._combo_region = ttk.Combobox(rr2, values=["ap","na","eu","kr","latam","br"], width=8, state="readonly")
        self._combo_region.set(self.cfg.get("region","ap"))
        self._combo_region.pack(side="left", padx=4)

        # APIキー
        ar = tk.Frame(hi, bg=CARD); ar.pack(anchor="w", pady=3)
        tk.Label(ar, text="APIキー:", bg=CARD, fg=TEXT, font=FB, width=14, anchor="w").pack(side="left")
        self._entry_apikey = ttk.Entry(ar, width=46)
        self._entry_apikey.insert(0, self.cfg.get("henrik_api_key",""))
        self._entry_apikey.pack(side="left", padx=4)

        ttk.Button(hi, text="✔ 保存", command=self._save_henrik).pack(anchor="w", pady=(8,0))
        tk.Label(hi, text="APIキーの取得: https://docs.henrikdev.xyz",
                 bg=CARD, fg=DIM, font=FS).pack(anchor="w", pady=(4,0))

        # ── Tesseract設定 ──────────────────────
        tk.Label(f, text="OCR設定（Tesseract）", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(22,6))
        oc = tk.Frame(f, bg=CARD); oc.pack(fill="x", padx=20)
        oi = tk.Frame(oc, bg=CARD); oi.pack(fill="x", padx=20, pady=14)

        pil_ok = HAS_PIL; ocr_ok = HAS_OCR; pyn_ok = HAS_PYNPUT
        statuses = [
            ("Pillow",      pil_ok, "pip install pillow"),
            ("pytesseract", ocr_ok, "pip install pytesseract"),
            ("pynput",      pyn_ok, "pip install pynput"),
        ]
        for lib, ok, cmd in statuses:
            row = tk.Frame(oi, bg=CARD); row.pack(anchor="w", pady=2)
            tk.Label(row, text="✔" if ok else "✘", bg=CARD,
                     fg=GREEN if ok else ACCENT, font=FB, width=2).pack(side="left")
            tk.Label(row, text=f"{lib}", bg=CARD, fg=TEXT if ok else DIM,
                     font=FB, width=14, anchor="w").pack(side="left")
            if not ok:
                tk.Label(row, text=f"→  {cmd}", bg=CARD, fg=ORANGE, font=FS).pack(side="left", padx=6)

        tk.Label(oi, text="\nTesseract.exe パス:", bg=CARD, fg=DIM, font=FS).pack(anchor="w")
        tr = tk.Frame(oi, bg=CARD); tr.pack(fill="x")
        self._entry_tess = ttk.Entry(tr, width=54)
        self._entry_tess.insert(0, self.cfg.get("tesseract_path",""))
        self._entry_tess.pack(side="left")
        ttk.Button(tr, text="保存", style="Sub.TButton",
                   command=self._save_tess).pack(side="left", padx=6)
        tk.Label(oi, text="Tesseract をインストール: https://github.com/tesseract-ocr/tesseract/releases",
                 bg=CARD, fg=DIM, font=FS).pack(anchor="w", pady=(4,0))

        # ── DPI設定 ────────────────────────────
        tk.Label(f, text="デフォルト DPI", bg=BG, fg=ACCENT2, font=FH).pack(anchor="w",padx=20,pady=(18,6))
        dc = tk.Frame(f, bg=CARD); dc.pack(fill="x", padx=20)
        di = tk.Frame(dc, bg=CARD); di.pack(fill="x", padx=20, pady=12)
        rr = tk.Frame(di, bg=CARD); rr.pack(anchor="w")
        tk.Label(rr, text="DPI:", bg=CARD, fg=TEXT, font=FB).pack(side="left")
        self._entry_dpi = ttk.Entry(rr, width=8)
        self._entry_dpi.insert(0, str(self.cfg.get("dpi",800)))
        self._entry_dpi.pack(side="left", padx=8)
        ttk.Button(rr, text="保存", style="Sub.TButton", command=self._save_dpi).pack(side="left")

        # データフォルダ
        tk.Label(f, text=f"データ保存フォルダ:  {APP_DIR}",
                 bg=BG, fg=DIM, font=FS).pack(anchor="w",padx=20,pady=(16,4))
        ttk.Button(f, text="📂 フォルダを開く", style="Sub.TButton",
                   command=lambda: os.startfile(APP_DIR)).pack(anchor="w",padx=20)

        # 初期値をUIに反映
        self._load_hotkey_to_ui()
        self._preview_hotkey()

    # ─────────────────────────────────────
    #  設定ヘルパー
    # ─────────────────────────────────────
    def _load_hotkey_to_ui(self):
        """保存済みホットキーをチェックボックス・エントリに反映"""
        parts = self.cfg.get("hotkey_parts", ["ctrl","shift","s"])
        mods  = {"ctrl","shift","alt","win"}
        for m, v in self._mod_vars.items():
            v.set(m in parts)
        # メインキー
        main = [p for p in parts if p.lower() not in mods]
        if main:
            mk = main[0]
            if mk.startswith("f") and mk[1:].isdigit():
                self._combo_fkey.set(mk.upper())
                self._entry_mainkey.delete(0,"end")
            else:
                self._entry_mainkey.delete(0,"end")
                self._entry_mainkey.insert(0, mk)
                self._combo_fkey.current(0)
        disp = hotkey_display(parts)
        if hasattr(self,"_lbl_hk_cur"):    self._lbl_hk_cur.config(text=disp)
        if hasattr(self,"_lbl_hk_home"):   self._lbl_hk_home.config(text=f"  ⌨  {disp}  ")

    def _preview_hotkey(self, *_):
        parts = self._collect_hotkey_parts()
        self._lbl_hk_preview.config(text=hotkey_display(parts) if parts else "（キーを選択してください）")

    def _collect_hotkey_parts(self):
        parts = [m for m, v in self._mod_vars.items() if v.get()]
        fkey = self._combo_fkey.get()
        if fkey and fkey != "(なし)":
            parts.append(fkey.lower())
        else:
            mk = self._entry_mainkey.get().strip().lower()
            if mk: parts.append(mk)
        return parts

    def _apply_hotkey(self):
        parts = self._collect_hotkey_parts()
        if not parts:
            messagebox.showerror("エラー","ホットキーを設定してください"); return
        # 修飾キーだけの場合を防ぐ
        mods = {"ctrl","shift","alt","win"}
        has_main = any(p not in mods for p in parts)
        if not has_main:
            messagebox.showerror("エラー","修飾キーだけでなくメインキーも設定してください"); return
        self.cfg["hotkey_parts"] = parts
        save_cfg(self.cfg)
        self.hotkey.start(parts)
        disp = hotkey_display(parts)
        self._lbl_hk_cur.config(text=disp)
        self._lbl_hk_home.config(text=f"  ⌨  {disp}  ")
        self._update_status_bar(True)
        messagebox.showinfo("設定完了", f"ホットキーを  {disp}  に設定しました！")

    def _save_henrik(self):
        self.cfg["player_name"]    = self._entry_pname.get().strip()
        self.cfg["player_tag"]     = self._entry_ptag.get().strip().lstrip("#")
        self.cfg["region"]         = self._combo_region.get().strip()
        self.cfg["henrik_api_key"] = self._entry_apikey.get().strip()
        save_cfg(self.cfg)
        # PUUIDを取得してconfigに保存
        threading.Thread(target=self._fetch_puuid, daemon=True).start()
        messagebox.showinfo("保存","Henrikdev API設定を保存しました。PUUIDを取得中…")

    def _fetch_puuid(self):
        import urllib.request, urllib.parse as up
        try:
            name    = self.cfg.get("player_name","")
            tag     = self.cfg.get("player_tag","")
            api_key = self.cfg.get("henrik_api_key","")
            url = (f"https://api.henrikdev.xyz/valorant/v1/account/"
                   f"{up.quote(name,safe='')}/{up.quote(tag,safe='')}")
            req = urllib.request.Request(url, headers={"Authorization": api_key})
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read().decode("utf-8","replace"))
            puuid = data.get("data",{}).get("puuid","")
            if puuid:
                self.cfg["player_puuid"] = puuid
                save_cfg(self.cfg)
                self.after(0, lambda p=puuid: messagebox.showinfo("完了", f"PUUID取得成功！\n{p[:16]}…"))
            else:
                self.after(0, lambda: messagebox.showwarning("警告", "PUUIDの取得に失敗しました。名前・タグを確認してください"))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("エラー", f"PUUID取得失敗: {e}"))

    def _save_tess(self):
        p = self._entry_tess.get().strip()
        self.cfg["tesseract_path"] = p
        save_cfg(self.cfg)
        if HAS_OCR and p:
            pytesseract.pytesseract.tesseract_cmd = p
        messagebox.showinfo("保存","Tesseractパスを保存しました")

    def _save_dpi(self):
        try:
            d = int(self._entry_dpi.get().strip())
            self.cfg["dpi"] = d; save_cfg(self.cfg)
            messagebox.showinfo("保存",f"DPI を {d} に保存しました")
        except:
            messagebox.showerror("エラー","整数で入力してください")

    # ─────────────────────────────────────
    #  感度保存（手動）
    # ─────────────────────────────────────
    def _save_sens_manual(self):
        try:
            v = float(self._e_sv.get())
        except:
            messagebox.showerror("エラー","感度の値が正しくありません"); return
        dpi_s = self._e_sd.get().strip()
        try: dpi = int(dpi_s) if dpi_s else self.cfg.get("dpi")
        except: dpi = self.cfg.get("dpi")
        note = self._e_sn.get().strip()
        self._save_sens_db(v, dpi, note, method="手動")
        for w in (self._e_sv, self._e_sd, self._e_sn):
            w.delete(0,"end")
        self._e_se.configure(state="normal"); self._e_se.delete(0,"end"); self._e_se.configure(state="readonly")
        messagebox.showinfo("保存",f"感度 {v} を記録しました！")
        self.refresh_all()

    def _save_sens_db(self, sens, dpi, note, method="手動", ss_path=None):
        edpi = sens * dpi if dpi else None
        now  = datetime.now().strftime("%Y-%m-%d %H:%M")
        c = db()
        c.execute("INSERT INTO sensitivity_log (sensitivity,dpi,edpi,note,screenshot,method,changed_at)"
                  " VALUES (?,?,?,?,?,?,?)",
                  (sens, dpi, edpi, note, None, method, now))
        c.commit(); c.close()

    # ─────────────────────────────────────
    #  データ読み込み・更新
    # ─────────────────────────────────────
    def refresh_all(self):
        self._load_current()
        self._load_sens_tree()
        self._load_recent()
        self._update_match_combo()
        self.refresh_analysis()
        self._load_hotkey_to_ui()
        self._update_status_bar(self.hotkey._listener is not None)

    def _load_current(self):
        row = db().execute(
            "SELECT sensitivity,dpi,edpi,changed_at FROM sensitivity_log ORDER BY id DESC LIMIT 1"
        ).fetchone()
        if row:
            self._lbl_sens.config(text=str(row[0]))
            self._lbl_edpi.config(text=f"eDPI: {row[2]:.1f}" if row[2] else f"DPI: {row[1] or '--'}")
            self._lbl_since.config(text=f"変更日: {row[3]}")
        else:
            self._lbl_sens.config(text="--")

    def _load_sens_tree(self):
        for i in self._tree_sens.get_children(): self._tree_sens.delete(i)
        rows = db().execute("""
            SELECT s.id,s.sensitivity,s.dpi,s.edpi,COUNT(m.id),s.method,s.changed_at,s.note
            FROM sensitivity_log s LEFT JOIN match_stats m ON m.sensitivity_id=s.id
            GROUP BY s.id ORDER BY s.id DESC""").fetchall()
        for n, r in enumerate(rows):
            edpi = f"{r[3]:.1f}" if r[3] else "--"
            tag = "e" if n%2==0 else "o"
            self._tree_sens.insert("","end", values=(r[0],r[1],r[2] or "--",edpi,r[4],r[5] or "--",r[6],r[7] or ""), tags=(tag,))
        self._tree_sens.tag_configure("e", background=BG2)
        self._tree_sens.tag_configure("o", background=CARD)

    def _load_recent(self):
        for i in self._tree_recent.get_children(): self._tree_recent.delete(i)
        rows = db().execute("""
            SELECT m.match_date,m.agent,m.map,m.kills,m.deaths,m.assists,
                   m.acs,m.hs_percent,m.won,s.sensitivity
            FROM match_stats m JOIN sensitivity_log s ON s.id=m.sensitivity_id
            ORDER BY m.id DESC LIMIT 30""").fetchall()
        for r in rows:
            self._tree_recent.insert("","end", values=(
                r[0] or "--", r[1] or "--", r[2] or "--",
                f"{r[3]}/{r[4]}/{r[5]}",
                f"{r[6]:.0f}" if r[6] else "--",
                f"{r[7]:.1f}%" if r[7] else "--",
                "✔ 勝" if r[8] else "✘ 負",
                r[9]))

    def _update_match_combo(self):
        rows = db().execute(
            "SELECT id,sensitivity,changed_at FROM sensitivity_log ORDER BY id DESC").fetchall()
        vals = [f"[{r[0]}] 感度 {r[1]}  ({r[2]})" for r in rows]
        self._combo_msens["values"] = vals
        if vals: self._combo_msens.current(0); self._filter_matches()

    def _filter_matches(self, *_):
        sel = self._combo_msens.get()
        if not sel: return
        sid = int(re.match(r"\[(\d+)\]", sel).group(1))
        for i in self._tree_matches.get_children(): self._tree_matches.delete(i)
        
        # ▼ 修正: 日時が新しい順に並び替え
        rows = db().execute("""
            SELECT id,match_date,agent,map,kills,deaths,assists,acs,damage,hs_percent,won,note
            FROM match_stats WHERE sensitivity_id=? ORDER BY match_date DESC, id DESC""", (sid,)).fetchall()
        
        for n, r in enumerate(rows):
            tag = "e" if n%2==0 else "o"
            
            # ▼ 修正: valuesからr[0]を外し、代わりに text=str(r[0]) として裏側にIDを隠し持つ
            self._tree_matches.insert("","end", text=str(r[0]), values=(
                r[1] or "--", r[2] or "--", r[3] or "--",
                r[4],r[5],r[6],
                f"{r[7]:.0f}" if r[7] else "--",
                f"{r[8]:.0f}" if r[8] else "--",
                f"{r[9]:.1f}%" if r[9] else "--",
                "✔ 勝" if r[10] else "✘ 負",
                r[11] or ""), tags=(tag,))
                
        self._tree_matches.tag_configure("e", background=BG2)
        self._tree_matches.tag_configure("o", background=CARD)
    # ─────────────────────────────────────
    #  分析
    # ─────────────────────────────────────
    def refresh_analysis(self):
        for w in self._analysis_frame.winfo_children(): w.destroy()

        # ── メイン集計クエリ ──────────────────────────────────────
        raw = db().execute("""
            WITH rs_agg AS (
                SELECT match_stats_id,
                       AVG(kills)  AS avg_k,
                       AVG(damage) AS avg_d,
                       SUM(hs_hits)    AS sum_hs,
                       SUM(total_hits) AS sum_total
                FROM round_stats
                GROUP BY match_stats_id
            )
            SELECT s.sensitivity,
                   COUNT(m.id)                                   AS cnt,
                   AVG(CAST(m.kills AS REAL)/NULLIF(m.deaths,0)) AS kd,
                   AVG(m.acs)                                    AS acs,
                   AVG(m.hs_percent)                             AS hs,
                   SUM(m.won)                                    AS wins,
                   COUNT(m.won)                                  AS total,
                   MIN(s.dpi)                                    AS dpi,
                   MIN(s.edpi)                                   AS edpi,
                   MIN(s.changed_at)                             AS first_used,
                   MAX(s.changed_at)                             AS last_used,
                   COUNT(DISTINCT s.id)                          AS periods,
                   AVG(rs.avg_k)                                 AS avg_round_kills,
                   AVG(rs.avg_d)                                 AS avg_round_dmg,
                   CASE WHEN SUM(rs.sum_total)>0
                        THEN CAST(SUM(rs.sum_hs) AS REAL)/SUM(rs.sum_total)*100
                        ELSE NULL END                            AS round_hs_pct
            FROM sensitivity_log s
            LEFT JOIN match_stats m ON m.sensitivity_id = s.id
            LEFT JOIN rs_agg rs ON rs.match_stats_id = m.id
            GROUP BY s.sensitivity
            ORDER BY s.sensitivity""").fetchall()

        if not raw:
            ttk.Label(self._analysis_frame, text="データがありません。", foreground=DIM).pack(pady=40)
            return

        # ── ACS 標準偏差を感度ごとに計算 ─────────────────────────
        # SQLiteにはSTDEV関数がないためPython側で計算する
        acs_rows = db().execute("""
            SELECT s.sensitivity, m.acs
            FROM sensitivity_log s
            JOIN match_stats m ON m.sensitivity_id = s.id
            WHERE m.acs IS NOT NULL
        """).fetchall()
        # {感度値: [acs, acs, ...]}
        acs_by_sens = {}
        for sens_v, acs_v in acs_rows:
            acs_by_sens.setdefault(sens_v, []).append(acs_v)

        def calc_stability(vals):
            """安定度 = 平均ACS ÷ 標準偏差（高いほど安定）。2試合未満はNone"""
            if len(vals) < 2:
                return None
            mean = sum(vals) / len(vals)
            sd = math.sqrt(sum((x - mean) ** 2 for x in vals) / (len(vals) - 1))
            return mean / sd if sd > 0 else None

        stdev_map = {s: calc_stability(v) for s, v in acs_by_sens.items()}

        # ── 「試合数少」の閾値を動的決定 ──────────────────────────
        # 方針：全感度の総試合数を母数とし、その 10% を下限、
        #       ただし絶対値として最低 5、最大 10 でクランプする。
        #       → 全体が50試合なら閾値5、100試合なら10、それ以上も10。
        total_all = sum(r[1] for r in raw if r[1])
        few_threshold = max(5, min(10, round(total_all * 0.10)))

        with_acs  = [(r[0], r[3]) for r in raw if r[1] and r[3]]
        best_sens = max(with_acs, key=lambda x: x[1])[0] if with_acs else None

        # ════════════════════════════════════════════
        #  説明パネル（タブ上部）
        # ════════════════════════════════════════════
        info_bg = "#0d1a0d"
        info = tk.Frame(self._analysis_frame, bg=info_bg)
        info.pack(fill="x", pady=(0, 4))
        tk.Label(info,
            text="💡 同じ感度値は統合して表示しています（複数の期間に使った感度も1行にまとめています）\n"
                 "💡 上段に[試合全体]、下段に[ラウンド毎]のスタッツを表示しています",
            bg=info_bg, fg=GREEN, font=FS, justify="left").pack(padx=12, pady=(8, 2), anchor="w")

        # ACS安定度（標準偏差）の解説ボックス
        stdev_bg = "#101828"
        stdev_box = tk.Frame(self._analysis_frame, bg=stdev_bg, bd=0)
        stdev_box.pack(fill="x", padx=10, pady=(0, 8))
        tk.Label(stdev_box,
            text=(
                "  安定度 = 平均ACS ÷ 標準偏差（高いほど安定）\n"
                "\n"
                "   S  4.2以上    → 非常に安定\n"
                "   A  3.6〜4.2  → 安定\n"
                "   B  3.0〜3.6  → 普通\n"
                "   C  2.4〜3.0  → やや不安定\n"
                "   D  2.4未満   → 不安定\n"
                "\n"
                f"  ⚠ 試合数が全体の約10%（現在の閾値: {few_threshold}試合）未満の感度には"
                " 「⚠試合数少」バッジを表示します。\n"
            ),
            bg=stdev_bg, fg=TEXT, font=FS, justify="left").pack(anchor="w", padx=18, pady=(0, 10))

        # ════════════════════════════════════════════
        #  感度カード一覧
        # ════════════════════════════════════════════
        for r in raw:
            sens, cnt, kd, acs, hs, wins, total, dpi_v, edpi_v, \
                first_used, last_used, periods, avg_rk, avg_rdmg, round_hs = r
            if not cnt:
                continue

            is_best  = (sens == best_sens)
            is_few   = (cnt < few_threshold)          # 試合数が閾値未満
            stdev    = stdev_map.get(sens)            # ACS標準偏差（Noneの場合あり）

            bg_c = "#0e2a1a" if is_best else CARD
            card = tk.Frame(self._analysis_frame, bg=bg_c)
            card.pack(fill="x", pady=4, padx=10)

            # ── ヘッダー行 ─────────────────────────────
            hdr = tk.Frame(card, bg=bg_c)
            hdr.pack(fill="x", padx=10, pady=(6, 2))

            tk.Label(hdr, text=f"感度: {sens}", bg=bg_c, fg=ACCENT,
                     font=("NotoSansJP", 14, "bold")).pack(side="left")

            period_txt = f"{periods}回" if periods else "--"
            edpi_disp  = f"{edpi_v:.0f}" if edpi_v else "--"
            tk.Label(hdr,
                text=f"  (eDPI: {edpi_disp})  |  試合数: {cnt}回  |  使用期間数: {period_txt}",
                bg=bg_c, fg=DIM, font=FB).pack(side="left")

            # 右端バッジ群（右から積む → 最右が最初にpack）
            if is_best:
                tk.Label(hdr, text="★ 最高ACS", bg=bg_c, fg=GREEN,
                         font=("NotoSansJP", 10, "bold")).pack(side="right", padx=(4, 0))
            if is_few:
                tk.Label(hdr, text="⚠ 試合数少", bg="#2a1500", fg=ORANGE,
                         font=("NotoSansJP", 9, "bold"), padx=6, pady=2).pack(side="right", padx=(4, 0))

            # ── 上段：試合全体スタッツ ──────────────────
            row1 = tk.Frame(card, bg=bg_c)
            row1.pack(fill="x", padx=14, pady=(4, 1))
            tk.Label(row1, text="[試合全体]", bg=bg_c, fg=ACCENT2,
                     font=("NotoSansJP", 10, "bold"), width=10, anchor="w").pack(side="left")

            win_s = f"{wins/total*100:.0f}%" if (total and wins is not None) else "--"
            win_c = GREEN if (total and wins and wins / total >= 0.5) else ("#e74c3c" if total else TEXT)
            kd_v  = f"{kd:.2f}" if kd else "--"
            kd_c  = GREEN if (kd and kd >= 1.0) else ("#e74c3c" if kd else TEXT)
            acs_v = f"{acs:.0f}" if acs else "--"

            tk.Label(row1, text=f"勝率: {win_s}", bg=bg_c, fg=win_c, font=FB, width=14, anchor="w").pack(side="left")
            tk.Label(row1, text=f"K/D: {kd_v}",  bg=bg_c, fg=kd_c,  font=FB, width=12, anchor="w").pack(side="left")
            tk.Label(row1, text=f"ACS: {acs_v}", bg=bg_c, fg=TEXT,  font=FB, width=12, anchor="w").pack(side="left")

            # ACS安定度ラベル（stdevが取れた場合のみ）
            if stdev is not None:
                if stdev >= 4.2:
                    stdev_mark, stdev_fg = "S  非常に安定", GREEN
                elif stdev >= 3.6:
                    stdev_mark, stdev_fg = "A  安定",       ACCENT2
                elif stdev >= 3.0:
                    stdev_mark, stdev_fg = "B  普通",       TEXT
                elif stdev >= 2.4:
                    stdev_mark, stdev_fg = "C  やや不安定", YELLOW
                else:
                    stdev_mark, stdev_fg = "D  不安定",     "#e74c3c"
                tk.Label(row1,
                    text=f"安定度: {stdev:.2f}  {stdev_mark}",
                    bg=bg_c, fg=stdev_fg, font=FB, anchor="w").pack(side="left", padx=(8, 0))
            elif cnt == 1:
                tk.Label(row1, text="安定度: 試合数1につき計算不能",
                         bg=bg_c, fg=DIM, font=FS, anchor="w").pack(side="left", padx=(8, 0))
            # acs自体がない場合は何も出さない（--だと紛らわしい）

            # ── 下段：ラウンド毎スタッツ ────────────────
            row2 = tk.Frame(card, bg=bg_c)
            row2.pack(fill="x", padx=14, pady=1)
            tk.Label(row2, text="[ラウンド]", bg=bg_c, fg=YELLOW,
                     font=("NotoSansJP", 10, "bold"), width=10, anchor="w").pack(side="left")

            rk_v   = f"{avg_rk:.2f}"    if avg_rk   is not None else "--"
            rdmg_v = f"{avg_rdmg:.0f}"  if avg_rdmg is not None else "--"
            rhs_v  = f"{round_hs:.1f}%" if round_hs is not None else "--"

            tk.Label(row2, text=f"平均キル: {rk_v}",  bg=bg_c, fg=TEXT, font=FB, width=14, anchor="w").pack(side="left")
            tk.Label(row2, text=f"平均DMG: {rdmg_v}", bg=bg_c, fg=TEXT, font=FB, width=12, anchor="w").pack(side="left")
            tk.Label(row2, text=f"HS率: {rhs_v}",     bg=bg_c, fg=TEXT, font=FB, width=12, anchor="w").pack(side="left")

            # ── フッター：使用期間 ───────────────────────
            row3 = tk.Frame(card, bg=bg_c)
            row3.pack(fill="x", padx=14, pady=(1, 8))
            period_disp = first_used if first_used == last_used \
                else f"{first_used[:10]} 〜 {last_used[:10]}"
            tk.Label(row3, text=f"使用期間: {period_disp}",
                     bg=bg_c, fg=DIM, font=FS).pack(side="left", padx=(80, 0))
    # ─────────────────────────────────────
    #  CRUD ヘルパー
    # ─────────────────────────────────────
    def _calc_edpi(self, *_):
        try:
            self._e_se.configure(state="normal")
            self._e_se.delete(0,"end")
            self._e_se.insert(0, f"{float(self._e_sv.get()) * int(self._e_sd.get()):.1f}")
            self._e_se.configure(state="readonly")
        except: pass

    def _field(self, parent, label, width, ro=False):
        f = ttk.Frame(parent); f.pack(side="left", padx=(0,14))
        ttk.Label(f, text=label, font=FS, foreground=DIM).pack(anchor="w")
        e = ttk.Entry(f, width=width)
        if ro: e.configure(state="readonly")
        e.pack(); return e

    def _open_add_sens(self):
        self.nb.select(self.tab_sens); self._e_sv.focus()

    def _open_add_match(self):
        MatchDialog(self)

    def _sens_dbl(self, *_):
        sel = self._tree_sens.selection()
        if not sel: return
        sid = self._tree_sens.item(sel[0])["values"][0]
        self.nb.select(self.tab_matches)
        for i, v in enumerate(self._combo_msens["values"]):
            if str(v).startswith(f"[{sid}]"):
                self._combo_msens.current(i); self._filter_matches(); break

    def _del_sens(self):
        sel = self._tree_sens.selection()
        if not sel: messagebox.showinfo("","削除する感度を選択してください"); return
        sid = self._tree_sens.item(sel[0])["values"][0]
        if messagebox.askyesno("確認","この感度とその試合データをすべて削除しますか？"):
            c = db()
            # ▼追加: 紐づくラウンドスタッツを先に一括削除
            c.execute("DELETE FROM round_stats WHERE match_stats_id IN (SELECT id FROM match_stats WHERE sensitivity_id=?)", (sid,))
            
            c.execute("DELETE FROM match_stats WHERE sensitivity_id=?", (sid,))
            c.execute("DELETE FROM sensitivity_log WHERE id=?", (sid,))
            c.commit(); c.close(); self.refresh_all()

    def _del_match(self):
        sel = self._tree_matches.selection()
        if not sel: messagebox.showinfo("","削除する試合を選択してください"); return
        
        # ▼ 修正: 裏側に隠しておいた text 属性からDBのIDを取得する
        mid = int(self._tree_matches.item(sel[0])["text"])
        
        if messagebox.askyesno("確認","この試合を削除しますか？"):
            c = db()
            # 紐づくラウンドスタッツを先に削除（これは維持）
            c.execute("DELETE FROM round_stats WHERE match_stats_id=?", (mid,))
            
            c.execute("DELETE FROM match_stats WHERE id=?", (mid,))
            c.commit(); c.close(); self.refresh_all() 

    def _set_as_current(self):
        sel = self._tree_sens.selection()
        if not sel:
            messagebox.showinfo("", "感度を選択してください")
            return
        vals = self._tree_sens.item(sel[0])["values"]
        sens = vals[1]
        dpi  = vals[2]
        try: dpi = int(dpi)
        except: dpi = self.cfg.get("dpi", 800)

        if not messagebox.askyesno("確認", f"感度 {sens} を現在の感度として設定しますか？\n"
                                            "（履歴に新しいエントリとして追加されます）"):
            return
        self._save_sens_db(float(sens), dpi, note="履歴から再設定", method="手動")
        self.refresh_all()
        messagebox.showinfo("完了", f"感度 {sens} を現在の感度に設定しました")

# ─────────────────────────────────────────────
#  OCR確認ダイアログ（ホットキー発火後）
# ─────────────────────────────────────────────
class SensConfirmDialog(tk.Toplevel):
    def __init__(self, parent, detected, img, ss_path, raw):
        super().__init__(parent)
        self.parent   = parent
        self.ss_path  = ss_path
        self.detected = detected
        self.title("感度を確認して保存")
        self.configure(bg=BG)
        self.geometry("600x520")
        self.resizable(False, False)
        self.grab_set()
        self.attributes("-topmost", True)
        self._build(img, raw)

    def _build(self, img, raw):
        # タイトル
        tk.Label(self, text="📸  スクリーンショットから感度を検出", bg=BG, fg=ACCENT2,
                 font=FH).pack(pady=(16,6))

        # スクショサムネイル
        try:
            thumb = img.copy(); thumb.thumbnail((560, 200))
            from PIL import ImageTk
            self._thumb_img = ImageTk.PhotoImage(thumb)
            tk.Label(self, image=self._thumb_img, bg=BG).pack(pady=4)
        except: pass

        # 検出結果
        result_frame = tk.Frame(self, bg=CARD); result_frame.pack(fill="x", padx=20, pady=8)
        ri = tk.Frame(result_frame, bg=CARD); ri.pack(padx=16, pady=12)

        if self.detected is not None:
            tk.Label(ri, text="✔ 感度を検出しました", bg=CARD, fg=GREEN,
                     font=("NotoSansJP",11,"bold")).pack(anchor="w")
        else:
            tk.Label(ri, text="⚠ 感度を自動検出できませんでした（手動で入力してください）",
                     bg=CARD, fg=YELLOW, font=FB).pack(anchor="w")

        # 感度入力欄（検出値が初期値）
        row = tk.Frame(ri, bg=CARD); row.pack(anchor="w", pady=(8,4))
        tk.Label(row, text="感度:", bg=CARD, fg=TEXT, font=FB).pack(side="left")
        self._e_val = ttk.Entry(row, width=12, font=("NotoSansJP",14,"bold"))
        if self.detected is not None:
            self._e_val.insert(0, str(self.detected))
        self._e_val.pack(side="left", padx=8)

        # DPI
        row2 = tk.Frame(ri, bg=CARD); row2.pack(anchor="w", pady=4)
        tk.Label(row2, text="DPI:", bg=CARD, fg=TEXT, font=FB).pack(side="left")
        self._e_dpi = ttk.Entry(row2, width=8)
        self._e_dpi.insert(0, str(self.parent.cfg.get("dpi",800)))
        self._e_dpi.pack(side="left", padx=8)

        # メモ
        row3 = tk.Frame(ri, bg=CARD); row3.pack(anchor="w", pady=4)
        tk.Label(row3, text="メモ:", bg=CARD, fg=TEXT, font=FB).pack(side="left")
        self._e_note = ttk.Entry(row3, width=28)
        self._e_note.pack(side="left", padx=8)

        # OCRテキスト（折り畳み可能）
        if raw:
            tk.Label(ri, text=f"OCR生テキスト（一部）: {raw[:80].strip()!r}",
                     bg=CARD, fg=DIM, font=FS).pack(anchor="w", pady=(4,0))

        # ボタン
        btn = tk.Frame(self, bg=BG); btn.pack(pady=12)
        ttk.Button(btn, text="💾  この感度で保存", command=self._save).pack(side="left", padx=8)
        ttk.Button(btn, text="キャンセル", style="Sub.TButton",
                   command=self.destroy).pack(side="left")

    def _save(self):
        try:
            v = float(self._e_val.get())
        except:
            messagebox.showerror("エラー","感度の値が正しくありません"); return
        try: dpi = int(self._e_dpi.get())
        except: dpi = self.parent.cfg.get("dpi")
        note = self._e_note.get().strip()
        self.parent._save_sens_db(v, dpi, note, method="スクショ", ss_path=self.ss_path)
        messagebox.showinfo("保存",f"感度 {v} を記録しました！")
        self.parent.refresh_all()
        self.destroy()


# ─────────────────────────────────────────────
#  一括インポート確認ダイアログ（Henrikdev API）
# ─────────────────────────────────────────────
class BulkImportDialog(tk.Toplevel):
    def __init__(self, parent, sid, matches, since_str, until_str, name, api_key="", region="ap", my_puuid=""):
        super().__init__(parent)
        self.parent   = parent
        self.sid      = sid
        self.matches  = matches   # [(match_dict, match_dt), ...]
        self.name     = name
        self.api_key  = api_key
        self.region   = region
        self.my_puuid = my_puuid
        self.title("試合を一括インポート")
        self.configure(bg=BG)
        self.geometry("980x640")
        self.grab_set()
        self._check_vars = []
        self._build(since_str, until_str)

    def _build(self, since_str, until_str):
        tk.Label(self, text="📥  試合を一括インポート", bg=BG, fg=ACCENT2, font=FH).pack(pady=(14,2))
        tk.Label(self,
                 text=f"期間: {since_str}  〜  {until_str or '現在'}   |   {len(self.matches)} 試合が見つかりました",
                 bg=BG, fg=TEXT, font=FB).pack()
        tk.Label(self,
                 text="インポートしたい試合にチェックを入れて「選択した試合を登録」を押してください。ラウンド別スタッツも自動で保存されます。",
                 bg=BG, fg=DIM, font=FS).pack(pady=(2,6))

        btn_row = tk.Frame(self, bg=BG); btn_row.pack(anchor="w", padx=20)
        ttk.Button(btn_row, text="全て選択", style="Sub.TButton",
                   command=self._select_all).pack(side="left", padx=4)
        ttk.Button(btn_row, text="全て解除", style="Sub.TButton",
                   command=self._deselect_all).pack(side="left", padx=4)

        wrap = tk.Frame(self, bg=BG); wrap.pack(fill="both", expand=True, padx=20, pady=6)
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        sb     = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        self._inner = tk.Frame(canvas, bg=BG)
        self._inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self._inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ヘッダー
        hdr = tk.Frame(self._inner, bg=BG2); hdr.pack(fill="x", pady=(0,2))
        for txt, w in [("✔",30),("日時",150),("エージェント",105),("マップ",95),
                       ("K/D/A",80),("ACS",62),("DMG",70),("HS%",62),("勝敗",55)]:
            tk.Label(hdr, text=txt, bg=BG2, fg=ACCENT2,
                     font=("NotoSansJP",9,"bold"), width=w//7, anchor="center").pack(side="left",padx=2)

        self._check_vars = []
        for i, (match, match_dt) in enumerate(self.matches):
            var = tk.BooleanVar(value=True)
            self._check_vars.append(var)
            stats = self._extract(match)
            bg_c  = BG2 if i%2==0 else CARD
            row   = tk.Frame(self._inner, bg=bg_c); row.pack(fill="x", pady=1)

            tk.Checkbutton(row, variable=var, bg=bg_c,
                           activebackground=bg_c,
                           selectcolor="#ffffff",
                           fg="#00ff88",
                           activeforeground="#00ff88").pack(side="left", padx=4)

            date_disp = match_dt.strftime("%Y-%m-%d %H:%M")
            kda = f"{stats.get('kills','?')}/{stats.get('deaths','?')}/{stats.get('assists','?')}"
            acs = f"{stats['acs']:.0f}"   if stats.get('acs')  else "--"
            dmg = f"{stats['damage']:.0f}" if stats.get('damage') else "--"
            hs  = f"{stats['hs']:.1f}%"   if stats.get('hs')   else "--"
            won = "✔ 勝" if stats.get('won')==1 else "✘ 負" if stats.get('won')==0 else "--"

            for txt, w in [(date_disp,150),(stats.get('agent','--'),105),(stats.get('map','--'),95),
                           (kda,80),(acs,62),(dmg,70),(hs,62),(won,55)]:
                tk.Label(row, text=txt, bg=bg_c, fg=TEXT,
                         font=FS, width=w//7, anchor="center").pack(side="left", padx=2)

        existing = db().execute(
            "SELECT COUNT(*) FROM match_stats WHERE sensitivity_id=?", (self.sid,)).fetchone()[0]
        if existing:
            tk.Label(self, text=f"※ この感度にはすでに {existing} 試合が登録されています（重複に注意）",
                     bg=BG, fg=YELLOW, font=FS).pack()

        bf = tk.Frame(self, bg=BG); bf.pack(pady=10)
        ttk.Button(bf, text="💾 選択した試合を登録",
                   command=self._save).pack(side="left", padx=8)
        ttk.Button(bf, text="キャンセル", command=self.destroy,
                   style="Sub.TButton").pack(side="left")

    def _extract(self, match):
        """Henrikdev v4 のmatchデータからスタッツ抽出"""
        players = match.get("players", [])
        me = None
        # PUUID優先で照合（最も確実）
        if self.my_puuid:
            for p in players:
                if p.get("puuid","") == self.my_puuid:
                    me = p; break
        # PUUIDがなければ名前+タグで照合
        if not me:
            for p in players:
                if (p.get("name","").lower() == self.name.lower() and
                        p.get("tag","").lower() == self.parent.cfg.get("player_tag","").lower()):
                    me = p; break
        # それでもなければ名前のみで照合
        if not me:
            for p in players:
                if p.get("name","").lower() == self.name.lower():
                    me = p; break
        if not me and players:
            me = players[0]

        meta      = match.get("metadata", {})
        stats_raw = me.get("stats", {}) if me else {}
        team_id   = (me.get("team_id","") if me else "").lower()
        my_puuid  = self.my_puuid or (me.get("puuid","") if me else "")

        # ダメージはdealt
        dmg_field = stats_raw.get("damage", {})
        dmg = dmg_field.get("dealt") if isinstance(dmg_field, dict) else dmg_field

        # ACS = score / ラウンド数
        rounds     = match.get("rounds", [])
        num_rounds = len(rounds) if rounds else 1
        raw_score  = stats_raw.get("score", 0) or 0
        acs        = raw_score / num_rounds if num_rounds > 0 else None

        # HS%
        hs_h  = stats_raw.get("headshots", 0) or 0
        bd_h  = stats_raw.get("bodyshots",  0) or 0
        lg_h  = stats_raw.get("legshots",   0) or 0
        total = hs_h + bd_h + lg_h
        hs_pct = (hs_h / total * 100) if total > 0 else None

        result = {
            "kills":    stats_raw.get("kills"),
            "deaths":   stats_raw.get("deaths"),
            "assists":  stats_raw.get("assists"),
            "acs":      round(acs, 1) if acs else None,
            "damage":   dmg,
            "hs":       hs_pct,
            "agent":    me.get("agent",{}).get("name","") if me else "",
            "map":      meta.get("map",{}).get("name","") if isinstance(meta.get("map"),dict) else meta.get("map",""),
            "match_id": meta.get("match_id",""),
        }

        # 勝敗
        teams = match.get("teams", [])
        for t in teams:
            if t.get("team_id","").lower() == team_id:
                result["won"] = 1 if t.get("won") else 0
                break

        # ラウンド別スタッツ
        # 構造: rounds[].stats[] に player.puuid と stats{score,kills,headshots,bodyshots,legshots}
        # ダメージは damage_events[] の合計
        round_stats = []
        for rnd in rounds:
            rnd_num = rnd.get("id", len(round_stats))
            for pstat in rnd.get("stats", []):
                pid = pstat.get("player",{}).get("puuid","")
                if my_puuid and pid != my_puuid:
                    continue
                if not my_puuid and pstat.get("player",{}).get("name","").lower() != self.name.lower():
                    continue
                st      = pstat.get("stats", {})
                r_hs    = st.get("headshots", 0) or 0
                r_bd    = st.get("bodyshots",  0) or 0
                r_lg    = st.get("legshots",   0) or 0
                # ダメージはdamage_eventsを合計
                r_dmg   = sum(e.get("damage",0) for e in pstat.get("damage_events",[]))
                r_kills = st.get("kills", 0) or 0
                # デスは winning_team が相手チームかつ自分が死んだかどうかで判定
                round_stats.append({
                    "round_num":  rnd_num,
                    "kills":      r_kills,
                    "score":      st.get("score", 0) or 0,
                    "damage":     r_dmg,
                    "hs_hits":    r_hs,
                    "total_hits": r_hs + r_bd + r_lg,
                })
                break

        result["round_stats"] = round_stats
        return result


    def _select_all(self):
        for v in self._check_vars: v.set(True)

    def _deselect_all(self):
        for v in self._check_vars: v.set(False)

    def _save(self):
        selected = [(m,dt) for (m,dt),v in zip(self.matches,self._check_vars) if v.get()]
        if not selected:
            messagebox.showinfo("","1つ以上の試合を選択してください"); return

        c   = db()
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        cnt = 0
        for match, match_dt in selected:
            stats     = self._extract(match)
            date_disp = match_dt.strftime("%Y-%m-%d %H:%M")

            cur = c.execute("""INSERT INTO match_stats
                (sensitivity_id,match_id,match_date,agent,map,kills,deaths,assists,
                 acs,hs_percent,damage,won,tracker_url,note,added_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (self.sid,
                 stats.get("match_id") or None,
                 date_disp,
                 stats.get("agent")  or None,
                 stats.get("map")    or None,
                 stats.get("kills"),
                 stats.get("deaths"),
                 stats.get("assists"),
                 stats.get("acs"),
                 stats.get("hs"),
                 stats.get("damage"),
                 stats.get("won", 0),
                 None, "Henrikdev一括取得", now))
            mid = cur.lastrowid

            # ラウンドスタッツを保存
            for rs in stats.get("round_stats", []):
                c.execute("""INSERT INTO round_stats
                    (match_stats_id,round_num,kills,score,damage,hs_hits,total_hits)
                    VALUES(?,?,?,?,?,?,?)""",
                    (mid, rs["round_num"], rs["kills"], rs["score"],
                     rs["damage"], rs["hs_hits"], rs["total_hits"]))
            cnt += 1

        c.commit(); c.close()

        messagebox.showinfo("完了", f"{cnt} 試合を登録しました！（ラウンド別スタッツも保存済み）")
        self.parent._lbl_bulk_status.config(text=f"✔ {cnt} 試合を登録しました", fg=GREEN)
        self.parent.refresh_all()
        self.destroy()


# ─────────────────────────────────────────────
#  試合追加ダイアログ
# ─────────────────────────────────────────────
class MatchDialog(tk.Toplevel):
    AGENTS = ["Astra","Breach","Brimstone","Chamber","Clove","Cypher","Deadlock","Fade",
              "Gekko","Harbor","Iso","Jett","KAYO","Killjoy","Neon","Omen","Phoenix",
              "Raze","Reyna","Sage","Skye","Sova","Tejo","Viper","Vyse","Waylay","Yoru"]
    MAPS   = ["Abyss","Ascent","Bind","Breeze","Corrode","Fracture","Haven","Icebox",
              "Lotus","Pearl","Split","Sunset"]

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("試合を記録")
        self.configure(bg=BG)
        self.geometry("530x580")
        self.resizable(False, False)
        self.grab_set()
        self._build()
        self._load_sens()

    def _build(self):
        p = dict(padx=20, pady=5)
        tk.Label(self, text="試合を記録", bg=BG, fg=ACCENT, font=FH).pack(pady=(14,4))

        # 感度
        f0=tk.Frame(self,bg=BG);f0.pack(fill="x",**p)
        tk.Label(f0,text="感度:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._cb_sens=ttk.Combobox(f0,width=36,state="readonly");self._cb_sens.pack(side="left",padx=8)

        # 日時
        f1=tk.Frame(self,bg=BG);f1.pack(fill="x",**p)
        tk.Label(f1,text="日時:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._e_date=ttk.Entry(f1,width=20);self._e_date.insert(0,datetime.now().strftime("%Y-%m-%d %H:%M"))
        self._e_date.pack(side="left",padx=8)

        # エージェント・マップ
        f2=tk.Frame(self,bg=BG);f2.pack(fill="x",**p)
        tk.Label(f2,text="エージェント:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._cb_agent=ttk.Combobox(f2,values=self.AGENTS,width=14,state="readonly");self._cb_agent.pack(side="left",padx=8)
        tk.Label(f2,text="マップ:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._cb_map=ttk.Combobox(f2,values=self.MAPS,width=12,state="readonly");self._cb_map.pack(side="left",padx=8)

        # KDA
        f3=tk.Frame(self,bg=BG);f3.pack(fill="x",**p)
        for lbl,attr in [("Kills:","_e_k"),("Deaths:","_e_d"),("Assists:","_e_a")]:
            tk.Label(f3,text=lbl,bg=BG,fg=TEXT,font=FB).pack(side="left")
            e=ttk.Entry(f3,width=6);e.pack(side="left",padx=(2,12));setattr(self,attr,e)

        # ACS HS%
        f4=tk.Frame(self,bg=BG);f4.pack(fill="x",**p)
        tk.Label(f4,text="ACS:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._e_acs=ttk.Entry(f4,width=8);self._e_acs.pack(side="left",padx=(2,15))
        tk.Label(f4,text="HS%:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._e_hs=ttk.Entry(f4,width=8);self._e_hs.pack(side="left",padx=(2,15))

        # 勝敗
        f5=tk.Frame(self,bg=BG);f5.pack(fill="x",**p)
        tk.Label(f5,text="勝敗:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._won=tk.IntVar(value=1)
        ttk.Radiobutton(f5,text="✔ 勝ち",variable=self._won,value=1).pack(side="left",padx=8)
        ttk.Radiobutton(f5,text="✘ 負け",variable=self._won,value=0).pack(side="left")

        # tracker.gg URL
        f6=tk.Frame(self,bg=BG);f6.pack(fill="x",**p)
        tk.Label(f6,text="tracker.gg URL:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._e_url=ttk.Entry(f6,width=32);self._e_url.pack(side="left",padx=8)
        ttk.Button(f6,text="📋貼付&取得",command=self._paste_scrape,style="Sub.TButton").pack(side="left")

        self._lbl_sc=tk.Label(self,text="",bg=BG,fg=DIM,font=FS);self._lbl_sc.pack()

        # メモ
        f7=tk.Frame(self,bg=BG);f7.pack(fill="x",**p)
        tk.Label(f7,text="メモ:",bg=BG,fg=TEXT,font=FB).pack(side="left")
        self._e_note=ttk.Entry(f7,width=36);self._e_note.pack(side="left",padx=8)

        # ボタン
        br=tk.Frame(self,bg=BG);br.pack(pady=14)
        ttk.Button(br,text="💾 保存",command=self._save).pack(side="left",padx=8)
        ttk.Button(br,text="キャンセル",command=self.destroy,style="Sub.TButton").pack(side="left")

        tk.Label(self,text="tracker.gg の試合URLを貼り付けるとスタッツの自動取得を試みます（失敗時は手動入力）",
                 bg=BG,fg=DIM,font=FS).pack(pady=(0,8))

    def _load_sens(self):
        rows=db().execute("SELECT id,sensitivity,changed_at FROM sensitivity_log ORDER BY id DESC").fetchall()
        vals=[f"[{r[0]}] 感度 {r[1]}  ({r[2]})" for r in rows]
        self._cb_sens["values"]=vals
        if vals: self._cb_sens.current(0)

    def _paste_scrape(self):
        try:
            url=self.clipboard_get().strip()
            self._e_url.delete(0,"end"); self._e_url.insert(0,url)
        except: pass
        url=self._e_url.get().strip()
        if not url: return
        self._lbl_sc.config(text="読み込み中…",fg=YELLOW)
        threading.Thread(target=self._scrape,args=(url,),daemon=True).start()

    def _parse_profile_url(self, url):
        """
        tracker.gg のプロフィールURL から name と tag を抽出する
        例: https://tracker.gg/valorant/profile/riot/NAME%23TAG/overview
        → ("NAME", "TAG")
        """
        import urllib.parse
        m = re.search(r'/riot/([^/]+)', url)
        if not m: return None, None
        raw = urllib.parse.unquote(m.group(1))
        if '#' in raw:
            name, tag = raw.split('#', 1)
        elif '%23' in raw:
            name, tag = raw.split('%23', 1)
        else:
            return raw, None
        return name, tag

    def _scrape(self, url):
        import urllib.parse
        try:
            # ── プロフィールURLかどうか判定 ──────────────
            is_profile = '/profile/riot/' in url and '/overview' in url

            if is_profile:
                # tracker.gg 非公式API でマッチ履歴を取得
                name, tag = self._parse_profile_url(url)
                if not name or not tag:
                    raise ValueError("URLからプレイヤー名を取得できませんでした")

                encoded = urllib.parse.quote(f"{name}#{tag}", safe='')
                api_url = (
                    f"https://api.tracker.gg/api/v2/valorant/standard/matches/riot/"
                    f"{urllib.parse.quote(name, safe='')}%23{urllib.parse.quote(tag, safe='')}?"
                    f"type=competitive"
                )
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "application/json, text/plain, */*",
                    "Referer": "https://tracker.gg/",
                    "Origin": "https://tracker.gg",
                }
                req = urllib.request.Request(api_url, headers=headers)
                with urllib.request.urlopen(req, timeout=15) as r:
                    raw_json = r.read().decode("utf-8", "replace")

                data = json.loads(raw_json)
                matches = data.get("data", {}).get("matches", [])
                if not matches:
                    raise ValueError("試合データが見つかりませんでした（取得できる試合がない可能性）")

                # 最新1試合を取得
                match = matches[0]
                segments = match.get("segments", [])

                # 自分のセグメントを探す（nameで照合）
                my_seg = None
                for seg in segments:
                    meta = seg.get("metadata", {})
                    pname = meta.get("platformUserHandle","")
                    if name.lower() in pname.lower():
                        my_seg = seg; break
                if not my_seg and segments:
                    my_seg = segments[0]

                if not my_seg:
                    raise ValueError("プレイヤーデータが見つかりませんでした")

                stats = my_seg.get("stats", {})
                meta  = match.get("metadata", {})

                d = {}
                def gv(key):
                    s = stats.get(key, {})
                    return s.get("value") if s else None

                k = gv("kills");   d["kills"]   = int(k)   if k is not None else None
                de= gv("deaths");  d["deaths"]  = int(de)  if de is not None else None
                a = gv("assists"); d["assists"] = int(a)   if a is not None else None
                acs= gv("score");  d["acs"]     = float(acs) if acs is not None else None
                hs = gv("headshotsPercentage")
                if hs is None: hs = gv("headshots_percentage")
                d["hs"] = float(hs) if hs is not None else None

                # 勝敗
                outcome = my_seg.get("metadata", {}).get("result", "")
                if not outcome:
                    outcome = match.get("metadata", {}).get("result","")
                if "victory" in outcome.lower() or "win" in outcome.lower():
                    d["won"] = 1
                elif "defeat" in outcome.lower() or "loss" in outcome.lower():
                    d["won"] = 0

                # エージェント・マップ
                agent = my_seg.get("metadata", {}).get("agentName","") or \
                        my_seg.get("metadata", {}).get("agent","")
                map_n = meta.get("mapName","") or meta.get("map","")
                if agent: d["agent"] = agent
                if map_n: d["map"]   = map_n

                # Noneを除去
                d = {k:v for k,v in d.items() if v is not None}
                self.after(0, lambda: self._fill(d))

            else:
                # 個別試合URL（従来方式）
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "text/html,application/xhtml+xml",
                }
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=15) as r:
                    html = r.read().decode("utf-8","replace")
                d={}
                for k,p in [("acs",   r'"acs"[^}]*?"value"\s*:\s*([\d.]+)'),
                             ("hs",    r'"headshots_percentage"[^}]*?"value"\s*:\s*([\d.]+)'),
                             ("kills", r'"kills"[^}]*?"value"\s*:\s*(\d+)'),
                             ("deaths",r'"deaths"[^}]*?"value"\s*:\s*(\d+)'),
                             ("assists",r'"assists"[^}]*?"value"\s*:\s*(\d+)')]:
                    m2=re.search(p,html,re.I)
                    if m2: d[k]=(float if k in ("acs","hs") else int)(m2.group(1))
                if re.search(r'class="[^"]*victory[^"]*"',html,re.I): d["won"]=1
                elif re.search(r'class="[^"]*defeat[^"]*"',html,re.I): d["won"]=0
                self.after(0, lambda: self._fill(d))

        except Exception as e:
            self.after(0, lambda: self._lbl_sc.config(
                text=f"取得失敗: {e}", fg=ACCENT))

    def _fill(self, d):
        if not d:
            self._lbl_sc.config(text="データが見つかりませんでした（手動入力してください）", fg=YELLOW)
            return
        def se(w,v): w.delete(0,"end"); w.insert(0,str(v))
        if "kills"   in d: se(self._e_k,   d["kills"])
        if "deaths"  in d: se(self._e_d,   d["deaths"])
        if "assists" in d: se(self._e_a,   d["assists"])
        if "acs"     in d: se(self._e_acs, f"{d['acs']:.0f}")
        if "hs"      in d: se(self._e_hs,  f"{d['hs']:.1f}")
        if "won"     in d: self._won.set(d["won"])
        if "agent"   in d:
            for i, v in enumerate(self._cb_agent["values"]):
                if v.lower() == d["agent"].lower():
                    self._cb_agent.current(i); break
        if "map"     in d:
            for i, v in enumerate(self._cb_map["values"]):
                if v.lower() == d["map"].lower():
                    self._cb_map.current(i); break
        self._lbl_sc.config(text=f"✔ {len(d)}項目取得しました（内容を確認してください）", fg=GREEN)

    def _save(self):
        sel=self._cb_sens.get()
        if not sel: messagebox.showerror("エラー","感度を選択してください"); return
        sid=int(re.match(r"\[(\d+)\]",sel).group(1))
        def si(e):
            v=e.get().strip(); return int(v) if v else None
        def sf(e):
            v=e.get().strip(); return float(v) if v else None
        try: k,d,a,acs,hs=si(self._e_k),si(self._e_d),si(self._e_a),sf(self._e_acs),sf(self._e_hs)
        except: messagebox.showerror("エラー","数値が正しくありません"); return
        c=db()
        c.execute("""INSERT INTO match_stats
            (sensitivity_id,match_date,agent,map,kills,deaths,assists,
             acs,hs_percent,won,tracker_url,note,added_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (sid,self._e_date.get().strip(),
             self._cb_agent.get() or None,self._cb_map.get() or None,
             k,d,a,acs,hs,self._won.get(),
             self._e_url.get().strip() or None,
             self._e_note.get().strip() or None,
             datetime.now().strftime("%Y-%m-%d %H:%M")))
        c.commit(); c.close()
        messagebox.showinfo("保存","試合を記録しました！")
        self.parent.refresh_all(); self.destroy()


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()

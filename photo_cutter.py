"""
一卡通相片裁切工具 v7  ·  Fabrica 設計系統
──────────────────────────────────────────────────────────────
全新改版：
  • 相片表模式（Grid Sheet）— 自動偵測格線，逐格裁切，OCR 學號
  • 批次照片模式（Batch Folder）— 資料夾逐張照片，比對 Excel 學生名單，
    自動配對、標示缺件、匯出報表 → 支援 600+ 學生大量處理
  • iPhone 等級精準裁切對話框（平移 + 縮放 + 三分構圖格）
  • OCR 修正：字元框間距判斷，只取號碼不帶姓名
  • 裁切修正：相片表格子已由格線偵測清楚分割
  • 資安強化：路徑消毒、ZIP 結構安全、記憶體內處理無落地暫存
  • Tier 授權系統：Free / Scholar / Pro / Enterprise
  • GitHub 自動更新通知

相依套件：Pillow, numpy, openpyxl（選用）, pytesseract（選用+Tesseract引擎）
           rembg + onnxruntime（選用，首次執行自動下載 ~170MB 模型，之後完全離線）
           PyJWT + cryptography（授權驗證）
"""
VERSION = '1.0.0'
GITHUB_REPO = 'FadedCantCode/Fabrica-CROP'

import os, io, re, csv, json, zipfile, threading, datetime, urllib.request
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont
import numpy as np
from PIL import Image, ImageTk, ImageOps, ImageDraw

try:
    import openpyxl; EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    import pytesseract
    # 自動偵測 Tesseract 安裝位置（Windows / macOS / Linux 常見路徑）
    _TESS_CANDIDATES = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",       # Windows 預設
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe", # Windows 32-bit
        r"C:\Users\Public\Tesseract-OCR\tesseract.exe",        # Windows 另一種
        "/usr/local/bin/tesseract",    # macOS Homebrew (Intel)
        "/opt/homebrew/bin/tesseract", # macOS Homebrew (Apple Silicon)
        "/usr/bin/tesseract",          # Linux apt
    ]
    for _p in _TESS_CANDIDATES:
        if os.path.isfile(_p):
            pytesseract.pytesseract.tesseract_cmd = _p
            break
    pytesseract.get_tesseract_version()
    OCR_OK = True
except Exception:
    OCR_OK = False

try:
    import rembg as _rembg_mod; REMBG_OK = True
except ImportError:
    REMBG_OK = False

_rembg_session = None

# ── 授權系統 (RSA JWT) ──────────────────────────────────────
try:
    import jwt as _jwt
    from cryptography.hazmat.primitives.serialization import load_pem_public_key as _load_pub
    _JWT_OK = True
except ImportError:
    _JWT_OK = False

# 公鑰（只能驗證，不能生成金鑰；私鑰只在 dev_console.py）
_PUBLIC_KEY_PEM = (
    "-----BEGIN PUBLIC KEY-----\n"
    "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA9lFjvPnHvCxS6kq2vfjX\n"
    "fY6yZ69iYn66BgtVBOezOPTTkG5VVEBZoVCrqlue/dDLV+RAtIaMDgBNP7fYwqD9\n"
    "nbA+RslBOaJOMDJmQnfsCf0QVfD4i4TZqipoD6kYX3STxXgYDlf8hxYoJNbnNj7Y\n"
    "0vi17BbFbwg4EVlZ/owmcw9acLIEYwHFwFtpyafPS6VNcNfX8pv3HbVlGAdRbTmi\n"
    "BS6jkTwL0AENS1MhJ1K5uc5wdy4+UKq4ZeN9ly7HyYF9PJvYeFdBM5EmjZQ8kP0J\n"
    "QJhL4KJBDiFFh2YsS/xa9xPMKMMSVLnvi038spoVUFUaLqh6CanZ2nPP5t+kuvm/\n"
    "hQIDAQAB\n"
    "-----END PUBLIC KEY-----\n"
)
_LICENSE_PATH = os.path.join(os.path.expanduser('~'), '.fabrica', 'photo_cutter.license')

# Tier 功能矩陣
TIER_LIMITS = {
    'free':       {'max_cells': 30,  'ocr': False, 'matting': False, 'batch': False},
    'scholar':    {'max_cells': 150, 'ocr': True,  'matting': False, 'batch': True},
    'pro':        {'max_cells': 0,   'ocr': True,  'matting': True,  'batch': True},
    'enterprise': {'max_cells': 0,   'ocr': True,  'matting': True,  'batch': True},
}
TIER_LABEL = {
    'free':       ('FREE',       '#8C8474'),
    'scholar':    ('SCHOLAR',    '#1A7FBF'),
    'pro':        ('PRO',        '#E8552F'),
    'enterprise': ('ENTERPRISE', '#1FA85B'),
}

class LicenseManager:
    """啟動時載入 license，提供 tier / 功能查詢。"""
    def __init__(self):
        self.payload = None
        self.tier = 'free'
        self.org = ''
        self._load()

    def _load(self):
        if not _JWT_OK or not os.path.exists(_LICENSE_PATH):
            return
        try:
            token = open(_LICENSE_PATH, 'r', encoding='utf-8').read().strip()
            pub = _load_pub(_PUBLIC_KEY_PEM.encode())
            self.payload = _jwt.decode(token, pub, algorithms=['RS256'],
                                        issuer='fabrica.studio')
            self.tier = self.payload.get('tier', 'free')
            self.org  = self.payload.get('org', '')
        except _jwt.ExpiredSignatureError:
            self.payload = {'_error': '授權金鑰已到期，請聯繫 Fabrica 更新。'}
        except Exception:
            self.payload = {'_error': '授權金鑰無效。'}

    def activate(self, token: str) -> tuple[bool, str]:
        """輸入金鑰字串，驗證並儲存。"""
        if not _JWT_OK:
            return False, '缺少套件：pip install PyJWT cryptography'
        try:
            pub = _load_pub(_PUBLIC_KEY_PEM.encode())
            payload = _jwt.decode(token.strip(), pub, algorithms=['RS256'],
                                   issuer='fabrica.studio')
            os.makedirs(os.path.dirname(_LICENSE_PATH), exist_ok=True)
            with open(_LICENSE_PATH, 'w', encoding='utf-8') as f:
                f.write(token.strip())
            self.payload = payload
            self.tier = payload.get('tier', 'free')
            self.org  = payload.get('org', '')
            return True, f"✓ 啟動成功！{TIER_LABEL[self.tier][0]} 方案 · {self.org}"
        except _jwt.ExpiredSignatureError:
            return False, '❌ 金鑰已到期，請聯繫 Fabrica 更新。'
        except Exception:
            return False, '❌ 金鑰無效，請確認複製正確。'

    def allowed(self, feature: str) -> bool:
        return TIER_LIMITS.get(self.tier, TIER_LIMITS['free']).get(feature, False)

    def max_cells(self) -> int:
        return TIER_LIMITS.get(self.tier, TIER_LIMITS['free'])['max_cells']

    def error(self) -> str | None:
        if self.payload and '_error' in self.payload:
            return self.payload['_error']
        return None

LICENSE = LicenseManager()


# ── GitHub 版本更新通知 ────────────────────────────────────
def check_github_update(repo: str, current: str, callback):
    """背景執行緒：查 GitHub releases，有新版就呼叫 callback(latest_version)。"""
    def _worker():
        try:
            url = f'https://api.github.com/repos/{repo}/releases/latest'
            req = urllib.request.Request(url, headers={'User-Agent': 'FabricaPhotoApp'})
            with urllib.request.urlopen(req, timeout=5) as r:
                data = json.loads(r.read())
            latest = data.get('tag_name','').lstrip('v')
            if latest and latest != current:
                callback(latest)
        except Exception:
            pass   # 靜默失敗，不影響主程式
    threading.Thread(target=_worker, daemon=True).start()

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

# ── 安全設定 ────────────────────────────────────────────────
# 防止 Decompression Bomb：限制解碼圖像最大像素數（50MP 已遠超任何學生證用途）
Image.MAX_IMAGE_PIXELS = 50_000_000

# ── Fabrica 設計 Token ──────────────────────────────────────
C = {
    "bg":        "#F5F1E6",
    "surface":   "#FBF8F0",
    "surface2":  "#F1EAD8",
    "accent":    "#E8552F",
    "accent_dk": "#C0411F",
    "ink":       "#18140F",
    "muted":     "#8C8474",
    "line":      "#E4DCC9",
    "ok":        "#1FA85B",
    "warn":      "#D8A000",
    "err":       "#D8452F",
    "header":    "#18140F",
    "hfg":       "#FBF8F0",
}
FONT = "Microsoft JhengHei"


# ════════════════════════════════════════════════════════════
#  安全工具
# ════════════════════════════════════════════════════════════

_UNSAFE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')

def safe_filename(s: str, max_len: int = 80) -> str:
    """消毒檔名：移除危險字元，防止路徑穿越。"""
    s = _UNSAFE.sub('_', str(s).strip())
    s = re.sub(r'\.\.', '_', s)
    return s[:max_len] if s else 'unnamed'


# ════════════════════════════════════════════════════════════
#  影像核心
# ════════════════════════════════════════════════════════════

def _smooth(p, w):
    return np.convolve(p, np.ones(w)/w, mode='same') if w > 1 else p

def _bands(prof, min_fill, min_size, sw):
    p = _smooth(prof, sw); active = p > min_fill
    res, i, n = [], 0, len(p)
    while i < n:
        if active[i]:
            j = i
            while j < n and active[j]: j += 1
            if j - i >= min_size: res.append((i, j))
            i = j
        else:
            i += 1
    return res

def detect_grid(img, force_cols=None, force_rows=None, dark_thr=232):
    """自動偵測欄/列。回傳 (col_bands, photo_row_bands)。"""
    g = np.asarray(img.convert('RGB')).mean(axis=2)
    H, W = g.shape
    col_b = _bands((g < dark_thr).mean(axis=0), 0.04, int(W*0.04), max(3, int(W*0.004)))
    row_b = _bands((g < dark_thr).mean(axis=1), 0.03, int(H*0.02), max(3, int(H*0.004)))
    if row_b:
        hmax = max(e-s for s,e in row_b)
        photo_rows = [b for b in row_b if (b[1]-b[0]) >= 0.5*hmax]
    else:
        photo_rows = []
    def adj(bands, target):
        if target is None or not bands: return bands
        if len(bands) > target:
            bands = sorted(bands, key=lambda b: b[1]-b[0], reverse=True)[:target]
        return sorted(bands, key=lambda b: b[0])
    return adj(col_b, force_cols), adj(photo_rows, force_rows)

def _fill_ratio(arr, thr=235):
    return float((arr.mean(axis=2) < thr).mean())

# ── 邊框清理（僅用於批次個別照片，相片表格子不需要）──

def _is_border_line(vec):
    """
    判斷一維像素向量是否為「跨全寬的格線」。
    格線：暗點分布在至少 65% 的欄/列寬度，且整體偏暗或稀疏虛線。
    頭髮、臉、衣服：暗點集中在局部，無法通過 spread 檢測。
    """
    n = len(vec)
    if n < 6: return False
    g = vec.astype(float)
    dark = g < 140
    seg = max(1, n // 8)
    spread = sum(1 for i in range(0, n, seg) if dark[i:i+seg].any()) / max(1, n // seg)

    # ── Case 1：實心暗線（跨全寬）──
    if g.mean() < 158 and spread >= 0.65: return True

    # ── Case 2：實體相片紙邊（亮底稀疏暗點）──
    # 實測特徵：mean≈226-233（不是純白但也不暗）、dark%≈1.5-21%、bright%≈78-80%
    # 與頭髮的區別：頭髮 mean<180，不符合 mean>212 條件。
    # 與純白紙的區別：純白 mean>240，不符合 mean<240 條件。
    # 不要求 spread（相片貼斜時 spread 只有 12-25%）。
    if 212 < g.mean() < 240 and 0.01 < dark.mean() < 0.45 and (g > 210).mean() > 0.55:
        return True

    return False

def strip_photo_borders(arr, inset=2, max_frac=0.08):
    """
    去掉實體相片邊框：掃描紙白邊 + 虛線格線。

    演算法：「跳過白區，掃到第一個非邊框就停」
    ──────────────────────────────────────────
    WHITE 門檻：mean > 230 且 dark% < 1.5%
      → 涵蓋純白掃描紙（dark=0%）以及掃描陰影漸層（dark=0%，mean=233-246）
    跳過 WHITE 後：符合 _is_border_line → 切；否則立刻停止。

    同時正確處理兩情境：
      RIGHT 虛線：skip(dark=0%) → skip(dark=0%) → BORDER(dark=1.5%)→cut → stop
      TOP 頭髮  ：skip(dark=0%) → content(dark=3.1%,mean=248>240) → 立即停 cut=0 ✓
    """
    h, w = arr.shape[:2]
    t, b, l, r = inset, h-inset, inset, w-inset
    if b <= t or r <= l: return arr
    a = arr[t:b, l:r]
    g = a.mean(axis=2)
    sh, sw = g.shape
    mr = max(2, int(sh * max_frac))
    mc = max(2, int(sw * max_frac))

    def seq_cut(get_vec, n_check):
        cut = 0
        for k in range(n_check):
            v = get_vec(k).astype(float)
            mean_k = v.mean()
            dark_k = (v < 140).mean()
            if mean_k > 230 and dark_k < 0.015:   # 白區 → 跳過
                continue
            if _is_border_line(v):                 # 邊框 → 切
                cut = k + 1
            else:                                  # 內容(頭髮/臉) → 停
                break
        return cut

    tt = seq_cut(lambda k: g[k],        mr)
    bb = seq_cut(lambda k: g[sh-1-k],   mr)
    ll = seq_cut(lambda k: g[:, k],     mc)
    rr = seq_cut(lambda k: g[:, sw-1-k],mc)

    result = a[tt:sh-bb if bb else sh, ll:sw-rr if rr else sw]
    return result if result.size > 0 else arr


def render_cell(cell, out_size=(420, 530)):
    """依 rotation 旋轉再縮放至輸出尺寸。"""
    im = cell['_orig']
    r = cell['rotation'] % 360
    if r == 90:  im = im.transpose(Image.ROTATE_270)
    elif r == 180: im = im.transpose(Image.ROTATE_180)
    elif r == 270: im = im.transpose(Image.ROTATE_90)
    return im.resize(out_size, Image.LANCZOS)


# ════════════════════════════════════════════════════════════
#  去背白底（AI 摳圖）
# ════════════════════════════════════════════════════════════

def _get_rembg_session():
    """懶載入 rembg session（首次呼叫下載模型 ~170MB，之後完全離線）。"""
    global _rembg_session
    if _rembg_session is None:
        # 優先用人像專用模型，沒有就用通用 u2net
        for model in ('u2net_human_seg', 'u2net'):
            try:
                _rembg_session = _rembg_mod.new_session(model)
                break
            except Exception:
                continue
    return _rembg_session

def matting_to_white(src_img, out_size=(420, 530)):
    """
    AI 摳圖：把人像從背景分離，置中貼在純白底上。
    解決：實體相片虛線框、雜色背景、掃描陰影。
    需要 rembg + onnxruntime（pip install rembg onnxruntime）。
    首次執行自動下載模型（~170MB），之後完全在本機離線執行。
    回傳 420×530 純白底 RGB 圖片。
    """
    if not REMBG_OK:
        raise RuntimeError("請先安裝：pip install rembg onnxruntime")
    sess = _get_rembg_session()
    if sess is None:
        raise RuntimeError("rembg 模型載入失敗，請確認網路連線以下載模型")

    rgba = _rembg_mod.remove(src_img, session=sess)          # RGBA，背景透明
    alpha = np.array(rgba)[:, :, 3]

    # 找人像邊框
    rows = np.any(alpha > 10, axis=1)
    cols = np.any(alpha > 10, axis=0)
    if not rows.any():
        return src_img.resize(out_size, Image.LANCZOS)       # fallback

    rmin, rmax = np.where(rows)[0][[0, -1]]
    cmin, cmax = np.where(cols)[0][[0, -1]]

    # 稍微加 margin，避免頭頂被裁到
    pad_v = max(6, int((rmax - rmin) * 0.04))
    pad_h = max(6, int((cmax - cmin) * 0.04))
    h_src, w_src = alpha.shape
    rmin = max(0, rmin - pad_v);  rmax = min(h_src, rmax + pad_v)
    cmin = max(0, cmin - pad_h);  cmax = min(w_src, cmax + pad_h)

    subject = rgba.crop((cmin, rmin, cmax, rmax))
    sw, sh = subject.size
    ow, oh = out_size

    # 縮放：寬度最多 90%，高度最多 92%（留頭頂空間）
    scale = min(ow * 0.90 / sw, oh * 0.92 / sh)
    nw, nh = max(1, int(sw * scale)), max(1, int(sh * scale))
    subject = subject.resize((nw, nh), Image.LANCZOS)

    # 貼在白底，水平置中，垂直稍偏上（頭頂留約 4%）
    white = Image.new('RGB', out_size, (255, 255, 255))
    x = (ow - nw) // 2
    y = max(int(oh * 0.04), (oh - nh) // 2)
    white.paste(subject.convert('RGB'), (x, y), mask=subject.split()[3])
    return white

def extract_cells(img, col_bands, photo_rows, empty_thr=0.02, out_size=(420, 530)):
    """
    相片表格子擷取。
    格線偵測已把格子清楚分割，只做最小固定 2px 內縮即可，
    絕對不用紙張色偵測（會誤切白衣）或內容比例分析（會誤切頭頂）。
    """
    arr = np.asarray(img.convert('RGB'))
    cells = []
    for ri, (y1, y2) in enumerate(photo_rows):
        for ci, (x1, x2) in enumerate(col_bands):
            raw = arr[y1:y2, x1:x2]
            if raw.size == 0 or _fill_ratio(raw) < empty_thr: continue
            # strip_photo_borders：先 2px 固定內縮，再偵測並削掉實體相片虛線邊框
            photo = strip_photo_borders(raw)
            orig = Image.fromarray(photo)
            rotation = 90 if (orig.width > orig.height * 1.5) else 0
            cd = {'row': ri, 'col': ci, '_orig': orig, 'rotation': rotation,
                  'id': '', '_cell_box': (x1, y1, x2, y2), 'dup': False}
            cd['image'] = render_cell(cd, out_size)
            cells.append(cd)
    return cells


# ════════════════════════════════════════════════════════════
#  OCR 學號
# ════════════════════════════════════════════════════════════

def _digit_cluster(bw, cfg, psm):
    """字元框間距法：只取最左側連續號碼段，遇大空白（姓名前）停止。"""
    try:
        boxes = pytesseract.image_to_boxes(bw, config=f'--psm {psm} {cfg}')
    except: return ''
    chars = []
    for line in boxes.splitlines():
        p = line.split()
        if len(p) >= 5 and p[0].isdigit():
            chars.append((int(p[1]), int(p[3]), p[0]))
    if not chars: return ''
    chars.sort(key=lambda c: c[0])
    gaps = sorted(g for g in (chars[i][0]-chars[i-1][1] for i in range(1,len(chars))) if g > 0)
    medg = gaps[len(gaps)//2] if gaps else 0
    out, prev = chars[0][2], chars[0][1]
    for xa, xb, ch in chars[1:]:
        if len(out) >= 6 and (xa-prev) > max(medg*1.6, 4): break
        out += ch; prev = xb
    return out if 5 <= len(out) <= 9 else ''

def ocr_id_below(img, box, next_row_top=None):
    """擷取照片下方學號文字並辨識，多門檻投票。"""
    if not OCR_OK: return ''
    x1, y1, x2, y2 = box
    ph = y2 - y1
    top = y2 + int(ph*0.02)
    lim = next_row_top - int(ph*0.05) if next_row_top else y2+int(ph*0.45)
    bot = min(y2+int(ph*0.40), lim, img.height)
    if bot - top < 8: return ''
    strip = img.crop((x1, top, x2, bot)).convert('L')
    strip = strip.resize((strip.width*3, strip.height*3), Image.LANCZOS)
    strip = ImageOps.autocontrast(strip)
    cfg = '-c tessedit_char_whitelist=0123456789'
    from collections import Counter
    votes = Counter()
    for thr in (135, 160):
        bw = strip.point(lambda v: 0 if v < thr else 255)
        for psm in ('7','6'):
            n = _digit_cluster(bw, cfg, psm)
            if n: votes[n] += 1
    if not votes: return ''
    ranked = sorted(votes.items(), key=lambda kv: (kv[1], 6<=len(kv[0])<=8, -len(kv[0])), reverse=True)
    return ranked[0][0]

def id_from_filename(fn: str) -> str:
    """嘗試從檔名提取 6-8 碼學號。"""
    stem = os.path.splitext(os.path.basename(fn))[0]
    m = re.findall(r'\b\d{6,8}\b', stem)
    return m[0] if m else ''


# ════════════════════════════════════════════════════════════
#  資料層
# ════════════════════════════════════════════════════════════

def load_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return {}, []
    hdrs = [str(c) if c is not None else f'欄{i+1}' for i,c in enumerate(rows[0])]
    data = {h: [] for h in hdrs}
    for row in rows[1:]:
        for h,v in zip(hdrs, row):
            data[h].append(str(v).strip() if v is not None else '')
    return data, hdrs

def guess_id_col(data, hdrs):
    kw = ('學號','座號','卡號','id','no','number','編號')
    for h in hdrs:
        if any(k in h.lower() for k in kw): return h
    best, bs = None, 0.0
    for h in hdrs:
        vals = [v for v in data[h] if v]
        if not vals: continue
        s = sum(1 for v in vals if re.fullmatch(r'\d{5,9}', v)) / len(vals)
        if s > bs: best, bs = h, s
    return best if bs >= 0.6 else None

def scan_folder(folder):
    """掃描資料夾，回傳 [{path, filename, auto_id}]。
    安全：驗證每個路徑的 realpath 仍在 folder 內，防止 symlink 逃逸。"""
    real_root = os.path.realpath(folder)
    try:
        entries = [e for e in os.scandir(folder)
                   if e.is_file(follow_symlinks=False)   # 不追蹤 symlink
                   and os.path.splitext(e.name)[1].lower() in IMAGE_EXTS
                   and os.path.realpath(e.path).startswith(real_root)]  # realpath 驗證
        return sorted([
            {'path': e.path, 'filename': e.name, 'auto_id': id_from_filename(e.name)}
            for e in entries
        ], key=lambda x: x['filename'].lower())
    except PermissionError as exc:
        raise RuntimeError(f'無法存取資料夾：{exc}')


# ════════════════════════════════════════════════════════════
#  iPhone 等級精準裁切對話框
# ════════════════════════════════════════════════════════════

class CropperDialog(tk.Toplevel):
    """
    固定比例裁切框（2:3），照片在框下平移/縮放。
    • 拖曳 → 平移照片
    • 滾輪 → 縮放照片
    • 鏽橘角標 + 三分構圖格線
    • 照片永遠完整覆蓋裁切框（不露出黑底）
    """
    OUT_W, OUT_H = 420, 530
    FRM_H = 468
    FRM_W = int(FRM_H * OUT_W / OUT_H)   # ≈ 371
    CV_W, CV_H = 660, 580

    def __init__(self, parent, cell):
        super().__init__(parent)
        self.title('精準裁切'); self.resizable(False, False)
        self.configure(bg=C['bg'])
        self.cell = cell
        self.src = cell['_orig'].copy()
        self.rotation = cell['rotation']
        self.zoom = 1.0; self.pan_x = 0.0; self.pan_y = 0.0
        self._drag = None; self._tk_img = None
        self._build(); self._fit()
        self.transient(parent); self.grab_set()

    def _get_src(self):
        r = self.rotation % 360
        if r == 90:   return self.src.transpose(Image.ROTATE_270)
        if r == 180:  return self.src.transpose(Image.ROTATE_180)
        if r == 270:  return self.src.transpose(Image.ROTATE_90)
        return self.src

    def _fit(self):
        """讓照片剛好填滿裁切框（cover 模式）。"""
        ph = self._get_src(); pw, pheight = ph.size
        self.zoom = max(self.FRM_W/pw, self.FRM_H/pheight)
        self.pan_x = self.pan_y = 0.0
        self._redraw()

    def _build(self):
        self.cv = tk.Canvas(self, width=self.CV_W, height=self.CV_H,
                            bg='#1a1a1a', highlightthickness=0)
        self.cv.pack()
        self.cv.bind('<ButtonPress-1>', lambda e: setattr(self,'_drag',(e.x,e.y)))
        self.cv.bind('<B1-Motion>', self._on_drag)
        self.cv.bind('<MouseWheel>', self._on_wheel)
        self.cv.bind('<Button-4>', lambda e: self._zoom(1.1))
        self.cv.bind('<Button-5>', lambda e: self._zoom(0.9))

        bar = tk.Frame(self, bg=C['surface'], highlightbackground=C['line'], highlightthickness=1)
        bar.pack(fill='x')
        def btn(p, t, c, k='ink'):
            bg={'primary':C['accent'],'ink':C['ink'],'ghost':C['surface2']}[k]
            fg=C['ink'] if k=='ghost' else 'white'
            return tk.Button(p, text=t, command=c, bg=bg, fg=fg, relief='flat',
                             font=(FONT,9,'bold'), cursor='hand2', bd=0)
        btn(bar,'↺', lambda: self._rot(-90)).pack(side='left',padx=8,pady=6,ipadx=6)
        btn(bar,'↻', lambda: self._rot(90)).pack(side='left',padx=2,pady=6,ipadx=6)
        btn(bar,'↔ 翻轉', self._flip).pack(side='left',padx=8,pady=6,ipadx=6)
        btn(bar,'⊡ 重置', self._fit,'ghost').pack(side='left',padx=2,pady=6,ipadx=6)
        tk.Label(bar,text='滾輪縮放・拖曳平移',bg=C['surface'],fg=C['muted'],font=(FONT,8)).pack(side='right',padx=10)
        btn(bar,'套用裁切',self._apply,'primary').pack(side='right',padx=10,pady=6,ipadx=14,ipady=3)

    def _on_drag(self, e):
        if self._drag:
            dx, dy = e.x-self._drag[0], e.y-self._drag[1]
            self.pan_x += dx; self.pan_y += dy; self._drag = (e.x, e.y)
            self._clamp(); self._redraw()

    def _on_wheel(self, e):
        self._zoom(1.1 if e.delta > 0 else 0.9)

    def _zoom(self, f):
        old = self.zoom
        self.zoom = max(0.3, min(10.0, self.zoom*f))
        s = self.zoom/old; self.pan_x *= s; self.pan_y *= s
        self._clamp(); self._redraw()

    def _rot(self, d):
        self.rotation = (self.rotation+d)%360; self._fit()

    def _flip(self):
        self.src = self.src.transpose(Image.FLIP_LEFT_RIGHT); self._redraw()

    def _clamp(self):
        """確保裁切框始終在照片內部（不露黑底）。"""
        ph = self._get_src(); pw, pheight = ph.size
        dw, dh = pw*self.zoom, pheight*self.zoom
        fw, fh = self.FRM_W, self.FRM_H
        cx0 = self.CV_W/2 - dw/2; cy0 = self.CV_H/2 - dh/2
        fx0 = (self.CV_W-fw)/2;   fy0 = (self.CV_H-fh)/2
        # pan limits so frame stays inside photo
        max_px = fx0 - cx0;       min_px = (fx0+fw) - (cx0+dw)
        max_py = fy0 - cy0;       min_py = (fy0+fh) - (cy0+dh)
        self.pan_x = max(min_px, min(max_px, self.pan_x))
        self.pan_y = max(min_py, min(max_py, self.pan_y))

    def _redraw(self):
        ph = self._get_src(); pw, pheight = ph.size
        rw, rh = max(1,int(pw*self.zoom)), max(1,int(pheight*self.zoom))
        rendered = ph.resize((rw, rh), Image.LANCZOS)
        # 合成畫布
        canvas = Image.new('RGB', (self.CV_W, self.CV_H), (26,26,26))
        px = int(self.CV_W/2 - rw/2 + self.pan_x)
        py = int(self.CV_H/2 - rh/2 + self.pan_y)
        canvas.paste(rendered, (px, py))
        # 框外暗化
        fw, fh = self.FRM_W, self.FRM_H
        fx0 = int((self.CV_W-fw)/2); fy0 = int((self.CV_H-fh)/2)
        fx1, fy1 = fx0+fw, fy0+fh
        ov = Image.new('RGBA', (self.CV_W, self.CV_H), (0,0,0,110))
        ImageDraw.Draw(ov).rectangle([fx0,fy0,fx1,fy1], fill=(0,0,0,0))
        canvas = canvas.convert('RGBA'); canvas.alpha_composite(ov)
        canvas = canvas.convert('RGB')
        d = ImageDraw.Draw(canvas)
        # 裁切框
        d.rectangle([fx0,fy0,fx1,fy1], outline='white', width=2)
        # 三分格線
        for i in (1,2):
            d.line([(fx0+fw*i//3,fy0),(fx0+fw*i//3,fy1)], fill=(255,255,255,60), width=1)
            d.line([(fx0,fy0+fh*i//3),(fx1,fy0+fh*i//3)], fill=(255,255,255,60), width=1)
        # 鏽橘角標
        CL, CW = 18, 3
        for cx,cy,dx,dy in [(fx0,fy0,1,1),(fx1,fy0,-1,1),(fx0,fy1,1,-1),(fx1,fy1,-1,-1)]:
            d.line([(cx,cy),(cx+dx*CL,cy)], fill=C['accent'], width=CW)
            d.line([(cx,cy),(cx,cy+dy*CL)], fill=C['accent'], width=CW)
        self._tk_img = ImageTk.PhotoImage(canvas)
        self.cv.delete('all')
        self.cv.create_image(0,0,anchor='nw',image=self._tk_img)

    def _apply(self):
        ph = self._get_src(); pw, pheight = ph.size
        fw, fh = self.FRM_W, self.FRM_H
        fx0 = (self.CV_W-fw)/2; fy0 = (self.CV_H-fh)/2
        px0 = self.CV_W/2 - pw*self.zoom/2 + self.pan_x
        py0 = self.CV_H/2 - pheight*self.zoom/2 + self.pan_y
        cx1 = max(0.0, (fx0-px0)/self.zoom)
        cy1 = max(0.0, (fy0-py0)/self.zoom)
        cx2 = min(float(pw), cx1+fw/self.zoom)
        cy2 = min(float(pheight), cy1+fh/self.zoom)
        if cx2-cx1 < 10 or cy2-cy1 < 10:
            messagebox.showwarning('裁切', '裁切區域太小，請縮小後重試'); return
        cropped = ph.crop((cx1,cy1,cx2,cy2))
        self.cell['_orig'] = cropped
        self.cell['rotation'] = 0
        self.cell['image'] = render_cell(self.cell)
        self.destroy()


# ════════════════════════════════════════════════════════════
#  進度對話框
# ════════════════════════════════════════════════════════════

class ProgressDlg(tk.Toplevel):
    def __init__(self, parent, title, maximum):
        super().__init__(parent)
        self.title(title); self.geometry('360x110')
        self.resizable(False,False); self.configure(bg=C['surface'])
        self.transient(parent); self.grab_set()
        self.lbl = tk.Label(self, text=title, bg=C['surface'], fg=C['ink'], font=(FONT,10))
        self.lbl.pack(pady=(16,6))
        self.bar = ttk.Progressbar(self, length=300, maximum=maximum)
        self.bar.pack(); self.update()
    def step(self, val, text=None):
        self.bar['value'] = val
        if text: self.lbl.config(text=text)
        self.update()
    def close(self): self.destroy()


class Tooltip:
    """任何 widget 都可以加 hover 提示氣泡，1 秒後出現。"""
    def __init__(self, widget, text):
        self.widget = widget; self.text = text; self._job = None; self.tw = None
        widget.bind('<Enter>', self._enter); widget.bind('<Leave>', self._leave)
    def _enter(self, _e):
        self._job = self.widget.after(900, self._show)
    def _leave(self, _e):
        if self._job: self.widget.after_cancel(self._job); self._job = None
        if self.tw: self.tw.destroy(); self.tw = None
    def _show(self):
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tw = tk.Toplevel(self.widget); self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f'+{x}+{y}')
        tk.Label(self.tw, text=self.text, bg=C['ink'], fg=C['hfg'],
                 font=(FONT, 9), padx=10, pady=5, justify='left',
                 wraplength=280).pack()


class ZoomDialog(tk.Toplevel):
    """點擊預覽照片後彈出的大圖視窗，按 Esc 或空白鍵關閉。"""
    def __init__(self, parent, cell):
        super().__init__(parent)
        self.title('放大檢視（按 Esc 關閉）')
        self.configure(bg='white'); self.resizable(False, False)
        img = cell['image']
        w, h = img.size
        # 最大顯示 700×880（大螢幕），小螢幕也不超出
        scale = min(700/w, 880/h, 1.6)
        disp = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
        tk_img = ImageTk.PhotoImage(disp)
        lbl = tk.Label(self, image=tk_img, bg='white', cursor='hand2')
        lbl.image = tk_img; lbl.pack()
        info = cell.get('id','') + ('  ' + cell.get('excel_name','') if cell.get('excel_name') else '')
        if info.strip():
            tk.Label(self, text=info, bg='white', fg=C['ink'],
                     font=(FONT, 14, 'bold')).pack(pady=6)
        self.bind('<Escape>', lambda e: self.destroy())
        self.bind('<space>',  lambda e: self.destroy())
        lbl.bind('<Button-1>', lambda e: self.destroy())
        self.transient(parent); self.grab_set()


# ════════════════════════════════════════════════════════════
#  共用 UI 工具
# ════════════════════════════════════════════════════════════

def fab_btn(parent, text, cmd, kind='primary', **kw):
    bg = {'primary':C['accent'],'ink':C['ink'],'ghost':C['surface2'],'warn':C['warn']}[kind]
    fg = C['ink'] if kind == 'ghost' else 'white'
    return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, relief='flat',
                     font=(FONT,9,'bold'), cursor='hand2',
                     activebackground=C['accent_dk'] if kind=='primary' else bg,
                     activeforeground='white', bd=0, **kw)

def fab_card(parent, title):
    wrap = tk.Frame(parent, bg=C['bg']); wrap.pack_propagate(True)
    head = tk.Frame(wrap, bg=C['bg']); head.pack(fill='x', pady=(0,4))
    tk.Frame(head, bg=C['accent'], width=4, height=15).pack(side='left', padx=(0,7))
    tk.Label(head, text=title, bg=C['bg'], fg=C['ink'], font=(FONT,10,'bold')).pack(side='left')
    inner = tk.Frame(wrap, bg=C['surface'], highlightbackground=C['line'], highlightthickness=1)
    inner.pack(fill='both', expand=True)
    return wrap, inner

def check_duplicates(cells):
    from collections import Counter
    ids = [c['id'] for c in cells if c['id'].strip()]
    dups = {k for k,v in Counter(ids).items() if v > 1}
    for c in cells: c['dup'] = c['id'] in dups


# ════════════════════════════════════════════════════════════
#  相片表模式分頁
# ════════════════════════════════════════════════════════════

class SheetTab(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C['bg'])
        self.app = app
        self.img_path = ''; self.img = None
        self.cells = []; self.current = 0
        self.cols_var = tk.IntVar(value=5); self.rows_var = tk.IntVar(value=4)
        self.auto_var = tk.BooleanVar(value=True)
        self.excel_data = {}; self.excel_headers = []
        self.excel_names = []
        self._list_map = []
        self._ocr_running = False
        self._build()

    def _build(self):
        body = tk.Frame(self, bg=C['bg'])
        body.pack(fill='both', expand=True, padx=18, pady=12)

        # ── Step 1: 圖片 ──
        w1, s1 = fab_card(body, '步驟一　選擇相片表圖片')
        w1.pack(fill='x', pady=(0,8))
        fr = tk.Frame(s1, bg=C['surface']); fr.pack(fill='x', padx=12, pady=10)
        self.img_lbl = tk.Label(fr, text='尚未選擇…', bg=C['surface'], fg=C['muted'],
                                anchor='w', font=(FONT,9))
        self.img_lbl.pack(side='left', fill='x', expand=True)
        fab_btn(fr,'選擇圖片',self._pick_img).pack(side='right', ipadx=8, ipady=3)

        gr = tk.Frame(s1, bg=C['surface']); gr.pack(fill='x', padx=12, pady=(0,8))
        tk.Checkbutton(gr, text='自動偵測欄列', variable=self.auto_var, bg=C['surface'],
                       fg=C['ink'], selectcolor=C['surface'], activebackground=C['surface'],
                       font=(FONT,9), command=self._toggle_auto).pack(side='left', padx=(0,12))
        self.col_sp = tk.Spinbox(gr, from_=1, to=15, textvariable=self.cols_var,
                                 width=3, font=(FONT,9), state='disabled', relief='solid', bd=1)
        self.row_sp = tk.Spinbox(gr, from_=1, to=15, textvariable=self.rows_var,
                                 width=3, font=(FONT,9), state='disabled', relief='solid', bd=1)
        tk.Label(gr, text='欄：', bg=C['surface'], fg=C['ink'], font=(FONT,9)).pack(side='left')
        self.col_sp.pack(side='left', padx=(0,8))
        tk.Label(gr, text='列：', bg=C['surface'], fg=C['ink'], font=(FONT,9)).pack(side='left')
        self.row_sp.pack(side='left')
        self.grid_lbl = tk.Label(s1, text='', bg=C['surface'], fg=C['accent_dk'], font=(FONT,9))
        self.grid_lbl.pack(anchor='w', padx=12, pady=(0,8))

        # ── Step 2: Excel ──
        w2, s2 = fab_card(body, '步驟二　（選用）匯入 Excel 學號')
        w2.pack(fill='x', pady=(0,8))
        fr2 = tk.Frame(s2, bg=C['surface']); fr2.pack(fill='x', padx=12, pady=10)
        self.xl_lbl = tk.Label(fr2, text='尚未選擇…', bg=C['surface'], fg=C['muted'],
                               anchor='w', font=(FONT,9))
        self.xl_lbl.pack(side='left', fill='x', expand=True)
        fab_btn(fr2,'選擇 Excel',self._pick_excel,'ink').pack(side='right', ipadx=8, ipady=3)
        cr = tk.Frame(s2, bg=C['surface']); cr.pack(fill='x', padx=12, pady=(0,8))
        tk.Label(cr, text='學號欄位：', bg=C['surface'], fg=C['ink'], font=(FONT,9)).pack(side='left')
        self.col_combo = ttk.Combobox(cr, state='disabled', width=18, font=(FONT,9))
        self.col_combo.pack(side='left', padx=(4,10))
        self.col_combo.bind('<<ComboboxSelected>>', self._on_col_select)
        self.xl_status = tk.Label(s2, text='', bg=C['surface'], fg=C['ok'], font=(FONT,9))
        self.xl_status.pack(anchor='w', padx=12, pady=(0,8))

        # ── Step 3: 逐張確認 ──
        w3, s3 = fab_card(body, '步驟三　逐張確認學號')
        w3.pack(fill='both', expand=True, pady=(0,8))
        prow = tk.Frame(s3, bg=C['surface']); prow.pack(fill='both', expand=True, padx=12, pady=10)

        # 左：預覽（可點擊放大）
        pv = tk.Frame(prow, bg=C['surface']); pv.pack(side='left', padx=(0,12))
        self.photo_lbl = tk.Label(pv, bg=C['surface2'], width=160, height=200,
                                  text='尚未載入', font=(FONT,9), fg=C['muted'],
                                  highlightbackground=C['line'], highlightthickness=1,
                                  cursor='hand2')
        self.photo_lbl.pack()
        self.photo_lbl.bind('<Button-1>', lambda e: self._zoom_preview())
        Tooltip(self.photo_lbl, '點擊放大，可看清楚是否為正確學生\n（按 Esc 或再次點擊關閉）')
        tk.Label(pv, text='點擊可放大', bg=C['surface'], fg=C['muted'], font=(FONT,7)).pack()

        rotf = tk.Frame(pv, bg=C['surface']); rotf.pack(pady=3)
        b_left  = fab_btn(rotf,'↺ 左轉',lambda:self._rotate(-90),'ghost')
        b_right = fab_btn(rotf,'↻ 右轉',lambda:self._rotate(90),'ghost')
        b_left.pack(side='left',padx=2,ipadx=4,ipady=2)
        b_right.pack(side='left',padx=2,ipadx=4,ipady=2)
        Tooltip(b_left,  '逆時針旋轉 90°（快捷鍵 ←）')
        Tooltip(b_right, '順時針旋轉 90°（快捷鍵 →）')

        b_crop = fab_btn(pv,'✂ 精準裁切',self._open_cropper,'ink')
        b_crop.pack(pady=3,fill='x')
        Tooltip(b_crop, '開啟裁切視窗：拖曳移動、滾輪縮放\n可以去除邊框或調整構圖')

        b_mat = fab_btn(pv,'🪄 去背白底',self._apply_matting,
                        'primary' if (REMBG_OK and LICENSE.allowed('matting')) else 'ghost')
        b_mat.pack(pady=3,fill='x')
        if not REMBG_OK:
            mat_hint = '需安裝 rembg'
        elif not LICENSE.allowed('matting'):
            mat_hint = f'Pro 方案功能（目前 {TIER_LABEL[LICENSE.tier][0]}）'
            b_mat.config(state='disabled')
        else:
            mat_hint = '（首次執行需網路下載模型 ~170MB）'
        Tooltip(b_mat, f'用 AI 把人像摳出，貼在純白底\n{mat_hint}')

        self.prog_lbl = tk.Label(pv, text='', bg=C['surface'], fg=C['muted'], font=(FONT,8))
        self.prog_lbl.pack(pady=(4,0))

        # 右：學號＋姓名＋清單
        rv = tk.Frame(prow, bg=C['surface']); rv.pack(side='left', fill='both', expand=True)

        # 姓名（來自 Excel，對阿姨視覺確認用）
        self.name_lbl = tk.Label(rv, text='', bg=C['surface'], fg=C['accent_dk'],
                                  font=(FONT, 13, 'bold'))
        self.name_lbl.pack(anchor='w')

        tk.Label(rv, text='學號', bg=C['surface'], fg=C['ink'], font=(FONT,10,'bold')).pack(anchor='w')
        self.id_var = tk.StringVar()
        self.id_entry = tk.Entry(rv, textvariable=self.id_var, font=(FONT,22,'bold'),
                                 width=11, relief='flat', bd=0, fg=C['ink'],
                                 highlightbackground=C['accent'], highlightcolor=C['accent'],
                                 highlightthickness=2)
        self.id_entry.pack(anchor='w', pady=(2,1))
        self.id_entry.bind('<Return>', lambda e: self._next())
        self.id_entry.bind('<KP_Enter>', lambda e: self._next())
        Tooltip(self.id_entry, '輸入或修改學號，按 Enter 確認並跳到下一張')
        self.dup_lbl = tk.Label(rv, text='', bg=C['surface'], font=(FONT,8))
        self.dup_lbl.pack(anchor='w')

        nav = tk.Frame(rv, bg=C['surface']); nav.pack(anchor='w', pady=6)
        self.prev_btn = fab_btn(nav,'◀ 上一張',self._prev,'ink')
        self.prev_btn.pack(side='left', padx=(0,6), ipadx=8, ipady=3)
        self.prev_btn.config(state='disabled')
        self.next_btn = fab_btn(nav,'下一張 ▶',self._next)
        self.next_btn.pack(side='left', ipadx=8, ipady=3)
        self.next_btn.config(state='disabled')
        Tooltip(self.prev_btn, '回到上一張（快捷鍵 ← 方向鍵）')
        Tooltip(self.next_btn, '儲存學號並前往下一張（快捷鍵 → 方向鍵）')

        sf = tk.Frame(rv, bg=C['surface']); sf.pack(fill='x', pady=(6,2))
        tk.Label(sf, text='搜尋：', bg=C['surface'], fg=C['muted'], font=(FONT,8)).pack(side='left')
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *_: self._refresh_list())
        tk.Entry(sf, textvariable=self.search_var, font=(FONT,9), width=12,
                 relief='solid', bd=1).pack(side='left', padx=2)

        lb = tk.Frame(rv, bg=C['surface']); lb.pack(fill='both', expand=True)
        sb = tk.Scrollbar(lb); sb.pack(side='right', fill='y')
        self.listbox = tk.Listbox(lb, font=(FONT,9), height=7, yscrollcommand=sb.set,
                                  activestyle='none', bd=0, highlightthickness=1,
                                  highlightbackground=C['line'],
                                  selectbackground=C['accent'], selectforeground='white')
        self.listbox.pack(fill='both', expand=True)
        self.listbox.bind('<<ListboxSelect>>', self._on_list)
        sb.config(command=self.listbox.yview)

        # ── Step 4: 匯出 ──
        ef = tk.Frame(body, bg=C['bg']); ef.pack(fill='x')
        self.export_btn = fab_btn(ef,'步驟四　儲存 ZIP　→',self._export)
        self.export_btn.config(font=(FONT,11,'bold'), state='disabled')
        self.export_btn.pack(side='left', fill='x', expand=True, ipady=10)
        Tooltip(self.export_btn, '把全部照片打包成 ZIP 壓縮檔儲存\n每張以學號命名')

        self.quick_btn = fab_btn(ef,'⚡ 信任 OCR 直接存',self._quick_export,'ink')
        self.quick_btn.config(state='disabled')
        self.quick_btn.pack(side='left',padx=(6,0),ipady=10,ipadx=8)
        Tooltip(self.quick_btn, 'OCR 已自動填好學號，不需逐張確認\n直接跳過審閱步驟，一鍵存檔')

        fab_btn(ef,'🪄 全部去背',lambda:self._apply_matting(all_cells=True),
                'primary' if REMBG_OK else 'ghost').pack(side='left',padx=(6,0),ipady=10,ipadx=8)
        b_csv = fab_btn(ef,'CSV 報表',self._export_csv,'ghost')
        b_csv.pack(side='left',padx=(6,0),ipady=10,ipadx=10)
        Tooltip(b_csv, '匯出一份 Excel 可開啟的清單\n列出所有學號與狀態（已填／未填）')

        # 鍵盤快捷鍵
        self.bind_all('<Left>', lambda e: self._prev() if self.cells else None)
        self.bind_all('<Right>', lambda e: self._next() if self.cells else None)

    def _toggle_auto(self):
        st = 'disabled' if self.auto_var.get() else 'normal'
        self.col_sp.config(state=st); self.row_sp.config(state=st)

    def _pick_img(self):
        path = filedialog.askopenfilename(
            title='選擇相片表',
            filetypes=[('圖片','*.jpg *.jpeg *.png *.bmp'),('所有','*.*')])
        if not path: return
        self.img_path = path
        self.img_lbl.config(text=os.path.basename(path), fg=C['ink'])
        self._load_image()

    def _load_image(self):
        if not self.img_path: return
        self.img = Image.open(self.img_path).convert('RGB')
        fc = fr = None
        if not self.auto_var.get(): fc, fr = self.cols_var.get(), self.rows_var.get()
        pd = ProgressDlg(self.app, '正在偵測格線…', 100)
        col_bands, photo_rows = detect_grid(self.img, fc, fr)
        if not col_bands or not photo_rows:
            pd.close()
            messagebox.showwarning('偵測失敗','找不到照片格線，請改用手動指定欄列數。'); return
        self.cells = extract_cells(self.img, col_bands, photo_rows)
        pd.close()
        if not self.cells:
            messagebox.showwarning('沒有照片','未擷取到任何照片，請檢查設定。'); return

        # ── Tier gate：Free 最多 30 張 ──
        max_c = LICENSE.max_cells()
        if max_c and len(self.cells) > max_c:
            messagebox.showwarning(
                '超過方案上限',
                f'目前 {TIER_LABEL[LICENSE.tier][0]} 方案限制 {max_c} 張，\n'
                f'本相片表共 {len(self.cells)} 張，只處理前 {max_c} 張。\n\n'
                f'升級 Pro 方案可無限制使用。')
            self.cells = self.cells[:max_c]
        self.grid_lbl.config(text=f'偵測 {len(col_bands)}欄 × {len(photo_rows)}列，擷取 {len(self.cells)} 張')
        check_duplicates(self.cells)
        if self._load_session():
            pass
        # ── UI 立即顯示全部照片，不等 OCR ──
        self._refresh_list(); self._show(0)
        self.next_btn.config(state='normal')
        self.export_btn.config(state='normal')
        self.quick_btn.config(state='normal')
        if self.excel_data and self.col_combo.get():
            self._apply_excel(self.col_combo.get(), only_blank=True)
        # ── OCR 背景執行緒（Scholar/Pro/Enterprise 才有）──
        if OCR_OK and LICENSE.allowed('ocr'):
            self._ocr_running = True
            # 鎖定匯出，避免 OCR 未完成就存出空白學號檔名
            self.export_btn.config(state='disabled', text='⋯ OCR 辨識中，請稍候')
            self.quick_btn.config(state='disabled')
            self.grid_lbl.config(text=self.grid_lbl.cget('text') + '   ⋯ OCR 辨識中')
            row_tops = sorted({c['_cell_box'][1] for c in self.cells})
            def _ocr_worker():
                for k, c in enumerate(self.cells):
                    if not self._ocr_running: break
                    y = c['_cell_box'][1]
                    later = [t for t in row_tops if t > y]
                    c['id'] = ocr_id_below(self.img, c['_cell_box'],
                                           next_row_top=min(later) if later else None)
                    # 安全地更新 UI（必須在主執行緒）
                    self.app.after(0, lambda k=k: self._on_ocr_step(k))
                self.app.after(0, self._on_ocr_done)
            threading.Thread(target=_ocr_worker, daemon=True).start()

    def _on_ocr_step(self, k):
        """每張 OCR 完成後在主執行緒更新 UI（不凍結）。"""
        check_duplicates(self.cells); self._refresh_list()
        if k == self.current: self._show(self.current)

    def _on_ocr_done(self):
        """全部 OCR 完成，解鎖匯出按鈕，補填 Excel。"""
        self._ocr_running = False
        txt = self.grid_lbl.cget('text').replace('   ⋯ OCR 辨識中', '')
        filled = sum(1 for c in self.cells if c['id'].strip())
        self.grid_lbl.config(text=txt + f'   ✓ OCR 完成（{filled}/{len(self.cells)} 筆）')
        # 解鎖匯出
        self.export_btn.config(state='normal', text='步驟四　儲存 ZIP　→')
        self.quick_btn.config(state='normal')
        check_duplicates(self.cells); self._refresh_list()
        if self.excel_data and self.col_combo.get():
            self._apply_excel(self.col_combo.get(), only_blank=True)

    def _pick_excel(self):
        if not EXCEL_OK: messagebox.showerror('缺少套件','pip install openpyxl'); return
        path = filedialog.askopenfilename(title='選擇 Excel',
            filetypes=[('Excel','*.xlsx *.xls'),('所有','*.*')])
        if not path: return
        try: data, hdrs = load_excel(path)
        except Exception as e: messagebox.showerror('讀取失敗',str(e)); return
        self.excel_data, self.excel_headers = data, hdrs
        self.xl_lbl.config(text=os.path.basename(path), fg=C['ink'])
        self.col_combo.config(values=hdrs, state='readonly')
        g = guess_id_col(data, hdrs)
        if g:
            self.col_combo.set(g); self.xl_status.config(text=f'已自動選定欄「{g}」')
            self._on_col_select(None)
        else: self.xl_status.config(text='請手動選擇學號欄位 ↑')

    def _on_col_select(self, _e):
        col = self.col_combo.get()
        if not col or col not in self.excel_data: return
        if not self.cells: self.xl_status.config(text='已記住欄位，載入圖片後套用'); return
        only_blank = True
        if any(c['id'].strip() for c in self.cells):
            only_blank = not messagebox.askyesno('套用方式','已有 OCR 學號。\n覆蓋全部？（否=只填補空白）')
        self._apply_excel(col, only_blank)

    def _apply_excel(self, col, only_blank):
        vals = [v for v in self.excel_data[col] if v.strip()]
        # 嘗試找姓名欄
        name_col = next((h for h in self.excel_headers
                         if any(k in h for k in ('姓名','名字','name','Name'))), None)
        names = self.excel_data.get(name_col, []) if name_col else []
        n = 0
        for i, c in enumerate(self.cells):
            if i >= len(vals): break
            if only_blank and c['id'].strip(): continue
            c['id'] = vals[i]
            c['excel_name'] = names[i] if i < len(names) else ''
            n += 1
        self.xl_status.config(text=f'✓ Excel 填入 {n} 筆' + (f'，含姓名欄「{name_col}」' if name_col else ''))
        check_duplicates(self.cells); self._refresh_list()
        if self.cells: self._show(self.current)

    def _refresh_list(self):
        q = self.search_var.get().strip().lower()
        self.listbox.delete(0,'end'); self._list_map = []
        for i, c in enumerate(self.cells):
            sid = c['id']
            if q and q not in sid.lower() and q not in str(i+1): continue
            mark = '✓' if sid.strip() else '○'
            extra = ' ⚠ 重複' if c.get('dup') else ''
            self.listbox.insert('end', f'  {mark}  {i+1:3d}.  {sid or "（未填）"}{extra}')
            fg = C['warn'] if c.get('dup') else (C['ok'] if sid.strip() else C['err'])
            self.listbox.itemconfig(self.listbox.size()-1, fg=fg)
            self._list_map.append(i)

    def _show(self, idx):
        if not self.cells or not (0 <= idx < len(self.cells)): return
        self.current = idx; c = self.cells[idx]
        preview = c['image'].copy(); preview.thumbnail((160,200))
        tk_img = ImageTk.PhotoImage(preview)
        self.photo_lbl.config(image=tk_img, text=''); self.photo_lbl.image = tk_img

        # 進度（✓已填 / ○未填 / 總計）
        total = len(self.cells)
        filled = sum(1 for cc in self.cells if cc['id'].strip())
        remaining = total - filled
        prog_text = f'{idx+1} / {total}　　✓已填 {filled}'
        if remaining: prog_text += f'　⚠未填 {remaining}'
        self.prog_lbl.config(text=prog_text,
                             fg=C['err'] if remaining else C['ok'])

        # 姓名（來自 Excel）
        name = c.get('excel_name','')
        self.name_lbl.config(text=name if name else '')

        self.id_var.set(c['id']); self.id_entry.focus(); self.id_entry.selection_range(0,'end')
        self.prev_btn.config(state='normal' if idx > 0 else 'disabled')
        is_last = idx == total - 1
        self.next_btn.config(text='完成 ✓' if is_last else '下一張 ▶',
                             bg=C['accent_dk'] if is_last else C['accent'])
        self.dup_lbl.config(text='⚠ 此學號重複！' if c.get('dup') else '按 Enter 跳下一張',
                            fg=C['warn'] if c.get('dup') else C['accent_dk'])
        for lb_i, ci in enumerate(self._list_map):
            if ci == idx:
                self.listbox.selection_clear(0,'end')
                self.listbox.selection_set(lb_i); self.listbox.see(lb_i); break

    def _save_cur(self):
        if self.current < len(self.cells):
            self.cells[self.current]['id'] = self.id_var.get().strip()
        check_duplicates(self.cells); self._refresh_list()
        self._save_session()   # 每次確認學號就自動儲存進度

    def _prev(self): self._save_cur(); self._show(self.current-1)
    def _next(self):
        self._save_cur()
        if self.current < len(self.cells)-1: self._show(self.current+1)

    def _on_list(self, _e):
        sel = self.listbox.curselection()
        if not sel: return
        ci = self._list_map[sel[0]] if self._list_map else sel[0]
        if ci == self.current: return
        self._save_cur(); self._show(ci)

    def _open_cropper(self):
        if not self.cells: return
        self._save_cur()
        dlg = CropperDialog(self.app, self.cells[self.current])
        self.app.wait_window(dlg)   # 等對話框真正關閉後才刷新預覽
        self._show(self.current)

    def _apply_matting(self, all_cells=False):
        if not self.cells: return
        if not LICENSE.allowed('matting'):
            messagebox.showinfo('需要升級',
                f'AI 去背白底是 Pro / Enterprise 方案功能。\n'
                f'目前方案：{TIER_LABEL[LICENSE.tier][0]}\n\n'
                '點擊標頭的方案標籤可輸入授權金鑰升級。')
            return
        if not REMBG_OK:
            messagebox.showinfo('缺少套件',
                '請先安裝去背引擎：\n\npip install rembg onnxruntime\n\n'
                '首次執行會自動下載 AI 模型（~170MB），之後完全離線使用。')
            return
        targets = self.cells if all_cells else [self.cells[self.current]]
        if all_cells and not messagebox.askyesno('批次去背',
            f'將對全部 {len(targets)} 張套用 AI 去背白底。\n每張約需 2-5 秒，確定繼續？'): return

        pd = ProgressDlg(self.app, '正在 AI 去背…', len(targets)) if all_cells else None

        def _worker():
            errors = []
            for k, c in enumerate(targets):
                try:
                    src = render_cell(c)
                    white = matting_to_white(src)
                    c['_orig'] = white; c['rotation'] = 0; c['image'] = white.copy()
                except Exception as e:
                    errors.append(f"第{self.cells.index(c)+1}張：{e}")
                if pd: self.app.after(0, lambda v=k+1: pd.step(v, f'去背 {v}/{len(targets)}'))
            def _done():
                if pd: pd.close()
                if errors: messagebox.showwarning('部分失敗', '\n'.join(errors[:5]))
                self._show(self.current)
            self.app.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def _rotate(self, deg):
        if not self.cells: return
        c = self.cells[self.current]
        c['rotation'] = (c['rotation']+deg)%360; c['image'] = render_cell(c)
        self._show(self.current)

    def _export(self):
        if not self.cells: return
        # 安全網：OCR 還在背景跑（不應該發生，但以防萬一）
        if self._ocr_running:
            messagebox.showwarning('OCR 尚未完成',
                'OCR 學號辨識還在進行中，\n請等按鈕變回「步驟四　儲存 ZIP」再匯出。')
            return
        self._save_cur()
        empty = [i+1 for i,c in enumerate(self.cells) if not c['id'].strip()]
        if empty and not messagebox.askyesno('有空白學號',
            f'第 {empty} 張未填學號，將以 photo_N 命名。確定繼續？'): return
        dups = list({c['id'] for c in self.cells if c.get('dup')})
        if dups and not messagebox.askyesno('學號重複',
            f'以下學號重複：{dups}\n確定繼續？'): return
        base = os.path.splitext(os.path.basename(self.img_path))[0]
        zip_path = filedialog.asksaveasfilename(
            title='儲存 ZIP', defaultextension='.zip',
            initialfile=f'{base}_照片.zip', filetypes=[('ZIP','*.zip')])
        if not zip_path: return
        self.export_btn.config(state='disabled'); self.quick_btn.config(state='disabled')
        total = len(self.cells)
        pd = ProgressDlg(self.app, '正在儲存照片…', total)
        snap = list(self.cells)   # thread-safe snapshot

        def _worker():
            try:
                seen = {}; buf = io.BytesIO()
                with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                    for i, c in enumerate(snap):
                        sid = c['id'].strip()
                        name = safe_filename(sid) if sid else f'photo_{i+1}'
                        seen[name] = seen.get(name, 0) + 1
                        if seen[name] > 1: name = f'{name}_{seen[name]}'
                        ib = io.BytesIO()
                        render_cell(c).save(ib, 'JPEG', quality=95)
                        zf.writestr(f'{name}.jpg', ib.getvalue())
                        self.app.after(0, lambda v=i+1: pd.step(v, f'儲存 {v}/{total}'))
                bv = buf.getvalue()
                def _done():
                    pd.close()
                    with open(zip_path, 'wb') as f: f.write(bv)
                    self._save_session()
                    self.export_btn.config(state='normal'); self.quick_btn.config(state='normal')
                    self.app.bell()
                    messagebox.showinfo('完成 ✓', f'已儲存 {total} 張\n\n📁 {zip_path}')
                self.app.after(0, _done)
            except Exception as e:
                def _err(m=str(e)):
                    pd.close()
                    self.export_btn.config(state='normal'); self.quick_btn.config(state='normal')
                    messagebox.showerror('儲存失敗', m)
                self.app.after(0, _err)

        threading.Thread(target=_worker, daemon=True).start()

    def _zoom_preview(self):
        """點擊預覽照片，彈出大圖視窗。"""
        if not self.cells: return
        ZoomDialog(self.app, self.cells[self.current])

    def _quick_export(self):
        """信任 OCR 結果，跳過逐張確認，直接匯出。"""
        if not self.cells: return
        empty = [i+1 for i,c in enumerate(self.cells) if not c['id'].strip()]
        filled = len(self.cells) - len(empty)
        msg = f'OCR 已自動填入 {filled} / {len(self.cells)} 個學號。\n'
        if empty:
            msg += f'第 {empty[:5]}{"…等" if len(empty)>5 else ""} 張尚未填，將以 photo_N 命名。\n'
        msg += '\n確定直接儲存，不逐張確認？'
        if not messagebox.askyesno('⚡ 信任 OCR 直接存', msg): return
        self._export()

    def _save_session(self):
        """把目前所有學號和旋轉角度存到 .session.json，供下次開啟回復。"""
        if not self.img_path or not self.cells: return
        try:
            path = self.img_path + '.session.json'
            data = {
                'version': 1,
                'img_path': self.img_path,
                'saved_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'cells': [{'id': c['id'], 'rotation': c['rotation'],
                            'excel_name': c.get('excel_name','')} for c in self.cells],
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # session 儲存失敗不影響主流程

    def _load_session(self) -> bool:
        """嘗試讀取 .session.json 回復上次進度，回傳是否成功套用。"""
        if not self.img_path or not self.cells: return False
        path = self.img_path + '.session.json'
        if not os.path.exists(path): return False
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data.get('img_path') != self.img_path: return False
            saved = data.get('cells', [])
            if len(saved) != len(self.cells): return False
            saved_at = data.get('saved_at', '上次')
            if not messagebox.askyesno('發現上次進度',
                f'偵測到 {saved_at} 的未完成紀錄（{len(saved)} 張）。\n\n'
                '要回復上次填寫的學號嗎？\n（選「否」會從頭開始）'):
                return False
            for c, s in zip(self.cells, saved):
                c['id'] = s.get('id', '')
                c['rotation'] = s.get('rotation', 0)
                c['excel_name'] = s.get('excel_name', '')
                c['image'] = render_cell(c)
            check_duplicates(self.cells)
            return True
        except Exception:
            return False

    def _export_csv(self):
        if not self.cells: messagebox.showinfo('','請先載入相片表'); return
        self._save_cur()
        path = filedialog.asksaveasfilename(
            title='儲存 CSV 報表', defaultextension='.csv',
            initialfile='裁切報表.csv', filetypes=[('CSV','*.csv')])
        if not path: return
        with open(path,'w',newline='',encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['序號','學號','狀態','備註'])
            for i, c in enumerate(self.cells):
                sid = c['id'].strip()
                w.writerow([i+1, sid, '已填' if sid else '未填',
                             '學號重複' if c.get('dup') else ''])
        messagebox.showinfo('完成',f'報表已儲存：{path}')


# ════════════════════════════════════════════════════════════
#  批次照片模式分頁
# ════════════════════════════════════════════════════════════

class BatchTab(tk.Frame):
    """
    資料夾 → 逐張個別照片 → 比對 Excel 學生名單
    支援 600+ 學生。自動配對、缺件標示、批次匯出。
    """
    STATUS_COLOR = {'matched': '#1FA85B', 'unmatched': '#D8A000',
                    'missing': '#D8452F', 'orphan': '#8C8474'}

    def __init__(self, parent, app):
        super().__init__(parent, bg=C['bg'])
        self.app = app
        self.folder = ''
        self.excel_data = {}; self.excel_headers = []
        self.items = []       # 全部條目
        self.filtered = []    # 目前顯示的條目（indices into self.items）
        self.current = -1
        self.filter_var = tk.StringVar(value='all')
        self._build()

    # item 結構:
    # { type:'matched'|'unmatched'|'missing'|'orphan',
    #   path:str|None, filename:str, id:str, auto_id:str,
    #   excel_name:str, _img:PIL|None }

    def _build(self):
        body = tk.Frame(self, bg=C['bg']); body.pack(fill='both',expand=True,padx=18,pady=12)

        # ── 控制列 ──
        ctrl = tk.Frame(body, bg=C['bg']); ctrl.pack(fill='x', pady=(0,10))

        wf, sf = fab_card(ctrl, '資料來源')
        wf.pack(side='left', fill='both', expand=True, padx=(0,8))
        fab_btn(sf,'選擇照片資料夾',self._pick_folder).pack(anchor='w',padx=12,pady=(10,4),ipadx=6,ipady=3)
        self.folder_lbl = tk.Label(sf, text='尚未選擇…', bg=C['surface'], fg=C['muted'],
                                   font=(FONT,8), anchor='w')
        self.folder_lbl.pack(anchor='w',padx=12,pady=(0,4))
        fab_btn(sf,'選擇 Excel 學生名單',self._pick_excel,'ink').pack(anchor='w',padx=12,pady=(0,4),ipadx=6,ipady=3)
        self.excel_lbl = tk.Label(sf, text='尚未選擇…', bg=C['surface'], fg=C['muted'],
                                  font=(FONT,8), anchor='w')
        self.excel_lbl.pack(anchor='w',padx=12,pady=(0,8))

        ws, ss = fab_card(ctrl, '統計')
        ws.pack(side='left', fill='y')
        self.stat_lbl = tk.Label(ss, text='—', bg=C['surface'], fg=C['ink'],
                                 font=(FONT,10), justify='left', padx=12, pady=10)
        self.stat_lbl.pack()

        # ── 篩選鈕 ──
        ff = tk.Frame(body, bg=C['bg']); ff.pack(fill='x', pady=(0,6))
        for text, val in [('全部','all'),('已配對','matched'),('未配對','unmatched'),
                          ('缺照片','missing'),('孤立照片','orphan')]:
            rb = tk.Radiobutton(ff, text=text, variable=self.filter_var, value=val,
                                bg=C['bg'], fg=C['ink'], selectcolor=C['bg'],
                                activebackground=C['bg'], font=(FONT,9), cursor='hand2',
                                command=self._apply_filter)
            rb.pack(side='left', padx=(0,10))

        # ── 主清單 + 預覽 ──
        mid = tk.Frame(body, bg=C['bg']); mid.pack(fill='both', expand=True)

        # 清單
        lf = tk.Frame(mid, bg=C['bg']); lf.pack(side='left', fill='both', expand=True, padx=(0,10))
        lb_f = tk.Frame(lf, bg=C['surface'], highlightbackground=C['line'], highlightthickness=1)
        lb_f.pack(fill='both', expand=True)
        sb = tk.Scrollbar(lb_f); sb.pack(side='right', fill='y')
        self.listbox = tk.Listbox(lb_f, font=(FONT,9), yscrollcommand=sb.set,
                                  activestyle='none', bd=0, highlightthickness=0,
                                  selectbackground=C['accent'], selectforeground='white')
        self.listbox.pack(fill='both', expand=True)
        self.listbox.bind('<<ListboxSelect>>', self._on_list)
        sb.config(command=self.listbox.yview)

        # 預覽面板
        pf = tk.Frame(mid, bg=C['surface'], width=280,
                      highlightbackground=C['line'], highlightthickness=1)
        pf.pack(side='left', fill='y'); pf.pack_propagate(False)
        self.preview = tk.Label(pf, bg=C['surface2'], width=240, height=300,
                                text='選擇項目', font=(FONT,9), fg=C['muted'])
        self.preview.pack(padx=16, pady=16)
        self.status_lbl = tk.Label(pf, text='', bg=C['surface'], fg=C['ink'],
                                   font=(FONT,9), wraplength=240, justify='left')
        self.status_lbl.pack(padx=16, pady=(0,8))
        tk.Label(pf, text='手動指定學號', bg=C['surface'], fg=C['muted'], font=(FONT,8)).pack(anchor='w', padx=16)
        self.batch_id_var = tk.StringVar()
        ie = tk.Entry(pf, textvariable=self.batch_id_var, font=(FONT,14,'bold'),
                      width=12, relief='flat', bd=0, fg=C['ink'],
                      highlightbackground=C['accent'], highlightcolor=C['accent'],
                      highlightthickness=2)
        ie.pack(padx=16, pady=(2,6))
        ie.bind('<Return>', lambda e: self._assign_id())
        fab_btn(pf,'套用學號',self._assign_id).pack(padx=16,pady=(0,6),fill='x',ipadx=4,ipady=3)
        fab_btn(pf,'✂ 精準裁切',self._open_cropper,'ink').pack(padx=16,pady=(0,16),fill='x',ipadx=4,ipady=3)

        # ── 匯出列 ──
        ef = tk.Frame(body, bg=C['bg']); ef.pack(fill='x', pady=(8,0))
        self.export_btn = fab_btn(ef,'批次匯出 ZIP　→',self._export)
        self.export_btn.config(font=(FONT,12,'bold'), state='disabled')
        self.export_btn.pack(side='left', fill='x', expand=True, ipady=10)
        self.matting_var = tk.BooleanVar(value=False)
        tk.Checkbutton(ef, text='🪄 去背白底', variable=self.matting_var,
                       bg=C['bg'], fg=C['ink'] if REMBG_OK else C['muted'],
                       selectcolor=C['bg'], activebackground=C['bg'],
                       font=(FONT,9,'bold'),
                       state='normal' if REMBG_OK else 'disabled').pack(side='left',padx=8)
        fab_btn(ef,'CSV 報表',self._export_csv,'ghost').pack(side='left',padx=(6,0),ipady=10,ipadx=10)

    def _pick_folder(self):
        folder = filedialog.askdirectory(title='選擇照片資料夾')
        if not folder: return
        self.folder = folder
        self.folder_lbl.config(text=folder, fg=C['ink'])
        self._build_items()

    def _pick_excel(self):
        if not EXCEL_OK: messagebox.showerror('缺少套件','pip install openpyxl'); return
        path = filedialog.askopenfilename(title='選擇 Excel 學生名單',
            filetypes=[('Excel','*.xlsx *.xls'),('所有','*.*')])
        if not path: return
        try: data, hdrs = load_excel(path)
        except Exception as e: messagebox.showerror('讀取失敗',str(e)); return
        self.excel_data, self.excel_headers = data, hdrs
        self.excel_lbl.config(text=os.path.basename(path), fg=C['ink'])
        self._build_items()

    def _build_items(self):
        if not self.folder: return
        try: photos = scan_folder(self.folder)
        except RuntimeError as e: messagebox.showerror('錯誤',str(e)); return

        # 讀 Excel 學號 + 姓名
        excel_ids, excel_names = [], []
        if self.excel_data:
            id_col = guess_id_col(self.excel_data, self.excel_headers)
            name_col = next((h for h in self.excel_headers if '姓名' in h or 'name' in h.lower()), None)
            if id_col:
                excel_ids = self.excel_data[id_col]
                excel_names = self.excel_data[name_col] if name_col else ['']*len(excel_ids)

        excel_set = set(excel_ids)
        matched_ids = set()
        items = []

        for p in photos:
            auto_id = p['auto_id']
            if auto_id and auto_id in excel_set:
                idx = excel_ids.index(auto_id)
                items.append({'type':'matched','path':p['path'],'filename':p['filename'],
                              'id':auto_id,'auto_id':auto_id,
                              'excel_name':excel_names[idx] if idx < len(excel_names) else '',
                              '_img':None})
                matched_ids.add(auto_id)
            else:
                items.append({'type':'orphan' if excel_ids else 'unmatched',
                              'path':p['path'],'filename':p['filename'],
                              'id':auto_id,'auto_id':auto_id,'excel_name':'',
                              '_img':None})

        for eid, name in zip(excel_ids, excel_names):
            if eid not in matched_ids:
                items.append({'type':'missing','path':None,'filename':'',
                              'id':eid,'auto_id':eid,'excel_name':name,'_img':None})

        self.items = items
        self._update_stats()
        self._apply_filter()
        self.export_btn.config(state='normal' if items else 'disabled')

    def _update_stats(self):
        from collections import Counter
        cnt = Counter(it['type'] for it in self.items)
        self.stat_lbl.config(text=(
            f"總計　{len(self.items)} 筆\n"
            f"✓ 配對　{cnt.get('matched',0)}\n"
            f"? 未配對　{cnt.get('unmatched',0)+cnt.get('orphan',0)}\n"
            f"✗ 缺照片　{cnt.get('missing',0)}"))

    def _apply_filter(self):
        f = self.filter_var.get()
        if f == 'all':
            self.filtered = list(range(len(self.items)))
        elif f == 'unmatched':
            self.filtered = [i for i,it in enumerate(self.items) if it['type'] in ('unmatched','orphan')]
        else:
            self.filtered = [i for i,it in enumerate(self.items) if it['type'] == f]
        self._rebuild_list()

    def _rebuild_list(self):
        self.listbox.delete(0,'end')
        for fi in self.filtered:
            it = self.items[fi]
            if it['type'] == 'matched':
                txt = f"  ✓  {it['id']}  {it['excel_name']}  ←  {it['filename']}"
            elif it['type'] == 'missing':
                txt = f"  ✗  {it['id']}  {it['excel_name']}  ← 缺照片"
            else:
                txt = f"  ?  {it['filename']}  (ID: {it['id'] or '—'})"
            self.listbox.insert('end', txt)
            self.listbox.itemconfig(self.listbox.size()-1, fg=self.STATUS_COLOR.get(it['type'],C['ink']))

    def _on_list(self, _e):
        sel = self.listbox.curselection()
        if not sel: return
        fi = self.filtered[sel[0]]
        # ── 釋放前一張全圖（memory：只保留當前一張）──
        if self.current >= 0 and self.current != fi:
            prev = self.items[self.current]
            if prev.get('_img') is not None:
                prev['_img'] = None   # GC 會回收
        self.current = fi
        it = self.items[fi]
        self.batch_id_var.set(it['id'])
        status_map = {'matched':'已配對','missing':'缺照片','unmatched':'未配對','orphan':'無對應學生'}
        info = (f"狀態：{status_map.get(it['type'],'—')}\n"
                f"學號：{it['id'] or '—'}\n"
                f"姓名：{it['excel_name'] or '—'}\n"
                f"檔案：{it['filename'] or '—'}")
        self.status_lbl.config(text=info)
        if it['path']:
            if it['_img'] is None:
                try:
                    raw = Image.open(it['path']).convert('RGB')
                    cleaned = strip_photo_borders(np.asarray(raw))
                    it['_img'] = Image.fromarray(cleaned)
                except Exception: it['_img'] = None
            if it['_img']:
                thumb = it['_img'].copy(); thumb.thumbnail((240,300))
                tk_img = ImageTk.PhotoImage(thumb)
                self.preview.config(image=tk_img, text=''); self.preview.image = tk_img
                return
        self.preview.config(image='', text='缺照片' if it['type']=='missing' else '無法載入')
        self.preview.image = None

    def _assign_id(self):
        if self.current < 0: return
        new_id = self.batch_id_var.get().strip()
        it = self.items[self.current]
        old_type = it['type']
        it['id'] = new_id
        # 更新配對狀態
        excel_ids = []
        if self.excel_data:
            id_col = guess_id_col(self.excel_data, self.excel_headers)
            if id_col: excel_ids = self.excel_data[id_col]
        if new_id in excel_ids:
            it['type'] = 'matched'
            idx = excel_ids.index(new_id)
            name_col = next((h for h in self.excel_headers if '姓名' in h or 'name' in h.lower()), None)
            it['excel_name'] = self.excel_data[name_col][idx] if name_col and idx < len(self.excel_data[name_col]) else ''
        self._update_stats(); self._apply_filter()

    def _open_cropper(self):
        if self.current < 0: return
        it = self.items[self.current]
        if not it['path'] or it['_img'] is None: return
        cell = {'_orig': it['_img'].copy(), 'rotation': 0, 'image': None}
        cell['image'] = render_cell(cell)
        dlg = CropperDialog(self.app, cell)
        self.app.wait_window(dlg)   # 等套用後才繼續
        it['_img'] = cell['_orig']
        # 刷新預覽縮圖
        thumb = it['_img'].copy(); thumb.thumbnail((240, 300))
        tk_img = ImageTk.PhotoImage(thumb)
        self.preview.config(image=tk_img, text='')
        self.preview.image = tk_img

    def _export(self):
        matched = [it for it in self.items if it['type']=='matched' and it['path']]
        if not matched: messagebox.showinfo('','沒有可匯出的已配對照片'); return
        do_matting = REMBG_OK and self.matting_var.get()
        zip_path = filedialog.asksaveasfilename(
            title='批次匯出 ZIP', defaultextension='.zip',
            initialfile='批次照片.zip', filetypes=[('ZIP','*.zip')])
        if not zip_path: return
        label = '正在 AI 去背並匯出…' if do_matting else '正在匯出…'
        pd = ProgressDlg(self.app, label, len(matched))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf,'w',compression=zipfile.ZIP_DEFLATED) as zf:
            for k, it in enumerate(matched):
                try:
                    raw = Image.open(it['path']).convert('RGB')
                    cleaned = strip_photo_borders(np.asarray(raw))
                    img = Image.fromarray(cleaned)
                    if do_matting:
                        out = matting_to_white(img)         # AI 去背白底
                    else:
                        out = img.resize((420,530), Image.LANCZOS)
                    ib = io.BytesIO(); out.save(ib,'JPEG',quality=95)
                    name = safe_filename(it['id'])
                    zf.writestr(f'{name}.jpg', ib.getvalue())
                except Exception as e:
                    pd.close(); messagebox.showerror('錯誤',f"處理 {it['filename']} 失敗：{e}"); return
                pd.step(k+1, f'{"去背" if do_matting else "匯出"} {k+1}/{len(matched)}')
        pd.close()
        with open(zip_path,'wb') as f: f.write(buf.getvalue())
        suffix = '（AI 去背白底）' if do_matting else ''
        messagebox.showinfo('完成',f'已匯出 {len(matched)} 張{suffix}\n儲存：{zip_path}')

    def _export_csv(self):
        if not self.items: messagebox.showinfo('','請先載入資料'); return
        path = filedialog.asksaveasfilename(
            title='儲存 CSV 報表', defaultextension='.csv',
            initialfile='批次報表.csv', filetypes=[('CSV','*.csv')])
        if not path: return
        status_map = {'matched':'已配對','missing':'缺照片','unmatched':'未配對','orphan':'無對應'}
        with open(path,'w',newline='',encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['學號','姓名','狀態','照片檔名'])
            for it in self.items:
                w.writerow([it['id'],it['excel_name'],status_map.get(it['type'],'—'),it['filename']])
        messagebox.showinfo('完成',f'報表已儲存：{path}')


# ════════════════════════════════════════════════════════════
#  主視窗
# ════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('一卡通相片裁切工具')
        self.geometry('880x840')
        self.minsize(820,720)
        self.configure(bg=C['bg'])
        self._init_fonts(); self._build()

    def _init_fonts(self):
        fams = set(tkfont.families(self))
        def pick(prefs, default):
            for p in prefs:
                if p in fams: return p
            return default
        self.f_mono = pick(['JetBrains Mono','Geist Mono','Consolas','Courier New'],'Courier New')
        self.f_disp = pick(['Pixelify Sans','Consolas','Courier New'],'Courier New')
        st = ttk.Style(self)
        try: st.theme_use('clam')
        except: pass
        st.configure('Fab.TNotebook', background=C['bg'], tabmargins=[0,0,0,0])
        st.configure('Fab.TNotebook.Tab', background=C['surface2'], foreground=C['muted'],
                     font=(FONT,10,'bold'), padding=[18,8])
        st.map('Fab.TNotebook.Tab',
               background=[('selected',C['surface'])],
               foreground=[('selected',C['accent'])])

    def _build(self):
        outer = tk.Frame(self, bg=C['bg'], highlightbackground=C['ink'], highlightthickness=2)
        outer.pack(fill='both', expand=True)

        # 頂部標頭
        top = tk.Frame(outer, bg=C['header'], height=58)
        top.pack(fill='x'); top.pack_propagate(False)
        mark = tk.Frame(top, bg=C['hfg'], width=32, height=32)
        mark.pack(side='left', padx=(18,10)); mark.pack_propagate(False)
        tk.Label(mark, text='/', bg=C['hfg'], fg=C['ink'],
                 font=(self.f_disp,16,'bold')).pack(expand=True)
        tk.Label(top, text='一卡通相片裁切工具', bg=C['header'], fg=C['hfg'],
                 font=(FONT,14,'bold')).pack(side='left', pady=14)
        status = ('OCR READY' if OCR_OK else 'OCR OFFLINE')
        status += '  ·  ' + ('BG.AI READY' if REMBG_OK else 'BG.AI OFFLINE')
        status += '  ·  v' + VERSION
        tk.Label(top, text=status, bg=C['header'],
                 fg=C['accent'] if (OCR_OK or REMBG_OK) else C['muted'],
                 font=(self.f_mono,9,'bold')).pack(side='right', padx=20)

        # Tier badge
        tier_label, tier_color = TIER_LABEL.get(LICENSE.tier, ('FREE','#8C8474'))
        org_txt = f'  {LICENSE.org}' if LICENSE.org else ''
        tier_btn = tk.Button(top, text=f'{tier_label}{org_txt}',
                             bg=tier_color, fg='white',
                             font=(self.f_mono, 9, 'bold'), relief='flat',
                             cursor='hand2', padx=10, pady=2, bd=0,
                             command=self._open_license_dialog)
        tier_btn.pack(side='right', padx=(0,8), pady=12)

        # 分頁
        nb = ttk.Notebook(outer, style='Fab.TNotebook')
        nb.pack(fill='both', expand=True)
        self.sheet_tab = SheetTab(nb, self)
        self.batch_tab = BatchTab(nb, self)
        nb.add(self.sheet_tab, text='  相片表模式  ')
        # Batch tab only available for scholar/pro/enterprise
        if LICENSE.allowed('batch'):
            nb.add(self.batch_tab, text='  批次照片模式  ')
        else:
            nb.add(self.batch_tab, text='  批次模式（Pro）  ', state='disabled')

        # 狀態列
        tk.Frame(outer, bg=C['line'], height=1).pack(fill='x')
        self.status_bar = tk.Label(outer,
            text='本工具所有處理在本機完成，不上傳任何學生資料。',
            bg=C['surface2'], fg=C['muted'], font=(FONT,8), anchor='w', padx=10)
        self.status_bar.pack(fill='x')

        # 啟動時 license 錯誤提示
        if LICENSE.error():
            self.after(500, lambda: messagebox.showwarning(
                '授權問題', LICENSE.error() + '\n\n說明 → 輸入授權金鑰 可重新啟動。'))

        # GitHub 更新通知（背景）
        check_github_update(GITHUB_REPO, VERSION, self._on_update_found)

        # Help 選單
        menubar = tk.Menu(self)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label='輸入授權金鑰…', command=self._open_license_dialog)
        help_menu.add_separator()
        help_menu.add_command(label=f'版本 v{VERSION}', state='disabled')
        help_menu.add_command(label='檢查更新',
            command=lambda: check_github_update(GITHUB_REPO, VERSION,
                lambda v: self.after(0, lambda: messagebox.showinfo('有新版本', f'v{v} 已發布！')))
                or (None if True else None))
        menubar.add_cascade(label='說明', menu=help_menu)
        self.config(menu=menubar)

    def _on_update_found(self, latest: str):
        """有新版本時在主執行緒顯示 banner。"""
        def _show():
            self.status_bar.config(
                text=f'🚀  新版本 v{latest} 已發布！點此更新 →',
                fg=C['accent'], bg=C['surface2'], cursor='hand2')
            self.status_bar.bind('<Button-1>',
                lambda e: __import__('webbrowser').open(
                    f'https://github.com/{GITHUB_REPO}/releases/latest'))
        self.after(0, _show)

    def _open_license_dialog(self):
        """授權金鑰啟動對話框。"""
        dlg = tk.Toplevel(self)
        dlg.title('授權管理'); dlg.geometry('520x320')
        dlg.resizable(False, False); dlg.configure(bg=C['surface'])
        dlg.transient(self); dlg.grab_set()

        tk.Label(dlg, text='Fabrica Photo Cutter  ·  授權管理',
                 bg=C['surface'], fg=C['ink'], font=(FONT,12,'bold')).pack(pady=(20,4))

        tier_lbl, tier_clr = TIER_LABEL.get(LICENSE.tier, ('FREE','#8C8474'))
        tk.Label(dlg, text=f'目前方案：{tier_lbl}',
                 bg=C['surface'], fg=tier_clr, font=(FONT,11,'bold')).pack()
        limits = TIER_LIMITS.get(LICENSE.tier, TIER_LIMITS['free'])
        desc = (f"最多 {limits['max_cells'] or '無限'} 張"
                f"  ·  OCR {'✓' if limits['ocr'] else '✗'}"
                f"  ·  去背 {'✓' if limits['matting'] else '✗'}"
                f"  ·  批次 {'✓' if limits['batch'] else '✗'}")
        tk.Label(dlg, text=desc, bg=C['surface'], fg=C['muted'], font=(FONT,9)).pack(pady=4)

        tk.Frame(dlg, bg=C['line'], height=1).pack(fill='x', pady=12)
        tk.Label(dlg, text='輸入授權金鑰以升級方案：',
                 bg=C['surface'], fg=C['ink'], font=(FONT,10)).pack(anchor='w', padx=24)

        key_var = tk.StringVar()
        e = tk.Entry(dlg, textvariable=key_var, font=(FONT,9), width=56,
                     relief='solid', bd=1)
        e.pack(padx=24, pady=6, fill='x')

        msg_lbl = tk.Label(dlg, text='', bg=C['surface'], font=(FONT,9))
        msg_lbl.pack()

        def _activate():
            ok, msg = LICENSE.activate(key_var.get())
            msg_lbl.config(text=msg, fg=C['ok'] if ok else C['err'])
            if ok:
                self.after(1200, lambda: (dlg.destroy(), messagebox.showinfo(
                    '重新啟動', '授權已套用，請重新開啟程式以完整生效。')))

        fab_btn(dlg, '啟動授權', _activate).pack(pady=(4,8), ipadx=16, ipady=4)
        tk.Label(dlg, text='需要金鑰？聯繫 support@fabrica.studio',
                 bg=C['surface'], fg=C['muted'], font=(FONT,8)).pack()


if __name__ == '__main__':
    App().mainloop()

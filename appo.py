import os
import re
import sys
import csv
import json
import hashlib
import threading
import tempfile
import subprocess
import urllib.request
import urllib.error
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import webbrowser
from datetime import datetime

# ---------- Optional drag & drop ----------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ---------- Resource helper ----------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # set by PyInstaller
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ---------- Constants (app + updates) ----------
MAIL_DOMAIN = "tadcinnah.onmicrosoft.com"

APP_NAME = "Office365 Toplu Hesap Sihirbazı"
APP_VERSION = "1.2.0"  # bump when you ship a new build
APP_COPYRIGHT = "© 2025 BTI"

# --- Updater settings (configure these) ---
UPDATE_MANIFEST_URL = "https://example.com/office365_toplu_hesap_sihirbazi/manifest.json"
AUTO_CHECK_UPDATES = True               # check on startup
AUTO_CHECK_DELAY_MS = 2500              # delay after UI loads
PERIODIC_CHECK_HOURS = 0                # set >0 to re-check periodically (e.g., 6)
NETWORK_TIMEOUT_SEC = 20                # http timeout

OFFICE365_HEADERS = [
    'Kullanıcı adı', 'Ad', 'Soyadı', 'Görünen ad',
    'İş unvanı', 'Bölüm', 'İş yeri numarası', 'İş telefonu', 'Cep telefonu', 'Faks',
    'Alternatif e-posta adresi', 'Adres', 'Şehir', 'İl veya ilçe', 'Posta kodu', 'Ülke veya bölge'
]

# ---------- Globals ----------
LAST_DF = None  # keep last loaded dataframe for exports

# ---------- Normalization helpers ----------
_TR_MAP = str.maketrans('çğıöşüÇĞİÖŞÜ', 'cgiosuCGIOSU')

def normalize_text(s: str) -> str:
    """Lowercase, strip, remove Turkish diacritics; collapse inner spaces."""
    if not isinstance(s, str):
        return ""
    s = s.strip().lower()
    s = s.replace('\u2013', '-').replace('\u2014', '-').replace('\u2212', '-')
    s = re.sub(r'\s+', ' ', s)
    s = s.translate(_TR_MAP)
    return s

def normalize_col(col: str) -> str:
    s = normalize_text(col)
    s = s.replace('.', '').replace('_', ' ')
    s = re.sub(r'\s+', ' ', s)
    return s

def col_equals(col: str, target: str) -> bool:
    return normalize_col(col) == normalize_col(target)

def find_column(df: pd.DataFrame, search_for: str):
    target = normalize_col(search_for)
    for col in df.columns:
        if col_equals(col, search_for):
            return col
    for col in df.columns:
        c = normalize_col(col)
        if c.startswith(target) or re.search(rf'\b{re.escape(target)}\b', c):
            return col
    return None

# ---------- Data helpers ----------
def safe_read_excel(path, header=None, nrows=None):
    ext = os.path.splitext(path)[1].lower()
    try:
        return pd.read_excel(path, header=header, nrows=nrows)
    except Exception as e1:
        if ext in ('.xlsx', '.xlsm'):
            try:
                return pd.read_excel(path, header=header, nrows=nrows, engine='openpyxl')
            except Exception as e2:
                raise Exception(f"Excel dosyası okunamadı (.xlsx/.xlsm): {e2}") from e2
        elif ext == '.xls':
            try:
                import xlrd  # noqa: F401
            except Exception:
                raise Exception("'.xls' dosyaları için 'xlrd==1.2.0' gereklidir. Lütfen paketi kurun veya dosyayı .xlsx olarak kaydedin.")
            try:
                return pd.read_excel(path, header=header, nrows=nrows, engine='xlrd')
            except Exception as e3:
                raise Exception(f"Excel dosyası okunamadı (.xls): {e3}") from e3
        else:
            raise Exception(f"Desteklenmeyen uzantı: {ext}") from e1

def smart_read_excel(path):
    preview = safe_read_excel(path, header=None, nrows=5)
    header_keywords = r'\b(ad|adi|adı|soyad|soyadi|soyadı|ad soyad|adi soyadi|adı soyadı)\b'
    for idx, row in preview.iterrows():
        joined = ' '.join([normalize_text(x) for x in row if isinstance(x, str)])
        if re.search(header_keywords, joined):
            return safe_read_excel(path, header=idx)
    return safe_read_excel(path)

def clean_text(text):
    text = str(text)
    text = text.replace('*', '')
    text = re.sub(r'(?i)hk\.?', '', text)
    text = text.translate(_TR_MAP)
    text = re.sub(r'[^a-zA-Z]', '', text)
    return text.lower()

def clean_display(text):
    s = str(text)
    s = re.sub(r'(?i)hk\.?', '', s)
    s = s.replace('*', '')
    return s.strip()

def dedupe_username(username: str, seen: set) -> str:
    if username not in seen:
        seen.add(username)
        return username
    i = 2
    while f"{username}{i}" in seen:
        i += 1
    final = f"{username}{i}"
    seen.add(final)
    return final

def get_preview(file_path):
    global LAST_DF
    df = smart_read_excel(file_path)
    LAST_DF = df

    adi_col = find_column(df, 'Adı') or find_column(df, 'Adi') or find_column(df, 'Ad')
    soyadi_col = find_column(df, 'Soyadı') or find_column(df, 'Soyadi') or find_column(df, 'Soyad')
    adsoyadi_col = find_column(df, 'Adı Soyadı') or find_column(df, 'Adi Soyadi') or find_column(df, 'Ad Soyad')

    if adi_col and soyadi_col:
        mode = 'separate'
    elif adsoyadi_col:
        mode = 'combined'
    else:
        raise Exception('Dosyada uygun isim sütunu yok! (Adı+Soyadı veya Adı Soyadı)')

    preview_records = []
    seen_usernames = set()

    for idx, row in df.iterrows():
        if mode == 'separate':
            parts_ad = str(row.get(adi_col, "")).split()
            parts_soyad = str(row.get(soyadi_col, "")).split()
            cleaned_ad = [clean_text(p) for p in parts_ad if p]
            cleaned_soyad = [clean_text(p) for p in parts_soyad if p]
            base_username = ''.join(cleaned_ad + cleaned_soyad).strip()
            disp_ad = [clean_display(p) for p in parts_ad]
            disp_soyad = [clean_display(p) for p in parts_soyad]
            given = ' '.join(p.upper() for p in disp_ad if p)
            surname = ' '.join(p.upper() for p in disp_soyad if p)
        else:
            full = str(row.get(adsoyadi_col, "")).strip()
            cleaned = [clean_text(p) for p in full.split() if p]
            base_username = ''.join(cleaned).strip()
            disp_full = clean_display(full)
            parts = [p for p in disp_full.split() if p]
            given = parts[0].upper() if parts else ''
            surname = parts[-1].upper() if len(parts) > 1 else ''

        if not base_username or base_username == 'nan':
            base_username = f"user{idx+1}"

        username = dedupe_username(base_username, seen_usernames)
        mail = f"{username}@{MAIL_DOMAIN}"
        preview_records.append([mail, given, surname])

    return preview_records

# ---------- Filename sanitizer ----------
def sanitize_fs_name(name: str) -> str:
    s = (name or "").strip()
    s = re.sub(r'[\x00-\x1f]', '', s)
    s = s.replace('/', '_').replace('\\', '_')
    if os.name == 'nt':
        s = re.sub(r'[<>:"|?*]', '-', s)
        s = s.rstrip(' .')
        if re.fullmatch(r'(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])', s or ''):
            s = f'_{s}'
    return s or "Untitled"

# ---------- Update utilities ----------
def is_frozen() -> bool:
    return getattr(sys, 'frozen', False)

def current_executable_path() -> str:
    if is_frozen():
        return sys.executable
    return os.path.abspath(sys.argv[0])

def version_tuple(v: str):
    return tuple(int(x) for x in re.findall(r'\d+', v)[:4]) or (0,)

def fetch_manifest(url: str, timeout: int):
    req = urllib.request.Request(url, headers={'User-Agent': f'{APP_NAME}/{APP_VERSION}'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        if resp.getcode() != 200:
            raise RuntimeError(f"HTTP {resp.getcode()}")
        data = resp.read()
        return json.loads(data.decode('utf-8', errors='replace'))

def sha256_of_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()

def download_file(url: str, dst_path: str, timeout: int, progress_cb=None):
    req = urllib.request.Request(url, headers={'User-Agent': f'{APP_NAME}/{APP_VERSION}'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = resp.length or 0
        read = 0
        with open(dst_path, 'wb') as out:
            while True:
                chunk = resp.read(1024 * 64)
                if not chunk:
                    break
                out.write(chunk)
                read += len(chunk)
                if progress_cb and total:
                    progress_cb(read, total)

def begin_update_replace_and_restart(new_exe_path: str, target_exe: str):
    """
    On Windows: spawn a .bat that waits, copies new exe over the running one, relaunches, deletes itself.
    On non-Windows: save alongside and prompt user.
    """
    if os.name != 'nt':
        # Non-Windows fallback
        dst = target_exe + ".new"
        try:
            if os.path.abspath(new_exe_path) != os.path.abspath(dst):
                try:
                    os.replace(new_exe_path, dst)
                except Exception:
                    import shutil
                    shutil.copy2(new_exe_path, dst)
            messagebox.showinfo("Güncelleme",
                                f"Güncellenmiş dosya kaydedildi:\n{dst}\n"
                                f"Lütfen mevcut uygulamayı kapatıp bu dosyayı kullanın.")
        finally:
            return

    bat_content = f"""@echo off
setlocal enableextensions
set SRC="{new_exe_path}"
set DST="{target_exe}"
set RETRIES=60

:waitloop
>nul 2>nul (copy /y %SRC% %DST%)
if errorlevel 1 (
  timeout /t 1 /nobreak >nul
  set /a RETRIES-=1
  if %RETRIES% gtr 0 goto waitloop
  echo Kopyalama basarisiz.
  exit /b 1
)

start "" "%DST%"
del "%~f0" 2>nul
"""
    try:
        fd, bat_path = tempfile.mkstemp(prefix="update_", suffix=".bat")
        os.close(fd)
        with open(bat_path, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write(bat_content)
        # Launch the updater and exit current app
        subprocess.Popen(["cmd.exe", "/c", bat_path], close_fds=True, creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
    except Exception as e:
        messagebox.showerror("Güncelleme", f"Güncelleyici başlatılamadı:\n{e}")
        return

    try:
        root.destroy()
    except Exception:
        pass
    sys.exit(0)

def check_updates_async(silent=False):
    """Run in worker thread to avoid blocking UI."""
    def work():
        try:
            manifest = fetch_manifest(UPDATE_MANIFEST_URL, NETWORK_TIMEOUT_SEC)
            latest = str(manifest.get("version", "")).strip()
            url = manifest.get("url")
            sha = manifest.get("sha256", "").lower().strip()
            notes = manifest.get("notes", "")

            if not latest or not url:
                raise RuntimeError("Manifest geçersiz (version/url yok).")

            if version_tuple(latest) <= version_tuple(APP_VERSION):
                if not silent:
                    root.after(0, lambda: messagebox.showinfo("Güncellemeler", "En güncel sürümü kullanıyorsunuz."))
                return

            # Ask user
            def ask_user():
                msg = f"Yeni sürüm bulundu: {latest}\nMevcut: {APP_VERSION}"
                if notes:
                    msg += f"\n\nNotlar:\n{notes}"
                return messagebox.askyesno("Güncelleme Mevcut", msg + "\n\nİndirmek ve yüklemek ister misiniz?")

            proceed = tk.BooleanVar(value=False)
            def prompt():
                proceed.set(ask_user())
            root.after(0, prompt)
            # Wait main thread to set proceed
            while proceed.get() is False and proceed._name is None:
                pass  # Never happens; just to satisfy linter

            if not proceed.get():
                return

            # Progress dialog
            progress_win = {'win': None, 'bar': None, 'label': None}
            def open_progress():
                w = tk.Toplevel(root)
                w.title("Güncelleme İndiriliyor…")
                w.resizable(False, False)
                ttk.Label(w, text="Güncelleme indiriliyor…").pack(padx=12, pady=(12, 4))
                bar = ttk.Progressbar(w, length=300, mode='determinate', maximum=100)
                bar.pack(padx=12, pady=8)
                lbl = ttk.Label(w, text="0%")
                lbl.pack(padx=12, pady=(0, 12))
                progress_win['win'] = w
                progress_win['bar'] = bar
                progress_win['label'] = lbl
                w.grab_set()
                w.protocol("WM_DELETE_WINDOW", lambda: None)  # disable close
            root.after(0, open_progress)

            tmp_fd, tmp_path = tempfile.mkstemp(prefix="update_", suffix=".exe")
            os.close(tmp_fd)

            def on_progress(read, total):
                pct = int(read * 100 / max(1, total))
                def ui():
                    if progress_win['bar']:
                        progress_win['bar']['value'] = pct
                    if progress_win['label']:
                        progress_win['label']['text'] = f"{pct}%"
                root.after(0, ui)

            try:
                download_file(url, tmp_path, NETWORK_TIMEOUT_SEC, on_progress)
            except Exception as e:
                root.after(0, lambda: (progress_win['win'].destroy() if progress_win['win'] else None,
                                       messagebox.showerror("Güncelleme", f"İndirme hatası:\n{e}")))
                return

            # Verify hash if provided
            if sha:
                try:
                    got = sha256_of_file(tmp_path)
                    if got.lower() != sha:
                        root.after(0, lambda: (progress_win['win'].destroy() if progress_win['win'] else None,
                                               messagebox.showerror("Güncelleme", "Dosya bütünlük doğrulaması başarısız (SHA-256 uyuşmuyor).")))
                        return
                except Exception as e:
                    root.after(0, lambda: (progress_win['win'].destroy() if progress_win['win'] else None,
                                           messagebox.showerror("Güncelleme", f"SHA-256 hesaplanamadı:\n{e}")))
                    return

            # Close dialog, replace and restart
            def finish_and_replace():
                if progress_win['win']:
                    progress_win['win'].destroy()
                begin_update_replace_and_restart(tmp_path, current_executable_path())
            root.after(0, finish_and_replace)

        except urllib.error.URLError as e:
            if not silent:
                root.after(0, lambda: messagebox.showerror("Güncellemeler", f"Ağ hatası:\n{e}"))
        except Exception as e:
            if not silent:
                root.after(0, lambda: messagebox.showerror("Güncellemeler", str(e)))

    t = threading.Thread(target=work, daemon=True)
    t.start()

# ---------- UI callbacks (app logic) ----------
def select_file():
    path = filedialog.askopenfilename(title='Excel Dosyası Seç',
                                      filetypes=[('Excel Files', '*.xlsx;*.xls;*.xlsm')])
    if not path:
        return
    entry_file.delete(0, tk.END)
    entry_file.insert(0, path)
    refresh_preview(path)

def refresh_preview(path):
    try:
        for row in preview_tree.get_children():
            preview_tree.delete(row)
        for idx, rec in enumerate(get_preview(path), start=1):
            preview_tree.insert('', 'end', values=(idx, rec[0], rec[1], rec[2]))
        try:
            preview_tree.configure(height=min(20, len(preview_tree.get_children())))
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror('Hata', str(e))

def run_conversion():
    in_path = entry_file.get()
    if not in_path or not os.path.isfile(in_path):
        messagebox.showwarning('Uyarı', 'Lütfen geçerli bir Excel dosyası seçin.')
        return

    raw_class_name = entry_class.get().strip()
    if not raw_class_name:
        messagebox.showwarning('Uyarı', 'Lütfen "Sınıf Adı" bilgisini girin.')
        return

    class_name = sanitize_fs_name(raw_class_name)

    rows = preview_tree.get_children()
    if not rows:
        messagebox.showwarning('Uyarı', 'Önizleme boş olduğu için işlem iptal edildi.')
        return

    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    now = datetime.now()
    date_folder = now.strftime("%d.%m.%Y")
    out_dir = os.path.join(desktop_dir, "Kayıt Bilgileri", date_folder, class_name)
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception as e:
        messagebox.showerror('Hata', f'Klasör oluşturulamadı:\n{out_dir}\n{e}')
        return

    csv_path = os.path.join(out_dir, f"{class_name}.csv")
    tadmail_path = os.path.join(out_dir, "TADMail.txt")
    kisisel_path = os.path.join(out_dir, "KişiselMail.txt")
    telefon_path = os.path.join(out_dir, "Telefon.txt")

    try:
        # Write CSV (ensure "Bölüm" == "LS")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(OFFICE365_HEADERS)
            for item in rows:
                vals = preview_tree.item(item)['values']
                mail, given, surname = vals[1], vals[2], vals[3]
                visible = f"{given} {surname}".strip()

                # First 4 columns
                record = [mail, given, surname, visible]
                # Remaining columns placeholders
                extras = [''] * (len(OFFICE365_HEADERS) - 4)
                # Indexes: 0:İş unvanı, 1:Bölüm, ...
                extras[1] = "LS"  # force Bölüm to LS
                record += extras

                writer.writerow(record)

        # Write TADMail.txt
        mails = [str(preview_tree.item(item)['values'][1]) for item in rows]
        with open(tadmail_path, 'w', encoding='utf-8') as ftxt:
            ftxt.write('\n'.join(mails))

        # Write KişiselMail.txt (intentional trailing comma per line)
        if LAST_DF is not None:
            eposta_col = (
                find_column(LAST_DF, 'E-Posta') or
                find_column(LAST_DF, 'E Posta') or
                find_column(LAST_DF, 'Email') or
                find_column(LAST_DF, 'E-mail')
            )
            if eposta_col:
                personal = [str(x) for x in LAST_DF[eposta_col].dropna().astype(str).tolist()]
                with open(kisisel_path, 'w', encoding='utf-8') as fpers:
                    fpers.write('\n'.join([mail.strip() + ',' for mail in personal]))

        # Write Telefon.txt ("Ad Soyad Telefon")
        if LAST_DF is not None:
            phone_col = (
                find_column(LAST_DF, 'Cep Telefonu') or
                find_column(LAST_DF, 'Cep telefonu') or
                find_column(LAST_DF, 'Telefon')
            )
            adi_col = (
                find_column(LAST_DF, 'Adı') or
                find_column(LAST_DF, 'Adi') or
                find_column(LAST_DF, 'Ad')
            )
            soyadi_col = (
                find_column(LAST_DF, 'Soyadı') or
                find_column(LAST_DF, 'Soyadi') or
                find_column(LAST_DF, 'Soyad')
            )
            adsoyadi_col = (
                find_column(LAST_DF, 'Adı Soyadı') or
                find_column(LAST_DF, 'Adi Soyadi') or
                find_column(LAST_DF, 'Ad Soyad')
            )

            if phone_col:
                lines = []
                for _, row in LAST_DF.iterrows():
                    raw_phone = row.get(phone_col, "")
                    if pd.isna(raw_phone):
                        continue
                    phone_s = str(raw_phone).strip()
                    if not phone_s:
                        continue
                    if phone_s.startswith('+'):
                        phone = '+' + re.sub(r'\D', '', phone_s)
                    else:
                        phone = re.sub(r'\D', '', phone_s)
                    if not phone:
                        continue

                    if adi_col and soyadi_col:
                        name = f"{clean_display(row.get(adi_col, ''))} {clean_display(row.get(soyadi_col, ''))}".strip()
                    elif adsoyadi_col:
                        name = clean_display(row.get(adsoyadi_col, ''))
                    else:
                        name = ""

                    entry = f"{name} {phone}".strip() if name else phone
                    lines.append(entry)

                if lines:
                    with open(telefon_path, 'w', encoding='utf-8') as fphone:
                        fphone.write('\n'.join(lines))

        created_files = [os.path.basename(csv_path), os.path.basename(tadmail_path)]
        if os.path.exists(kisisel_path):
            created_files.append(os.path.basename(kisisel_path))
        if os.path.exists(telefon_path):
            created_files.append(os.path.basename(telefon_path))

        messagebox.showinfo(
            'Başarılı',
            'Kayıt tamamlandı.\n\n'
            f'Klasör: {out_dir}\n- ' + '\n- '.join(created_files)
        )

    except Exception as e:
        messagebox.showerror('Hata', str(e))

def show_about():
    win = tk.Toplevel(root)
    win.title("Hakkında")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    win.update_idletasks()
    w, h = 560, 420
    x = root.winfo_rootx() + (root.winfo_width() // 2) - (w // 2)
    y = root.winfo_rooty() + (root.winfo_height() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")

    wrapper = ttk.Frame(win, padding=16)
    wrapper.pack(fill="both", expand=True)

    ttk.Label(wrapper, text=APP_NAME, font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
    ttk.Label(wrapper, text=f"Sürüm {APP_VERSION}").pack(anchor="w", pady=(2, 0))
    ttk.Label(wrapper, text=APP_COPYRIGHT).pack(anchor="w", pady=(0, 8))

    frame_txt = ttk.Frame(wrapper)
    frame_txt.pack(fill="both", expand=True)

    txt = tk.Text(frame_txt, wrap="word", height=12, borderwidth=1, relief="solid")
    txt.pack(side="left", fill="both", expand=True)
    vsb = ttk.Scrollbar(frame_txt, orient="vertical", command=txt.yview)
    vsb.pack(side="right", fill="y")
    txt.configure(yscrollcommand=vsb.set)

    instructions = (
        "1. TadSOFT’tan hesaplarınızı Excel (.xlsx) formatında dışarı aktarın.\n"
        "2. Dosyayı seçip 'Dönüştür ve Kaydet' butonuna tıklayın.\n"
        "3. Microsoft Yönetim Panelinden Kullanıcılar > Etkin Kullanıcılar kısmına gelin.\n"
        "4. Birden Fazla Kullanıcı ekle butonuna tıklayın.\n"
        "5. Kullanıcı bilgilerini içeren bir CSV'yi karşıya yüklemek için kutucuğunu işaretleyin.\n"
        "6. Mavi Gözat butonuna basıp oluşturduğunuz dosyayı seçin.\n"
        "7. Hiçbir Lisans Atama (Önerilmez) seçeneğini seçerek hesapları açın.\n"
        "8. Hesapları açtıktan sonra Etkin Kullanıcılar bölümüne gidin.\n"
        "9. Filtre Kümesi > Özel Filtre Kümesi > Lisansızlar Filtresini uygulayın.\n"
        "10. Tüm hesapları seçip Parola Sıfırla butonuna basarak parolaları sıfırlayın.\n"
        "11. Tüm hesapları seçip Lisansları Yönet butonuyla lisans atayın.\n"
        "\nGüncellemeler:\n"
        "- Menü > Yardım > Güncellemeleri Denetle… ile elle kontrol edebilir,\n"
        "- Uygulama açılışında otomatik kontrol ettirebilirsiniz."
    )
    txt.insert("1.0", instructions)
    txt.configure(state="disabled")

    btn_row = ttk.Frame(wrapper)
    btn_row.pack(fill="x", pady=(8, 0))
    ttk.Button(btn_row, text="Kapat", command=win.destroy).pack(side="right")
    win.bind("<Escape>", lambda e: win.destroy())

def drop(event):
    files = event.data
    file_path = root.tk.splitlist(files)[0] if isinstance(files, str) else files
    entry_file.delete(0, tk.END)
    entry_file.insert(0, file_path)
    refresh_preview(file_path)

def on_double_click(event):
    item = preview_tree.identify_row(event.y)
    column = preview_tree.identify_column(event.x)
    if not item or not column:
        return
    x, y, width, height = preview_tree.bbox(item, column)
    value = preview_tree.set(item, column)
    entry = tk.Entry(preview_tree)
    entry.insert(0, value)
    entry.focus()

    def save_edit(event=None):
        preview_tree.set(item, column, entry.get())
        entry.destroy()

    entry.bind('<Return>', save_edit)
    entry.bind('<FocusOut>', save_edit)
    entry.place(x=x, y=y, width=width, height=height)

# ---------- Root window ----------
if 'TkinterDnD' in globals() and DND_AVAILABLE:
    root = TkinterDnD.Tk()
else:
    root = tk.Tk()
root.title(APP_NAME)

# Center window
window_width = 560
window_height = 720
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2) - 40
position_left = int(screen_width / 2 - window_width / 2)
root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

# Enable DnD on root/entry
if DND_AVAILABLE:
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop)

# Icon (optional)
icon_path = resource_path('icon.ico')
try:
    root.iconbitmap(icon_path)
except Exception:
    pass
try:
    img_icon = Image.open(icon_path)
    photo_icon = ImageTk.PhotoImage(img_icon)
    root.iconphoto(True, photo_icon)
except Exception:
    pass

# --- Menubar (upper bar) ---
menubar = tk.Menu(root)

help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Yardım…", command=show_about)
help_menu.add_separator()
help_menu.add_command(label="Güncellemeleri Denetle…", command=lambda: check_updates_async(silent=False))
menubar.add_cascade(label="Hakkında", menu=help_menu)

root.config(menu=menubar)
root.bind("<F1>", lambda e: show_about())

# ---------- Top form ----------
frm = ttk.Frame(root, padding=10)
frm.pack(fill='x')

frm.grid_columnconfigure(1, weight=1)
frm.grid_columnconfigure(4, weight=1)

tk.Label(frm, text='Excel Dosyası:').grid(row=0, column=0, sticky='e', pady=6, padx=(0, 6))

entry_file = ttk.Entry(frm)
entry_file.grid(row=0, column=1, sticky='ew', pady=6)
if DND_AVAILABLE:
    entry_file.drop_target_register(DND_FILES)
    entry_file.dnd_bind('<<Drop>>', drop)

ttk.Button(frm, text='Gözat...', command=select_file)\
    .grid(row=0, column=2, padx=4, pady=6, sticky='w')

ttk.Button(frm, text='Etkin Kullanıcılar',
           command=lambda: webbrowser.open('https://admin.cloud.microsoft/#/users'))\
    .grid(row=0, column=5, padx=4, pady=6, sticky='e')

ttk.Button(frm, text='Etkin Ekipler',
           command=lambda: webbrowser.open('https://admin.cloud.microsoft/#/groups'))\
    .grid(row=0, column=6, padx=(4, 0), pady=6, sticky='w')

# ---------- Class Name (Sınıf Adı) input ----------
class_frame = ttk.Frame(root, padding=(10, 0, 10, 6))
class_frame.pack(fill='x')
ttk.Label(class_frame, text='Sınıf Adı:').pack(side='left')
entry_class = ttk.Entry(class_frame)
entry_class.pack(side='left', fill='x', expand=True, padx=(6, 0))

# ---------- Preview ----------
preview_frame = ttk.LabelFrame(root, text='Önizleme', padding=0)
preview_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))

bg_canvas = tk.Canvas(preview_frame, highlightthickness=0, bd=0)
bg_canvas.pack(fill='both', expand=True, side='left')

scrollbar_y = ttk.Scrollbar(preview_frame, orient='vertical')
scrollbar_y.pack(side='right', fill='y')

preview_tree = ttk.Treeview(
    bg_canvas,
    columns=('index', 'mail', 'given', 'surname'),
    show='headings',
    style='Custom.Treeview'
)
style = ttk.Style()
style.configure('Custom.Treeview', background='#f9f9f9', fieldbackground='#f9f9f9')

preview_tree.heading('index', text='#')
preview_tree.column('index', width=40, anchor='center', stretch=False)
preview_tree.heading('mail', text='Mail')
preview_tree.column('mail', anchor='w', stretch=True)
preview_tree.heading('given', text='Ad')
preview_tree.column('given', width=120, anchor='w', stretch=False)
preview_tree.heading('surname', text='Soyadı')
preview_tree.column('surname', width=120, anchor='w', stretch=False)

preview_tree.configure(yscrollcommand=scrollbar_y.set)
scrollbar_y.configure(command=preview_tree.yview)

tree_window_id = bg_canvas.create_window((0, 0), window=preview_tree, anchor='nw')

# --- Watermark ---
_logo_original_rgba = None
_logo_tk = None
logo_image_id = None

def load_logo_rgba():
    global _logo_original_rgba
    logo_file = resource_path("logo.png")
    if not os.path.exists(logo_file):
        _logo_original_rgba = None
        return
    try:
        img = Image.open(logo_file).convert("RGBA")
        _logo_original_rgba = img
    except Exception:
        _logo_original_rgba = None

def draw_or_update_watermark():
    global _logo_tk, logo_image_id, _logo_original_rgba
    if _logo_original_rgba is None:
        return
    cw = max(1, bg_canvas.winfo_width())
    ch = max(1, bg_canvas.winfo_height())
    target = int(min(cw, ch) * 0.45)
    if target < 10:
        return
    w, h = _logo_original_rgba.size
    if w >= h:
        tw = target
        th = int(h * (target / w))
    else:
        th = target
        tw = int(w * (target / h))
    wm = _logo_original_rgba.resize((max(1, tw), max(1, th)), Image.LANCZOS).copy()
    wm.putalpha(90)
    _logo_tk = ImageTk.PhotoImage(wm)
    cx = cw // 2
    cy = ch // 2
    if logo_image_id is None:
        logo_image_id = bg_canvas.create_image(cx, cy, image=_logo_tk)
    else:
        bg_canvas.itemconfigure(logo_image_id, image=_logo_tk)
        bg_canvas.coords(logo_image_id, cx, cy)

def on_canvas_configure(event):
    bg_canvas.itemconfigure(tree_window_id, width=event.width, height=event.height)
    draw_or_update_watermark()

def on_tree_configure(event):
    bg_canvas.config(scrollregion=bg_canvas.bbox('all'))

bg_canvas.bind('<Configure>', on_canvas_configure)
preview_tree.bind('<Configure>', on_tree_configure)

preview_tree.bind('<MouseWheel>', lambda e: preview_tree.yview_scroll(int(-1 * (e.delta / 120)), 'units'))
preview_tree.bind('<Double-1>', on_double_click)

# Load logo once (original), then first draw
load_logo_rgba()
root.after(50, draw_or_update_watermark)

# ---------- Footer ----------
ttk.Button(root, text='Dönüştür ve Kaydet', command=run_conversion, width=30).pack(pady=(6, 20))

# ---------- Auto-update hooks ----------
def schedule_periodic_check():
    if PERIODIC_CHECK_HOURS and PERIODIC_CHECK_HOURS > 0:
        ms = int(PERIODIC_CHECK_HOURS * 3600 * 1000)
        root.after(ms, lambda: (check_updates_async(silent=True), schedule_periodic_check()))

if AUTO_CHECK_UPDATES:
    root.after(AUTO_CHECK_DELAY_MS, lambda: check_updates_async(silent=True))
    root.after(AUTO_CHECK_DELAY_MS + 100, schedule_periodic_check)

root.mainloop()

import os, sys, re, csv, time, shutil, subprocess, threading, glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

NOWIN = 0x08000000 if os.name == "nt" else 0
PIPE, STDOUT, DEVNULL = subprocess.PIPE, subprocess.STDOUT, subprocess.DEVNULL

VIDEO = {"dav", "avi", "mkv", "mov", "wmv", "flv", "mp4", "264", "h264", "265", "mpg",
         "mpeg", "ts", "webm", "3gp", "m4v", "vob", "asf", "mp3", "wav", "aac", "flac", "m4a", "ogg"}
IMAGE = {"png", "jpg", "jpeg", "bmp", "gif", "tif", "tiff", "webp", "ico", "ppm", "tga"}
WORD_SRC = {"doc", "docx", "rtf", "txt", "odt", "html", "htm"}
EXCEL_SRC = {"xls", "xlsx", "csv", "ods"}
PPT_SRC = {"ppt", "pptx", "odp"}
OFFICE = WORD_SRC | EXCEL_SRC | PPT_SRC

VIDEO_OUT = ["mp4", "avi", "mkv", "mov", "wmv", "webm", "gif", "mp3", "wav"]
IMAGE_OUT = ["png", "jpg", "bmp", "gif", "tiff", "webp", "ico", "pdf"]
PDF_OUT = ["docx", "png", "jpg", "txt"]
WORD_OUT = ["pdf", "docx", "doc", "odt", "rtf", "txt", "html"]
EXCEL_OUT = ["pdf", "xlsx", "xls", "csv", "ods", "html"]
PPT_OUT = ["pdf", "pptx", "ppt", "odp"]

# ---------- Renkler ----------
BG, CARD, CARD2 = "#0e1016", "#161923", "#1e2230"
BORDER, TRACK = "#262b3b", "#2a2f42"
ACCENT, ACCENT_H, ACCENT_D = "#7c5cff", "#9279ff", "#3a3566"
TEXT, MUTED = "#e8eaf2", "#8a90a6"
OK, ERR, WARN = "#34d399", "#f87171", "#fbbf24"


class Cancelled(Exception):
    pass


class Ctx:
    """İlerleme bildirimi ve iptal için ortak nesne."""
    def __init__(self):
        self.cb = lambda f: None
        self.cancel = False
        self.proc = None
        self.engine = ""

    def report(self, f):
        self.cb(f)


# ---------- Araç bulma ----------
def base_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))


def find_ffmpeg():
    for p in (os.path.join(base_dir(), "ffmpeg.exe"),
              os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "ffmpeg.exe")):
        if os.path.isfile(p):
            return p
    return shutil.which("ffmpeg")


def find_soffice():
    p = shutil.which("soffice")
    if p:
        return p
    for pat in (r"C:\Program Files\LibreOffice*\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice*\program\soffice.exe"):
        m = glob.glob(pat)
        if m:
            return m[0]
    return None


def run(cmd):
    return subprocess.run(cmd, capture_output=True, text=True, errors="ignore",
                          stdin=DEVNULL, creationflags=NOWIN)


def run_cancellable(cmd, ctx):
    p = subprocess.Popen(cmd, stdout=PIPE, stderr=PIPE, stdin=DEVNULL, text=True,
                         errors="ignore", creationflags=NOWIN)
    ctx.proc = p
    res = []
    t = threading.Thread(target=lambda: res.append(p.communicate()), daemon=True)
    t.start()
    while t.is_alive():
        if ctx.cancel:
            p.kill()
            raise Cancelled()
        time.sleep(0.2)
    return p.returncode, (res[0][1] if res else "")


# ---------- Dönüştürücüler ----------
def probe_duration(ff, pre, src):
    r = run([ff, "-hide_banner", *pre, "-i", src])
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", r.stderr)
    if not m:
        return None
    d = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3))
    return d if d > 0 else None


def conv_video(src, out, target, ctx):
    ff = find_ffmpeg()
    if not ff:
        raise RuntimeError("ffmpeg.exe bulunamadı (programın yanına koyun).")
    if target in ("mp3", "wav"):
        args = ["-vn"]
    elif target == "gif":
        args = ["-vf", "fps=10,scale=480:-1"]
    elif target == "webm":
        args = ["-c:v", "libvpx-vp9", "-crf", "32", "-b:v", "0", "-c:a", "libopus"]
    elif target == "wmv":
        args = ["-c:v", "wmv2", "-c:a", "wmav2"]
    else:
        args = ["-c:v", "libx264", "-preset", "fast", "-crf", "23", "-pix_fmt", "yuv420p"]
        args += ["-c:a", "mp3" if target == "avi" else "aac"]

    raw = src.lower().endswith((".dav", ".264", ".h264", ".265"))
    tries = [[], ["-f", "dhav"], ["-f", "h264"], ["-f", "hevc"]] if raw else [[]]
    err = ""
    for pre in tries:
        if ctx.cancel:
            raise Cancelled()
        dur = probe_duration(ff, pre, src)
        ctx.report(0.0 if dur else None)
        cmd = [ff, "-y", "-hide_banner", "-loglevel", "error", "-nostats",
               "-progress", "pipe:1", *pre, "-i", src, *args, out]
        p = subprocess.Popen(cmd, stdout=PIPE, stderr=STDOUT, stdin=DEVNULL, text=True,
                             errors="ignore", creationflags=NOWIN)
        ctx.proc = p
        errs = []
        for line in p.stdout:
            line = line.strip()
            m = re.match(r"out_time_(?:us|ms)=(\d+)", line)
            if m:
                if dur:
                    ctx.report(min(int(m.group(1)) / 1e6 / dur, 0.99))
            elif line and not re.match(r"^\w+=", line):
                errs.append(line)
        p.wait()
        if ctx.cancel:
            try:
                os.remove(out)
            except OSError:
                pass
            raise Cancelled()
        if p.returncode == 0 and os.path.isfile(out) and os.path.getsize(out) > 0:
            ctx.report(1.0)
            return
        err = " ".join(errs)[-300:]
    raise RuntimeError(err or "ffmpeg dönüştüremedi")


def conv_image(src, out, target, ctx):
    from PIL import Image
    ctx.report(0.3)
    im = Image.open(src)
    if target in ("jpg", "pdf") and im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGB")
    fmt = {"jpg": "JPEG", "tiff": "TIFF", "pdf": "PDF"}.get(target, target.upper())
    im.save(out, fmt)
    ctx.report(1.0)


def conv_pdf(src, out, target, outdir, stem, ctx):
    if target == "docx":
        ctx.report(None)
        from pdf2docx import Converter
        cv = Converter(src)
        cv.convert(out)
        cv.close()
        ctx.report(1.0)
        return
    import fitz
    doc = fitz.open(src)
    total = max(len(doc), 1)
    if target == "txt":
        with open(out, "w", encoding="utf-8") as f:
            for i, page in enumerate(doc):
                if ctx.cancel:
                    raise Cancelled()
                f.write(page.get_text() + "\n")
                ctx.report((i + 1) / total)
    else:
        for i, page in enumerate(doc, 1):
            if ctx.cancel:
                raise Cancelled()
            page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(
                os.path.join(outdir, f"{stem}_sayfa{i}.{target}"))
            ctx.report(i / total)
    doc.close()


WORD_FMT = {"docx": 16, "doc": 0, "rtf": 6, "html": 8, "odt": 23}
EXCEL_FMT = {"xlsx": 51, "xls": 56, "csv": 62, "html": 44, "ods": 60}
PPT_FMT = {"pptx": 24, "ppt": 1, "odp": 35}


def ms_office(src, ext, out, target, ctx):
    """Microsoft Office (COM) ile dönüştürür. Office kurulu değilse False döner."""
    if os.name != "nt":
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    if ext in WORD_SRC:
        progid = "Word.Application"
    elif ext in EXCEL_SRC:
        progid = "Excel.Application"
    else:
        progid = "PowerPoint.Application"
    src, out = os.path.abspath(src), os.path.abspath(out)
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.DispatchEx(progid)  # kullanıcının açık Office'ine dokunmaz
        except Exception:
            return False  # Office kurulu değil
        try:
            if progid.startswith("Word"):
                app.Visible = False
                app.DisplayAlerts = 0
                doc = app.Documents.Open(src, ConfirmConversions=False, ReadOnly=True,
                                         AddToRecentFiles=False)
                try:
                    if target == "pdf":
                        doc.ExportAsFixedFormat(out, 17)
                    elif target == "txt":
                        doc.SaveAs2(out, FileFormat=2, Encoding=65001)
                    else:
                        doc.SaveAs2(out, FileFormat=WORD_FMT[target])
                finally:
                    doc.Close(False)
            elif progid.startswith("Excel"):
                app.Visible = False
                app.DisplayAlerts = False
                wb = app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True)
                try:
                    if target == "pdf":
                        wb.ExportAsFixedFormat(0, out)
                    else:
                        wb.SaveAs(out, FileFormat=EXCEL_FMT[target])
                finally:
                    wb.Close(False)
            else:
                pres = app.Presentations.Open(src, ReadOnly=True, Untitled=False, WithWindow=False)
                try:
                    pres.SaveAs(out, 32 if target == "pdf" else PPT_FMT[target])
                finally:
                    pres.Close()
        finally:
            try:
                app.Quit()
            except Exception:
                pass
            del app
    finally:
        pythoncom.CoUninitialize()
    return os.path.isfile(out)


def read_text(path):
    for enc in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return ""


def conv_builtin(src, ext, out, target):
    """Hiçbir program kurulu değilse: basit dönüşümler (kalite sınırlı)."""
    if ext == "xlsx" and target == "csv":
        from openpyxl import load_workbook
        wb = load_workbook(src, read_only=True, data_only=True)
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")  # Türkçe Excel ; ayracını bekler
            for row in wb.worksheets[0].iter_rows(values_only=True):
                w.writerow(["" if c is None else c for c in row])
        wb.close()
    elif ext == "csv" and target == "xlsx":
        from openpyxl import Workbook
        text = read_text(src)
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        wb = Workbook()
        ws = wb.active
        for row in csv.reader(text.splitlines(), dialect):
            ws.append(row)
        wb.save(out)
    elif ext == "docx" and target == "txt":
        from docx import Document
        d = Document(src)
        lines = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                lines.append("\t".join(c.text for c in r.cells))
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    elif ext == "txt" and target == "docx":
        from docx import Document
        d = Document()
        for line in read_text(src).splitlines():
            d.add_paragraph(line)
        d.save(out)
    else:
        return False
    return True


def conv_office(src, ext, out, target, outdir, ctx):
    """Sırayla dener: 1) Microsoft Office  2) LibreOffice  3) yerleşik basit dönüştürücü."""
    ctx.report(None)
    ms_err = None
    try:
        if ms_office(src, ext, out, target, ctx):
            ctx.engine = "Microsoft Office"
            ctx.report(1.0)
            return
    except Cancelled:
        raise
    except Exception as e:
        ms_err = str(e).replace("\n", " ")[:200]
    so = find_soffice()
    if so:
        flt = "txt:Text" if target == "txt" else target
        code, err = run_cancellable([so, "--headless", "--convert-to", flt, "--outdir", outdir, src], ctx)
        if code != 0:
            raise RuntimeError(err.strip()[-300:] or "LibreOffice hatası")
        ctx.engine = "LibreOffice"
        ctx.report(1.0)
        return
    if conv_builtin(src, ext, out, target):
        ctx.engine = "yerleşik, sınırlı kalite"
        ctx.report(1.0)
        return
    msg = "Bu dönüşüm için Microsoft Office veya LibreOffice kurulu olmalı."
    if ms_err:
        msg = f"Microsoft Office hatası: {ms_err}. " + msg
    raise RuntimeError(msg)


def convert(src, target, outdir, ctx):
    ext = os.path.splitext(src)[1][1:].lower()
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(outdir, f"{stem}.{target}")
    if os.path.abspath(out) == os.path.abspath(src):
        out = os.path.join(outdir, f"{stem}_donusturuldu.{target}")
    if ext in VIDEO:
        conv_video(src, out, target, ctx)
    elif ext in IMAGE:
        conv_image(src, out, target, ctx)
    elif ext == "pdf":
        conv_pdf(src, out, target, outdir, stem, ctx)
    elif ext in OFFICE:
        conv_office(src, ext, out, target, outdir, ctx)
    else:
        raise RuntimeError(f"Desteklenmeyen biçim: .{ext}")
    return out


def targets_for(ext):
    if ext in VIDEO: return VIDEO_OUT
    if ext in IMAGE: return IMAGE_OUT
    if ext == "pdf": return PDF_OUT
    if ext in WORD_SRC: return [t for t in WORD_OUT if t != ext]
    if ext in EXCEL_SRC: return [t for t in EXCEL_OUT if t != ext]
    if ext in PPT_SRC: return [t for t in PPT_OUT if t != ext]
    return []


def fmt_t(s):
    s = int(max(s, 0))
    h, r = divmod(s, 3600)
    m, sec = divmod(r, 60)
    return f"{h}:{m:02d}:{sec:02d}" if h else f"{m:02d}:{sec:02d}"


# ---------- Arayüz parçaları ----------
class Btn(tk.Label):
    def __init__(self, master, text, command, bg=CARD2, hover="#2b3144", fg=TEXT,
                 font=("Segoe UI", 10, "bold"), padx=16, pady=8):
        super().__init__(master, text=text, bg=bg, fg=fg, font=font, padx=padx, pady=pady, cursor="hand2")
        self.cmd, self.base, self.hov, self.fgc = command, bg, hover, fg
        self.on = True
        self.bind("<Enter>", lambda e: self.on and self.config(bg=self.hov))
        self.bind("<Leave>", lambda e: self.on and self.config(bg=self.base))
        self.bind("<Button-1>", lambda e: self.on and self.cmd())

    def enable(self, on):
        self.on = on
        self.config(bg=self.base if on else "#232736", fg=self.fgc if on else "#5b6176",
                    cursor="hand2" if on else "arrow")


class Bar(tk.Canvas):
    def __init__(self, master):
        super().__init__(master, height=16, bg=CARD, highlightthickness=0)
        self.value, self.indet, self.pos = 0.0, False, 0.0
        self.bind("<Configure>", lambda e: self.draw())

    def set(self, v):
        self.indet = False
        self.value = max(0.0, min(1.0, v))
        self.draw()

    def start_indet(self):
        if not self.indet:
            self.indet = True
            self._anim()

    def _anim(self):
        if not self.indet:
            return
        self.pos = (self.pos + 0.018) % 1.3
        self.draw()
        self.after(20, self._anim)

    def draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        r = h / 2
        if w < 4 * r:
            return
        self.create_line(r, r, w - r, r, width=h, capstyle="round", fill=TRACK)
        span = w - 2 * r
        if self.indet:
            x1 = r + span * (self.pos - 0.3)
            x2 = r + span * self.pos
            x1, x2 = max(x1, r), min(x2, w - r)
            if x2 > x1:
                self.create_line(x1, r, x2, r, width=h, capstyle="round", fill=ACCENT)
        elif self.value > 0:
            self.create_line(r, r, r + max(span * self.value, 0.1), r, width=h, capstyle="round", fill=ACCENT)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Evrensel Dönüştürücü")
        self.geometry("780x780")
        self.minsize(700, 700)
        self.configure(bg=BG)
        self.files, self.ctx, self.running = [], None, False
        self.indet_mode, self.t0, self.last_out = False, 0.0, None
        self.style_widgets()
        self.build()
        self.after(100, self.dark_titlebar)
        self.bind("<Delete>", lambda e: self.remove_selected())

    def dark_titlebar(self):
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            for attr in (20, 19):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(ctypes.c_int(1)), 4)
        except Exception:
            pass

    def style_widgets(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("TCombobox", fieldbackground=CARD2, background=CARD2, foreground=TEXT,
                     arrowcolor=TEXT, bordercolor=BORDER, lightcolor=CARD2, darkcolor=CARD2, padding=6)
        st.map("TCombobox", fieldbackground=[("readonly", CARD2)], foreground=[("readonly", TEXT)],
               selectbackground=[("readonly", CARD2)], selectforeground=[("readonly", TEXT)])
        st.configure("Vertical.TScrollbar", background=CARD2, troughcolor=CARD, bordercolor=CARD,
                     arrowcolor=MUTED, lightcolor=CARD2, darkcolor=CARD2)
        self.option_add("*TCombobox*Listbox.background", CARD2)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

    def card(self, expand=False):
        outer = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        outer.pack(fill="both" if expand else "x", expand=expand, padx=20, pady=6)
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="both", expand=True, padx=16, pady=14)
        return inner

    def small(self, parent, text):
        return tk.Label(parent, text=text, bg=CARD, fg=MUTED, font=("Segoe UI", 8, "bold"))

    def build(self):
        tk.Label(self, text="Evrensel Dönüştürücü", bg=BG, fg=TEXT,
                 font=("Segoe UI", 21, "bold")).pack(anchor="w", padx=22, pady=(18, 0))
        tk.Label(self, text="DAV  •  MP4  •  PDF  •  Word  •  Excel  •  Görüntü", bg=BG, fg=MUTED,
                 font=("Segoe UI", 10)).pack(anchor="w", padx=22, pady=(0, 8))

        # Dosyalar
        c = self.card()
        head = tk.Frame(c, bg=CARD); head.pack(fill="x")
        self.small(head, "DOSYALAR").pack(side="left")
        self.count = tk.Label(head, text="0 dosya", bg=CARD, fg=MUTED, font=("Segoe UI", 9))
        self.count.pack(side="right")
        lw = tk.Frame(c, bg=CARD); lw.pack(fill="x", pady=(8, 8))
        self.lb = tk.Listbox(lw, height=6, bg=CARD2, fg=TEXT, selectbackground="#2d3350",
                             selectforeground=TEXT, borderwidth=0, highlightthickness=0,
                             font=("Segoe UI", 10), activestyle="none")
        sb = ttk.Scrollbar(lw, orient="vertical", command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")
        row = tk.Frame(c, bg=CARD); row.pack(fill="x")
        Btn(row, "＋  Dosya Ekle", self.add).pack(side="left")
        Btn(row, "Seçileni Sil", self.remove_selected).pack(side="left", padx=8)
        Btn(row, "Temizle", self.clear).pack(side="left")

        # Ayarlar
        c = self.card()
        self.small(c, "AYARLAR").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        tk.Label(c, text="Hedef biçim", bg=CARD, fg=TEXT, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w")
        self.fmt = ttk.Combobox(c, state="readonly", width=12, font=("Segoe UI", 10))
        self.fmt.grid(row=1, column=1, sticky="w", padx=12, pady=4)
        tk.Label(c, text="Kayıt klasörü", bg=CARD, fg=TEXT, font=("Segoe UI", 10)).grid(row=2, column=0, sticky="w")
        self.outvar = tk.StringVar(value="(kaynakla aynı klasör)")
        tk.Entry(c, textvariable=self.outvar, bg=CARD2, fg=TEXT, insertbackground=TEXT, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                 font=("Segoe UI", 10)).grid(row=2, column=1, sticky="ew", padx=12, pady=4, ipady=5)
        Btn(c, "Seç", self.pick, padx=14, pady=5).grid(row=2, column=2)
        c.columnconfigure(1, weight=1)

        # İlerleme
        c = self.card()
        top = tk.Frame(c, bg=CARD); top.pack(fill="x")
        self.status = tk.Label(top, text="Hazır", bg=CARD, fg=TEXT, font=("Segoe UI", 10), anchor="w")
        self.status.pack(side="left", fill="x", expand=True)
        self.pct = tk.Label(top, text="%0", bg=CARD, fg=ACCENT_H, font=("Segoe UI", 18, "bold"))
        self.pct.pack(side="right")
        self.bar = Bar(c); self.bar.pack(fill="x", pady=(8, 6))
        self.eta = tk.Label(c, text="Dosya ekleyip DÖNÜŞTÜR'e basın", bg=CARD, fg=MUTED,
                            font=("Segoe UI", 9), anchor="w")
        self.eta.pack(fill="x")

        # Düğmeler
        row = tk.Frame(self, bg=BG); row.pack(fill="x", padx=20, pady=6)
        self.go = Btn(row, "▶   DÖNÜŞTÜR", self.start, bg=ACCENT, hover=ACCENT_H, padx=34, pady=11,
                      font=("Segoe UI", 11, "bold"))
        self.go.pack(side="left")
        self.stop = Btn(row, "İptal", self.cancel, bg="#3a1f2a", hover="#55293a", fg=ERR, pady=11)
        self.stop.pack(side="left", padx=8); self.stop.enable(False)
        self.openb = Btn(row, "Klasörü Aç", self.open_out, pady=11)
        self.openb.pack(side="right"); self.openb.enable(False)

        # Günlük
        c = self.card(expand=True)
        self.small(c, "GÜNLÜK").pack(anchor="w", pady=(0, 6))
        self.log = tk.Text(c, height=6, bg=CARD2, fg=MUTED, relief="flat", borderwidth=0,
                           font=("Consolas", 9), state="disabled", highlightthickness=0, padx=8, pady=6)
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("ok", foreground=OK)
        self.log.tag_config("err", foreground=ERR)
        self.log.tag_config("info", foreground=TEXT)

    # ----- yardımcılar -----
    def say(self, t, tag="info"):
        self.log.config(state="normal"); self.log.insert("end", t + "\n", tag)
        self.log.see("end"); self.log.config(state="disabled")

    def set_item(self, i, icon, color=TEXT):
        self.lb.delete(i)
        self.lb.insert(i, f"  {icon}  {os.path.basename(self.files[i])}")
        self.lb.itemconfig(i, fg=color)

    def update_count(self):
        self.count.config(text=f"{len(self.files)} dosya")

    def add(self):
        if self.running:
            return
        for f in filedialog.askopenfilenames(title="Dosya seç"):
            if f not in self.files:
                self.files.append(f)
                self.lb.insert("end", "")
                self.set_item(len(self.files) - 1, "●", MUTED)
        if self.files:
            t = targets_for(os.path.splitext(self.files[0])[1][1:].lower())
            self.fmt["values"] = t
            if t and self.fmt.get() not in t:
                self.fmt.set(t[0])
        self.update_count()

    def remove_selected(self):
        if self.running:
            return
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i); del self.files[i]
        self.update_count()

    def clear(self):
        if self.running:
            return
        self.files.clear(); self.lb.delete(0, "end"); self.update_count()

    def pick(self):
        d = filedialog.askdirectory()
        if d:
            self.outvar.set(d)

    def open_out(self):
        if self.last_out and os.path.isdir(self.last_out):
            try:
                os.startfile(self.last_out)
            except Exception:
                pass

    def cancel(self):
        if self.ctx and self.running:
            self.ctx.cancel = True
            self.status.config(text="İptal ediliyor… (çalışan dosya bitince durur)")
            try:
                if self.ctx.proc:
                    self.ctx.proc.kill()
            except Exception:
                pass

    # ----- çalıştırma -----
    def start(self):
        if self.running:
            return
        if not self.files or not self.fmt.get():
            messagebox.showinfo("Bilgi", "Önce dosya ekleyin ve hedef biçimi seçin.")
            return
        self.running = True
        self.ctx = Ctx()
        self.t0 = time.time()
        self.go.enable(False); self.stop.enable(True); self.openb.enable(False)
        for i in range(len(self.files)):
            self.set_item(i, "●", MUTED)
        self.bar.set(0); self.pct.config(text="%0")
        self.eta.config(text="Hesaplanıyor…")
        self.tick()
        threading.Thread(target=self.work, daemon=True).start()

    def tick(self):
        if not self.running:
            return
        if self.indet_mode:
            self.eta.config(text=f"İşleniyor…   Geçen süre: {fmt_t(time.time() - self.t0)}")
        self.after(500, self.tick)

    def on_progress(self, i, n, frac):
        if not self.running:
            return
        name = os.path.basename(self.files[i])
        self.status.config(text=f"Dosya {i + 1}/{n}  ·  {name}")
        elapsed = time.time() - self.t0
        if frac is None:
            self.indet_mode = True
            self.bar.start_indet()
            self.pct.config(text=f"{i + 1}/{n}")
            return
        self.indet_mode = False
        overall = (i + frac) / n
        self.bar.set(overall)
        self.pct.config(text=f"%{int(overall * 100)}")
        if overall > 0.01 and elapsed > 1.5:
            left = elapsed * (1 - overall) / overall
            self.eta.config(text=f"Tahmini kalan süre: {fmt_t(left)}   ·   Geçen: {fmt_t(elapsed)}")
        else:
            self.eta.config(text="Kalan süre hesaplanıyor…")

    def work(self):
        target, n, ctx = self.fmt.get(), len(self.files), self.ctx
        ok = bad = 0
        cancelled = False
        for i, f in enumerate(self.files):
            if ctx.cancel:
                cancelled = True
                break
            od = self.outvar.get()
            if not os.path.isdir(od):
                od = os.path.dirname(f)
            self.last_out = od
            ctx.cb = lambda fr, i=i: self.after(0, self.on_progress, i, n, fr)
            self.after(0, self.set_item, i, "▶", ACCENT_H)
            self.after(0, self.on_progress, i, n, 0.0)
            self.after(0, self.say, f"▶ {os.path.basename(f)}  →  .{target}")
            try:
                ctx.engine, ctx.proc = "", None
                o = convert(f, target, od, ctx)
                ok += 1
                via = f"  [{ctx.engine}]" if ctx.engine else ""
                self.after(0, self.set_item, i, "✔", OK)
                self.after(0, self.say, f"   ✔ Tamam: {o}{via}", "ok")
            except Cancelled:
                cancelled = True
                self.after(0, self.set_item, i, "●", MUTED)
                break
            except Exception as e:
                bad += 1
                self.after(0, self.set_item, i, "✖", ERR)
                self.after(0, self.say, f"   ✖ Hata: {e}", "err")
        self.after(0, self.finish, ok, bad, cancelled, n)

    def finish(self, ok, bad, cancelled, n):
        self.running = False
        self.indet_mode = False
        self.go.enable(True); self.stop.enable(False); self.openb.enable(bool(self.last_out))
        total = fmt_t(time.time() - self.t0)
        if cancelled:
            self.status.config(text="İptal edildi")
            self.eta.config(text=f"Kullanıcı tarafından durduruldu · {total}")
            self.bar.set(self.bar.value)
            self.say("— İptal edildi —", "err")
        else:
            self.bar.set(1.0 if bad == 0 else self.bar.value)
            self.pct.config(text="%100" if bad == 0 else "—")
            self.status.config(text=f"Tamamlandı: {ok} başarılı" + (f", {bad} hatalı" if bad else ""))
            self.eta.config(text=f"Toplam süre: {total}")
            self.say("— Bitti —", "ok" if bad == 0 else "err")


if __name__ == "__main__":
    App().mainloop()

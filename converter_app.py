"""
万能 PDF 转换器 v1.8.2
功能：Word/Excel/PPT/图片/CAD 批量转 PDF
特性：代码深度精简、AutoCAD 引擎直连、比例重对齐修复、原生拖拽
"""

import os, threading, glob, winreg, subprocess, tempfile, queue, platform, ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

try:
    import customtkinter as ctk
    from PIL import Image
    import comtypes.client
    import fitz
except ImportError as e:
    print(f"Missing libraries: {e}")

def _find_accore():
    for p in [r"C:\Program Files\Autodesk\AutoCAD*", r"D:\Program Files\Autodesk\AutoCAD*",
              r"E:\Program Files\Autodesk\AutoCAD*", r"F:\cad\AutoCAD*"]:
        for f in glob.glob(p):
            exe = os.path.join(f, "accoreconsole.exe")
            if os.path.exists(exe): return exe
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"AutoCAD.Drawing\shell\open\command") as k:
            cmd = winreg.QueryValue(k, None)
            d = os.path.dirname(cmd.split('"')[1] if '"' in cmd else cmd.split()[0])
            exe = os.path.join(d, "accoreconsole.exe")
            if os.path.exists(exe): return exe
    except: pass
    return None

def _register_drop(hwnd, callback):
    u32, s32 = ctypes.windll.user32, ctypes.windll.shell32
    HDROP = wintypes.HANDLE
    s32.DragQueryFileW.argtypes = [HDROP, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    s32.DragQueryFileW.restype = wintypes.UINT
    s32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]

    is64 = platform.architecture()[0] == "64bit"
    PTR = ctypes.c_void_p if is64 else ctypes.c_uint
    SetWLP = u32.SetWindowLongPtrW if is64 else u32.SetWindowLongW
    GetWLP = u32.GetWindowLongPtrW if is64 else u32.GetWindowLongW
    SetWLP.argtypes = [wintypes.HWND, ctypes.c_int, PTR]
    SetWLP.restype = PTR
    GetWLP.argtypes = [wintypes.HWND, ctypes.c_int]
    GetWLP.restype = PTR
    
    WNDPROC = ctypes.WINFUNCTYPE(PTR, wintypes.HWND, wintypes.UINT, PTR, PTR)
    CallWP = u32.CallWindowProcW
    CallWP.argtypes = [PTR, wintypes.HWND, wintypes.UINT, PTR, PTR]
    CallWP.restype = PTR
    
    old = GetWLP(hwnd, -4)
    def handler(h, m, wp, lp):
        if m == 0x0233:
            try:
                hd = ctypes.cast(wp, HDROP)
                n = s32.DragQueryFileW(hd, 0xFFFFFFFF, None, 0)
                fs = []
                for i in range(n):
                    sz = s32.DragQueryFileW(hd, i, None, 0)
                    b = ctypes.create_unicode_buffer(sz + 1)
                    s32.DragQueryFileW(hd, i, b, sz + 1)
                    fs.append(b.value)
                s32.DragFinish(hd)
                callback(fs)
            except: pass
            return 0
        return CallWP(old, h, m, wp, lp)
    
    hook = WNDPROC(handler)
    s32.DragAcceptFiles(hwnd, True)
    SetWLP(hwnd, -4, ctypes.cast(hook, PTR))
    return hook

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PDFConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("万能 PDF 转换器 v1.8.2")
        self.geometry("800x700")
        self.input_paths = set()
        self.output_dir = ""
        self.is_busy = False
        self.accore = _find_accore()
        self.cad_center = tk.BooleanVar(value=True)
        self.cad_border = tk.BooleanVar(value=True)
        self.margin_top = tk.StringVar(value="3.5")
        self.margin_bottom = tk.StringVar(value="3.5")
        self.margin_left = tk.StringVar(value="3.5")
        self.margin_right = tk.StringVar(value="3.5")

        self._build_ui()
        self._dq = queue.Queue()
        self.after(100, self._poll_drops)
        self._hook = _register_drop(self.winfo_id(), self._dq.put)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)
        ctk.CTkLabel(self, text="万能 PDF 转换器", font=("Arial", 24, "bold")).grid(row=0, pady=20)
        
        bf = ctk.CTkFrame(self); bf.grid(row=1, padx=20, pady=10, sticky="ew")
        bf.grid_columnconfigure((0, 1, 2), weight=1)
        ctk.CTkButton(bf, text="选择文件", command=self._pick_f).grid(row=0, column=0, padx=10, pady=20)
        ctk.CTkButton(bf, text="选择目录", command=self._pick_d).grid(row=0, column=1, padx=10, pady=20)
        self._btn_out = ctk.CTkButton(bf, text="保存位置", command=self._pick_o, fg_color="gray")
        self._btn_out.grid(row=0, column=2, padx=10, pady=20)

        self._lbl = ctk.CTkLabel(self, text="当前待处理：0 个路径", text_color="#3498db")
        self._lbl.grid(row=2, pady=5)

        cf = ctk.CTkFrame(self); cf.grid(row=3, padx=20, pady=5, sticky="ew")
        ctk.CTkLabel(cf, text="CAD 选项:").pack(side="left", padx=10)
        ctk.CTkCheckBox(cf, text="居中", variable=self.cad_center).pack(side="left", padx=10)
        ctk.CTkCheckBox(cf, text="边框", variable=self.cad_border).pack(side="left", padx=10)
        
        ctk.CTkLabel(cf, text="边距(mm):  上").pack(side="left", padx=(10, 2))
        ctk.CTkEntry(cf, textvariable=self.margin_top, width=40).pack(side="left")
        ctk.CTkLabel(cf, text="下").pack(side="left", padx=2)
        ctk.CTkEntry(cf, textvariable=self.margin_bottom, width=40).pack(side="left")
        ctk.CTkLabel(cf, text="左").pack(side="left", padx=2)
        ctk.CTkEntry(cf, textvariable=self.margin_left, width=40).pack(side="left")
        ctk.CTkLabel(cf, text="右").pack(side="left", padx=2)
        ctk.CTkEntry(cf, textvariable=self.margin_right, width=40).pack(side="left")

        self._bar = ctk.CTkProgressBar(self); self._bar.grid(row=4, padx=20, pady=10, sticky="ew"); self._bar.set(0)
        self._log = ctk.CTkTextbox(self, height=220); self._log.grid(row=5, padx=20, pady=10, sticky="nsew")
        self._log.configure(state="disabled")
        self._msg("系统就绪。支持 Word/Excel/PPT/图片/CAD。")

        self.btn_go = ctk.CTkButton(self, text="开始执行转换", command=self._start, height=50, fg_color="#2ecc71")
        self.btn_go.grid(row=6, pady=20)

    def _poll_drops(self):
        try:
            while True:
                fs = self._dq.get_nowait()
                added = [f for f in fs if os.path.exists(f) and f not in self.input_paths]
                self.input_paths.update(added)
                if added: self._msg(f"添加 {len(added)} 个路径"); self._lbl.configure(text=f"当前待处理：{len(self.input_paths)} 个路径")
        except: pass
        finally: self.after(100, self._poll_drops)

    def _msg(self, t):
        self._log.configure(state="normal")
        self._log.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {t}\n")
        self._log.see("end"); self._log.configure(state="disabled")

    def _pick_f(self):
        f = filedialog.askopenfilenames(); self.input_paths.update(f); self._sync()
    def _pick_d(self):
        d = filedialog.askdirectory(); self.input_paths.add(d); self._sync()
    def _pick_o(self):
        d = filedialog.askdirectory(); self.output_dir = d; self._btn_out.configure(text=f"保存到: {os.path.basename(d)}")
    def _sync(self): self._lbl.configure(text=f"当前待处理：{len(self.input_paths)} 个路径")

    def _start(self):
        if not self.input_paths or self.is_busy: return
        self.is_busy = True; self.btn_go.configure(state="disabled")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        exts = ('.docx', '.xlsx', '.pptx', '.jpg', '.jpeg', '.png', '.bmp', '.dwg', '.dxf')
        all_fs = []
        for p in self.input_paths:
            if os.path.isfile(p) and p.lower().endswith(exts): all_fs.append(p)
            elif os.path.isdir(p):
                for r, _, fs in os.walk(p):
                    for f in fs:
                        if f.lower().endswith(exts): all_fs.append(os.path.join(r, f))
        
        total = len(all_fs)
        funcs = {'jpg':self._img, 'jpeg':self._img, 'png':self._img, 'bmp':self._img, 'docx':self._off, 'xlsx':self._off, 'pptx':self._off, 'dwg':self._cad, 'dxf':self._cad}
        
        for i, fp in enumerate(all_fs):
            self.after(0, lambda p=(i+1)/total: self._bar.set(p))
            name = os.path.basename(fp)
            self._msg(f"转换: {name}")
            try:
                od = self.output_dir or os.path.join(os.path.dirname(fp), "输出_PDF")
                os.makedirs(od, exist_ok=True)
                base = os.path.splitext(name)[0]
                dst = os.path.join(od, base + ".pdf")
                c = 1
                while os.path.exists(dst): dst = os.path.join(od, f"{base}({c}).pdf"); c += 1
                
                ext = name.rsplit('.', 1)[-1].lower()
                if ext in funcs:
                    funcs[ext](fp, dst)
                    is_cad = ext in ('dwg', 'dxf')
                    
                    try: top_mm = float(self.margin_top.get())
                    except ValueError: top_mm = 3.5
                    try: bot_mm = float(self.margin_bottom.get())
                    except ValueError: bot_mm = 3.5
                    try: l_mm = float(self.margin_left.get())
                    except ValueError: l_mm = 3.5
                    try: r_mm = float(self.margin_right.get())
                    except ValueError: r_mm = 3.5

                    self._add_b(dst, border=(self.cad_border.get() if is_cad else False),
                                margins=(top_mm, bot_mm, l_mm, r_mm))
            except Exception as e: self._msg(f"失败 {name}: {e}")
        
        self._msg("完成"); self.input_paths.clear(); self.after(0, self._sync)
        self.is_busy = False; self.after(0, lambda: self.btn_go.configure(state="normal"))

    def _img(self, s, d):
        i = Image.open(s); i.convert('RGB').save(d, "PDF")

    def _off(self, s, d):
        s, d = os.path.abspath(s), os.path.abspath(d)
        ext = s.rsplit('.', 1)[-1].lower()
        comtypes.CoInitialize()
        try:
            if ext == 'docx':
                app = comtypes.client.CreateObject("Word.Application")
                app.Visible = False; doc = app.Documents.Open(s); doc.SaveAs(d, 17); doc.Close(); app.Quit()
            elif ext == 'xlsx':
                app = comtypes.client.CreateObject("Excel.Application")
                app.Visible = False; wb = app.Workbooks.Open(s); wb.ExportAsFixedFormat(0, d); wb.Close(); app.Quit()
            elif ext == 'pptx':
                app = comtypes.client.CreateObject("Powerpoint.Application")
                pres = app.Presentations.Open(s, WithWindow=False); pres.SaveAs(d, 32); pres.Close(); app.Quit()
        finally: comtypes.CoUninitialize()

    def _cad(self, s, d):
        if not self.accore: raise Exception("No AutoCAD")
        s, d = os.path.abspath(s), os.path.abspath(d)
        off = '_C' if self.cad_center.get() else '0,0'
        m = '\n'.join([
            '_ctab', 'Model',
            '_regen',
            '_zoom', '_e',
            '_zoom', '_e',
            '_-plot', '_Y', '', 'DWG To PDF.pc3', 'ISO_A3_(420.00_x_297.00_MM)',
            '_M', '_L', '_N',
            '_E', '_F', off,
            '_Y', '.', '_Y', '',
            f'"{d}"', '_N', '_Y',
            '_quit', '_y',
        ]) + '\n'
        
        fd, scr = tempfile.mkstemp(suffix=".scr")
        with os.fdopen(fd, 'w', encoding='gbk') as f: f.write(m)
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = subprocess.SW_HIDE
            subprocess.run([self.accore, "/readonly", "/noaudit", "/i", s, "/s", scr], startupinfo=si, timeout=180)
            if not os.path.exists(d): raise Exception("CAD Export Fail")
        finally:
            try: os.remove(scr)
            except: pass

    def _add_b(self, p, border=True, margins=(3.5, 3.5, 3.5, 3.5)):
        try:
            doc = fitz.open(p)
            for pg in doc:
                if border:
                    r = pg.rect
                    mm2pt = 72 / 25.4
                    m_t = margins[0] * mm2pt
                    m_b = margins[1] * mm2pt
                    m_l = margins[2] * mm2pt
                    m_r = margins[3] * mm2pt
                    
                    x0 = max(r.x0 + m_l, r.x0)
                    y0 = max(r.y0 + m_t, r.y0)
                    x1 = max(r.x1 - m_r, x0 + 1)
                    y1 = max(r.y1 - m_b, y0 + 1)
                    if x1 > x0 and y1 > y0:
                        pg.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(0, 0, 0), fill=None, width=2)
            try:
                cat = doc.pdf_catalog()
                doc.xref_set_key(cat, "OpenAction", "[0 /XYZ null null 0.7]")
                doc.xref_set_key(cat, "PageLayout", "/SinglePage")
            except: pass
            doc.save(p + "_t.pdf"); doc.close()
            os.replace(p + "_t.pdf", p)
        except Exception as e:
            self._msg(f"显示优化略过: {e}")

if __name__ == "__main__":
    PDFConverterApp().mainloop()

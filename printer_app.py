"""Local Windows document printing; stdlib only, Office + signed SumatraPDF."""
from __future__ import annotations
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
import json
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import uuid

BASE = Path(__file__).resolve().parent
ENGINE = BASE / 'tools' / 'SumatraPDF-3.6.1-64.exe'
EXTENSIONS = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'}
POWERSHELL = str(Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe')
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def libreoffice() -> Path | None:
    candidates = [Path(os.environ.get('ProgramFiles', r'C:\Program Files'))/'LibreOffice/program/soffice.com',
                  Path(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'))/'LibreOffice/program/soffice.com']
    found = shutil.which('soffice')
    if found: candidates.insert(0, Path(found))
    return next((path for path in candidates if path.is_file()), None)


def convert_with_libreoffice(engine: Path, source: Path, target: Path, cancel: threading.Event | None):
    # An isolated profile prevents an existing LibreOffice window taking over the job.
    with tempfile.TemporaryDirectory(prefix='belge-lo-') as folder:
        work = Path(folder)
        profile = work/'profile'; (profile/'user').mkdir(parents=True)
        (profile/'user/registrymodifications.xcu').write_text(
            '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
            '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item></oor:items>', encoding='utf-8')
        command = [str(engine), '-env:UserInstallation='+profile.as_uri(), '--headless', '--nologo',
                   '--nodefault', '--norestore', '--convert-to', 'pdf', '--outdir', str(work), str(source.resolve())]
        with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=NO_WINDOW) as process:
            deadline = time.monotonic() + 120
            while True:
                try:
                    stdout, stderr = process.communicate(timeout=.2)
                    break
                except subprocess.TimeoutExpired:
                    if (cancel and cancel.is_set()) or time.monotonic() > deadline:
                        # Only this conversion's process tree; never the user's open Office windows.
                        subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'], capture_output=True, creationflags=NO_WINDOW)
                        process.communicate()
                        if cancel and cancel.is_set(): raise Cancelled()
                        raise RuntimeError('Belge dönüşümü zaman aşımına uğradı. Şifreli veya bozuk dosyayı kontrol edin.')
            output = work/(source.stem+'.pdf')
            if process.returncode or not output.is_file():
                detail = (stderr or stdout).decode('utf-8', errors='replace').strip()
                raise RuntimeError('Belge PDF’ye dönüştürülemedi. Şifreli veya bozuk olabilir. ' + detail[-400:])
            shutil.copyfile(output, target)


class Cancelled(Exception):
    pass


def installed_printers() -> dict:
    """Enumerate the same Winspool printers used by the PDF print engine."""
    spool = ctypes.WinDLL('winspool.drv', use_last_error=True)
    class PrinterInfo4(ctypes.Structure):
        _fields_ = [('name', wintypes.LPWSTR), ('server', wintypes.LPWSTR), ('attributes', wintypes.DWORD)]
    spool.EnumPrintersW.argtypes = [wintypes.DWORD, wintypes.LPWSTR, wintypes.DWORD, ctypes.c_void_p,
                                   wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(wintypes.DWORD)]
    needed, count = wintypes.DWORD(), wintypes.DWORD()
    spool.EnumPrintersW(6, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(count))
    if not needed.value: return {'printers': [], 'default': ''}
    buffer = ctypes.create_string_buffer(needed.value)
    if not spool.EnumPrintersW(6, None, 4, buffer, len(buffer), ctypes.byref(needed), ctypes.byref(count)):
        raise ctypes.WinError(ctypes.get_last_error())
    entries = ctypes.cast(buffer, ctypes.POINTER(PrinterInfo4))
    names = sorted({entries[i].name for i in range(count.value)})
    spool.GetDefaultPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    length = wintypes.DWORD()
    spool.GetDefaultPrinterW(None, ctypes.byref(length))
    default = ctypes.create_unicode_buffer(max(length.value, 1))
    spool.GetDefaultPrinterW(default, ctypes.byref(length))
    return {'printers': names, 'default': default.value}


def backend(job: dict, cancel: threading.Event | None = None, timeout: float = 120) -> dict:
    with tempfile.TemporaryDirectory(prefix='belge-islem-') as folder:
        request = Path(folder) / 'request.json'
        request.write_text(json.dumps(job, ensure_ascii=False), encoding='utf-8')
        with subprocess.Popen([POWERSHELL, '-NoLogo', '-NoProfile', '-NonInteractive', '-File',
                               str(BASE / 'office_backend.ps1'), '-Request', str(request)],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=NO_WINDOW) as process:
            deadline = time.monotonic() + timeout
            while True:
                try:
                    stdout, stderr = process.communicate(timeout=.2)
                    break
                except subprocess.TimeoutExpired:
                    if (cancel and cancel.is_set()) or time.monotonic() > deadline:
                        # Stop only this worker. Never kill an existing user Office session.
                        process.kill()
                        process.communicate()
                        if cancel and cancel.is_set():
                            raise Cancelled()
                        raise RuntimeError('İşlem zaman aşımına uğradı. Office iletişim kutularını kontrol edin.')
            output = stdout.decode('utf-8-sig', errors='replace').strip()
            try:
                result = json.loads(output)
            except ValueError:
                raise RuntimeError(stderr.decode('utf-8-sig', errors='replace').strip() or output or 'Office bağlantısı kurulamadı.')
            if process.returncode or result.get('ok') is False:
                raise RuntimeError(result.get('error', 'Belge dönüştürülemedi.'))
            return result


def convert(source: Path, target: Path, cancel: threading.Event | None = None) -> None:
    if not source.is_file():
        raise FileNotFoundError('Dosya bulunamadı: ' + str(source))
    if source.suffix.lower() not in EXTENSIONS:
        raise ValueError('Desteklenmeyen dosya türü.')
    if cancel and cancel.is_set():
        raise Cancelled()
    if source.suffix.lower() == '.pdf':
        shutil.copyfile(source, target)
    elif (engine := libreoffice()) is not None:
        convert_with_libreoffice(engine, source, target, cancel)
    else:
        backend({'action': 'convert', 'source': str(source.resolve()), 'target': str(target.resolve())}, cancel)
    with target.open('rb') as file:
        if file.read(5) != b'%PDF-':
            raise ValueError('Geçerli PDF oluşturulamadı; belge bozuk veya şifreli olabilir.')


@dataclass(frozen=True)
class PrintOptions:
    printer: str
    copies: int = 1
    pages: str = ''
    duplex: str = 'simplex'
    color: str = 'color'
    scale: str = 'fit'

    def command(self, pdf: Path) -> list[str]:
        if not self.printer:
            raise ValueError('Bir yazıcı seçin.')
        if not 1 <= self.copies <= 999:
            raise ValueError('Kopya sayısı 1–999 arasında olmalıdır.')
        if not valid_pages(self.pages):
            raise ValueError('Sayfa aralığı örneği: 1-3,5. Tümü için boş bırakın.')
        settings = ','.join(x for x in (self.pages, f'{self.copies}x', self.duplex, self.color, self.scale) if x)
        return [str(ENGINE), '-print-to', self.printer, '-print-settings', settings, '-silent', str(pdf)]


def valid_pages(value: str) -> bool:
    return not value or bool(re.fullmatch(r'[1-9]\d*(?:-[1-9]\d*)?(?:,[1-9]\d*(?:-[1-9]\d*)?)*', value))


def send_to_printer(pdf: Path, options: PrintOptions) -> None:
    # Fail clearly before submission if Windows denies the printer handle.
    spool = ctypes.WinDLL('winspool.drv', use_last_error=True)
    spool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE), ctypes.c_void_p]
    spool.ClosePrinter.argtypes = [wintypes.HANDLE]
    handle = wintypes.HANDLE()
    if not spool.OpenPrinterW(options.printer, ctypes.byref(handle), None):
        error = ctypes.get_last_error()
        if error == 5:
            raise PermissionError('Windows bu uygulamanın yazıcıya erişimini reddetti. Yazıcı izinlerini veya kurumun yazdırma politikasını kontrol edin. İş gönderilmedi.')
        raise OSError(error, 'Yazıcı açılamadı: ' + ctypes.FormatError(error))
    spool.ClosePrinter(handle)
    if not ENGINE.is_file():
        raise FileNotFoundError('PDF yazdırma bileşeni yok. tools klasörünü koruyun.')
    # Once submitted, let the current transfer finish; cancellation stops later jobs.
    try:
        result = subprocess.run(options.command(pdf), creationflags=NO_WINDOW, timeout=180)
    except subprocess.TimeoutExpired:
        raise RuntimeError('Yazıcı yanıt vermedi. Windows kuyruğunu kontrol edin; iş gönderilmiş olabilir. Otomatik tekrar yapılmadı.')
    if result.returncode:
        raise RuntimeError(f'Yazdırma başarısız (kod {result.returncode}). Yazıcıyı ve belgeyi kontrol edin.')


class App(tk.Tk):
    def __init__(self, initial=()):
        super().__init__()
        self.title('PrintBatch')
        self.geometry('1050x580')
        self.minsize(820, 440)
        self.events: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.cache = Path(tempfile.mkdtemp(prefix='BelgeYazdir-'))
        self.controls = []
        self.default_printer = ''
        self.status = tk.StringVar(value='Dosya ekleyin veya pencereye sürükleyin.')
        self.printer = tk.StringVar()
        self.copies = tk.StringVar(value='1')
        self.pages = tk.StringVar()
        self.duplex = tk.StringVar(value='Tek taraf')
        self.color = tk.StringVar(value='Renkli')
        self.scale = tk.StringVar(value='Sayfaya sığdır')
        bar = ttk.Frame(self, padding=8); bar.pack(fill='x')
        for label, action in [('Dosya ekle', self.add_dialog), ('Seçileni kaldır', self.remove),
                              ('Listeyi temizle', self.clear), ('Yazıcıları yenile', self.refresh_printers),
                              ('Seçileni önizle', lambda: self.start(True))]:
            button = ttk.Button(bar, text=label, command=action); button.pack(side='left', padx=3)
            self.controls.append((button, 'normal'))
        settings = ttk.Frame(self, padding=(8, 2)); settings.pack(fill='x')
        ttk.Label(settings, text='Yazıcı:').grid(row=0, column=0, padx=3, pady=6)
        self.printers = ttk.Combobox(settings, textvariable=self.printer, width=47, state='readonly')
        self.printers.grid(row=0, column=1, columnspan=3, sticky='w'); self.controls.append((self.printers, 'readonly'))
        ttk.Label(settings, text='Kopya:').grid(row=0, column=4, padx=5)
        spin = ttk.Spinbox(settings, from_=1, to=999, textvariable=self.copies, width=5)
        spin.grid(row=0, column=5); self.controls.append((spin, 'normal'))
        ttk.Label(settings, text='Sayfalar:').grid(row=0, column=6, padx=5)
        entry = ttk.Entry(settings, textvariable=self.pages, width=14); entry.grid(row=0, column=7)
        self.controls.append((entry, 'normal'))
        ttk.Label(settings, text='Boş = tümü; örnek: 1-3,5').grid(row=0, column=8, padx=5)
        for column, (variable, values) in enumerate([
            (self.duplex, ['Tek taraf', 'Çift taraf / uzun kenar', 'Çift taraf / kısa kenar']),
            (self.color, ['Renkli', 'Siyah beyaz']), (self.scale, ['Sayfaya sığdır', 'Gerçek boyut'])]):
            widget = ttk.Combobox(settings, textvariable=variable, values=values, width=25, state='readonly')
            widget.grid(row=1, column=column*3, columnspan=3, padx=3, pady=6, sticky='w')
            self.controls.append((widget, 'readonly'))
        table = ttk.Frame(self, padding=8); table.pack(fill='both', expand=True)
        self.files = ttk.Treeview(table, columns=('file', 'status', 'path'), show='headings', selectmode='extended')
        for key, label, width in [('file', 'Dosya', 230), ('status', 'Durum', 280), ('path', 'Konum', 460)]:
            self.files.heading(key, text=label); self.files.column(key, width=width, minwidth=80)
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.files.yview)
        self.files.configure(yscrollcommand=scroll.set)
        self.files.pack(side='left', fill='both', expand=True); scroll.pack(side='right', fill='y')
        self.files.bind('<Double-1>', self.show_details)
        bottom = ttk.Frame(self, padding=8); bottom.pack(fill='x')
        print_button = ttk.Button(bottom, text='Tümünü yazdır', command=lambda: self.start(False))
        print_button.pack(side='left'); self.controls.append((print_button, 'normal'))
        self.stop_button = ttk.Button(bottom, text='Kuyruğu durdur', command=self.stop, state='disabled')
        self.stop_button.pack(side='left', padx=8)
        ttk.Label(bottom, textvariable=self.status, wraplength=700).pack(side='left')
        ttk.Label(self, text='“Yazıcıya gönderildi” fiziksel baskının tamamlandığı anlamına gelmez. Durumu Windows yazıcı kuyruğundan izleyin.',
                  padding=(8, 0, 8, 8)).pack(fill='x')
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.after(100, self.poll)
        self.add_files(initial)
        self.refresh_printers()
        self.after(200, self.install_drop)

    def install_drop(self):
        """Native Windows file drops; callback is retained for the window lifetime."""
        if hasattr(self, 'drop_proc'):
            return
        try:
            user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
            user32.GetParent.argtypes = [wintypes.HWND]; user32.GetParent.restype = wintypes.HWND
            handle = user32.GetParent(self.winfo_id()) or self.winfo_id()
            pointer = ctypes.c_ssize_t
            prototype = ctypes.WINFUNCTYPE(pointer, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
            setter = user32.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user32.SetWindowLongW
            setter.argtypes = [wintypes.HWND, ctypes.c_int, pointer]; setter.restype = pointer
            user32.CallWindowProcW.argtypes = [pointer, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            user32.CallWindowProcW.restype = pointer
            shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
            shell32.DragFinish.argtypes = [wintypes.HANDLE]
            def callback(hwnd, message, wparam, lparam):
                if message == 0x233:
                    paths = []
                    try:
                        for index in range(shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)):
                            length = shell32.DragQueryFileW(wparam, index, None, 0)
                            buffer = ctypes.create_unicode_buffer(length + 1)
                            shell32.DragQueryFileW(wparam, index, buffer, length + 1)
                            paths.append(buffer.value)
                    finally:
                        shell32.DragFinish(wparam)
                    self.events.put(('drop', paths))
                    return 0
                return user32.CallWindowProcW(self.old_proc, hwnd, message, wparam, lparam)
            self.drop_proc = prototype(callback)
            self.old_proc = setter(handle, -4, ctypes.cast(self.drop_proc, ctypes.c_void_p).value)
            shell32.DragAcceptFiles(handle, True)
        except (OSError, AttributeError):
            self.status.set('Dosya ekle düğmesiyle belgelerinizi seçin.')

    def add_dialog(self):
        self.add_files(filedialog.askopenfilenames(filetypes=[('Belgeler', '*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx')]))

    def add_files(self, paths):
        if self.busy:
            return
        existing = {os.path.normcase(self.files.set(i, 'path')) for i in self.files.get_children()}
        skipped = 0
        for value in paths:
            path = Path(value).resolve()
            if not path.is_file() or path.suffix.lower() not in EXTENSIONS:
                skipped += 1; continue
            if os.path.normcase(str(path)) not in existing:
                self.files.insert('', 'end', values=(path.name, 'Hazır', str(path)))
                existing.add(os.path.normcase(str(path)))
        self.status.set(f'{len(existing)} belge.' + (f' {skipped} geçersiz/desteklenmeyen dosya eklenmedi.' if skipped else ''))

    def remove(self):
        if not self.busy:
            for item in self.files.selection(): self.files.delete(item)

    def clear(self):
        if not self.busy:
            for item in self.files.get_children(): self.files.delete(item)

    def show_details(self, _event=None):
        items = self.files.selection()
        if items:
            messagebox.showinfo('Belge durumu', self.files.set(items[0], 'path') + '\n\n' + self.files.set(items[0], 'status'))

    def refresh_printers(self):
        self.status.set('Yazıcılar okunuyor…')
        def worker():
            try: self.events.put(('printers', installed_printers()))
            except Exception as ex: self.events.put(('printer_error', str(ex)))
        threading.Thread(target=worker, daemon=True).start()

    def start(self, preview):
        if self.busy: return
        items = self.files.selection()[:1] if preview else self.files.get_children()
        if not items:
            messagebox.showinfo('PrintBatch', 'Önizleme için bir belge seçin.' if preview else 'Önce dosya ekleyin.'); return
        try:
            options = PrintOptions(self.printer.get(), int(self.copies.get()), self.pages.get().replace(' ', ''),
                {'Tek taraf': 'simplex', 'Çift taraf / uzun kenar': 'duplexlong', 'Çift taraf / kısa kenar': 'duplexshort'}[self.duplex.get()],
                'color' if self.color.get() == 'Renkli' else 'monochrome', 'fit' if self.scale.get() == 'Sayfaya sığdır' else 'noscale')
            if not preview: options.command(Path('validation.pdf'))
            if not ENGINE.is_file(): raise ValueError('tools klasöründeki PDF bileşeni bulunamadı.')
        except ValueError as ex:
            messagebox.showerror('Ayarları kontrol edin', str(ex)); return
        jobs = [(item, Path(self.files.set(item, 'path'))) for item in items]
        self.busy = True; self.cancel.clear()
        for control, _ in self.controls: control.configure(state='disabled')
        self.stop_button.configure(state='normal')
        threading.Thread(target=self.run_jobs, args=(jobs, options, preview), daemon=True).start()

    def run_jobs(self, jobs, options, preview):
        sent = errors = 0
        for item, source in jobs:
            if self.cancel.is_set(): break
            target = self.cache / (uuid.uuid4().hex + '.pdf')
            try:
                self.events.put(('row', item, 'PDF hazırlanıyor…'))
                convert(source, target, self.cancel)
                if self.cancel.is_set(): raise Cancelled()
                if preview:
                    viewer = subprocess.Popen([str(ENGINE), '-reuse-instance', str(target)], creationflags=NO_WINDOW)
                    threading.Thread(target=viewer.wait, daemon=True).start()
                    self.events.put(('row', item, 'Önizleme açıldı'))
                else:
                    self.events.put(('row', item, 'Yazıcıya gönderiliyor…'))
                    send_to_printer(target, options)
                    self.events.put(('row', item, 'Yazıcıya gönderildi')); sent += 1
            except Cancelled:
                self.events.put(('row', item, 'Durduruldu')); break
            except Exception as ex:
                self.events.put(('row', item, 'Hata: ' + str(ex))); errors += 1
            finally:
                if not preview:
                    try: target.unlink(missing_ok=True)
                    except OSError: pass
        self.events.put(('done', sent, errors, preview, self.cancel.is_set()))

    def stop(self):
        self.cancel.set(); self.status.set('Durduruluyor… Gönderilen iş Windows kuyruğunda kalır.')

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == 'row':
                    if self.files.exists(event[1]): self.files.set(event[1], 'status', event[2])
                    self.status.set(event[2])
                elif event[0] == 'printers':
                    values = event[1]['printers']; self.printers.configure(values=values)
                    if self.printer.get() not in values:
                        self.printer.set(event[1].get('default') or (values[0] if values else ''))
                    self.status.set(f'{len(values)} yazıcı hazır.' if values else 'Yazıcı bulunamadı.')
                elif event[0] == 'printer_error': self.status.set('Yazıcılar okunamadı: ' + event[1])
                elif event[0] == 'drop': self.add_files(event[1])
                elif event[0] == 'done':
                    self.busy = False
                    for control, state in self.controls: control.configure(state=state)
                    self.stop_button.configure(state='disabled')
                    _, sent, errors, preview, stopped = event
                    summary = ('Önizleme hazır.' if not errors else 'Önizleme başarısız.') if preview else f'{sent} gönderildi, {errors} hata.'
                    self.status.set(summary + (' Kuyruk durduruldu.' if stopped else ''))
        except queue.Empty: pass
        self.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo('İşlem sürüyor', 'Önce kuyruğu durdurun ve mevcut işlemin bitmesini bekleyin.'); return
        shutil.rmtree(self.cache, ignore_errors=True)
        self.destroy()


if __name__ == '__main__':
    App(sys.argv[1:]).mainloop()

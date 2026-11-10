import hashlib
import json
import os
import shutil
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from printer_app import App, BASE, ENGINE, NO_WINDOW, POWERSHELL, Cancelled, PrintOptions, backend, convert, installed_printers, valid_pages


class LogicTests(unittest.TestCase):
    def test_ranges_and_safe_arguments(self):
        for value in ['', '1', '1-3,5', '3-1']:
            self.assertTrue(valid_pages(value))
        for value in ['0', 'abc', '1,', '-1', '1;2']:
            self.assertFalse(valid_pages(value))
        command = PrintOptions('HP printer name', 2, '1-3,5', 'duplexlong', 'monochrome').command(Path('a b.pdf'))
        self.assertEqual(command[2], 'HP printer name')
        self.assertEqual(command[4], '1-3,5,2x,duplexlong,monochrome,fit')
        self.assertEqual(command[-1], 'a b.pdf')
        with self.assertRaises(ValueError): PrintOptions('HP', 0).command(Path('x.pdf'))

    def test_invalid_and_cancelled(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'x.docx'; target = Path(directory) / 'x.pdf'
            with self.assertRaises(FileNotFoundError): convert(source, target)
            source = Path(directory) / 'x.txt'; source.write_text('hello')
            with self.assertRaises(ValueError): convert(source, target)
            source = Path(directory) / 'invalid.pdf'; source.write_bytes(b'not a pdf')
            with self.assertRaises(ValueError): convert(source, target)
            cancel = threading.Event(); cancel.set()
            with self.assertRaises(Cancelled): convert(source, target, cancel)

    def test_job_queue_continues_after_failure_and_stops_on_cancel(self):
        class QueueApp:
            pass
        import queue
        app = QueueApp(); app.events = queue.Queue(); app.cancel = threading.Event()
        with tempfile.TemporaryDirectory() as directory:
            app.cache = Path(directory)
            first, second = Path('a.pdf'), Path('b.pdf')
            def conversion(source, target, _):
                if source == first: raise ValueError('broken')
                target.write_bytes(b'%PDF-test')
            with patch('printer_app.convert', side_effect=conversion), patch('printer_app.send_to_printer') as send:
                App.run_jobs(app, [('one', first), ('two', second)], PrintOptions('HP'), False)
                self.assertEqual(send.call_count, 1)
                events = list(app.events.queue)
                self.assertEqual(events[-1], ('done', 1, 1, False, False))
            app.cancel.set()
            with patch('printer_app.convert') as conversion:
                App.run_jobs(app, [('one', first)], PrintOptions('HP'), False)
                conversion.assert_not_called()


@unittest.skipUnless(os.environ.get('PRINTER_INTEGRATION') == '1', 'Set PRINTER_INTEGRATION=1 for real Office/GUI checks')
class IntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = BASE / 'test-output'; cls.folder.mkdir(exist_ok=True)
        subprocess.run([sys.executable, str(BASE/'tests/prepare_modern.py')], check=True, timeout=30)
        for extension in ['doc', 'xls', 'ppt']:
            shutil.copyfile(BASE/'tests/fixtures'/f'sample.{extension}', cls.folder/f'sample document.{extension}')

    def test_all_six_office_formats_and_pdf(self):
        from pypdf import PdfReader
        report = []
        for extension in ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']:
            with self.subTest(extension=extension):
                source = self.folder / f'sample document.{extension}'
                before = hashlib.sha256(source.read_bytes()).hexdigest()
                target = self.folder / f'{extension}.pdf'
                convert(source, target)
                reader = PdfReader(target)
                self.assertGreater(len(reader.pages), 0)
                text = '\n'.join(page.extract_text() for page in reader.pages)
                if extension in ['docx', 'xlsx', 'pptx']:
                    self.assertIn('Printer' if extension == 'xlsx' else 'Belge', text)
                else:
                    self.assertTrue(text.strip())
                self.assertEqual(before, hashlib.sha256(source.read_bytes()).hexdigest())
                report.append({'format': extension, 'pages': len(reader.pages), 'text': text, 'source_unchanged': True})
        source = self.folder/'docx.pdf'; target = self.folder/'passthrough.pdf'
        convert(source, target)
        self.assertEqual(source.read_bytes(), target.read_bytes())
        (self.folder/'conversion-report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding='utf-8')

    def test_printers_gui_and_preview(self):
        result = installed_printers()
        self.assertTrue(result['printers'])
        (self.folder/'printers.json').write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding='utf-8')
        app = App()
        try:
            source = self.folder/'docx.pdf'
            app.add_files([source, source, self.folder/'sample document.docx'])
            self.assertEqual(len(app.files.get_children()), 2)
            app.update(); app.install_drop()
            self.assertTrue(hasattr(app, 'drop_proc'))
            app.files.selection_set(app.files.get_children()[0])
            app.start(True)
            deadline = time.monotonic() + 30
            while app.busy and time.monotonic() < deadline:
                app.update(); time.sleep(.05)
            self.assertFalse(app.busy)
            self.assertEqual(app.files.set(app.files.get_children()[0], 'status'), 'Önizleme açıldı')
            from PIL import ImageGrab
            app.update()
            ImageGrab.grab(bbox=(app.winfo_rootx(), app.winfo_rooty(), app.winfo_rootx()+app.winfo_width(), app.winfo_rooty()+app.winfo_height())).save(self.folder/'application.png')
        finally:
            app.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)

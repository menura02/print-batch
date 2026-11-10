"""One-page virtual printer smoke test. Save dialog is completed interactively."""
from pathlib import Path
import json
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from printer_app import BASE, PrintOptions, send_to_printer
try:
    send_to_printer(BASE/'test-output/docx.pdf', PrintOptions('Microsoft Print to PDF', pages='1'))
except Exception as error:
    (BASE/'test-output/virtual-print-result.json').write_text(json.dumps({'submitted': False, 'error': str(error)}, ensure_ascii=False), encoding='utf-8')
    raise
else:
    (BASE/'test-output/virtual-print-result.json').write_text(json.dumps({'submitted': True, 'printer': 'Microsoft Print to PDF'}), encoding='utf-8')

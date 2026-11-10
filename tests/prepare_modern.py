from pathlib import Path
from docx import Document
from openpyxl import Workbook
from pptx import Presentation

folder = Path(__file__).resolve().parents[1] / 'test-output'
folder.mkdir(exist_ok=True)
doc = Document()
doc.add_paragraph('Belge Yazdır test: çğıöşü ÇĞİÖŞÜ')
doc.save(folder / 'sample document.docx')
book = Workbook()
book.active.append(['Product', 'Quantity'])
book.active.append(['Printer', 3])
book.active.print_area = 'A1:B2'
book.save(folder / 'sample document.xlsx')
slides = Presentation()
slide = slides.slides.add_slide(slides.slide_layouts[0])
slide.shapes.title.text = 'Belge Yazdır test'
slide.placeholders[1].text = 'PowerPoint document'
slides.save(folder / 'sample document.pptx')

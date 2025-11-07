import io
import os
import re
import sys
import argparse
import yaml
from pathlib import Path
from typing import Iterator
from docx import Document
from docx.shared import Mm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

__version__ = '0.2.0'

class Text2Docx:
  """text typesetter"""

  head_parser = re.compile(r'(\{\w+\})|([^{}]*)')
  head_align = {
    (False, False): WD_ALIGN_PARAGRAPH.CENTER,
    (True, False):  WD_ALIGN_PARAGRAPH.RIGHT,
    (False, True):  WD_ALIGN_PARAGRAPH.LEFT,
    (True, True):   WD_ALIGN_PARAGRAPH.CENTER,
  }

  def __init__(self, textin) -> None:
    self.conf = self.load_conf()
    self.args = self.get_args()
    if not self.args.raw:
      textin = io.TextIOWrapper(textin.buffer, encoding='utf-8')
    self.doc = Document()
    self.set_section(self.doc.sections[0])
    self.set_style(self.doc.styles['Normal'])
    if self.args.col:
      self.set_multicolumn(self.doc.sections[0], self.args.col)
    if self.args.sample:
      self.set_sample()
    else:
      self.typeset(textin)

  def load_conf(self) -> dict:
    fpath = Path(__file__).resolve().parent / 'config.yaml'
    with open(fpath, 'r', encoding='utf8') as f:
      conf = yaml.safe_load(f)
    return conf

  def get_args(self) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
      prog='python -m text2docx',
      description='text typesetter')
    parser.add_argument('--raw', help='suppress stdin encoding',
      action='store_true')
    parser.add_argument('--out', help='output filename')
    parser.add_argument('--page', help='page size',
      choices=self.conf['page'].keys())
    parser.add_argument('--landscape', help='landscape',
      action='store_true')
    parser.add_argument('--margin', help='margin mm',
      type=float,
      nargs=4, metavar=('top','bottom','left','right'))
    parser.add_argument('--col', help='multi column',
      type=int,
      choices=(2,3))
    parser.add_argument('--size', help='font pt',
      type=float)
    parser.add_argument('--font', help='font',
      choices=self.conf['font'].keys())
    parser.add_argument('--eafont', help='eastasia font',
      choices=self.conf['eafont'].keys())
    parser.add_argument('--sample', help='font sample',
      action='store_true')
    parser.add_argument('--do', help='operation',
      choices=['print', 'edit', 'open'])
    head_args = parser.add_mutually_exclusive_group()
    head_args.add_argument('--number', help='page number on header',
      action='store_true')
    head_args.add_argument('--header', help='header')
    parser.add_argument('--footer', help='footer')
    parser.set_defaults(**self.conf['default'])
    return parser.parse_args()

  def set_section(self, sect) -> None:
    if self.args.landscape:
      sect.orientation = WD_ORIENT.LANDSCAPE
      (sect.page_height,
       sect.page_width) = map(Mm, self.conf['page'][self.args.page])
    else:
      sect.orientation = WD_ORIENT.PORTRAIT
      (sect.page_width,
       sect.page_height) = map(Mm, self.conf['page'][self.args.page])
    (sect.top_margin,
     sect.bottom_margin,
     sect.left_margin,
     sect.right_margin) = map(Mm, self.args.margin)
    (sect.header_distance,
     sect.footer_distance) = map(Mm, [5, 5])
    if self.args.number:
      self.set_head(sect.header, self.conf['head_number'])
    elif self.args.header:
      self.set_head(sect.header, self.args.header)
    if self.args.footer:
      self.set_head(sect.footer, self.args.footer)

  def set_multicolumn(self, sect, num) -> None:
    sectPr = sect._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'), str(num))

  def set_style(self, sty) -> None:
    sty.font.size = Pt(self.args.size)
    sty.font.name = self.conf['font'].get(self.args.font, self.args.font)
    sty.element.rPr.rFonts.set(qn('w:eastAsia'),
      self.conf['eafont'].get(self.args.eafont, self.args.eafont))

  def save(self) -> None:
    self.doc.save(self.args.out)
    if self.args.do:
      os.startfile(self.args.out, operation=self.args.do)

  def typeset(self, textin, sep=None) -> None:
    sep = sep or self.conf['pagesep']
    for page in self.paginate(textin, sep):
      if page == sep:
        self.doc.add_page_break()
      else:
        self.doc.add_paragraph(page)

  def paginate(self, textin, sep) -> Iterator[str]:
    page = []
    for line in textin:
      while True:
        part = line.partition(sep)
        page.append(part[0])
        if part[1] == '':
          break
        else:
          yield ''.join(page)
          yield sep
          page = []
          line = part[2]
          if not line.rstrip():
            break
    if page:
      yield ''.join(page)

  def set_head(self, head, fcode) -> None:
    par = head.paragraphs[0]
    par.alignment = self.head_align[(fcode.startswith(' '), fcode.endswith(' '))]
    for m in self.head_parser.finditer(fcode):
      if m.group(1):
        self.add_field(par, m.group(1)[1:-1])
      else:
        par.add_run(m.group(2))

  def add_field(self, par, text) -> None:
    run = par.add_run()
    run._r.append(OxmlElement('w:fldChar'))
    run._r[-1].set(qn('w:fldCharType'), 'begin')
    run._r.append(OxmlElement('w:instrText'))
    run._r[-1].text = text
    run._r.append(OxmlElement('w:fldChar'))
    run._r[-1].set(qn('w:fldCharType'), 'end')

  def set_sample(self):
    table = self.doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = 'font name'
    hdr[1].text = ''
    for k,fn in self.conf['font'].items():
      row = table.add_row().cells
      row[0].width = Mm(50)
      row[1].width = Mm(150)
      row[0].text = '{0}\n(--font {1})'.format(fn,k)
      r = row[1].paragraphs[0].add_run(self.conf['sample']['font'])
      r.font.name = fn
    for k,fn in self.conf['eafont'].items():
      row = table.add_row().cells
      row[0].width = Mm(50)
      row[1].width = Mm(150)
      row[0].text = '{0}\n(--eafont {1})'.format(fn,k)
      r = row[1].paragraphs[0].add_run(self.conf['sample']['eafont'])
      r.font.name = self.conf['font']['lc']
      r._element.rPr.rFonts.set(qn('w:eastAsia'), fn)

if __name__ == '__main__':
  d = Text2Docx(sys.stdin)
  d.save()

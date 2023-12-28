import io
import os
import sys
import argparse
from docx import Document
from docx.shared import Mm, Pt
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT

__version__ = '0.1.2'

class Text2Docx:
  """text typesetter"""

  NEWPAGE = '\x0C'

  PAGE = {
    'a3': (297, 420),
    'b4': (257, 364),
    'a4': (210, 297),
    'b5': (182, 257),
    'a5': (148, 210),
    'hagaki': (100, 148),
  }

  FONT = {
    'lc': 'Lucida Console',
    'lst': 'Lucida Sans Typewriter',
  }

  EAFONT = {
    'biz': 'BIZ UDゴシック',
    'hg': 'HGｺﾞｼｯｸ',
    'hge': 'HGｺﾞｼｯｸE',
    'hgm': 'HGｺﾞｼｯｸM',
    'meiryo': 'メイリオ',
    'yu': '游ゴシック',
    'ms': 'ＭＳ ゴシック',
  }

  def __init__(self, st) -> None:
    self.set_args()
    if not self.args.raw:
      st = io.TextIOWrapper(st.buffer, encoding='utf-8')
    self.doc = Document()
    self.set_style(self.doc.styles['Normal'])
    self.typeset(st)

  def set_args(self) -> None:
    parser = argparse.ArgumentParser(
      prog='python -m text2docx',
      description='text typesetter')
    parser.add_argument('--raw', help='suppress stdin encoding',
      action='store_true')
    parser.add_argument('--out', help='output filename',
      default='output.docx')
    parser.add_argument('--page', help='page size',
      default='a4', choices=self.PAGE.keys())
    parser.add_argument('--landscape', help='landscape',
      action='store_true')
    parser.add_argument('--margin', help='margin mm',
      default=(10,10,10,10), type=float,
      nargs=4, metavar=('top','bottom','left','right'))
    parser.add_argument('--size', help='font pt',
      default=14, type=float)
    parser.add_argument('--font', help='font',
      default='lc',  choices=self.FONT.keys())
    parser.add_argument('--eafont', help='eastasia font',
      default='hge', choices=self.EAFONT.keys())
    parser.add_argument('--do', help='operation',
      choices=['print', 'edit', 'open'])
    self.args = parser.parse_args()

  def set_section(self, sect) -> None:
    if self.args.landscape:
      sect.orientation = WD_ORIENT.LANDSCAPE
      (sect.page_height,
       sect.page_width) = map(Mm, self.PAGE[self.args.page])
    else:
      sect.orientation = WD_ORIENT.PORTRAIT
      (sect.page_width,
       sect.page_height) = map(Mm, self.PAGE[self.args.page])
    (sect.top_margin,
     sect.bottom_margin,
     sect.left_margin,
     sect.right_margin) = map(Mm, self.args.margin)
    (sect.header_distance,
     sect.footer_distance) = map(Mm, [5, 5])

  def set_style(self, sty) -> None:
    sty.font.size = Pt(self.args.size)
    sty.font.name = self.FONT.get(self.args.font, self.args.font)
    sty.element.rPr.rFonts.set(qn('w:eastAsia'),
      self.EAFONT.get(self.args.eafont, self.args.eafont))

  def save(self) -> None:
    self.doc.save(self.args.out)
    if self.args.do:
      os.startfile(self.args.out, operation=self.args.do)

  def typeset(self, st):
    for page in self.pagination(st):
      if page == self.NEWPAGE:
        self.doc.add_page_break()
      else:
        self.doc.add_paragraph(page)
        self.set_section(self.doc.sections[-1])

  def pagination(self, st):
    page = []
    for line in st:
      while True:
        part = line.partition(self.NEWPAGE)
        page.append(part[0])
        if part[1] == '':
          break
        else:
          yield ''.join(page)
          yield self.NEWPAGE
          page = []
          line = part[2]
          if not line.rstrip():
            break
    if page:
      yield ''.join(page)

if __name__ == '__main__':
  d = Text2Docx(sys.stdin)
  d.save()

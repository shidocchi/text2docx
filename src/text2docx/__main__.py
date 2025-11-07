import sys
from . import Text2Docx

if __name__ == '__main__':
  d = Text2Docx(sys.stdin)
  d.save()

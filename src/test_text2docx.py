import unittest
from text2docx import Text2Docx

class Text2DocxTest(unittest.TestCase):

  def setUp(self):
    pass

  def tearDown(self):
    pass

  def test_paginate_0(self):
    '''pagination for no FF line'''
    inputs = ['あああ\n', 'いいい\n', 'ううう\n']
    outputs = ['あああ\nいいい\nううう\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_1(self):
    '''pagination for FF line of the beginning'''
    inputs = ['あああ\n', '\fいいい\n', 'ううう\n']
    outputs = ['あああ\n', '\f', 'いいい\nううう\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_2(self):
    '''pagination for FF line of the end'''
    inputs = ['あああ\n', 'いいい\f\n', 'ううう\n']
    outputs = ['あああ\nいいい', '\f', 'ううう\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_3(self):
    '''pagination for FF line of the middle'''
    inputs = ['あああ\n', 'いいい\fううう\n', 'えええ\n']
    outputs = ['あああ\nいいい', '\f', 'ううう\nえええ\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_4(self):
    '''pagination for FF line with prefix text'''
    inputs = ['\fあああ\n', 'いいい\n', 'ううう\n']
    outputs = ['', '\f', 'あああ\nいいい\nううう\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_5(self):
    '''pagination for FF line with suffix text'''
    inputs = ['あああ\n', 'いいい\n', 'ううう\n\f']
    outputs = ['あああ\nいいい\nううう\n', '\f']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_6(self):
    '''pagination for FF line with prefix and suffix text'''
    inputs = ['\fあああ\n', 'いいい\n', 'ううう\n\f']
    outputs = ['', '\f', 'あああ\nいいい\nううう\n', '\f']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

  def test_paginate_7(self):
    '''pagination for multi-FF line'''
    inputs = ['あああ\n', 'いいい\fううう\fえええ\n', 'おおお\n']
    outputs = ['あああ\nいいい', '\f', 'ううう', '\f', 'えええ\nおおお\n']
    t = Text2Docx()
    self.assertEqual(outputs, list(t.paginate(inputs)))

if __name__ == '__main__':
  unittest.main()

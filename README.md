# text2docx
text typesetter


## INSTALL

```
$ python -m pip install git+https://github.com/shidocchi/text2docx.git
```

## USAGE

```
usage: python -m text2docx ...

text typesetter

optional arguments:
  -h, --help            show this help message and exit
  --raw                 suppress stdin encoding
  --out OUT             output filename
  --page {a3,b4,a4,b5,a5,hagaki}
                        page size
  --landscape           landscape
  --margin top bottom left right
                        margin mm
  --size SIZE           font pt
  --font {lc,lst}       font
  --eafont {biz,hg,hge,hgm,meiryo,yu,ms}
                        eastasia font
  --do {print,edit,open}
                        operation
```

```
$ cat sample.txt | python -m text2docx --do open
```

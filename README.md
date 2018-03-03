# googletranslator
mix = True, the new doc present both original and target language line by line.
    = Flase, only present target language.

use google translator API to translate word format document and write back into file with same format
```
from docx import Document
from googletrans import Translator

def work_doc(filename,destination = 'zh-CN', mix = True):
    tx = lambda t : Translator().translate(t,dest=destination).text
    doc = Document(filename)
    for p in doc.paragraphs:        
        txd = tx(p.text)

        p.text = p.text + ('\n' + txd if mix else '')

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txd = tx(cell.text)
                p.text = cell.text + ('\n' + txd if mix else '')


    f = filename.replace('.doc',destination + '.doc')
    doc.save(f)
if __name__ == '__main__':
    filename = 'p1.docx'
    work_doc(filename)
```

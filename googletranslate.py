from docx import Document
from googletrans import Translator
import datetime

def work_doc(filename,destination = 'zh-CN', mix = True):
    tx = lambda t : Translator().translate(t,dest=destination).text
    dlt = datetime.datetime.now()
    print('start to translate ....{0}...to...{1}'.format(filename,destination))
    doc = Document(filename)
    for cnt, p in enumerate(doc.paragraphs):
        if len(p.text) > 0 :
            try:        
                p.text = p.text + ('\n' + tx(p.text) if mix else '')
            except:
                print('Error: ',p.text)
    print('translated {0} paragraphs in {1} secs'.format(cnt,(datetime.datetime.now() - dlt).total_seconds()))
    dlt = datetime.datetime.now()
    for cnt,table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                if len(cell.text) > 0:
                    try:
                        cell.text = cell.text + ('\n' + tx(cell.text) if mix else '')
                    except :
                        print('Error: ',cell.text )

    print('translated {0} paragraphs in {1} secs'.format(cnt,(datetime.datetime.now() - dlt).total_seconds()))
    f = filename.replace('.doc',destination + '.doc')
    doc.save(f)
    print('done, save the tranlated file as: {0}'.format(f))
if __name__ == '__main__':
    filename = 'P2.docx'
    work_doc(filename)

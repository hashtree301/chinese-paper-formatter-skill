import sys, win32com.client
word = win32com.client.DispatchEx('Word.Application')
word.Visible = False
doc = word.Documents.Open(r'D:\gemini云南植物园\1.docx')
c = 0
for p in doc.Paragraphs:
    t = p.Range.Text
    if t.startswith(chr(0xf0b7)):
        p.Range.Characters(1).Text = '● '
        c += 1
    elif t.startswith('o\t'):
        p.Range.Characters(1).Text = '●'
        c += 1
print('Replaced:', c)
doc.Save()
doc.Close()
word.Quit()

import win32com.client as win32

word = win32.gencache.EnsureDispatch('Word.Application')
# 是否可视化
word.Visible = 0
# 打开原始文件
file_path = r"C:\Users\Administrator\Desktop\马克思主义哲学史\pythonProject\Doc1.docx"
# doc2 = Document(r"C:\Users\Administrator\Desktop\马克思主义哲学全书  - 副本.docx")
# 打开
doc = word.Documents.Open(file_path)
doc.Footnotes.Add(doc.Sections.Range)



doc.Close()
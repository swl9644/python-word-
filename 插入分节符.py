from win32com.client import Dispatch
app = Dispatch('Word.Application')
app.Visible = True
doc = app.Documents.Open('D:\Pycharm\排版\马克思主义哲学史.docx')
# 运行下句代码后，s获得新建文档的光标焦点，也就是图中的回车符前
s = app.Selection

doc.Range(0, 0).Select()
s.InsertBreak(2)
s.Find.ClearFormatting
s.Find.ParagraphFormat.OutlineLevel = 5

while s.Find.Execute():
	s.InsertBreak(2)
	s = app.Selection
	s.MoveDown(4,1)


# search_list = ["①","②","③","④","⑤","⑥","@"]
# for search_string in search_list:
# 	doc.Range(0, 0).Select()
# 	while s.Find.Execute(search_string):
# 		range_start = app.Selection.Start.numerator
# 		r = doc.Range(range_start,range_start+1)
# 		doc.Footnotes.Add(Range = r,Reference ="",Text ="")


# doc.Save()
# doc.Close()
# app.Quit()
from win32com.client import Dispatch
app = Dispatch('Word.Application')
app.Visible = False
doc = app.Documents.Open('D:\Pycharm\排版\保卫唯物辩证法 (徐崇温) .docx') # 运行下句代码后，s获得新建文档的光标焦点，也就是图中的回车符前
s = app.Selection
s.FootnoteOptions.Location = 0
s.FootnoteOptions.NumberingRule = 2
s.FootnoteOptions.StartingNumber = 1
s.FootnoteOptions.NumberStyle = 0
s.FootnoteOptions.LayoutColumns = 0

doc.Range(0, 0).Select()
search_list = ["①","②","③","④","⑤","⑥","@","⑦","⑧","⑨"]
for search_string in search_list:
	doc.Range(0, 0).Select()
	while s.Find.Execute(search_string):
		range_start = app.Selection.Start.numerator
		r = doc.Range(range_start,range_start+1)
		doc.Footnotes.Add(Range = r,Reference ="",Text ="")


doc.Save()
doc.Close()
app.Quit()
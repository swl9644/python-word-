import win32com.client as win32
import pyperclip

data = []
list2=[]
start =0
end = 0
# 读取整理好顺序的文件，并转换为列表
with open(r"C:\Users\Administrator\Desktop\排版\新建文本文档.txt", "r",encoding="utf-8") as f:
    for line in f.readlines():
        line = line.strip('\n')  #去掉列表中每一个元素的换行符
        data.append(line)

# 打开word应用程序
word = win32.gencache.EnsureDispatch('Word.Application')
# 是否可视化
word.Visible = 0
# 打开原始文件
file_path = r"C:\Users\Administrator\Desktop\排版\马克思主义哲学全书 .docx"
# doc2 = Document(r"C:\Users\Administrator\Desktop\马克思主义哲学全书  - 副本.docx")
# 打开
doc = word.Documents.Open(file_path)
# 打开修改文件
doc_new = word.Documents.Open(r"C:\Users\Administrator\Desktop\排版\最终版本.docx")
# 设置列表和循环，按顺序重复进行查找粘贴操作
for findstring in data:
    # 设置活动界面，不设置会导致无法查找到结果
    search_range = doc.Content
    # 光标移动至开头
    doc.Range(0, 0).Select()
    # 设置查找的内容，格式
    findstringnum=len(findstring)
    print(findstringnum)
    search_range.Find.ClearFormatting
    search_range.Find.ParagraphFormat.OutlineLevel=2
    search_range.Find.Text = findstring
    #开始查找
    while search_range.Find.Execute(FindText=findstring):
        # 根据标题的长度判断是否是所要查找的标题
        if len(search_range.Paragraphs(1).Range()) -1 == findstringnum:
            # print(len(search_range.Paragraphs(1).Range()))
            pyperclip.copy('')#清空剪切板
            search_range.Select()#选择查找到的内容
            word.Selection.MoveLeft()#光标左移
            start = word.Selection.Start.numerator#读取查找内容开始位置
            # print(start)
            word.Selection.GoTo(11,2,1)#到下一个标题
            end = word.Selection.Start.numerator#读取标题下内容的结束位置
            # 选取光标start到光标end的内容
            doc.Range(start, end).Select()
            # 剪切，复制用Copy（）
            word.Selection.Copy()
            # 粘贴的目标文件末尾
            doc_new.Characters.Last.Paste()
            break
    else:
        # 如果查找不到内容，则记录在文本文件中
        with open(r"C:\Users\Administrator\Desktop\排版\异常.txt", "a", encoding="utf-8") as f:
            f.writelines(findstring+"\n")
# 关闭两个文件
doc.Close()
doc_new.Close()


data = []

with open(r"C:\Users\Administrator\Desktop\排版\新建文本文档.txt", "r",encoding="utf-8") as f:
    for line in f.readlines():
        line = line.strip('\n')  #去掉列表中每一个元素的换行符
        data.append(line)
print(data)
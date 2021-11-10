#2021-10-25 V1.0
#2021-11-7 v1.1 识别标颜色的文本，并增加<color></color>标记
#2021-11-8 V1.2 识别链接文本，并增加<a href></a>标记
#2021-11-9 V2.0 优化了代码结构
#2021-11-10 V2.1 改了一处BUG


#引用python-docs库处理docx文件
import docx
#时间模块，用来操作时间
import datetime
#调用OS、shutil模块，用来创建文件夹\复制文件
import os, shutil
#调用json模块，处理json文件
import json

#输入并获取原始公告（需要在同一根目录)
file_original = input("请输入文件名\n(要求后缀为.docx）：")
gonggao = docx.Document(file_original)
#初始化处理后的公告
xingonggao = []
#获取公告语言版本，并初始化对应的json格式
language = input("公告语言类型\n1.中文 2.葡语 3.西班牙语 4.英语以及其他：")
chn = {"title":"公告"}
eng = {"title": "Notice"}
pu = {"title":"Anúncio"}
xi = {"title":"Anuncio"}

#处理原始公告
for duanluo in gonggao.paragraphs:
	run = duanluo.runs
	for r in run:
		#找出带https链接的文本，并标注<a href=""></a>
		if "https:" in r.text:
			r.text = f'<a href="{r.text}">{r.text}</a>'
		#找出带颜色标记的部分，并为其标注<color></color>
		if r.font.color.type == True:
			r.text = f"<color=#ffc815>{r.text}</color>"
	#把每一个段落处理成文本，并且转化为字符串,把每一行后面都加一个\n
	duanluo = duanluo.text
	duanluo = duanluo.split()
	duanluo.append("\n")
	xingonggao = xingonggao + duanluo

#新公告处理成字符串，注意空格
xingonggao = " ".join(xingonggao)

#根据语言处理成对应的公告配置,并且保存为对应的json和公告专用文件
if language == "1":
	chn["content"] = xingonggao
	announcement = chn
	houzhui = ("中文", "Chinese")
elif language == "2":
	pu["content"] = xingonggao
	announcement = pu
	houzhui = ("葡萄牙语", "Portuguese")
elif language == "3":
	xi["content"] = xingonggao
	announcement = xi
	houzhui = ("西班牙语", "Spanish")
elif language == "4" :
	eng["content"] = xingonggao
	announcement = eng
	houzhui = ("英语", 
		"English", "French", "German", "Indonesia", "Italian", "Polish", "Russian", "Thai", "Turkish")
#按照日期创建文件夹，然后在里面存放公告，需要注意乱码问题
wenjianjia = f"{str(datetime.date.today())}{houzhui[0]}"
if not os.path.exists(wenjianjia):
	os.makedirs(wenjianjia)
file_path = f"{wenjianjia}/announcement4.json"	
with open(file_path, "w",encoding = "utf-8") as f:
	json.dump(announcement, f, ensure_ascii = False)
#复制json文件，并且重命名为对应的公告配置文件
for h in houzhui[1:]:
	shutil.copyfile(file_path, f"{wenjianjia}/announcement4.json_{h}")
#通过对比Key来进行本地化语言表的批量替换脚本
#2021-11-26 V1.0基本功能实现
#2021-11-29 V1.1 优化代码
#2021-11-29 V1.2 (1)删除多余的print(2)处理空值造成的问题
#2021-11-30 V1.3 （1）解决Unnamed空数据列问题（2）支持xls后缀（3）印尼语列标签修正
#2021-12-2 V1.4   增加进度条显示效果
#2021-12-7 V1.5   处理_x000D_字符问题

#利用pandas库来进行数据操作
import pandas as pd
from tqdm import tqdm
import re

print("=======================1读取源文件并分析文件格式========================================")
csv = "Localization.csv"
file_new = input("请输入需要处理的文件(包括后缀excel/csv):")

original = pd.read_csv(csv, encoding="utf_8_sig")
filetype = file_new.split(".")
if filetype[-1] in ["xlsx", "xls"]:
	new = pd.read_excel(file_new)
elif filetype[-1] == "csv":
	new = pd.read_csv(file_new, encoding="utf_8_sig")

#处理excel中的Unnamed空数据问题
cols = [col for col in new if not col.startswith('Unnamed:')]
new = new[cols]
print("========================2读取KEY======================================")
key_original = original.iloc[:, 0]
key_new = new.iloc[:, 0]

print("=========================3通过对比Key，找出交集，提取出需要被替换的内容======================================")
language = ["Chinese", "English", "German",
	"French", "Spanish", "Portuguese",
	"Russian", "Indonesia", "Thai",
	"Polish", "Turkish", "Italian"]

intersection = new[key_new.isin(key_original)]
#降维处理数据，先把Key列提取出来，再把KEY值处理成一个列表用来循环核对
list_a = [v for k, v in intersection.iterrows()]
list_key= [i["Key"] for i in list_a]

print("=========================4找出需要替换的Key并执行替换======================================")
original = original.set_index("Key")
new = new.set_index("Key")

for i in tqdm(list_key):
	#处理空值nan，以免报错
	if pd.isnull(i) == False:
		for l in language:
			original.loc[i, l] = new.loc[i, l]

print("=========================5处理特殊字符======================================")
#处理_x000D_字符问题
original = original.replace(to_replace=r'_x000D_', value='', regex=True)

print("=========================6保存文档======================================")
original.to_csv(csv, encoding="utf_8_sig")

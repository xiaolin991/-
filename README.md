# 环境要求
1. 请安装python3   安装教程: https://zhuanlan.zhihu.com/p/111168324
2. 依赖的库文件:re,openpyxl   使用命令行  输入"pip install openpyxl",按Enter
# 如何使用
1. 将要分析的文本粘贴 到 Data/text.txt文本文件中。
2. 设置好配置文件config.py。
3. 在DataExtract.py 文件目录中运行命令行，输入'python data_extract.py'。   参考：https://blog.csdn.net/LYJ_viviani/article/details/51817755#:~:text=%E6%96%B9%E6%B3%95%E4%B8%80%EF%BC%9A,%E5%9B%9E%E8%BD%A6%E5%B0%B1%E5%8F%AF%E4%BB%A5%E4%BA%86%E3%80%82
4. 提取出的数据保存在 Data/output.xlsx,将Excel中的内容进行保存。
# 注意
## text文本要求:
1. 要查询的文本请用“姓名”进行分隔，否则会发生错误。
2. 文本文件请使用**UTF-8**编码。
3. 输入到文本文件(.txt)后要进行保存，同时关闭编辑器。（保证没有其他软件正在用这个txt文件）
## Excel文件要求
1.在程序运行之前确保保证没有其他软件正在用这个txt文件，如果有，请关闭其相关软件。
## config.py
可以配置:
1. 文本文件路径
2. Excel文件保存路径
3. 查询相关信息的正则表达式

# 下载
下载DataExtract.zip即可

#!/usr/bin/env python
# encoding: utf-8
# File: DataExtract.py
# coding=utf-8
# Author: Aili
# Date: December 30, 2023
# Description: '功能:文本信息提取，使用Excel保存结果'

import re  # 正则表达式
import openpyxl  # Excel文件操作
import config


setting = config.setting
excel_path = config.excel_path
passage_num = 0
print('----------------------------------------------')
print('功能:文本信息提取，使用Excel保存结果')
print('提示: 文本请使用.txt文本文件保存，编码格式为:UTF-8编码')
print('文本路径: '+ config.text_filepath + ' 作用: 保存要分析的文件')
print('配置文件路径:  config.py    作用:配置需要查询的信息')
print('Excel文件路径: '+ config.excel_path +'作用:保存分析后的结果  提示: 在运行该软件前请关闭其它读取该文件的软件')
print('-----------------------------------------')
header = []
for ele in setting:
    head_ele = ele[0]
    header.append(head_ele)
print('------------------------------------------')
print('您要查找的字段名为:')
print(header)
print('------------------------------------------')
print('')
print('')
print('开始查找:')
result = []


# 根据配置和文本，返回line
def analysis_text (setting,text):
    print('----------------------')
    print("请检查提取出的文本和数字：")
    # 使用for循环遍历列表
    line = []      # 存放写入Excel的信息
    line_str = []  # 存放查找的信息
    for item in setting:
        # 检查索引是否在有效范围内
        # index_to_check = len(setting) - 1

        # if 0 <= index_to_check < len(item):
        #   element_at_index = item[index_to_check]
        #    print(f"Element at index {index_to_check}: {element_at_index}")
            # 查找元素并插入
        element = re.findall(item[1], text)
        # 判断是否匹配成功
        if element:
            line_str.append(element[0])
            ele = re.findall(item[2], element[0])  # 提取数字
            if ele:
                print(element,'->' ,ele[-1])   # 方便查错
                line.append(ele[-1])
            else:
                print(element,'->' ,'\' \'') # 方便查错
                line.append('')
        else:
            print(element,'->','\' \'')# 方便查错
            line_str.append('')
            line.append('')
    print('写入到Exclel的数据为:',line)
        # else:
        #   print(f"Index {index_to_check} is out of range.")
    print('----------------------------------------')
    print('这一段文本是:')
    print(text)
    print('')
    # print('这个文本匹配的搜索结果是:')
    # print(line_str)
    # print('')
    # print('这一段写入Excel数据是:')
    # print(line)
    print('----------------------------------------')
    return line

# 打开文件
text_filepath = config.text_filepath
with open(text_filepath, 'r', encoding='utf-8') as file:
    # 读取文件内容
    file_content = file.read()

# 对文本内容进行分割
# 使用正则表达式按照"姓名："进行分割
    pattern = re.compile(r'姓名')
    text_list = re.split(pattern, file_content)
    for ele in text_list:
        if(ele):
            ele  = '姓名' + ele
            # print('元素：')
            # print(ele)
            print('\n')
            passage_num += 1
            print('这是第',passage_num,'段:')  # 提示这是第几段。
            result.append(analysis_text(setting,ele))

print('')
print('')
# print('保存到Excel中的数据是:')
# print(result)
print('--------------------------------------------')

# 将结果写入excel

# 创建一个新的Excel工作簿和工作表
workbook = openpyxl.Workbook()
sheet = workbook.active

sheet.append(header)

# 示例数据
data = result

# 将数据写入工作表

# 写入数据
for row in data:
    sheet.append(row)

# 保存工作簿到文件
workbook.save(excel_path)
print('')
print('提取出的数据已保存到:'+ excel_path)
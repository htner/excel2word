#!/usr/bin/env python
# -*- coding: utf-8 -*- 
import os
import re
import xlrd
import sys
from mailmerge import MailMerge


## template_names=['temp'] # 其他模板文件，加到这个list里即可
date_convert_cols=['tm_time'] # 需要转换日期格式的列名
template_type='_template.docx'
out_col_name=1   # 默认以第一行为输出文件, 如果需要第n行，修改为n
out_doc_type='.docx'

def file_filter(f):
    if f[-14:] in [template_type]:
        print('  模板源文件：' + f)
        return True
    else:
        return False

def batch(maindir):
    print('当前路径：' +  maindir)
    files = os.listdir(maindir+'/.')
    template_files = list(filter(file_filter, files))

    for f in os.listdir(maindir+'/.'):
        if not os.path.splitext(f)[1] == '.xlsx' and not os.path.splitext(f)[1] == '.xls':
            continue
        print('  数据源文件：' + f)
        # 打开Excel文件
        xl = xlrd.open_workbook(os.path.join(maindir, f))
        print('  Sheet：' + str(xl.sheet_names()))
        # 读取第一个表
        table = xl.sheet_by_name(xl.sheet_names()[0])

        # 获取表中行数
        nrows = table.nrows
        
        # 获取该表总列数
        ncols = table.ncols

        for template_name in template_files:
            print('    正在处理模板：' + template_name)
            path_name = os.path.join(maindir, 'out')
            if not os.path.exists(path_name):
                os.makedirs(path_name)
            print('    将保存到' + path_name)

            for i in range(1, nrows):  # 循环逐行打印
                # 第一行为表头，不作为填充数据
                doc = MailMerge(maindir + '/' + template_name)  # 打开模板文件
                # 以下为填充模板中对应的域，
                print('      正在处理：' + str(table.row_values(i)[0]))

                mergeInfo = {}
                for j in range(0, ncols):
                    if str(table.row_values(0)[j]) in date_convert_cols:
                        mergeInfo[str(table.row_values(0)[j])] = str(excel_date_convert(table.row_values(i)[j]))
                    else:
                        mergeInfo[str(table.row_values(0)[j])] = str(table.row_values(i)[j])

                # print(mergeInfo)

                doc.merge(**mergeInfo)

                # doc.merge(
                    # name=str(table.row_values(i)[0]),
                    # gender=str(table.row_values(i)[1]),
                    # birthday=excel_date_convert(table.row_values(i)[2]),
                    # id_card=str(table.row_values(i)[3]),
                    # register_address=str(table.row_values(i)[4]),
                    # phone_number=str(table.row_values(i)[5]),
                    # home_address=str(table.row_values(i)[6]),
                    # loan_balance=str(table.row_values(i)[7]),
                    # due_plus_date=str(excel_date_convert(table.row_values(i)[8])),
                # )
                # os.chdir(path_name)
                word_name = os.path.join(path_name, template_name[0:-14] + '_'+ str(table.row_values(i)[out_col_name-1]) +'_'+str(i)+ out_doc_type)
                print("        正在保存 " + word_name)
                doc.write(word_name)
                print("        保存成功")
                doc.close()
                doc = None

def excel_date_convert(excel_date):
    temp_tuple = re.split('/| ', str(excel_date))    # 字符式
    print(temp_tuple)
    # temp_tuple = xlrd.xldate_as_tuple(excel_date, 0)  # 内置
    format_date='{0}年{1}月{2}日'.format(temp_tuple[0], temp_tuple[1], temp_tuple[2])
    return format_date

if __name__ == '__main__':
    batch(sys.argv[1])

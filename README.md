# excel2word

## 需求
通过excel表中搜集的数据批量导出各种类型的word文档，诸如立案申请书、证据清单等等。

## 数据源
数据源:Excel保存名字、性别、出身年月等

模板包含：立案申请书模板.docx、证据清单模板.docx等。

本方法主要参考 

[使用python将Excel数据填充Word模板并生成Word](https://www.dact.dev/post/excel2word_py/)

[dact.dev](https://www.dact.dev/post/excel2word_py/)

来实现

并在原脚本的基础上，解决了一些问题，并做了一些优化，包括但不限于如下:
```
修复macos下的依赖问题
一键执行
带空格的目录的兼容
日志打印优化
优化自动匹配域的功能
```

## 安装
### windos下python3的安装
- 请注意，Python 3.11.8不能在Windows 7或更早版本上使用。Windows 7或更早版本请安装其他版本。
- 下载 https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe
- 点击 python-3.11.8-amd64.exe 执行，出现安装对话框 选中选项“Add Python 3.6 to PATH”,将Python程序的路径加入到Path环境变量中, 然后点击 Install now
- 点击 install_lib.bat 执行

### mac下python3的安装
mac下使用终端执行 python3 自动跳出xcode安装，直接安装即可 因python2存在一些中文编码的问题，故建议使用python3

#### MailMerge ImportError
``` shell
ImportError: cannot import name 'MailMerge' from 'mailmerge'
```
需要安装的包是 docx-mailmerge 而不是 mailmerge
``` Python
pip3 uninstall mailmerge
pip3 install docx-mailmerge
```
然后就可以正常import


#### xlrd.biffh.XLRDError
提示 不支持xlsx文件
``` shell
xlrd.biffh.XLRDError: Excel xlsx file; not supported
``` 
```  shell
The previous version, xlrd 1.2.0, may appear to work, but it could also expose you to potential security vulnerabilities. With that warning out of the way, if you still want to give it a go, type the following command:
``` 
xlrd版本过高，限定在1.2.0版本即可。

``` 
pip3 install xlrd==1.2.0
``` 

## 测试
+ python3 excel2word.py .   其中.是指输出目录, 输出在输入目录的out目录里
+ 到outtest目录里面查看是否生成了相应的三个文件

## 用法

### 制作模板文件
按照execl表的格式，制作相应的模板文件，需在 插入/域/邮件合并/MERGEFIELD 增加相应的对象
``` 
注意域的名字要和excel的第一列中名字相同，例如 name
``` 
### 执行
+ 移除例子里的 '.xlsx' '.docx' 文件
+ 将模板文件, execl文件放到当前目录，模板文件必须命名为"xxx_template.docx"格式
+ 如果有特殊的需求，修改 excel2word.py
+ python3 excel2word.py input # 其中是指input输出目录，输出在输入目录的out目录里

### 高阶用法
修改 excel2word.py, 解锁更多用法

- 默认 out_col_name 为1，指的是用第一列作为生成的文件名，可以修改为其他列
- date_convert_cols 参数定义的列会修改为 {xxxx年xx月xx日} 的格式为输出


### 支持日期格式转换
直接对xlrd的xldate_as_tuple方法生成的tuple做处理，提取对应的年月日，不依赖datetime、timestamp之类的包，简单高效。

- 可以按需修改 excel_date_convert 函数生成不同的日期格式
- 目前源码中已经修改为文本的处理方式

``` Python
def excel_date_convert(excel_date):
    temp_tuple = xlrd.xldate_as_tuple(excel_date, 0)
    format_date='{0}年{1}月{2}日'.format(temp_tuple[0], temp_tuple[1], temp_tuple[2])
    return format_date
```

## R
https://www.dact.dev/post/excel2word_py/

https://www.jianshu.com/p/b876a0d1940a

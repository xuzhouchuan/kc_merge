使用方法：
1.安装python和pip, http://blog.csdn.net/lengqi0101/article/details/61921399
2.安装python 依赖包 pip install -r requirements.txt
3.python ./merge_excel.py -d 放各个公司excel的目录名 [-o 输出文件名]
这是一个自己用的，处理excel合并的自动化程序，代码只能用屎来形容，特别烂。
总结一下烂在什么地方：
1.目录组织不行，一个文件处理了所有事情，就是个呆子
2.程序结构不行，因为每个excel处理都太细化了，这个表格怎么合并，那个表格怎么合并，都是一些碎的操作，当时是一面看一面写，没有抽象出来操作开成好用的kpi，也没有配置、log
3.易用性也不行，现在使用的主要是xlrd、wlwt、openpyxl三个库，前边两个只处理xls格式，后面两个只处理xlsx格式，不好处理啊，最好是写成service，你直接填固定格式的表上就可以了，写到数据库里，多么规范

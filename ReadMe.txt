1.修改文件xlrt/WorkBook.py line 55: encoding="utf-8"
2.修改文件xltr/UnicodeUtils.py line 49:
    elif s is None:
        us = unicode('', encoding)

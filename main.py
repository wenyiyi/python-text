#coding:utf-8
import os
import docx
import re
import sys
import subprocess
import win32com.client
from docx import Document
from operatedoc import RemoteWord

#主文件

#旧的字符串
old_text = '*****'
#新的字符串
new_text = '*****'
#要遍历的目录
fileDir = "d:\\test"

# filelist
def listFiles(dirPath):
    fileList = []
    for root, dirs, files in os.walk(dirPath):
        for fileObj in files:
            fileList.append(os.path.join(root, fileObj))
    return fileList

#修改docx的方法
def replace_text(old_text, new_text,doc,fileObj):
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in inline:
                if old_text in i.text:
                    text = i.text.replace(old_text, new_text)
                    i.text = text
    doc.save(fileObj)

def main():
    fileList = listFiles(fileDir)
    for fileObj in fileList:
        if os.path.splitext(fileObj)[1] == '.txt':
            f = open(fileObj, 'r+')
            all_the_lines = f.readlines()
            f.seek(0)
            f.truncate()
            for line in all_the_lines:
                f.write(line.replace(old_text, new_text))
            f.close()
        elif os.path.splitext(fileObj)[1] == '.docx':
              print("准备修改"+fileObj)
              doc = Document(fileObj)
              replace_text(old_text, new_text, doc, fileObj)
              print(fileObj+"修改成功！")
        elif os.path.splitext(fileObj)[1] == '.doc':
            print("准备修改" + fileObj)
            doc = RemoteWord(fileObj)  # 初始化一个doc对象
            doc.replace_doc(old_text ,new_text ) # 替换doc文本内容
            doc.close()
            print(fileObj + "修改成功")

if __name__ == '__main__':
    main()

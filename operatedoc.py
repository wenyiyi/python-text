#coding:utf-8
import os
import win32com.client


# 处理doc的类
class RemoteWord:
  def __init__(self, filename=None):
      self.xlApp=win32com.client.DispatchEx('Word.Application')
      self.xlApp.Visible=0
      self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
      if filename:
          self.filename=filename
          if os.path.exists(self.filename):
              self.doc=self.xlApp.Documents.Open(filename)
          else:
              self.doc = self.xlApp.Documents.Add()    #创建新的文档
              self.doc.SaveAs(filename)
      else:
          self.doc=self.xlApp.Documents.Add()
          self.filename=''

  def add_doc_end(self, string):
      '''在文档末尾添加内容'''
      rangee = self.doc.Range()
      rangee.InsertAfter('\n'+string)

  def add_doc_start(self, string):
      '''在文档开头添加内容'''
      rangee = self.doc.Range(0, 0)
      rangee.InsertBefore(string+'\n')

  def insert_doc(self, insertPos, string):
      '''在文档insertPos位置添加内容'''
      rangee = self.doc.Range(0, insertPos)
      if (insertPos == 0):
          rangee.InsertAfter(string)
      else:
          rangee.InsertAfter('\n'+string)

  def replace_doc(self,old_text,new_text):
      '''替换文字'''
      self.xlApp.Selection.Find.ClearFormatting()
      self.xlApp.Selection.Find.Replacement.ClearFormatting()
      self.xlApp.Selection.Find.Execute(old_text ,False, False, False, False, False, True, 1, True, new_text ,2)

  def save(self):
      '''保存文档'''
      self.doc.Save()

  def save_as(self, filename):
      '''文档另存为'''
      self.doc.SaveAs(filename)

  def close(self):
      '''保存文件、关闭文件'''
      self.save()
      self.xlApp.Documents.Close()
      self.xlApp.Quit()

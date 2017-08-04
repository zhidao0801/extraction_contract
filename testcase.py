#!/usr/bin/env python
# -*- coding:utf-8 -*-
import unittest
from newchouqu import extract_title,extract_number,extract_money
import textract
import re
filel='/home/mylinux/workspace/a/0.doc'
class testcase(unittest.TestCase):
    def read_testdata(self,rootdir):
        with open(rootdir,'r') as f:
            file_list=f.readlines()
        return file_list
    def read_word(self,filel):
        doccontent = textract.process(filel)
        return doccontent 
    """
    def test_extract_title(self):
        filelist = self.read_testdata('/home/mylinux/workspace/a/title.txt')
        for i in range(len(filelist)):
            if len(re.sub(r'\n|\s*?','',filelist[i]))==0:
                break
            print filelist[i]
            content=self.read_word(filelist[i].split('@')[0])
            self.assertEqual(extract_title(content),re.sub(r'\n','',filelist[i].split('@')[1]),'matcherror')
    
    def test_extract_number(self):
        filelist = self.read_testdata('/home/mylinux/workspace/a/number.txt')
        print len(filelist)
        for i in range(len(filelist)):
            if len(re.sub(r'\n|\s*?','',filelist[i]))==0:
                break
            print filelist[i]
            content=self.read_word(filelist[i].split('@')[0])
            self.assertEqual(extract_number(content),re.sub(r'\n','',filelist[i].split('@')[1]),'matcherror')
    """
    def test_extract_money(self):
        filelist = self.read_testdata('/home/mylinux/workspace/a/money.txt')
        print len(filelist)
        for i in range(len(filelist)):
            if len(re.sub(r'\n|\s*?','',filelist[i]))==0:
                break
            print filelist[i]
            content=self.read_word(filelist[i].split('@')[0])
            self.assertEqual(extract_money(content),re.sub(r'\n','',filelist[i].split('@')[1]),'matcherror')
if __name__ == '__main__':
    unittest.main()

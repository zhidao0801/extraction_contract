#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import textract
import re
from docx2html import convert
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
import sys
reload(sys)
sys.setdefaultencoding('utf8')
#n=os.system('/home/mylinux/workspace/test.sh')
pathlist=[]
def getpathlist(rootdir):
    rootdir=rootdir
    filelist=os.listdir(rootdir)
    global pathlist
    for i in range(len(filelist)):
        path = os.path.join(rootdir,filelist[i])
        if os.path.isfile(path):
            pathlist.append(path)
        elif os.path.isdir(path):
            getpathlist(path)
        else:
            pass
    return pathlist
def read_word(path):
    doccontent = textract.process(path)
    return doccontent   
def extract_title(content):
    doccontent=content
    a=doccontent.split('\n')
    hetongmingcheng=""
    if re.search(r'\|(\s*?)合同名称(\s*?)\|',doccontent,re.S):      
        for i in range(len(a)):
            if "合同名称" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "合同名称" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    hetongmingcheng=hetongmingcheng+a[k].split('|')[q+1]  
        hetongmingcheng=re.sub(r"(\s*)|(\n)","",hetongmingcheng)
        print hetongmingcheng    
    elif re.search(r'(.*?)(合同编号|工程项目名称|闽宏价合|建设单位|项目名称)',doccontent,re.S):
        hetongmingcheng=re.sub(r'(\n)|(\s*)|(\|)','',re.search(r'(.*?)(合同编号|工程项目名称|闽宏价合|建设单位|项目名称)',doccontent,re.S).group(1))
        if "：" in hetongmingcheng or ":" in hetongmingcheng:
            a=re.split(r"：|:",hetongmingcheng)
            hetongmingcheng=a[-1]
        print hetongmingcheng 
    elif re.search(r'订立的',doccontent,re.S):
        for i in range(len(a)):
            if '订立的' in a[i]:
                j=i+1
                while len(re.sub(r'(\s*)|(\n)','',a[j]))==0:
                    j=j+1
                k=j
                while len(re.sub(r'(\s*)|(\n)','',a[k]))>0:
                    k=k+1
                for w in range(j,k):
                    hetongmingcheng=hetongmingcheng+a[w]
                hetongmingcheng=re.sub(r"(\s*)|(\n)","",hetongmingcheng)
                print hetongmingcheng
                return hetongmingcheng
    else:
        hetongmingcheng=''
        print hetongmingcheng
    return hetongmingcheng
def extract_number(content):
    doccontent=content
    hetongbianhao=""
    if re.search(r'合同编号(.*?)(:|：)\n',doccontent,re.S):
        hetongbianhao= re.search(r'合同编号(.*?)(:|：)(.*?)\n',doccontent,re.S).group(3)
        print hetongbianhao
    elif re.search(r'合同编号(\s*?)\|?',doccontent,re.S):
        a=doccontent.split('\n')
        for i in range(len(a)):
            if "合同编号" in a[i]:
                for q in range(len(a[i].split('|'))):
                     if "合同编号" in a[i].split('|')[q]:
                         break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    hetongbianhao=hetongbianhao+a[k].split('|')[q+1]  
                hetongbianhao=re.sub(r"(\s*)|(\n)","",hetongbianhao)
                print hetongbianhao
    else :
        hetongbianhao=''
        print hetongbianhao
    return hetongbianhao
def extract_money(content): 
    hetongjiner=""
    doccontent=content
    if re.search(r'人民币大写(.*?)）',re.sub(r'\n','',doccontent),re.S):  
        c=re.sub(r'(\n|\s*)','',re.search(r'人民币大写(.*?)）',re.sub(r'\n','',doccontent),re.S).group(1))   
        hetongjiner= re.sub(r'[^\d*?|\.|万]','',c)
        print hetongjiner
    elif re.search(r'人民币（大写）(.*?)）',re.sub(r'\n','',doccontent),re.S):
        c=re.sub(r'(\n|\s*)','',re.search(r'人民币（大写）(.*?)）',re.sub(r'\n','',doccontent),re.S).group(1))  
        hetongjiner= re.sub(r'[^\d*?|\.|万]','',c)
        print hetongjiner
    elif re.search(r'报酬总额(.*?)）',re.sub(r'\n','',doccontent),re.S):
        c=re.sub(r'(\n|\s*)','',re.search(r'报酬总额(.*?)）',re.sub(r'\n','',doccontent),re.S).group(1))
        hetongjiner= re.sub(r'[^\d*?|\.|万]','',c)
        print hetongjiner
    elif re.search(r'人民币￥(.*?)）',re.sub(r'\n','',doccontent),re.S):
        c=re.sub(r'(\n|\s*)','',re.search(r'人民币￥(.*?)）',re.sub(r'\n','',doccontent),re.S).group(1))
        hetongjiner= re.sub(r'[^\d*?|\.|万]','',c)
        print hetongjiner
    elif re.search(r'送审总价',re.sub(r'\n','',doccontent),re.S):
        c=re.sub(r'(\n|\s*)','',re.search(r'送审总价(.*?)(:|：)(.*?)元',re.sub(r'\n','',doccontent),re.S).group(3))
        hetongjiner = re.sub(r'[^\d*?|\.|万]','',c)
        print hetongjiner
    elif "合同金额" in doccontent:
        a=doccontent.split('\n')
        for i in range(len(a)):
            if "合同金额" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "合同金额" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    hetongjiner=hetongjiner+a[k].split('|')[q+1]  
        hetongjiner=re.sub(r"(\s*)|(\n)","",hetongjiner).split('元')[0]
        print hetongjiner 
    else :
        hetongjiner=''
        print hetongjiner
    return hetongjiner 
def extract_secondparty(content):
    doccontent=content
    hetongduifang=""
    if re.search(r'受托方(.*?)(:|：)(.*?)\n',doccontent,re.S):
        hetongduifang= re.search(r'受托方(.*?)(:|：)(.*?)\n',doccontent,re.S).group(3) 
        hetongduifang=re.sub(r"(\s*)|(\n)","",hetongduifang)
        print hetongduifang
            
    elif re.search(r'乙(\s*?)方(:|：)(.*?)\n',doccontent,re.S):
        hetongduifang= re.search(r'乙(\s*?)方(:|：)(.*?)\n',doccontent,re.S).group(3)
        hetongduifang=re.sub(r"(\s*)|(\n)","",hetongduifang)
        print hetongduifang
            
    elif re.search(r'受托方(.*?)\n',doccontent):
        hetongduifang = re.search(r'受托方(.*?)\n',doccontent,re.S).group(1)
        hetongduifang=re.sub(r"(\s*)|(\n)","",hetongduifang)
        print hetongduifang
            
    elif re.search(r'承(\s*?)包(\s*?)人(.*?)(:|：)(.*?)\n',doccontent,re.S):
        hetongduifang = re.search(r'承(\s*?)包(\s*?)人(.*?)(:|：)(.*?)\n',doccontent,re.S).group(5)
        hetongduifang=re.sub(r"(\s*)|(\n)","",hetongduifang)
        print hetongduifang
            
    elif re.search(r"(对方当事人)|(合同签约对方)",doccontent):
        a=doccontent.split('\n')
        for i in range(len(a)):
            if re.search(r"(对方当事人)|(合同签约对方)",a[i]) and len(a[i].split('|')) >2:
                for q in range(len(a[i].split('|'))):
                    if re.search(r"(对方当事人)|(合同签约对方)",a[i].split('|')[q]):
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    hetongduifang=hetongduifang+a[k].split('|')[q+1] 
        hetongduifang=re.sub(r"(\s*)|(\n)","",hetongduifang)
        print hetongduifang
            
    else:
        hetongduifang=''
        print hetongduifang
    return hetongduifang
def extract_undertakingdepartment(content):
    doccontent=content
    a=doccontent.split('\n')
    chengbanbumen=""
    for i in range(len(a)):
        if "承办部门" in a[i] and "承办部门意见" not in a[i]:
            for q in range(len(a[i].split('|'))):
                if "承办部门" in a[i].split('|')[q]:
                    break
            j=i+1
            while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                j=j+1
            for k in range(i,j):
                chengbanbumen=chengbanbumen+a[k].split('|')[q+1]
               
        else:
            pass          
    chengbanbumen=re.sub(r'\n|\s*?','',chengbanbumen)      
    print chengbanbumen
    return chengbanbumen 
def extract_dateofsigning(content):
    qiandingrqi=""
    doccontent=content
    if re.search(r'(签订日期|签订时间)(.*?)(：|:)(.*?)\n',doccontent,re.S):
        qiandingrqi=re.search(r'(签订日期|签订时间)(.*?)(：|:)(.*?)\n',doccontent,re.S).group(4)
        qiandingrqi=re.sub(r'(\n|\s*)','',qiandingrqi)
        qiandingrqi=re.sub(r'(年|月|日)','-',qiandingrqi)
        str_list=list(qiandingrqi)
            
        for u in range(len(str_list)):
            if '-' in str_list[len(str_list)-1-u]:
                str_list[len(str_list)-1-u]=''
            else:
                break
        print ''.join(str_list)
        qiandingrqi=''.join(str_list)
    elif "签订日期" in doccontent:
        a=doccontent.split('\n')
        for i in range(len(a)):
            if "签订日期" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "签订日期" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    qiandingrqi=qiandingrqi+a[k].split('|')[q+1]  
        qiandingrqi=re.sub(r"(\s*)|(\n)","",qiandingrqi)
        print qiandingrqi
           
    else:
        qiandingrqi=''
        print qiandingrqi
    return qiandingrqi
def extract_undertaker(content):
    doccontent=content
    a=doccontent.split('\n')
    chengbanren=""
    if re.search(r'\|(.*?)承办人(.*?)\|',doccontent,re.M):
        for i in range(len(a)):
            if "承办人" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "承办人" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    chengbanren=chengbanren+a[k].split('|')[q+1]
               
            else:
                pass          
        chengbanren=re.sub(r'\n|\s*?','',chengbanren)  
    else:
        pass    
    print chengbanren
    return chengbanren 
def extract_typeofcontract(content): 
    doccontent=content
    a=doccontent.split('\n')
    typeofcontract=""
    if re.search(r'\|(.*?)合同类型(.*?)\|',doccontent,re.M):
        for i in range(len(a)):
            if "合同类型" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "合同类型" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    typeofcontract=typeofcontract+a[k].split('|')[q+1]
               
            else:
                pass          
        typeofcontract=re.sub(r'\n|\s*?','',typeofcontract)  
    else:
        
        pass    
    print typeofcontract
    return typeofcontract
def extract_dateprepared(content):
    doccontent=content
    a=doccontent.split('\n')
    dateprepared=""
    if re.search(r'\|(.*?)拟稿时间(.*?)\|',doccontent,re.M):
        for i in range(len(a)):
            if "拟稿时间" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "拟稿时间" in a[i].split('|')[q]:
                        break
                j=i+1
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:
                    j=j+1
                for k in range(i,j):
                    dateprepared=dateprepared+a[k].split('|')[q+1]
            else:
                pass          
        dateprepared=re.sub(r'\n|\s*?','',dateprepared)  
    else:
        pass    
    print dateprepared
    return dateprepared
def extract_numberofcontractcopies(content):
    doccontent=content
    a=doccontent.split('\n')
    numberofcontractcopies=""
    if re.search(r'一式(.*?)份',re.sub(r'\n','',doccontent)):
        numberofcontractcopies=re.search(r'一式(.*)份',re.sub(r'\n','',doccontent)).group(1)
        numberofcontractcopies=re.sub(r'\n|\s*?','',numberofcontractcopies)
        if "份" in numberofcontractcopies:
            a=numberofcontractcopies.split('份')
            numberofcontractcopies=a[0] 
            numberofcontractcopies=re.sub(r'\s*?|_*?','',numberofcontractcopies)
    elif re.search(r'\|(.*?)正本(.*?)\|',doccontent,re.S):
        for i in range(len(a)):
            if "正本" in a[i]:
                for q in range(len(a[i].split('|'))):
                    if "正本" in a[i].split('|')[q]:
                        break
                j=i+1
                
                while len(a[j].split('|')) == len(a[i].split('|')) and len(re.sub(r"(\s*)","",a[j].split('|')[q]))==0:   
                    j=j+1
                for k in range(i,j):
                    numberofcontractcopies=numberofcontractcopies+a[k].split('|')[q-1]
                
            else:
                pass          
    else :
        pass  
    print numberofcontractcopies
    return numberofcontractcopies
def init_table(path):
    book = xlwt.Workbook(encoding='utf8', style_compression=0)
    sheet = book.add_sheet('aa', cell_overwrite_ok=True) 
    sheet.write(0, 0, '合同名称') 
    sheet.write(0, 1, '合同编号')
    sheet.write(0, 2, '合同金额')
    sheet.write(0, 3, '合同签约对方')
    sheet.write(0, 4, '承办部门')
    sheet.write(0, 5, '签订日期')
    sheet.write(0, 6,'承办人')
    sheet.write(0,7,'合同类型')
    sheet.write(0,8,'拟稿时间')
    sheet.write(0,9,'合同份数')
    book.save(path) 
def insert_table(path,result):
    rb = xlrd.open_workbook(path)
    table = rb.sheet_by_index(0)
    nrow = table.nrows	
    wb=copy(rb)
    sheet=wb.get_sheet(0)
    q=nrow
    for i in range(len(result)):  
        sheet.write(q,i,result[i]+u'')
        wb.save(path) 

def main(rootdir,path):
    
    pathlist = os.listdir(rootdir)
    filelist=getpathlist(pathlist)
    if os.path.exists(path):
        pass  
    else:
        init_table(path)
    for i in filelist:
        print i
        result=[]
        content=read_word(i)
        title=extract_title(content)
        number=extract_number(content)
        money=extract_money(content)
        secondparty=extract_secondparty(content)
        undertakingdepartment=extract_undertakingdepartment(content)
        dateofsigning=extract_dateofsigning(content)
        undertaker=extract_undertaker(content)
        typeofcontract=extract_typeofcontract(content)
        dateprepared=extract_dateprepared(content)
        umberofcontractcopies=extract_numberofcontractcopies(content)
        result.append(title)
        result.append(number)
        result.append(money)
        result.append(secondparty)
        result.append(undertakingdepartment)
        result.append(dateofsigning)
        result.append(undertaker)
        result.append(typeofcontract)
        result.append(dateprepared)
        insert_table(path,result)
    print len(filelist)
if __name__ =="__main__":
    rootdir = '/home/mylinux/workspace/150823'
    main(rootdir,'/home/mylinux/workspace/b.xlsx')

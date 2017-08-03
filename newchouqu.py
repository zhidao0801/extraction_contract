#!/usr/bin/env python
# -*- coding:utf-8 -*-


"""
   本程序适用于Linux系统下抽取doc或docx文件里面的内容,
   抽取的内容包括 合同名称、合同编号、合同金额、乙方、
   承办部门、签订时间、承办人、合同类型、拟稿时间、合同份数 等信息。
""" 
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
pathlist=[]

def getpathlist(rootdir):
    """
    返回rootdir文件夹下面所有的文件及子文件夹下面所有的文件
    """
    filelist=os.listdir(rootdir) # 返回rootdir文件夹下面所有的文件及子文件夹
    global pathlist
    for i in range(len(filelist)):
        path = os.path.join(rootdir,filelist[i]) #生成完整路径
        if os.path.isfile(path):
            pathlist.append(path) #如果是文件，则把文件路径加入到pathlist列表中
        elif os.path.isdir(path):
            getpathlist(path) #如果是文件夹，则把递归调用函数getpathlist
        else:
            pass
    return pathlist
def read_word(filelist):
	""" 提取word文档里面的内容"""
    doccontent = textract.process(filelist) 
    return doccontent   
def extract_title(content):
	""" 提取word文档里面的合同名称信息"""
    doccontent=content
    a=doccontent.split('\n')
    hetongmingcheng=""
    if re.search(r'\|(\s*?)合同名称(\s*?)\|',doccontent,re.S):   #判断word文档里面的内容是否是表格   
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
    elif re.search(r'(.*?)(合同编号|工程项目名称|闽宏价合|建设单位|项目名称)',doccontent,re.S): #提取合同名称
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
    """ 提取word文档里面的合同编号信息"""
    doccontent=content
    hetongbianhao=""
    if re.search(r'合同编号(.*?)(:|：)\n',doccontent,re.S):
        hetongbianhao= re.search(r'合同编号(.*?)(:|：)(.*?)\n',doccontent,re.S).group(3) # 提取文本 合同编号 后面的内容
        print hetongbianhao
    elif re.search(r'合同编号(\s*?)\|?',doccontent,re.S): # 提取表格里面的合同编号
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
    """提取合同金额"""
    hetongjiner=""
    doccontent=content 
    if re.search(r'人民币大写(.*?)）',re.sub(r'\n','',doccontent),re.S): 
        c=re.sub(r'(\n|\s*)','',re.search(r'人民币大写(.*?)）',re.sub(r'\n','',doccontent),re.S).group(1))  # 去除换行符和空格 
        hetongjiner= re.sub(r'[^\d*?|\.|万]','',c) #保留具体的数字金额
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
    elif "合同金额" in doccontent: # 提取表格里面的合同金额
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
    """ 提取文件里面的乙方信息"""
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
def extract_undertakingdepartment(content): # 
    """ 提取承办部门信息 """
    doccontent=content
    a=doccontent.split('\n')
    chengbanbumen=""
    for i in range(len(a)): # 提取表格里面的承办部门信息
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
    """提取签订日期"""
    qiandingrqi=""
    doccontent=content
    if re.search(r'(签订日期|签订时间)(.*?)(：|:)(.*?)\n',doccontent,re.S): #提取word文本里面的签订日期
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
    elif "签订日期" in doccontent: # 提取表格里面的签订日期
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
    """提取文件承办人的信息"""
    doccontent=content
    a=doccontent.split('\n')
    chengbanren=""
    if re.search(r'\|(.*?)承办人(.*?)\|',doccontent,re.M):  #提取表格承办人里面的信息
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
    """提取word里面的合同类型信息"""
    doccontent=content
    a=doccontent.split('\n')
    typeofcontract=""
    if re.search(r'\|(.*?)合同类型(.*?)\|',doccontent,re.M): # 提取表格里面的合同类型
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
    """  提取word表格里面的拟稿时间 """
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
    """提取word文件里面的合同份数"""
    doccontent=content
    a=doccontent.split('\n')
    numberofcontractcopies=""
    if re.search(r'一式(.*?)份',re.sub(r'\n','',doccontent)):  
        numberofcontractcopies=re.search(r'一式(.*)份',re.sub(r'\n','',doccontent)).group(1)
        numberofcontractcopies=re.sub(r'\n|\s*?','',numberofcontractcopies)
        if "份" in numberofcontractcopies:
            a=numberofcontractcopies.split('份')
            numberofcontractcopies=a[0]        # 保留合同文本份数的具体数字
            numberofcontractcopies=re.sub(r'\s*?|_*?','',numberofcontractcopies)
    elif re.search(r'\|(.*?)正本(.*?)\|',doccontent,re.S):  # 提取word表格里面的信息
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
    """初始化表格"""
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
    book.save(path)  # 初始化后的表格保存在path路径下
def insert_table(path,result_all):
    """ 将数据插入表格中"""
    if os.path.exists(path):   #判断path路径下的表格是否存在，如果存在则pass，否则初始化表格
        pass  
    else:
        init_table(path)
    rb = xlrd.open_workbook(path)
    table = rb.sheet_by_index(0)
    nrow = table.nrows	
    wb=copy(rb)
    sheet=wb.get_sheet(0)
    q=nrow
    for result in result_all:  # 将抽取的内容插入表格
        for i in range(len(result)):  
            sheet.write(q,i,result[i]+u'')
            wb.save(path) 
def extract_all(filelists):  # 
    """获得所有的信息，包括合同名称、合同金额、合同编号、乙方、承办部门、签订日期、合同份数等信息"""
    filelists=filelists
    result_all=[]
    for i in filelists:
        print i
        result=[]
        content=read_word(i)
        title=extract_title(content) # 返回合同标题
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
        result.append(umberofcontractcopies)
        result_all.append(result)
    return result_all
def main(rootdir,path):
	"""程序主函数，获得rootdir 下所有的文件，获得 抽取的所有内容，将获得的内容插入表格中"""
    filelists=getpathlist(rootdir)  
    result_all=extract_all(filelists) 
    insert_table(path,result_all) 
if __name__ =="__main__":
    rootdir = '/home/mylinux/workspace/150823'
    main(rootdir,'/home/mylinux/workspace/b.xlsx')

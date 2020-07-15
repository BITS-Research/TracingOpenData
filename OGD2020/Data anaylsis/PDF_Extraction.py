#!/usr/bin/env python
# encoding: utf-8
import os.path
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
import pandas as pd
import re
import time
import timeout_decorator
from xlrd import open_workbook
from xlutils.copy import copy
#from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
import eventlet#导入eventlet这个模块

# 1. Transform pdf to txt and return txt
# 2. Extract emails and save to txt file
def parse(path):
    doc_text = ''
    try:
        fp = open(path, 'rb')
        praser = PDFParser(fp)
        doc = PDFDocument(praser)
        praser.set_document(doc)

        if not doc.is_extractable:
            print('from pdfminer.pdfdocument import PDFDocument')
        else:
            rsrcmgr = PDFResourceManager()
            laparams = LAParams()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)
                layout = device.get_result()
                for x in layout:
                    if (isinstance(x, LTTextBoxHorizontal)):
                        results = x.get_text()
                        results = results.replace('\n', ' ')
                        doc_text = doc_text + str(results)

            #emailRegex = re.compile(r'([a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+)', re.VERBOSE)
            Email = re.findall(r'([a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+)', doc_text)
            print('e-mail')
            print(Email)
            with open(r'D:\暑研\UW\email-data.txt', 'a', encoding='utf-8') as f2:#email-data.txt
                for email in Email:
                    f2.write(email[0]+', ')
    except:
        pass

    return doc_text

# Read portals, save as list and return the list
def url_list(url_path):
    df = pd.read_excel(url_path, usecols=[1], names=None)
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[0])
    urls = []
    for url in result:
        if url[-1] == '/':
            url1 = url[:-1]
        else:
            url1 = url

        #delete 'http' and 'www'
        if url1[:7] == 'http://':
            url2 = url1[7:]
        else:
            url2 = url1

        if url2[:8] == 'https://':
            url3 = url2[8:]
        else:
            url3 = url2

        if url3[:4] == 'www.':
            url4 = url3[4:]
        else:
            url4 = url3

        urls.append(url4)
    return urls

# 1. Extract the portals appear in txt transformed from pdf article as list
# 2. Save the list into xls

def find_url(file,url_list, results):
    portals = []
    for url in url_list:
        if str(results).find(url)!= -1:
            portals.append(url)
            temp = set(portals)
            portals = list(temp)
            print(url)
            print(str(results).find(url))#health.data.ny.gov
    print(portals)

    r_xls = open_workbook(r'D:\暑研\UW\-data.xls')  # 读取excel文件  -data   test
    row = r_xls.sheets()[0].nrows
    excel = copy(r_xls)
    table = excel.get_sheet(0)

    table.write(row, 0, file)
    table.write(row, 1, str(portals))

    excel.save(r'D:\暑研\UW\-data.xls')


if __name__ == '__main__':
    url_path = r'D:\暑研\UW\OGD2020_ portals.xlsx'
    url_list = url_list(url_path)
    folder = r'D:\暑研\UW\pdf\IEEE'  #data.gov1015  -data.gov
    files = os.listdir(folder)
    #print(files)
    # Iterate the pdf folder
    for file in files:
        print(file)
        print(time.localtime(time.time()))
        path = os.path.join(folder, file)
        # Parse pdf and extract emails
        results = parse(path)
        if results!='':
            #Extract portals
            find_url(file, url_list, results)
        else:
            pass
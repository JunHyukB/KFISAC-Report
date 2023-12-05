# This is a sample Python script.
import numpy as np
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
# 문단 정렬
from docx.enum.text import WD_ALIGN_PARAGRAPH
# 문자 스타일 변경
from docx.enum.style import WD_STYLE_TYPE
# 가장 기본적인 기능(문서 열기, 저장, 글자 쓰기 등등)
from docx import Document
import pandas as pd
from datetime import datetime, timedelta
import re

"""
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.
"""
def docx():
    doc = Document()
    doc.save('/Users/jun/Desktop/test.docx')

def excel(filename):
    df = pd.read_excel(filename, engine='openpyxl')
    df = df.replace(np.nan,'',regex=True)

    html(df.head())
    text_name = ""
    text_src = ""
    text_dst = ""
    p = re.compile("[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+") #ip 정규표현삭

    for i in df.index:
        if df['탐지유형'][i] != "":
            text_name = text_name + df['탐지유형'][i] + "(" + df['탐지 이벤트'][i] + "건)\n"

        if df['공격대상'][i] != "":
            if p.match(df['공격대상'][i]):
                if text_src != (df['공격대상'][i] + "(" + df.iat[i,5] + "), "): #중복제거 로직 아직 미완
                    text_src = text_src + df['공격대상'][i] + "(" + df.iat[i,5] + "), "
            else:
                if df.iat[i, 7] == "" and text_src !="":
                    text_src = text_src[:-2]
                    text_src = text_src + "\n"

        if p.match(df.iat[i,7]):
            text_dst = text_dst + df.iat[i,7] + "(tcp/" + df.iat[i,8] + "), "
        else:
            if df.iat[i,7] == "" and text_dst != "":
                text_dst = text_dst[:-2]
                text_dst = text_dst + "\n"
    text_src = text_src[:-2]
    text_dst = text_dst[:-2]
    list1 = text_name.split("\n")
    list2 = text_src.split("\n")
    list3 = text_dst.split("\n")


    for j in range(0,len(list1)):
        print(list1[j])
        print(list2[j])
        print(list3[j])

def html(excelfile):
    html_text = """<!DOCTYPE html><html><head><title>report</title><meta charset="UTF-8"></head><body>""" + excelfile.to_html() + """</body></html>"""
    html_file = open("/Users/jun/Desktop/report.html",'w')
    html_file.write(html_text)
    html_file.close()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #print_hi('PyCharm')
    a = datetime.today() - timedelta(3)
    filename = "/Users/jun/Desktop/탐지분석_"+a.strftime("%Y-%m-%d")+".xlsx"
    excel(filename)

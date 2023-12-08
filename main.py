# This is a sample Python script.
import numpy as np
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
# 문단 정렬
#from docx.enum.text import WD_ALIGN_PARAGRAPH
# 문자 스타일 변경
#from docx.enum.style import WD_STYLE_TYPE
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

    text_name = ""
    text_src = ""
    text_dst = ""
    text_date = ""
    p = re.compile("[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+") #ip 정규표현식 적용
    r = re.compile("[0-9]+\-[0-9]+\-[0-9].+") #날짜 정규표현식

    for i in df.index:
        # 공격유형 로직
        if df['탐지유형'][i] != "":                                  #탐지유형 열이 공백이 아니면 로직 태움
            if text_name != "":                                      #엔터로 탐지유형을 구분하지만 그냥 추가하면 처음에 공백이 들어가므로 변수에 값이 없는 첫 루프면 안넣게 조건 추가
                text_name = text_name + "\n"
            if text_name.endswith(df['탐지유형'][i]):                   #중복값을 추가하지 않기 위해 체크 로직
                print('error')
            else:
                text_name = text_name + df['탐지유형'][i] + "(" + df['탐지 이벤트'][i] + "건)" #문자열에 탐지유형 추가

        # 출발지 로직
        if df['공격대상'][i] != "":                                 #공격대상 열이 공백이 아닌 것만 체크
            if p.match(df['공격대상'][i]):                          #공격대상열에 정규표현식에 해당하는 것만 체크
                """if text_src != (df['공격대상'][i] + "(" + df.iat[i,5] + "), "): #중복제거 로직 아직 미완
                    text_src = text_src + df['공격대상'][i] + "(" + df.iat[i,5] + "), " """
                dup_src = text_src.split('(')
                if dup_src[len(dup_src)-2].endswith(df['공격대상'][i]):      #중복제거 로직 미완성
                    print('error')
                else:
                    text_src = text_src + df['공격대상'][i] + "(" + df.iat[i, 5] + "), "
            else:
                if df.iat[i, 7] == "" and text_src !="":            #탐지유형별로 출발지를 구분하기 위해 출발지 문자열이 공백이 아니며(처음공백 피하기 위해) 특정위치가 공백이면 엔터 추가하게 로직 구성
                    text_src = text_src[:-2]                        #마지막 ip는 ", "을 삭제
                    text_src = text_src + "\n"
        #목적지 로직
        if p.match(df.iat[i,7]):
            dup_dst = text_dst.split('(')
            if dup_dst[len(dup_dst)-2].endswith(df.iat[i,7]):
                print("error")
            else:
                text_dst = text_dst + df.iat[i,7] + "(tcp/" + df.iat[i,8] + "), "   #ip형식이면 추가
        else:
            if df.iat[i,7] == "" and text_dst != "":                            #목적지 문자열이 공백이 아니며 목적지 나열이 끝나면 공백추가
                text_dst = text_dst[:-2]                                        #마지막 ip는 ", "을 삭제
                text_dst = text_dst + "\n"

        #탐지시간 로직
        if r.match(df.iat[i,11]):
            text_date = text_date + df.iat[i,11][11:-2] + " | "
        else:
            if df.iat[i,11] == "" and text_date != "":
                text_date = text_date[:-3]
                text_date = text_date + "\n"

    text_src = text_src[:-2]
    text_dst = text_dst[:-2]        #출발지 목적지 마직막에 ", " 제거 탐지유형은 하나씩만 있으므로 없어도 됨
    text_date = text_date[:-3]

    list1 = text_name.split("\n")
    list2 = text_src.split("\n")
    list3 = text_dst.split("\n")    #문자열 엔터기준으로 나누기
    list4 = text_date.split("\n")

    excel_report = ""
    report_date = df.iat[4,11][:11]
    report_date = report_date.replace("-","/")
    for j in range(0,len(list1)):
        list4[j] = report_date + list4[j]#str(datetime.today().strftime("%Y/%m/%d")) + " " + list4[j]
        excel_report = excel_report + "<h1>" +list1[j] + "</h1><h2>" + list2[j] + "</h2><h2>" + list3[j] + "</h2><h2>" + list4[j] + "</h2>"
    html(excel_report)

    print(list1)
    print(list2)
    print(list3)
    print(list4)

def html(excelfile):
    #html_text = """<!DOCTYPE html><html><head><title>report</title><meta charset="UTF-8"></head><body>""" + excelfile.to_html() + """</body></html>"""
    html_text = """<!DOCTYPE html><html><head><title>report</title><meta charset="UTF-8"></head><body>""" + excelfile + """</body></html>"""
    html_file = open("/Users/aibikeiyeongeumboheom/Desktop/report.html",'w')
    html_file.write(html_text)
    html_file.close()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #print_hi('PyCharm')
    a = datetime.today() - timedelta(2)
    filename = "/Users/aibikeiyeongeumboheom/Desktop/탐지분석_"+a.strftime("%Y-%m-%d")+".xlsx"
    excel(filename)

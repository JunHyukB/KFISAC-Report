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
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

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
    now_date=""
    p = re.compile("[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+") #ip 정규표현식 적용
    r = re.compile("[0-9]+\-[0-9]+\-[0-9].+") #날짜 정규표현식

    for i in df.index:
        # 공격유형 로직
        if df['탐지유형'][i] != "":                                  #탐지유형 열이 공백이 아니면 로직 태움
            if text_name != "":                                      #엔터로 탐지유형을 구분하지만 그냥 추가하면 처음에 공백이 들어가므로 변수에 값이 없는 첫 루프면 안넣게 조건 추가
                text_name = text_name + "\n"
            if text_name.endswith(df['탐지유형'][i]):                   #중복값을 추가하지 않기 위해 체크 로직
                print('중복_name')
            else:
                text_name = text_name + df['탐지유형'][i] + "(" + df['탐지 이벤트'][i] + "건)" #문자열에 탐지유형 추가

        # 출발지 로직
        if df['공격대상'][i] != "":                                 #공격대상 열이 공백이 아닌 것만 체크
            if p.match(df['공격대상'][i]):                          #공격대상열에 정규표현식에 해당하는 것만 체크
                """if text_src != (df['공격대상'][i] + "(" + df.iat[i,5] + "), "): #중복제거 로직 아직 미완
                    text_src = text_src + df['공격대상'][i] + "(" + df.iat[i,5] + "), " """
                dup_src = text_src.split('(')
                if dup_src[len(dup_src)-2].endswith(df['공격대상'][i]) and not text_src.endswith("\n"):      #중복제거 로직 미완성, 출발지가 연속적이지 않으면 검사안됨
                    print('중복_src')
                    #print(dup_src)
                else:
                    text_src = text_src + df['공격대상'][i] + "(" + df.iat[i, 5] + "), "
            else:
                if df.iat[i, 7] == "" and text_src !="":            #탐지유형별로 출발지를 구분하기 위해 출발지 문자열이 공백이 아니며(처음공백 피하기 위해) 특정위치가 공백이면 엔터 추가하게 로직 구성
                    text_src = text_src[:-2]                        #마지막 ip는 ", "을 삭제
                    text_src = text_src + "\n"
        #목적지 로직
        if p.match(df.iat[i,7]):
            dup_dst = text_dst.split('(')
            if dup_dst[len(dup_dst)-2].endswith(df.iat[i,7]) and not text_dst.endswith("\n"):#마지막ip가 중복되면 추가안함,공격유형바뀌면 검사 예외
                print("중복_dst")
            else:
                text_dst = text_dst + df.iat[i,7] + "(tcp/" + df.iat[i,8] + "), "   #ip형식이면 추가
        else:
            if df.iat[i,7] == "" and text_dst != "":                            #목적지 문자열이 공백이 아니며 목적지 나열이 끝나면 공백추가
                text_dst = text_dst[:-2]                                        #마지막 ip는 ", "을 삭제
                text_dst = text_dst + "\n"
        #탐지시간 로직
        if r.match(df.iat[i,11]):
            """if text_date == "" or text_date.endswith("\n"):
                text_date = text_date + df.iat[i,11][:11]
                text_date = text_date.replace("-","/")
                print(text_date)""" #날짜 입력 로직이지만 정렬위해 폐기
            if text_date == "":
                now_date = now_date + df.iat[i,11][:11]
                now_date = now_date.replace("-", "/") #최초 1회 오늘 일자 저장
            text_date = text_date + df.iat[i,11][11:-2] + " | "
        else:
            if df.iat[i,11] == "" and text_date != "":
                text_date = text_date[:-3]
                text_date = text_date + "\n"

    text_src = text_src[:-2]
    text_dst = text_dst[:-1]        #출발지 목적지 마직막에 ", " 제거 탐지유형은 하나씩만 있으므로 없어도 됨
    text_date = text_date[:-3]

    list1 = text_name.split("\n")
    list2 = text_src.split("\n")
    list3 = text_dst.split("\n")    #문자열 엔터기준으로 나누기
    list4 = text_date.split("\n")

    """정렬로직 추가해야함"""


    for g in range(len(list4)):#일자 추가
        list4[g] = now_date + list4[g]

    """excel_report = ""
    report_date = df.iat[4,11][:11]
    report_date = report_date.replace("-","/")
    for j in range(0, len(list1)):
        list4[j] = report_date + list4[j]#str(datetime.today().strftime("%Y/%m/%d")) + " " + list4[j]
        excel_report = excel_report + "<p>" + "◼︎ " + list1[j] + "</p><table><tr><td>출발지</td><td>" + list2[j] + "</td></tr><tr><td>목적지</td><td>" + list3[j] + "</td></tr><tr><td>탐지시간</td><td>" + list4[j] + "</td></tr></table>"
    html(excel_report)

    print(list1)
    print(list2)
    print(list3)
    print(list4)"""
    call_html(list1,list2,list3,list4)

def call_html(list1,list2,list3,list4):
    html("""<!DOCTYPE html><html><head><title>report</title><meta charset="UTF-8"></head><body><table>""")
    for j in range(0, len(list1)):
        # print(list1[j])
        # print(list2[j])
        # print(list3[j])
        html("◼︎" + list1[j])
        html("<table border='1' width='700>")
        html("<tr height='30'><th rowspan='1'width='100' height='25'>출발지</th><td>" + list2[j] + "</td></tr>")
        html("<tr height='30'><th rowspan='1'>목적지</th><td>" + list3[j] + "</td></tr>")
        html("<tr height='30'><th rowspan='1'>탐지시간</th><td>"+list4[j]+"</td></tr>")
        html("<tr height='30'><th rowspan='1'>탐지결과</th><td>test</td></tr>")
        html("<tr height='30'><th rowspan='1'>분석내용</th><td>test</td></tr>")
        html("<tr height='30'><th rowspan='1'>조치결과</th><td>test</td></tr></table><p></p>")
    html("""</body></html>""")

def html(excelfile):
    #html_text = """<!DOCTYPE html><html><head><title>report</title><meta charset="UTF-8"></head><body>""" + excelfile.to_html() + """</body></html>"""
    html_text = ""
    html_text = html_text + excelfile

    #html_file = open("/Users/jun/Desktop/report.html",'w')
    html_file = open("/Users/jun/Desktop/report.html", 'a+')
    html_file.write(html_text)
    html_file.close()

"""
def tk_file():
    #files = filedialog.askopenfilename(initialdir="./",title="hi")
    #return files
    try:
        files = filedialog.askopenfilename(initialdir="./",title="hi")
        tk.messagebox.showinfo("kFisac Report", "파일 선택 성공!")
        #text.insert(1.0, files)
        return files
    except:
        tk.messagebox.showinfo("kFisac Report", "파일 에러 실패!")
"""

def tk_excel():
    try:
        files = filedialog.askopenfilename(initialdir="./", title="hi")
        excel(files)
        tk.messagebox.showinfo("kFisac Report", "성공!")
    except:
        tk.messagebox.showinfo("kFisac Report", "보고서 실패!")
def tk_console():
    vWindow = tk.Tk()
    vWindow.title("kFisac Report")
    vWindow.geometry("640x480+950+500")
    vWindow.resizable(False, False)

    label=tk.Label(vWindow, text="일일보고 생성기")
    label.pack()

    #button1 = tk.Button(vWindow, width=30, height=10, text="파일 선택", command=text.insert(1.0, tk_file))
    #button1.pack()
    #text = tk.Text(vWindow, width=42, height=1)
    #text.pack()

    button2 = tk.Button(vWindow, width=30, height=10, text="보고서 실행", command=tk_excel)
    button2.place(x=180,y=150)
    #button2.pack(side="Center")

    vWindow.mainloop()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #print_hi('PyCharm')
    a = datetime.today() - timedelta(4)
    #filename = "/Users/aibikeiyeongeumboheom/Desktop/탐지분석_"+a.strftime("%Y-%m-%d")+".xlsx"
    #filename = "/Users/jun/Desktop/탐지분석_"+a.strftime("%Y-%m-%d")+".xlsx"

    tk_console() #실행
    #filename = tk_file()
    #excel(filename)

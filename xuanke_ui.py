# -*- coding:utf-8 -*-
import bs4
import xml
import xlwt
from lxml import etree
from openpyxl import Workbook
from prettytable import PrettyTable
import requests
import tkinter as tk
from tkinter import *
import tkinter.messagebox as messagebox
import Tkinter, tkFileDialog
from tkinter.filedialog import askdirectory

def save(url,file):
# url = "http://sydney.edu.au/handbooks/engineering_PG/coursework/units_of_study/information_technology/information_technology_descriptions.shtml"
# url=raw_input()
    page = requests.get(url)

    html = page.text
    # result = etree.parse(html,etree.HTMLParser())
    total=[]
    type = []
    each=[]
    eachchoose=[]
    html2=etree.HTML(html)
    name = html2.xpath('//div[@class="uosList"]/div[@class="uosListTitle"]/strong/text()')
    # for i in name
    # #  print i.text
    choose = html2.xpath('//div[@class="visibleFields"]/span[@class="subInTitle"]/text()')

    details = html2.xpath('//div[@class="uosList"]//div[@class="visibleFields"]/text()')
    # number=1
    # for y in details:
    #     each.append(y)
    #     number=number+1;
    #     if(number%4==0):
    #         total.append(each);
    #         each=[];
    x=0
    for i in range(0,len(choose),1):

        if("Credit" in choose[i]):
            if(each!=[]):
                total.append(each)
                type.append(eachchoose)
            eachchoose=[]
            each=[]
            # print "\n"
            # print " Name: "+name[x];
            eachchoose.append("Name");
            each.append(name[x]);
            x+=1

            # print choose[i],details[i]

            each.append(details[i]);
            eachchoose.append((choose[i]));
        else:
            # print choose[i],details[i]
            each.append(details[i])
            eachchoose.append((choose[i]));
    # for i in range(0,len(name),1):
    #
    #     # l1=x.xpath('span[@class="subInTitle"]')
    #     print name[i],choose[i],details[i];
    # print type

    table = PrettyTable();
    wb = Workbook()
    ws = wb.create_sheet("comment", 0)
    ws.append(["Name", "Credit points:" ,"Session:","Classes:","Prerequisites:","Assumed knowledge:","Assessment:","Mode of delivery"])
    for numb in range(0,len(total),1):
        dic = dict(zip(type[numb], total[numb]));
        result = [];
        if(dic.has_key("Name")):
            result.append(dic["Name"]);
        else:
            result.append("");

        if (dic.has_key(" Credit points: ")):
            result.append(dic[" Credit points: "]);
        else:
            result.append("")

        if (dic.has_key(" Session: ")):
            result.append(dic[" Session: "]);
        else:
            result.append("")

        if (dic.has_key(" Classes: ")):
            result.append(dic[" Classes: "]);
        else:
            result.append("")

        if (dic.has_key(" Prerequisites: ")):
            result.append(dic[" Prerequisites: "]);
        else:
            result.append("")

        if(dic.has_key(" Assumed knowledge: ")):
            result.append(dic[" Assumed knowledge: "]);
        else:
            result.append("")

        if (dic.has_key(" Assessment: ")):
            result.append(dic[" Assessment: "]);
        else:
            result.append("")

        if (dic.has_key(" Mode of delivery: ")):
            result.append(dic[" Mode of delivery: "]);

        else:
            result.append("");


        ws.append(result)
        # print result


    # result
    wb.save(file)
    wb.close();


class Application:
    def __init__(self, master):
        self.master=master
        self.frame = tk.Frame(self.master)
        self.path = StringVar()
        self.Button3 = tk.Entry(self.frame)
        self.Button5 = tk.Entry(self.frame, textvariable=self.path)
        # self.pack()
        self.createWidgets()
        self.frame.pack(fill=BOTH,expand=1)

    def createWidgets(self):
        # self.nameInput = Entry(self)
        # self.nameInput.pack()

        Button2 = tk.Label(self.frame, text='对应科目网址')  # ui完成
        # Button3 = tk.Entry(self.frame)
        Button4 = tk.Label(self.frame, text='保存文件路径')
        # Button5 = tk.Entry(self.frame,textvariable = self.path)
        Button6 = tk.Button(self.frame,text="选择文件路径",command=self.selectPath);
        Button7 = tk.Button(self.frame, text="生成",command=self.generate)

        # Button1.pack(fill=BOTH,expand=1)
        Button2.pack(fill=BOTH, expand=1)
        self.Button3.pack(fill=BOTH, expand=1)
        Button4.pack(fill=BOTH, expand=1)
        self.Button5.pack(fill=BOTH, expand=1)
        Button6.pack(fill=BOTH, expand=1)
        Button7.pack(fill=BOTH, expand=1)




    def selectPath(self):
        path_ = askdirectory()
        self.path.set(path_+"/xuanke.xlsx")

    def generate(self):
        save(self.Button3.get(),self.Button5.get())



root=tk.Tk()
root.geometry('700x500')
app = Application(root)
# 设置窗口标题:

app.master.title('选课excel生成器')
# 主消息循环:
root.mainloop()


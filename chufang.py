# -*- coding: utf-8 -*-
"""
开药并打印处方;
v1.0

"""
import json
import os
import pyautogui
import sys
import tkinter as tk
import win32api
import win32print

from docx import Document


def read_data():
    with open('data_back.json', encoding='utf-8') as f:
        lines = f.readlines()

        if lines:
            last_line = lines[-1]

            info_dict = json.loads(last_line)

        f.close()
        return info_dict


def write_date(info_dict, name, age, weight, gender, rx_list):
    info_dict['name'] = name
    info_dict['age'] = age
    info_dict['weight'] = weight
    info_dict['rx'] = rx_list
    info_dict['gender'] = gender

    with open('data_back.json', 'a', encoding='utf-8') as f:
        f.write('\n')
        json.dump(info_dict, f, ensure_ascii=False)

        f.close()


drug_dic = {
    '止嗽口服液': 'zskfy',
    '赖氨酸磷酸氢钙': 'xks',
    '脑蛋白水解物口服液': 'wt',
    '头孢丙烯颗粒': 'yls',
    '双氯芬酸钾栓12.5mg(<16kg)': 'ptnx',
    '双氯芬酸钾栓50mg': 'ptnd',
    '硫酸亚铁糖浆': 'lsyt',
    '阿奇霉素颗粒(希舒美)': 'xsm'

}


class P(object):
    def __init__(self, w, a):
        self.w = w
        self.a = a

    def yls(self):
        if self.w in [12, 13, 14, 15]:
            dot = 0.125
        elif self.w in [x2 for x2 in range(16, 21)]:
            dot = 0.16
        elif self.w in [x2 for x2 in range(21, 25)]:
            dot = 0.2
        elif self.w in [x2 for x2 in range(25, 37)]:
            dot = 0.25
        elif self.w in [x2 for x2 in range(37, 50)]:
            dot = 0.375
        elif self.w in [x2 for x2 in range(50, 70)]:
            dot = 0.5
        return ['头孢丙烯颗粒(银力舒)*1盒',
                '                    sig: {} Bid po'.format(dot)]

    def zskfy(self):
        if self.a < 2:
            dot = 3
        elif 2 <= self.a <= 6:
            dot = 5
        elif 6 <= self.a <= 12:
            dot = 7.5
        else:
            dot = 10
        return ['止嗽口服液10ml*1盒',
                '              sig: {}ml Tid po'.format(dot)]

    def wt(self):
        if self.a < 2:
            dot = 3
        elif 2 <= self.a <= 6:
            dot = 5
        elif 6 <= self.a <= 11:
            dot = 7.5
        else:
            dot = 10
        return ['脑蛋白水解物口服液(万通)10ml*1盒',
                '              sig: {}ml Tid po'.format(dot)]

    def xks(self):
        if self.a < 2:
            dot = 0.5

        else:
            dot = 1
        return ['赖氨酸磷酸氢钙(小K斯)*3盒',
                '                sig: {}包 Tid po'.format(dot)]

    def ptnx(self):
        if 5 <= self.w < 7:
            dot = 5
        elif self.w in [x2 for x2 in range(7, 9)]:
            dot = 6
        elif self.w in [x2 for x2 in range(9, 10)]:
            dot = 8
        elif self.w in [x2 for x2 in range(10, 12)]:
            dot = 10
        elif self.w in [x2 for x2 in range(12, 16)]:
            dot = 12.5
        return ['双氯芬酸钾栓 12.5mg*1盒',
                '              sig: {}mg sos 肛塞'.format(dot)]

    def ptnd(self):
        if self.w in [x2 for x2 in range(16, 25)]:
            dot = 15
        elif self.w in [x2 for x2 in range(25, 35)]:
            dot = 25
        elif self.w in [x2 for x2 in range(35, 40)]:
            dot = 35
        elif self.w in [x2 for x2 in range(40, 50)]:
            dot = 40
        elif self.w >= 50:
            dot = 50
        return ['双氯芬酸钾栓 50mg*1盒',
                '              sig: {}mg sos 肛塞'.format(dot)]

    def lsyt(self):
        if self.a < 1:
            dot = 1.5
        elif 1 <= self.a <= 5:
            dot = 3
        elif 6 <= self.a <= 12:
            dot = 5
        return ['硫酸亚铁糖浆 *1盒',
                '              sig: {}ml tid po'.format(dot)]

    def xsm(self):
        if self.w < 10:
            dot = self.w * 0.01
        elif self.w in [10, 11, 12, 13, 14]:
            dot = 0.1
        elif self.w in [x2 for x2 in range(15, 20)]:
            dot = 0.15
        elif self.w in [x2 for x2 in range(20, 25)]:
            dot = 0.2
        elif self.w in [x2 for x2 in range(25, 30)]:
            dot = 0.25
        elif self.w in [x2 for x2 in range(30, 40)]:
            dot = 0.3
        elif self.w in [x2 for x2 in range(40, 50)]:
            dot = 0.4
        else:
            dot = 0.5
        return ['阿奇霉素颗粒(希舒美)0.1*1盒',
                '              sig: {} qd po (吃3天停4天)'.format(dot)]


class Ky(object):
    drug = ('阿奇霉素颗粒(希舒美)',
            '止嗽口服液',
            '赖氨酸磷酸氢钙',
            '脑蛋白水解物口服液',
            '头孢丙烯颗粒',
            '双氯芬酸钾栓12.5mg(<16kg)',
            '双氯芬酸钾栓50mg',
            '硫酸亚铁糖浆')
    drug_sd = []

    def __init__(self, root, name, gender, age, weight):
        self.root = root
        self.name = name
        self.gender = gender
        self.age = age
        self.weight = weight
        self.root.wm_attributes('-topmost', 1)
        self.root.overrideredirect(True)

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        ww = 500
        wh = 650
        self.root.geometry(
            "%dx%d+%d+%d" %
            (ww, wh, (sw - ww) / 2, (sh - wh) / 2))
        self.root.after(1, lambda: self.root.focus_force())
        self.val = tk.StringVar()

        self.f1 = tk.Frame(self.root)
        self.f1.grid(pady=25)

        tk.Label(self.f1, text='姓名').grid(row=0, column=0)
        self.e1 = tk.Entry(self.f1)
        self.e1.insert(0, str(self.name))
        self.e1.focus_set()
        self.e1.grid(row=0, column=1)
        tk.Label(self.f1, text='性别').grid(row=0, column=2)
        self.e2 = tk.Entry(self.f1)
        self.e2.insert(0, str(self.gender))
        self.e2.grid(row=0, column=3)
        tk.Label(self.f1, text='年龄').grid(row=1, column=0)
        self.e3 = tk.Entry(self.f1)
        self.e3.insert(0, str(self.age))
        self.e3.grid(row=1, column=1)
        tk.Label(self.f1, text='体重').grid(row=1, column=2)
        self.e4 = tk.Entry(self.f1)
        self.e4.insert(0, str(self.weight))
        self.e4.grid(row=1, column=3)

        self.f = tk.Frame(self.root)
        self.f.grid(padx=40)
        self.croll = tk.Scrollbar(self.f, orient=tk.VERTICAL)
        self.croll.grid(row=0, column=1, sticky=tk.N + tk.S)
        self.val.set(self.drug)
        self.lb = tk.Listbox(self.f,
                             listvariable=self.val,
                             selectmode=tk.MULTIPLE,
                             bd=1,
                             font=('微软雅黑', 10),
                             selectbackground='brown',
                             width=50,
                             height=25,
                             yscrollcommand=self.croll.set)
        self.lb.grid(row=0, column=0)
        self.croll.config(command=self.lb.yview)

        self.fdrug = tk.Frame(self.root)
        self.fdrug.grid(pady=10)

        self.bt = tk.Button(self.fdrug,
                            text='确 定',
                            command=self.ok,
                            width=12,
                            height=2,
                            activebackground='grey',
                            relief='groove')
        self.bt.grid(row=0, column=0)
        self.btc = tk.Button(self.fdrug,
                             text='取 消',
                             command=self.can,
                             width=12,
                             height=2,
                             activebackground='grey',
                             relief='groove')
        self.btc.grid(row=0, column=1)

    def ok(self):
        self.name = self.e1.get()
        self.gender = self.e2.get()
        self.age = int(self.e3.get())
        self.weight = int(self.e4.get())

        for each in self.lb.curselection():
            self.drug_sd.append(self.lb.get(each))
        self.root.destroy()

    def can(self):
        self.root.destroy()
        sys.exit(0)


root = tk.Tk()
p_info = read_data()
app = Ky(
    root,
    p_info['name'],
    p_info['gender'],
    p_info['age'],
    p_info['weight'])
root.mainloop()

drug_code = []
text_to_w = []
if app.drug_sd:
    for each_d in app.drug_sd:
        drug_code.append(drug_dic[each_d])

if drug_code:
    a = P(app.weight, app.age)
    for each_c in drug_code:
        aa = getattr(a, each_c)
        text_to_w.append(aa())

text_1 = [app.name, app.gender, str(app.age)]
text_2 = text_to_w
write_date(p_info, app.name, app.age, app.weight, app.gender, text_to_w)
text_show = ''
for each in text_2:
    text_show += each[0] + ' ' + each[1].lstrip('              ') + '\n'
pyautogui.alert(text=str(text_show), title=str(app.name) +
                ' ' + str(app.age) + '岁 ' + str(app.weight) + 'kg:')
d = Document('.\\chufang.docx')

t = d.tables

t1 = t[0]
for i in range(1, len(t1.columns), 2):
    t1.cell(0, i).text = ''
    run = t1.cell(0, i).paragraphs[0].add_run(text_1[(i - 1) // 2])
    run.font.name = '宋体'
    run.font.size = 240000

t2 = t[1]
l = 0
for row in range(0, len(t2.rows)):
    t2.cell(0, row).text = ''
for j in range(0, len(text_2)):

    for k in range(0, 2):
        run1 = t2.cell(0, l).paragraphs[0].add_run(text_2[j][k])
        run1.font.name = '宋体'
        run1.font.size = 240000
        l += 1

d.save('.\\temp.docx')


def print_file(filename):
    #    open(filename,'r')
    win32api.ShellExecute(
        0,
        'print',
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        '.',
        0)


print_file(os.path.abspath('.\\temp.docx'))
os.remove('.\\temp.docx')

# -*- coding: utf-8 -*-
import json
import os
import sys
import tkinter as tk
import win32api
import win32print

from docx import Document


def write_date(info_dict, name, age, weight, gender, rx_list):
    info_dict['name'] = name
    info_dict['age'] = age
    info_dict['weight'] = weight
    info_dict['rx'] = rx_list
    info_dict['gender'] = gender

    with open('data_back.json', 'a', encoding='utf-8') as f:
        f.write('\n')
        json.dump(info_dict, f, ensure_ascii=False, sort_keys=True)

        f.close()


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
        else:
            dot = '  '
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
                '                sig: {}包 Bid po'.format(dot)]

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
        elif self.w in [x2 for x2 in range(16, 25)]:
            dot = 15
        elif self.w in [x2 for x2 in range(25, 35)]:
            dot = 25
        elif self.w in [x2 for x2 in range(35, 40)]:
            dot = 35
        elif self.w in [x2 for x2 in range(40, 50)]:
            dot = 40
        elif self.w >= 50:
            dot = 50
        else:
            dot = '  '
        return ['双氯芬酸钾栓 12.5mg*1盒',
                '              sig: {}mg sos 肛塞'.format(dot)]

    def ptnd(self):
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
        elif self.w in [x2 for x2 in range(16, 25)]:
            dot = 15
        elif self.w in [x2 for x2 in range(25, 35)]:
            dot = 25
        elif self.w in [x2 for x2 in range(35, 40)]:
            dot = 35
        elif self.w in [x2 for x2 in range(40, 50)]:
            dot = 40
        elif self.w >= 50:
            dot = 50
        else:
            dot = '  '
        return ['双氯芬酸钾栓 50mg*1盒',
                '              sig: {}mg sos 肛塞'.format(dot)]

    def lsyt(self):
        if self.a < 1:
            dot = 1.5
        elif 1 <= self.a <= 5:
            dot = 3
        elif 6 <= self.a <= 12:
            dot = 5
        else:
            dot = '  '
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
    drug = {
        '止嗽口服液': 'zskfy',
        '赖氨酸磷酸氢钙': 'xks',
        '脑蛋白水解物口服液': 'wt',
        '头孢丙烯颗粒': 'yls',
        '双氯芬酸钾栓12.5mg(<16kg)': 'ptnx',
        '双氯芬酸钾栓50mg': 'ptnd',
        '硫酸亚铁糖浆': 'lsyt',
        '阿奇霉素颗粒(希舒美)': 'xsm'

    }

    def __init__(self, root):
        self.root = root
        self.p_info = self.read_data()
        self.name, self.gender, self.age, self.weight = [
            self.p_info['name'],
            self.p_info['gender'],
            self.p_info['age'],
            self.p_info['weight']
        ]
        self.root.wm_attributes('-topmost', 1)
        self.root.overrideredirect(True)

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        ww = 650
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
        self.e1.bind('<FocusIn>', self.on_select)
        self.e1.focus_set()
        self.e1.grid(row=0, column=1)
        tk.Label(self.f1, text='性别').grid(row=0, column=2)
        self.e2 = tk.Entry(self.f1)
        self.e2.insert(0, str(self.gender))
        self.e2.bind('<FocusIn>', self.on_select)
        self.e2.grid(row=0, column=3)
        tk.Label(self.f1, text='年龄').grid(row=1, column=0)
        self.e3 = tk.Entry(self.f1)
        self.e3.insert(0, str(self.age))
        self.e3.bind('<FocusIn>', self.on_select)
        self.e3.grid(row=1, column=1)
        tk.Label(self.f1, text='体重').grid(row=1, column=2)
        self.e4 = tk.Entry(self.f1)
        self.e4.insert(0, str(self.weight))
        self.e4.bind('<FocusIn>', self.on_select)
        self.e4.grid(row=1, column=3)

        self.f = tk.Frame(self.root)
        self.f.grid(padx=40)
        self.croll = tk.Scrollbar(self.f, orient=tk.VERTICAL)
        self.croll.grid(row=0, column=1, sticky=tk.N + tk.S)
        self.val.set(tuple(sorted(self.drug)))
        self.lb = tk.Listbox(self.f,
                             listvariable=self.val,
                             selectmode=tk.MULTIPLE,
                             bd=1,
                             font=('微软雅黑', 10),
                             selectbackground='brown',
                             width=50,
                             height=9,
                             yscrollcommand=self.croll.set)
        self.lb.bind('<<ListboxSelect>>', self.show_info)
        self.lb.grid(row=0, column=0)
        self.croll.config(command=self.lb.yview)
        tk.Label(root, text='开药情况预览（可以直接修改预览信息后打印）').grid(pady=10)
        # text
        bg_frame_text = tk.Frame(root)
        bg_frame_text.grid(padx=30)
        scroll_bar_text = tk.Scrollbar(bg_frame_text, orient=tk.VERTICAL)
        scroll_bar_text.grid(row=0, column=1, sticky=tk.N + tk.S)
        self.text = tk.Text(bg_frame_text, width=80, height=20, yscrollcommand=scroll_bar_text.set)
        self.text.grid(row=0, column=0)
        scroll_bar_text.config(command=self.text.yview)

        self.fdrug = tk.Frame(self.root)
        self.fdrug.grid(pady=10)

        self.bt = tk.Button(self.fdrug,
                            text='打 印',
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

    def read_data(self):
        with open('data_back.json', encoding='utf-8') as f:
            lines = f.readlines()

            if lines:
                last_line = lines[-1]

                info_dict = json.loads(last_line)
                return info_dict
            else:
                return {"age": 0, "name": "", "weight": 0, "gender": "", "rx": []}

    def ok(self):
        self.name = self.e1.get()
        self.gender = self.e2.get()
        self.age = int(self.e3.get())
        self.weight = int(self.e4.get())
        res = self.text.get(1.0, tk.END)
        self.text_to_w = res.split('\n')

        self.root.destroy()

    def can(self):
        self.root.destroy()
        sys.exit(0)

    def on_select(self, event):
        # event.widget 即触发事件的控件，可以直接用它的方法如.select_range()等
        event.widget.select_range(0, tk.END)

    def show_info(self, event):
        try:
            self.name = self.e1.get()
            self.gender = self.e2.get()
            self.age = int(self.e3.get())
            self.weight = int(self.e4.get())
            self.drug_sd = []

            for each in self.lb.curselection():
                self.drug_sd.append(self.lb.get(each))
            drug_code = []

            if self.drug_sd:
                for each_d in self.drug_sd:
                    drug_code.append(self.drug[each_d])
            text_to_show = []
            if drug_code:

                a = P(self.weight, self.age)
                for each_c in drug_code:
                    aa = getattr(a, each_c)

                    text_to_show.append('\n'.join(aa()))

            text_data = '\n'.join(text_to_show)
            self.text.delete(1.0, 'end')
            self.text.insert('end', text_data)
        except IndexError:
            pass


def to_word(text_header, text_body):
    d = Document('.\\chufang.docx')

    t = d.tables

    t1 = t[0]
    for i in range(1, len(t1.columns), 2):
        t1.cell(0, i).text = ''
        run = t1.cell(0, i).paragraphs[0].add_run(text_header[(i - 1) // 2])
        run.font.name = '宋体'
        run.font.size = 240000

    t2 = t[1]
    l = 0
    for row in range(0, len(t2.rows)):
        t2.cell(0, row).text = ''
    for j in range(0, len(text_body)):
        run1 = t2.cell(0, l).paragraphs[0].add_run(text_body[j])
        run1.font.name = '宋体'
        run1.font.size = 240000
        l += 1

    d.save('.\\temp.docx')


def print_file(filename):
    win32api.ShellExecute(
        0,
        'print',
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        '.',
        0)
    os.remove(filename)


if __name__ == '__main__':
    while True:
        root = tk.Tk()
        app = Ky(root)
        root.mainloop()

        p_info = {}
        write_date(p_info, app.name, app.age, app.weight, app.gender, app.text_to_w)

        text_head = [app.name, app.gender, str(app.age)]
        text_body = app.text_to_w
        to_word(text_head, text_body)

        print_file(os.path.abspath('.\\temp.docx'))

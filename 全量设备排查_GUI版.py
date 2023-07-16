import os
import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import *

win = tk.Tk()
win.title("离线可视化分析软件Beta Code by LYC")
win.geometry('300x200+800+400')
win.iconphoto(False, tk.PhotoImage(file='监控摄像头_surveillance-cameras.png'))
frame = tk.Frame(win)
# 周数输入控件
label_1 = tk.Label(win, text="1.请输入周数")
label_1.grid(row=0, column=0)
text_v = tk.StringVar()
entry_1 = tk.Entry(win, textvariable=text_v, width=30)
entry_1.grid(row=1, column=0)


def visualise_analysis():
    week = entry_1.get()
    visualise_data_files = os.listdir(os.getcwd() + "\\离线台账\\")
    offline_visualise_list = pd.DataFrame(
        columns=['设备名称', '设备厂家'])
    for i in range(len(visualise_data_files)):
        if i == 0:
            visualise_data = pd.read_excel(os.getcwd() + "\\离线台账\\" + visualise_data_files[i], header=1)
            visualise_offline_data = list(visualise_data[visualise_data.是否在线 == "离线"].设备名称)
            offline_visualise_list.设备名称 = visualise_data.设备名称
            offline_visualise_list.设备厂家 = visualise_data.设备厂家
            offline_visualise_list.insert(loc=len(offline_visualise_list.columns), column='2', value=0)
            for k in range(len(offline_visualise_list)):
                if list(offline_visualise_list.设备名称)[k] in visualise_offline_data:
                    offline_visualise_list.iloc[k, i + 2] = 1
                else:
                    offline_visualise_list.iloc[k, i + 2] = 0
        else:
            visualise_data = pd.read_excel(os.getcwd() + "\\离线台账\\" + visualise_data_files[i], header=1)
            visualise_offline_data = list(visualise_data[visualise_data.是否在线 == "离线"].设备名称)
            offline_visualise_list.insert(loc=len(offline_visualise_list.columns),
                                          column=str(len(offline_visualise_list.columns)), value=0)
            for k in range(len(offline_visualise_list)):
                if list(offline_visualise_list.设备名称)[k] in visualise_offline_data:
                    offline_visualise_list.iloc[k, i + 2] = 1
                else:
                    offline_visualise_list.iloc[k, i + 2] = 0

    offline_visualise_list.fillna(value=0, inplace=True)
    offline_visualise_list.to_excel(os.getcwd() + "\\分析结果\\" + week + "周在线情况.xlsx")
    showinfo('提示', '分析完成')


label_4 = tk.Label(win, text="2.请点击开始")
label_4.grid(row=0, column=4)
Start_button = tk.Button(win, text='开始分析', command=visualise_analysis)
Start_button.grid(row=1, column=4)

win.mainloop()

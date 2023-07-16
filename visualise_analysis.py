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
        columns=['设备名称', '设备厂家', list(range(len(visualise_data_files)))])
    for i in range(len(visualise_data_files)):
        if i == 0:
            visualise_data = pd.read_excel(os.getcwd() + "\\离线台账\\" + visualise_data_files[i], header=1)
            visualise_offline_data = visualise_data[visualise_data.是否在线 == "离线"].reset_index()
            offline_visualise_list.设备名称 = visualise_offline_data.设备名称
            offline_visualise_list.设备厂家 = visualise_offline_data.设备厂家
            offline_visualise_list.iloc[:, 2] = 1
        else:
            visualise_data = pd.read_excel(os.getcwd() + "\\离线台账\\" + visualise_data_files[i], header=1)
            visualise_offline_data = visualise_data[visualise_data.是否在线 == "离线"].reset_index()
            for j in range(len(visualise_offline_data.设备名称)):
                if list(visualise_offline_data.设备名称)[j] in list(offline_visualise_list.设备名称):
                    offline_visualise_list.iloc[j, i + 1] = 1
                else:
                    offline_visualise_list.loc[len(offline_visualise_list.index)] = [
                        list(visualise_offline_data.设备名称)[j], list(visualise_offline_data.设备厂家)[j],
                        0,
                        0, 0, 0, 0, 0, 0]
                    offline_visualise_list.iloc[len(offline_visualise_list.index) - 1, i + 1] = 1
    #offline_visualise_list = offline_visualise_list.fillna(value=0)
    #offline_visualise_list_48h = offline_visualise_list.copy()
    #for i in range(1, len(offline_visualise_list_48h.columns) - 1):
#        offline_visualise_list_48h.iloc[:, i] = offline_visualise_list_48h.iloc[:, i] + offline_visualise_list_48h.iloc[
 #                                                                                       :,
  #                                                                                  i + 1]

    #offline_visualise_list_48h.columns = ['设备名称', '第一个48小时', '第二个48小时', '第三个48小时', '第四个48小时',
   #                                       '第五个48小时', '第六个48小时', '周五不在线']
    #offline_visualise_list_48h = offline_visualise_list_48h.drop('周五不在线', axis=1)
    offline_visualise_list.to_excel(os.getcwd() + "\\分析结果\\" + week + "周在线情况.xlsx")
    #offline_visualise_list_48h.to_excel(os.getcwd() + "\\分析结果\\" + week + "周分析结果.xlsx")
    showinfo('提示', '分析完成')


label_4 = tk.Label(win, text="2.请点击开始")
label_4.grid(row=0, column=4)
Start_button = tk.Button(win, text='开始分析', command=visualise_analysis)
Start_button.grid(row=1, column=4)

win.mainloop()

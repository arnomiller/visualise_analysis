import os
import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import *

week = 'test'
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
# offline_visualise_list_48h.to_excel(os.getcwd() + "\\分析结果\\" + week + "周分析结果.xlsx")
showinfo('提示', '分析完成')

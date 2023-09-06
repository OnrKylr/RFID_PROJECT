# -*- coding: utf-8 -*-

import serial
import openpyxl
import keyboard  # keyboard kütüphanesi
import pandas as pd
import json
from datetime import datetime


ser = serial.Serial('COM5', 9600)

#reg_data = {}
#df = pd.read_json("C:/Users/MSI/Desktop/export_dataframe.json")
with open("/export_dataframe.json") as f:
    data = json.load(f)
    df = pd.DataFrame(data)
column_name = "RFIDs"
dictionary = df[column_name].unique().tolist()

while True:
    line = ser.readline().decode("utf-8" , errors="backslashreplace").strip() #encode edilemeyen karakter backslash ile degistirilerek cozuldu
    if line in dictionary:
        index = dictionary.index(line)
        row = df.iloc[index]
        print(row)
    else:
        print("yok aga")
    if line == "q" :
        break


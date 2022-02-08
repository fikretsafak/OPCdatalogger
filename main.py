import OpenOPC
import csv
import time
import pandas as pd
import os
import numpy as np


reports = []
Deger = []
zaman = []
config = open('config.txt','r')

f=open(time.strftime('%d%m%Y')+"_RAPORLAMA.csv","a")
writer=csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
line = config.readline()

while line!='':
    if "[" in line:
        print ("Config reading..")
        report_name = line.split("[")[1].split("]")[0].strip()
        path = config.readline().split("=")[1].strip()
        dist_list = config.readline().split("=")[1].strip()
        dist_time = config.readline().split("=")[1].strip()
        read_time = config.readline().split("=")[1].strip()
        OPC_Name = config.readline().split("=")[1].strip()
        Tag = config.readline().split("=")[1].strip()
        Tags=Tag.split(",")
        reports.append([report_name, path, dist_list, dist_time,read_time,OPC_Name,Tags])
    line = config.readline()

opc = OpenOPC.client()
opc.connect(OPC_Name)
uzunluk=len(Tags)
writer.writerow(Tags)

results = opc.read(Tags)

for j in range(uzunluk):
    Deger.append([results[j][1]])
for x in range(20):
    results = opc.read(Tags)
    for k in range(uzunluk):
        Deger[k]=results[k][1]
    fc = str(Deger)
    Deger[0]=str(time.strftime('%d/%m/%Y %X'))
    csv =Deger
    print(csv)
    writer.writerow(csv)
    time.sleep(float(read_time))   
f.close()

csv_export = pd.read_csv(time.strftime('%d%m%Y')+"_RAPORLAMA.csv")
excel_export = pd.ExcelWriter(time.strftime('%d%m%Y')+"_RAPORLAMA.xlsx")
csv_export.to_excel(excel_export, index = False)
  
excel_export.save()

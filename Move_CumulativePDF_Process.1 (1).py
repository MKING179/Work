# -*- coding: utf-8 -*-
"""
Created on Tue Apr  4 21:41:35 2023

@author: cannk
"""
import shutil
import glob
import re
import os 
from pathlib import Path
from datetime import datetime, timedelta
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import math

##UPDATE THE SOURCE AND DESTINATION PATHS - FILE PREFIX IS DYNAMIC
SourcePath = "C:\\Users\\mking\\Downloads"
DestinationPath = "C:\\Users\\mking\\OneDrive\\Rep 451\\"

#FilePrefix= "Need"

#function blocks that are needed for the "need" calculations 
#this is also construct for the check boxes
def assign_oil_need(row):
    if row <= 500:
         return int(500 - row)
    elif row > 500 and row <= 1499:
         return int(1500 - row)
    elif row > 1499 and row <= 2499:
         return int(2500 - row)
    elif row > 2499 and row <= 3499:
         return int(3500 - row)
    elif row > 3499 and row <= 4499:
         return int(4500 - row)
    else:
         return 0
     
def assign_tire_need(row):
    if row <= 2000:
         return int(2001 - row)
    elif row > 2001 and row <= 4000:
         return int(4001 - row)
    elif row > 4001 and row <= 6001:
         return int(6001 - row)
    elif row > 6001:
         return 0
    else:
         return 0

def assign_battery_need(row):
    if row <= 500:
         return int(500 - row)
    elif row > 500 and row <= 1499:
         return int(1500 - row)
    elif row > 1499 and row <= 2499:
         return int(2500 - row)
    elif row > 2499 and row <= 3499:
         return int(3500 - row)
    elif row > 3499 and row <= 4499:
         return int(4500 - row)
    else:
         return 0

def assign_brake_need(row):
    if row <= 500:
         return int(500 - row)
    elif row > 500 and row <= 1499:
         return int(1500 - row)
    elif row > 1499 and row <= 2499:
         return int(2500 - row)
    elif row > 2499 and row <= 3499:
         return int(3500 - row)
    elif row > 3499 and row <= 4499:
         return int(4500 - row)
    else:
         return 0

def assign_helmet_need(row):
    if row <= 500:
         return int(500 - row)
    elif row > 500 and row <= 1499:
         return int(1500 - row)
    elif row > 1499 and row <= 2499:
         return int(2500 - row)
    elif row > 2499 and row <= 3499:
         return int(3500 - row)
    elif row > 3499 and row <= 4499:
         return int(4500 - row)
    else:
         return 0
     
def assign_apparel_need_2500(row):
    if row <= 2500:
         return int(2500 - row)
    elif row >= 5000 and row < 7500:
         return int(7500 - row)
    elif row >= 7500 and row < 10000:
         return int(10000 - row)
    elif row >= 10000:
         return int(0 - row)
    else:
         return 0
     
def assign_apparel_need_1500(row):
    if row < 1500:
        return int(1500 - row)
    elif row >= 1500 and row <2500:
        return int(2500 - row)
    elif row >= 2500 and row < 5000:
         return int(5000 - row)
    elif row >= 5000 and row < 7500:
         return int(7500 - row)
    elif row >= 7500 and row < 10000:
         return int(10000 - row)
    elif row > 10000:
         return int(0 - row)
    else:
         return 0

def assign_apparel_need_slippery(row):
    if row < 1500:
        return int(1500 - row)
    elif row >= 1500 and row <2500:
        return int(2500 - row)
    elif row >= 2500 and row < 5000:
         return int(5000 - row)
    elif row >= 5000 and row < 7500:
         return int(7500 - row)
    elif row >= 7500:
         return int(0 - row)
    else:
         return 0 

print(SourcePath)
print(DestinationPath)

directoryList = os.listdir(SourcePath)

for f in glob.glob(DestinationPath+'\\*\\*\\', recursive=True):
    #print(f)
    for files in os.listdir(SourcePath):
        fileMatch = re.compile(r"^.*(?=(\_))") #Folder Finder
        res = fileMatch.search(files)
       ## print(res)
        if res:
            if res.group() in str(f):
                print((res.group()))
                OutputFolder = (res.group())
                visitDate = datetime.now()+timedelta(1)
                visitDateFriendly = visitDate.strftime('%m%d%Y')
                Path(f+visitDateFriendly+'\\').mkdir(parents=True, exist_ok=True)
                Dest = f+visitDateFriendly+'\\'+files
                Sour = SourcePath+'\\'+files
                os.rename(Sour, Dest)
                
                if "CumulativeReport" in files: 
                    reader = PyPDF2.PdfReader(open(f + "\\CumulativeEditable.pdf", "rb"))          
                    # Count number of pages in our pdf file
                    number_of_pages = len(reader.pages)
                    # print number of pages in the pdf file
                    print("Number of pages in this pdf: " + str(number_of_pages))
                    
                    fields = reader.get_form_text_fields()
                    fields == {"key": "value", "key2": "value2"}
                    
                    boxes = reader.get_fields('/Btn')
                    print(f+visitDateFriendly +"\\" + (res.group())+ "_CumulativeReport.xlsx")
                    ##print(boxes)
                    excel_file = pd.ExcelFile(f+visitDateFriendly +"\\" + (res.group())+ "_CumulativeReport.xlsx")
                    print(excel_file)
                    sheetname = excel_file.sheet_names
                    
                    df = pd.DataFrame(sheetname)
                    
                    for i in df.index:
                        if 'Cumulative Tire Tube' in df[0][i]:
                            tire_df = pd.read_excel(excel_file, sheet_name=df[0][i])  
                        elif 'Cumulative Apparel' in df[0][i]:
                            apparel_df = pd.read_excel(excel_file, sheet_name=df[0][i]) 
                        elif 'Cumulative Battery' in df[0][i]:
                            battery_df = pd.read_excel(excel_file, sheet_name=df[0][i]) 
                        elif 'Cumulative Brake' in df[0][i]:
                            brake_df = pd.read_excel(excel_file, sheet_name=df[0][i]) 
                        elif 'Cumulative Helmet' in df[0][i]:
                            helmet_df = pd.read_excel(excel_file, sheet_name=df[0][i])
                        elif 'Cumulative Oil-Chemical' in df[0][i]:
                            oil_df = pd.read_excel(excel_file, sheet_name=df[0][i]) 
                        else: 
                            print('Moving On')
                    
                    #locate the correct columns for calcs
                    tire = tire_df.iloc[[0],[0, 1, 2, 4, 16]]
                    apparel = apparel_df.iloc[:,[0, 1, 2, 4, 16]]
                    battery = battery_df.iloc[[0],[0, 1, 2, 4, 16]]
                    brake = brake_df.iloc[[0],[0, 1, 2, 4, 16]]
                    helmet = helmet_df.iloc[[0],[0, 1, 2, 4, 16]]
                    oil = oil_df.iloc[[0],[0, 1, 2, 4, 16]]
                    
                    #turn nulls into 0's for calculation work
                    oil_n = oil.fillna(0)
                    tire_n = tire.fillna(0)
                    apparel_n = apparel.fillna(0)
                    battery_n = battery.fillna(0)
                    brake_n = brake.fillna(0)
                    helmet_n = helmet.fillna(0)
                    
                    #turn 0.0 into 0's for calculation work
                    oil0 = oil_n.round(0)
                    tire0 = tire_n.round(0)
                    apparel0 = apparel_n.round(0)
                    battery0 = battery_n.round(0)
                    brake0 = brake_n.round(0)
                    helmet0 = helmet_n.round(0)
                    
                    print(brake0)
                    
                    #criterion = df2['a'].map(lambda x: x.startswith('t'))
                    apparel_index2500= apparel0['BRAND'].map(lambda x: x.startswith('ALPINESTARS'))
                    apparel_2500_alp = apparel0[apparel_index2500]
                    apparel_2500_alp = apparel_2500_alp.append(apparel_2500_alp.sum(numeric_only=True), ignore_index=True)
                    apparel_2500_alp = apparel_2500_alp[apparel_2500_alp['BRAND'].isnull()]
                    print(apparel_2500_alp)
                    
                    
                    apparel_index2500i= apparel0['BRAND'].map(lambda x: x.startswith('ICON'))
                    apparel_2500_icon = apparel0[apparel_index2500i]
                    
                    apparel_index1500= apparel0['BRAND'].map(lambda x: x.startswith('ARCTIVA'))
                    apparel_1500 = apparel0[apparel_index1500]
                    
                    apparel_index1500z= apparel0['BRAND'].map(lambda x: x.startswith('Z1R'))
                    apparel_1500z = apparel0[apparel_index1500z]
                    
                    apparel_index1500t= apparel0['BRAND'].map(lambda x: x.startswith('THOR'))
                    apparel_1500t = apparel0[apparel_index1500t]
                    
                    apparel_index1500m= apparel0['BRAND'].map(lambda x: x.startswith('MOOSE'))
                    apparel_1500m = apparel0[apparel_index1500m]
                    
                    apparel_index1500Slip= apparel0['BRAND'].map(lambda x: x.startswith('SLIPPERY'))
                    apparel_Slippery = apparel0[apparel_index1500Slip]
                     
                    oil0['Need']=oil0['TY'].apply(assign_oil_need) 
                    tire0['Need']=tire0['TY'].apply(assign_tire_need) 
                    apparel_2500_alp['Need']=apparel_2500_alp['TY'].apply(assign_apparel_need_2500) 
                    apparel_2500_icon['Need']=apparel_2500_icon['TY'].apply(assign_apparel_need_2500) 
                    apparel_1500['Need']=apparel_1500['TY'].apply(assign_apparel_need_1500) 
                    apparel_1500z['Need']=apparel_1500z['TY'].apply(assign_apparel_need_1500) 
                    apparel_1500t['Need']=apparel_1500t['TY'].apply(assign_apparel_need_1500) 
                    apparel_1500m['Need']=apparel_1500m['TY'].apply(assign_apparel_need_1500)  
                    apparel_Slippery['Need']=apparel_Slippery['TY'].apply(assign_apparel_need_slippery) 
                    battery0['Need']=battery0['TY'].apply(assign_battery_need) 
                    brake0['Need']=brake0['TY'].apply(assign_brake_need)
                    helmet0['Need']=helmet0['TY'].apply(assign_helmet_need)  
                    
                    need_oil = int(oil0.iloc[0][2] - oil0.iloc[0][4] + oil0['Need'])
                    need_tire = int(tire0.iloc[0][2] - tire0.iloc[0][4] + tire0['Need'])
                    need_battery = int(battery0.iloc[0][2] - battery0.iloc[0][4] + battery0['Need'])
                    need_brake = int(brake0.iloc[0][2] - brake0.iloc[0][4] + brake0['Need'])
                    need_helmet = int(helmet0.iloc[0][2] - helmet0.iloc[0][4] + helmet0['Need'])
                    
                    writer = PdfWriter()
                    writer.add_page(reader.pages[0])
                    writer.add_page(reader.pages[1])
                    
                    ##template_pdf = PdfReader(reader)
                    if len(apparel_2500_alp) > 0:
                        need_apparel_alpine = int(apparel_2500_alp.iloc[0][2] - apparel_2500_alp.iloc[0][4] + apparel_2500_alp['Need'])
                    #Alpinestars
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Astars": apparel_2500_alp.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Astars": apparel_2500_alp.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Astars": apparel_2500_alp.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Astars": apparel_2500_alp.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Astars": apparel_2500_alp.iloc[0][4]})
                    else: 
                        print('No Alpinestars in the report')
                    if len(apparel_2500_icon) > 0:   
                        need_apparel_icon = int(apparel_2500_icon.iloc[0][2] - apparel_2500_icon.iloc[0][4] + apparel_2500_icon['Need'])
                    ##Icon
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Icon": apparel_2500_icon.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Icon": apparel_2500_icon.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Icon": apparel_2500_icon.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Icon": apparel_2500_icon.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Icon": apparel_2500_icon.iloc[0][4]})
                    
                    else: 
                        print('No Icon in the report')
                    if len(apparel_1500) > 0:
                        need_apparel_arctiva = int(apparel_1500.iloc[0][2] - apparel_1500.iloc[0][4] + apparel_1500['Need'])
                    ##Arctiva
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Arctiva": apparel_1500.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Arctiva": apparel_1500.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Arctiva": apparel_1500.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Arctiva": apparel_1500.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Arctiva": apparel_1500.iloc[0][4]})
                    else:
                       print('No Arctiva in the report')
                    if len(apparel_1500z) > 0:
                        need_apparel_z1r = int(apparel_1500z.iloc[0][2] - apparel_1500z.iloc[0][4] + apparel_1500z['Need'])
                    ##Z1R
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Z1R": apparel_1500z.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Z1R": apparel_1500z.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Z1R": apparel_1500z.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Z1R": apparel_1500z.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Z1R": apparel_1500z.iloc[0][4]})
                    else:
                        print('No Z1R in the report')
                    if len(apparel_1500t) > 0:
                        need_apparel_thor = int(apparel_1500t.iloc[0][2] - apparel_1500t.iloc[0][4] + apparel_1500t['Need'])
                    ##Thor
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Thor": apparel_1500t.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Thor": apparel_1500t.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Thor": apparel_1500t.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Thor": apparel_1500t.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Thor": apparel_1500t.iloc[0][4]})
                    
                    else:
                        print('No Thor in the report')
                    if len(apparel_1500m) > 0:
                        need_apparel_mooose = int(apparel_1500m.iloc[0][2] - apparel_1500m.iloc[0][4] + apparel_1500m['Need'])
                    ##Moose
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Moose": apparel_1500m.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Moose": apparel_1500m.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Moose": apparel_1500m.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Moose": apparel_1500m.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Moose": apparel_1500m.iloc[0][4]})
                    else:
                        print('No Moose in the report')
                    if len(apparel_Slippery) > 0:
                        need_apparel_slippery = int(apparel_Slippery.iloc[0][2] - apparel_Slippery.iloc[0][4] + apparel_Slippery['Need'])
                    ##Slippery
                        writer.update_page_form_field_values(
                            writer.pages[1], {"LY Slippery": apparel_Slippery.iloc[0][1]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"TY Slippery": apparel_Slippery.iloc[0][2]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Current Slippery": apparel_Slippery.iloc[0][3]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Need Slippery": apparel_Slippery.iloc[0][5]})
                        writer.update_page_form_field_values(
                            writer.pages[1], {"Roll-Off Slippery": apparel_Slippery.iloc[0][4]})
                    else:
                        print('No Slippery in the report')
                    #print(oil0)  
                     #q2 = roll off
                     #need = TY-q2 + next tier floor
                    
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Dealer Number": (res.group())})
                    
                    writer.update_page_form_field_values(
                        writer.pages[0], {"LY Tires": tire0.iloc[0][1]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"TY Tires": tire0.iloc[0][2]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Current Tires": tire0.iloc[0][3]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Need Tires": tire0.iloc[0][5]})
                    if tire0.iloc[0][2] <= 2000:
                        writer.update_page_form_field_values(writer.pages[0], {"TireL1": "/Yes"})
                    elif tire0.iloc[0][2] <= 4000 and tire0.iloc[0][2]>= 2001:
                        writer.update_page_form_field_values(writer.pages[0], {"TireL2": "/Yes"})
                    elif tire0.iloc[0][2] <= 6000 and tire0.iloc[0][2]>= 4001:
                        writer.update_page_form_field_values(writer.pages[0], {"TireL3": "/Yes"})
                    elif tire0.iloc[0][2]>= 6001:
                        writer.update_page_form_field_values(writer.pages[0], {"TireL4": "/Yes"})
                    else: 
                        print('error: tires checkbox')
                    
                    
                    writer.update_page_form_field_values(
                        writer.pages[0], {"LY Battery": battery0.iloc[0][1]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"TY Battery": battery0.iloc[0][2]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Current Battery": battery0.iloc[0][3]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Need Battery": battery0.iloc[0][5]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Roll-Off Battery": battery0.iloc[0][4]})
                    if battery0.iloc[0][2] <= 1499 and battery0.iloc[0][2] >= 500:
                        writer.update_page_form_field_values(writer.pages[0], {"Btry1": "/Yes"})
                    elif battery0.iloc[0][2] <= 2499 and battery0.iloc[0][2]>= 1500:
                        writer.update_page_form_field_values(writer.pages[0], {"Btry2": "/Yes"})
                    elif battery0.iloc[0][2] <= 3499 and battery0.iloc[0][2]>= 2500:
                        writer.update_page_form_field_values(writer.pages[0], {"Btry3": "/Yes"})
                    elif battery0.iloc[0][2]<= 4499 and battery0.iloc[0][2]>= 3500:
                        writer.update_page_form_field_values(writer.pages[0], {"Btry4": "/Yes"})
                    elif battery0.iloc[0][2]>= 4500:
                        writer.update_page_form_field_values(writer.pages[0], {"Btry5": "/Yes"})
                    else: 
                        print('error: battery checkbox')
                    
                    writer.update_page_form_field_values(
                        writer.pages[0], {"LY Oil/Chem": oil0.iloc[0][1]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"TY Oil/Chem": oil0.iloc[0][2]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Current Oil/Chem": oil0.iloc[0][3]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Need Oil/Chem": oil0.iloc[0][5]})
                    writer.update_page_form_field_values(
                        writer.pages[0], {"Roll-Off Oil/Chem": oil0.iloc[0][4]})
                    if oil0.iloc[0][2] <= 1499 and oil0.iloc[0][2] >= 500:
                        writer.update_page_form_field_values(writer.pages[0], {"ChemL1": "/Yes"})
                    elif oil0.iloc[0][2] <= 2499 and oil0.iloc[0][2]>= 1500:
                        writer.update_page_form_field_values(writer.pages[0], {"ChemL2": "/Yes"})
                    elif oil0.iloc[0][2] <= 3499 and oil0.iloc[0][2]>= 2500:
                        writer.update_page_form_field_values(writer.pages[0], {"ChemL3": "/Yes"})
                    elif oil0.iloc[0][2]<= 4499 and oil0.iloc[0][2]>= 3500:
                        writer.update_page_form_field_values(writer.pages[0], {"ChemL4": "/Yes"})
                    elif oil0.iloc[0][2]>= 4500:
                        writer.update_page_form_field_values(writer.pages[0], {"ChemL5": "/Yes"})
                    else: 
                        print('error: oil_chem checkbox')
                    
                    writer.update_page_form_field_values(
                        writer.pages[1], {"LY Brakes": brake0.iloc[0][1]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"TY Brakes": brake0.iloc[0][2]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Current Brakes": brake0.iloc[0][3]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Need Brakes": brake0.iloc[0][5]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Roll-Off Brakes": brake0.iloc[0][4]})
                    if brake0.iloc[0][2] <= 1499 and brake0.iloc[0][2] >= 500:
                        writer.update_page_form_field_values(writer.pages[1], {"BrakeL1": "/Yes"})
                    elif brake0.iloc[0][2] <= 2499 and brake0.iloc[0][2]>= 1500:
                        writer.update_page_form_field_values(writer.pages[1], {"BrakeL2": "/Yes"})
                    elif brake0.iloc[0][2] <= 3499 and brake0.iloc[0][2]>= 2500:
                        writer.update_page_form_field_values(writer.pages[1], {"BrakeL3": "/Yes"})
                    elif brake0.iloc[0][2]>= 3500:
                        writer.update_page_form_field_values(writer.pages[1], {"BrakeL4": "/Yes"})
                    else: 
                        print('error: brake checkbox')
                    
                    #FIND SECOND PAGE DETAILS
                    writer.update_page_form_field_values(
                        writer.pages[1], {"LY Helmet": helmet0.iloc[0][1]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"TY Helmet": helmet0.iloc[0][2]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Current Helmet": helmet0.iloc[0][3]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Need Helmet": helmet0.iloc[0][5]})
                    writer.update_page_form_field_values(
                        writer.pages[1], {"Roll-Off Helmet": helmet0.iloc[0][4]})
                    if helmet0.iloc[0][2] <= 1499 and helmet0.iloc[0][2] >= 500:
                        writer.update_page_form_field_values(writer.pages[1], {"HelmetL1": "/Yes"})
                    elif helmet0.iloc[0][2] <= 2499 and helmet0.iloc[0][2]>= 1500:
                        writer.update_page_form_field_values(writer.pages[1], {"HelmetL2": "/Yes"})
                    elif helmet0.iloc[0][2] <= 3499 and helmet0.iloc[0][2]>= 2500:
                        writer.update_page_form_field_values(writer.pages[1], {"HelmetL3": "/Yes"})
                    elif helmet0.iloc[0][2]<= 4499 and helmet0.iloc[0][2]>= 3500:
                        writer.update_page_form_field_values(writer.pages[1], {"HelmetL4": "/Yes"})
                    elif helmet0.iloc[0][2]>= 4500:
                        writer.update_page_form_field_values(writer.pages[1], {"HelmetL5": "/Yes"})
                    else: 
                        print('error: helmet checkbox')
                     
                    with Path(f + "\\CumulativeEditable_filled.pdf", mode ="wb") as output_stream:
                        writer.write(output_stream)
                else: 
                    print('not a cumulative report')
            else:
                print('did not find files for folder')
        else:
           print('did not find folder')
         

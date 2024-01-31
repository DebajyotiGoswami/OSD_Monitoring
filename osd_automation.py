from pandas import DataFrame , Series
import pandas as pd
import numpy as np
import openpyxl as xl
from datetime import date

ccc_list = [3332103 , 3332201 , 3332202 , 3332208 , 3332300 , 3332301 , 3332401 , 3332402]
columns = ['CCC Code', 'MRU', 'Con Id', 'Name', 'Address','Base Class', 'Device', 'Period', 'OSD', 'OSD Lakh', 'MOBILE NO']


data = pd.read_excel('kalyani_total_osd.xlsx')

live_non_govt_dom_condition = (data['Base Class'] == 'D') & (data['Discon Status'].isna()) & (data['Govt Status'].isna())
live_non_govt_comm_condition = (data['Base Class'] == 'C') & (data['Discon Status'].isna()) & (data['Govt Status'].isna())
live_non_govt_ind_condition = (data['Base Class'] == 'I') & (data['Discon Status'].isna()) & (data['Govt Status'].isna())
live_non_govt_agri_condition = (data['Base Class'] == 'A') & (data['Discon Status'].isna()) & (data['Govt Status'].isna())
live_non_govt_oth_condition = (~data['Base Class'].isin(['D' , 'C' , 'I' , 'A'])) & (data['Discon Status'].isna()) & (data['Govt Status'].isna())

dom_osd = data[live_non_govt_dom_condition]
comm_osd = data[live_non_govt_comm_condition]
ind_osd = data[live_non_govt_ind_condition]
agri_osd = data[live_non_govt_agri_condition]
oth_osd = data[live_non_govt_oth_condition]

dom_osd_10000 = dom_osd[dom_osd['OSD'] > 10000]
dom_osd_5000 = dom_osd[dom_osd['OSD'] > 5000]
dom_osd_3000 = dom_osd[dom_osd['OSD'] > 3000]
dom_osd_1000 = dom_osd[dom_osd['OSD'] > 1000]
dom_osd_500 = dom_osd[dom_osd['OSD'] > 500]
dom_osd_5000_to_10000 = dom_osd[dom_osd['OSD'].between(5000 , 9999)]
dom_osd_3000_to_5000 = dom_osd[dom_osd['OSD'].between(3000 , 4999)]
dom_osd_1000_to_3000 = dom_osd[dom_osd['OSD'].between(1000 , 3000)]
dom_osd_500_to_1000 = dom_osd[dom_osd['OSD'].between(500 , 1000)]

comm_osd_500 = comm_osd[comm_osd['OSD'] > 500]

ind_osd_summary = ind_osd.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
ind_osd_summary = pd.concat([ind_osd_summary , ind_osd_summary.sum().to_frame().T])
ind_osd_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

dom_osd_summary = dom_osd.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
dom_osd_summary = pd.concat([dom_osd_summary , dom_osd_summary.sum().to_frame().T])
dom_osd_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

dom_osd_10000_summary = dom_osd_10000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
dom_osd_10000_summary = pd.concat([dom_osd_10000_summary , dom_osd_10000_summary.sum().to_frame().T])
dom_osd_10000_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

dom_osd_5000_to_10000_summary = dom_osd_5000_to_10000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
dom_osd_5000_to_10000_summary = pd.concat([dom_osd_5000_to_10000_summary , dom_osd_5000_to_10000_summary.sum().to_frame().T])
dom_osd_5000_to_10000_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

dom_osd_3000_to_5000_summary = dom_osd_3000_to_5000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
dom_osd_3000_to_5000_summary = pd.concat([dom_osd_3000_to_5000_summary , dom_osd_3000_to_5000_summary.sum().to_frame().T])
dom_osd_3000_to_5000_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

comm_osd_summary = comm_osd.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
comm_osd_summary = pd.concat([comm_osd_summary , comm_osd_summary.sum().to_frame().T])
comm_osd_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

comm_osd_500_summary = comm_osd_500.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
comm_osd_500_summary = pd.concat([comm_osd_500_summary , comm_osd_500_summary.sum().to_frame().T])
comm_osd_500_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

agri_osd_summary = agri_osd.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
agri_osd_summary = pd.concat([agri_osd_summary , agri_osd_summary.sum().to_frame().T])
agri_osd_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

oth_osd_summary = oth_osd.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
oth_osd_summary = pd.concat([oth_osd_summary , oth_osd_summary.sum().to_frame().T])
oth_osd_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

def top_100(non_govt_live_df):
    top_100_df = []
    for each_ccc in ccc_list:
        sorted_temp = non_govt_live_df[ (non_govt_live_df['CCC Code'] == int(each_ccc) )]
        sorted_temp = sorted_temp.sort_values(by = 'OSD' , ascending = False)[:100]
        top_100_df.append(sorted_temp)

    top_100_df = pd.concat(top_100_df)
    return top_100_df

top_100_dom = top_100(dom_osd)
top_100_dom_summary = top_100_dom.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
top_100_dom_summary = pd.concat([top_100_dom_summary , top_100_dom_summary.sum().to_frame().T])
top_100_dom_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

top_100_comm = top_100(comm_osd)
top_100_comm_summary = top_100_comm.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0)
top_100_comm_summary = pd.concat([top_100_comm_summary , top_100_comm_summary.sum().to_frame().T])
top_100_comm_summary.rename(index = {0 : 'TOTAL'} , inplace = True)

with pd.ExcelWriter(str(date.today())+'-OSD-SUMMARY.xlsx' , engine="xlsxwriter") as writer:
 
    dom_osd_summary.to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 0)
    worksheet = writer.sheets['summary']
    worksheet.write(0 , 1 , "DOM-ALL-OSD")

    dom_osd_10000_summary.to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 4)
    worksheet.write(0 , 5 , "DOM-ABOVE-10000")

    dom_osd_5000_to_10000_summary.to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 8)
    worksheet.write(0 , 9 , "DOM-5000-TO-10000")

    dom_osd_3000_to_5000_summary.to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 12)
    worksheet.write(0 , 13 , "DOM-3000-TO-5000")

    top_100_dom_summary.to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 16)
    worksheet.write(0 , 17 , "DOM-TOP-100-OSD")

    comm_osd_summary.to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 0)
    worksheet.write(13 , 1 , "COMM-ALL-OSD")

    comm_osd_500_summary.to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 4)
    worksheet.write(13 , 5 , "COMM-ABOVE-500")

    top_100_comm_summary.to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 8)
    worksheet.write(13 , 9 , "COMM-TOP-100-OSD")

    ind_osd_summary.to_excel(writer , sheet_name = 'summary' , startrow = 28 , startcol = 0)
    worksheet.write(26 , 1 , "IND-ALL-OSD")


    dom_osd_10000[columns].to_excel(writer, sheet_name= 'dom_10K', startrow= 0 , index = False)
    dom_osd_5000_to_10000[columns].to_excel(writer, sheet_name= 'dom_5K_to_10K', startrow= 0 , index = False)
    dom_osd_3000_to_5000[columns].to_excel(writer, sheet_name= 'dom_3K_to_5K', startrow= 0 , index = False)
    top_100_dom[columns].to_excel(writer, sheet_name= 'dom_top_100', startrow= 0 , index = False)
    comm_osd_500[columns].to_excel(writer, sheet_name= 'comm_500', startrow= 0 , index = False)
    top_100_comm[columns].to_excel(writer, sheet_name= 'comm_top_100', startrow= 0 , index = False)
    ind_osd[columns].to_excel(writer, sheet_name= 'ind_osd_all', startrow= 0 , index = False)

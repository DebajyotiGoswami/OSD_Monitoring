import pandas as pd
from pandas import Series , DataFrame
import numpy as np
import openpyxl as xl
from datetime import date

df = pd.read_excel('kalyani_total_osd.xlsx')
pd.set_option('display.max_columns' , 100)

ccc_list = [3332103 , 3332201 , 3332202 , 3332208 , 3332300 , 3332301 , 3332401 , 3332402]
columns = ['CCC Code', 'MRU', 'Con Id', 'Name', 'Address','Base Class', 'Device', 'Period', 'OSD', 'OSD Lakh', 'MOBILE NO']
base_data = pd.read_excel('kalyani_total_osd.xlsx')

#non_govt_live data
non_govt_live = base_data[(base_data['Govt Status'].isnull() & base_data['Discon Status'].isnull())]

#following queries are to find out non government and live OSD data from Dom , Comm , Ind and Agri
non_govt_live_comm = non_govt_live[(non_govt_live['Base Class'] == 'C')]
non_govt_live_dom = non_govt_live[(non_govt_live['Base Class'] == 'D')]
non_govt_live_ind = non_govt_live[(non_govt_live['Base Class'] == 'I')]
non_govt_live_agri = non_govt_live[(non_govt_live['Base Class'] == 'A')]
non_govt_live_oth = non_govt_live[(~non_govt_live['Base Class'].isin(['D' , 'A' , 'C' , 'I']))]

#following queries are to find OSD based data on non government and live OSD
#we can change queries as required by us like OSD > 500 or OSD > 5000
non_govt_live_comm_5000 = non_govt_live_comm[non_govt_live_comm['OSD'] > 5000]
non_govt_live_comm_500 = non_govt_live_comm[non_govt_live_comm['OSD'] > 500]

non_govt_live_dom_5000 = non_govt_live_dom[non_govt_live_dom['OSD'] > 5000]
non_govt_live_dom_3000 = non_govt_live_dom[non_govt_live_dom['OSD'] > 3000]
non_govt_live_dom_1000 = non_govt_live_dom[non_govt_live_dom['OSD'] > 1000]

non_govt_live_comm_500_6_months = non_govt_live_comm_500[non_govt_live_comm_500['DAYS DIFFERENCE'].isin(['1 Year' , '6 Months'])]
non_govt_live_dom_5000_6_months = non_govt_live_dom_5000[non_govt_live_dom_5000['DAYS DIFFERENCE'].isin(['1 Year' , '6 Months'])]

# temp = non_govt_live_comm_500.groupby('CCC Code')#.agg[{'Con Id' : 'count'}]
# temp_summary = temp.agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'})

#top 100 list
def top_100(non_govt_live_df):
    top_100_df = []
    for each_ccc in ccc_list:
        sorted_temp = non_govt_live_df[ (non_govt_live_df['CCC Code'] == int(each_ccc) )]
        sorted_temp = sorted_temp.sort_values(by = 'OSD' , ascending = False)[:100]
        top_100_df.append(sorted_temp)


    top_100_df = pd.concat(top_100_df)
    return top_100_df

top_100_dom = top_100(non_govt_live_dom)
top_100_comm = top_100(non_govt_live_comm)
# top_100_dom.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'})

#so the following variable we have created to get different reports :
### (1) non_govt_live_comm_5000 - checked ok
### (2) non_govt_live_comm_500 - checked ok
### (3) non_govt_live_dom_5000 - checked ok
### (4) non_govt_live_dom_1000 - checked ok
### (5) non_govt_live_comm_500_6_months - checked ok
### (6) non_govt_live_dom_5000_6_months - checked ok
### (7) top_100_dom - ********* not checked yet ************
### (8) top_100_comm - *********** not checked yet **********

# non_govt_live_dom_5000_6_months.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'})

#### writting our output to an excel file
with pd.ExcelWriter(str(date.today())+'-total-OSD.xlsx' , engine="xlsxwriter") as writer:
    non_govt_live_comm_5000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 0)
    worksheet = writer.sheets['summary']
    
    worksheet.write(0 , 0 , "COMM-5000")
    non_govt_live_comm_500.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 4)
    worksheet.write(0 , 4 , "COMM-500")
    non_govt_live_dom_5000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 8)
    worksheet.write(0 , 8 , "DOM-5000")
    non_govt_live_dom_1000.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 2 , startcol = 12)
    worksheet.write(0 , 12 , "DOM-1000")
    non_govt_live_comm_500_6_months.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 0)
    worksheet.write(13 , 0 , "COMM-500-6-MONTHS")
    non_govt_live_dom_5000_6_months.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 4)
    worksheet.write(13 , 4 , "DOM-5000-6-MONTHS")
    non_govt_live_ind.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 15 , startcol = 8)
    worksheet.write(13 , 8 , "IND-ALL")
    top_100_dom.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 28 , startcol = 0)
    worksheet.write(26 , 0 , "DOM-TOP-100")
    top_100_comm.groupby('CCC Code').agg({'Con Id' : 'count' , 'OSD Lakh' : 'sum'}).reindex(ccc_list , fill_value = 0).to_excel(writer , sheet_name = 'summary' , startrow = 28 , startcol = 4)
    worksheet.write(26 , 4 , "COMM-TOP-100")

    non_govt_live_comm_5000[columns].to_excel(writer, sheet_name= 'comm_5000', startrow= 0 , index = False)
    non_govt_live_comm_500[columns].to_excel(writer, sheet_name= 'comm_500', startrow= 0 , index = False)
    non_govt_live_dom_5000[columns].to_excel(writer, sheet_name= 'dom_5000', startrow= 0 , index = False)
    non_govt_live_dom_1000[columns].to_excel(writer, sheet_name= 'dom_1000', startrow= 0 , index = False)
    non_govt_live_ind[columns].to_excel(writer, sheet_name= 'ind_all', startrow= 0 , index = False)
    non_govt_live_comm_500_6_months[columns].to_excel(writer, sheet_name= 'comm_500_6_months', startrow= 0 , index = False)
    non_govt_live_dom_5000_6_months[columns].to_excel(writer, sheet_name= 'dom_5000_6_months', startrow= 0 , index = False)
    top_100_dom[columns].to_excel(writer, sheet_name= 'top_100_dom', startrow= 0 , index = False)
    top_100_comm[columns].to_excel(writer, sheet_name= 'top_100_comm', startrow= 0 , index = False)

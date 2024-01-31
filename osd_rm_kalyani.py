#! python3
import pyperclip, openpyxl, os, time, datetime

def find_cell(sheet, heading):
    '''
    get the row and column number of the cell which contains a particular heading
    argument --- a sheet type object of a excel file
    return --- two integer: row and column index
    '''
    max_row= sheet.max_row
    max_column= sheet.max_column

    for i in range(1,max_row+1):
        for j in range(1, max_column+1):
            try:
                if sheet.cell(row= i, column= j).value and sheet.cell(row= i, column= j).value.strip()== heading:
                    return i,j
            except:
                pass

def clear_file(sheet):
    max_row= sheet.max_row
    for rowNum in range(2, max_row+1):
        sheet.cell(row= rowNum, column= 1).value= None
        sheet.cell(row= rowNum, column= 2).value= None
        sheet.cell(row= rowNum, column= 3).value= None
    return


def calculation(filename):
    fileObj= openpyxl.load_workbook(filename)
    sheet= fileObj.active
    conid_row, conid_col= find_cell(sheet, "Consumer Id")
    class_row, class_col= find_cell(sheet, "Base Class")
    dept_row, dept_col= find_cell(sheet, "Nature of Conn")
    govt_row, govt_col= find_cell(sheet, "Gov/Non-Gov")
    osd_row, osd_col= find_cell(sheet, "D2 Net O/S")
    dis_row, dis_col= find_cell(sheet, "Discon Status")

    max_row, max_col= sheet.max_row, sheet.max_column


    conListObj= openpyxl.load_workbook('list_con.xlsx')
    conListSheet= conListObj[filename[:3]]
    clear_file(conListSheet)

    index_10000, index_5000, index_1000, index_500, index_less= 1, 1, 1, 1, 1

    count_10000, osd_10000, count_5000, osd_5000, count_1000, osd_1000, count_500, osd_500, count_less, osd_less, count_ind, osd_ind, count_agri, osd_agri= 0, 0, 0,0,0,0,0,0,0,0,0,0,0,0

    '''
    govt_class= ['AGRI IRRI.','JUDICIAL', 'GRAM PANCHAYET', 'Municipality', 'Nature of Conn.', 'PHE','IRRIG. AND WATER WAYS',
                'POLICE', 'POST AND TELEGRAPH', 'PWD(ELECTRICAL)', 'STATE GOVT.','AGRI MECH','MINOR IRRIG.',
                 'DEPT. OF HEALTH','PRI HEALTH CENTRE','MDTW','ZILLA PARISAD','PRIMARY SCHOOL','B.D.O. OFFICE',
                 'PWD(CONST. BOARD)','DEPT. OF EDU.','C.A.D.A.','JUDICIAL','PWD','PWD(NH DIVISION)','HIGH SCHOOL',
                 'PWD(ROADS)','DEPT. OF AGRI.']
    '''
    govt_class = []
    
    for i in range(1, max_row+1):
        if sheet.cell(row= i, column= class_col).value not in ('I','A'):        #('W', 'H', 'A'):
            if sheet.cell(row= i, column= dis_col).value is None and sheet.cell(row= i, column= govt_col).value in ('', None) and sheet.cell(row= i, column= dept_col).value not in govt_class:
                if sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value> 10000:
                    conId= str(sheet.cell(row= i, column= conid_col).value)
                    conListSheet.cell(row= index_10000, column= 1).value= conId
                    conListSheet.cell(row= index_10000+ 1, column= 1).value= '0'+conId
                    index_10000+= 2
                    count_10000+= 1
                    osd_10000+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value> 5000:
                    conId= str(sheet.cell(row= i, column= conid_col).value)
                    conListSheet.cell(row= index_5000, column= 1).value= conId
                    conListSheet.cell(row= index_5000+ 1, column= 1).value= '0'+conId
                    index_5000+= 2
                    count_5000+= 1
                    osd_5000+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value> 1000:
                    conId= str(sheet.cell(row= i, column= conid_col).value)
                    conListSheet.cell(row= index_1000, column= 2).value= conId
                    conListSheet.cell(row= index_1000+ 1, column= 2).value= '0'+conId
                    index_1000+= 2
                    count_1000+= 1
                    osd_1000+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value> 500:
                    conId= str(sheet.cell(row= i, column= conid_col).value)
                    conListSheet.cell(row= index_500, column= 2).value= conId
                    conListSheet.cell(row= index_500+ 1, column= 2).value= '0'+conId
                    index_500+= 2
                    count_500+= 1
                    osd_500+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value>=200:
                    conId= str(sheet.cell(row= i, column= conid_col).value)
                    conListSheet.cell(row= index_less, column= 3).value= conId
                    conListSheet.cell(row= index_less+ 1, column= 3).value= '0'+ conId
                    index_less+= 2
                    count_less+= 1
                    osd_less+= sheet.cell(row= i, column= osd_col).value/100000
        elif sheet.cell(row= i, column= class_col).value== 'I' and sheet.cell(row= i, column= dis_col).value is None and sheet.cell(row= i, column= govt_col).value in ('', None) and sheet.cell(row= i, column= dept_col).value not in govt_class and sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value > 200:
                count_ind+= 1
                osd_ind+= sheet.cell(row= i, column= osd_col).value/100000
        else:
            if sheet.cell(row= i, column= class_col).value== 'A' and sheet.cell(row= i, column= dis_col).value is None and sheet.cell(row= i, column= govt_col).value in ('', None) and sheet.cell(row= i, column= dept_col).value not in govt_class and sheet.cell(row= i, column= osd_col).value and sheet.cell(row= i, column= osd_col).value > 200:
                count_agri+= 1
                osd_agri+= sheet.cell(row= i, column= osd_col).value/100000

    conListObj.save('list_con.xlsx')
    print("Work completed for filename :", filename)
    return count_less, osd_less, count_500, osd_500, count_1000, osd_1000, count_5000, osd_5000, count_10000, osd_10000,  count_ind, osd_ind, count_agri, osd_agri      

def get_new_sheet(filename):
    '''
    create a name with current date and set it to the current sheet name
    argument ---  a sheet which needs to be modified.
    return --- New sheet name
    '''
    fileObj= openpyxl.load_workbook(filename)       #opening the file
    date= '.'.join(str(datetime.date.today()- datetime.timedelta(1)).split('-')[::-1])     #get yesterday's date
    fileObj.active.title= date                            #rename the current sheet   
    fileObj.save(filename)
    return date

if __name__== '__main__':
    start= time.time()
    print("preparing OSD report for RM.....")
    input("If you are ready to go, then press enter")
    #path= os.getcwd()
    #os.chdir(path+ '\..\Excels')
    input_file= 'OSD_rm_kalyani.xlsx'
    new_sheet_name= get_new_sheet(input_file)
    fileObj= openpyxl.load_workbook(input_file)
    sheet= fileObj[new_sheet_name]

    sheet['C9'], sheet['D9'], sheet['E9'], sheet['F9'], sheet['G9'], sheet['H9'], sheet['I9'], sheet['J9'], sheet['K9'], sheet['L9'], sheet['Q9'], sheet['R9'], sheet['U9'], sheet['V9']= calculation("103.xlsx")
    sheet['C10'], sheet['D10'], sheet['E10'], sheet['F10'], sheet['G10'], sheet['H10'], sheet['I10'], sheet['J10'], sheet['K10'], sheet['L10'], sheet['Q10'], sheet['R10'], sheet['U10'], sheet['V10']= calculation("201.xlsx")
    sheet['C11'], sheet['D11'], sheet['E11'], sheet['F11'], sheet['G11'], sheet['H11'], sheet['I11'], sheet['J11'], sheet['K11'], sheet['L11'], sheet['Q11'], sheet['R11'], sheet['U11'], sheet['V11']= calculation("202.xlsx")
    sheet['C12'], sheet['D12'], sheet['E12'], sheet['F12'], sheet['G12'], sheet['H12'], sheet['I12'], sheet['J12'], sheet['K12'], sheet['L12'], sheet['Q12'], sheet['R12'], sheet['U12'], sheet['V12']= calculation("208.xlsx")
    sheet['C13'], sheet['D13'], sheet['E13'], sheet['F13'], sheet['G13'], sheet['H13'], sheet['I13'], sheet['J13'], sheet['K13'], sheet['L13'], sheet['Q13'], sheet['R13'], sheet['U13'], sheet['V13']= calculation("300.xlsx")
    sheet['C14'], sheet['D14'], sheet['E14'], sheet['F14'], sheet['G14'], sheet['H14'], sheet['I14'], sheet['J14'], sheet['K14'], sheet['L14'], sheet['Q14'], sheet['R14'], sheet['U14'], sheet['V14']= calculation("301.xlsx")
    sheet['C15'], sheet['D15'], sheet['E15'], sheet['F15'], sheet['G15'], sheet['H15'], sheet['I15'], sheet['J15'], sheet['K15'], sheet['L15'], sheet['Q15'], sheet['R15'], sheet['U15'], sheet['V15']= calculation("401.xlsx")
    sheet['C16'], sheet['D16'], sheet['E16'], sheet['F16'], sheet['G16'], sheet['H16'], sheet['I16'], sheet['J16'], sheet['K16'], sheet['L16'], sheet['Q16'], sheet['R16'], sheet['U16'], sheet['V16']= calculation("402.xlsx")


    fileObj.save(input_file)
    #os.chdir(path)
    end= time.time()
    print("Finished in %.2f minutes. Please check your file" %float((end-start)/60))
    input("Press ENTER")


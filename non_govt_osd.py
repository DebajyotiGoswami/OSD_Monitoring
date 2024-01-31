#! python3
import pyperclip, openpyxl, os, time

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
            if str(sheet.cell(row= i, column= j).value and sheet.cell(row= i, column= j).value).strip()== heading:
                return i,j

def calculation(filename):
    fileObj= openpyxl.load_workbook(filename)
    sheet= fileObj.active
    conid_row, conid_col= find_cell(sheet, 'Consumer Id')
    class_row, class_col= find_cell(sheet, 'Base Class')
    dept_row, dept_col= find_cell(sheet, 'Nature of Conn')
    osd_row, osd_col= find_cell(sheet, 'D2 Net O/S')
    dis_row, dis_col= find_cell(sheet, "Discon Status")
    govt_row, govt_col= find_cell(sheet, "Gov/Non-Gov")

    max_row, max_col= sheet.max_row, sheet.max_column

    live_count= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}
    live_osd= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}
    temp_count= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}
    temp_osd= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}
    perm_count= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}
    perm_osd= {'agri': 0, 'com': 0, 'dom': 0, 'ind': 0, 'other': 0}

    '''
    govt_class= ['AGRI IRRI.','JUDICIAL', 'GRAM PANCHAYET', 'Municipality', 'Nature of Conn.', 'PHE','IRRIG. AND WATER WAYS',
                'POLICE', 'POST AND TELEGRAPH', 'PWD(ELECTRICAL)', 'STATE GOVT.','AGRI MECH','MINOR IRRIG.',
                 'DEPT. OF HEALTH','PRI HEALTH CENTRE','MDTW','ZILLA PARISAD','PRIMARY SCHOOL','B.D.O. OFFICE',
                 'PWD(CONST. BOARD)','DEPT. OF EDU.','C.A.D.A.','JUDICIAL','PWD','PWD(NH DIVISION)','HIGH SCHOOL',
                 'PWD(ROADS)','DEPT. OF AGRI.']
    '''
    govt_class = []
    for i in range(1, max_row+1):
        if sheet.cell(row= i, column= dept_col).value not in govt_class and sheet.cell(row= i, column= govt_col).value in ('', None) and sheet.cell(row= i, column= osd_col).value is not None and sheet.cell(row= i, column= osd_col).value>200:
            if sheet.cell(row= i, column= dis_col).value is None:
                if sheet.cell(row= i, column= class_col).value== 'A':
                    live_count['agri']+= 1
                    live_osd['agri']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'C':
                    live_count['com']+= 1
                    live_osd['com']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'D':
                    live_count['dom']+= 1
                    live_osd['dom']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'I':
                    live_count['ind']+= 1
                    live_osd['ind']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value not in ('A' , 'C' , 'D' , 'I'):
                    print(sheet.cell(row = i , column = conid_col).value)
                    live_count['other']+= 1
                    live_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
                elif sheet.cell(row= i, column= class_col).value not in ('H','G'):
                    live_count['other']+= 1
                    live_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
            elif sheet.cell(row= i, column= dis_col).value== "Temprory Disconnected":
                if sheet.cell(row= i, column= class_col).value== 'A':
                    temp_count['agri']+= 1
                    temp_osd['agri']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'C':
                    temp_count['com']+= 1
                    temp_osd['com']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'D':
                    temp_count['dom']+= 1
                    temp_osd['dom']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'I':
                    temp_count['ind']+= 1
                    temp_osd['ind']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value not in ('A' , 'C' , 'D' , 'I'):
                    temp_count['other']+= 1
                    temp_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
                elif sheet.cell(row= i, column= class_col).value not in ('H','G'):
                    live_count['other']+= 1
                    live_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
            elif sheet.cell(row= i, column= dis_col).value in ("Deemed Disconnection"):
                if sheet.cell(row= i, column= class_col).value== 'A':
                    perm_count['agri']+= 1
                    perm_osd['agri']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'C':
                    perm_count['com']+= 1
                    perm_osd['com']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'D':
                    perm_count['dom']+= 1
                    perm_osd['dom']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value== 'I':
                    perm_count['ind']+= 1
                    perm_osd['ind']+= sheet.cell(row= i, column= osd_col).value/100000
                elif sheet.cell(row= i, column= class_col).value not in ('A' , 'C' , 'D' , 'I'):
                    perm_count['other']+= 1
                    perm_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
                elif sheet.cell(row= i, column= class_col).value not in ('H','G'):
                    live_count['other']+= 1
                    live_osd['other']+= sheet.cell(row= i, column= osd_col).value/100000
                '''
    
    print("Work completed for filename :", filename)
    return live_count['agri'], live_osd['agri'], live_count['com'], live_osd['com'], live_count['dom'], live_osd['dom'], live_count['ind'], live_osd['ind'], live_count['other'], live_osd['other'], temp_count['agri'], temp_osd['agri'], temp_count['com'], temp_osd['com'], temp_count['dom'], temp_osd['dom'], temp_count['ind'], temp_osd['ind'], temp_count['other'], temp_osd['other'], perm_count['agri'], perm_osd['agri'], perm_count['com'], perm_osd['com'], perm_count['dom'], perm_osd['dom'], perm_count['ind'], perm_osd['ind'], perm_count['other'], perm_osd['other']
        
if __name__== '__main__':
    start= time.time()
    #path= os.getcwd()
    #os.chdir(path+ '\..\Excels')
    input("If you are good to go. Press enter")

    input_file= "NON_GOVT_OSD.xlsx"
    fileObj= openpyxl.load_workbook(input_file)
    live_sheet= fileObj['LIVE']
    discon_sheet= fileObj['DISCON']

    live_sheet['E5'], live_sheet['E6'], live_sheet['F5'], live_sheet['F6'], live_sheet['G5'],live_sheet['G6'], live_sheet['H5'], live_sheet['H6'], live_sheet['I5'], live_sheet['I6'],discon_sheet['E5'], discon_sheet['E6'], discon_sheet['F5'], discon_sheet['F6'], discon_sheet['G5'],discon_sheet['G6'], discon_sheet['H5'], discon_sheet['H6'], discon_sheet['I5'], discon_sheet['I6'],discon_sheet['E27'], discon_sheet['E28'], discon_sheet['F27'], discon_sheet['F28'], discon_sheet['G27'],discon_sheet['G28'], discon_sheet['H27'], discon_sheet['H28'], discon_sheet['I27'], discon_sheet['I28']= calculation("103.xlsx")
    live_sheet['E7'], live_sheet['E8'], live_sheet['F7'], live_sheet['F8'], live_sheet['G7'],live_sheet['G8'], live_sheet['H7'], live_sheet['H8'], live_sheet['I7'], live_sheet['I8'],discon_sheet['E7'], discon_sheet['E8'], discon_sheet['F7'], discon_sheet['F8'], discon_sheet['G7'],discon_sheet['G8'], discon_sheet['H7'], discon_sheet['H8'], discon_sheet['I7'], discon_sheet['I8'],discon_sheet['E29'], discon_sheet['E30'], discon_sheet['F29'], discon_sheet['F30'], discon_sheet['G29'],discon_sheet['G30'], discon_sheet['H29'], discon_sheet['H30'], discon_sheet['I29'], discon_sheet['I30']= calculation("201.xlsx")
    live_sheet['E9'], live_sheet['E10'], live_sheet['F9'], live_sheet['F10'], live_sheet['G9'],live_sheet['G10'], live_sheet['H9'], live_sheet['H10'], live_sheet['I9'], live_sheet['I10'],discon_sheet['E9'], discon_sheet['E10'], discon_sheet['F9'], discon_sheet['F10'], discon_sheet['G9'],discon_sheet['G10'], discon_sheet['H9'], discon_sheet['H10'], discon_sheet['I9'], discon_sheet['I10'],discon_sheet['E31'], discon_sheet['E32'], discon_sheet['F31'], discon_sheet['F32'], discon_sheet['G31'],discon_sheet['G32'], discon_sheet['H31'], discon_sheet['H32'], discon_sheet['I31'], discon_sheet['I32']= calculation("202.xlsx")
    live_sheet['E11'], live_sheet['E12'], live_sheet['F11'], live_sheet['F12'], live_sheet['G11'],live_sheet['G12'], live_sheet['H11'], live_sheet['H12'], live_sheet['I11'], live_sheet['I12'],discon_sheet['E11'], discon_sheet['E12'], discon_sheet['F11'], discon_sheet['F12'], discon_sheet['G11'],discon_sheet['G12'], discon_sheet['H11'], discon_sheet['H12'], discon_sheet['I11'], discon_sheet['I12'],discon_sheet['E33'], discon_sheet['E34'], discon_sheet['F33'], discon_sheet['F34'], discon_sheet['G33'],discon_sheet['G34'], discon_sheet['H33'], discon_sheet['H34'], discon_sheet['I33'], discon_sheet['I34']= calculation("208.xlsx")
    live_sheet['E13'], live_sheet['E14'], live_sheet['F13'], live_sheet['F14'], live_sheet['G13'],live_sheet['G14'], live_sheet['H13'], live_sheet['H14'], live_sheet['I13'], live_sheet['I14'],discon_sheet['E13'], discon_sheet['E14'], discon_sheet['F13'], discon_sheet['F14'], discon_sheet['G13'],discon_sheet['G14'], discon_sheet['H13'], discon_sheet['H14'], discon_sheet['I13'], discon_sheet['I14'],discon_sheet['E35'], discon_sheet['E36'], discon_sheet['F35'], discon_sheet['F36'], discon_sheet['G35'],discon_sheet['G36'], discon_sheet['H35'], discon_sheet['H36'], discon_sheet['I35'], discon_sheet['I36']= calculation("300.xlsx")

    live_sheet['E15'], live_sheet['E16'], live_sheet['F15'], live_sheet['F16'], live_sheet['G15'],live_sheet['G16'], live_sheet['H15'], live_sheet['H16'], live_sheet['I15'], live_sheet['I16'],discon_sheet['E15'], discon_sheet['E16'], discon_sheet['F15'], discon_sheet['F16'], discon_sheet['G15'],discon_sheet['G16'], discon_sheet['H15'], discon_sheet['H16'], discon_sheet['I15'], discon_sheet['I16'],discon_sheet['E37'], discon_sheet['E38'], discon_sheet['F37'], discon_sheet['F38'], discon_sheet['G37'],discon_sheet['G38'], discon_sheet['H37'], discon_sheet['H38'], discon_sheet['I37'], discon_sheet['I38']= calculation("301.xlsx")
    live_sheet['E17'], live_sheet['E18'], live_sheet['F17'], live_sheet['F18'], live_sheet['G17'],live_sheet['G18'], live_sheet['H17'], live_sheet['H18'], live_sheet['I17'], live_sheet['I18'],discon_sheet['E17'], discon_sheet['E18'], discon_sheet['F17'], discon_sheet['F18'], discon_sheet['G17'],discon_sheet['G18'], discon_sheet['H17'], discon_sheet['H18'], discon_sheet['I17'], discon_sheet['I18'],discon_sheet['E39'], discon_sheet['E40'], discon_sheet['F39'], discon_sheet['F40'], discon_sheet['G39'],discon_sheet['G40'], discon_sheet['H39'], discon_sheet['H40'], discon_sheet['I39'], discon_sheet['I40']= calculation("401.xlsx")
    live_sheet['E19'], live_sheet['E20'], live_sheet['F19'], live_sheet['F20'], live_sheet['G19'],live_sheet['G20'], live_sheet['H19'], live_sheet['H20'], live_sheet['I19'], live_sheet['I20'],discon_sheet['E19'], discon_sheet['E20'], discon_sheet['F19'], discon_sheet['F20'], discon_sheet['G19'],discon_sheet['G20'], discon_sheet['H19'], discon_sheet['H20'], discon_sheet['I19'], discon_sheet['I20'],discon_sheet['E41'], discon_sheet['E42'], discon_sheet['F41'], discon_sheet['F42'], discon_sheet['G41'],discon_sheet['G42'], discon_sheet['H41'], discon_sheet['H42'], discon_sheet['I41'], discon_sheet['I42']= calculation("402.xlsx")

    fileObj.save(input_file)
    #os.chdir(path)
    end= time.time()
    print("Finished in %.2f minutes. Please check your file" %float((end-start)/60))
    input("Press ENTER")


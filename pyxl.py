import numpy as np
import pandas as pd

location = input("Enter the file location: ")  #file location = C:\Users\DELL\Desktop\responsibility\data\student_db.xlsx
data = pd.ExcelFile(r"{}".format(location))
sheet = pd.read_excel(r"{}".format(location),sheet_name=0)
sheet.head()
#converting everything to "object data type --> sheet = sheet.astype(str)"


#conditions

sheet1 = sheet
#converts data of all "object" columns to title case (eg. - Anubhav Kumar)
sheet1 = sheet1.applymap(lambda s:s.title() if type(s) == str else s)
sheet1.columns = map(str.title, sheet1.columns)#converts Titles of all columns to title case (eg. - Class Roll)
flag1 = 0
flag2 = 0
flag3 = 0
flag4 = 0

def export(final_sheet):
    split = location.split('\\')
    split.pop(len(split) - 1)
    raw_export = "\\".join(split)
    mandatory1 = '\\'
    mandatory2 = '.xlsx'
    pyxl_logo = '(pyXL)'
    print("")
    print("Exporting your sheet...")
    print("")
    print("Give a name to your sheet")
    export_file_name = input("File Name : ")
    final_export_location = raw_export + mandatory1 + export_file_name + pyxl_logo + mandatory2

    writer = pd.ExcelWriter(final_export_location , engine='xlsxwriter')
    final_sheet.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    print("")
    print("Your sheet '{}' has been exported to location - '{}'".format(export_file_name + pyxl_logo, raw_export))
    return final_sheet

def shortened(final_sheet):
    global flag4
    print("")
    print("Shorten sheet")
    print("")
    print("Which columns do you want in Your sheet ?")
    clmns = input("Enter column names seperated with COMMAS : ").split(',')
    errors = []
    #---------Handling for Column Name-------------
    for index,item in enumerate(clmns):
        if clmns[index] not in sheet1.columns:
            errors.append(item)
            
    if len(errors) > 0:
        flag4 = flag4 + 1
        print(flag4)
        if flag4 <= 3:
            print("")
            print("Error : Incorrect Column Names {}".format(errors))
            final_sheet = shortened(final_sheet)
        else:
            print("You seem to be SLEEPY..Go have a Coffee Break !!")
            print("Error : Incorrect Column Names {}".format(errors))
            print("")
            print("We are exporting your sheet")
            final_sheet = export(final_sheet)
    #--------------------------------------------------
    else:
        print("")
        print("Your Input : ",clmns)
        final_sheet = final_sheet.loc[:,clmns]
        print("")
        print("Great, Shortening Complete. '{} rows x {} columns' found".format(len(final_sheet), len(clmns)))
        final_sheet = export(final_sheet)
    return final_sheet
    
    
def further(final_sheet):
    global flag3
    global flag2
    print("1.) Run more Conditions    2.) Shorten Sheet     3.) Export Sheet ")
    choice = input("Enter Your Choice (1, 2 or 3) : ")
    #---------Handling Overloaded Conditions-------------
    if choice == '1':
        flag3 = flag3 + 1 
        print(flag3)
        if flag3 == 3:
            print("NO MORE CONDITIONS")
            print("Your Sheet is being exported...")
            export(final_sheet)
        else:
            print("")
            print("Run More Conditions")
            print("")
            final_sheet = initial()
            final_sheet = further(final_sheet)         
              
    elif choice == '2':
        final_sheet = shortened(final_sheet)    
    elif choice == '3':
        export(final_sheet)
    else:
    #---------Handling Invalid further() input-------------
        flag2 = flag2 + 1
        print(flag2)
        if flag2 == 3:
            print("You seem to be SLEEPY..Go have a Coffee Break !!")
            raise Exception("'{}' is not a valid input!!".format(choice))
        print("'{}' is not a valid input!!".format(choice))      
        final_sheet = further(final_sheet)
    return final_sheet


def operators(cc1,dtype):

    global flag1
    global sheet1
    operator = input("Operator: ")
    if (dtype == 'datetime64[ns]'):
        dob = input("Data in DD-MM-YYYY: ")
        intdata = pd.Timestamp(dob)
    else:
        intdata = float(input("Data: "))
    
    if operator == '>':
        edited_sheet = sheet1[sheet1["{}".format(cc1)] > intdata]#& (sheet["Gender"] == 'Male')]
        
    elif operator == '<':
        edited_sheet = sheet1[sheet1["{}".format(cc1)] < intdata]
        
    elif operator == '>=':
         edited_sheet = sheet1[sheet1["{}".format(cc1)] >= intdata]  #sheet[sheet['Date of Birth (DD-MM-YYYY)'] >= t]
        
    elif operator == '<=':
        edited_sheet = sheet1[sheet1["{}".format(cc1)] <= intdata]
        
    elif operator == '!=':
        edited_sheet = sheet1[sheet1["{}".format(cc1)] != intdata]
        
    elif operator == '=':
        edited_sheet = sheet1[sheet1["{}".format(cc1)] == intdata]            
    else:
    #-------------------------Handling Operators-------------------------    
        flag1 = flag1 + 1
        print(flag1)
        if flag1 == 3:
            print("You seem to be SLEEPY..Go have a Coffee Break !!")
            raise Exception("'{}' is not an Operator!".format(operator))
        print("Error: '{}' is not an Operator! Try using '>' or '<'".format(operator))
        edited_sheet = operators(cc1,dtype)     
        #raise Exception("'{}' is not an Operator! Try using '>' or '<'".format(operator))
        #raise SystemExit("'{}' is not an operator".format(operator))
    sheet1 = edited_sheet
    #edited_sheet = edited_sheet[['{}'.format(clmns[0]),'{}'.format(clmns[1])]]
    return edited_sheet
    
            
def char(cc1):
    global sheet1
    chardata = input("Data :")
    edited_sheet = sheet1[sheet1["{}".format(cc1)] == chardata]
    sheet1 = edited_sheet
    #edited_sheet = edited_sheet[['{}'.format(clmns[0]),'{}'.format(clmns[1])]]
    return edited_sheet    
    
    
def cond(cc1,dtype):
    if dtype == 'int64':
        mid_sheet = operators(cc1,dtype)
    elif dtype == 'float64':
        mid_sheet = operators(cc1,dtype)
    elif dtype == 'datetime64[ns]':
        mid_sheet = operators(cc1,dtype)
    else:
        mid_sheet = char(cc1)
    return mid_sheet
    
            
#cc = []
def initial():
    nconditions = int(input("No. of conditions u wanna have: "))
    if nconditions == 0:
        result_sheet = sheet1
        print("")
        print("No Conditions!!")
        print("")
    for i in range(nconditions):
        print("Condition {}:".format(i+1))
        cc1 = input("Column Name: ")
        if cc1 not in (sheet1.columns): 
            flag = 0
    #---------Handling for Column Name-------------
            while(cc1 not in (sheet1.columns)):
                print("Error : '{}' is not a Column Name!".format(cc1))
                cc1 = input("Re-enter Column Name: ")
                flag = flag + 1
                print(flag)
                if (flag == 2 and cc1 not in sheet1.columns):
                    print("You seem to be SLEEPY..Go have a Coffee Break !!")
                    raise Exception("'{}' is not a Column Name!".format(cc1))
                    #raise KeyError("'{}' is not a Column Name!".format(cc1))   
        dtype = sheet1['{}'.format(cc1)].dtype
        print("type  : ",dtype)
        result_sheet = cond(cc1,dtype)
        if len(result_sheet) == 0:
            print("")
            print("SORRY, No data found!!")
            print("")
        else:
            print("")
            print("{} rows x {} columns found".format(len(result_sheet),len(result_sheet.columns)))
            print("")
    return result_sheet
    #cc.append(cc1)
    
final_sheet = initial()



further_sheet = further(final_sheet)
further_sheet


#print("")
#print("Sample Output :")
#print("")
#print(final_sheet[['Class Roll','Univ. Roll', 'Student Full Name']].head())
#proceed = input("Continue ?")

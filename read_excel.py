import xlrd
import os
import datetime
import pandas as pd


def read_content(readthrough):
# To open Workbook 
    wb = xlrd.open_workbook(readthrough, logfile=open(os.devnull, 'wb')) 
    sheet = wb.sheet_by_index(0) 
      
    # For row 0 and column 0 
    sheet.cell_value(0, 0)

    Filename=loc
    datause=[]
    data=[]
    Analysis_Summary=''
    Start_date=''
    End_date=''
    Project_Name=''
    Test_Name=''
    Duration=''
    Maximum_Running_Vusers=''
    Transaction_Name=''
    SLA_Status=''
    Minimum=''
    Average=''
    Maximum=''
    Std_Deviation=''
    Percent90=''
    Pass=''
    Fail=''
    Stop=''
    Values=''
    final_listing=[]
    base_values = {}

    for i in range(0,sheet.nrows):
        datause.append(sheet.row_values(i))

    datareuse = list(filter(None,datause))

    for i in range(len(datareuse)):
        row=datareuse[i]
        if((len(row[0]))==0 or row[0]=='Test Description:' or row[0]=='Controller Run Time:' or row[0]=='User Notes:' or row[0]=='Statistics Summary'\
           or row[0]=='Total Throughput (bytes):' or row[0]=='Average Throughput (bytes/second):' or row[0]=='Total Hits:' or row[0]=='Average Hits per Second:' or row[0]=='Total Errors:' or row[0]=='5 Worst Transactions' \
           or row[0]=='No valid SLA rules for transactions available' or row[0]=='Transaction Name' or row[0]=='Application Under Test Errors'\
           or row[0]=='Transaction Summary' or row[0]=='Transactions:' or row[0]=='Action_Transaction'\
           or row[0]=='Transaction Name'):
            continue
        else:
            data.append(row)

    for i in range(len(data)):
        row=data[i]
        for j in row:
            #For ANalysis Summary
            if j=='Analysis Summary':
                Values='Analysis_Summary'
            elif j=='Project Name:':
                Values='Project_Name'
            elif j=='Test Name:':
                Values='Test_Name'
            elif j=='Duration:':
                Values='Duration'
            elif j=='Maximum Running Vusers:':
                Values='Maximum_Running_Vusers'
            elif j=='dates':
                Values='Transaction_Name'
            break
        
        if  Values=='Analysis_Summary':
            Analysis_Summary=(' '.join([str(elem) for elem in row])).replace('Analysis Summary','').replace('Period:','')
            Start_date=(Analysis_Summary.split("-"))[0]
            End_date=(Analysis_Summary.split("-"))[1]
        elif  Values=='Project_Name':
            Project_Name=(' '.join([str(elem) for elem in row])).replace('Project Name:','')
        elif  Values=='Test_Name':
            Test_Name=(' '.join([str(elem) for elem in row])).replace('Test Name:','')
        elif  Values=='Duration':
            Duration=(' '.join([str(elem) for elem in row])).replace('Duration:','')
        elif  Values=='Maximum_Running_Vusers':
            Maximum_Running_Vusers=(' '.join([str(elem) for elem in row])).replace('Maximum Running Vusers:','')
        elif  Values=='Transaction_Name':
            Tags=row[0]
            New_Tags=(Tags.split('_'))
            listing=[]
            listing.append(readthrough.strip())
            listing.append(Start_date.strip())
            listing.append(End_date.strip())
            listing.append(Project_Name.strip())
            listing.append(Test_Name.strip())
            listing.append(Duration.strip())
            listing.append(Maximum_Running_Vusers.strip())
            listing.append(row[0].strip())
            listing.append(row[1].strip())
            listing.append(row[2].strip())
            listing.append(row[3].strip())
            listing.append(row[4].strip())
            listing.append(row[5].strip())
            listing.append(row[6].strip())
            listing.append(row[7].strip())
            listing.append(row[8].strip())
            listing.append(row[9].strip())
            listing.append(New_Tags)
            final_listing.append(listing)
    return(final_listing)

def write_to_file(list_of_content,writeinto):
    f= open(writeinto,"w+")
    final_listing=list_of_content[1:]
    for l in range(len(final_listing)):
        val=final_listing[l]
        f.write('|'.join([str(elem) for elem in val]))
        f.write('\n')
    f.close()

def move_file(readthrough,move_to):
    os.rename(readthrough, move_to)



# Give the location of the folder where files will be deposited
loc = ("D:\\Python\\Scripts\\Data")
writeinto_location="D:\\Python\\Scripts\\Output"
moveinto_location="D:\\Python\\Scripts\\Moved_Data"

for root, dirs, files in os.walk(loc):
    for filename in files:
        readthrough=loc+'\\'+filename
        print(readthrough)

        #Give the location of the folder where contents will be kept

        now = datetime.datetime.now()
        date_filename=(now.isoformat()).replace(':','').replace('.','')
        writeinto=writeinto_location+'\\datafile_'+date_filename+'.txt'
        print(writeinto)

        #Give the location of the folder where read files will be moved
        
        now = datetime.datetime.now()
        temp = os.path.splitext(filename)[0]
        extension=os.path.splitext(filename)[1]
        new_filename=os.path.basename(temp)
        move_to=moveinto_location+'\\'+new_filename+date_filename+extension
        print(move_to)

        try:
            list_of_content=read_content(readthrough)
            write_to_file(list_of_content,writeinto)
            move_file(readthrough,move_to)
        except:
            print('Error')
 
    
        

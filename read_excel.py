import xlrd
import os
import datetime


def read_content(readthrough):
# To open Workbook 
    wb = xlrd.open_workbook(readthrough, logfile=open(os.devnull, 'w')) 
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
    count=0
    
    for i in range(0,sheet.nrows):
        datause.append(sheet.row_values(i))

    data = list(filter(None,datause))
    
    for i in range(len(data)):

        row=data[i]
        if row[0]=='Analysis Summary':
            Analysis_Summary=(' '.join([str(elem) for elem in row])).replace('Analysis Summary','').replace('Period:','')
            Start_date=(Analysis_Summary.split("-"))[0]
            End_date=(Analysis_Summary.split("-"))[1]
        elif row[0]=='Project Name:':
             Project_Name=(' '.join([str(elem) for elem in row])).replace('Project Name:','')
        elif row[0]=='Test Name:':
            Test_Name=(' '.join([str(elem) for elem in row])).replace('Test Name:','')
        elif row[0]=='Duration:':
            Duration=(' '.join([str(elem) for elem in row])).replace('Duration:','')
        elif row[0]=='Maximum Running Vusers:':
            Maximum_Running_Vusers=(' '.join([str(elem) for elem in row])).replace('Maximum Running Vusers:','')
        elif  row[0]+row[1]=='Transaction NameSLA Status':

            j=count
            while j<len(data):
                values=data[j]
                Tags=values[0]
                New_Tags=(Tags.split('_'))
                listing=[]
                listing.append(readthrough.strip())
                listing.append(Start_date.strip())
                listing.append(End_date.strip())
                listing.append(Project_Name.strip())
                listing.append(Test_Name.strip())
                listing.append(Duration.strip())
                listing.append(Maximum_Running_Vusers.strip())
                listing.append(values[0].strip())
                listing.append(values[1].strip())
                listing.append(values[2].strip())
                listing.append(values[3].strip())
                listing.append(values[4].strip())
                listing.append(values[5].strip())
                listing.append(values[6].strip())
                listing.append(values[7].strip())
                listing.append(values[8].strip())
                listing.append(values[9].strip())
                listing.append(New_Tags)
                final_listing.append(listing)
                j=j+1
            break
        count=count+1
                
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
        read_content(readthrough)

        try:
            list_of_content=read_content(readthrough)
            write_to_file(list_of_content,writeinto)
            move_file(readthrough,move_to)
        except:
            print('Error')

    
        

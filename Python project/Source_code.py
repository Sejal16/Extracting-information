
import sys
import pandas as pd

logs=open("C:\logFile\log.txt","a")  #opening the log file
logs.write("Name of the user: ")
name=input("Enter your name: ")
logs.write(name)
logs.write("\n")

print("Hello, "+name+"! We help you to search and read information from an excel file")
file_name=input("Enter the path of your excel file: ")


try:
   xls=pd.ExcelFile(file_name)
except IOError:  #Exception used when file is not found
    print("File not found, Please enter the correct path of your file")
    sys.exit()  # Program terminates

# Read information from a particular file
def read_OneSheet(sheet):                  
    logs=open("C:\logFile\log.txt","a")   #opening log file
    logs.write("Reading data from ")
    logs.write(sheet)
    logs.write("\n")
    df=pd.read_excel(xls,sheet,header=None)
    print()
    print(df)
    q=df.to_string()
    logs.write(q)           #Updating log file
    logs.write("\n")
    
    

# Read information from a particular file
def read_allSheet():
    logs=open("C:\logFile\log.txt","a")
    logs.write("Reading data from all the sheets\n")
    logs.write("\n")
    
    sheet_name=xls.sheet_names  #storing all the sheet names
    for tab in sheet_name:
        print()
        print("******************* "+tab+" *******************")
        logs.write("******************* ")
        logs.write(tab)
        logs.write("******************* \n")
        df=pd.read_excel(xls,tab,header=None)
        
        print()
        print(df)           #Printing the data of the given sheet
        q=df.to_string()
        logs.write(q)      #Updating log file
        logs.write("\n")
        
    


# Find the value of the given field
def field_value(sheet,field):
    
    logs=open("C:\logFile\log.txt","a")    #opening log file
    df=pd.read_excel(xls,sheet,header=None)
    

    df.set_index(0 ,inplace=True)   #sets the first column as index column
    indices = pd.Series(df.index)   #stores all the field elements
    x=indices[indices== field].index[0]  #Checks if field enter by the user is present or not
    li=list(df.iloc[x,0:])      #stores the values of the corresponding fields
    print(field+": ",li)
    
    logs.write(field)
    logs.write(": ")
    for i in li:        
        logs.write(str(i))    #Updating log file
        logs.write(" ")        
    logs.write("\n")

         

ch=1
while True:   #Infinite loop is running
    
     print()
     #Concept of Menu driven is applied
     print("1.List the names of all the sheets")
     print("2.Read any one sheet")
     print("3.Read all the sheets")
     print("4.Display the value to the field")
     print("Press -1 to end the program")
     
     ch=int(input("Choose: "))
     logs=open("C:\logFile\log.txt","a")  
     if ch==1:
         print(xls.sheet_names)         #Printing the name of the sheets in the excel file
         logs.write("Sheet names are ")
         for i in xls.sheet_names:
             logs.write(i)              #Updating the log file
             logs.write(" ")
         logs.write("\n")
         print()
     elif ch==2:
         sheet=input("Enter the name of sheeet: ")
         read_OneSheet(sheet)
     elif ch==3:
         read_allSheet()
     elif ch==4:
         sheet=input("Enter the name of sheeet: ")
         field=input("Enter the field to obtain corresponding values: ")
         print()
         field_value(sheet,field)
         print()
         
     elif ch==-1:        #terminates the program
         logs.close()
         break             
     else:
         print("Wrong choice")
  

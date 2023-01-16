import pandas as pd
import openpyxl
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth',None)
columnLetters=['D, E','F, G','H, I','J, K','L, M','N, O','P, Q']

#print(schedule)
# Search through the chosen names for the shift that starts with 8 write it to the excel sheet

days=["Monday","Tuesday","Wednesday","Thursday", "Friday", "Saturday", "Sunday"]
srcfile = openpyxl.load_workbook('Cashiers_BreakSheet.xlsx', read_only=False, keep_vba=True)





# TO-DO 
# Create an automated way so that we can increment the rows and colums were the data will be writen in the excel sheet
#Try and sort the data before searching to decrease time complexity
k=0
for k in range(7):
    shiftStart=8
    schedule=pd.read_excel("Cashier_Schedule.xlsx","schedule",usecols="C,{}".format(columnLetters[k]), skiprows=8 ,nrows=16)
    sheetname = srcfile.get_sheet_by_name(days[k])
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart and not( "CS"  in str(schedule.iloc[i,2]))):
         #write the iloc[0] ,1 ,2 to the excel sheet
         sheetname['B5'] = str(schedule.iloc[i,0])
         sheetname['C5'] = str(shiftStart)
         sheetname['E5'] = str(shiftStart+3)
         sheetname['G5'] = str(schedule.iloc[i,2])
         break       
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['B6'] = str(schedule.iloc[i,0])
            sheetname['C6'] = str(shiftStart)
            sheetname['E6'] = str(shiftStart+3)
            sheetname['G6'] = str(schedule.iloc[i,2])
            break
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['B7'] = str(schedule.iloc[i,0])
            sheetname['C7'] = str(shiftStart)
            sheetname['E7'] = str(1)
            sheetname['G7'] = str(schedule.iloc[i,2])
            
            break
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['B8'] = str(schedule.iloc[i,0])
            sheetname['C8'] = str(shiftStart)
            sheetname['E8'] = str(2)
            sheetname['G8'] = str(schedule.iloc[i,2])
            
            break
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart):
             #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['B9'] = str(schedule.iloc[i,0])
            sheetname['C9'] = str(shiftStart)
            sheetname['E9'] = str(3)
            sheetname['G9'] = str(schedule.iloc[i,2])
            
            break
    shiftStart+=1
    #shifts after 1
    shiftStart=1
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart and not( "CS"  in str(schedule.iloc[i,2])) and not( "HC"  in str(schedule.iloc[i,2]))):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['B10'] = str(schedule.iloc[i,0])
            sheetname['C10'] = str(shiftStart)
            sheetname['E10'] = str(shiftStart+3)
            sheetname['G10'] = str(schedule.iloc[i,2])
        
            break
    shiftStart+=1        
    for i in range(15):
            
        if(schedule.iloc[i,1]==shiftStart and not( "CS" and "HC" in str(schedule.iloc[i,2]))):
            
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J5'] = str(schedule.iloc[i,0])
            sheetname['K5'] = str(shiftStart)
            sheetname['M5'] = str(shiftStart+3)
            sheetname['O5'] = str(schedule.iloc[i,2])
            
            break
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart and not( "CS" and "HC"in str(schedule.iloc[i,2]))):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J6'] = str(schedule.iloc[i,0])
            sheetname['K6'] = str(shiftStart)
            sheetname['M6'] = str(shiftStart+3)
            sheetname['O6'] = str(schedule.iloc[i,2])
            
            break  
    shiftStart+=1        
    for i in range(15):
     if(schedule.iloc[i,1]==shiftStart and not( "CS" and "HC" in str(schedule.iloc[i,2]))):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J7'] = str(schedule.iloc[i,0])
            sheetname['K7'] = str(shiftStart)
            sheetname['M7'] = str(shiftStart+3)
            sheetname['O7'] = str(schedule.iloc[i,2])[0]
            
            break  
        
    #Western Union and Customer service Schedule -----------------------
    #for CS who are cashiers
    shiftStart=12
    for j in range(3):
        shiftStart+=1
        if(shiftStart==13):
            shiftStart=1    
        for i in range(7):
         if(schedule.iloc[i,1]==shiftStart and "CS" in str(schedule.iloc[i,2])):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J10'] = str(schedule.iloc[i,0])
            sheetname['K10'] = str(shiftStart)
            sheetname['M10'] = str("C.S")
            sheetname['O10'] = str(schedule.iloc[i,2])[0]
            shiftStart+=1   
            break  
    #for HC who are cashiers
    shiftStart=12
    for j in range(3):
        shiftStart+=1
        if(shiftStart==13):
            shiftStart=1    
        for i in range(7):
         if(schedule.iloc[i,1]==shiftStart and "HC" in str(schedule.iloc[i,2])):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J9'] = str(schedule.iloc[i,0])
            sheetname['K9'] = str(shiftStart)
            sheetname['M9'] = str("H.C")
            sheetname['O9'] = str(schedule.iloc[i,2])[0]
            shiftStart+=1   
            break  

    #for WU who are cashiers
    shiftStart=7
    for j in range(4):
        shiftStart+=1
        if(shiftStart==13):
            shiftStart=1    
        for i in range(7):
         if(schedule.iloc[i,1]==shiftStart and "WU" in str(schedule.iloc[i,2])):
            
            sheetname['J7'] = str(schedule.iloc[i,0])
            sheetname['K7'] = str(shiftStart)
            sheetname['M7'] = str("W.U")
            sheetname['O7'] = str(schedule.iloc[i,2])[0]
            shiftStart+=1   
            break  
    shiftStart=8
    for i in range(7):
     if(schedule.iloc[i,1]==shiftStart and "CS" in str(schedule.iloc[i,2])):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J8'] = str(schedule.iloc[i,0])
            sheetname['K8'] = str(shiftStart)
            sheetname['M8'] = str("C.S")
            sheetname['O8'] = str(schedule.iloc[i,2])[0]
            
            break           
    shiftStart+=1                                              
    schedule=pd.read_excel("Cashier_Schedule.xlsx","schedule",usecols="C,{}".format(columnLetters[k]) ,nrows=7)
 

    shiftStart=7;   
    for i in range(7):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J7'] = str(schedule.iloc[i,0])
            sheetname['K7'] = str(shiftStart) 
            sheetname['M7'] = str("C.S")
            sheetname['O7'] = str(schedule.iloc[i,2])
        
            break    
        
    shiftStart=9;   
    for i in range(7):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J7'] = str(schedule.iloc[i,0])
            sheetname['K7'] = str(shiftStart) 
            sheetname['M7'] = str("W.U")
            sheetname['O7'] = str(schedule.iloc[i,2])
        
            break    
    shiftStart+=1       
    for i in range(7):
     if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J8'] = str(schedule.iloc[i,0])
            sheetname['K8'] = str(shiftStart)
            sheetname['M8'] = str("W.U")
            sheetname['O8'] = str(schedule.iloc[i,2])
            
            break 
    for j in range(3):
        shiftStart+=1
        if(shiftStart==13):
            shiftStart=1    
        for i in range(7):
         if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J9'] = str(schedule.iloc[i,0])
            sheetname['K9'] = str(shiftStart)
            sheetname['M9'] = str(shiftStart+3)
            sheetname['O9'] = str(schedule.iloc[i,2])
            shiftStart+=1   
            break    
        
        shiftStart=1
    for j in range(4):
        shiftStart+=1  
        if(shiftStart==13):
            shiftStart=1     
        for i in range(7):
         if(schedule.iloc[i,1]==shiftStart):
            #write the iloc[0] ,1 ,2 to the excel sheet
            sheetname['J10'] = str(schedule.iloc[i,0])
            sheetname['K10'] = str(shiftStart)
            sheetname['M10'] = str("C.S")
            sheetname['O10'] = str(schedule.iloc[i,2]) 
            break              
            
            

        
srcfile.save('Schedule.xlsm')
import xlsxwriter
import xlrd
import pandas as pd
import json
from datetime import date, timedelta


Input_File = r"C:\Users\moyella\Desktop\Report Formatter\Working Script\20211019_Input.xls"
#Path of input file

#Open the file (workbook)
wb = xlrd.open_workbook(Input_File)

#Empty list for users within a spreadsheet
Users = []
Missing_Dates = []

#open the file
wb = xlrd.open_workbook(Input_File)
sheet = wb.sheet_by_index(0)
#*sheet.cell_value(0, 0)
 
for i in range(sheet.nrows):
	name = sheet.cell_value(i,1)
	if name not in Users:
		Users.append(name)
		

sorted_Users=sorted(Users)
print(sorted_Users)
print(len(sorted_Users))

#Function to create new worksheets


df = pd.read_excel('20211019_Input.xls')

#df['Date'] = df['Date'].astype(str)
df = df.sort_values(by="Date")

#df['Date']=df['Date'].dt.strftime('%d/%m/%Y')


print(df)


sdate = date(2021, 9, 17)   # start date
edate = date(2021, 10, 16)   # end date

delta = edate - sdate       # as timedelta

for i in range(delta.days + 1):
	day = sdate + timedelta(days=i)
	#fdate=day.strftime('%d/%m/%Y')
	Missing_Dates.append(day)


Missing_Dates=pd.to_datetime(Missing_Dates)	
	
	

for nme in Users:
	df_name = df[df["Name"] == nme] 
	for dte in Missing_Dates:
		if dte in df_name['Date'].values:
			continue
		else:
			print(nme,dte,"not present")
			new_row = {'Name':nme, 'Date':dte, 'Distinct Views':0, 'Distinct Edits':0}
			df=df.append(new_row, ignore_index=True)
			
			
	del df_name
	
print(df)
	

#print("finito")
	
#print(Missing_Dates)
print(df['Date'])



df = df.sort_values(by="Date")

df['Date']=df['Date'].dt.strftime('%d/%m/%Y')





#workbook = xlsxwriter.Workbook('myfile.xlsx')
#worksheet = workbook.add_worksheet()

Report = xlsxwriter.Workbook("Monthly Reviewer Statistics - 15-10-2021.xlsx")

Dict={n: grp.loc[n].to_dict('index') for n, grp in df.set_index(['Name', 'Date']).groupby(level='Name')} 

output_d = (json.dumps(Dict, indent=2))
print(output_d)
for x in sorted_Users:
	#Creates a new tab called (x)
	worksheet=Report.add_worksheet(x)
	
	print(x)
	
	row = 0
	col = 1
	
	

	for key in Dict:
		if key == x:
			worksheet.write(0, 0,"Date")
			worksheet.write(1, 0,"Distinct Edits")
			worksheet.write(2, 0,"Distinct Views")
			
			#row += 1
		
		
			totals = {'Distinct Edits':0,'Distinct Views':0}
			for item in Dict[key]:
				worksheet.write(row,col,item)
				row+=1
				for field in ['Distinct Edits','Distinct Views']:
					amount = Dict[key][item][field]
					worksheet.write(row,col,amount)
					totals[field] += amount
					row+=1
				# for j in Dict[key][item].values():
					# worksheet.write(row,col,j)
					# row+=1
				col+=1
				row-=3
			worksheet.write(row,col,'Grand Total')
			for field in ['Distinct Edits','Distinct Views']:
				worksheet.write(row+1,col,totals[field])
				row+=1

	
	#create a new tab within the report spreadsheet and name the tab x (the user)

Report.close()

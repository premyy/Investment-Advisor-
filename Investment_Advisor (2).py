import gspread
import time
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

while(True):
#--------------------Taking Permissions---------------------------------------

 scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
 account = ServiceAccountCredentials.from_json_keyfile_name('investment-advisor-370208-90c66c1b7986.json',scope)
 access = gspread.authorize(account)

#-------------------Opening Spreadsheets & getting Worksheets --------------------------------------

 Project = access.open('4a_Unit3_Project')     # MAIN SPREADSHEET
 BSE500 = Project.get_worksheet(0)             # SHEET1
 IncomeExpense = Project.get_worksheet(1)      # SHEET2    // Headers should be unique (changes done)
 FINAL_REPORT1 = Project.get_worksheet(2)      # SHEET3
 FINAL_REPORT2 = Project.get_worksheet(3)      # SHEET4
 Stock_Insights = Project.get_worksheet(4)

#-----------------------------Subtask_1 Start-----------------------------
                             
#-------------------Using Pandas for calculation , filtering , sorting etc----------------------
#-----Income / Expense Worksheet (data)-------
 data = pd.DataFrame(IncomeExpense.get_all_records())
#-----Net income (C7) -----------
 net_income = data[(data.Income_Expense == 'Income')]['INR'].sum()
 FINAL_REPORT1.update('C7', net_income)
#-----Net Expense (C8) ----------
 net_expense = data[(data.Income_Expense == 'Expense')]['INR'].sum()
 FINAL_REPORT1.update('C8', net_expense)
#-----Available Money (C24) ------
 Available_Money = net_income - net_expense
 FINAL_REPORT1.update('C24', Available_Money)
#-----Money_Message (D23&D24) ------
 FINAL_REPORT1.update('D23', "Message For You Dear Investor:")
 FINAL_REPORT1.update('D25', "Python Script Update Rate: Avg Every 15 Sec")
 if(Available_Money<0):
    FINAL_REPORT1.update('D24', "Investments cant be done, no amount .Investments Company will not be shown")
 else:
    FINAL_REPORT1.update('D24', "Investments can be done.Please Select Investment Profile for having company suggestions")
#-----Food (C10)-----------------
 Food = data[(data.Category == 'Food')]['INR'].sum()
 FINAL_REPORT1.update('C10', Food)
#-----Other (C11)-----------------
 Other = data[(data.Category == 'Other')]['INR'].sum()
 FINAL_REPORT1.update('C11', Other)
#-----Transportation (C12)-----------------
 Transportation = data[(data.Category == 'Transportation')]['INR'].sum()
 FINAL_REPORT1.update('C12', Transportation)
#-----Social Life (C13)-----------------
 Social_Life = data[(data.Category == 'Social Life')]['INR'].sum()
 FINAL_REPORT1.update('C13', Social_Life)
#-----Household (C14)-----------------
 Household = data[(data.Category == 'Household')]['INR'].sum()
 FINAL_REPORT1.update('C14', Household)
#-----Apparel (C15)-----------------
 Apparel = data[(data.Category == 'Apparel')]['INR'].sum()
 FINAL_REPORT1.update('C15',Apparel)
#-----Education (C16)-----------------
 Education  = data[(data.Category == 'Education')]['INR'].sum()
 FINAL_REPORT1.update('C16', Education )
#-----Salary (C17)-----------------
 Salary = data[(data.Category == 'Salary')]['INR'].sum()
 FINAL_REPORT1.update('C17',Salary)
#-----Allowance (C18)-----------------
 Allowance= data[(data.Category == 'Allowance')]['INR'].sum()
 FINAL_REPORT1.update('C18', Allowance)
#-----Beauty(C19)-----------------
 Beauty = data[(data.Category == 'Beauty')]['INR'].sum()
 FINAL_REPORT1.update('C19', Beauty)
#-----Gift (C20)-----------------
 Gift  = data[(data.Category == 'Gift')]['INR'].sum()
 FINAL_REPORT1.update('C20', Gift )
#-----Petty cash (C21)-----------------
 Petty_cash = data[(data.Category == 'Petty cash')]['INR'].sum()
 FINAL_REPORT1.update('C21', Petty_cash)

#-----------------------------Subtask_1 End-------------------------------


#-----------------------------Subtask_2 Start-----------------------------

#----------------------------DropDown Input (C27) Start--------------------------------
 sheetName = "Final_Report1" # Sheet_Name.
 sheetId = Project.worksheet(sheetName).id # Extracting Sheet_ID
 body = {
    "requests": [
        {
            "updateCells": {
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": 26,
                    "endRowIndex": 27,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                },
                "rows": [
                    {
                        "values": [
                            {
                                "dataValidation": {
                                    "condition": {
                                        "values": [
                                            {
                                                "userEnteredValue": "High Risk Taking"
                                            },
                                            {
                                                "userEnteredValue": "Risk Taking"
                                            },
                                            {
                                                "userEnteredValue": "Moderate Risk Taking"
                                            },
                                            {
                                                "userEnteredValue": "Low Risk Taking"
                                            }
                                        ],
                                        "type": "ONE_OF_LIST"
                                    },
                                    "showCustomUi": True
                                }
                            }
                        ]
                    }
                ],
                "fields": "dataValidation"
            }
        }
    ]
 }
 Project.batch_update(body)
#----------------------------DropDown Input (C27) End--------------------------------

#-----------------Response to User Selection(C27)-----------------------------

 risk_profile_input = FINAL_REPORT1.acell('C27').value

#1 - High Risk Taking
#2 - Risk Taking
#3 - Moderate Risk Taking
#4 - Low Risk Taking

#----Most Common Task (is  Task 1,2) for all Risk profile input logics Start-------------------------------

#------------------------------------TASK 1----------------------------------------------------------------
#Make a new column in Gsheet 1 named “Delta” and populate it with (52 Week High - price)/(52 week High)

#-----importing as pandas dataframe-----------------------------------------
 stockmarket = pd.DataFrame(BSE500.get_all_records())
 stockmarket['Delta'] =(stockmarket["52 Week High"]-stockmarket["Price"])/stockmarket["52 Week High"]
 Delta = stockmarket[['Delta']]
#-----uploding the new delta row created to gsheet--------------------------
 BSE500.update('AP1',[Delta.columns.values.tolist()] + Delta.values.tolist())

#------------------------------------TASK 2----------------------------------------------------------------

#Filter out those where Delta column is positive (>0)
 stockmarket2 = stockmarket[(stockmarket['Delta']>0)]

#----Most Common Task (is  Task 1,2) for all Risk profile input logics End-------------------------------

#------Data Cleaning of required Columns[10-Year Return(%),Dividend Per Share,Market Cap(Cr)]--------------

#-1) Nan Filteration of column
#-2) proper datatype 

 stockmarket2['10-Year Return(%)'] = stockmarket2['10-Year Return(%)'].fillna(0)
 stockmarket2['10-Year Return(%)'] = pd.to_numeric(stockmarket2['10-Year Return(%)'])
 #stockmarket['Market Cap(Cr)']=stockmarket['Market Cap(Cr)'].str.replace(',','')
 stockmarket2['Market Cap(Cr)'] = stockmarket2['Market Cap(Cr)'].fillna(0)
 stockmarket2['Market Cap(Cr)']=pd.to_numeric(stockmarket['Market Cap(Cr)'])
 stockmarket2['Dividend Per Share'] = stockmarket2['Dividend Per Share'].fillna(0)
 stockmarket2['Dividend Per Share'] = pd.to_numeric(stockmarket2['Dividend Per Share'])
 FINAL_REPORT2.batch_clear(["B2:B8"])
 FINAL_REPORT2.batch_clear(["C2:C8"])

#1----------------High Risk Taking Filters Start ------------------------------------
 if(risk_profile_input=="High Risk Taking" and Available_Money>0):
    High_Risk_Taking = stockmarket2[(stockmarket2['Market Cap(Cr)']<2000) & (stockmarket2['10-Year Return(%)']<8)]
    High_Risk_Taking1=  High_Risk_Taking.sort_values(['Dividend Per Share'],ascending=False).head(7)
    High_Risk_Taking2= High_Risk_Taking1[['Company']]
    FINAL_REPORT2.update('B2', High_Risk_Taking2.values.tolist())
      #-----Investment Money Slpit (C2:C8) ------
    if(len(High_Risk_Taking2.values.tolist())>=5):
     Investment_Money = Available_Money/5
     Investment_Money1 =  {"Money":[Investment_Money,Investment_Money,Investment_Money,Investment_Money,Investment_Money]}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
    else:
     Investment_Money = Available_Money/ len(High_Risk_Taking2.values.tolist())
     LIST=[]
     print(len(High_Risk_Taking2.values.tolist()))
     for i in  range(len(High_Risk_Taking2.values.tolist())):
        LIST.append(Investment_Money)
     Investment_Money1 =  {"Money":LIST}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
     LIST.clear()
#1------------------High Risk Taking Filters End ------------------------------------

#2----------------Risk Taking Filters Start ------------------------------------
 if(risk_profile_input=="Risk Taking" and Available_Money>0):
    Risk_Taking = stockmarket2[(stockmarket2['Market Cap(Cr)']<5000) & (stockmarket2['Market Cap(Cr)']>2000) & (stockmarket2['10-Year Return(%)']<15) & (stockmarket2['10-Year Return(%)']>8)]
    Risk_Taking1= Risk_Taking.sort_values(['Dividend Per Share'],ascending=False).head(7)
    Risk_Taking2= Risk_Taking1[['Company']]
    FINAL_REPORT2.update('B2', Risk_Taking2.values.tolist())
   #-----Investment Money Slpit (C2:C8) ------
    if(len(Risk_Taking2.values.tolist())>=5):
     Investment_Money = Available_Money/5
     Investment_Money1 =  {"Money":[Investment_Money,Investment_Money,Investment_Money,Investment_Money,Investment_Money]}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
    else:
     Investment_Money = Available_Money/ len(Risk_Taking2.values.tolist())
     LIST=[]
     print(len(Risk_Taking2.values.tolist()))
     for i in  range(len(Risk_Taking2.values.tolist())):
        LIST.append(Investment_Money)
     Investment_Money1 =  {"Money":LIST}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
     LIST.clear()




#2------------------Risk Taking Filters End ------------------------------------


#3----------------Moderate Risk Taking Filters Start ------------------------------------
 if(risk_profile_input=="Moderate Risk Taking" and Available_Money>0):
    Moderate_Risk_Taking = stockmarket2[(stockmarket2['Market Cap(Cr)']<15000) & (stockmarket2['Market Cap(Cr)']>5000) & (stockmarket2['10-Year Return(%)']<20) & (stockmarket2['10-Year Return(%)']>15)]
    Moderate_Risk_Taking1= Moderate_Risk_Taking.sort_values(['Dividend Per Share'],ascending=False).head(7)
    Moderate_Risk_Taking2= Moderate_Risk_Taking1[['Company']]
    FINAL_REPORT2.update('B2', Moderate_Risk_Taking2.values.tolist())
    #-----Investment Money Slpit (C2:C8) ------
    if(len(Moderate_Risk_Taking2.values.tolist())>=5):
     Investment_Money = Available_Money/5
     Investment_Money1 =  {"Money":[Investment_Money,Investment_Money,Investment_Money,Investment_Money,Investment_Money]}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
    else:
     Investment_Money = Available_Money/ len(Moderate_Risk_Taking2.values.tolist())
     LIST=[]
     print(len(Moderate_Risk_Taking2.values.tolist()))
     for i in  range(len(Moderate_Risk_Taking2.values.tolist())):
        LIST.append(Investment_Money)
     Investment_Money1 =  {"Money":LIST}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
     LIST.clear()

#3------------------Moderate Risk Taking Filters End ------------------------------------


#4----------------Low Risk Taking Filters Start ------------------------------------
 if(risk_profile_input=="Low Risk Taking"and Available_Money>0):
    Low_Risk_Taking = stockmarket2[ (stockmarket2['Market Cap(Cr)']>15000) & (stockmarket2['10-Year Return(%)']>20)]
    Low_Risk_Taking1= Low_Risk_Taking.sort_values(['Dividend Per Share'],ascending=False).head(7)
    Low_Risk_Taking2= Low_Risk_Taking1[['Company']]
    FINAL_REPORT2.update('B2',Low_Risk_Taking2.values.tolist())
    print(len(Low_Risk_Taking2.values.tolist()))
    #-----Investment Money Slpit (C2:C8) ------
    if(len(Low_Risk_Taking2.values.tolist())>=5):
     Investment_Money = Available_Money/5
     Investment_Money1 =  {"Money":[Investment_Money,Investment_Money,Investment_Money,Investment_Money,Investment_Money]}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
    else:
     Investment_Money = Available_Money/ len(Low_Risk_Taking2.values.tolist())
     LIST=[]
     for i in  len(Low_Risk_Taking2.values.tolist()):
        LIST.append(Investment_Money)
     Investment_Money1 =  {"Money":LIST}
     Money = pd.DataFrame(Investment_Money1)
     FINAL_REPORT2.update('C2:C8', Money.values.tolist())
     LIST.clear()
    # 10 - sec update complete code 
#4------------------Low Risk Taking Filters End ------------------------------------

#-----------------------------Subtask_2 End-----------------------------

#-----------------------------Subtask_3 Start-----------------------------
# 1)Compare the median of column Enterprise Value(Cr) across different Sectors.
# 1)For instance, what is the median enterprise value of Technology sector as compared to Services sector
 Stock_Insights.clear()
 sector = {"sectors":stockmarket['Sector'].sort_values(ascending=True).unique()}
 sector1 = pd.DataFrame(sector)
 Stock_Insights.update('A1:A',[['Sectors']]+sector1.values.tolist())
 #stockmarket['Enterprise Value(Cr)']=stockmarket['Enterprise Value(Cr)'].str.replace(',','').fillna(0)
 stockmarket['Enterprise Value(Cr)']=pd.to_numeric(stockmarket['Enterprise Value(Cr)'])
 sector_values=stockmarket.groupby('Sector').median()
 sector_values=sector_values.round(3).fillna(0)
 sector_values=pd.DataFrame(sector_values['Enterprise Value(Cr)'])
 Stock_Insights.update('B1',[sector_values.columns.values.tolist()])
 Stock_Insights.update('B2',sector_values.values.tolist())

# 2)Try to find a relation between Dividend Per Share with Market Cap(Cr)
# 3)Count the companies in different Industry with positive and negative 3-Year Return.




# 3)For instance how many companies in Drugs & Pharma industry have positive 3-Year Return and how many have that negative. 
# 3)Basis this, decide which industry would you recommend someone to invest if the same return is followed
 #-------------------------Data Cleaning --------------------------------
 stockmarket['3-Year Return']=stockmarket['3-Year Return'].fillna(0)
 stockmarket['3-Year Return']=pd.to_numeric(stockmarket['3-Year Return'])
 #-------------------------Data Cleaning --------------------------------
 #---------------Positive 3 yr-------------------------------
 Positive_C=stockmarket[(stockmarket['3-Year Return']>0)]['Industry'].sort_values(axis= 0, ascending = True).unique()
 Positive_C=pd.DataFrame(Positive_C)
 Stock_Insights.update('A20',[["Industry"]]+Positive_C.values.tolist())
 Positive=stockmarket[(stockmarket['3-Year Return']>0)][['Company','Industry']].groupby('Industry').count()
 Stock_Insights.update('B20',[["3 Year Return (+) Company"]]+Positive.values.tolist())
 #---------------Negative 3 yr-------------------------------
 Negative_C=stockmarket[(stockmarket['3-Year Return']<0)]['Industry'].sort_values(axis= 0, ascending = True).unique()
 Negative_C=pd.DataFrame(Negative_C)
 Stock_Insights.update('C20',[["Industry"]]+Negative_C.values.tolist())
 Negative=stockmarket[(stockmarket['3-Year Return']<0)][['Company','Industry']].groupby('Industry').count()
 Stock_Insights.update('D20',[["3 Year Return (-) Company"]]+Negative.values.tolist())

# 4)Come up with any one KPI which can help define the best stock across different Sector, 
# 4)you may need to learn a little bit of Finance for the same

#-----------------------------Subtask_3 End-----------------------------
 time.sleep(10)

# -*- coding: utf-8 -*-
"""
Created on Mon Mar 22 14:32:20 2021

@author: khambarde
"""
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 18:37:59 2020

@author: khambarde
"""
# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 15:06:52 2020

@author: vsharma
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Oct 27 15:41:48 2020

@author: vsharma
"""
#cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
def CensusGenerationFun():
    import xlsxwriter
    from datetime import datetime
    import pyodbc 
    import os
    import pandas as pd
    from Masking_Data import Mask_SSN
    from Masking_Data import Mask_DOB
    #from dateutil.parser import parse
    # Start of Main       
    ################## SQL SERVER CONNECTION#####################################################################
    db = pyodbc.connect("Driver={SQL Server};"
                          "Server=NDS-AA-02;"
                          "Database=HPU_DEV;"
                          "uid=RPA;pwd=nds1@2020;")
    cursor = db.cursor()
    
    # Client ID##############################
    sqlExecSP="{call USP_GetEMailLogID}"
    dfMailListToBeClassified = pd.read_sql_query(sql=sqlExecSP, con=db)
      
    EMAILID = int(dfMailListToBeClassified['ID'][0])
    CLIENTID = dfMailListToBeClassified['CLIENTID'][0]
    emailSender = dfMailListToBeClassified['FROM_EMAIL_ADDRESS'][0]
    # ClientID = int(CLIENTID)
    # InEmailID = int(EMAILID)
    ClientID = int(471)
    InEmailID = int(1945)
    
    try:
        ########################################
        # Get dates in diffferent formats
        arg = (ClientID,InEmailID)
        cursor.execute("{CALL dbo.CheckInEmailStatus(?,?)}",arg)
        EmailStatus_All = cursor.fetchall()
        EmailStatus_All_Row_0 = EmailStatus_All[0]
        # print(int(EmailStatus_All_Row_0[0]))
        if int(EmailStatus_All_Row_0[0])>0:   
            print(ClientID)
            print(InEmailID)
            cursor.execute("{CALL dbo.GET_DATES_IN_DIFF_FORMS(?)}",ClientID)
            Date_All = cursor.fetchall()
            Date_All_Row_0 = Date_All[0]
            ##############################################
            MonthOfReport_MMM_YY = Date_All_Row_0[0]
            MonthOfReport_MMM_yy_Minus_1 = Date_All_Row_0[1]
            MonthOfReport_MMM_yy_Minus_2 = Date_All_Row_0[2]
            MonthOfReport_MMM_yy_Minus_3 = Date_All_Row_0[3]
            MonthOfReport_MM_yy = Date_All_Row_0[4].strip()
            MonthOfReport_MM_yy_Minus_1 = Date_All_Row_0[5]
            MonthOfReport_MM_yy_Minus_2 = Date_All_Row_0[6]
            MonthOfReport_MM_yy_Minus_3 = Date_All_Row_0[7]
            MonthOfReport_MM_yy_Minus_4 = Date_All_Row_0[8]
            MonthOfReport_MM_yy_Minus_5 = Date_All_Row_0[9]
            MonthOfReport_MM_yy_Minus_6 = Date_All_Row_0[10]
            MonthOfReport_MM_yy_Plus_1 = Date_All_Row_0[11]
            DATE_OF_REPORT = Date_All_Row_0[12]
            #DayAndMonthOfReport = Date_All_Row_0[13]
            DateOfReport_Month_Year_Obj = datetime.strptime(MonthOfReport_MMM_YY,'%Y-%m-%d')
            DateOfReport_Month_Year = datetime.strftime(DateOfReport_Month_Year_Obj,'%d %b %Y')
            
            MonthOfReport_Excel = datetime.strptime(MonthOfReport_MMM_YY,'%Y-%m-%d')
            MonthOfReport_Minus_1_Excel = datetime.strptime(MonthOfReport_MMM_yy_Minus_1,'%Y-%m-%d')
            MonthOfReport_Minus_2_Excel = datetime.strptime(MonthOfReport_MMM_yy_Minus_2,'%Y-%m-%d')
            MonthOfReport_Minus_3_Excel = datetime.strptime(MonthOfReport_MMM_yy_Minus_3,'%Y-%m-%d')
            ##############################################
            # Get Client information
            cursor.execute("{CALL dbo.GetClientInfoFromClientID(?)}",ClientID)
            Client_info = cursor.fetchall()
            Client_info_Row_0 = Client_info[0]
            #######################################
            PAY_DUE_DATE = Client_info_Row_0[0]  
            REINSTATEMENT_DATE = Client_info_Row_0[1].strip()
            CENSUS_CUT_OFF_DATE = Client_info_Row_0[2].strip()
            CENSUS_FOLDER_PATH = Client_info_Row_0[3].strip()
            try:
                CLIENT_STATE = Client_info_Row_0[4].strip()
            except:
                CLIENT_STATE = ''
            try:    
                POLICYHOLDER = Client_info_Row_0[5].strip()
            except:
                POLICYHOLDER = ''
            try:    
                NAMED_MOTORCARRRIER = Client_info_Row_0[6].strip()
            except:
                NAMED_MOTORCARRRIER = ''
            try:    
                OA_POLICY_NUMBER = Client_info_Row_0[7].strip()
            except:
                OA_POLICY_NUMBER = ''
            try:                
                CL_POLICY = Client_info_Row_0[8].strip()
            except:
                CL_POLICY = ''
            try:                
                CL_POLICY_NUMBER = Client_info_Row_0[9].strip()
            except:
                CL_POLICY_NUMBER = ''
            OARATE = Client_info_Row_0[10]
            CLRATE = Client_info_Row_0[11]
            DUESRATE = Client_info_Row_0[12]
            CLIENT_EFFECTIVE_DATE = Client_info_Row_0[13]
            CLIENT_EXPIRATION_DATE = Client_info_Row_0[14]
            OA_COMMISSION_RATE = Client_info_Row_0[15]
            CLIENT_INVOICE = Client_info_Row_0[16]
            INTERNAL_INVOICE = Client_info_Row_0[17]
            LATE_FEE_AMOUNT = Client_info_Row_0[18]
            REINSTATEMENT_AMOUNT = Client_info_Row_0[19]
            CL_COMMISSION_RATE = Client_info_Row_0[20]
            NAMED_MOTORCARRRIER_PLANE = Client_info_Row_0[21]
            print(NAMED_MOTORCARRRIER_PLANE)
            
          
            
             
            
            ADDCOUNT_Final = 0
            DELETECOUNT_Final = 0
            ADD_DELETECOUNT_Final = 0
            DEBITCOUNT_Final = 0 
            CREDITCOUNT_Final = 0
            
            cursor.execute("{CALL USP_GetRateInfo(?)}",ClientID)
            Rate_info = cursor.fetchall()
            IS_PA_RATE = Rate_info[0][0]
              
            #CHANGE HERE FOR BROKER_ID = 1
            BROKER_ID = Rate_info[0][4]
            if BROKER_ID is None:
                BROKER_ID = 0
            print(BROKER_ID)
            
            
            
            if IS_PA_RATE == False:
                IS_PA_RATE = 0
                PA_COMMISSION_RATE = 0
                PA_RATE = 0
        
            else:    
                PA_RATE = Rate_info[0][1]                
                IS_CL_RATE = Rate_info[0][2]
                PA_COMMISSION_RATE = Rate_info[0][3]    
            
            IS_CL_RATE = Rate_info[0][2]
        
        
            #USP_GetLumpsum_cl_Month_YearDetails
            cursor.execute("{CALL USP_GetLumpsum_cl_Month_YearDetails(?)}",ClientID)
            Rate_info = cursor.fetchall()
            LUMPSUM_CL_AMOUNT = 0
            
            if len(Rate_info) >=1:                    
                IS_LUMPSUM_CL = Rate_info[0][0]
                LUMPSUM_CL_AMOUNT = Rate_info[0][1]
                LUMPSUM_CL_MONTH_YEAR = Rate_info[0][2]
                print(LUMPSUM_CL_MONTH_YEAR)
                
                if IS_LUMPSUM_CL:
                    import datetime
                    LumpsumDate = datetime.datetime.strptime(LUMPSUM_CL_MONTH_YEAR,'%Y-%m-%d')
                    Month_Lumpsum = LumpsumDate.strftime('%m')
                    Year_Lumpsum = LumpsumDate.strftime('%y')
                    
                    CensusDate = datetime.datetime.strptime(MonthOfReport_MMM_YY,'%Y-%m-%d')    
                    Month_Census = CensusDate.strftime('%m')
                    Year_Census = CensusDate.strftime('%y')                
                    
                    if Month_Lumpsum == Month_Census and Year_Lumpsum == Year_Census:
                        Lumpsum_Flag = 'Yes'
                    else:
                        Lumpsum_Flag = 'No'
                        LUMPSUM_CL_AMOUNT = 0
                else:
                    Lumpsum_Flag = 'No'
                    LUMPSUM_CL_AMOUNT = 0
        
                                      
            from datetime import datetime
            
            
            cursor.execute("{CALL USP_Number_Of_Credit_Debit(?)}",ClientID)
            Number_Of_Creditebit = cursor.fetchall()
        
            if len(Number_Of_Creditebit) == 1:
                US_Number_Of_Credit_Debit = Number_Of_Creditebit[0][0]
            
        
            def FileNameBroker():
                BROKER_FILE_NAME = NAMED_MOTORCARRRIER_PLANE+"_Census Report_" + DateOfReport_Month_Year+"_To_Broker.xlsx"
                Broker_File_Path = CENSUS_FOLDER_PATH + "\\" + BROKER_FILE_NAME
                try:
                    if os.path.exists(Broker_File_Path):
                        os.remove(Broker_File_Path,ignore_errors=True)
                except:
                    pass      
            

                workbook_Broker = xlsxwriter.Workbook(Broker_File_Path)
                worksheet_Broker = workbook_Broker.add_worksheet()
                worksheet_Broker.set_column('A:A', 25)
                worksheet_Broker.set_column('B:B', 25)
                worksheet_Broker.set_column('C:C', 20)
                worksheet_Broker.set_column('D:D', 10)
                worksheet_Broker.set_column('E:E', 15)
                worksheet_Broker.set_column('F:F', 15)
                worksheet_Broker.set_column('G:G', 13)
                worksheet_Broker.set_column('H:H', 13)
                worksheet_Broker.set_column('I:I', 13)
                worksheet_Broker.set_column('J:J', 15)
                worksheet_Broker.set_column('K:K', 10)
                worksheet_Broker.set_column('L:L', 15)
                worksheet_Broker.set_column('M:M', 13)
                worksheet_Broker.set_column('N:N', 13)
                worksheet_Broker.set_column('O:O', 13)
                worksheet_Broker.set_column('P:P', 13)
                worksheet_Broker.set_landscape()
                worksheet_Broker.set_paper(1)
                worksheet_Broker.set_print_scale(52)
                worksheet_Broker.set_margins(left=0.45,right=0.25,top=0.75, bottom=0.75)

                return Broker_File_Path,workbook_Broker,worksheet_Broker,BROKER_FILE_NAME
            
            def FileNameInternal():
                INTERNAL_FILE_NAME = NAMED_MOTORCARRRIER_PLANE+"_Census Report_" + DateOfReport_Month_Year+"_Internal.xlsx"
                Internal_File_Path = CENSUS_FOLDER_PATH + "\\" + INTERNAL_FILE_NAME
                try:
                    if os.path.exists(Internal_File_Path):
                        os.remove(Internal_File_Path,ignore_errors=True)
                except:
                    pass      
            
                workbook_Internal = xlsxwriter.Workbook(Internal_File_Path)
                worksheet_Internal = workbook_Internal.add_worksheet()
                worksheet_Internal.set_column(0,0, 20)
                worksheet_Internal.set_column('B:B', 20)
                worksheet_Internal.set_column('C:C', 20)
                worksheet_Internal.set_column('D:D', 10)
                worksheet_Internal.set_column('E:E', 15)
                worksheet_Internal.set_column('F:F', 15)
                worksheet_Internal.set_column('G:G', 10)
                worksheet_Internal.set_column('H:H', 10)
                worksheet_Internal.set_column('I:I', 10)
                worksheet_Internal.set_column('J:J', 15)
                worksheet_Internal.set_column('K:K', 10)
                worksheet_Internal.set_column('L:L', 15)
                worksheet_Internal.set_column('M:M', 13)
                worksheet_Internal.set_column('N:N', 10)
                worksheet_Internal.set_column('O:O', 10)
                worksheet_Internal.set_column('P:P', 10)
                worksheet_Internal.set_landscape()
                worksheet_Internal.set_paper(1)
                worksheet_Internal.set_print_scale(52)
                return Internal_File_Path,workbook_Internal,worksheet_Internal,INTERNAL_FILE_NAME
            
            def FileNameClient():
                CLIENT_FILE_NAME = NAMED_MOTORCARRRIER_PLANE+"_Census Report_" + DateOfReport_Month_Year+"_To_Client.xlsx"
                Client_File_Path = CENSUS_FOLDER_PATH + "\\" + CLIENT_FILE_NAME
                try:
                    if os.path.exists(Client_File_Path):
                        os.remove(Client_File_Path,ignore_errors=True)
                except:
                    pass      
            
                workbook_Client = xlsxwriter.Workbook(Client_File_Path)
                worksheet_Client = workbook_Client.add_worksheet()
                worksheet_Client.set_column('A:A', 25)
                worksheet_Client.set_column('B:B', 25)
                worksheet_Client.set_column('C:C', 20)
                worksheet_Client.set_column('D:D', 10)
                worksheet_Client.set_column('E:E', 15)
                worksheet_Client.set_column('F:F', 15)
                worksheet_Client.set_column('G:G', 13)
                worksheet_Client.set_column('H:H', 13)
                worksheet_Client.set_column('I:I', 13)
                worksheet_Client.set_column('J:J', 15)
                worksheet_Client.set_column('K:K', 13)
                worksheet_Client.set_column('L:L', 15)
                worksheet_Client.set_column('M:M', 13)
                worksheet_Client.set_column('N:N', 13)
                worksheet_Client.set_column('O:O', 13)
                worksheet_Client.set_column('P:P', 13)
                worksheet_Client.set_landscape()
                worksheet_Client.set_paper(1)
                worksheet_Client.set_print_scale(52)
                return Client_File_Path,workbook_Client,worksheet_Client,CLIENT_FILE_NAME
            ##############################################################################
            def FieldCopyFun(workbook,worksheet):
            ########### HEADER INFO WRITE ########
            #######################################
                Row = "1"
                #####################################
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'POLICYHOLDER',cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, POLICYHOLDER, cell_format)
                ######################################
                Row = "2"
                #################################
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'NAMED MOTORCARRIER',cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, NAMED_MOTORCARRRIER, cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('F'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12','align': 'centre'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('G'+Row, 'STATE', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12','align': 'centre'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('H'+Row, CLIENT_STATE, cell_format)
                
                ######################################
                Row = "3"
                #################################
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'OCCUPATIONAL ACCIDENT POLICY NUMBER',cell_format) 
                
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12'})
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_left(1)
                cell_format.set_bg_color('yellow')
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, OA_POLICY_NUMBER,cell_format) 
                
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_right(1)
                cell_format.set_font_name('Arial')
                worksheet.write('F'+Row, '', cell_format)
                
                from datetime import datetime
                cell_format = workbook.add_format({'bold': True,'font_size': '12','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('G'+Row, 'EFFECTIVE',cell_format) 
                
                
                CLIENT_EFFECTIVE_DATE_Format = datetime.strptime(CLIENT_EFFECTIVE_DATE,'%Y-%m-%d')                            
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_bg_color('yellow')
                cell_format.set_font_name('Arial')
                cell_format.set_num_format('mm/dd/yy')
                worksheet.write('H'+Row, CLIENT_EFFECTIVE_DATE_Format,cell_format) 
                
                #cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12','align': 'centre'})
                #cell_format.set_border(1)
                #cell_format.set_bg_color('yellow')
                #cell_format.set_font_name('Arial')
                #cell_format.set_num_format('mmm-yy')
                #worksheet.write('I'+Row, date_time,cell_format)
                
                ######################################
                Row = "4"
                #################################
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_left(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('F'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('G'+Row, 'EXPIRATION',cell_format) 
                
                try:
                   CLIENT_EXPIRATION_DATE_Format = datetime.strptime(CLIENT_EXPIRATION_DATE,'%Y-%m-%d')
                except:
                   CLIENT_EXPIRATION_DATE_Format = ''                             
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_bg_color('yellow')
                cell_format.set_font_name('Arial')
                cell_format.set_num_format('mm/dd/yy')
                worksheet.write('H'+Row, CLIENT_EXPIRATION_DATE_Format,cell_format)
                
                ######################################
                Row = "5"
                #################################
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'CONTINGENT LIABILITY',cell_format) 
                
                cell_format = workbook.add_format({'bold': True,'font_color': 'red','font_size': '12','align': 'centre'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_top(1)
                cell_format.set_right(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, 'YES/NO?', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_bg_color('yellow')
                cell_format.set_font_name('Arial')
                worksheet.write('G'+Row, CL_POLICY,cell_format) 
                
                ######################################
                Row = "6"
                #################################
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'CONTINGENT LIABILITY POLICY NUMBER',cell_format) 
                
                #cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12'})
                #cell_format.set_border(1)
                #cell_format.set_font_name('Arial')
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                #worksheet = workbook.add_worksheet()
                merge_format = workbook.add_format({
                    'bold': 1,
                    'border': 1,
                    'align': 'right',
                    'font_name':'Arial',
                    'font_color':'blue',
                    'font_size': '12'})
                
                worksheet.merge_range('F6:G6',CL_POLICY_NUMBER, merge_format)
                #################################
                
                
                #worksheet.write('F'+Row, 'Q0014-CL18-0027513Q',cell_format)
                
                ######################################
                Row = "7"
                #########################################
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'DATE OF MONTHLY REPORT',cell_format) 
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('F'+Row, '', cell_format)
                
                
                #Date3 = '2020-10-27'
                #DATE_OF_REPORT_Format = datetime.strptime(DATE_OF_REPORT,'%Y-%m-%d')
                cell_format = workbook.add_format({'bold': True,'font_color': 'blue','font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_bg_color('yellow')
                cell_format.set_font_name('Arial')
                cell_format.set_num_format('dd-mmm-yy')
                worksheet.write('G'+Row, DATE_OF_REPORT,cell_format)
                
                
                ######################################
                Row = "8"
                #########################################
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'TOTAL NUMBER OF DRIVERS',cell_format) 
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('B'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('C'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('D'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('E'+Row, '', cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '12'})
                #cell_format.set_bg_color('yellow')
                cell_format.set_top(1)
                cell_format.set_bottom(1)
                cell_format.set_font_name('Arial')
                worksheet.write('F'+Row, '', cell_format)
                
                
                
                ######################################
                Row = str(int(Row) + 2)
                #########################################
                cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                worksheet.write('A'+Row, 'FIRSTNAME',cell_format)
                worksheet.write('B'+Row, 'LASTNAME',cell_format)
                
                cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue','align': 'centre'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                
                worksheet.write('C'+Row, 'APPLICANT SSN',cell_format)
                worksheet.write('D'+Row, 'DOB',cell_format)
                worksheet.write('E'+Row, 'EFFDATE',cell_format)
                worksheet.write('F'+Row, 'TERMDATE',cell_format)
                worksheet.write('G'+Row, 'STATE',cell_format)
                worksheet.write('H'+Row, 'RFI / RF',cell_format)
                worksheet.write('I'+Row, '#DRIVERS',cell_format)
                worksheet.write('J'+Row, 'STATUS',cell_format)
                worksheet.write('K'+Row, 'MONTH',cell_format)
                worksheet.write('L'+Row, 'TOTALRATE',cell_format)
                worksheet.write('M'+Row, 'OARATE',cell_format)
                
                if IS_CL_RATE == 1 and IS_PA_RATE == 1:
                    worksheet.write('N'+Row, 'CLRATE',cell_format)
                    worksheet.write('O'+Row, 'PARATE',cell_format)
                    worksheet.write('P'+Row, 'DUES',cell_format)
                
                
                elif IS_CL_RATE == 1:
                    worksheet.write('N'+Row, 'CLRATE',cell_format)
                    worksheet.write('O'+Row, 'DUES',cell_format)
                
                elif IS_PA_RATE == 1:
                    
                    worksheet.write('N'+Row, 'PARATE',cell_format)
                    worksheet.write('O'+Row, 'DUES',cell_format)
                
                else:
                    worksheet.write('N'+Row, 'DUES',cell_format)
        
        
                
                # ADD ACTIVE DRIVERS  ################################
                # Get active drivers
                DataRowFirst = str(int(Row) + 1)
                ADDCOUNT = 0
                DELETECOUNT = 0
                ADD_DELETECOUNT = 0
                DEBITCOUNT = 0 
                CREDITCOUNT = 0
                TOTAL_DRIVER_COUNT = 0 
                cursor.execute("{CALL dbo.GetACTIVEDriversFromClientID(?)}",ClientID)
                Active_Drivers = cursor.fetchall()
                for i in range(0,len(Active_Drivers),1):
                    Row = str(int(Row) + 1)
                    Active_Drivers_Row_i = Active_Drivers[i]
                    FirstName = Active_Drivers_Row_i[0]
                    LastName = Active_Drivers_Row_i[1]
                    DRIVER_SSN = Active_Drivers_Row_i[2]
                    DRIVER_SSN = Mask_SSN(DRIVER_SSN)
                    DRIVER_STATE = Active_Drivers_Row_i[3]
                    RFI_RF = Active_Drivers_Row_i[4]
                    DRIVER_DOB = Active_Drivers_Row_i[5]
                    DRIVER_EFF_DATE = Active_Drivers_Row_i[6]
                    try:
                        #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                        DRIVER_DOB_Excel = Mask_DOB(DRIVER_DOB)
                    except:
                        DRIVER_DOB_Excel = ''
                    try:   
                        DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                    except:
                        DRIVER_EFF_DATE_Excel = ''
                    
                    DRIVER_STATUS = 'ACTIVE'
                    No_of_Drivers = 1
                    OARATE_Final = OARATE
                    CLRATE_Final = CLRATE
                    DUESRATE_Final = DUESRATE
                    if IS_PA_RATE:                
                        PA_RATE_Final = PA_RATE
                    else:
                        PA_RATE_Final = 0
                    
                    #PA_RATE_Final fetch from database
                    if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                    
                    elif IS_CL_RATE==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                    
                    elif IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                    
                    else:
                        Total_Rate = float(OARATE_Final)+int(DUESRATE_Final)
                    
                    cell_format = workbook.add_format({'font_size': '11'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, FirstName,cell_format) # First Name
                    worksheet.write('B'+Row, LastName,cell_format) # Last Name
                    
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    
                    worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                    worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                    worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                    worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                    TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                    worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                    
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mm/dd/yy')
                    worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                    worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                    worksheet.write('F'+Row, '',cell_format) # TERMDATE                                     
                    
                    cell_format = workbook.add_format({'num_format': '#,##0.00','align': 'right'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                    worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                    
                    if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                       
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                        worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                        worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_CL_RATE == 1:
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_PA_RATE == 1:
                        worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    else:
                        worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mmm-yy')
                    #print(MonthOfReport_MMM_YY)
                    
                    worksheet.write('K'+Row, MonthOfReport_Excel,cell_format) # MONTH
                
                # ADD ADD DRIVERS  ################################
                # Get Add drivers
                cursor.execute("{CALL dbo.GetADDDriversFromClientID(?)}",ClientID)
                Add_Drivers = cursor.fetchall()
                for i in range(0,len(Add_Drivers),1):
                    second_flag = ''
                    first_flag = ''
                    ADDCOUNT = ADDCOUNT + 1
                    Row = str(int(Row) + 1)
                    Add_Drivers_Row_i = Add_Drivers[i]
                    FirstName = Add_Drivers_Row_i[0]
                    LastName = Add_Drivers_Row_i[1]
                    DRIVER_SSN = Add_Drivers_Row_i[2]
                    DRIVER_SSN = Mask_SSN(DRIVER_SSN)
                    DRIVER_DOB = Add_Drivers_Row_i[3]
                    DRIVER_EFF_DATE = Add_Drivers_Row_i[4]
                    DRIVER_STATE = Add_Drivers_Row_i[5]
                    RFI_RF = Add_Drivers_Row_i[6]
                    EffDate_MM_YY = Add_Drivers_Row_i[7]
                    EffDayInt = int(Add_Drivers_Row_i[8])
                    RETRO_ACTIVE_DAYS = Add_Drivers_Row_i[9]
                    if RETRO_ACTIVE_DAYS is None:
                        RETRO_ACTIVE_DAYS = 0
                    else:
                        RETRO_ACTIVE_DAYS = int(Add_Drivers_Row_i[9])
        
                    #print(EffDayInt)
                    
                    #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                    #DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                    try:
                        #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                        DRIVER_DOB_Excel = Mask_DOB(DRIVER_DOB)
                    except:
                        DRIVER_DOB_Excel = ''
                    try:   
                        DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                    except:
                        DRIVER_EFF_DATE_Excel = ''
                    
                    DRIVER_STATUS = 'ADD'
                    if ((EffDate_MM_YY == MonthOfReport_MM_yy) and (EffDayInt>15)) or (EffDate_MM_YY == MonthOfReport_MM_yy_Plus_1):
                        No_of_Drivers = 0
                        OARATE_Final = 0
                        CLRATE_Final = 0
                        DUESRATE_Final = 0
                        PA_RATE_Final = 0
                    else:    
                        No_of_Drivers = 1
                        OARATE_Final = float(OARATE)
                        CLRATE_Final = float(CLRATE)
                        DUESRATE_Final = float(DUESRATE)
                        if IS_PA_RATE:
                            PA_RATE_Final = float(PA_RATE)
                        else:
                            PA_RATE_Final = 0
        
                        
                    
                     #PA_RATE_Final fetch from database
                    if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                    
                    elif IS_CL_RATE==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                    
                    elif IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                    
                    else:
                        Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                                          
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, FirstName,cell_format) # First Name
                    worksheet.write('B'+Row, LastName,cell_format) # Last Name
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    
                    worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                    worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                    worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                    worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                    TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                    worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mm/dd/yy')
                    worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                    worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                    worksheet.write('F'+Row, '',cell_format) # TERMDATE
                   
                    
                    
                    cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                    worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                    
                    if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                       
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                        worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                        worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_CL_RATE == 1:
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_PA_RATE == 1:
                        worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    else:
                        worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mmm-yy')
                    
                    worksheet.write('K'+Row, MonthOfReport_Excel,cell_format) # MONTH
        
                    if ((EffDate_MM_YY == MonthOfReport_MM_yy_Minus_1) and (EffDayInt<16)) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_2) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_3) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_4) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_5) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_6):
                        first_flag = 'yes'
                        Row = str(int(Row) + 1)
                        DEBITCOUNT = DEBITCOUNT + 1
                        DRIVER_STATUS = 'DB'
                        No_of_Drivers = 1
                        OARATE_Final = float(OARATE)
                        CLRATE_Final = float(CLRATE)
                        DUESRATE_Final = float(DUESRATE)
                        if IS_PA_RATE:
                            PA_RATE_Final = float(PA_RATE)
                        else:
                            PA_RATE_Final = 0
        
                                          #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                        cell_format.set_font_name('Arial')
                    
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, '',cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                            
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                           
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                                                                    
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                    
                    if ((EffDate_MM_YY == MonthOfReport_MM_yy_Minus_2) and US_Number_Of_Credit_Debit==2 and (EffDayInt<16)) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_3 and US_Number_Of_Credit_Debit == 2) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_4 and US_Number_Of_Credit_Debit == 2) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_5 and US_Number_Of_Credit_Debit == 2) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_6 and US_Number_Of_Credit_Debit == 2):
                        second_flag = 'yes'
                        Row = str(int(Row) + 1)
                        DEBITCOUNT = DEBITCOUNT + 1
                        DRIVER_STATUS = 'DB'
                        No_of_Drivers = 1
                        OARATE_Final = float(OARATE)
                        CLRATE_Final = float(CLRATE)
                        DUESRATE_Final = float(DUESRATE)
                        if IS_PA_RATE:
                            PA_RATE_Final = float(PA_RATE)
                        else:
                            PA_RATE_Final = 0
        
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                       
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, '',cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                            
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                           
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_2_Excel,cell_format) # MONTH    
                    
                    if RETRO_ACTIVE_DAYS ==1 and first_flag != 'yes':
                       Row = str(int(Row) + 1)
                       DEBITCOUNT = DEBITCOUNT + 1
                       DRIVER_STATUS = 'DB'
                       No_of_Drivers = 1
                       OARATE_Final = float(OARATE)
                       CLRATE_Final = float(CLRATE)
                       DUESRATE_Final = float(DUESRATE)
                                         #PA_RATE_Final fetch from database
                       if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                           Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                       
                       elif IS_CL_RATE==1:
                           Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                       
                       elif IS_PA_RATE ==1:
                           Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                       
                       else:
                           Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                       
                       cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                       cell_format.set_font_name('Arial')
                   
                       worksheet.write('A'+Row, FirstName,cell_format) # First Name
                       worksheet.write('B'+Row, LastName,cell_format) # Last Name
                       
                       cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                       cell_format.set_font_name('Arial')
                       
                       worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                       worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                       worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                       worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                       TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                       worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                       
                       cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                       cell_format.set_font_name('Arial')
                       cell_format.set_num_format('mm/dd/yy')
                       worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                       worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                       worksheet.write('F'+Row, '',cell_format) # TERMDATE
                      
                       
                       
                       cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                       cell_format.set_font_name('Arial')
                       worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                       worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                           
                       if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                          
                           worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                           worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                           worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                       
                       elif IS_CL_RATE == 1:
                           worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                           worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                       
                       elif IS_PA_RATE == 1:
                           worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                           worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                       else:
                           worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                       
                   
                       
                       cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                       cell_format.set_font_name('Arial')
                       cell_format.set_num_format('mmm-yy')
                       worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
           
                    #########
                    
                    if RETRO_ACTIVE_DAYS ==1 and US_Number_Of_Credit_Debit == 2 and second_flag != 'yes':               
                        Row = str(int(Row) + 1)
                        DEBITCOUNT = DEBITCOUNT + 1
                        DRIVER_STATUS = 'DB'
                        No_of_Drivers = 1
                        OARATE_Final = float(OARATE)
                        CLRATE_Final = float(CLRATE)
                        DUESRATE_Final = float(DUESRATE)
                                          #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                        cell_format.set_font_name('Arial')
                    
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, '',cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                            
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                           
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                    
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                      
                        
                
        
                
                
                # ADD ADD_DELETE DRIVERS  ################################
                # Get Add_Delete drivers
                cursor.execute("{CALL dbo.GetADD_DELETEDriversFromClientID(?,?)}",ClientID,InEmailID)
                Add_Delete_Drivers = cursor.fetchall()
                
                for i in range(0,len(Add_Delete_Drivers),1):
                    ADD_DELETECOUNT = ADD_DELETECOUNT + 1
                    Row = str(int(Row) + 1)
                    Add_Delete_Drivers_Row_i = Add_Delete_Drivers[i]
                    FirstName = Add_Delete_Drivers_Row_i[0]
                    LastName = Add_Delete_Drivers_Row_i[1]
                    DRIVER_SSN = Add_Delete_Drivers_Row_i[2]
                    DRIVER_SSN = Mask_SSN(DRIVER_SSN)
                    DRIVER_DOB = Add_Delete_Drivers_Row_i[3]
                    DRIVER_EFF_DATE = Add_Delete_Drivers_Row_i[4]
                    DRIVER_STATE = Add_Delete_Drivers_Row_i[5]
                    RFI_RF = Add_Delete_Drivers_Row_i[6]
                    EffDate_MM_YY = Add_Delete_Drivers_Row_i[7]
                    EffDayInt = int(Add_Delete_Drivers_Row_i[8])
                    TERMINATION_DATE = Add_Delete_Drivers_Row_i[9]
                    effday = Add_Delete_Drivers_Row_i[10]
                    ADD_REQUEST_EMAILID = Add_Delete_Drivers_Row_i[11]
                   
                    #date difference for broker id = 1
                    #2021-03-18 YYYY-MM-DD
                    cursor.execute("{CALL dbo.USP_BROKER_CONFIG(?)}",BROKER_ID)
                    BROKER_ID_Config = cursor.fetchall()
                    if len(BROKER_ID_Config)>=1:
                        Flag_ADD_DELETE_PRE_SETTIGS = BROKER_ID_Config[0][11]
                        BROKERid = BROKER_ID_Config[0][0]
                        ADD_DELETE_PRE_SETTIGS = BROKER_ID_Config[0][3]
                        if Flag_ADD_DELETE_PRE_SETTIGS is None:
                            Flag_ADD_DELETE_PRE_SETTIGS = 0
                        
                        if ADD_DELETE_PRE_SETTIGS is None:
                            ADD_DELETE_PRE_SETTIGS = 0
                        else:
                            ADD_DELETE_PRE_SETTIGS =0
                            BROKERid = 500000                                                
                    else:
                        ADD_DELETE_PRE_SETTIGS =0
                        BROKERid = 500000
                    
                    
                    
                    f_date = TERMINATION_DATE[8:10]
                    l_date = effday[8:10]
                    if int(f_date) < int(l_date):
                        delta = int(l_date) - int(f_date)
                    else:
                         delta = int(f_date) - int(l_date)
                         
                    delta = int(f_date) - int(l_date) 
                    print(delta)
                    '''
                    
                    from datetime import datetime
                    import pandas as pd
                    from datetime import date
                    '''

                    #############

                    
                    if ADD_REQUEST_EMAILID == InEmailID:
            
                        try:
                            #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                            DRIVER_DOB_Excel = Mask_DOB(DRIVER_DOB)
                        except:
                            DRIVER_DOB_Excel = ''
                        try:   
                            DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                        except:
                            DRIVER_EFF_DATE_Excel = ''
        
                        try:
                            TERMINATION_DATE_Excel = datetime.strptime(TERMINATION_DATE,'%Y-%m-%d')
                        except:
                            TERMINATION_DATE_Excel = ''  
                             
                        
                        DRIVER_STATUS = 'ADD/DELETE'
                        
                        if DRIVER_EFF_DATE == TERMINATION_DATE:
                            
                            No_of_Drivers = 0
                            OARATE_Final = 0
                            CLRATE_Final = 0
                            DUESRATE_Final = 0
                            PA_RATE_Final = 0
                                
                        elif  (EffDate_MM_YY[0:2] != TERMINATION_DATE[5:7]) or  (EffDate_MM_YY[0:4] != TERMINATION_DATE[0:4]):
                            print('here')
                            No_of_Drivers = 1
                            OARATE_Final = float(OARATE)
                            CLRATE_Final = float(CLRATE)
                            DUESRATE_Final = float(DUESRATE) 
                            PA_RATE_Final = float(PA_RATE)
                            if IS_PA_RATE:
                                PA_RATE_Final = float(PA_RATE)
                            else:
                                PA_RATE_Final = 0
        
                        elif (EffDate_MM_YY[0:2] == TERMINATION_DATE[5:7] and int(effday[8:11]) <= 15) and int(TERMINATION_DATE[5:7]) ==int(MonthOfReport_MM_yy[0:2]):
                            No_of_Drivers = 1
                            OARATE_Final = float(OARATE)
                            CLRATE_Final = float(CLRATE)
                            DUESRATE_Final = float(DUESRATE) 
                            PA_RATE_Final = float(PA_RATE)
                            if IS_PA_RATE:
                                PA_RATE_Final = float(PA_RATE)
                            else:
                                PA_RATE_Final = 0
                                               
                        
                        else:    
                            No_of_Drivers = 0
                            OARATE_Final = 0
                            CLRATE_Final = 0
                            DUESRATE_Final = 0
                            PA_RATE_Final = 0
                            
                            
                        #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            print(Total_Rate)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'pink'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'pink','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                            
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                           
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                    
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Excel,cell_format) # MONTH


                        flag = 'No'
                        if ((EffDate_MM_YY == MonthOfReport_MM_yy_Minus_1) and (EffDayInt<16)) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_2) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_3) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_4) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_5) or (EffDate_MM_YY == MonthOfReport_MM_yy_Minus_6):
                            print('DB')
                            flag = 'Yes'
                            Row = str(int(Row) + 1)
                            DEBITCOUNT = DEBITCOUNT + 1
                            # print('execyted this')
                            DRIVER_STATUS = 'DB'
                            No_of_Drivers = 1
                            OARATE_Final = float(OARATE)
                            CLRATE_Final = float(CLRATE)
                            DUESRATE_Final = float(DUESRATE)
                            if IS_PA_RATE:
                                PA_RATE_Final = float(PA_RATE)
                            else:
                                PA_RATE_Final = 0
                                
                                              #PA_RATE_Final fetch from database
                            if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_CL_RATE==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            else:
                                Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink'})
                            cell_format.set_font_name('Arial')
                        
                            worksheet.write('A'+Row, FirstName,cell_format) # First Name
                            worksheet.write('B'+Row, LastName,cell_format) # Last Name
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            
                            worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                            worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                            worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                            worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                            TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                            worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mm/dd/yy')
                            worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                            worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                            # worksheet.write('F'+Row, '',cell_format) # TERMDATE
                            worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                           
                            
                            
                            cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'pink','align': 'right'})
                            cell_format.set_font_name('Arial')
                            worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                            worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                                                    
                            if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                                worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                                worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_CL_RATE == 1:
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_PA_RATE == 1:
                                worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            else:
                                worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                    
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mmm-yy')
                            worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                            
                        
                        import datetime
                        end_date = datetime.datetime(int(TERMINATION_DATE[0:4]),int(TERMINATION_DATE[5:7]),int(TERMINATION_DATE[8:10]))
                        start_date = datetime.datetime(int(effday[0:4]),int(effday[5:7]),int(effday[8:10]))
                        num_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
                        if (num_months >6):
                            Row = str(int(Row) + 1)
                            print('DB')
                            DEBITCOUNT = DEBITCOUNT + 1
                            DRIVER_STATUS = 'DB'
                            No_of_Drivers = 1
                            OARATE_Final = float(OARATE)
                            CLRATE_Final = float(CLRATE)
                            DUESRATE_Final = float(DUESRATE)
                            if IS_PA_RATE:
                                PA_RATE_Final = float(PA_RATE)
                            else:
                                PA_RATE_Final = 0
                            
                            #PA_RATE_Final fetch from database
                            if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_CL_RATE==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            else:
                                Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink'})
                            cell_format.set_font_name('Arial')
                            worksheet.write('A'+Row, FirstName,cell_format) # First Name
                            worksheet.write('B'+Row, LastName,cell_format) # Last Name
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            
                            worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                            worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                            worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                            worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                            TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                            worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mm/dd/yy')
                            worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                            worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                            # worksheet.write('F'+Row, '',cell_format) # TERMDATE
                            worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                           
                            
                            
                            cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'pink','align': 'right'})
                            cell_format.set_font_name('Arial')
                            worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                            worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                            
                            if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                                worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                                worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_CL_RATE == 1:
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_PA_RATE == 1:
                                worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            else:
                                worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                                
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'pink','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mmm-yy')
                            worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                        
                        from datetime import datetime
                        if delta >=ADD_DELETE_PRE_SETTIGS and BROKER_ID == BROKERid and flag != "Yes" and int(Flag_ADD_DELETE_PRE_SETTIGS) ==1:
                            Row = str(int(Row) + 1)
                            print('DB')
                            DEBITCOUNT = DEBITCOUNT + 1
                            DRIVER_STATUS = 'DB'
                            No_of_Drivers = 1
                            OARATE_Final = float(OARATE)
                            CLRATE_Final = float(CLRATE)
                            DUESRATE_Final = float(DUESRATE)
                            if IS_PA_RATE:
                                PA_RATE_Final = float(PA_RATE)
                            else:
                                PA_RATE_Final = 0
                            
                                                      #PA_RATE_Final fetch from database
                            if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_CL_RATE==1:
                                Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                            
                            elif IS_PA_RATE ==1:
                                Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                            
                            else:
                                Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'blue'})
                            cell_format.set_font_name('Arial')
                            worksheet.write('A'+Row, FirstName,cell_format) # First Name
                            worksheet.write('B'+Row, LastName,cell_format) # Last Name
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            
                            worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                            worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                            worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                            worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                            TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                            worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                            
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mm/dd/yy')
                            worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                            worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                            # worksheet.write('F'+Row, '',cell_format) # TERMDATE
                            worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                           
                            
                            
                            cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'blue','align': 'right'})
                            cell_format.set_font_name('Arial')
                            worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                            worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                            
                            if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                                worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                                worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_CL_RATE == 1:
                                worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            
                            elif IS_PA_RATE == 1:
                                worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                                worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                            else:
                                worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                                
                            cell_format = workbook.add_format({'font_size': '11','font_color': 'blue','align': 'centre'})
                            cell_format.set_font_name('Arial')
                            cell_format.set_num_format('mmm-yy')
                            worksheet.write('K'+Row, MonthOfReport_Minus_2_Excel,cell_format) # MONTH    
                      
                    
                
                # DELETE DRIVERS  ################################
                # Get Delete drivers
                from datetime import datetime
                cursor.execute("{CALL dbo.GetDeletedDriversFromClientID(?)}",ClientID)
                Delete_Drivers = cursor.fetchall()
                for i in range(0,len(Delete_Drivers),1):
                    second_flag = ''
                    first_flag = ''
                    DELETECOUNT = DELETECOUNT + 1
                    Row = str(int(Row) + 1)
                    Delete_Drivers_Row_i = Delete_Drivers[i]
                    #print(Delete_Drivers_Row_i)
                    FirstName = Delete_Drivers_Row_i[0]
                    LastName = Delete_Drivers_Row_i[1]
                    DRIVER_SSN = Delete_Drivers_Row_i[2]
                    DRIVER_SSN = Mask_SSN(DRIVER_SSN)
                    DRIVER_DOB = Delete_Drivers_Row_i[3]
                    DRIVER_EFF_DATE = Delete_Drivers_Row_i[4]
                    TERMINATION_DATE = Delete_Drivers_Row_i[5]
                    DRIVER_STATE = Delete_Drivers_Row_i[6]
                    RFI_RF = Delete_Drivers_Row_i[7]
                    RETRO_ACTIVE_DAYS_DELETE = Delete_Drivers_Row_i[10]
                    if RETRO_ACTIVE_DAYS_DELETE is None:
                        RETRO_ACTIVE_DAYS_DELETE = 0
                    # break
        
                    try:
                        Termination_Date_MM_YY = Delete_Drivers_Row_i[8]
                    except:
                        Termination_Date_MM_YY = ''
                    try:    
                        TerminationDayInt = int(Delete_Drivers_Row_i[9])
                    except:
                        TerminationDayInt = 0
                    
                    
                    #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                    #DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                    #TERMINATION_DATE_Excel = datetime.strptime(TERMINATION_DATE,'%Y-%m-%d')
        
                    try:
                        #DRIVER_DOB_Excel = datetime.strptime(DRIVER_DOB,'%Y-%m-%d')
                        DRIVER_DOB_Excel = Mask_DOB(DRIVER_DOB)
                    except:
                        DRIVER_DOB_Excel = ''
                    try:   
                        DRIVER_EFF_DATE_Excel = datetime.strptime(DRIVER_EFF_DATE,'%Y-%m-%d')
                    except:
                        DRIVER_EFF_DATE_Excel = ''
        
                    try:
                        TERMINATION_DATE_Excel = datetime.strptime(TERMINATION_DATE,'%Y-%m-%d')
                    except:
                        TERMINATION_DATE_Excel = '' 
                    
                    DRIVER_STATUS = 'DELETE'
                    No_of_Drivers = 0
                    
                    if (Termination_Date_MM_YY == MonthOfReport_MM_yy) and (TerminationDayInt>1):
                       No_of_Drivers = 1
                    if (Termination_Date_MM_YY == MonthOfReport_MM_yy) and (TerminationDayInt==1):
                       No_of_Drivers = 0
                    if No_of_Drivers == 1:
                       OARATE_Final = float(OARATE)
                       CLRATE_Final = float(CLRATE)
                       DUESRATE_Final = float(DUESRATE)
                       if IS_PA_RATE:
                           PA_RATE_Final = float(PA_RATE)
                       else:
                           PA_RATE_Final = 0
                            
                    else:
                       OARATE_Final = 0
                       CLRATE_Final = 0
                       DUESRATE_Final = 0 
                       PA_RATE_Final = 0
                       
                       #PA_RATE_Final fetch from database
                    if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                    
                    elif IS_CL_RATE==1:
                        Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                    
                    elif IS_PA_RATE ==1:
                        Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                    
                    else:
                        Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'red'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, FirstName,cell_format) # First Name
                    worksheet.write('B'+Row, LastName,cell_format) # Last Name
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    
                    worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                    worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                    worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                    worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                    TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                    worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mm/dd/yy')
                    worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                    worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                    worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                                                           
                    cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'red','align': 'right'})
                    cell_format.set_font_name('Arial')
                    worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                    worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
        
                    
                    if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                        worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                        worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_CL_RATE == 1:
                        worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    elif IS_PA_RATE == 1:
                        worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                        worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                    else:
                        worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                    cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                    cell_format.set_font_name('Arial')
                    cell_format.set_num_format('mmm-yy')
                    worksheet.write('K'+Row, MonthOfReport_Excel,cell_format) # MONTH
                   
                    # print((Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_1))
                    # print('start here')
                    # print(FirstName)
                    # print(Termination_Date_MM_YY)
                    
                    # print(MonthOfReport_MM_yy_Minus_2)
                    # print((Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_2))
                    
                    if ((Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_1) and (TerminationDayInt==1)) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_2) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_3) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_4) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_5) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_6):
                        first_flag = 'yes'
                        Row = str(int(Row) + 1)
                        CREDITCOUNT = CREDITCOUNT + 1
                        DRIVER_STATUS = 'CR'
                        No_of_Drivers = -1
                        OARATE_Final = float(OARATE)*No_of_Drivers
                        CLRATE_Final = float(CLRATE)*No_of_Drivers
                        DUESRATE_Final = float(DUESRATE)*No_of_Drivers
                        if IS_PA_RATE:
                           PA_RATE_Final = float(PA_RATE)*No_of_Drivers
                        else:
                           PA_RATE_Final = 0
                            
                                          #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                       
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'red','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                                   
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                                               
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                        
                    if ((Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_2 and US_Number_Of_Credit_Debit==2) and (TerminationDayInt==1)) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_3 and US_Number_Of_Credit_Debit == 2) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_4 and US_Number_Of_Credit_Debit == 2) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_5 and US_Number_Of_Credit_Debit == 2) or (Termination_Date_MM_YY == MonthOfReport_MM_yy_Minus_6 and US_Number_Of_Credit_Debit == 2):
                        second_flag = 'yes'
                        
                        print('This is run')
                        Row = str(int(Row) + 1)
                        CREDITCOUNT = CREDITCOUNT + 1
                        DRIVER_STATUS = 'CR'
                        No_of_Drivers = -1
                        OARATE_Final = float(OARATE)*No_of_Drivers
                        CLRATE_Final = float(CLRATE)*No_of_Drivers
                        DUESRATE_Final = float(DUESRATE)*No_of_Drivers
                        if IS_PA_RATE:
                           PA_RATE_Final = float(PA_RATE)*No_of_Drivers
                        else:
                           PA_RATE_Final = 0
                           
                        #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                       
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'red','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                        
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_2_Excel,cell_format) # MONTH    
                    
                    
                    ##
                     
                    if RETRO_ACTIVE_DAYS_DELETE == 1 and first_flag != 'yes':
        
                        Row = str(int(Row) + 1)
                        CREDITCOUNT = CREDITCOUNT + 1
                        DRIVER_STATUS = 'CR'
                        No_of_Drivers = -1
                        OARATE_Final = float(OARATE)*No_of_Drivers
                        CLRATE_Final = float(CLRATE)*No_of_Drivers
                        DUESRATE_Final = float(DUESRATE)*No_of_Drivers
                        if IS_PA_RATE:
                           PA_RATE_Final = float(PA_RATE)*No_of_Drivers
                        else:
                           PA_RATE_Final = 0
                            
                                          #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                       
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'red','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                                   
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                                               
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_1_Excel,cell_format) # MONTH
                    
                    if RETRO_ACTIVE_DAYS_DELETE == 1 and US_Number_Of_Credit_Debit ==2 and second_flag != 'yes':
                        Row = str(int(Row) + 1)
                        CREDITCOUNT = CREDITCOUNT + 1
                        DRIVER_STATUS = 'CR'
                        No_of_Drivers = -1
                        OARATE_Final = float(OARATE)*No_of_Drivers
                        CLRATE_Final = float(CLRATE)*No_of_Drivers
                        DUESRATE_Final = float(DUESRATE)*No_of_Drivers
                        if IS_PA_RATE:
                           PA_RATE_Final = float(PA_RATE)*No_of_Drivers
                        else:
                           PA_RATE_Final = 0
                           
                        #PA_RATE_Final fetch from database
                        if IS_CL_RATE == 1 and IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_CL_RATE==1:
                            Total_Rate = float(OARATE_Final)+float(CLRATE_Final)+float(DUESRATE_Final)
                        
                        elif IS_PA_RATE ==1:
                            Total_Rate = float(OARATE_Final)+float(PA_RATE_Final)+float(DUESRATE_Final)
                        
                        else:
                            Total_Rate = float(OARATE_Final)+float(DUESRATE_Final)
                       
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('A'+Row, FirstName,cell_format) # First Name
                        worksheet.write('B'+Row, LastName,cell_format) # Last Name
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        
                        worksheet.write('C'+Row, DRIVER_SSN,cell_format) # SSN
                        worksheet.write('G'+Row, DRIVER_STATE,cell_format) # STATE
                        worksheet.write('H'+Row, RFI_RF,cell_format) # RFI / RF
                        worksheet.write('I'+Row, No_of_Drivers,cell_format) # NO of DRIVERS
                        TOTAL_DRIVER_COUNT = TOTAL_DRIVER_COUNT + No_of_Drivers
                        worksheet.write('J'+Row, DRIVER_STATUS,cell_format) # STATUS
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mm/dd/yy')
                        worksheet.write('D'+Row, DRIVER_DOB_Excel,cell_format) # DOB
                        worksheet.write('E'+Row, DRIVER_EFF_DATE_Excel,cell_format) # EFFDATE
                        worksheet.write('F'+Row, TERMINATION_DATE_Excel,cell_format) # TERMDATE
                       
                        
                        
                        cell_format = workbook.add_format({'num_format': '#,##0.00','font_color': 'red','align': 'right'})
                        cell_format.set_font_name('Arial')
                        worksheet.write('L'+Row, Total_Rate,cell_format) # TOTALRATE
                        worksheet.write('M'+Row, OARATE_Final,cell_format) # OARATE
                        
                        if IS_CL_RATE ==1 and IS_PA_RATE == 1:
                               
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE
                            worksheet.write('O'+Row, PA_RATE_Final,cell_format) # CLRATE
                            worksheet.write('P'+Row, DUESRATE_Final,cell_format)  # DUES
                    
                        elif IS_CL_RATE == 1:
                            worksheet.write('N'+Row, CLRATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        elif IS_PA_RATE == 1:
                            worksheet.write('N'+Row, PA_RATE_Final,cell_format) # CLRATE                    
                            worksheet.write('O'+Row, DUESRATE_Final,cell_format)  # DUES
                        else:
                            worksheet.write('N'+Row, DUESRATE_Final,cell_format)  # DUES
                        
                        cell_format = workbook.add_format({'font_size': '11','font_color': 'red','align': 'centre'})
                        cell_format.set_font_name('Arial')
                        cell_format.set_num_format('mmm-yy')
                        worksheet.write('K'+Row, MonthOfReport_Minus_2_Excel,cell_format) # MONTH    
                
                if IS_CL_RATE ==1 and IS_PA_RATE !=1:
                    
                    DataRowLast = Row
                    Row = str(int(Row) + 2)
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, '',cell_format)
                    worksheet.write('B'+Row, '',cell_format)
                    worksheet.write('C'+Row, '',cell_format)
                    worksheet.write('D'+Row, '',cell_format)
                    worksheet.write('E'+Row, '',cell_format)
                    worksheet.write('F'+Row, '',cell_format)
                    worksheet.write('G'+Row, '',cell_format)
                    worksheet.write('H'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=sum(M"+DataRowFirst+":M"+DataRowLast+")",cell_format)
                    worksheet.write("N"+Row, "=sum(N"+DataRowFirst+":N"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=sum(I"+DataRowFirst+":I"+DataRowLast+")",cell_format) # SUM of Drivers
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_right(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("O"+Row, "=sum(O"+DataRowFirst+":O"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("G8", "=I"+Row,cell_format) # SUM of Drivers
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row = str(int(Row) + 2)
                    worksheet.write('H'+Row, 'TOTALS',cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('I'+Row, '#Drivers',cell_format)
                    worksheet.write('J'+Row, 'OA Rate',cell_format)
                    worksheet.write('K'+Row, 'CL Rate',cell_format)
                    worksheet.write('L'+Row, 'Dues',cell_format)
                    worksheet.write('M'+Row, 'OA GP',cell_format)
                    worksheet.write('N'+Row, 'CL GP',cell_format)
                    worksheet.write('O'+Row, 'Dues',cell_format)
                    Row = str(int(Row) + 1)
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')                
                    worksheet.write('H'+Row, 'Paying For',cell_format)
                    Row_Minus_3 = str(int(Row) - 3)
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=I"+Row_Minus_3,cell_format)
                    cell_format = workbook.add_format({'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('J'+Row, OARATE,cell_format)
                    worksheet.write('K'+Row, CLRATE,cell_format)
                    worksheet.write('L'+Row, DUESRATE,cell_format)
                    worksheet.write('M'+Row, "=I"+Row+"*"+"J"+Row,cell_format)
                    worksheet.write('N'+Row, "=I"+Row+"*"+"K"+Row,cell_format)
                    worksheet.write('O'+Row, "=I"+Row+"*"+"L"+Row,cell_format)
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('H'+Row, 'Gross Premium Due High Point Underwriters',cell_format)
                    Row_Minus_4 = str(int(Row) - 4)
                    worksheet.write('I'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=M"+Row_Minus_4,cell_format)
                    worksheet.write("N"+Row, "=N"+Row_Minus_4,cell_format)
                    worksheet.write("O"+Row, "=O"+Row_Minus_4,cell_format)
                    Row = str(int(Row) + 1)
                
                elif IS_CL_RATE !=1 and IS_PA_RATE ==1:
                    
                    DataRowLast = Row
                    Row = str(int(Row) + 2)
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, '',cell_format)
                    worksheet.write('B'+Row, '',cell_format)
                    worksheet.write('C'+Row, '',cell_format)
                    worksheet.write('D'+Row, '',cell_format)
                    worksheet.write('E'+Row, '',cell_format)
                    worksheet.write('F'+Row, '',cell_format)
                    worksheet.write('G'+Row, '',cell_format)
                    worksheet.write('H'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=sum(M"+DataRowFirst+":M"+DataRowLast+")",cell_format)
                    worksheet.write("N"+Row, "=sum(N"+DataRowFirst+":N"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=sum(I"+DataRowFirst+":I"+DataRowLast+")",cell_format) # SUM of Drivers
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_right(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("O"+Row, "=sum(O"+DataRowFirst+":O"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("G8", "=I"+Row,cell_format) # SUM of Drivers
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row = str(int(Row) + 2)
                    worksheet.write('H'+Row, 'TOTALS',cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('I'+Row, '#Drivers',cell_format)
                    worksheet.write('J'+Row, 'OA Rate',cell_format)
                    worksheet.write('K'+Row, 'PA Rate',cell_format)
                    worksheet.write('L'+Row, 'Dues',cell_format)
                    worksheet.write('M'+Row, 'OA GP',cell_format)
                    worksheet.write('N'+Row, 'PA GP',cell_format)
                    worksheet.write('O'+Row, 'Dues',cell_format)
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')                
                    worksheet.write('H'+Row, 'Paying For',cell_format)
                    Row_Minus_3 = str(int(Row) - 3)
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=I"+Row_Minus_3,cell_format)
                    cell_format = workbook.add_format({'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('J'+Row, OARATE,cell_format)
                    worksheet.write('K'+Row, PA_RATE,cell_format)
                    worksheet.write('L'+Row, DUESRATE,cell_format)
                    worksheet.write('M'+Row, "=I"+Row+"*"+"J"+Row,cell_format)
                    worksheet.write('N'+Row, "=I"+Row+"*"+"K"+Row,cell_format)
                    worksheet.write('O'+Row, "=I"+Row+"*"+"L"+Row,cell_format)
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('H'+Row, 'Gross Premium Due High Point Underwriters',cell_format)
                    Row_Minus_4 = str(int(Row) - 4)
                    worksheet.write('I'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=M"+Row_Minus_4,cell_format)
                    worksheet.write("N"+Row, "=N"+Row_Minus_4,cell_format)
                    worksheet.write("O"+Row, "=O"+Row_Minus_4,cell_format)
                    Row = str(int(Row) + 1)
                    
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:
                    
                    DataRowLast = Row
                    Row = str(int(Row) + 2)
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, '',cell_format)
                    worksheet.write('B'+Row, '',cell_format)
                    worksheet.write('C'+Row, '',cell_format)
                    worksheet.write('D'+Row, '',cell_format)
                    worksheet.write('E'+Row, '',cell_format)
                    worksheet.write('F'+Row, '',cell_format)
                    worksheet.write('G'+Row, '',cell_format)
                    worksheet.write('H'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    
                    worksheet.write("P"+Row, "=sum(P"+DataRowFirst+":P"+DataRowLast+")",cell_format)
                    worksheet.write("M"+Row, "=sum(M"+DataRowFirst+":M"+DataRowLast+")",cell_format)
                    worksheet.write("N"+Row, "=sum(N"+DataRowFirst+":N"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=sum(I"+DataRowFirst+":I"+DataRowLast+")",cell_format) # SUM of Drivers
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_right(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("O"+Row, "=sum(O"+DataRowFirst+":O"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("G8", "=I"+Row,cell_format) # SUM of Drivers
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row = str(int(Row) + 2)
                    worksheet.write('G'+Row, 'TOTALS',cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('H'+Row, '#Drivers',cell_format)
                    worksheet.write('I'+Row, 'OA Rate',cell_format)
                    worksheet.write('J'+Row, 'CL Rate',cell_format)
                    worksheet.write('K'+Row, 'PA Rate',cell_format)
                    worksheet.write('L'+Row, 'Dues',cell_format)
                    worksheet.write('M'+Row, 'OA GP',cell_format)
                    worksheet.write('N'+Row, 'CL GP',cell_format)
                    worksheet.write('O'+Row, 'PA GP',cell_format)
                    worksheet.write('P'+Row, 'Dues',cell_format)
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')                
                    worksheet.write('G'+Row, 'Paying For',cell_format)
                    Row_Minus_3 = str(int(Row) - 3)
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("H"+Row, "=I"+Row_Minus_3,cell_format)
                    cell_format = workbook.add_format({'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('I'+Row, OARATE,cell_format)
                    worksheet.write('J'+Row, CLRATE,cell_format)
                    worksheet.write('K'+Row, PA_RATE,cell_format)
                    worksheet.write('L'+Row, DUESRATE,cell_format)
                    worksheet.write('M'+Row, "=H"+Row+"*"+"I"+Row,cell_format)
                    worksheet.write('N'+Row, "=H"+Row+"*"+"J"+Row,cell_format)
                    worksheet.write('O'+Row, "=H"+Row+"*"+"K"+Row,cell_format)
                    worksheet.write('P'+Row, "=H"+Row+"*"+"L"+Row,cell_format)
        
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('G'+Row, 'Gross Premium Due High Point Underwriters',cell_format)
                    Row_Minus_4 = str(int(Row) - 4)
                    worksheet.write('I'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=M"+Row_Minus_4,cell_format)
                    worksheet.write("N"+Row, "=N"+Row_Minus_4,cell_format)
                    worksheet.write("O"+Row, "=O"+Row_Minus_4,cell_format)
                    worksheet.write("P"+Row, "=P"+Row_Minus_4,cell_format)
        
                    Row = str(int(Row) + 1)
                    
                else:                        
                    DataRowLast = Row
                    Row = str(int(Row) + 2)
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('A'+Row, '',cell_format)
                    worksheet.write('B'+Row, '',cell_format)
                    worksheet.write('C'+Row, '',cell_format)
                    worksheet.write('D'+Row, '',cell_format)
                    worksheet.write('E'+Row, '',cell_format)
                    worksheet.write('F'+Row, '',cell_format)
                    worksheet.write('G'+Row, '',cell_format)
                    worksheet.write('H'+Row, '',cell_format)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    worksheet.write("M"+Row, "=sum(M"+DataRowFirst+":M"+DataRowLast+")",cell_format)
                    worksheet.write("N"+Row, "=sum(N"+DataRowFirst+":N"+DataRowLast+")",cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("I"+Row, "=sum(I"+DataRowFirst+":I"+DataRowLast+")",cell_format) # SUM of Drivers
                    cell_format = workbook.add_format({'num_format': '#,##0.00','bold': True,'font_size': '11'})
                    cell_format.set_top(1)
                    cell_format.set_bottom(1)
                    cell_format.set_right(1)
                    cell_format.set_font_name('Arial')
                    # worksheet.write("O"+Row, "=sum(O"+DataRowFirst+":O"+DataRowLast+")",cell_format)
                    # cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write("G8", "=I"+Row,cell_format) # SUM of Drivers
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row = str(int(Row) + 2)
                    
                    
                    
                    worksheet.write('I'+Row, 'TOTALS',cell_format)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('J'+Row, '#Drivers',cell_format)
                    worksheet.write('K'+Row, 'OA Rate',cell_format)
                    #worksheet.write('K'+Row, 'CL Rate',cell_format)
                    worksheet.write('L'+Row, 'Dues',cell_format)
                    worksheet.write('M'+Row, 'OA GP',cell_format)
                    # worksheet.write('N'+Row, 'CL GP',cell_format)
                    worksheet.write('N'+Row, 'Dues',cell_format)
                    Row = str(int(Row) + 1)
                    
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')  
        
                      
                    worksheet.write('I'+Row, 'Paying For',cell_format)
                    Row_Minus_3 = str(int(Row) - 3)
                    cell_format = workbook.add_format({'font_size': '11','align': 'centre'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    
                    worksheet.write("J"+Row, "=I"+Row_Minus_3,cell_format)
                    cell_format = workbook.add_format({'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet.write('K'+Row, OARATE,cell_format)
                    #worksheet.write('K'+Row, CLRATE,cell_format)
                    worksheet.write('L'+Row, DUESRATE,cell_format)
                    # worksheet.write('L'+Row, "=I"+Row+"*"+"J"+Row,cell_format)
                    worksheet.write('M'+Row, "=J"+Row+"*"+"K"+Row,cell_format)
                    worksheet.write('N'+Row, "=J"+Row+"*"+"L"+Row,cell_format)
                    Row = str(int(Row) + 1)
                    cell_format = workbook.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    
                    
                    worksheet.write('I'+Row, 'Gross Premium Due High Point Underwriters',cell_format)
                    Row_Minus_4 = str(int(Row) - 4)
                    worksheet.write('J'+Row, '',cell_format)
                    worksheet.write('K'+Row, '',cell_format)
                    #worksheet.write('K'+Row, '',cell_format)
                    worksheet.write('L'+Row, '',cell_format)
                    # worksheet.write("L"+Row, "=L"+Row_Minus_4,cell_format)
                    worksheet.write("M"+Row, "=M"+Row_Minus_4,cell_format)
                    worksheet.write("N"+Row, "=N"+Row_Minus_4,cell_format)
                    Row = str(int(Row) + 1)
                    
                    
                return workbook,worksheet,Row,ADDCOUNT,DELETECOUNT,ADD_DELETECOUNT,DEBITCOUNT,CREDITCOUNT,TOTAL_DRIVER_COUNT
            
            def CalculationsFunBroker(workbook_Broker,worksheet_Broker,Row_Broker):
                
                if IS_CL_RATE == 1 and IS_PA_RATE !=1:                        
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'red'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('H'+Row_Broker, 'Less Commission',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','align': 'centre','font_color': 'red','num_format': '0%'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('J'+Row_Broker,OA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('K'+Row_Broker,CL_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('L'+Row_Broker,'0%' ,cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_1+"*J"+Row_Broker,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_1+"*K"+Row_Broker,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_1+"*L"+Row_Broker,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('H'+Row_Broker, 'Net Premiums To High Point Underwriters',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('K'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('L'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')      
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    Row_Broker_Minus_2 = str(int(Row_Broker) - 2)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_2+"-M"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_2+"-N"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_2+"-O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    
                    if Lumpsum_Flag =='Yes':
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Premium Due'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Annual = 'Annual CL Flat Rate Fee'  
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Annual, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker))
                        worksheet_Broker.write("J"+Row_Broker,LUMPSUM_CL_AMOUNT,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        Total_Payable_to_HPU = 'Total Payable to HPU'
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Total Payable to HPU'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_2 = str(int(Row_Broker)-2)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write('J'+Row_Broker, "=J"+Row_Broker_Minus_2+"+J"+Row_Broker_Minus_1,cell_format)
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                    else:                            
                        print('no')
                        # Row_Broker = str(int(Row_Broker) + 1)
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        worksheet_Broker.write('H'+Row_Broker, 'Total Payable to HPU',cell_format)
                        worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                
                elif IS_CL_RATE != 1 and IS_PA_RATE ==1:                        
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'red'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('H'+Row_Broker, 'Less Commission',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','align': 'centre','font_color': 'red','num_format': '0%'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('J'+Row_Broker,OA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('K'+Row_Broker,PA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('L'+Row_Broker,'0%' ,cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_1+"*J"+Row_Broker,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_1+"*K"+Row_Broker,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_1+"*L"+Row_Broker,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('H'+Row_Broker, 'Net Premiums To High Point Underwriters',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('K'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('L'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')      
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    Row_Broker_Minus_2 = str(int(Row_Broker) - 2)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_2+"-M"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_2+"-N"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_2+"-O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    
                    if Lumpsum_Flag =='Yes':
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Premium Due'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Annual = 'Annual CL Flat Rate Fee'  
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Annual, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker))
                        worksheet_Broker.write("J"+Row_Broker,LUMPSUM_CL_AMOUNT,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        Total_Payable_to_HPU = 'Total Payable to HPU'
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Due to HPU'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_2 = str(int(Row_Broker)-2)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write('J'+Row_Broker, "=J"+Row_Broker_Minus_2+"+J"+Row_Broker_Minus_1,cell_format)
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                    else:                            
                        print('no')
                        # Row_Broker = str(int(Row_Broker) + 1)
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        worksheet_Broker.write('G'+Row_Broker, 'Total Payable to HPU',cell_format)
                        worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:
                    print('her')
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'red'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('G'+Row_Broker, 'Less Commission',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    
                    
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','align': 'centre','font_color': 'red','num_format': '0%'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('I'+Row_Broker,OA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('J'+Row_Broker,CL_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('K'+Row_Broker,PA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('L'+Row_Broker,'0%' ,cell_format)
                    
                    
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_1+"*I"+Row_Broker,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_1+"*J"+Row_Broker,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_1+"*K"+Row_Broker,cell_format)
                    worksheet_Broker.write('P'+Row_Broker, "=P"+Row_Broker_Minus_1+"*L"+Row_Broker,cell_format)
        
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('G'+Row_Broker, 'Net Premiums To High Point Underwriters',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('K'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('L'+Row_Broker, '',cell_format)
                    
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')      
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    Row_Broker_Minus_2 = str(int(Row_Broker) - 2)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_2+"-M"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_2+"-N"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_2+"-O"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('P'+Row_Broker, "=P"+Row_Broker_Minus_2+"-P"+Row_Broker_Minus_1,cell_format)
                    
        
                     ############add here code for lumpsum
        
                    if Lumpsum_Flag =='Yes':
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Premium Due'
                        worksheet_Broker.merge_range('G'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker))
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Annual = 'Annual CL Flat Rate Fee'  
                        worksheet_Broker.merge_range('G'+Row_Broker+':I'+Row_Broker,Annual, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker))
                        worksheet_Broker.write("J"+Row_Broker,LUMPSUM_CL_AMOUNT,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        Total_Payable_to_HPU = 'Total Payable to HPU'
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Due to HPU'
                        worksheet_Broker.merge_range('G'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_2 = str(int(Row_Broker)-2)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write('J'+Row_Broker, "=J"+Row_Broker_Minus_2+"+J"+Row_Broker_Minus_1,cell_format)
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                    else:                            
                        Row_Broker = str(int(Row_Broker) + 1)
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        worksheet_Broker.write('G'+Row_Broker, 'Total Payable to HPU',cell_format)
                        worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                else:
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'red'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('I'+Row_Broker, 'Less Commission',cell_format)
                    worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','align': 'centre','font_color': 'red','num_format': '0%'})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('K'+Row_Broker,OA_COMMISSION_RATE ,cell_format)
                    # worksheet_Broker.write('K'+Row_Broker,PA_COMMISSION_RATE ,cell_format)
                    worksheet_Broker.write('L'+Row_Broker,'0%' ,cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_1+"*K"+Row_Broker,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_1+"*L"+Row_Broker,cell_format)
                    # worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_1+"*L"+Row_Broker,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    worksheet_Broker.write('I'+Row_Broker, 'Net Premiums To High Point Underwriters',cell_format)
                    worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('K'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('L'+Row_Broker, '',cell_format)
                    worksheet_Broker.write('M'+Row_Broker, '',cell_format)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')      
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    Row_Broker_Minus_2 = str(int(Row_Broker) - 2)
                    worksheet_Broker.write('M'+Row_Broker, "=M"+Row_Broker_Minus_2+"-M"+Row_Broker_Minus_1,cell_format)
                    worksheet_Broker.write('N'+Row_Broker, "=N"+Row_Broker_Minus_2+"-N"+Row_Broker_Minus_1,cell_format)
                    # worksheet_Broker.write('O'+Row_Broker, "=O"+Row_Broker_Minus_2+"-O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 1)
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
         
                    
                    if Lumpsum_Flag =='Yes':
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Premium Due'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1+"+P"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Annual = 'Annual CL Flat Rate Fee'  
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Annual, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker))
                        worksheet_Broker.write("J"+Row_Broker,LUMPSUM_CL_AMOUNT,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                        Total_Payable_to_HPU = 'Total Payable to HPU'
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        Net_Due_to_HPU = 'Net Due to HPU'
                        worksheet_Broker.merge_range('H'+Row_Broker+':I'+Row_Broker,Net_Due_to_HPU, cell_format)
                        worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_2 = str(int(Row_Broker)-2)
                        Row_Broker_Minus_1 = str(int(Row_Broker)-1)
                        worksheet_Broker.write('J'+Row_Broker, "=J"+Row_Broker_Minus_2+"+J"+Row_Broker_Minus_1,cell_format)
                        Row_Broker = str(int(Row_Broker) + 1)
                        
                    else:                            
                        # Row_Broker = str(int(Row_Broker) + 1)
                        print('no')
                        cell_format.set_font_name('Arial')
                        cell_format.set_bg_color('#FFFFA8')
                        worksheet_Broker.write('I'+Row_Broker, 'Total Payable to HPU',cell_format)
                        # worksheet_Broker.write('J'+Row_Broker, '',cell_format)
                        Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                        worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1,cell_format)                    
                        Row_Broker = str(int(Row_Broker) + 1)
                    
                return workbook_Broker,worksheet_Broker,Row_Broker
            
            def CalculationsInternal(workbook_Broker,worksheet_Broker,Row_Broker):
                if IS_PA_RATE:
                    PA_RATE_Final = float(PA_RATE)
                else:
                    PA_RATE_Final = 0
                
                if IS_CL_RATE == 1 and IS_PA_RATE !=1:  
                    ############add here code for lumpsum
        
                      
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    cell_format.set_bg_color('#FFFFA8')
                    
                    worksheet_Broker.write('H'+Row_Broker, 'Total Payable to HPU',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 2)
                    
                elif IS_CL_RATE != 1 and IS_PA_RATE ==1:    
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    cell_format.set_bg_color('#FFFFA8')
                    worksheet_Broker.write('H'+Row_Broker, 'Total Payable to HPU',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 2)
                    
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:                        
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    cell_format.set_bg_color('#FFFFA8')
                    worksheet_Broker.write('G'+Row_Broker, 'Total Payable to HPU',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 2)
                
                else:
                    cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                    cell_format.set_border(1)
                    cell_format.set_font_name('Arial')
                    cell_format.set_bg_color('#FFFFA8')
                    worksheet_Broker.write('H'+Row_Broker, 'Total Payable to HPU',cell_format)
                    worksheet_Broker.write('I'+Row_Broker, '',cell_format)
                    Row_Broker_Minus_1 = str(int(Row_Broker) - 1)
                    worksheet_Broker.write("J"+Row_Broker, "=M"+Row_Broker_Minus_1+"+N"+Row_Broker_Minus_1+"+O"+Row_Broker_Minus_1,cell_format)
                    Row_Broker = str(int(Row_Broker) + 2)
                
                merge_format = workbook_Broker.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'centre',
                        'bg_color':'yellow',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                
                PayDueDateExcel = 'PAYMENT DUE DATE: '+  PAY_DUE_DATE  
                
                if IS_CL_RATE == 1 and IS_PA_RATE !=1:
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                elif IS_CL_RATE != 1 and IS_PA_RATE ==1:
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:
                    worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                else:
                    worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
        
        
        
                Row_Broker = str(int(Row_Broker) + 1)
                merge_format = workbook_Broker.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'centre',
                        'bg_color':'#FFCCF9',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                
                PayDueDateExcel = 'LATE FEE IF RECEIVED AFTER: '+  PAY_DUE_DATE  
                if IS_CL_RATE == 1 and IS_PA_RATE !=1:
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                
                elif IS_CL_RATE != 1 and IS_PA_RATE ==1:
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:
                    worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                else:
                    worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                
                
        
        
                cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                cell_format.set_bg_color('#FFCCF9')
                
                
                if IS_CL_RATE == 1 and IS_PA_RATE !=1:
                    worksheet_Broker.write('M'+Row_Broker, LATE_FEE_AMOUNT,cell_format)
                                
                elif IS_CL_RATE == 1 and IS_PA_RATE !=1:
                    worksheet_Broker.write('M'+Row_Broker, LATE_FEE_AMOUNT,cell_format)
                elif IS_CL_RATE == 1 and IS_PA_RATE ==1:                        
                    worksheet_Broker.write('M'+Row_Broker, LATE_FEE_AMOUNT,cell_format)
                else:
                    worksheet_Broker.write('M'+Row_Broker, LATE_FEE_AMOUNT,cell_format)
                    
                Row_Broker = str(int(Row_Broker) + 1)
                merge_format = workbook_Broker.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'centre',
                        'bg_color':'#FFCCF9',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                
                PayDueDateExcel = 'REINSTATEMENT FEE + LATE FEE IF RECEIVED AFTER: '+  REINSTATEMENT_DATE 
                
                
                if IS_CL_RATE ==1 and IS_PA_RATE !=1:                
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            
                elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                    worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
        
                elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                    worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                else:
                    worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            
                
                cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                cell_format.set_bg_color('#FFCCF9')
                worksheet_Broker.write('M'+Row_Broker, REINSTATEMENT_AMOUNT,cell_format)
                Row_Broker = str(int(Row_Broker) + 1)
                return workbook_Broker,worksheet_Broker,Row_Broker
                
            ############# BROKER FILE #####################################
        #        BROKER_FILE_NAME = NAMED_MOTORCARRRIER_PLANE+" _Census Report_" + DateOfReport_Month_Year+"_To_Broker.xlsx"
        #        Broker_File_Path = CENSUS_FOLDER_PATH + "\\" + BROKER_FILE_NAME
        #        try:
        #            if os.path.exists(Broker_File_Path):
        #              os.remove(Broker_File_Path,ignore_errors=True)
        #        except:
        #            pass      
        #        
        #        workbook_Broker = xlsxwriter.Workbook(Broker_File_Path)
        #        worksheet_Broker = workbook_Broker.add_worksheet()
            Broker_File_Path,workbook_Broker,worksheet_Broker,BROKER_FILE_NAME = FileNameBroker()
            
            workbook_Broker,worksheet_Broker,Row_Broker,ADDCOUNT_Final,DELETECOUNT_Final,ADD_DELETECOUNT_Final,DEBITCOUNT_Final,CREDITCOUNT_Final,TOTAL_DRIVER_COUNT = FieldCopyFun(workbook_Broker,worksheet_Broker)
            
            workbook_Broker,worksheet_Broker,Row_Broker = CalculationsFunBroker(workbook_Broker,worksheet_Broker,Row_Broker)
            Row_Broker = str(int(Row_Broker) + 1)
            merge_format = workbook_Broker.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'centre',
                        'bg_color':'yellow',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                
            PayDueDateExcel = 'PAYMENT DUE DATE: '+  PAY_DUE_DATE  
            if IS_CL_RATE ==1 and IS_PA_RATE !=1:                
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            
            elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                
            
            elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            else:
                worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            
                
            Row_Broker = str(int(Row_Broker) + 1)
            merge_format = workbook_Broker.add_format({
                    'bold': 1,
                    'border': 1,
                    'align': 'centre',
                    'bg_color':'#FFCCF9',
                    'font_name':'Arial',
                    'font_color':'blue',
                    'font_size': '11'})
            PayDueDateExcel = 'LATE FEE IF RECEIVED AFTER: '+  PAY_DUE_DATE  
            if IS_CL_RATE ==1 and IS_PA_RATE !=1: 
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            else:
                worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
        
            cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
            cell_format.set_border(1)
            cell_format.set_font_name('Arial')
            cell_format.set_bg_color('#FFCCF9')
            worksheet_Broker.write('M'+Row_Broker, LATE_FEE_AMOUNT,cell_format)
            Row_Broker = str(int(Row_Broker) + 1)
            merge_format = workbook_Broker.add_format({
                    'bold': 1,
                    'border': 1,
                    'align': 'centre',
                    'bg_color':'#FFCCF9',
                    'font_name':'Arial',
                    'font_color':'blue',
                    'font_size': '11'})
            PayDueDateExcel = 'REINSTATEMENT FEE + LATE FEE IF RECEIVED AFTER: '+  REINSTATEMENT_DATE  
            if IS_CL_RATE ==1 and IS_PA_RATE !=1: 
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                worksheet_Broker.merge_range('H'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                worksheet_Broker.merge_range('G'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
            else:
                worksheet_Broker.merge_range('I'+Row_Broker+':L'+Row_Broker,PayDueDateExcel, merge_format)
                
            cell_format = workbook_Broker.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
            cell_format.set_border(1)
            cell_format.set_font_name('Arial')
            cell_format.set_bg_color('#FFCCF9')
            worksheet_Broker.write('M'+Row_Broker, REINSTATEMENT_AMOUNT,cell_format)
            workbook_Broker.close()
            GROSS_OA = TOTAL_DRIVER_COUNT*OARATE
            GROSS_CL = TOTAL_DRIVER_COUNT*CLRATE
            GROSS_PA = TOTAL_DRIVER_COUNT*PA_RATE
            GROSS_DUES = TOTAL_DRIVER_COUNT*DUESRATE
            
            OA_Commission_Val =  float(OA_COMMISSION_RATE.replace("%",""))*0.01*float(GROSS_OA)
            NET_OA = GROSS_OA - OA_Commission_Val
            
            CL_Commission_Val =  float(CL_COMMISSION_RATE.replace("%",""))*0.01*float(GROSS_CL)
            NET_CL = GROSS_CL - CL_Commission_Val
            
            NET_TOTAL = NET_OA + NET_CL + float(GROSS_DUES)
            #print(MonthOfReport_MMM_YY)
            args = (ClientID,MonthOfReport_MMM_YY,ADDCOUNT_Final,DELETECOUNT_Final,ADD_DELETECOUNT_Final,CREDITCOUNT_Final,DEBITCOUNT_Final,
                    TOTAL_DRIVER_COUNT,OARATE,CLRATE,DUESRATE,OA_COMMISSION_RATE,GROSS_OA,GROSS_CL,GROSS_DUES,OA_Commission_Val,NET_OA,NET_CL,GROSS_DUES,NET_TOTAL,
                    LATE_FEE_AMOUNT,REINSTATEMENT_AMOUNT,NET_TOTAL,InEmailID,CENSUS_FOLDER_PATH,CENSUS_CUT_OFF_DATE,PA_RATE,GROSS_PA,LUMPSUM_CL_AMOUNT)
            #print(args)
            cursor.execute("{CALL dbo.INSERT_INTO_CLIENT_CENSUS_DETAILS(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}",args)
            db.commit()
            args = (ClientID,'BROKER',MonthOfReport_MMM_YY,InEmailID,BROKER_FILE_NAME)
            cursor.execute("{CALL dbo.INSERT_INTO_SENT_EMAIL_LOG(?,?,?,?,?)}",args)
            db.commit()
            ############# INTERNAL FILE #####################################
            if int(INTERNAL_INVOICE) == 1:
                if IS_PA_RATE:
                    PA_RATE_Final = float(PA_RATE)
                else:
                    PA_RATE_Final = 0
        #            INTERNAL_FILE_NAME = NAMED_MOTORCARRRIER_PLANE+" _Census Report_" + DateOfReport_Month_Year+"_Internal.xlsx"
        #            Internal_File_Path = CENSUS_FOLDER_PATH + "\\" + INTERNAL_FILE_NAME
        #            try:
        #                if os.path.exists(Internal_File_Path):
        #                  os.remove(Internal_File_Path,ignore_errors=True)
        #            except:
        #                pass      
        #            workbook_Internal = xlsxwriter.Workbook(Internal_File_Path)
        #            worksheet_Internal = workbook_Internal.add_worksheet()
                
                Internal_File_Path,workbook_Internal,worksheet_Internal,INTERNAL_FILE_NAME = FileNameInternal()
                
                workbook_Internal,worksheet_Internal,Row_Internal,ADDCOUNT_Final,DELETECOUNT_Final,ADD_DELETECOUNT_Final,DEBITCOUNT_Final,CREDITCOUNT_Final,TOTAL_DRIVER_COUNT = FieldCopyFun(workbook_Internal,worksheet_Internal)
                workbook_Internal,worksheet_Internal,Row_Internal = CalculationsFunBroker(workbook_Internal,worksheet_Internal,Row_Internal)
                Row_Internal = str(int(Row_Internal) + 1)
                
                merge_format = workbook_Internal.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'left',
                        'bg_color':'#FFFFA8',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                if IS_CL_RATE ==1 and IS_PA_RATE !=1: 
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Paid(Check#)', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Paid(Check#)', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                    worksheet_Internal.merge_range('G'+Row_Internal+':I'+Row_Internal,'Paid(Check#)', merge_format)
                else:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Paid(Check#)', merge_format)
                               
                cell_format = workbook_Internal.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                cell_format.set_bg_color('#FFFFA8')
                worksheet_Internal.write('J'+Row_Internal, '',cell_format)
                Row_Internal = str(int(Row_Internal) + 1)
                
                merge_format = workbook_Internal.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'left',
                        'bg_color':'#FFFFA8',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                if IS_CL_RATE ==1 and IS_PA_RATE !=1: 
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Overpaid / underpaid', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Overpaid / underpaid', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                    worksheet_Internal.merge_range('G'+Row_Internal+':I'+Row_Internal,'Overpaid / underpaid', merge_format)
                else:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Overpaid / underpaid', merge_format)
                    
                cell_format = workbook_Internal.add_format({'bold': True,'font_size': '11','font_color': 'blue','num_format': 44})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                cell_format.set_bg_color('#FFFFA8')
                worksheet_Internal.write('J'+Row_Internal, '',cell_format)
                Row_Internal = str(int(Row_Internal) + 2)
                
                merge_format = workbook_Internal.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'left',
                        'bg_color':'#FFFFA8',
                        'font_name':'Arial',
                        'font_color':'blue',
                        'font_size': '11'})
                if IS_CL_RATE ==1 and IS_PA_RATE !=1: 
                   worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Date Deposited', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE !=1:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Date Deposited', merge_format)
                elif IS_CL_RATE ==1 and IS_PA_RATE ==1:
                    worksheet_Internal.merge_range('G'+Row_Internal+':I'+Row_Internal,'Date Deposited', merge_format)
                else:
                    worksheet_Internal.merge_range('H'+Row_Internal+':I'+Row_Internal,'Date Deposited', merge_format)
                    
                
                cell_format = workbook_Internal.add_format({'bold': True,'font_size': '11','font_color': 'blue'})
                cell_format.set_border(1)
                cell_format.set_font_name('Arial')
                cell_format.set_bg_color('#FFFFA8')
                cell_format.set_num_format('mm/dd/yy')
                worksheet_Internal.write('J'+Row_Internal, '',cell_format)
                workbook_Internal.close()
                
            ############# INTERNAL FILE #####################################
            if int(CLIENT_INVOICE) == 1:
                if IS_PA_RATE:
                    PA_RATE_Final = float(PA_RATE)
                else:
                    PA_RATE_Final = 0
           
                
                Client_File_Path,workbook_Client,worksheet_Client,CLIENT_FILE_NAME = FileNameClient()
                
                workbook_Client,worksheet_Client,Row_Client,ADDCOUNT_Final,DELETECOUNT_Final,ADD_DELETECOUNT_Final,DEBITCOUNT_Final,CREDITCOUNT_Final,TOTAL_DRIVER_COUNT = FieldCopyFun(workbook_Client,worksheet_Client)
                
                workbook_Client,worksheet_Client,Row_Client = CalculationsInternal(workbook_Client,worksheet_Client,Row_Client)
                
                workbook_Client.close()
                GROSS_OA = TOTAL_DRIVER_COUNT*OARATE
                GROSS_CL = TOTAL_DRIVER_COUNT*CLRATE
                GROSS_PA = TOTAL_DRIVER_COUNT*PA_RATE
                GROSS_DUES = TOTAL_DRIVER_COUNT*DUESRATE
                
                OA_Commission_Val =  0
                NET_OA = 0
                
                CL_Commission_Val =  0
                NET_CL = 0
                
                TOTAL_DUE_AMT = GROSS_OA + GROSS_CL + GROSS_DUES
                
                args = (ClientID,MonthOfReport_MMM_YY,ADDCOUNT_Final,DELETECOUNT_Final,ADD_DELETECOUNT_Final,CREDITCOUNT_Final,DEBITCOUNT_Final,
                    TOTAL_DRIVER_COUNT,OARATE,CLRATE,DUESRATE,'0%',GROSS_OA,GROSS_CL,GROSS_DUES,OA_Commission_Val,NET_OA,NET_CL,'0','0',
                    LATE_FEE_AMOUNT,REINSTATEMENT_AMOUNT,TOTAL_DUE_AMT,InEmailID,CENSUS_FOLDER_PATH,CENSUS_CUT_OFF_DATE,PA_RATE,GROSS_PA,LUMPSUM_CL_AMOUNT)
        
                cursor.execute("{CALL dbo.INSERT_INTO_CLIENT_CENSUS_DETAILS(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}",args)
                db.commit()
                args = (ClientID,'CLIENT',MonthOfReport_MMM_YY,InEmailID,CLIENT_FILE_NAME)
                cursor.execute("{CALL dbo.INSERT_INTO_SENT_EMAIL_LOG(?,?,?,?,?)}",args)
                db.commit()
                
            args = (ClientID,MonthOfReport_MMM_YY,CENSUS_CUT_OFF_DATE)
            cursor.execute("{CALL dbo.UPDATE_LAST_ACTIVITY_Date(?,?,?)}",args)
            db.commit()
    except:
        db.rollback()
#####END ####     

# def CensusModule():
        
#     import pyodbc 
#     import pandas as pd
    
#     conn_str = (r'DRIVER={SQL Server};'
#                r'SERVER=NDS-AA-02;'
#                r'DATABASE=HPU;'
#                r'Trusted_Connection=no;'
#                r'UID=RPA;'
#                r'PWD=nds1@2020;'
#                r'autocommit=True'
#            )
#     cnxn = pyodbc.connect(conn_str)
#     sqlExecSP="{call USP_GetEMailLogID}"
#     dfMailListToBeClassified = pd.read_sql_query(sql=sqlExecSP, con=cnxn)
      
#     EMAILID = int(dfMailListToBeClassified['ID'][0])
#     CLIENTID = dfMailListToBeClassified['CLIENTID'][0]
#     emailSender = dfMailListToBeClassified['FROM_EMAIL_ADDRESS'][0]
#     CLIENTID = int(CLIENTID)
#     IN_EMAILID = int(EMAILID)
    
#     # Get Client information
#     db = pyodbc.connect("Driver={SQL Server};"
#                      "Server=NDS-AA-02;"
#                      "Database=HPU;"
#                      "uid=RPA;pwd=nds1@2020;")
#     cursor = db.cursor()
#     cursor.execute("{CALL USP_GetRateInfo(?)}",int(CLIENTID))
#     Rate_info = cursor.fetchall()
    
#     PRORATE_FLAG = Rate_info[0][-1]
#     if PRORATE_FLAG:
#         from ProRateCensus import ProRateCensusGenerationFun
#         ProRateCensusGenerationFun()
#         #digital census
#         from ProRateModule import ProRate
#         ProRate()
#         from PDF_Invoice import PDF_INVOICE_MODULE
#         PDF_INVOICE_MODULE()
        
#     else:    
#         CensusGenerationFun()
#         from DigitalCensus import DigitalCensus1
#         DigitalCensus1()
#         from PDF_Invoice import PDF_INVOICE_MODULE
#         PDF_INVOICE_MODULE()
        
    
# CensusModule()
CensusGenerationFun()
    
    
    
    

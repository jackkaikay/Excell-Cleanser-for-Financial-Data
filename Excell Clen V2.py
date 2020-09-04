import pandas as pd
import os

workbook = 'M:\\2020\MiDES\\20M.058 Association of Colleges - Finance Projection Collection and Reporting\\6. Data Team\Data Outputs\Downloaded data\Problem Files\Financial Health\\FH Calculator v1.4a.xlsm'
workbook2 = 'M:\\2020\MiDES\\20M.058 Association of Colleges - Finance Projection Collection and Reporting\\6. Data Team\Data Outputs\Downloaded data\Problem Files\Monthly Cash Flows\\HBC July 2020 financial return - cashflow - submitted.xlsx'







#########WorkbookOnea#############


def sheetOne():
    sheet1 = pd.read_excel(workbook, sheet_name='1. Financial Health')
    sheet1.columns = ['blank', 'Key', 'Name', '2020', '2021', '2022']
    sheet1 = sheet1[['Key', 'Name', '2020', '2021', '2022']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[6:9]
    sheet2a = sheet2a[['Key','Name','2020','2021','2022']]

    sheet2a_added_info = ['Ratios','Ratios','Ratios']
    sheet2a_added_info2 = [workbook,workbook,workbook]
    sheet2a_added_info3 = [UKPRN,UKPRN,UKPRN]

    sheet2a['Category'] = sheet2a_added_info
    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3

    #sheet1 part 2
    sheet2b = sheet1.iloc[12:15]
    sheet2b = sheet2b[['Key','Name','2020','2021','2022']]

    sheet2b_added_info = ['Calculation of grade','Calculation of grade','Calculation of grade']
    sheet2b_added_info2 = [workbook,workbook,workbook]
    sheet2b_added_info3 = [UKPRN,UKPRN,UKPRN]

    sheet2b['Category'] = sheet2b_added_info
    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3

    # sheet1 part 3
    sheet2c = sheet1.iloc[16:17]
    sheet2c = sheet2c[['Key', 'Name', '2020', '2021', '2022']]

    sheet2c_added_info = ['Total Points']
    sheet2c_added_info2 = [workbook]
    sheet2c_added_info3 = [UKPRN]

    sheet2c['Category'] = sheet2c_added_info
    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3

    # sheet1 part 4
    sheet2d = sheet1.iloc[18:21]
    sheet2d = sheet2d[['Key', 'Name', '2020', '2021', '2022']]

    sheet2d_added_info = ['Financial Health Grade (automated)', 'Automated moderation - EFS', 'Automated moderation - EBITDA score zero']
    sheet2d_added_info2 = [workbook, workbook, workbook]
    sheet2d_added_info3 = [UKPRN, UKPRN, UKPRN]

    sheet2d['Category'] = sheet2d_added_info
    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3

    # sheet1 part 5
    sheet2e = sheet1.iloc[22:23]
    sheet2e = sheet2e[['Key', 'Name', '2020', '2021', '2022']]

    sheet2e_added_info = ['College Self Assessment']
    sheet2e_added_info2 = [workbook]
    sheet2e_added_info3 = [UKPRN]

    sheet2e['Category'] = sheet2e_added_info
    sheet2e['Excel Title'] = sheet2e_added_info2
    sheet2e['UKPRN'] = sheet2e_added_info3




    print(sheet1.to_string())
    print(sheet2a.to_string())
    print(sheet2b.to_string())
    print(sheet2c.to_string())
    print(sheet2d.to_string())
    print(sheet2e.to_string())

    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d,sheet2e])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Financial Health')

def sheetTwo():
    sheet1 = pd.read_excel(workbook, sheet_name='2. Ratios')
    sheet1.columns = ['blank', 'Key', 'Name', '2020', '2021', '2022']
    sheet1 = sheet1[['Key', 'Name', '2020', '2021', '2022']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[7:8]
    sheet2a = sheet2a[['Key','Name','2020','2021','2022']]

    sheet2a_added_info = ["Adjusted income used in ratio analysis (£'000')"]
    sheet2a_added_info2 = [workbook]
    sheet2a_added_info3 = [UKPRN]

    sheet2a['Category'] = sheet2a_added_info
    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3

    #sheet1 part 2
    sheet2b = sheet1.iloc[10:12]
    sheet2b = sheet2b[['Key','Name','2020','2021','2022']]

    sheet2b_added_info = ['Liquidity','Liquidity']
    sheet2b_added_info2 = [workbook,workbook]
    sheet2b_added_info3 = [UKPRN,UKPRN]

    sheet2b['Category'] = sheet2b_added_info
    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3

    #sheet1 part 3
    sheet2c = sheet1.iloc[14:15]
    sheet2c = sheet2c[['Key','Name','2020','2021','2022']]

    sheet2c_added_info = ['Gearing']
    sheet2c_added_info2 = [workbook]
    sheet2c_added_info3 = [UKPRN]

    sheet2c['Category'] = sheet2c_added_info
    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3

    #sheet1 part 4
    sheet2d = sheet1.iloc[17:25]
    sheet2d = sheet2d[['Key','Name','2020','2021','2022']]

    sheet2d_added_info = ['Margin','Margin','Margin','Margin','Margin','Margin','Margin','Margin']
    sheet2d_added_info2 = [workbook,workbook,workbook,workbook,workbook,workbook,workbook,workbook]
    sheet2d_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2d['Category'] = sheet2d_added_info
    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3


    print(sheet1.to_string())
    print(sheet2a.to_string())
    print(sheet2b.to_string())
    print(sheet2c.to_string())
    print(sheet2d.to_string())

    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Ratios')

def sheetThree():
    sheet1 = pd.read_excel(workbook, sheet_name='3. Income & Expenditure')
    sheet1.columns = ['blank', 'Key', 'Name','Misc', '2020', '2021', '2022']
    sheet1 = sheet1[['Name', '2020', '2021', '2022']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[8:14]
    sheet2a = sheet2a[['Name','2020','2021','2022']]

    sheet2a_added_info = ["Income" , "Income" , "Income" ,"Income" , "Income", "Income"]
    sheet2a_added_info2 = [workbook, workbook , workbook , workbook , workbook, workbook]
    sheet2a_added_info3 = [UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN]

    sheet2a['Category'] = sheet2a_added_info
    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3

    #sheet1 part 2
    sheet2b = sheet1.iloc[16:19]
    sheet2b = sheet2b[['Name','2020','2021','2022']]

    sheet2b_added_info = ["Expenditure" , "Expenditure" , "Expenditure" ]
    sheet2b_added_info2 = [workbook, workbook , workbook ]
    sheet2b_added_info3 = [UKPRN, UKPRN, UKPRN]

    sheet2b['Category'] = sheet2b_added_info
    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3

    #sheet1 part 3
    sheet2c = sheet1.iloc[20:21]
    sheet2c = sheet2c[['Name','2020','2021','2022']]

    sheet2c_added_info = ["Expenditure" ]
    sheet2c_added_info2 = [workbook]
    sheet2c_added_info3 = [UKPRN]

    sheet2c['Category'] = sheet2c_added_info
    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3


    #sheet1 part 4
    sheet2d = sheet1.iloc[23:27]
    sheet2d = sheet2d[['Name','2020','2021','2022']]

    sheet2d_added_info = ["Interest, taxation, depreciation & amortisation", "Interest, taxation, depreciation & amortisation" , "Interest, taxation, depreciation & amortisation" , "Interest, taxation, depreciation & amortisation" ]
    sheet2d_added_info2 = [workbook,workbook,workbook,workbook]
    sheet2d_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2d['Category'] = sheet2d_added_info
    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3

    #sheet1 part 5
    sheet2e = sheet1.iloc[28:29]
    sheet2e = sheet2e[['Name','2020','2021','2022']]

    sheet2e_added_info = ["Interest, taxation, depreciation & amortisation"]
    sheet2e_added_info2 = [workbook]
    sheet2e_added_info3 = [UKPRN]

    sheet2e['Category'] = sheet2e_added_info
    sheet2e['Excel Title'] = sheet2e_added_info2
    sheet2e['UKPRN'] = sheet2e_added_info3


    #sheet1 part 6
    sheet2f = sheet1.iloc[31:35]
    sheet2f = sheet2f[['Name','2020','2021','2022']]

    sheet2f_added_info = ["Other gains and losses", "Other gains and losses", "Other gains and losses", "Other gains and losses"]
    sheet2f_added_info2 = [workbook, workbook, workbook, workbook]
    sheet2f_added_info3 = [UKPRN, UKPRN, UKPRN, UKPRN]

    sheet2f['Category'] = sheet2f_added_info
    sheet2f['Excel Title'] = sheet2f_added_info2
    sheet2f['UKPRN'] = sheet2f_added_info3

    #sheet1 part 7
    sheet2g = sheet1.iloc[36:37]
    sheet2g = sheet2g[['Name','2020','2021','2022']]

    sheet2g_added_info = ["Other gains and losses"]
    sheet2g_added_info2 = [workbook]
    sheet2g_added_info3 = [UKPRN]

    sheet2g['Category'] = sheet2g_added_info
    sheet2g['Excel Title'] = sheet2g_added_info2
    sheet2g['UKPRN'] = sheet2g_added_info3


    #sheet1 part 8
    sheet2h = sheet1.iloc[39:47]
    sheet2h = sheet2h[['Name','2020','2021','2022']]

    sheet2h_added_info = ["Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios", "Adjustments to surplus/(deficit) for ratios"]
    sheet2h_added_info2 = [workbook, workbook, workbook, workbook, workbook, workbook, workbook, workbook]
    sheet2h_added_info3 = [UKPRN, UKPRN, UKPRN,UKPRN , UKPRN, UKPRN, UKPRN, UKPRN]

    sheet2h['Category'] = sheet2h_added_info
    sheet2h['Excel Title'] = sheet2h_added_info2
    sheet2h['UKPRN'] = sheet2h_added_info3

    #sheet1 part 9
    sheet2i = sheet1.iloc[52:56]
    sheet2i = sheet2i[['Name','2020','2021','2022']]

    sheet2i_added_info = ["Covid-19", "Covid-19", "Covid-19", "Covid-19"]
    sheet2i_added_info2 = [workbook, workbook, workbook, workbook]
    sheet2i_added_info3 = [UKPRN, UKPRN, UKPRN, UKPRN]

    sheet2i['Category'] = sheet2i_added_info
    sheet2i['Excel Title'] = sheet2i_added_info2
    sheet2i['UKPRN'] = sheet2i_added_info3

    print(sheet1.to_string())
    print(sheet2a.to_string())
    print(sheet2b.to_string())
    print(sheet2c.to_string())
    print(sheet2d.to_string())
    print(sheet2e.to_string())
    print(sheet2f.to_string())
    print(sheet2g.to_string())
    print(sheet2h.to_string())
    print(sheet2i.to_string())

    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d, sheet2e, sheet2f, sheet2g, sheet2h, sheet2i])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Income & Expenditure')

def sheetFour():
    #sheet1 = pd.read_excel(workbook, sheet_name='3a. Sensitivity Analysis')
    sheet1 = pd.read_excel(workbook, sheet_name='3a. I&E Sensitivities')
    sheet1.columns = ['Scenario', 'Code', 'Area of Sensitivity', 'Type', '2020', '2021', '2022','Financial Health Grade & Score Impact','2024','2025']
    sheet1 = sheet1[['Scenario', 'Code', 'Area of Sensitivity', 'Type','2020', '2021', '2022','Financial Health Grade & Score Impact','2024','2025']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[6:7]
    sheet2a = sheet2a[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2a_added_info2 = [workbook]
    sheet2a_added_info3 = [UKPRN]
    sheet2a_added_info4 = sheet1.loc[7]['Type']
    sheet2a_added_info5 = sheet1.loc[10]['Type']

    sheet2a['Description'] = sheet2a_added_info4
    sheet2a['Mitigation'] = sheet2a_added_info5

    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3



    #sheet1 part 2
    sheet2b = sheet1.iloc[12:13]
    sheet2b = sheet2b[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2b_added_info2 = [workbook]
    sheet2b_added_info3 = [UKPRN]
    sheet2b_added_info4 = sheet1.loc[13]['Type']
    sheet2b_added_info5 = sheet1.loc[16]['Type']

    sheet2b['Description'] = sheet2b_added_info4
    sheet2b['Mitigation'] = sheet2b_added_info5

    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3


    #sheet1 part 3
    sheet2c = sheet1.iloc[18:19]
    sheet2c = sheet2c[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2c_added_info2 = [workbook]
    sheet2c_added_info3 = [UKPRN]
    sheet2c_added_info4 = sheet1.loc[19]['Type']
    sheet2c_added_info5 = sheet1.loc[22]['Type']

    sheet2c['Description'] = sheet2c_added_info4
    sheet2c['Mitigation'] = sheet2c_added_info5

    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3

    #sheet1 part 4
    sheet2d = sheet1.iloc[24:25]
    sheet2d = sheet2d[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2d_added_info2 = [workbook]
    sheet2d_added_info3 = [UKPRN]
    sheet2d_added_info4 = sheet1.loc[25]['Type']
    sheet2d_added_info5 = sheet1.loc[28]['Type']

    sheet2d['Description'] = sheet2d_added_info4
    sheet2d['Mitigation'] = sheet2d_added_info5

    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3

    #sheet1 part 5
    sheet2e = sheet1.iloc[30:31]
    sheet2e = sheet2e[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2e_added_info2 = [workbook]
    sheet2e_added_info3 = [UKPRN]
    sheet2e_added_info4 = sheet1.loc[31]['Type']
    sheet2e_added_info5 = sheet1.loc[34]['Type']

    sheet2e['Description'] = sheet2e_added_info4
    sheet2e['Mitigation'] = sheet2e_added_info5

    sheet2e['Excel Title'] = sheet2e_added_info2
    sheet2e['UKPRN'] = sheet2e_added_info3

    #sheet1 part 6
    sheet2f = sheet1.iloc[36:37]
    sheet2f = sheet2f[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2f_added_info2 = [workbook]
    sheet2f_added_info3 = [UKPRN]
    sheet2f_added_info4 = sheet1.loc[37]['Type']
    sheet2f_added_info5 = sheet1.loc[40]['Type']

    sheet2f['Description'] = sheet2f_added_info4
    sheet2f['Mitigation'] = sheet2f_added_info5

    sheet2f['Excel Title'] = sheet2f_added_info2
    sheet2f['UKPRN'] = sheet2f_added_info3

    #sheet1 part 7
    sheet2g = sheet1.iloc[42:43]
    sheet2g = sheet2g[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2g_added_info2 = [workbook]
    sheet2g_added_info3 = [UKPRN]
    sheet2g_added_info4 = sheet1.loc[43]['Type']
    sheet2g_added_info5 = sheet1.loc[46]['Type']

    sheet2g['Description'] = sheet2g_added_info4
    sheet2g['Mitigation'] = sheet2g_added_info5

    sheet2g['Excel Title'] = sheet2g_added_info2
    sheet2g['UKPRN'] = sheet2g_added_info3


    #sheet1 part 8
    sheet2h = sheet1.iloc[48:49]
    sheet2h = sheet2h[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2h_added_info2 = [workbook]
    sheet2h_added_info3 = [UKPRN]
    sheet2h_added_info4 = sheet1.loc[49]['Type']
    sheet2h_added_info5 = sheet1.loc[52]['Type']

    sheet2h['Description'] = sheet2h_added_info4
    sheet2h['Mitigation'] = sheet2h_added_info5

    sheet2h['Excel Title'] = sheet2h_added_info2
    sheet2h['UKPRN'] = sheet2h_added_info3

    #sheet1 part 9
    sheet2i = sheet1.iloc[54:55]
    sheet2i = sheet2i[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2i_added_info2 = [workbook]
    sheet2i_added_info3 = [UKPRN]
    sheet2i_added_info4 = sheet1.loc[55]['Type']
    sheet2i_added_info5 = sheet1.loc[58]['Type']

    sheet2i['Description'] = sheet2i_added_info4
    sheet2i['Mitigation'] = sheet2i_added_info5

    sheet2i['Excel Title'] = sheet2i_added_info2
    sheet2i['UKPRN'] = sheet2i_added_info3

    #sheet1 part 10
    sheet2j = sheet1.iloc[60:61]
    sheet2j = sheet2j[['Scenario', 'Code', 'Area of Sensitivity', '2020', '2021','2022','Financial Health Grade & Score Impact']]

    sheet2j_added_info2 = [workbook]
    sheet2j_added_info3 = [UKPRN]
    sheet2j_added_info4 = sheet1.loc[61]['Type']
    sheet2j_added_info5 = sheet1.loc[64]['Type']

    sheet2j['Description'] = sheet2j_added_info4
    sheet2j['Mitigation'] = sheet2j_added_info5

    sheet2j['Excel Title'] = sheet2j_added_info2
    sheet2j['UKPRN'] = sheet2j_added_info3


    print(sheet1.to_string())

    print(sheet2a.to_string())
    print(sheet2b.to_string())
    print(sheet2c.to_string())
    print(sheet2d.to_string())
    print(sheet2e.to_string())
    print(sheet2f.to_string())
    print(sheet2g.to_string())
    print(sheet2h.to_string())
    print(sheet2i.to_string())
    print(sheet2j.to_string())


    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d, sheet2e, sheet2f, sheet2g, sheet2h, sheet2i,sheet2j])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Sensitivity Analysis')

def sheetFive():
    sheet1 = pd.read_excel(workbook, sheet_name='4. Balance Sheet')
    sheet1.columns = ['blank', 'Key', 'Name','Misc', '2020', '2021', '2022']
    sheet1 = sheet1[['Name', '2020', '2021', '2022']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[8:12]
    sheet2a = sheet2a[['Name','2020','2021','2022']]

    sheet2a_added_info = ["Non current assets" , "Non current assets" , "Non current assets" ,"Non current assets" , ]
    sheet2a_added_info2 = [workbook, workbook , workbook , workbook ]
    sheet2a_added_info3 = [UKPRN, UKPRN, UKPRN, UKPRN]

    sheet2a['Category'] = sheet2a_added_info
    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3

    # sheet1 part 2
    sheet2b = sheet1.iloc[13:14]
    sheet2b = sheet2b[['Name', '2020', '2021', '2022']]

    sheet2b_added_info = ["Non current assets" ]
    sheet2b_added_info2 = [workbook]
    sheet2b_added_info3 = [UKPRN]

    sheet2b['Category'] = sheet2b_added_info
    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3


    # sheet1 part 3
    sheet2c = sheet1.iloc[16:22]
    sheet2c = sheet2c[['Name', '2020', '2021', '2022']]

    sheet2c_added_info = ["Current Assets", "Current Assets", "Current Assets","Current Assets","Current Assets","Current Assets"]
    sheet2c_added_info2 = [workbook,workbook,workbook,workbook,workbook,workbook]
    sheet2c_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2c['Category'] = sheet2c_added_info
    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3


    # sheet1 part 4
    sheet2d = sheet1.iloc[24:34]
    sheet2d = sheet2d[['Name', '2020', '2021', '2022']]

    sheet2d_added_info = ["Creditors: amounts falling due within 1 year", "Creditors: amounts falling due within 1 year", "Creditors: amounts falling due within 1 year", "Creditors: amounts falling due within 1 year", "Creditors: amounts falling due within 1 year","Creditors: amounts falling due within 1 year","Creditors: amounts falling due within 1 year","Creditors: amounts falling due within 1 year","Creditors: amounts falling due within 1 year","Creditors: amounts falling due within 1 year"]
    sheet2d_added_info2 = [workbook,workbook,workbook,workbook,workbook,workbook,workbook,workbook,workbook,workbook]
    sheet2d_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2d['Category'] = sheet2d_added_info
    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3

    # sheet1 part 5
    sheet2e = sheet1.iloc[35:36]
    sheet2e = sheet2e[['Name', '2020', '2021', '2022']]

    sheet2e_added_info = ["Creditors: amounts falling due within 1 year"]
    sheet2e_added_info2 = [workbook]
    sheet2e_added_info3 = [UKPRN]

    sheet2e['Category'] = sheet2e_added_info
    sheet2e['Excel Title'] = sheet2e_added_info2
    sheet2e['UKPRN'] = sheet2e_added_info3

    # sheet1 part 6
    sheet2f = sheet1.iloc[37:38]
    sheet2f = sheet2f[['Name', '2020', '2021', '2022']]

    sheet2f_added_info = ["Creditors: amounts falling due within 1 year"]
    sheet2f_added_info2 = [workbook]
    sheet2f_added_info3 = [UKPRN]

    sheet2f['Category'] = sheet2f_added_info
    sheet2f['Excel Title'] = sheet2f_added_info2
    sheet2f['UKPRN'] = sheet2f_added_info3



    # sheet1 part 7
    sheet2g = sheet1.iloc[40:46]
    sheet2g = sheet2g[['Name', '2020', '2021', '2022']]

    sheet2g_added_info = ["Creditors: amounts falling due after 1 year","Creditors: amounts falling due after 1 year","Creditors: amounts falling due after 1 year","Creditors: amounts falling due after 1 year","Creditors: amounts falling due after 1 year","Creditors: amounts falling due after 1 year",]
    sheet2g_added_info2 = [workbook,workbook,workbook,workbook,workbook,workbook]
    sheet2g_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2g['Category'] = sheet2g_added_info
    sheet2g['Excel Title'] = sheet2g_added_info2
    sheet2g['UKPRN'] = sheet2g_added_info3

    # sheet1 part 8
    sheet2h = sheet1.iloc[48:51]
    sheet2h = sheet2h[['Name', '2020', '2021', '2022']]

    sheet2h_added_info = ["Provisions","Provisions","Provisions"]
    sheet2h_added_info2 = [workbook,workbook,workbook]
    sheet2h_added_info3 = [UKPRN,UKPRN,UKPRN]

    sheet2h['Category'] = sheet2h_added_info
    sheet2h['Excel Title'] = sheet2h_added_info2
    sheet2h['UKPRN'] = sheet2h_added_info3

    # sheet1 part 9
    sheet2i = sheet1.iloc[52:53]
    sheet2i = sheet2i[['Name', '2020', '2021', '2022']]

    sheet2i_added_info = ["Provisions"]
    sheet2i_added_info2 = [workbook]
    sheet2i_added_info3 = [UKPRN]

    sheet2i['Category'] = sheet2i_added_info
    sheet2i['Excel Title'] = sheet2i_added_info2
    sheet2i['UKPRN'] = sheet2i_added_info3

    # sheet1 part 10
    sheet2j = sheet1.iloc[56:61]
    sheet2j = sheet2j[['Name', '2020', '2021', '2022']]

    sheet2j_added_info = ["Reserves","Reserves","Reserves","Reserves","Reserves"]
    sheet2j_added_info2 = [workbook,workbook,workbook,workbook,workbook]
    sheet2j_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2j['Category'] = sheet2j_added_info
    sheet2j['Excel Title'] = sheet2j_added_info2
    sheet2j['UKPRN'] = sheet2j_added_info3


    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d, sheet2e, sheet2f, sheet2g, sheet2h, sheet2i,sheet2j])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Balance Sheet')

def sheetSix():
    sheet1 = pd.read_excel(workbook, sheet_name='5. Memorandum inputs')
    sheet1.columns = ['blank', 'Key', 'Name','Misc', '2020', '2021', '2022']
    sheet1 = sheet1[['Name', '2020', '2021', '2022']]


    #sheet1 part 1
    sheet2a = sheet1.iloc[10:12]
    sheet2a = sheet2a[['Name','2020','2021','2022']]

    sheet2a_added_info = ["Income" , "Income"  ]
    sheet2a_added_info2 = [workbook, workbook  ]
    sheet2a_added_info3 = [UKPRN, UKPRN]

    sheet2a['Category'] = sheet2a_added_info
    sheet2a['Excel Title'] = sheet2a_added_info2
    sheet2a['UKPRN'] = sheet2a_added_info3


    #sheet1 part 2
    sheet2b = sheet1.iloc[14:17]
    sheet2b = sheet2b[['Name','2020','2021','2022']]

    sheet2b_added_info = ["FRS 102 (28) LGPS - Net Return on Pension scheme" , "FRS 102 (28) LGPS - Net Return on Pension scheme" , "FRS 102 (28) LGPS - Net Return on Pension scheme"]
    sheet2b_added_info2 = [workbook, workbook , workbook]
    sheet2b_added_info3 = [UKPRN, UKPRN, UKPRN]

    sheet2b['Category'] = sheet2b_added_info
    sheet2b['Excel Title'] = sheet2b_added_info2
    sheet2b['UKPRN'] = sheet2b_added_info3

    #sheet1 part 3
    sheet2c = sheet1.iloc[18:19]
    sheet2c = sheet2c[['Name','2020','2021','2022']]

    sheet2c_added_info = ["Income"]
    sheet2c_added_info2 = [workbook]
    sheet2c_added_info3 = [UKPRN]

    sheet2c['Category'] = sheet2c_added_info
    sheet2c['Excel Title'] = sheet2c_added_info2
    sheet2c['UKPRN'] = sheet2c_added_info3


    #sheet1 part 4
    sheet2d = sheet1.iloc[20:24]
    sheet2d = sheet2d[['Name','2020','2021','2022']]

    sheet2d_added_info = ["Income","Income","Income","Income"]
    sheet2d_added_info2 = [workbook,workbook,workbook,workbook]
    sheet2d_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2d['Category'] = sheet2d_added_info
    sheet2d['Excel Title'] = sheet2d_added_info2
    sheet2d['UKPRN'] = sheet2d_added_info3

    #sheet1 part 5
    sheet2e = sheet1.iloc[25:26]
    sheet2e = sheet2e[['Name','2020','2021','2022']]

    sheet2e_added_info = ["Income"]
    sheet2e_added_info2 = [workbook]
    sheet2e_added_info3 = [UKPRN]

    sheet2e['Category'] = sheet2e_added_info
    sheet2e['Excel Title'] = sheet2e_added_info2
    sheet2e['UKPRN'] = sheet2e_added_info3


    #sheet1 part 6
    sheet2f = sheet1.iloc[29:34]
    sheet2f = sheet2f[['Name','2020','2021','2022']]

    sheet2f_added_info = ["Expenditure","Expenditure","Expenditure","Expenditure","Expenditure"]
    sheet2f_added_info2 = [workbook,workbook,workbook,workbook,workbook]
    sheet2f_added_info3 = [UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]

    sheet2f['Category'] = sheet2f_added_info
    sheet2f['Excel Title'] = sheet2f_added_info2
    sheet2f['UKPRN'] = sheet2f_added_info3

    #sheet1 part 7
    sheet2g = sheet1.iloc[36:38]
    sheet2g = sheet2g[['Name','2020','2021','2022']]

    sheet2g_added_info = ["Staff Costs","Staff Costs"]
    sheet2g_added_info2 = [workbook,workbook]
    sheet2g_added_info3 = [UKPRN,UKPRN]

    sheet2g['Category'] = sheet2g_added_info
    sheet2g['Excel Title'] = sheet2g_added_info2
    sheet2g['UKPRN'] = sheet2g_added_info3

    #sheet1 part 8
    sheet2h = sheet1.iloc[41:42]
    sheet2h = sheet2h[['Name','2020','2021','2022']]

    sheet2h_added_info = ["Balance Sheet"]
    sheet2h_added_info2 = [workbook]
    sheet2h_added_info3 = [UKPRN]

    sheet2h['Category'] = sheet2h_added_info
    sheet2h['Excel Title'] = sheet2h_added_info2
    sheet2h['UKPRN'] = sheet2h_added_info3

    #sheet1 part 9
    sheet2i = sheet1.iloc[45:46]
    sheet2i = sheet2i[['Name','2020','2021','2022']]

    sheet2i_added_info = ["Auto - Calculations"]
    sheet2i_added_info2 = [workbook]
    sheet2i_added_info3 = [UKPRN]

    sheet2i['Category'] = sheet2i_added_info
    sheet2i['Excel Title'] = sheet2i_added_info2
    sheet2i['UKPRN'] = sheet2i_added_info3


    #sheet1 part 10
    sheet2j = sheet1.iloc[47:48]
    sheet2j = sheet2j[['Name','2020','2021','2022']]

    sheet2j_added_info = ["Auto - Calculations"]
    sheet2j_added_info2 = [workbook]
    sheet2j_added_info3 = [UKPRN]

    sheet2j['Category'] = sheet2j_added_info
    sheet2j['Excel Title'] = sheet2j_added_info2
    sheet2j['UKPRN'] = sheet2j_added_info3

    print(sheet1.to_string())
    print(sheet2a.to_string())
    print(sheet2b.to_string())
    print(sheet2c.to_string())
    print(sheet2d.to_string())
    print(sheet2e.to_string())
    print(sheet2f.to_string())
    print(sheet2g.to_string())
    print(sheet2h.to_string())
    print(sheet2i.to_string())
    print(sheet2j.to_string())


    df_sheet1_clen = pd.concat([sheet2a,sheet2b,sheet2c,sheet2d, sheet2e, sheet2f, sheet2g, sheet2h, sheet2i,sheet2j])
    print(df_sheet1_clen.to_string())

    df_sheet1_clen.to_excel(writer,sheet_name='Memorandum inputs')



#########WorkbookTwo#############
def workbookTwo():
    sheet1 = pd.read_excel(workbook2, sheet_name='Monthly cash flow')

    UKPRN = sheet1.loc[2].values[0]

    print(UKPRN)
    print(sheet1)


    filename = os.path.basename(os.path.normpath(workbook2))
    print(filename)
    sheet1.columns = ['Key','Month 1', 'Month 2','Month 3', 'Month 4', 'Month 5' ,'Month 6', 'Month 7', 'Month 8', 'Month 9', 'Month 10', 'Month 11', 'Month 12', 'Month 13', 'Month 14', 'Month 15', 'Month 16', 'Month 17', 'Month 18', 'Month 19', 'Month 20', 'Month 21', 'Month 22', 'Month 23', 'Month 24', 'Month 25', 'Month 26', 'Month 27']

    sheet1 = sheet1.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,26,34,37,39,45,51,47,54,56,28,41,58,65,77,63,69,72,74,76,79,82,84,87,89,93])


    sheet1_added_info1 = [filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename,filename]
    sheet1_added_info2 = [UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN, UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN,UKPRN]


    sheet1['Excel Title'] = sheet1_added_info1
    sheet1['UKPRN'] = sheet1_added_info2



    print(sheet1.to_string())
    sheet1.to_excel(UKPRN + filename,sheet_name='Monthly cash flow')


#########FinalExcelProductionArea#############

#writer = pd.ExcelWriter('10002111_FH Calculator v1.4a.xlsx')

#########WorkbookOne#############

#sheetOne()
#sheetTwo()
#sheetThree()

#sheetFive()
#sheetSix()

#writer.save()
#########WorkbookTwo#############
workbookTwo()

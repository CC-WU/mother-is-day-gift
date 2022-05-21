from datetime import datetime
from datetime import date, timedelta
import pandas as pd
import numpy as np

date = datetime.today().strftime('%Y%m%d')
sheet = datetime.today().strftime('%m%d')

FilePath = "COVID-19確診名冊.xlsx"

try:
    df1 = pd.read_excel("COVID-19確診名冊.xlsx", sheet_name=sheet, converters = {u'Unnamed: 0':str, u'通報日期':str, u'通報單號':str, u'證號':str, u'姓名':str, u'接觸者人數':str, u'處理人員':str, u'處理日期':str, u'派案':str, u'Unnamed: 9':str})
except:
        df1 = pd.read_excel("COVID-19確診名冊.xlsx", sheet_name=sheet, converters = {u'Unnamed: 0':str, u'通報日期':str, u'通報單號':str, u'證號':str, u'姓名':str, u'接觸者人數':str, u'處理人員':str, u'處理日期':str, u'派案':str})
df2 = pd.read_excel("Reports_"+date+".xlsx", sheet_name="通報單列表", converters = {u'Unnamed: 0':str, u'通報單號':str, u'證號':str, u'姓名':str})
df1 = df1.astype(str)
df2 = df2.astype(str)
df1.fillna('', inplace=True)
df2.fillna('', inplace=True)
df1.drop('Unnamed: 0', inplace=True, axis=1)
df1 = df1.rename(columns={'Unnamed: 0': ''})
df2 = df2.rename(columns={'Unnamed: 0': ''})

try:
    df1 = df1.rename(columns={'Unnamed: 9': ''})
except:
    pass
    
df2["通報日期"] = sheet
frames = [df1, df2]

df = pd.concat(frames)
df_final = df.drop_duplicates(subset=['通報日期', '通報單號', '證號', '姓名'], keep = 'first')
df_final.index = np.arange(1, len(df_final)+1)

df_final.fillna('', inplace=True)
df_final['接觸者人數'] = df_final['接觸者人數'].replace({'nan':''})
df_final['處理人員'] = df_final['處理人員'].replace({'nan':''})
df_final['處理日期'] = df_final['處理日期'].replace({'nan':''})
df_final['派案'] = df_final['派案'].replace({'nan':''})
try:    
    df_final[''] = df_final[''].replace({'nan':''})
except:
    pass

writer = pd.ExcelWriter(FilePath, mode='a', engine = 'openpyxl', if_sheet_exists = 'replace')

df_final.to_excel(writer, sheet_name = sheet, index=True)
# writer.save()
writer.close()
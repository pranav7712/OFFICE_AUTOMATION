import pandas as pd
import numpy as np
from fiscalyear import *

df=pd.read_excel("G:\OFFICE AUTOMATION FILES\COMBINE GSTR2A\SAMPLE_032021_R2AGSTR2A_all_combined.xlsx",sheet_name=4)

df['PAN_Number']=df["GSTIN_of_Supplier"].apply(lambda x:x[2:12:1])

df=df.replace(np.nan,"",regex=True)

df["PAN_3_Way_Key"]=np.where(df["Sheet_Name"]=="B2BA",df["PAN_Number"]+"/"+df["Inv_CN_DN_Number_Revised"]+"/"
                             +df["Inv_CN_DN_Date_Revised_Unique"],df["PAN_Number"] + "/" + df["Inv_CN_DN_Number_Original"]
                             + "/" + df["Inv_CN_DN_Date_Unique"])

df["PAN_2_Way_Key_PAN_InvNo"]=np.where(df["Sheet_Name"]=="B2BA",df["PAN_Number"]+"/"+df["Inv_CN_DN_Number_Revised"]
                                       ,df["PAN_Number"] + "/" + df["Inv_CN_DN_Number_Original"])

df["PAN_3_Way_Key_PAN_InvDt"]=np.where(df["Sheet_Name"]=="B2BA",df["PAN_Number"]+"/"+df["Inv_CN_DN_Date_Revised_Unique"]
                                       ,df["PAN_Number"] +"/" + df["Inv_CN_DN_Date_Unique"])


df.to_excel("G:\OFFICE AUTOMATION FILES\COMBINE GSTR2A\SASSSSMPLE_032021_R2AGSTR2A_all_combined.xlsx",sheet_name="Combined",index=False)
print(df)


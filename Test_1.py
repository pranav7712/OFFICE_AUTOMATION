import pandas as pd

df=pd.read_excel('G:\OFFICE AUTOMATION FILES\COMBINE GSTR2A\SAMPLE_032021_R2AGSTR2A_all_combined.xlsx',sheet_name=4)


# df['Ultimately_Unique']=df["Sheet_Name"]+"/"+df["Supply_Attract_Reverse_Charge"]+df["GSTR_1_5_Filing_Status"]+"/"+df["Unique_ID"]


df['Ultimately_Unique']=str(df["Sheet_Name"]+"/"+df["Supply_Attract_Reverse_Charge"]+"/"+df["Unique_ID"])

df['PAN_Number']=str(df["GSTIN_of_Supplier"])


df.to_excel('G:\OFFICE AUTOMATION FILES\COMBINE GSTR2A\SASSSAMPLE_032021_R2AGSTR2A_all_combined.xlsx',sheet_name="Combined",index=False)
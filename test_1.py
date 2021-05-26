import pandas as pd

dicteg={"col1":["Vc1","Vc2","Vc3"],"col2":["Name1","Name2","Name3"]}

df1=pd.DataFrame(dicteg)
print(df1)

dicteg2={"col2":["Vc1","Vc2","Vc3"],"col4":["Name1","Name2","Name3"]}

df2=pd.DataFrame(dicteg2)
print(df2)

df3=pd.concat([df1,df2],axis=0)
print(df3)

df4=df1.append(df2)
print(df4)

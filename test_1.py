import datacompy
import pandas as pd
import numpy as np
import ast


df1=pd.read_excel("G:\OFFICE AUTOMATION FILES\Difference Finder\Sales Report_Original.xlsx",sheet_name=0)
df2=pd.read_excel("G:\OFFICE AUTOMATION FILES\Difference Finder\Sales Report_Revised.xlsx",sheet_name=0)

listcommoncol=["","","Segment"]

commoncol=np.intersect1d(df2.columns,df1.columns,listcommoncol)

print(commoncol)

print("the len is {} and the common cols are {}".format(len(commoncol),",".join(commoncol)))

a_set=set(commoncol)
b_set=set(listcommoncol)

common=list(set(a_set.intersection(b_set)))

print(common)
# e_1="Pranav"
# e_2="Pratik"
# e_3="Tulshyan"
#
#
# for x in (e_3,e_2,e_1):
#     new=list.append(x)
#
# print(new)

str1 = "Pranav"
str2 = ""
str3 = ""

listfinal=[]

listfinal.append(str3)
listfinal.append(str2)
listfinal.append(str1)

print(type(listfinal))
# listcomcol = []
#
# for x in (str1, str2, str3):
#     a=ast.parse(x,mode='eval')
#     print(eval(compile(a,"",mode="eval")))


print(listfinal)

ppt=["Prana"]

ppt.append("Pranav")

print(ppt)
# comparison=datacompy.Compare(df1, df2,join_columns="S.N")
#
# print (comparison.report())
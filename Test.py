import pandas as pd


df1= pd.DataFrame({'Student': ['Ram','Rohan','Shyam','Mohan'],
        'Grade': ['A','C','B','Ex']})
df2 = pd.DataFrame({'Student': ['Ram','Shyam','Raunak'],
        'Grade': ['A','B','F']})

df3=pd.merge(df1,df2, how='inner')

print(df1)
print(df2)
print(df3)
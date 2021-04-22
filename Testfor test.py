import pandas as pd

print("Heloo \n \t World")
a=2*3
print ("a")
print('Hello World'+ str(a))


a={"Month":["April","May","June"],
   "Country":["Nepal","India","USA"],
   "PCI":[500,400,1000]

}

print(a)

df=pd.DataFrame(a)

print(df)
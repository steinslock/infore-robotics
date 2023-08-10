import pandas as pd

data = pd.DataFrame((['23-10-12'],['22-03-15']), columns=['time'])

a = '23-05-20'
data1 = pd.DataFrame([a],columns=['time'])
data = data._append(data1)

print(data)


abc = ['1', '2', '3']  # index start from 0
df_tmp = pd.DataFrame([[abc[0], abc[1], abc[2]]],columns=['time', 'num', 'data'])
print(df_tmp)

import pandas as pn

data = pn.read_csv('titanic.csv', index_col='PassengerId')
res = data['Age'].describe()

print(res[1], res[5])

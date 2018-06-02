import pandas as pn

data = pn.read_csv('titanic.csv', index_col='PassengerId')
res = data.corr(method='pearson')

print(res)

#random matrica, creattion
import numpy as np
x = np.random.normal(loc=1, scale=10, size=(1000, 50))
#loc - average (mean) normal destribution, scale - variance

m = np.mean(x, axis=0)
std = np.std(x, 0)
X_result = ((x - m) / std)

'y = np.random.randint(0, high=5, size=(5, 3))'

y = np.array([[4, 5, 0],
              [1, 9, 3],
              [5, 3, 6],
              [7, 0, 1],
              [5, 4, 0],
             [3, 6, 7],
             [5, 2, 3]])

s = np.sum(y, axis=1)

A = np.eye(3, k=1)
B = np.eye(3)
r = np.hstack((B, A))
print(r)

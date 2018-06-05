import numpy as np
import matplotlib.pyplot as plt
from scipy import signal

from sklearn.decomposition import FastICA, PCA

# generate sample data
np.random.seed(0)
n_sample = 2000  # x axis
time = np.linspace(0, 8, n_sample)

s1 = np.sin(2 * time)   # sinusoidal signal
s2 = np.sign(np.sin(3 * time))  # square signal
s3 = signal.sawtooth(2 * np.pi * time)  # saw tooth signal

S = np.c_[s1, s2, s3]
S += 0.2 * np.random.normal(size=S.shape)  # Add noise

S /= S.std(axis=0)  # standardize data

#  Mix data
A = np.array([[1, 1, 1], [0.5, 2, 1.0], [1.5, 1.0, 2.0]])  # Mixing matrix
X = np.dot(S, A.T)  # Generate observations

# Compute ICA
ica = FastICA(n_components=3)
S_ = ica.fit_transform(X)  #  Reconstruct signals
A_ = ica.mixing_  # Get estimated mixing matrix

# We can prove that the ICA model applies by reverting the unmixing
assert np.allclose(X, np.dot(S_, A_.T) + ica.mean_)

# for comparison, compute PCA
pca = PCA(n_components=3)
H = pca.fit_transform(X)

######################
#  Plot results

plt.figure()

models = [X, S, S_, H]
names = ['Observation (mixed signal)',
         'True Sources',
         'ICA recovered signals',
         'PCA recovered signals']
colors = ['red', 'steelblue', 'orange']

for ii, (model, name) in enumerate(zip(models, names), 1):
    plt.subplot(4, 1, ii)
    plt.title(name)
    for sig, color in zip(model.T, colors):
        plt.plot(sig, color=color)

plt.subplots_adjust(0.09, 0.04, 0.94, 0.94, 0.26, 0.46)
plt.show()


# https://www.youtube.com/watch?v=GfIQlql-i2k

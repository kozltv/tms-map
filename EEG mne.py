import numpy as np
import mne
from mne.datasets import sample

data_path = '/Users/kseniya/Desktop/MY/Lab/ЭЭГ данные/markina/neurobar_ppy_1.vhdr'
raw = mne.channels.read_montage(data_path)


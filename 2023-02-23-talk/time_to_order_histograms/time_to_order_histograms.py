#! /usr/bin/env python3


import pickle

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


inv_and_orders = pd.read_csv('inventory_and_orders.csv', index_col=0, skipfooter=1)\
        .iloc[:, 0:12]

# Note to self: shape is (10000, 51): 10000 universes x 51 time steps
with open('order_streams.pkl', 'rb') as pf:
    order_streams = pickle.load(pf)

times_to_reorder = np.zeros((order_streams.shape[0], 4), dtype='int8')

cols = ['worst_case', 'industry', 'heuristic', 'optimized']

for p, policy in enumerate(cols):
    for i in range(order_streams.shape[0]):
        curr_inv = inv_and_orders.loc['inventory'].values \
                + inv_and_orders.loc[policy].values
        t = 0
        while np.all(curr_inv >= 0):
            t += 1
            curr_inv[order_streams[i, t]] -= 1

        times_to_reorder[i, p] = t

times_frame = pd.DataFrame(times_to_reorder, columns=cols)

plt.style.use('seaborn-deep')
bins = np.arange(0, 45)

plt.hist(times_frame.values, bins, label=times_frame.columns)
plt.legend(loc='upper right')
plt.xlabel('Orders before the next re-order')
plt.ylabel('Count of universes')
plt.show()

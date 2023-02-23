'''Order optimization functions.'''

import numpy as np


def heuristic(inv_arr, order_streams, order_size):
    # To build the order with maximum expected time to next reorder, track the
    # logical inventory from today assuming we never ordered more shirts, until
    # we have exactly ORDER_SIZE backorders.
    #
    backorders = np.zeros(inv_arr.shape)
    sim_size = order_streams.shape[0]

    # TODO While runtime is just a few seconds for SIM_SIZE=1e4, this is
    # inefficient.  Cleverly vectorize?
    for i in range(sim_size):
        curr_inv_arr = inv_arr.copy()
        j = 0
        while curr_inv_arr[curr_inv_arr < 0].sum() > -1 * order_size:
            curr_inv_arr[order_streams[i][j]] -= 1
            j += 1

        # Select the sizes with negative logical inventory after this simulation of
        # shirts sent out, then accumulate them in the backorders array.
        curr_neg_inv = curr_inv_arr.copy()
        curr_neg_inv[curr_neg_inv > 0] = 0
        backorders += -1 * curr_neg_inv

    assert backorders.sum() == sim_size * order_size

    backorders_divided = backorders / sim_size
    backorders_rounded = np.rint(backorders_divided).astype('int')

    # Sometimes rounding errors pile up, and we want exactly ORDER_SIZE shirts in
    # the order. When this happens, adjust the shirts with the biggest rounding
    # errors. In other words, add shirts to the sizes where backorders_divided had
    # the largest fractional parts below 0.5, or subtract shirts from the sizes
    # where backorders_divided had the smallest fractional parts above 0.5.
    sum_error = backorders_rounded.sum() - order_size
    if sum_error != 0:
        rounding_errors = backorders_divided - backorders_rounded
        adjustment_indices = np.argsort(rounding_errors)
        if sum_error > 0:
            for i in range(sum_error):
                # Round down the quantities that had been rounded up the most
                idx_to_decrement = adjustment_indices[i]
                backorders_rounded[idx_to_decrement] -= 1
        elif sum_error < 0:
            for i in range(-1 * sum_error):
                # Round up the quantities that had been rounded down the most
                idx_to_increment = np.flip(adjustment_indices)[i]
                backorders_rounded[idx_to_increment] += 1

    assert backorders_rounded.sum() == order_size

    return backorders_rounded

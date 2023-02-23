'''Order optimization functions.'''

import copy
import numpy as np
import pandas as pd


def worst_case(inv_arr, order_streams, order_size):
    '''Conservatively target equal inventory for every size.

    This guarantees a minimum level of performance in a worst-case scenario
    where the future brings a uniform stream of unexpected (e.g., W2XL or MXS)
    orders.
    '''
    order = np.zeros(inv_arr.shape, dtype='int')
    curr_inv_arr = inv_arr.copy()

    # Tiebreaker as we build the order
    _, counts = np.unique(order_streams, return_counts=True)

    for i in range(order_size):
        most_popular_count = counts.min()  # Seed the tiebreaker below

        scarcest_qty = np.min(curr_inv_arr)
        for j in range(len(curr_inv_arr)):
            if curr_inv_arr[j] == scarcest_qty:
                if counts[j] >= most_popular_count:
                    size_to_order = j
                    most_popular_count = counts[j]

        order[size_to_order] += 1
        curr_inv_arr[size_to_order] += 1

    return order


def industry(inv_arr, prior_size_hist, order_size):
    '''Build order to attempt to match industry expectations.

    Policy: Ensure we end up with at least one of every size. But beyond that,
    attempt to match the size distribution quoted in the industry blog.
    '''
    order = np.zeros(inv_arr.shape, dtype='int')
    curr_inv_arr = inv_arr.copy()
    industry_dist = np.array(list(prior_size_hist.values()))

    for i in range(order_size):
        scarcest_qty = np.min(curr_inv_arr)

        if scarcest_qty < 1:
            # First ensure we end up with at least one of every size, after
            # filling backorders
            size_to_order = np.argmin(curr_inv_arr)
            order[size_to_order] += 1
            curr_inv_arr[size_to_order] += 1
        else:
            # Add a shirt wherever we're farthest below the industry
            # distribution
            assert np.min(curr_inv_arr) >= 1
            dist_diff = curr_inv_arr / curr_inv_arr.sum() - industry_dist

            size_to_order = np.argmin(dist_diff)
            order[size_to_order] += 1
            curr_inv_arr[size_to_order] += 1

    return order


def heuristic(inv_arr, order_streams, order_size):
    '''Use a heuristic to get a near-optimal order.

    To build the order with maximum expected time to next reorder, track the
    logical inventory from today assuming we never ordered more shirts, until
    we have exactly ORDER_SIZE backorders.
    '''
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


def optimized(inv_arr, order_streams, order_size, max_dist=2):
    '''Do a small-scale brute-force search around the heuristic solution.'''

    sizes=len(inv_arr)
    heuristic_order = heuristic(inv_arr, order_streams, order_size)

    # Every adjustment of the order can be decomposed into steps of moving a
    # shirt from one size to another.
    elem_adjustments = set()
    for i in range(sizes):
        for j in range(sizes):
            if i != j:
                adj = [0 for _ in range(sizes)]
                adj[i] = 1
                adj[j] = -1
                elem_adjustments.add(tuple(adj))

    if max_dist > 1:
        curr_adjustments = copy.deepcopy(elem_adjustments)
        for stage in range(1, max_dist):
            new_adjustments = copy.deepcopy(curr_adjustments)
            for ca in curr_adjustments:
                for ea in elem_adjustments:
                    new_adjustments.add(
                            tuple(ca[i] + ea[i] for i in range(sizes)))
            curr_adjustments = new_adjustments
        adjustments = new_adjustments
    else:
        adjustments = elem_adjustments

    # Ensure the starting point itself is in the mix
    adjustments.add(tuple(0 for _ in range(sizes)))

    # Trial orders: increment heuristic order by the adjustments and throw out
    # any orders with negative quantities
    trials_with_negatives = np.array(list(adjustments)) + heuristic_order
    trial_orders = trials_with_negatives[np.all(trials_with_negatives >= 0, axis=1)]

    # Accumulate the order streams to build a larger data structure, of
    # the backorder after each order assuming we didn't order anything.
    snapshots = np.zeros((sizes, order_streams.shape[0], order_streams.shape[1]),
            dtype='int8')
    snapshots[:, :, 0] = inv_arr[:, np.newaxis]

    # TODO didn't seem to unduly hit runtime but I'm sure there's a better
    # vectorized way. I couldn't figure out a correct one at the time of
    # writing.
    for t in range(1, order_streams.shape[1]):
        snapshots[:, :, t] = snapshots[:, :, t - 1]
        for j in range(order_streams.shape[0]):
            snapshots[order_streams[j, t], j, t] -= 1

    best_mean_reorder_time = 0
    for order in trial_orders:
        curr_snapshots = snapshots + order[:, np.newaxis, np.newaxis]
        # TODO Bizarre to bring Pandas in just for this. It's a facile way to
        # get the reorder time for each stream of orders
        reorder_times = pd.DataFrame(
                np.argwhere(curr_snapshots == -1))\
            .groupby(1)[2]\
            .min()\
            .values
        mean_reorder_time = np.mean(reorder_times)

        if mean_reorder_time > best_mean_reorder_time:
            best_order = order
            best_mean_reorder_time = mean_reorder_time

    return best_order

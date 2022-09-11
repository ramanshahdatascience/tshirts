# A t-shirt reorder: a case study in optimal inventory management

This repo contains the homemade inventory system I use to manage the flow of
promotional t-shirts I send out as gifts as part of my marketing.

Key assets:

- `tshirt_inventory.xlsx` (included as `tshirt_inventory_redacted.xlsx`): a
  spreadsheet to capture and organize the flow of individual shirts going out,
  boxes of bulk shirts coming in from the screen printing shop, and tabulations
  of these to track when I'm running out of certain sizes.
- `inventory_to_shippo_labels.py`: a script to turn the inventory spreadsheet
  into a CSV file understood by [Shippo](https://goshippo.com), which draws
  upon a crib sheet to turn t-shirt sizes into order weights for the postage.
- `address_parser.py`: a hand-rolled international mailing address parser, used
  by `inventory_to_shippo_labels.csv`. It only has parsing rules for the few
  countries I've sent t-shirts to - definitely not comprehensive for public
  use.
- `build_order.py`: the interesting thing from a data scientist's point of
  view: a script that uses Bayesian inference (soon, multilevel models) and
  discrete optimization (soon, in the nontrivial sense) to build an optimal
  re-order for a box of shirts. In other words, it answers, "if I have this
  inventory on hand, and that order history, how many of each size should I
  buy?"

## Usage

```
# Create a CSV suitable for import into Shippo
./inventory_to_shippo_labels.py tshirt_inventory.xlsx xxxx-yy-zz-labels.csv

# Optimal order to console
./build_order.py tshirt_inventory.xlsx
./build_order.py tshirt_inventory.xlsx -o console

# Write optimal order to the "Hypothetical order" row in the "inventory" sheet
./build_order.py tshirt_inventory.xlsx -o hypothetical

# Finalize optimal order to the "incoming" sheet with today's date
./build_order.py tshirt_inventory.xlsx -o final
```

As of writing, the commands that write the optimal order data to the
spreadsheet are fiddly. One must **save the workbook in Excel, then close the
workbook** to avoid the program crashing.

## Prior

According to [this blog post by a t-shirt
wholesaler](https://www.theadairgroup.com/blog/2020/06/01/shirt-order-size-distribution-what-sizes-to-order-for-t-shirts/),
the global t-shirt size distribution is as follows:

- XS: 1 percent
- S: 7 percent
- M: 28 percent
- L: 30 percent
- XL: 20 percent
- 2XL: 12 percent
- 3XL: 2 percent

This information can feed into an informative prior for building an optimal
order.

#! /usr/bin/env python3


import copy
import csv
import math
import re
from sys import argv
import warnings

import openpyxl

# Measured masses of t-shirts by size, in oz
SHIRT_WEIGHTS = {'MXS': 3.45 / 1,
                 'MS': 11.45 / 3,
                 'MM': 38.80 / 9,
                 'ML': 33.35 / 7,
                 'MXL': 21.15 / 4,
                 'M2XL': 17.60 / 3,
                 'M3XL': 6.55 / 1,
                 'WS': 9.75 / 3,
                 'WM': 18.05 / 5,
                 'WL': 16.55 / 4,
                 'WXL': 17.95 / 4,
                 'W2XL': 9.95 / 2}

# Empirical, based on finished packages:
# MM: 7.65 oz
# ML: 8.45 oz
BALANCE_OF_SHIPMENT = \
    ((7.65 + 8.45) - (SHIRT_WEIGHTS['MM'] + SHIRT_WEIGHTS['ML'])) / 2.0 # ~3.51 oz

# Shipping weights are actually uncertain due to manufacturing variance and
# taping technique. To demonstrate the uncertainty, I weighed a box (3.15 oz)
# and card with envelope (0.40 oz) and the sum (3.55 oz) is more than the
# empirical balance of shipment, which includes a label and tape. I saw
# individual shirt weights vary ~0.5 oz min to max. Add a small margin of
# safety to reduce the risk of buying too little postage (noting that having
# too much margin jacks typical MM orders, which weigh close to 8 oz, from 8 oz
# to 9 oz, which incurs a big price increase).
MARGIN_OF_SAFETY = 0.15

# Hard-code USA addresses
SHIPPO_FIELDS = {'Order Number': None,
                 'Order Date': None,
                 'Recipient Name': None,
                 'Company': None,
                 'Email': None,
                 'Phone': None,
                 'Street Line 1': None,
                 'Street Number': None,
                 'Street Line 2': None,
                 'City': None,
                 'State/Province': None,
                 'Zip/Postal Code': None,
                 'Country': 'USA',
                 'Item Title': None,
                 'SKU': None,
                 'Quantity': None,
                 'Item Weight': None,
                 'Item Weight Unit': None,
                 'Item Price': None,
                 'Item Currency': None,
                 'Order Weight': None,
                 'Order Weight Unit': 'oz',
                 'Order Amount': None,
                 'Order Currency': None}

STATE_AND_ZIP = re.compile(r'(.*) ([A-Z]{2}),? ([0-9]{5})$')

def shippo_details(size_text, name_text, address_text):
    '''Parse spreadsheet row into Shippo fields.'''
    result = copy.deepcopy(SHIPPO_FIELDS)

    # Real addresses and not notes to hand-deliver have parts, thus commas
    if address_text.find(', ') > -1:
        state_and_zip_match = STATE_AND_ZIP.search(address_text.strip())
        if state_and_zip_match:
            result['Recipient Name'] = name_text
            result['State/Province'] = state_and_zip_match.group(2)
            result['Zip/Postal Code'] = state_and_zip_match.group(3)

            result['Order Weight'] = math.ceil(
                SHIRT_WEIGHTS[size_text] +
                BALANCE_OF_SHIPMENT +
                MARGIN_OF_SAFETY)

            address_city = state_and_zip_match.group(1).strip(',').split(', ')
            result['City'] = address_city[-1]
            address = address_city[:-1]

            if len(address) == 1:
                for unit_identifier in ['Apt', '#', 'Apartment', 'Unit']:
                    partition_idx = address[0].find(unit_identifier)
                    if partition_idx > -1:
                        result['Street Line 1'] = \
                            address[0][:partition_idx].strip()
                        result['Street Line 2'] = \
                            address[0][partition_idx:].strip()
                        break
                if result['Street Line 1'] is None:
                    result['Street Line 1'] = address[0]
            elif len(address) == 2:
                result['Street Line 1'] = address[0].strip()
                result['Street Line 2'] = address[1].strip()
            else:
                raise(Exception)

        else:
            # Don't try on non-US addresses for now; they are rare
            # TODO parse common international addresses
            warnings.warn(
                f'Possible international address found: {address_text}')
            result = None
    else:
        result = None

    return result


wb = openpyxl.load_workbook(filename=argv[1], data_only=True)
results = []
assert wb['outgoing']['C1'].value == 'name'
assert wb['outgoing']['D1'].value == 'address'
assert wb['outgoing']['E1'].value == 'shipped'

ir = wb['outgoing'].iter_rows()
next(ir)  # Header row
for row in ir:
    size = row[1].value
    if '2XL' in size or '3XL' in size:
        warnings.warn('Shirts bigger than XL may need a bigger box.')

    name = row[2].value
    address = row[3].value
    shipped = row[4].value

    if shipped != 'Y':
        fields = shippo_details(size, name, address)

        if fields is not None:
            results.append(fields)

with open(argv[2], 'w') as cf:
    writer = csv.DictWriter(cf, fieldnames=SHIPPO_FIELDS.keys())

    writer.writeheader()
    for result in results:
        writer.writerow(result)

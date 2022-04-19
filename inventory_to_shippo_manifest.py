#! /usr/bin/env python3


import copy
import csv
import re
from sys import argv
import warnings

import openpyxl

# Hard-code a 9 oz package sent to USA addresses
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
                 'Order Weight': 9,
                 'Order Weight Unit': 'oz',
                 'Order Amount': None,
                 'Order Currency': None}

STATE_AND_ZIP = re.compile(r'(.*) ([A-Z]{2}),? ([0-9]{5})$')

def address_cell_to_shippo_details(name_text, address_text):
    '''Parse my comma-separated addresses into Shippo fields.'''
    result = copy.deepcopy(SHIPPO_FIELDS)

    # Real addresses and not notes to hand-deliver have parts, thus commas
    if address_text.find(', ') > -1:
        state_and_zip_match = STATE_AND_ZIP.search(address_text.strip())
        if state_and_zip_match:
            result['Recipient Name'] = name_text
            result['State/Province'] = state_and_zip_match.group(2)
            result['Zip/Postal Code'] = state_and_zip_match.group(3)

            # TODO Look up postage weight as a function of shirt size.
            # Empirically, measuring finished packages with a card, tape, and label:
            # MM: 7.65 oz
            # ML: 8.45 oz
            # It will save some pennies to buy <= 8 oz of postage when
            # applicable, and >= XL shirts may require >= 10 oz. (The biggest
            # shirts probably won't fit in the standard packaging and will
            # require a Flat Rate Box or similar.) To do this, easiest will
            # probably be to measure the shirt weights and extrapolate the
            # packed weight.

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
    name = row[2].value
    address = row[3].value
    shipped = row[4].value

    if shipped != 'Y':
        fields = address_cell_to_shippo_details(name, address)

        if fields is not None:
            results.append(fields)

with open(argv[2], 'w') as cf:
    writer = csv.DictWriter(cf, fieldnames=SHIPPO_FIELDS.keys())

    writer.writeheader()
    for result in results:
        writer.writerow(result)

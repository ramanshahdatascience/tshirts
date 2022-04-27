#! /usr/bin/env python3

'''Usage:

./inventory_to_shippo_labels.py tshirt_inventory.xlsx labels.csv'''


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

DEFAULT_COUNTRY = 'US'
COUNTRIES = {
    'US': {'fields': ['City', 'State/Province', 'Zip/Postal Code'],
           'postal_code_regex': re.compile(r'[0-9]{5}(-[0-9]{4})?')},
    'GB': {'fields': ['City', 'Zip/Postal Code'],
           'postal_code_regex': re.compile(
               r'[A-Z][A-Z0-9]{1,3} [0-9][A-Z]{2}')}}

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
                 'Country': None,
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

def shippo_details(size_text, name_text, address_text):
    '''Parse spreadsheet row into Shippo fields.'''
    result = copy.deepcopy(SHIPPO_FIELDS)

    # Real addresses and not notes to hand-deliver have parts, thus commas
    if address_text.find(', ') > -1:
        result['Recipient Name'] = name_text
        result['Order Weight'] = math.ceil(
            SHIRT_WEIGHTS[size_text] +
            BALANCE_OF_SHIPMENT +
            MARGIN_OF_SAFETY)

        address_fields = _address_fields(address_text)
        result.update(address_fields)

    else:
        result = None

    return result

def _address_fields(address_text):
    fields = {}

    if address_text[-2:] in COUNTRIES:
        country = address_text[-2:]
        fields['Country'] = country
        remainder = address_text[:-2].strip().strip(',')
    elif COUNTRIES[DEFAULT_COUNTRY]['postal_code_regex'].search(address_text):
        country = DEFAULT_COUNTRY
        fields['Country'] = country
        remainder = address_text.strip().strip(',')
    else:
        raise Exception(f'Parse error on "{address_text}".')

    for field in COUNTRIES[country]['fields'][::-1]:
        # Fill out the country's address schema parsing the address text from
        # the right
        if field == 'Zip/Postal Code':
            # Advance to last postal code match (sometimes things like
            # five-digit street addresses will break the naive regex match)
            regex = COUNTRIES[country]['postal_code_regex']
            for postcode_match in re.finditer(regex, remainder):
                pass

            postcode_loc = postcode_match.span()
            fields[field] = postcode_match.group().strip()

            assert remainder[postcode_loc[1]:].strip().strip(',') == ''
            remainder = remainder[:postcode_loc[0]].strip().strip(',')

        elif field == 'State/Province':
            # Assumes all state/province codes are all-caps abbreviations
            regex = re.compile(r'[A-Z]+')
            for prov_match in re.finditer(regex, remainder):
                pass

            prov_loc = prov_match.span()
            fields[field] = prov_match.group().strip()

            assert remainder[prov_loc[1]:].strip().strip(',') == ''
            remainder = remainder[:prov_loc[0]].strip().strip(',')

        else:
            loc = remainder.rfind(',')
            fields[field] = remainder[loc + 1:].strip()
            remainder = remainder[:loc].strip().strip(',')

    street_address_parts = remainder.split(',')
    if len(street_address_parts) == 1:
        for unit_marker in ['Apt', 'Apartment', 'Unit', '#']:
            loc = remainder.find(unit_marker)
            if loc > -1:
                fields['Street Line 1'] = remainder[:loc].strip()
                fields['Street Line 2'] = remainder[loc:].strip()
                break

        if 'Street Line 1' not in fields:
            fields['Street Line 1'] = remainder.strip().strip(',')

    elif len(street_address_parts) == 2:
        fields['Street Line 1'] = street_address_parts[0].strip()
        fields['Street Line 2'] = street_address_parts[1].strip()
    else:
        raise Exception

    return fields

wb = openpyxl.load_workbook(filename=argv[1], data_only=True)
results = []
assert wb['outgoing']['C1'].value == 'name'
assert wb['outgoing']['D1'].value == 'address'
assert wb['outgoing']['E1'].value == 'shipped'

ir = wb['outgoing'].iter_rows()
next(ir)  # Header row
for row in ir:
    size = row[1].value

    name = row[2].value
    address = row[3].value
    shipped = row[4].value

    if shipped != 'Y':
        if re.compile(r'[0-9]XL').search(size):
            warnings.warn('Shirts bigger than XL may need a bigger box.')
        fields = shippo_details(size, name, address)

        if fields is not None:
            results.append(fields)

with open(argv[2], 'w') as cf:
    writer = csv.DictWriter(cf, fieldnames=SHIPPO_FIELDS.keys())

    writer.writeheader()
    for result in results:
        writer.writerow(result)

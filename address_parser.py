#! /usr/bin/env python3

import re
from sys import argv

import openpyxl

DEFAULT_COUNTRY = 'US'
COUNTRIES = {
    'US': {'fields': ['City', 'State/Province', 'Zip/Postal Code'],
           'regexes': {
               'Zip/Postal Code': re.compile(r'[0-9]{5}(-[0-9]{4})?'),
               'State/Province': re.compile(r'[A-Z][A-Z]')}},
    'FR': {'fields': ['Zip/Postal Code', 'City'],
           'regexes': {
               'Zip/Postal Code': re.compile(r'[0-9]{5}')}},
    'GB': {'fields': ['City', 'Zip/Postal Code'],
           'regexes': {
               'Zip/Postal Code': re.compile(
                   r'[A-Z][A-Z0-9]{1,3} [0-9][A-Z]{2}')}},
    'IE': {'fields': ['City', 'State/Province', 'Zip/Postal Code'],
           'regexes': {
               'Zip/Postal Code': re.compile(
                   r'[A-Z][0-9][0-9W] [A-Z0-9]{4}'),
               'State/Province': re.compile(
                   r'County [A-Za-z]*|Co\. [A-Za-z]*')}}
    }

def _filter_notes_to_self(address_text, fxn, default_country=DEFAULT_COUNTRY,
        all_country_metadata=COUNTRIES):
    '''Return None when address_text is a note to self (no commas).'''
    if address_text.find(',') == -1:
        return None
    else:
        return fxn(address_text, default_country, all_country_metadata)

def _country(address_text, default_country=DEFAULT_COUNTRY,
        all_country_metadata=COUNTRIES):
    '''Determine country from address text.

    In our format we demand either an address in the default_country, matching
    that country's regexes for address fields, or qualify an address with ",
    XX" at the end, where XX is an ISO 3166-1 alpha-2 country code.
    '''
    possible_country = address_text.split(',')[-1].strip()
    default_country_regexes = [regex for _, regex in
            all_country_metadata[default_country]['regexes'].items()]

    if re.match(r'[A-Z][A-Z]$', possible_country):
        rest_of_address = address_text.rsplit(',', maxsplit=1)[0]
        return possible_country, rest_of_address
    elif all(re.search(regex, address_text)
            for regex in default_country_regexes):
        return default_country, address_text
    else:
        raise(Exception('Could not parse country for ' + address_text + '.'))

def _tokenize(address_text, country_metadata):
    '''Break up addresses on commas, unit of building designators, and the
    country's regexes.'''
    boundaries = [0, len(address_text)]
    # Commas, before and after
    for match in re.finditer(',', address_text):
        boundaries.append(match.span()[0])
        boundaries.append(match.span()[1])

    # Unit-of-building designators, before
    # TODO Should this be localized too? This is English speaking stuff.
    for unit_marker in ['Apt', 'Apartment', 'Unit', '#']:
        for match in re.finditer(unit_marker, address_text):
            boundaries.append(match.span()[0])

    # State/province and postal code boundaries, localized via regexes, before
    # and after. We do the last such match, to avoid parsing (say) a five-digit
    # street address as a US zipcode.
    for _, regex in country_metadata['regexes'].items():
        for match in re.finditer(regex, address_text):
            # We do the last such match, to avoid parsing (say) a five-digit
            # street address as a US zipcode.
            pass

        boundaries.append(match.span()[0])
        boundaries.append(match.span()[1])

    sb = sorted(boundaries)

    result = []
    for substring in [address_text[sb[i]:sb[i+1]] for i in range(len(sb) - 1)]:
        # Filter for substrings containing alphanumeric data (vs. solely
        # garbage like punctuation or whitespace)
        if re.search(r'[A-Za-z0-9]', substring):
            result.append(substring.strip())

    return result

def _parse(address_text, default_country=DEFAULT_COUNTRY,
        all_country_metadata=COUNTRIES):
    country, unqualified_address = \
            _country(address_text, default_country, all_country_metadata)
    tokens = _tokenize(unqualified_address, all_country_metadata[country])

    fields = all_country_metadata[country]['fields']
    regexes = all_country_metadata[country]['regexes']

    result = {'Country': country}
    # Work from the end: street address part can have variable fields
    offset = len(tokens) - len(fields)
    assert offset >= 1

    result['Street Line 1'] = tokens[0]
    if offset > 1:
        result['Street Line 2'] = ', '.join(tokens[1:offset])

    for i, field in enumerate(fields):
        token = tokens[i + offset]
        if field in regexes:
            assert re.match(regexes[field], token)
        result[field] = token

    return result

def parse(address_text, default_country=DEFAULT_COUNTRY,
        all_country_metadata=COUNTRIES):
    return _filter_notes_to_self(address_text, _parse,
            default_country, all_country_metadata)


if __name__ == '__main__':
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

        # Omit recipients that are unconfirmed (and thus have no shipping address
        # and size)
        if all(i is not None for i in (name, address)):
            print(name)
            print(address)
            print('=>', parse(address))
            print('---')
            print()

import openpyxl
import pprint
import pandas as pd
from operator import itemgetter

# We both want to read and write to the spreadsheet
WORKBOOK_NAME = "ProperVildleder2019.xlsx"
RANGES = 'A4:H'
NUM_COLS = 8
SHEETS = ['Tema', 'Mad']


def get_unique_el(elements, sort_by=(0, 1)):
    """Returns a list of unique list, sorted by the
    n'th element of the lists. sort_by supports tuples and ints.
    """
    uniques = [list(x) for x in set(tuple(x) for x in elements)]
    sort = sorted(uniques, key=itemgetter(*sort_by))
    return sort


def propagate_down(values, columnID=0):
    # Propagates the value of a given cell down through empty cells, in a given
    # column
    for i in range(len(values)-1):
        if values[i+1][columnID] is None:
            values[i+1][columnID] = values[i][columnID]
    return values


def prepare_sheet_results(values, cols=[0, 1, 5, 6, 7]):
    # Values are returned as a list of lists, each of which contain the
    # cell contents. Default columns are
    # - Category
    # - receipt number
    # - price
    # - payment method
    # - name

    # Exclude empty rows - they have all None values.
    values = [x for x in values if x != [None]*NUM_COLS]

    # propagate down the category value
    values = propagate_down(values, 0)

    # Keep only certain columns.
    usable_values = [[row[i] for i in cols] for row in values]

    # Combine category and receipt number - pad with a single 0
    new_cat = [f'{row[0]}{int(row[1]):02d}' for row in usable_values]

    # include the new category in the listings (category, price, payment, name)
    combined = [[z[0], *z[1][2:]] for z in zip(new_cat, usable_values)]

    return combined


def get_values_from_sheet(sheet):
    # Find the max row count of the sheet and get the cells
    maxrow = sheet.max_row
    cells = sheet[f"{RANGES}{sheet.max_row}"]

    # Extract the raw values of each cell and store it in a list of lists
    values = [[i.value for i in rows] for rows in cells]

    # prep the values and return the sheet
    prepped_values = prepare_sheet_results(values)
    return prepped_values


def combine_sheets(sheet_lists):
    combined = []
    for sheet in sheet_lists:
        combined = combined + sheet
    return combined


def process_data(sheet):
    # The vildleder has the structure of vildleder -> name -> payment method ->
    # receipt numbers and price
    vildleder = {}
    names = []
    methods = ["kontokort", "vejlederkort", "personlig"]

    # Get the names
    names = [row[-1] for row in sheet]
    unique_names = list(set(names))

    # Create the dictionary template
    for name in unique_names:
        vildleder[name] = {method: {"kvitteringer": [], "beloeb": 0}
                           for method in methods}

    # populate the dictionary
    for row in sheet:
        receipt, price, method, name = row
        method = methods[int(method)]
        vildleder[name][method]['beloeb'] += price
        vildleder[name][method]['kvitteringer'].append(receipt)

    return vildleder


def main():
    wb = openpyxl.load_workbook(WORKBOOK_NAME)
    sheet = wb['Mad']
    values = get_values_from_sheet(sheet)
    vildleder = process_data(values)

    pprint.pprint(vildleder)


if __name__ == '__main__':
    main()

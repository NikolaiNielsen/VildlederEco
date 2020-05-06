"""
Forfatter: Nikolai Nielsen (NikolaiNielsen@outlook.dk)
Et script til at hjælpe med at gøre Vildlederøkonomien lidt nemmere:
- Gennemgår alle arkene set i variablen "SHEETS", finder alle indtastninger, og
  laver en oversigt over, hvad hver person har udlagt, under hvilken kategori,
  og med hvilken betalingsmetode.

Fremtidige forhåbninger:
- Automatisk genkendelse af, hvor de relevante data er, i hvert ark, så det er
  mere robust
- Grafisk brugerflade for dataindtastning
- Grafisk brugerflade til valg af fil

BRUG AF DETTE PROGRAM
- For at benytte dette program skal du have en lokal kopi af regnskabet i
  xlsx-format. Du skal ændre "WORKBOOK_NAME" til den relative sti til
  regnearket. Hvis regnearket er i samme mappe som dette program, kan du bare
  sætte WORKBOOK_NAME til at være filens navn (husk fil-typen!).
- Du skal ændre "SHEETS" til en liste af strings over hvilke ark, der skal
  tjekkes igennem. Som regel behøver denne ikke at blive ændret fra de normale,
  med mindre, der bliver ændret kategorier af øko-gruppen efter dette program
  er oprettet.
- Programmet forventer at data er indtastet fra kolonne A til H, og fra række 4
  og nedad. Det forventes af kolonnerne er som følger:
  - A: Bilagskode - kategorien af indkøbet. eksempelvis "Tema"
  - B: Bilagsnummer - Indkøbets nummer. Starter med 01, så 02 og så videre.
                      Sammen med bilagskoden udgør dette en samlet "ID" for
                      købet, eksempelvis Tema01.
  - C: NBI Rekvireringsnummer - som regel ikke vigtig. Bruges ikke her
  - D: Tekst - beskrivelse af købet
  - E: Status på bilag - Hvor ligger kvitteringen henne? (den skal helst i
                         mappen eller digitalt på dropbox eller lign.)
  - F: Beløb - hvor meget har dette indkøb kostet?
  - G: metode - hvordan indkøbet er foretaget. Der er 3 måder pt, set under
                "PAYMENT_METHODS". I arket forventes det at dette er et heltal,
                svarende til positionen i PAYMENT_METHODS (så 0=kontokort, etc)
  - H: Navn - hvem har foretaget købet? Der skelnes mellem store og små
              bogstaver!
"""

import openpyxl
import pprint
from decimal import Decimal
from operator import itemgetter

WORKBOOK_NAME = "ProperVildleder2019.xlsx"
SHEETS = ['Tema', 'Mad']
PAYMENT_METHODS = ["kontokort", "vejlederkort", "personlig"]

RANGES = 'A4:H'
NUM_COLS = 8


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
    values = [[i.internal_value for i in rows] for rows in cells]

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
    # Get the names
    names = [row[-1] for row in sheet]
    unique_names = list(set(names))

    # Create the dictionary template
    for name in unique_names:
        vildleder[name] = {method: {"kvitteringer": [], "beloeb": Decimal(0)}
                           for method in PAYMENT_METHODS}

    # populate the dictionary
    for row in sheet:
        receipt, price, method, name = row
        method = PAYMENT_METHODS[int(method)]
        vildleder[name][method]['beloeb'] += round(Decimal(price), 2)
        vildleder[name][method]['kvitteringer'].append(receipt)

    return vildleder


def main():
    wb = openpyxl.load_workbook(WORKBOOK_NAME)

    sheets = []
    for name in SHEETS:
        sheet = wb[name]
        values = get_values_from_sheet(sheet)
        sheets.append(values)

    sheet = combine_sheets(sheets)

    vildleder = process_data(sheet)

    pprint.pprint(vildleder)


if __name__ == '__main__':
    main()

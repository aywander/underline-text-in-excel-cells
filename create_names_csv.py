# This script reads a simple two-column text file of member names
# and outputs a larger variety of name arrangements in csv format
# which can be used with a VB script to underline these names in
# an Excel sheet.
#   The two-column data is separated by a comma.
# Left column, kanji/kana name <Surname> <First name> (space between surname and first name)
# Right column romaji name <First names> <Surname> (space between first name and surname)
# Middle names possible, but only for romaji name.
# E.g.
# ----- Begin file -----
# 佐藤 健二, Kenji Sato
# ワーグナー アレックス, Alex Y. Wagner
# フィン ラセル, Finn Edgar Russel
# -----  End file  -----

import re
import csv

def create_names_ja(name):

    # Last name, First name
    ln, fn = re.split('[ 　]+', name)

    return [f'{ln} {fn}', f'{ln}{fn}', f'{ln}　{fn}']

def create_names_en(name):

    # Last name, First name, Middle name
    name = re.split('[ ]+', name)
    fn, ln, mn = name[0], name[-1], name[1:-1]

    # Process middle names
    if mn:

        # Only use initial for middle names
        mn = [f'{m[0]}.' for m in mn]

        # Get rid of duplicate dots introduced by above
        mn = [' ' + re.sub('[.]+', '.', m) for m in mn]

        mn = ''.join(mn)

    else:
        mn = ''

    # All common combinations. Note there is a space before mn
    return [f'{ln} {fn}', f'{fn} {ln}',
            f'{ln} {fn[0]}.', f'{fn[0]}. {ln}',
            f'{ln}, {fn[0]}.', f'{ln}, {fn}',
            f'{ln} {fn}{mn}', f'{ln} {fn[0]}.{mn}',
            f'{ln}, {fn}{mn}', f'{ln}, {fn[0]}.{mn}',
            f'{fn}{mn} {ln}', f'{fn[0]}.{mn} {ln}',
            ]

# Read text file of member names
with open('member_names.txt', 'r') as fh:
    content = fh.readlines()
    content = [[ n.strip() for n in c.split(',')] for c in content]

# Write csv file
with open(f'member_names.csv', 'w', newline='') as csv_file:
    for c in content:

        print(f'Processing {c[0]} & {c[1]} ...')

        writer = csv.writer(csv_file, dialect='excel')
        names_ja = create_names_ja(c[0])
        names_en = create_names_en(c[1])

        writer.writerow(names_ja + names_en)

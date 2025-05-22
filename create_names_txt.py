# This script reads the yaml files of members from the TAG webpage
# and creates a simple text file with all the names that can then be used by
# create_names_csv.py

from pathlib import Path
from ruamel.yaml import YAML

do_selected = False
selected_positions = True

out_file_name = 'member_names_phd.txt'
positions_selection = ['professor', 'associate_professor', 'lecturer', 'assistant_professor',
                       'postdoc', 'postdoc_student']

# out_file_name = 'member_names_phd.txt'
# positions_selection = ['d6', 'd5', 'd4', 'd3', 'd2', 'd1']

# out_file_name = 'member_names_mas.txt'
# positions_selection = ['m3', 'm2', 'm1']

members_dir = Path('/Users/ayw/dev/TheoreticalAstrophysicsGroup/_members')

if do_selected:
    files = [
        Path('kirihara_takanobu.html'),
        Path('nakano_yusho.html'),
        Path('ichimura_issei.html'),
        Path('tanaka_rei.html'),
    ]
    files_en = [members_dir / 'en' / f for f in files]
    files_ja = [members_dir / 'ja' / f for f in files]

else:
    files_en = list((members_dir / 'en').glob('*.html'))
    files_ja = list((members_dir / 'ja').glob('*.html'))


# Use ruamel.yaml
yaml = YAML()
yaml.preserve_quotes = True
yaml.allow_duplicate_keys = True

with open(out_file_name, 'w') as fh:

    # Assume all arguments are filenames
    for fen, fja in zip(files_en, files_ja):

        print(f'Processing {fen.name}...')
        print(f'           {fja.name}...')
        print(f'')

        # loader expects more after the ending ---
        # so use load_all and only take first item
        # Dump data into language specific dir
        data_en = list(yaml.load_all(fen))[0]
        data_ja = list(yaml.load_all(fja))[0]

        if selected_positions:
            if data_en['position'] in positions_selection:
                fh.write(f'{data_ja["name"]}, {data_en["name"]}\n')

        else:
            fh.write(f'{data_ja["name"]}, {data_en["name"]}\n')

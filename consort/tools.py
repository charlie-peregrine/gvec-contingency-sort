# tools.py, Charlie Jordan, 10/20/2025

import time
from typing import Callable, Iterable
from pathlib import Path
from .contingency import Contingency

SEP_LINE = '/* ' + '='*105 + '\n'

# dump all contingencies to a file
def dump_contingencies(file_name: str | Path, contingencies: Iterable[Contingency],
            sort_func: Callable[[Contingency],tuple] = lambda x: (x.submitter, x.nerc_cat, x.id)):
    with open(file_name, 'w') as file:
        # write file name and time at the top of the file
        file.write(f"/* {Path(file_name).name}, Generated {time.asctime()}\n")
        
        # write category summary at the top of the file
        file.write(SEP_LINE)
        cat_numbers = get_cat_numbers(contingencies)
        file.write('/* Category Count:\n')
        for cat, num in cat_numbers:
            file.write(f'/*\t\t{num:>4} - {cat}\n')

        # keep track of submitters to place contingency grouping comments
        prev_submitter = ''
        for con in sorted(contingencies, key=sort_func):
            if con.submitter != prev_submitter:
                prev_submitter = con.submitter
                file.write(SEP_LINE)
                file.write(
                    f'/* \t\t\t\t\tContingency Definitions Submitted by: {con.submitter}\n'
                )
            file.write(SEP_LINE)
            file.write(con.full_str)
        file.write(SEP_LINE)
        file.write("END\n")

def get_cat_numbers(ls: Iterable[Contingency]):
    cat_numbers = {}
    for con in ls:
        if con.nerc_cat not in cat_numbers:
            cat_numbers[con.nerc_cat] = 1
        else:
            cat_numbers[con.nerc_cat] += 1
    return tuple(sorted(cat_numbers.items()))

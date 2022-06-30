## Adapted: https://www.geeksforgeeks.org/how-to-iterate-through-excel-rows-in-python/
## https://stackoverflow.com/questions/22613272/how-to-access-the-real-value-of-a-cell-using-the-openpyxl-module-for-python
## User interaction should be implemented
## https://www.nature.com/articles/nmeth.4104
# Will be integrated into Indel searcher

import os

# import module
import pandas as pd

# load excel with its path
df = pd.read_excel("AandT.xlsx")

PATH = "./References"

if not os.path.exists(PATH):
    os.makedirs(PATH)

INDEX_COL = "INDEX"
WIDE_TARGET_COL = "Wide target sequence"
RP_BINDING_COL = "RP binding site"
TYPE_COL = "Nuc type"

target_path = os.path.join(PATH, "Target_region.txt")

with open(target_path, 'w') as target:
    # iterate through excel and display data
    for i, row in df.iterrows():
        print("\n")
        print("Row ", i, " data :")
        wide_target_sequence = row[WIDE_TARGET_COL]
        rev_homology_sequence = row[RP_BINDING_COL]

        chunk = wide_target_sequence + rev_homology_sequence
        print(chunk, end=" ")

        type = row[TYPE_COL]
        category = row[INDEX_COL]

        if type == "AsCas12f" and category == "AsFullMatch":
            UP_CONTEXT = 34
            PAM_LENGTH = 4
            WEDGE_POS = 20
            WINDOW_SIZE = 10
        elif type == "AsCas12f":
            UP_CONTEXT = 4
            PAM_LENGTH = 4
            WEDGE_POS = 20
            WINDOW_SIZE = 10

        elif type == "TnpB" and category == "TFullMatch":
            UP_CONTEXT = 30
            PAM_LENGTH = 5
            WEDGE_POS = 18
            WINDOW_SIZE = 10

        elif type == "TnpB":
            UP_CONTEXT = 4
            PAM_LENGTH = 5
            WEDGE_POS = 18
            WINDOW_SIZE = 10

        # Save as files
        target.write(f"{chunk[UP_CONTEXT: UP_CONTEXT + PAM_LENGTH + WEDGE_POS + WINDOW_SIZE]}\n")

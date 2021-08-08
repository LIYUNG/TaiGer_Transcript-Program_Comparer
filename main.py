# import numpy as np
import sys
import pandas as pd
from sorter_function import *
program_idx = []


df_programs_selection = pd.read_excel(
    'C:/Users/steve/Downloads/Course_Sorting_Test/Programs.xlsx')

if df_programs_selection.columns[0] != 'Program' or df_programs_selection.columns[1] != 'Choose':
    print("Error: Please check the program selection xlsx file.")
    sys.exit()

for idx, choose in enumerate(df_programs_selection['Choose']):
    if(choose == 'Yes'):
        program_idx.append(idx)
func(program_idx)

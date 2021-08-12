# import numpy as np
import sys, os
import pandas as pd
from sorter_function import *

if __name__ == "__main__":

    if(len(sys.argv) == 2):
        print("arg: " + sys.argv[1])  # TODO use sys.argv[1] as filename.xlsx
    else:
        print("Error: Please select the transcript excel as argument")
        print("Please type: ")
        print("python main.py <path_to_courses>.xlsx ")
        sys.exit()

    program_idx = []

    program_selection_path = os.getcwd() + '\Programs.xlsx'
    print(os.getcwd())

    df_programs_selection = pd.read_excel(
        program_selection_path)

    if df_programs_selection.columns[0] != 'Program' or df_programs_selection.columns[1] != 'Choose':
        print("Error: Please check the program selection xlsx file.")
        sys.exit()

    for idx, choose in enumerate(df_programs_selection['Choose']):
        if(choose == 'Yes'):
            program_idx.append(idx)
    func(program_idx, sys.argv[1])

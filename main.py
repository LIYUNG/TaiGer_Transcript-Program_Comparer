# import numpy as np
import sys, os
import pandas as pd
from EE_sorter import *
from CS_sorter import *

if __name__ == "__main__":

    if(len(sys.argv) == 3):
        print("arg: " + sys.argv[1])  # TODO use sys.argv[1] as filename.xlsx
    else:
        print("Error: Please select the transcript excel as argument")
        print("Please type: ")
        print("python main.py <path_to_courses>.xlsx <cs/ee>")
        sys.exit()

    program_idx = []
    program_selection_path = ''
    if sys.argv[2] == 'cs':
        program_selection_path = os.getcwd() + '\CS_Programs.xlsx'
        print(os.getcwd())
    elif sys.argv[2] == 'ee':
        program_selection_path = os.getcwd() + '\EE_Programs.xlsx'
        print(os.getcwd())
    else:
        print("Please specify program group: cs ee")
        sys.exit()

    df_programs_selection = pd.read_excel(
        program_selection_path)

    if df_programs_selection.columns[0] != 'Program' or df_programs_selection.columns[1] != 'Choose':
        print("Error: Please check the program selection xlsx file.")
        sys.exit()

    for idx, choose in enumerate(df_programs_selection['Choose']):
        if(choose == 'Yes'):
            program_idx.append(idx)
    if sys.argv[2] == 'cs':
        CS_sorter(program_idx, sys.argv[1])
    elif sys.argv[2] == 'ee':
        EE_sorter(program_idx, sys.argv[1])

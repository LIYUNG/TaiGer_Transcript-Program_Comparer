# import numpy as np
from dotenv import load_dotenv
import sys
import os
import pandas as pd
from database.BiomedicalEngineering.BOE_sorter import BOE_sorter
from database.ElectricalEngineering.EE_sorter import EE_sorter
from database.ComputerScience.CS_sorter import CS_sorter
from database.MechanicalEngineering.ME_sorter import ME_sorter
from database.MaterialsScience.MTL_sorter import MTL_sorter
from database.Psychology.PSY_sorter import PSY_sorter
from database.Management.MGM_sorter import MGM_sorter
from database.DataScience_BusinessIntelligence.DSBI_sorter import DSBI_sorter
from database.TransportationEngineering.TE_sorter import TE_sorter
file_path = os.path.realpath(__file__)
file_path = os.path.dirname(file_path)


print('Before load_dotenv()', os.getenv('MODE'))

load_dotenv()

print('After load_dotenv()', os.getenv('MODE'))


if __name__ == "__main__":

    Pre_Sales_Version = 1
    Premium_Version = 2
    Generated_Version = 0
    if(len(sys.argv) == 3):
        print("premium version")
        # print("arg: " + sys.argv[1])  # TODO use sys.argv[1] as filename.xlsx
        Generated_Version = Premium_Version
    elif(len(sys.argv) == 4):
        print("pre-sales version")
        Generated_Version = Pre_Sales_Version
    else:
        print("Error: Please select the transcript excel as argument")
        print("Please type: ")
        print("python main.py <path_to_courses>.xlsx <cs/ee>")
        sys.exit()

    program_idx = []
    program_selection_path = ''
    program_group_to_file_path = {
        'cs': '/database/ComputerScience/CS_Programs.xlsx',
        'boe': '/database/BiomedicalEngineering/BOE_Programs.xlsx',
        'ee': '/database/ElectricalEngineering/EE_Programs.xlsx',
        'me': '/database/MechanicalEngineering/ME_Programs.xlsx',
        'mgm': '/database/Management/MGM_Programs.xlsx',
        'dsbi': '/database/DataScience_BusinessIntelligence/DSBI_Programs.xlsx',
        'psy': '/database/Psychology/PSY_Programs.xlsx',
        'mtl': '/database/MaterialsScience/MTL_Programs.xlsx',
        'cme': '/database/Materials_Science/CME_Programs.xlsx',
        'te': '/database/TransportationEngineering/TE_Programs.xlsx',
    }
    program_group = sys.argv[2]
    if program_group in program_group_to_file_path:
        program_selection_path = file_path + \
            program_group_to_file_path[program_group]
    else:
        print("Please specify program group: cs ee me")
        sys.exit()

    df_programs_selection = pd.read_excel(
        program_selection_path)

    if df_programs_selection.columns[0] != 'Program' or df_programs_selection.columns[1] != 'Choose':
        print("Error: Please check the program selection xlsx file.")
        sys.exit()

    for idx, choose in enumerate(df_programs_selection['Choose']):
        if(choose == 'Yes'):
            program_idx.append(idx)

    sort_functions = {
        'cs': CS_sorter,
        'boe': BOE_sorter,
        'ee': EE_sorter,
        'me': ME_sorter,
        'mgm': MGM_sorter,
        'dsbi': DSBI_sorter,
        'psy': PSY_sorter,
        'mtl': MTL_sorter,
        'te': TE_sorter,
    }

    if program_group in sort_functions:
        sort_functions[program_group](
            program_idx, sys.argv[1], program_group.upper(), Generated_Version)

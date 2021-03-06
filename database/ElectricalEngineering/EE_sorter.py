import xlsxwriter
from CourseSuggestionAlgorithms import *
from util import *
from database.ElectricalEngineering.EE_KEYWORDS import *
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import pandas as pd
import sys
import os
env_file_path = os.path.realpath(__file__)
env_file_path = os.path.dirname(env_file_path)

# Global variable:
column_len_array = []


def TUM_EI(transcript_sorted_group_map, df_transcript_array, writer):
    print("Create TUM EI sheet")
    # TODO: implement the mapping from the existing courses to program's requirement
    start_row = 0
    for idx, sortedcourses in enumerate(df_transcript_array):
        sortedcourses.to_excel(
            writer, sheet_name='TUM_EI', startrow=start_row, index=False)
        start_row += len(sortedcourses.index) + 2
    print("Save to TUM_EI")


def RWTH_EI(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'RWTH_Aachen_EI'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_ECTS': 28}
    PROG_SPEC_PHYSIK_PARAM = {
        'Program_Category': 'Physics', 'Required_ECTS': 10}
    PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM = {
        'Program_Category': 'Electronics and Circuits Module', 'Required_ECTS': 34}
    PROG_SPEC_PROGRAMMIERUNG_PARAM = {
        'Program_Category': 'Programming and Computer science', 'Required_ECTS': 12}
    PROG_SPEC_SYSTEM_THEORIE_PARAM = {
        'Program_Category': 'System_Theory', 'Required_ECTS': 8}
    PROG_SPEC_THEORETICAL_EECS_EI_PARAM = {
        'Program_Category': 'Theoretical_Module_EECS', 'Required_ECTS': 8}
    PROG_SPEC_ANWENDUNG_MODULE_PARAM = {
        'Program_Category': 'Application_Module_EECS', 'Required_ECTS': 20}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_PHYSIK_PARAM,  # ??????
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # ??????
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # ????????????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_THEORETICAL_EECS_EI_PARAM,  # ??????????????????
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # ????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_PHYSIK_PARAM,  # ??????
        PROG_SPEC_PHYSIK_PARAM,  # ????????????
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # ??????
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # ??????
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # ????????????
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # ????????????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ??????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ????????????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ??????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ???????????????
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # ??????
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # ????????????
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # ??????
        PROG_SPEC_OTHERS,  # ?????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_THEORETICAL_EECS_EI_PARAM,  # ??????????????????
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # ??????????????????
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # ????????????
        PROG_SPEC_OTHERS,  # ??????,??????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)


def STUTTGART_EI(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'STUTTGART_EI'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    # df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_ECTS': 24}
    PROG_SPEC_PHY_EXP_PARAM = {
        'Program_Category': 'Physics Experiment', 'Required_ECTS': 6}
    PROG_SPEC_MICROELECTRONICS_PARAM = {
        'Program_Category': 'Microelectronics', 'Required_ECTS': 9}
    PROG_SPEC_INTRO_ELECT_ENG_PROJ_PARAM = {
        'Program_Category': 'Intro. Electrical Engineering and project', 'Required_ECTS': 9}
    PROG_SPEC_INTRO_PROGRAMMING_ENG_PARAM = {
        'Program_Category': 'Intro. Programming and project', 'Required_ECTS': 6}
    PROG_SPEC_INTRO_SOFTWARE_SYSTEM_PARAM = {
        'Program_Category': 'Intro. Software System', 'Required_ECTS': 3}
    PROG_SPEC_ENERGIETECHNIK_PARAM = {
        'Program_Category': 'Energy Technique', 'Required_ECTS': 9}
    PROG_SPEC_SCHALTUNGSTECHNIK_PARAM = {
        'Program_Category': 'Circuits Technology', 'Required_ECTS': 9}
    PROG_SPEC_ELEKTRODYNAMIK_PARAM = {
        'Program_Category': 'Electromagnetics', 'Required_ECTS': 9}
    PROG_SPEC_NACHRICHTENTECHNIK_PARAM = {
        'Program_Category': 'Communication Engineering', 'Required_ECTS': 9}
    # Grundlagen der Informationsverarbeitung
    PROG_SPEC_INTRO_INFOR_VERARBEITUNG_PARAM = {
        'Program_Category': 'Intro. Information processing', 'Required_ECTS': 6}
    PROG_SPEC_SIGNAL_SYSTEM_PARAM = {
        'Program_Category': 'Signals and Systems', 'Required_ECTS': 6}
    PROG_SPEC_SCHWERPUNKTE_PARAM = {        # ?????????????????????????????????????????????????????????????????????Technische Informatik???????????????????????????????????????????????????
        'Program_Category': 'Advanced Modules', 'Required_ECTS': 6}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_PHY_EXP_PARAM,    # ????????????
        PROG_SPEC_MICROELECTRONICS_PARAM,  # ?????????
        PROG_SPEC_INTRO_ELECT_ENG_PROJ_PARAM,  # ??????????????????
        PROG_SPEC_INTRO_PROGRAMMING_ENG_PARAM,  # ?????????????????????
        # ?????????????????? Objektorientierung, Design pattern ??????????????????, UML
        PROG_SPEC_INTRO_SOFTWARE_SYSTEM_PARAM,
        PROG_SPEC_ENERGIETECHNIK_PARAM,         # ????????????
        PROG_SPEC_SCHALTUNGSTECHNIK_PARAM,      # ?????????
        PROG_SPEC_ELEKTRODYNAMIK_PARAM,         # ?????????
        PROG_SPEC_NACHRICHTENTECHNIK_PARAM,     # ????????????
        PROG_SPEC_INTRO_INFOR_VERARBEITUNG_PARAM,   # ???????????? ???????????? ????????? ????????????
        PROG_SPEC_SIGNAL_SYSTEM_PARAM,          # ???????????????
        PROG_SPEC_SCHWERPUNKTE_PARAM,           # ????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_PHY_EXP_PARAM,  # ??????
        PROG_SPEC_PHY_EXP_PARAM,  # ????????????
        PROG_SPEC_INTRO_INFOR_VERARBEITUNG_PARAM,  # ??????
        PROG_SPEC_INTRO_PROGRAMMING_ENG_PARAM,  # ??????
        PROG_SPEC_INTRO_SOFTWARE_SYSTEM_PARAM,  # ????????????
        PROG_SPEC_SCHWERPUNKTE_PARAM,  # ????????????
        PROG_SPEC_MICROELECTRONICS_PARAM,  # ??????
        PROG_SPEC_INTRO_ELECT_ENG_PROJ_PARAM,  # ????????????
        PROG_SPEC_SCHALTUNGSTECHNIK_PARAM,  # ??????
        PROG_SPEC_SIGNAL_SYSTEM_PARAM,  # ????????????
        PROG_SPEC_ELEKTRODYNAMIK_PARAM,  # ??????
        PROG_SPEC_ENERGIETECHNIK_PARAM,  # ????????????
        PROG_SPEC_NACHRICHTENTECHNIK_PARAM,  # ??????
        PROG_SPEC_SCHWERPUNKTE_PARAM,  # ?????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_ELEKTRODYNAMIK_PARAM,  # ??????????????????
        PROG_SPEC_SCHWERPUNKTE_PARAM,  # ??????????????????
        PROG_SPEC_SCHWERPUNKTE_PARAM,  # ????????????
        PROG_SPEC_OTHERS,  # ??????,??????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)


# FPSO: https://portal.mytum.de/archiv/kompendium_rechtsangelegenheiten/fachpruefungsordnungen/2018-56-Neufassung-FPSO-MA-CE-FINAL.pdf
# https://www.ei.tum.de/fileadmin/tueifei/www/Studium_Satzungen_Modullisten_Studienfuehrer/BScEI_Stu_richt_empf.pdf
# https://www.ei.tum.de/fileadmin/tueifei/www/Studium_Satzungen_Modullisten_Studienfuehrer/Modulliste_BSEI_PO20181.pdf
def TUM_MSCE(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_MSCE'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    # df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'H??here_Mathematik', 'Required_ECTS': 30}
    PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Elektrotechnik', 'Required_ECTS': 66}
    PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Kommunikationstechnik', 'Required_ECTS': 30}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # ????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ?????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????,??????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)


def TUM_MSPE(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_MSPE'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    # df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'H??here_Mathematik', 'Required_ECTS': 30}
    # Grundlagen der Elektrotechnik, Vertiefung Energietechnik
    # Schaltungstechnik, Elektrische Felder und Wellen,Festk??rperphysik und
    # Bauelemente, Hochspannungstechnik und Energie-??bertragungstechnik,
    # elektrische Maschinen, etc.
    PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Elektrotechnik', 'Required_ECTS': 45}
    # Grundlagen des Maschinenwesens, Vertiefung Energietechnik
    # (Technische Mechanik, Thermodynamik, Str??mungsmechanik,
    # W??rme-und Stoff??bertragung, Maschinendynamik, etc.)
    PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM = {
        'Program_Category': 'Grundlagen_Maschinenwesen', 'Required_ECTS': 45}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM,  # ????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ?????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # ??????????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM,  # ??????,????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)

# TODO: to finish it/ or move this program to biology related
# https://portal.mytum.de/archiv/kompendium_rechtsangelegenheiten/fachpruefungsordnungen/2020-98-FPSO-Ma-Neuroengineering-FINAL-22-12-2020.pdf/download


def TUM_MSNE(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_MSNE'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    # df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'H??here_Mathematik', 'Required_ECTS': 32}
    # Naturwissenschaftliche Grundlagen (Physik, Biochemie, Neuroscience)
    PROG_SPEC_GRUNDLAGE_NATUR_PARAM = {
        'Program_Category': 'Natural Science (Physics, Biochem., neuroscience', 'Required_ECTS': 45}
    # Bio-und Medizintechnische Ingenieurgrundlagen oder Psychologie
    PROG_SPEC_GRUNDLAGE_BIO_PARAM = {
        'Program_Category': 'Bio and medical engineering', 'Required_ECTS': 40}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_NATUR_PARAM,  # ???????????? ?????? ????????????
        PROG_SPEC_GRUNDLAGE_BIO_PARAM,  # ????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_NATUR_PARAM,  # ??????
        PROG_SPEC_GRUNDLAGE_NATUR_PARAM,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ???????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????
        PROG_SPEC_OTHERS,  # ?????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????????????????
        PROG_SPEC_OTHERS,  # ??????????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????,????????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)


# Requirement: https://www.tuhh.de/t3resources/tuhh/download/studium/studieninteressierte/Fachspezifische_Kenntnisse_Master/Fachspezifische-Anforderung-2016-MM.pdf
# https://www.tuhh.de/alt/tuhh/education/degree-courses/international-study-programs/how-and-when-to-apply/specific-requirements.html

def TUHH_MICROELECTRONICS(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUHH_MICROELECTRONICS'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp = []
    df_category_courses_sugesstion_data_temp = []
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    # df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_ECTS': 30}
    PROG_SPEC_COMPUTER_SCIENCE_ENG_PARAM = {
        'Program_Category': 'Computer Science', 'Required_ECTS': 18}
    PROG_SPEC_SYSTEM_THEORIE_PARAM = {
        'Program_Category': 'Control Theory', 'Required_ECTS': 6}
    PROG_SPEC_PHY_EXP_PARAM = {
        'Program_Category': 'Physics', 'Required_ECTS': 6}
    PROG_SPEC_ELECTRICAL_ENG_PARAM = {  # Electrical engineering(direct/alternating current, electronics)
        'Program_Category': 'Fundamental Electrical Engineering', 'Required_ECTS': 12}
    PROG_SPEC_ELECTRICAL_MTL_PARAM = {  # Materials in electrical engineering,
        'Program_Category': 'Materials in Electrical Engineering', 'Required_ECTS': 3}
    PROG_SPEC_METHOD_DATA_PROCESSING_PARAM = {
        'Program_Category': 'Measurements: Methods and data processing', 'Required_ECTS': 3}
    PROG_SPEC_CIRCUIT_THEORY_PARAM = {
        'Program_Category': 'Circuit theory', 'Required_ECTS': 6}
    PROG_SPEC_TRANS_LINE_PARAM = {  # Transmission Line
        'Program_Category': 'Transmission Line', 'Required_ECTS': 6}
    PROG_SPEC_SIGNAL_SYSTEM_PARAM = {
        'Program_Category': 'Signals and systems', 'Required_ECTS': 6}
    PROG_SPEC_THEORY_ELECTRICAL_ENG_PARAM = {
        'Program_Category': 'Theoretical Electrical Engineering', 'Required_ECTS': 12}
    PROG_SPEC_SEMICONDUCTOR_CIRCUIT_PARAM = {
        'Program_Category': 'Semiconductor and electronics devices', 'Required_ECTS': 6}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_ECTS': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_COMPUTER_SCIENCE_ENG_PARAM,  # ????????????
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # ????????????
        PROG_SPEC_PHY_EXP_PARAM,           # ??????
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # ??????????????????
        PROG_SPEC_ELECTRICAL_MTL_PARAM,  # ????????????
        PROG_SPEC_METHOD_DATA_PROCESSING_PARAM,  # ????????????
        PROG_SPEC_CIRCUIT_THEORY_PARAM,  # ?????????
        PROG_SPEC_TRANS_LINE_PARAM,     # ????????? ??????
        PROG_SPEC_SIGNAL_SYSTEM_PARAM,  # ????????????
        PROG_SPEC_THEORY_ELECTRICAL_ENG_PARAM,  # ?????????
        PROG_SPEC_SEMICONDUCTOR_CIRCUIT_PARAM,  # ????????????????????? ???????????? aka?????????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # ?????????
        PROG_SPEC_MATH_PARAM,  # ??????
        PROG_SPEC_PHY_EXP_PARAM,  # ??????
        PROG_SPEC_PHY_EXP_PARAM,  # ????????????
        PROG_SPEC_COMPUTER_SCIENCE_ENG_PARAM,  # ??????
        PROG_SPEC_COMPUTER_SCIENCE_ENG_PARAM,  # ??????
        PROG_SPEC_COMPUTER_SCIENCE_ENG_PARAM,  # ????????????
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # ????????????
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # ??????
        PROG_SPEC_METHOD_DATA_PROCESSING_PARAM,  # ????????????
        PROG_SPEC_CIRCUIT_THEORY_PARAM,  # ??????
        PROG_SPEC_SIGNAL_SYSTEM_PARAM,  # ????????????
        PROG_SPEC_THEORY_ELECTRICAL_ENG_PARAM,  # ??????
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # ????????????
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # ??????
        PROG_SPEC_SEMICONDUCTOR_CIRCUIT_PARAM,  # ?????????
        PROG_SPEC_ELECTRICAL_MTL_PARAM,  # ????????????
        PROG_SPEC_TRANS_LINE_PARAM,  # ??????????????????
        PROG_SPEC_OTHERS,  # ??????????????????
        PROG_SPEC_OTHERS,  # ????????????
        PROG_SPEC_OTHERS,  # ??????,??????
        PROG_SPEC_OTHERS  # ??????
    ]

    # Development check
    if len(program_category_map) != len(df_transcript_array):
        print("program_category_map size: " + str(len(program_category_map)))
        print("df_transcript_array size:  " + str(len(df_transcript_array)))
        print("Please check the number of program_category_map again!")
        sys.exit()

    #####################################################################
    ####################### End #########################################
    #####################################################################

    WriteToExcel(writer, program_name, program_category, program_category_map,
                 transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array)


program_sort_function = [TUM_EI,
                         RWTH_EI,
                         STUTTGART_EI,
                         TUM_MSCE,
                         TUM_MSPE,
                         TUM_MSNE,
                         TUHH_MICROELECTRONICS]


def EE_sorter(program_idx, file_path, abbrev, Generated_Version):
    print('After load_dotenv()', os.getenv('MODE'))

    basic_classification_en = {
        '?????????': [EE_CALCULUS_KEY_WORDS_EN, EE_CALCULUS_ANTI_KEY_WORDS_EN, ['???', '???']],
        '??????': [EE_MATH_KEY_WORDS_EN, EE_MATH_ANTI_KEY_WORDS_EN],
        '??????': [EE_PHYSICS_KEY_WORDS_EN, EE_PHYSICS_ANTI_KEY_WORDS_EN, ['???', '???']],
        '????????????': [EE_PHYSICS_EXP_KEY_WORDS_EN, EE_PHYSICS_EXP_ANTI_KEY_WORDS_EN, ['???', '???']],
        '??????': [EE_INTRO_COMPUTER_SCIENCE_KEY_WORDS_EN, EE_INTRO_COMPUTER_SCIENCE_ANTI_KEY_WORDS_EN],
        '??????': [EE_PROGRAMMING_KEY_WORDS_EN, EE_PROGRAMMING_ANTI_KEY_WORDS_EN],
        '????????????': [EE_SOFTWARE_SYSTEM_KEY_WORDS_EN, EE_SOFTWARE_SYSTEM_ANTI_KEY_WORDS_EN],
        '????????????': [EE_CONTROL_THEORY_KEY_WORDS_EN, EE_CONTROL_THEORY_ANTI_KEY_WORDS_EN],
        '??????': [EE_ELECTRONICS_KEY_WORDS_EN, EE_ELECTRONICS_ANTI_KEY_WORDS_EN, ['???', '???']],
        '????????????': [EE_ELECTRONICS_EXP_KEY_WORDS_EN, EE_ELECTRONICS_EXP_ANTI_KEY_WORDS_EN, ['???', '???']],
        '??????': [EE_ELECTRO_CIRCUIT_KEY_WORDS_EN, EE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS_EN, ['???', '???']],
        '????????????': [EE_SIGNAL_SYSTEM_KEY_WORDS_EN, EE_SIGNAL_SYSTEM_ANTI_KEY_WORDS_EN],
        '??????': [EE_ELECTRO_MAGNET_KEY_WORDS_EN, EE_ELECTRO_MAGNET_ANTI_KEY_WORDS_EN, ['???', '???']],
        '????????????': [EE_POWER_ELECTRO_KEY_WORDS_EN, EE_POWER_ELECTRO_ANTI_KEY_WORDS_EN, ['???', '???']],
        '??????': [EE_COMMUNICATION_KEY_WORDS_EN, EE_COMMUNICATION_ANTI_KEY_WORDS_EN, ['???', '???']],
        '?????????': [EE_SEMICONDUCTOR_KEY_WORDS_EN, EE_SEMICONDUCTOR_ANTI_KEY_WORDS_EN],
        '????????????': [EE_ELEC_MATERIALS_KEY_WORDS_EN, EE_ELEC_MATERIALS_ANTI_KEY_WORDS_EN],
        '??????????????????': [EE_HF_RF_THEO_INFO_KEY_WORDS_EN, EE_HF_RF_THEO_INFO_ANTI_KEY_WORDS_EN],
        '??????????????????': [EE_ADVANCED_ELECTRO_KEY_WORDS_EN, EE_ADVANCED_ELECTRO_ANTI_KEY_WORDS_EN],
        '??????????????????': [EE_APPLICATION_ORIENTED_KEY_WORDS_EN, EE_APPLICATION_ORIENTED_ANTI_KEY_WORDS_EN],
        '??????': [EE_MACHINE_RELATED_KEY_WORDS_EN, EE_MACHINE_RELATED_ANTI_KEY_WORDS_EN],
        '??????': [USELESS_COURSES_KEY_WORDS_EN, USELESS_COURSES_ANTI_KEY_WORDS_EN], }

    basic_classification_zh = {
        '?????????': [EE_CALCULUS_KEY_WORDS, EE_CALCULUS_ANTI_KEY_WORDS, ['???', '???']],
        '??????': [EE_MATH_KEY_WORDS, EE_MATH_ANTI_KEY_WORDS],
        '??????': [EE_PHYSICS_KEY_WORDS, EE_PHYSICS_ANTI_KEY_WORDS, ['???', '???']],
        '????????????': [EE_PHYSICS_EXP_KEY_WORDS, EE_PHYSICS_EXP_ANTI_KEY_WORDS, ['???', '???']],
        '??????': [EE_INTRO_COMPUTER_SCIENCE_KEY_WORDS, EE_INTRO_COMPUTER_SCIENCE_ANTI_KEY_WORDS],
        '??????': [EE_PROGRAMMING_KEY_WORDS, EE_PROGRAMMING_ANTI_KEY_WORDS],
        '????????????': [EE_SOFTWARE_SYSTEM_KEY_WORDS, EE_SOFTWARE_SYSTEM_ANTI_KEY_WORDS],
        '????????????': [EE_CONTROL_THEORY_KEY_WORDS, EE_CONTROL_THEORY_ANTI_KEY_WORDS],
        '??????': [EE_ELECTRONICS_KEY_WORDS, EE_ELECTRONICS_ANTI_KEY_WORDS, ['???', '???']],
        '????????????': [EE_ELECTRONICS_EXP_KEY_WORDS, EE_ELECTRONICS_EXP_ANTI_KEY_WORDS, ['???', '???']],
        '??????': [EE_ELECTRO_CIRCUIT_KEY_WORDS, EE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS, ['???', '???']],
        '????????????': [EE_SIGNAL_SYSTEM_KEY_WORDS, EE_SIGNAL_SYSTEM_ANTI_KEY_WORDS],
        '??????': [EE_ELECTRO_MAGNET_KEY_WORDS, EE_ELECTRO_MAGNET_ANTI_KEY_WORDS, ['???', '???']],
        '????????????': [EE_POWER_ELECTRO_KEY_WORDS, EE_POWER_ELECTRO_ANTI_KEY_WORDS, ['???', '???']],
        '??????': [EE_COMMUNICATION_KEY_WORDS, EE_COMMUNICATION_ANTI_KEY_WORDS, ['???', '???']],
        '?????????': [EE_SEMICONDUCTOR_KEY_WORDS, EE_SEMICONDUCTOR_ANTI_KEY_WORDS],
        '????????????': [EE_ELEC_MATERIALS_KEY_WORDS, EE_ELEC_MATERIALS_ANTI_KEY_WORDS],
        '??????????????????': [EE_HF_RF_THEO_INFO_KEY_WORDS, EE_HF_RF_THEO_INFO_ANTI_KEY_WORDS],
        '??????????????????': [EE_ADVANCED_ELECTRO_KEY_WORDS, EE_ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '??????????????????': [EE_APPLICATION_ORIENTED_KEY_WORDS, EE_APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '??????': [EE_MACHINE_RELATED_KEY_WORDS, EE_MACHINE_RELATED_ANTI_KEY_WORDS],
        '??????': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    Classifier(program_idx, file_path, abbrev, env_file_path,
               basic_classification_en, basic_classification_zh, column_len_array, program_sort_function, Generated_Version)

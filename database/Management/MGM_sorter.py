import xlsxwriter
from CourseSuggestionAlgorithms import *
from util import *
from database.Management.MGM_KEYWORDS import *
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import pandas as pd
import sys
import os
env_file_path = os.path.realpath(__file__)
env_file_path = os.path.dirname(env_file_path)

# Global variable:
column_len_array = []

# FPSO: https://www.tum.de/fileadmin/w00bfo/www/Studium/Studienangebot/Lesbare_Fassung/Master/Managem._Techn._LB_AS_3._AS_28052021.pdf


def TUM_MMT(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_MMT'
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

    PROG_SPEC_BWL_PARAM = {
        'Program_Category': 'Betriebswirtschaftliche Module', 'Required_CP': 25}  # 20 PUnkto
    PROG_SPEC_EMPIRIAL_METHODE_PARAM = {
        'Program_Category': 'Empirische Methoden', 'Required_CP': 6}            # 10 Punkte
    PROG_SPEC_OPERATION_RESEARCH_PARAM = {
        'Program_Category': 'Operations Research', 'Required_CP': 6}            # 10 Punkte
    PROG_SPEC_VWL_PARAM = {
        'Program_Category': 'Volkswirtschaftliche Module', 'Required_CP': 10}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_BWL_PARAM,  # 管理
        PROG_SPEC_EMPIRIAL_METHODE_PARAM,
        PROG_SPEC_OPERATION_RESEARCH_PARAM,  # 作業研究
        PROG_SPEC_VWL_PARAM,  # 經濟
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
    PROG_SPEC_OTHERS,  # 微積分
    PROG_SPEC_OTHERS,  # 數學
    PROG_SPEC_VWL_PARAM,  # 經濟
    PROG_SPEC_BWL_PARAM,  # 管理
    PROG_SPEC_OTHERS,  # 會計
    PROG_SPEC_OTHERS,  # 統計
    PROG_SPEC_OTHERS,  # 金融
    PROG_SPEC_OTHERS,  # 行銷
    PROG_SPEC_OPERATION_RESEARCH_PARAM,  # 作業研究
    PROG_SPEC_EMPIRIAL_METHODE_PARAM,  # 觀察研究
    PROG_SPEC_OTHERS,  # 程式
    PROG_SPEC_OTHERS,  # 資料科學
    PROG_SPEC_OTHERS,  # 論文
    PROG_SPEC_OTHERS  # 其他
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


def TUM_CONSUMER_SCIENCE(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name='TUM_CONSUMER_SCIENCE'
    print("Create " + program_name + " sheet")
    df_transcript_array_temp=[]
    df_category_courses_sugesstion_data_temp=[]
    for idx, df in enumerate(df_transcript_array):
        df_transcript_array_temp.append(df.copy())
    for idx, df in enumerate(df_category_courses_sugesstion_data):
        df_category_courses_sugesstion_data_temp.append(df.copy())
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    # Statistik, Empirische Forschungsmethoden, Quantitative Methoden, Mathematik
    PROG_SPEC_EMPIRIAL_METHODE_PARAM={
        'Program_Category': 'BWL, Quantitative Method, Mathematik', 'Required_CP': 15} # 10 PUnkto
    #  Bachelorarbeit, eines Projekts, eines wissenschaftlichen Aufsatzes
    PROG_SPEC_BACHELORARBEIT_PARAM = {
        'Program_Category': 'Bachelor Thesis', 'Required_CP': 5}                # 10 Punkte
    # quantitativen Entscheidungsunterstützung mit Methoden des Operations Research 
    PROG_SPEC_BWL_PARAM = {
        'Program_Category': 'BWL', 'Required_CP': 6}                           # 10 Punkte
    # VWL mind. 5 Credits oder Module aus dem Bereich Consumer Behavior mind. 5 Credits
    PROG_SPEC_VWL_PARAM = {
        'Program_Category': 'Volkswirtschaftliche Module', 'Required_CP': 10}   # 10 Punkte
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_EMPIRIAL_METHODE_PARAM,  # 觀察研究, 研究方法, 量化分析, 數學
        PROG_SPEC_BACHELORARBEIT_PARAM, # 論文
        PROG_SPEC_BWL_PARAM,  # 企業管理
        PROG_SPEC_VWL_PARAM,  # 經濟
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_OTHERS,  # 微積分
        PROG_SPEC_EMPIRIAL_METHODE_PARAM,  # 數學
        PROG_SPEC_VWL_PARAM,  # 經濟
        PROG_SPEC_BWL_PARAM,  # 管理
        PROG_SPEC_OTHERS,  # 會計
        PROG_SPEC_EMPIRIAL_METHODE_PARAM,  # 統計
        PROG_SPEC_OTHERS,  # 金融
        PROG_SPEC_OTHERS,  # 行銷
        PROG_SPEC_OTHERS,  # 作業研究
        PROG_SPEC_EMPIRIAL_METHODE_PARAM,  # 觀察研究
        PROG_SPEC_OTHERS,  # 程式
        PROG_SPEC_OTHERS,  # 資料科學
        PROG_SPEC_BACHELORARBEIT_PARAM,  # 論文
        PROG_SPEC_OTHERS  # 其他
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


program_sort_function = [TUM_MMT, TUM_CONSUMER_SCIENCE]


def MGM_sorter(program_idx, file_path, abbrev):
    basic_classification_en = {
        '微積分': [MGM_CALCULUS_KEY_WORDS_EN, MGM_CALCULUS_ANTI_KEY_WORDS_EN],
        '數學': [MGM_MATH_KEY_WORDS_EN, MGM_MATH_ANTI_KEY_WORDS_EN],
        '經濟': [MGM_ECONOMICS_KEY_WORDS_EN, MGM_ECONOMICS_ANTI_KEY_WORDS_EN],
        '企業': [MGM_BUSINESS_KEY_WORDS_EN, MGM_BUSINESS_ANTI_KEY_WORDS_EN],
        '管理': [MGM_MANAGEMENT_KEY_WORDS_EN, MGM_MANAGEMENT_ANTI_KEY_WORDS_EN],
        '會計': [MGM_ACCOUNTING_KEY_WORDS_EN, MGM_ACCOUNTING_ANTI_KEY_WORDS_EN],
        '統計': [MGM_STATISTICS_KEY_WORDS_EN, MGM_STATISTICS_ANTI_KEY_WORDS_EN],
        '金融': [MGM_FINANCE_KEY_WORDS_EN, MGM_FINANCE_ANTI_KEY_WORDS_EN],
        '行銷': [MGM_MARKETING_KEY_WORDS_EN, MGM_MARKETING_ANTI_KEY_WORDS_EN],
        '作業研究': [MGM_OP_RESEARCH_KEY_WORDS_EN, MGM_OP_RESEARCH_ANTI_KEY_WORDS_EN],
        '觀察研究': [MGM_EP_RESEARCH_KEY_WORDS_EN, MGM_EP_RESEARCH_ANTI_KEY_WORDS_EN],
        '程式': [MGM_PROGRAMMING_KEY_WORDS_EN, MGM_PROGRAMMING_ANTI_KEY_WORDS_EN],
        '資料科學': [MGM_DATA_SCIENCE_KEY_WORDS_EN, MGM_DATA_SCIENCE_ANTI_KEY_WORDS_EN],
        '論文': [MGM_BACHELOR_THESIS_KEY_WORDS_EN, MGM_BACHELOR_THESIS_ANTI_KEY_WORDS_EN],
        '其他': [USELESS_COURSES_KEY_WORDS_EN, USELESS_COURSES_ANTI_KEY_WORDS_EN], }

    basic_classification_zh = {
        '微積分': [MGM_CALCULUS_KEY_WORDS, MGM_CALCULUS_ANTI_KEY_WORDS],
        '數學': [MGM_MATH_KEY_WORDS, MGM_MATH_ANTI_KEY_WORDS],
        '經濟': [MGM_ECONOMICS_KEY_WORDS, MGM_ECONOMICS_ANTI_KEY_WORDS],
        '企業': [MGM_BUSINESS_KEY_WORDS, MGM_BUSINESS_ANTI_KEY_WORDS],
        '管理': [MGM_MANAGEMENT_KEY_WORDS, MGM_MANAGEMENT_ANTI_KEY_WORDS],
        '會計': [MGM_ACCOUNTING_KEY_WORDS, MGM_ACCOUNTING_ANTI_KEY_WORDS],
        '統計': [MGM_STATISTICS_KEY_WORDS, MGM_STATISTICS_ANTI_KEY_WORDS],
        '金融': [MGM_FINANCE_KEY_WORDS, MGM_FINANCE_ANTI_KEY_WORDS],
        '行銷': [MGM_MARKETING_KEY_WORDS, MGM_MARKETING_ANTI_KEY_WORDS],
        '作業研究': [MGM_OP_RESEARCH_KEY_WORDS, MGM_OP_RESEARCH_ANTI_KEY_WORDS],
        '觀察研究': [MGM_EP_RESEARCH_KEY_WORDS, MGM_EP_RESEARCH_ANTI_KEY_WORDS],
        '程式': [MGM_PROGRAMMING_KEY_WORDS, MGM_PROGRAMMING_ANTI_KEY_WORDS],
        '資料科學': [MGM_DATA_SCIENCE_KEY_WORDS, MGM_DATA_SCIENCE_ANTI_KEY_WORDS],
        '論文': [MGM_BACHELOR_THESIS_KEY_WORDS, MGM_BACHELOR_THESIS_ANTI_KEY_WORDS],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    Classifier(program_idx, file_path, abbrev, env_file_path,
               basic_classification_en, basic_classification_zh, column_len_array, program_sort_function)

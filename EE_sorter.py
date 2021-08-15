from numpy import nan
import sys
import os
import pandas as pd
from cell_formatter import red_out_failed_subject
from EE_KEYWORDS import *
from util import *
from alogrithms import *

import xlsxwriter

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
    df_transcript_array_temp = df_transcript_array
    df_category_courses_sugesstion_data_temp = df_category_courses_sugesstion_data
    #####################################################################
    ############## Program Specific Parameters ##########################
    #####################################################################

    # Create transcript_sorted_group to program_category mapping
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Höhere Mathematik', 'Required_CP': 28}
    PROG_SPEC_PHYSIK_PARAM = {
        'Program_Category': 'Physik', 'Required_CP': 10}
    PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM = {
        'Program_Category': 'Elektrotechnik_Schaltungstechnik', 'Required_CP': 34}
    PROG_SPEC_PROGRAMMIERUNG_PARAM = {
        'Program_Category': 'Programmierung', 'Required_CP': 12}
    PROG_SPEC_SYSTEM_THEORIE_PARAM = {
        'Program_Category': 'System_Theorie', 'Required_CP': 8}
    PROG_SPEC_VERTIEFUNG_EI_PARAM = {
        'Program_Category': 'Vertiefung_EI', 'Required_CP': 8}
    PROG_SPEC_ANWENDUNG_MODULE_PARAM = {
        'Program_Category': 'Anwendung_Module', 'Required_CP': 20}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_PHYSIK_PARAM,  # 物理
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # 資訊
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # 控制系統
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # 電子電路電磁
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 電機專業選修
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # 應用科技
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_PHYSIK_PARAM,  # 物理
        PROG_SPEC_PHYSIK_PARAM,  # 物理實驗
        PROG_SPEC_PROGRAMMIERUNG_PARAM,  # 資訊
        PROG_SPEC_SYSTEM_THEORIE_PARAM,  # 控制系統
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # 電子
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # 電路
        PROG_SPEC_ELEKTROTECHNIK_SCHALTUNGSTECHNIK_PARAM,  # 電磁
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 半導體
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 電機專業選修
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # 應用科技
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

    df_PROG_SPEC_CATES, df_PROG_SPEC_CATES_COURSES_SUGGESTION = ProgramCategoryInit(
        program_category)

    transcript_sorted_group_list = list(transcript_sorted_group_map)

    # Courses: mapping the students' courses to program-specific category
    df_PROG_SPEC_CATES = CoursesToProgramCategoryMapping(
        df_PROG_SPEC_CATES, program_category_map, transcript_sorted_group_list, df_transcript_array_temp)

    # Suggestion courses: mapping the sugesstion courses to program-specific category
    df_PROG_SPEC_CATES_COURSES_SUGGESTION = CoursesToProgramCategoryMapping(
        df_PROG_SPEC_CATES_COURSES_SUGGESTION, program_category_map, transcript_sorted_group_list, df_category_courses_sugesstion_data_temp)

    # append 總學分 for each program category
    for idx, trans_cat in enumerate(df_PROG_SPEC_CATES):
        category_credits_sum = {'學分': df_PROG_SPEC_CATES[idx]['學分'].sum(
        ), 'Required_CP': program_category[idx]['Required_CP']}
        df_PROG_SPEC_CATES[idx] = df_PROG_SPEC_CATES[idx].append(
            category_credits_sum, ignore_index=True)

    # Write to Excel
    start_row = 0
    for idx, sortedcourses in enumerate(df_PROG_SPEC_CATES):
        sortedcourses.to_excel(
            writer, sheet_name=program_name, startrow=start_row, header=True, index=False)
        df_PROG_SPEC_CATES_COURSES_SUGGESTION[idx].to_excel(
            writer, sheet_name=program_name, startrow=start_row, startcol=5, header=True, index=False)
        start_row += max(len(sortedcourses.index),
                         len(df_PROG_SPEC_CATES_COURSES_SUGGESTION[idx].index)) + 2

    # Formatting
    workbook = writer.book
    worksheet = writer.sheets[program_name]
    red_out_failed_subject(workbook, worksheet, 1, start_row)

    for df in df_PROG_SPEC_CATES:
        for i, col in enumerate(df.columns):
            # set the column length
            worksheet.set_column(i, i, column_len_array[i] * 2)
    print("Save to " + program_name)


def STUTTGART_EI(df_transcript_array, writer):
    print("Create Stuttgart_EI sheet")
    for idx, sortedcourses in enumerate(df_transcript_array):
        print(sortedcourses)
        print("")
    # TODO: implement the mapping from the existing courses to program's requirement
    start_row = 0
    for idx, sortedcourses in enumerate(df_transcript_array):
        sortedcourses.to_excel(
            writer, sheet_name='Uni_Stuttgart_EI', startrow=start_row, index=False)
        start_row += len(sortedcourses.index) + 2
    print("Save to Stuttgart_EI")


program_sort_function = [TUM_EI, RWTH_EI, STUTTGART_EI]


def EE_sorter(program_idx, file_path):

    Database_Path = os.getcwd() + '\\database\\'
    Output_Path = os.getcwd() + '\\output\\'

    Database_file_name = 'EE_Course_database.xlsx'
    input_file_name = os.path.split(file_path)
    input_file_name = input_file_name[1]
    print("input file name " + input_file_name)

    df_transcript = pd.read_excel(file_path,
                                  sheet_name='Transcript_Sorting')
    # Verify the format of transcript_course_list.xlsx
    if df_transcript.columns[0] != '所修科目' or df_transcript.columns[1] != '學分' or df_transcript.columns[2] != '成績':
        print("Error: Please check the student's transcript xlsx file.")
        print("Header: column_1 = 所修科目, column_2 = 學分, column_3 = 成績")
        sys.exit()

    df_database = pd.read_excel(Database_Path+Database_file_name,
                                sheet_name='All_EE_Courses')
    # Verify the format of EE_Course_database.xlsx
    if df_database.columns[0] != '所有科目':
        print("Error: Please check the EE database xlsx file.")
        sys.exit()
    df_database['所有科目'] = df_database['所有科目'].fillna('-')

    # unify course naming convention
    Naming_Convention(df_transcript)

    sorted_courses = []
    # EE
    transcript_sorted_group_map = {
        '微積分': [CALCULUS_KEY_WORDS, CALCULUS_ANTI_KEY_WORDS, ['一', '二']],
        '數學': [MATH_KEY_WORDS, MATH_ANTI_KEY_WORDS],
        '物理': [PHYSICS_KEY_WORDS, PHYSICS_ANTI_KEY_WORDS, ['一', '二']],
        '物理實驗': [PHYSICS_EXP_KEY_WORDS, PHYSICS_EXP_ANTI_KEY_WORDS, ['一', '二']],
        '資訊': [PROGRAMMING_KEY_WORDS, PROGRAMMING_ANTI_KEY_WORDS],
        '控制系統': [CONTROL_THEORY_KEY_WORDS, CONTROL_THEORY_ANTI_KEY_WORDS],
        '電子': [ELECTRONICS_KEY_WORDS, ELECTRONICS_ANTI_KEY_WORDS, ['一', '二']],
        '電路': [ELECTRO_CIRCUIT_KEY_WORDS, ELECTRO_CIRCUIT_ANTI_KEY_WORDS, ['一', '二']],
        '電磁': [ELECTRO_MAGNET_KEY_WORDS, ELECTRO_MAGNET_ANTI_KEY_WORDS, ['一', '二']],
        '半導體': [SEMICONDUCTOR_KEY_WORDS, SEMICONDUCTOR_ANTI_KEY_WORDS],
        '電機專業選修': [ADVANCED_ELECTRO_KEY_WORDS, ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '專業應用課程': [APPLICATION_ORIENTED_KEY_WORDS, APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    suggestion_courses_sorted_group_map = {
        '微積分': [[], CALCULUS_ANTI_KEY_WORDS],
        '數學': [[], MATH_ANTI_KEY_WORDS],
        '物理': [[], PHYSICS_ANTI_KEY_WORDS],
        '資訊': [[], PROGRAMMING_ANTI_KEY_WORDS],
        '控制系統': [[], CONTROL_THEORY_ANTI_KEY_WORDS],
        '電子': [[], ELECTRONICS_ANTI_KEY_WORDS],
        '電路': [[], ELECTRO_CIRCUIT_ANTI_KEY_WORDS],
        '電磁': [[], ELECTRO_MAGNET_ANTI_KEY_WORDS],
        '半導體': [[], SEMICONDUCTOR_ANTI_KEY_WORDS],
        '電機專業選修': [[], ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '專業應用課程': [[], APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '其他': [[], USELESS_COURSES_ANTI_KEY_WORDS], }

    category_data = []
    df_category_data = []
    category_courses_sugesstion_data = []
    df_category_courses_sugesstion_data = []
    for idx, cat in enumerate(transcript_sorted_group_map):
        category_data = {cat: [], '學分': [], '成績': []}
        df_category_data.append(pd.DataFrame(data=category_data))
        df_category_courses_sugesstion_data.append(
            pd.DataFrame(data=category_courses_sugesstion_data, columns=['建議修課']))

    # 基本分類課程 (與學程無關)
    df_category_data = CourseSorting(
        df_transcript, df_category_data, transcript_sorted_group_map)

    # 基本分類電機課程資料庫
    df_category_courses_sugesstion_data = DatabaseCourseSorting(
        df_database, df_category_courses_sugesstion_data, transcript_sorted_group_map)

    print(df_category_courses_sugesstion_data)
    # TODO: screening used matched keywords and keep not-yet matched keyword to screenning the suggestion courses
    # TODO: suggestion courses not work exactly
    # TODO: 樹狀篩選? 微積分:[一,二] 同時有含 微積分、一  的，就從recommendation拿掉

    for idx, cat in enumerate(df_category_data):
        df_category_courses_sugesstion_data[idx]['建議修課'] = df_category_courses_sugesstion_data[idx]['建議修課'].str.replace(
            '(', '', regex=False)
        df_category_courses_sugesstion_data[idx]['建議修課'] = df_category_courses_sugesstion_data[idx]['建議修課'].str.replace(
            ')', '', regex=False)

    # TODO: replace the following algorithm 1
    # for idx, cat in enumerate(df_category_data):
    #     temp_array = cat[cat.columns[0]].tolist()
    #     # print(temp_array)
    #     df_category_courses_sugesstion_data[idx] = df_category_courses_sugesstion_data[idx][
    #         ~df_category_courses_sugesstion_data[idx]['建議修課'].str.contains('|'.join(temp_array))]

    # algorithm 2 :
    df_category_courses_sugesstion_data = SuggestionCourseAlgorithm(
        df_category_data, transcript_sorted_group_map, df_category_courses_sugesstion_data)

    output_file_name = 'generated_' + input_file_name
    writer = pd.ExcelWriter(
        Output_Path+output_file_name, engine='xlsxwriter')

    # for each_cat in df_category_data:
    #     sorted_courses.append(each_cat)
    sorted_courses = df_category_data

    start_row = 0
    for idx, sortedcourses in enumerate(sorted_courses):
        sortedcourses.to_excel(
            writer, sheet_name='General', startrow=start_row, index=False)
        start_row += len(sortedcourses.index) + 2
    workbook = writer.book
    worksheet = writer.sheets['General']
    global column_len_array

    red_out_failed_subject(workbook, worksheet, 1, start_row)

    for i, col in enumerate(df_transcript.columns):
        # find length of column i
        column_len = df_transcript[col].astype(str).str.len().max()
        # Setting the length if the column header is larger
        # than the max column value length
        column_len_array.append(max(column_len, len(col)))
        # set the column length
        worksheet.set_column(i, i, column_len_array[i] * 2)

    # Modify to column width for "Required_CP"
    column_len_array.append(6)

    for idx in program_idx:
        program_sort_function[idx](
            transcript_sorted_group_map,
            sorted_courses,
            df_category_courses_sugesstion_data,
            writer)

    writer.save()

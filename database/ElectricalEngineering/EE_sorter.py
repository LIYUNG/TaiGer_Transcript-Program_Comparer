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
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # 電力電子
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 通訊
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 半導體
        PROG_SPEC_VERTIEFUNG_EI_PARAM,  # 電機專業選修
        PROG_SPEC_ANWENDUNG_MODULE_PARAM,  # 應用科技
        PROG_SPEC_OTHERS,  # 力學,機械
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
        'Program_Category': 'Höhere_Mathematik', 'Required_CP': 30}
    PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Elektrotechnik', 'Required_CP': 66}
    PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Kommunikationstechnik', 'Required_CP': 30}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 基礎電機電子
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # 基礎通訊
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 資訊
        PROG_SPEC_OTHERS,  # 控制系統
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電子
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電路
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電磁
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電力電子
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # 通訊
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 半導體
        PROG_SPEC_GRUNDLAGE_KOMMUNIKATIONSTECHNIK_PARAM,  # 電機專業選修
        PROG_SPEC_OTHERS,  # 應用科技
        PROG_SPEC_OTHERS,  # 力學,機械
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
        'Program_Category': 'Höhere_Mathematik', 'Required_CP': 30}
    # Grundlagen der Elektrotechnik, Vertiefung Energietechnik
    # Schaltungstechnik, Elektrische Felder und Wellen,Festkörperphysik und
    # Bauelemente, Hochspannungstechnik und Energie-übertragungstechnik,
    # elektrische Maschinen, etc.
    PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Grundlagen_Elektrotechnik', 'Required_CP': 45}
    # Grundlagen des Maschinenwesens, Vertiefung Energietechnik
    # (Technische Mechanik, Thermodynamik, Strömungsmechanik,
    # Wärme-und Stoffübertragung, Maschinendynamik, etc.)
    PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM = {
        'Program_Category': 'Grundlagen_Maschinenwesen', 'Required_CP': 45}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 基礎電機電子
        PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM,  # 基礎機械
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_OTHERS,  # 資訊
        PROG_SPEC_OTHERS,  # 控制系統
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電子
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電路
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電磁
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電力電子
        PROG_SPEC_OTHERS,  # 通訊
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 半導體
        PROG_SPEC_GRUNDLAGE_ELEKTROTECHNIK_PARAM,  # 電機專業選修
        PROG_SPEC_OTHERS,  # 應用科技
        PROG_SPEC_GRUNDLAGE_MASCHINEN_PARAM,  # 力學,機械相關
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


program_sort_function = [TUM_EI, RWTH_EI, STUTTGART_EI, TUM_MSCE, TUM_MSPE]


def EE_sorter(program_idx, file_path):

    Database_Path = env_file_path + '/'
    Output_Path = os.path.split(file_path)
    Output_Path = Output_Path[0]
    Output_Path = Output_Path + '/output/'
    print("output file path " + Output_Path)

    if not os.path.exists(Output_Path):
        print("create output folder")
        os.makedirs(Output_Path)

    Database_file_name = 'EE_Course_database.xlsx'
    input_file_name = os.path.split(file_path)
    input_file_name = input_file_name[1]
    print("input file name " + input_file_name)

    df_transcript = pd.read_excel(file_path,
                                  sheet_name='Transcript_Sorting')
    # Verify the format of transcript_course_list.xlsx
    if '所修科目' not in df_transcript.columns or '學分' not in df_transcript.columns or '成績' not in df_transcript.columns:
        print("Error: Please check the student's transcript xlsx file.")
        print(" There must be 所修科目, 學分 and 成績 in student's course excel file.")
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
        '微積分': [EE_CALCULUS_KEY_WORDS, EE_CALCULUS_ANTI_KEY_WORDS, ['一', '二']],
        '數學': [EE_MATH_KEY_WORDS, EE_MATH_ANTI_KEY_WORDS],
        '物理': [EE_PHYSICS_KEY_WORDS, EE_PHYSICS_ANTI_KEY_WORDS, ['一', '二']],
        '物理實驗': [EE_PHYSICS_EXP_KEY_WORDS, EE_PHYSICS_EXP_ANTI_KEY_WORDS, ['一', '二']],
        '資訊': [EE_PROGRAMMING_KEY_WORDS, EE_PROGRAMMING_ANTI_KEY_WORDS],
        '控制系統': [EE_CONTROL_THEORY_KEY_WORDS, EE_CONTROL_THEORY_ANTI_KEY_WORDS],
        '電子': [EE_ELECTRONICS_KEY_WORDS, EE_ELECTRONICS_ANTI_KEY_WORDS, ['一', '二']],
        '電路': [EE_ELECTRO_CIRCUIT_KEY_WORDS, EE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS, ['一', '二']],
        '電磁': [EE_ELECTRO_MAGNET_KEY_WORDS, EE_ELECTRO_MAGNET_ANTI_KEY_WORDS, ['一', '二']],
        '電力電子': [EE_POWER_ELECTRO_KEY_WORDS, EE_POWER_ELECTRO_ANTI_KEY_WORDS, ['一', '二']],
        '通訊': [EE_COMMUNICATION_KEY_WORDS, EE_COMMUNICATION_ANTI_KEY_WORDS, ['一', '二']],
        '半導體': [EE_SEMICONDUCTOR_KEY_WORDS, EE_SEMICONDUCTOR_ANTI_KEY_WORDS],
        '電機專業選修': [EE_ADVANCED_ELECTRO_KEY_WORDS, EE_ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '專業應用課程': [EE_APPLICATION_ORIENTED_KEY_WORDS, EE_APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '力學': [EE_MACHINE_RELATED_KEY_WORDS, EE_MACHINE_RELATED_ANTI_KEY_WORDS],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    suggestion_courses_sorted_group_map = {
        '微積分': [[], EE_CALCULUS_ANTI_KEY_WORDS],
        '數學': [[], EE_MATH_ANTI_KEY_WORDS],
        '物理': [[], EE_PHYSICS_ANTI_KEY_WORDS],
        '資訊': [[], EE_PROGRAMMING_ANTI_KEY_WORDS],
        '控制系統': [[], EE_CONTROL_THEORY_ANTI_KEY_WORDS],
        '電子': [[], EE_ELECTRONICS_ANTI_KEY_WORDS],
        '電路': [[], EE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS],
        '電磁': [[], EE_ELECTRO_MAGNET_ANTI_KEY_WORDS],
        '電力電子': [[], EE_POWER_ELECTRO_ANTI_KEY_WORDS],
        '通訊': [[], EE_COMMUNICATION_ANTI_KEY_WORDS],
        '半導體': [[], EE_SEMICONDUCTOR_ANTI_KEY_WORDS],
        '電機專業選修': [[], EE_ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '專業應用課程': [[], EE_APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '力學': [[], EE_MACHINE_RELATED_ANTI_KEY_WORDS],
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

    for idx, cat in enumerate(df_category_data):
        df_category_courses_sugesstion_data[idx]['建議修課'] = df_category_courses_sugesstion_data[idx]['建議修課'].str.replace(
            '(', '', regex=False)
        df_category_courses_sugesstion_data[idx]['建議修課'] = df_category_courses_sugesstion_data[idx]['建議修課'].str.replace(
            ')', '', regex=False)

    # 樹狀篩選 微積分:[一,二] 同時有含 微積分、一  的，就從recommendation拿掉
    # algorithm :
    df_category_courses_sugesstion_data = SuggestionCourseAlgorithm(
        df_category_data, transcript_sorted_group_map, df_category_courses_sugesstion_data)

    output_file_name = 'analyzed_' + input_file_name
    writer = pd.ExcelWriter(
        Output_Path+output_file_name, engine='xlsxwriter')

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
    print("output data at: " + Output_Path + output_file_name)
    print("Students' courses analysis and courses suggestion in EE area finished! ")
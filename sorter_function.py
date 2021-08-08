from numpy import nan
import sys
import os
import pandas as pd
from cell_formatter import red_out_failed_subject

import xlsxwriter
KEY_WORDS = 0
ANTI_KEY_WORDS = 1
CALCULUS_KEY_WORDS = ['微積分']
CALCULUS_ANTI_KEY_WORDS = ['asdgladfj;l']
MATH_KEY_WORDS = ['數學', '代數', '微分', '函數', '機率', '離散', '複變', '數值', '向量']
MATH_ANTI_KEY_WORDS = ['asdgladfj;l']
PROGRAMMING_KEY_WORDS = ['計算機', '演算', '資料', '物件', '運算',
                         '資電', '作業系統', '資料結構', '軟體', '編譯器', '程式設計', '程式語言', 'Python', 'C++', 'C語言']
PROGRAMMING_ANTI_KEY_WORDS = ['asdgladfj;l']
PHYSICS_KEY_WORDS = ['物理']
PHYSICS_ANTI_KEY_WORDS = ['半導體', '元件']
CHEMISTRY_KEY_WORDS = ['化學']
CHEMISTRY_ANTI_KEY_WORDS = ['asdgladfj;l']
ELECTRONICS_KEY_WORDS = ['電子']
ELECTRONICS_ANTI_KEY_WORDS = ['專題', '電力', '固態', '自動化']
ELECTRO_CIRCUIT_KEY_WORDS = ['電路', '訊號', '數位邏輯', '邏輯設計', '信號與系統']
ELECTRO_CIRCUIT_ANTI_KEY_WORDS = ['超大型', '專題']
ELECTRO_MAGNET_KEY_WORDS = ['電磁']
ELECTRO_MAGNET_ANTI_KEY_WORDS = ['asdgladfj;l', '專題']
ADVANCED_ELECTRO_KEY_WORDS = ['微波', '積體電路', '自動化', '天線', '網路', '高頻', '無線', '藍芽', '晶片',
                              '類比', '數位訊號', '通信',  '通訊', '微算機', '微處理', '電波', 'VLSI', '固態', '嵌入式', '人工智慧', '無線網路', '機器學習', '消息']
ADVANCED_ELECTRO_ANTI_KEY_WORDS = ['asdgladfj;l']
SEMICONDUCTOR_KEY_WORDS = ['半導體', '元件']
SEMICONDUCTOR_ANTI_KEY_WORDS = ['專題']
APPLICATION_ORIENTED_KEY_WORDS = ['電力', '生醫', '能源', '光機電', '電動機',
                                  '電機', '影像', '深度學習', '光電', '應用', '綠能', '雲端運算', '醫學工程', '再生能源']
APPLICATION_ORIENTED_ANTI_KEY_WORDS = ['asdgladfj;l']

MECHANICS_KEY_WORDS = ['熱力學', '動力', '靜力', '材料力', '摩擦', '流體']
MECHANICS_ANTI_KEY_WORDS = ['asdgladfj;l']
CONTROL_THEORY_KEY_WORDS = ['控制']
CONTROL_THEORY_ANTI_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_ANTI_KEY_WORDS = ['asdgladfj;l']

# Global variable:
column_len_array = []


def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def TUM_EI(transcript_sorted_group_map, df_transcript_array, writer):
    print("Create TUM EI sheet")
    # TODO: implement the mapping from the existing courses to program's requirement
    start_row = 0
    for idx, sortedcourses in enumerate(df_transcript_array):
        sortedcourses.to_excel(
            writer, sheet_name='TUM_EI', startrow=start_row, index=False)
        start_row += len(sortedcourses.index) + 2
    print("Save to TUM_EI")


def RWTH_EI(transcript_sorted_group_map, df_transcript_array, writer):
    print("Create RWTH_Aachen_EI sheet")
    df_transcript_array_temp = df_transcript_array

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

    df_PROG_SPEC_CATES = []
    for idx, cat in enumerate(program_category):
        PROG_SPEC_CAT = {cat['Program_Category']: [],
                         '學分': [], '成績': [], 'Required_CP': cat['Required_CP']}
        df_PROG_SPEC_CATES.append(pd.DataFrame(data=PROG_SPEC_CAT))

    transcript_sorted_group_list = list(transcript_sorted_group_map)

    # N to 1 mapping
    for idx, trans_cat in enumerate(df_transcript_array_temp):
        # append sorted courses to program's category
        categ = program_category_map[idx]['Program_Category']
        trans_cat.rename(
            columns={transcript_sorted_group_list[idx]: categ}, inplace=True)

        # TODO find the idx corresponding to program's category
        idx_temp = -1
        for idx2, cat in enumerate(df_PROG_SPEC_CATES):
            if categ == cat.columns[0]:
                print(cat.columns[0])
                idx_temp = idx2
                break
        df_PROG_SPEC_CATES[idx_temp] = df_PROG_SPEC_CATES[idx_temp].append(
            trans_cat, ignore_index=True)

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
            writer, sheet_name='RWTH_Aachen_EI', startrow=start_row, header=True, index=False)
        start_row += len(sortedcourses.index) + 2

    # Formatting
    workbook = writer.book
    worksheet = writer.sheets['RWTH_Aachen_EI']
    red_out_failed_subject(workbook, worksheet, 1, start_row)
    for df in df_PROG_SPEC_CATES:
        for i, col in enumerate(df.columns):
            # set the column length
            worksheet.set_column(i, i, column_len_array[i] * 2)
    print("Save to RWTH_Aachen_EI")


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


def func(program_idx):

    # Input_Path = os.getcwd() + '\\train_data\\'
    Input_Path = os.getcwd() + '\\database\\'
    Output_Path = os.getcwd() + '\\output\\'

    input_file_name = 'EE_Course_database.xlsx'
    # input_file_name = 'template.xlsx'
    # input_file_name = 'testdata1.xlsx'
    # input_file_name = 'testdata2.xlsx'
    # input_file_name = 'testdata3.xlsx'

    df_transcript = pd.read_excel(Input_Path+input_file_name,
                                  sheet_name='Transcript_Sorting')
    # Verify the format of transcript_course_list.xlsx
    if df_transcript.columns[0] != '所修科目' or df_transcript.columns[1] != '學分' or df_transcript.columns[2] != '成績':
        print("Error: Please check the student's transcript xlsx file.")
        sys.exit()

    df_transcript['所修科目'] = df_transcript['所修科目'].fillna('-')
    sorted_courses = []

    transcript_sorted_group_map = {
        '微積分': [CALCULUS_KEY_WORDS, CALCULUS_ANTI_KEY_WORDS],
        '數學': [MATH_KEY_WORDS, MATH_ANTI_KEY_WORDS],
        '物理': [PHYSICS_KEY_WORDS, PHYSICS_ANTI_KEY_WORDS],
        '資訊': [PROGRAMMING_KEY_WORDS, PROGRAMMING_ANTI_KEY_WORDS],
        '控制系統': [CONTROL_THEORY_KEY_WORDS, CONTROL_THEORY_ANTI_KEY_WORDS],
        '電子': [ELECTRONICS_KEY_WORDS, ELECTRONICS_ANTI_KEY_WORDS],
        '電路': [ELECTRO_CIRCUIT_KEY_WORDS, ELECTRO_CIRCUIT_ANTI_KEY_WORDS],
        '電磁': [ELECTRO_MAGNET_KEY_WORDS, ELECTRO_MAGNET_ANTI_KEY_WORDS],
        '半導體': [SEMICONDUCTOR_KEY_WORDS, SEMICONDUCTOR_ANTI_KEY_WORDS],
        '電機專業選修': [ADVANCED_ELECTRO_KEY_WORDS, ADVANCED_ELECTRO_ANTI_KEY_WORDS],
        '專業應用課程': [APPLICATION_ORIENTED_KEY_WORDS, APPLICATION_ORIENTED_ANTI_KEY_WORDS],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    category_data = []
    df_category_data = []
    for idx, cat in enumerate(transcript_sorted_group_map):
        category_data = {cat: [], '學分': [], '成績': []}
        df_category_data.append(pd.DataFrame(data=category_data))

    for idx, subj in enumerate(df_transcript['所修科目']):
        if subj == '-':
            continue
        for idx2, cat in enumerate(transcript_sorted_group_map):
            if(idx2 == len(transcript_sorted_group_map) - 1):
                temp = {cat: subj, '學分': df_transcript['學分'][idx],
                        '成績': df_transcript['成績'][idx]}
                df_category_data[idx2] = df_category_data[idx2].append(
                    temp, ignore_index=True)
                continue
            # filter subject by keywords. and exclude subject by anti_keywords
            if any(keywords in subj for keywords in transcript_sorted_group_map[cat][KEY_WORDS] if not any(anti_keywords in subj for anti_keywords in transcript_sorted_group_map[cat][ANTI_KEY_WORDS])):
                temp_string = str(df_transcript['成績'][idx])
                if((isfloat(temp_string) and float(temp_string) < 60)):  # failed subject not count
                    continue
                temp = {cat: subj, '學分': df_transcript['學分'][idx],
                        '成績': df_transcript['成績'][idx]}
                df_category_data[idx2] = df_category_data[idx2].append(
                    temp, ignore_index=True)
                break

    output_file_name = 'generated_' + input_file_name
    writer = pd.ExcelWriter(
        Output_Path+output_file_name, engine='xlsxwriter')

    for each_cat in df_category_data:
        sorted_courses.append(each_cat)

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

    print(column_len_array)
    # Modify to column width for "Required_CP"
    column_len_array.append(6)

    for idx in program_idx:
        program_sort_function[idx](
            transcript_sorted_group_map,
            sorted_courses,
            writer)

    writer.save()

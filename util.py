import pandas as pd
from CourseSuggestionAlgorithms import *
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import gc
import sys, os

KEY_WORDS = 0
ANTI_KEY_WORDS = 1
DIFFERENTIATE_KEY_WORDS = 2

# naming convention


def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def ProgramCategoryInit(program_category):
    df_PROG_SPEC_CATES = []
    df_PROG_SPEC_CATES_COURSES_SUGGESTION = []
    for idx, cat in enumerate(program_category):
        PROG_SPEC_CAT = {cat['Program_Category']: [],
                         '學分': [], '成績': [], 'Required_CP': cat['Required_CP']}
        PROG_SPEC_CATES_COURSES_SUGGESTION = {cat['Program_Category']: [],
                                              }
        df_PROG_SPEC_CATES.append(pd.DataFrame(data=PROG_SPEC_CAT))
        df_PROG_SPEC_CATES_COURSES_SUGGESTION.append(
            pd.DataFrame(data=PROG_SPEC_CATES_COURSES_SUGGESTION))
    return df_PROG_SPEC_CATES, df_PROG_SPEC_CATES_COURSES_SUGGESTION


def CheckTemplateFormat(df_transcript):
    if '所修科目_中文' not in df_transcript.columns or '所修科目_英語' not in df_transcript.columns or '學分' not in df_transcript.columns or '成績' not in df_transcript.columns:
        print("Error: Please check the student's transcript xlsx file.")
        print(" There must be 所修科目_中文, 所修科目_英文, 學分 and 成績 in student's course excel file.")
        sys.exit()


def CheckDBFormat(df_database):
    if '所有科目' not in df_database.columns:
        print("Error: Please check the database xlsx file.")
        sys.exit()


def isOutputEnglish(df_transcript):
    
    if ~df_transcript['所修科目_英語'].isnull().any():
        # output English version
        print("Output English Version")
        return True

    # print(df_transcript['所修科目_中文'].isnull().any())
    if ~df_transcript['所修科目_中文'].isnull().any():
        print("Output Chinese Version")
        return False  # output CHinese version

    print("所修科目_英語 所修科目_中文 學分 和成績 noat make. Please make sure you fill the template correctly")
    sys.exit()

def DataPreparation(df_database, df_transcript):
    df_database['所有科目'] = df_database['所有科目'].fillna('-')
    if '所有科目_英語' in df_database.columns:
        df_database['所有科目_英語'] = df_database['所有科目_英語'].str.lower()
    # unify course naming convention
    df_transcript = Naming_Convention_ZH(df_transcript)
    df_transcript = Naming_Convention_EN(df_transcript)
    print("Prepared data successfully.")
    return df_database, df_transcript

def Naming_Convention_ZH(df_course):
    # modify data in the same
    df_course['所修科目_中文'] = df_course['所修科目_中文'].fillna('-')

    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '1', '一', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '2', '二', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '3', '三', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '(', '', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '（', '', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        ')', '', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        '）', '', regex=False)
    df_course['所修科目_中文'] = df_course['所修科目_中文'].str.replace(
        ' ', '', regex=False)    
    return df_course

def Naming_Convention_EN(df_course):
    ## English Version only:
    df_course['所修科目_英語'] = df_course['所修科目_英語'].fillna('-')

    df_course['所修科目_英語'] = df_course['所修科目_英語'].str.lower()

    df_course['所修科目_英語'] = df_course['所修科目_英語'].str.replace(
        '(', '', regex=False)
    df_course['所修科目_英語'] = df_course['所修科目_英語'].str.replace(
        '（', '', regex=False)
    df_course['所修科目_英語'] = df_course['所修科目_英語'].str.replace(
        ')', '', regex=False)
    df_course['所修科目_英語'] = df_course['所修科目_英語'].str.replace(
        '）', '', regex=False)
    return df_course


# mapping courses to target programs category
def CoursesToProgramCategoryMapping(df_PROG_SPEC_CATES, program_category_map, transcript_sorted_group_list, df_transcript_array_temp, isSuggestionCourse):
    for idx, trans_cat in enumerate(df_transcript_array_temp):
        # append sorted courses to program's category
        categ = program_category_map[idx]['Program_Category']
        trans_cat.rename(
            columns={transcript_sorted_group_list[idx]: categ}, inplace=True)
        # find the idx corresponding to program's category
        idx_temp = -1
        for idx2, cat in enumerate(df_PROG_SPEC_CATES):
            if categ == cat.columns[0]:
                # print(cat.columns[0])
                idx_temp = idx2
                break
        # remove the redundant suggestion courses mapping to "Others" because those categories in Others are not advanced courses.
        if isSuggestionCourse:
            if idx != len(df_transcript_array_temp) - 1 and idx_temp == len(df_PROG_SPEC_CATES) - 1:
                continue

        df_PROG_SPEC_CATES[idx_temp] = df_PROG_SPEC_CATES[idx_temp].append(
            trans_cat, ignore_index=True)
    return df_PROG_SPEC_CATES


# course sorting
def CourseSorting(df_transcript, df_category_data, transcript_sorted_group_map, column_name_en_zh):
    # print(df_transcript[column_name_en_zh])
    # df_transcript = df_transcript.dropna()
    for idx, subj in enumerate(df_transcript[column_name_en_zh]):
        if subj == '-':
            continue
        for idx2, cat in enumerate(transcript_sorted_group_map):
            # Put the rest of courses to Others
            if(idx2 == len(transcript_sorted_group_map) - 1):
                temp = {cat: subj, '學分': df_transcript['學分'][idx],
                        '成績': df_transcript['成績'][idx]}
                df_category_data[idx2] = df_category_data[idx2].append(
                    temp, ignore_index=True)
                continue
            
            # filter subject by keywords. and exclude subject by anti_keywords
            if any(keywords in subj for keywords in transcript_sorted_group_map[cat][KEY_WORDS] if not any(anti_keywords in subj for anti_keywords in transcript_sorted_group_map[cat][ANTI_KEY_WORDS])):
                temp_string = str(df_transcript['成績'][idx])
                # failed subject not count
                if((isfloat(temp_string) and float(temp_string) < 60 and float(temp_string) and float(temp_string)> 4.5)):
                    continue
                temp = {cat: subj, '學分': df_transcript['學分'][idx],
                        '成績': df_transcript['成績'][idx]}
                df_category_data[idx2] = df_category_data[idx2].append(
                    temp, ignore_index=True)
                break
    return df_category_data


def DatabaseCourseSorting(df_database, df_category_courses_sugesstion_data, transcript_sorted_group_map, column_name_en_zh):
    for idx, subj in enumerate(df_database[column_name_en_zh]):
        if subj == '-':
            continue
        for idx2, cat in enumerate(transcript_sorted_group_map):
            # Put the rest of courses to Others
            if(idx2 == len(transcript_sorted_group_map) - 1):
                temp = {'建議修課': subj}
                df_category_courses_sugesstion_data[idx2] = df_category_courses_sugesstion_data[idx2].append(
                    temp, ignore_index=True)
                continue

            # filter database by keywords. and exclude subject by anti_keywords
            if any(keywords in subj for keywords in transcript_sorted_group_map[cat][KEY_WORDS] if not any(anti_keywords in subj for anti_keywords in transcript_sorted_group_map[cat][ANTI_KEY_WORDS])):
                temp = {'建議修課': subj}
                df_category_courses_sugesstion_data[idx2] = df_category_courses_sugesstion_data[idx2].append(
                    temp, ignore_index=True)
                break
    return df_category_courses_sugesstion_data


def AppendCreditsCount(df_PROG_SPEC_CATES, program_category):
    for idx, trans_cat in enumerate(df_PROG_SPEC_CATES):
        category_credits_sum = {'學分': df_PROG_SPEC_CATES[idx]['學分'].sum(
        ), 'Required_CP': program_category[idx]['Required_CP']}
        df_PROG_SPEC_CATES[idx] = df_PROG_SPEC_CATES[idx].append(
            category_credits_sum, ignore_index=True)
    return df_PROG_SPEC_CATES


def WriteToExcel(writer, program_name, program_category, program_category_map, transcript_sorted_group_map, df_transcript_array_temp, df_category_courses_sugesstion_data_temp, column_len_array):
    df_PROG_SPEC_CATES, df_PROG_SPEC_CATES_COURSES_SUGGESTION = ProgramCategoryInit(
        program_category)

    transcript_sorted_group_list = list(transcript_sorted_group_map)

    # Courses: mapping the students' courses to program-specific category
    df_PROG_SPEC_CATES = CoursesToProgramCategoryMapping(
        df_PROG_SPEC_CATES, program_category_map, transcript_sorted_group_list, df_transcript_array_temp, False)

    # Suggestion courses: mapping the sugesstion courses to program-specific category
    df_PROG_SPEC_CATES_COURSES_SUGGESTION = CoursesToProgramCategoryMapping(
        df_PROG_SPEC_CATES_COURSES_SUGGESTION, program_category_map, transcript_sorted_group_list, df_category_courses_sugesstion_data_temp, True)

    # append 總學分 for each program category
    df_PROG_SPEC_CATES = AppendCreditsCount(
        df_PROG_SPEC_CATES, program_category)

    # drop the Others, 建議修課
    # for idx, trans_cat in enumerate(df_PROG_SPEC_CATES_COURSES_SUGGESTION):
    #     if(idx == len(df_PROG_SPEC_CATES_COURSES_SUGGESTION) - 1):
    #         df_PROG_SPEC_CATES_COURSES_SUGGESTION[idx].drop(
    #             columns=['Others', '建議修課'], inplace=True)

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
    # red_out_insufficient_credit(workbook, worksheet)

    # print("writer")
    # print(writer['A1'])
    # print("worksheet")
    # print(worksheet)
    for df in df_PROG_SPEC_CATES:
        # print(df)
        for i, col in enumerate(df.columns):
            # print(i)
            # set the column length
            worksheet.set_column(i, i, column_len_array[i] * 2)
    gc.collect()  # Forced GC
    print("Save to " + program_name)


def Classifier(program_idx, file_path, abbrev, env_file_path, basic_classification_en, basic_classification_zh, column_len_array, program_sort_function):

    program_idx, file_path, abbrev
    Database_Path = env_file_path + '/'
    Output_Path = os.path.split(file_path)
    Output_Path = Output_Path[0]
    Output_Path = Output_Path + '/output/'
    print("output file path " + Output_Path)

    if not os.path.exists(Output_Path):
        print("create output folder")
        os.makedirs(Output_Path)

    Database_file_name = abbrev + '_Course_database.xlsx'
    input_file_name = os.path.split(file_path)
    input_file_name = input_file_name[1]
    print("input file name " + input_file_name)

    df_transcript = pd.read_excel(file_path,
                                  sheet_name='Transcript_Sorting')
    # TODO: move the checking mechanism to util.py!
    # Verify the format of transcript_course_list.xlsx
    CheckTemplateFormat(df_transcript)
    print("Checked input template successfully.")

    sheetname = 'All_' + abbrev + '_Courses'
    df_database = pd.read_excel(Database_Path+Database_file_name,
                                sheet_name=sheetname)
    # Verify the format of Course_database.xlsx
    CheckDBFormat(df_database)
    print("Checked database successfully.")

    Englist_Version = isOutputEnglish(df_transcript)
    # Data preparation
    df_database, df_transcript = DataPreparation(df_database, df_transcript)

    sorted_courses = []
    transcript_sorted_group_map = {}

    if Englist_Version:
        transcript_sorted_group_map = basic_classification_en
    else:
        transcript_sorted_group_map = basic_classification_zh

    category_data = []
    df_category_data = []
    category_courses_sugesstion_data = []
    df_category_courses_sugesstion_data = []
    for idx, cat in enumerate(transcript_sorted_group_map):
        category_data = {cat: [], '學分': [], '成績': []}
        df_category_data.append(pd.DataFrame(data=category_data))
        df_category_courses_sugesstion_data.append(
            pd.DataFrame(data=category_courses_sugesstion_data, columns=['建議修課']))

    if Englist_Version:
        # 基本分類課程 (與申請學程無關)
        df_category_data = CourseSorting(
            df_transcript, df_category_data, transcript_sorted_group_map, "所修科目_英語")
        # 基本分類電機課程資料庫
        df_category_courses_sugesstion_data = DatabaseCourseSorting(
            df_database, df_category_courses_sugesstion_data, transcript_sorted_group_map, "所有科目_英語")
    else:
        # 基本分類課程 (與申請學程無關)
        df_category_data = CourseSorting(
            df_transcript, df_category_data, transcript_sorted_group_map, "所修科目_中文")
        # 基本分類電機課程資料庫
        df_category_courses_sugesstion_data = DatabaseCourseSorting(
            df_database, df_category_courses_sugesstion_data, transcript_sorted_group_map, "所有科目")

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

    red_out_failed_subject(workbook, worksheet, 1, start_row)

    for i, col in enumerate(df_transcript.columns):
        # find length of column i

        column_len = df_transcript[col].astype(str).str.len().max()
        # Setting the length if the column header is larger
        # than the max column value length
        if i == 1:
            column_len_array.append(len(col))
        else:
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
    print("Students' courses analysis and courses suggestion in " +
          abbrev + " area finished! ")

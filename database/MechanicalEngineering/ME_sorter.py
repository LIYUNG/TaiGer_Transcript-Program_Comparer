import xlsxwriter
from CourseSuggestionAlgorithms import *
from util import *
from database.MechanicalEngineering.ME_KEYWORDS import *
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import pandas as pd
import sys
import os
env_file_path = os.path.realpath(__file__)
env_file_path = os.path.dirname(env_file_path)

# Global variable:
column_len_array = []


def RWTH_AUTO(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'RWTH_Aachen_AUTO'
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

    PROG_SPEC_MECHANIK_PARAM = {
        'Program_Category': 'Mechanik', 'Required_CP': 18}
    PROG_SPEC_MASCHINENGESTALTUNG_PARAM = {
        'Program_Category': 'Maschinengestaltung', 'Required_CP': 13}
    PROG_SPEC_THERMODYNAMIK_PARAM = {
        'Program_Category': 'Thermodynamik', 'Required_CP': 7}
    PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM = {
        'Program_Category': 'Wärm_und_Stoffübertragung', 'Required_CP': 6}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde', 'Required_CP': 8}
    PROG_SPEC_CONTROL_TECHNIQUE_PARAM = {
        'Program_Category': 'Regelungstechnik', 'Required_CP': 6}
    PROG_SPEC_STROEMUNGSMECHANIK_PARAM = {
        'Program_Category': 'Strömungsmechanik I', 'Required_CP': 6}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Höhere Mathematik', 'Required_CP': 17}
    PROG_SPEC_FAHRZEUGTECHNIK_PARAM = {
        'Program_Category': 'Fahrzeugtechnik', 'Required_CP': 22}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MECHANIK_PARAM,  # 力學
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機械繪圖
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_FAHRZEUGTECHNIK_PARAM,  # 汽車
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機械設計
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_MECHANIK_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_FAHRZEUGTECHNIK_PARAM,  # 車輛
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

# FPSO: https://www.tum.de/fileadmin/w00bfo/www/Studium/Studienangebot/Lesbare_Fassung/Master/Maschinenwesen_MA_LB_3._AS_28052021.pdf


def TUM_MW(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_MW'
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
        'Program_Category': 'Höhere Mathematik 1,2,3', 'Required_CP': 19}
    PROG_SPEC_MECHANIK_PARAM = {
        'Program_Category': 'Technische Mechanik 1,2,3', 'Required_CP': 19}
    PROG_SPEC_MASCHINENELEMENTE_PARAM = {
        'Program_Category': 'Maschinenelemente 1,2', 'Required_CP': 15}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde 1,2', 'Required_CP': 10}
    PROG_SPEC_THERMODYNAMIK_PARAM = {
        'Program_Category': 'Thermodynamik', 'Required_CP': 6}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_MECHANIK_PARAM,  # 力學
        PROG_SPEC_MASCHINENELEMENTE_PARAM,  # 機械
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MASCHINENELEMENTE_PARAM,  # 機械設計
        PROG_SPEC_MASCHINENELEMENTE_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_OTHERS,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_MECHANIK_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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

# TODO


def TUHH_MECHATRONICS(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUHH_MECHATRONICS'
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
    PROG_SPEC_CALCULUS_PARAM = {
        'Program_Category': 'Calculus', 'Required_CP': 12}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_CP': 18}
    PROG_SPEC_MECHANICS_PARAM = {
        'Program_Category': 'Mechanics', 'Required_CP': 21}
    PROG_SPEC_CONTROL_THEORY_PARAM = {
        'Program_Category': 'Control Theory', 'Required_CP': 6}
    PROG_SPEC_MECHATRONICS_PARAM = {
        'Program_Category': 'Mechatronics', 'Required_CP': 6}
    PROG_SPEC_MATERIAL_PROPERTIES_TESTING_PARAM = {
        'Program_Category': 'Materials Science', 'Required_CP': 6}
    PROG_SPEC_MANUFACTURING_PARAM = {
        'Program_Category': 'Manufacturing Porcesses', 'Required_CP': 6}
    PROG_SPEC_MEASUREMENT_PARAM = {
        'Program_Category': 'Measurement Technology', 'Required_CP': 6}
    PROG_SPEC_THERMODYNAMICS_PARAM = {
        'Program_Category': 'Thermodynamics', 'Required_CP': 6}
    PROG_SPEC_COMPUTER_SCIENCE_PARAM = {
        'Program_Category': 'Computer Science', 'Required_CP': 6}
    PROG_SPEC_MECHANICAL_DESIGN_PARAM = {
        'Program_Category': 'Mechanical Engineering Design', 'Required_CP': 12}
    PROG_SPEC_ELECTRICAL_ENG_PARAM = {
        'Program_Category': 'Electrical Engineering', 'Required_CP': 12}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_CALCULUS_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_MECHANICS_PARAM,  # 力學
        PROG_SPEC_CONTROL_THEORY_PARAM,  # 控制理論
        PROG_SPEC_MECHATRONICS_PARAM,  # 機電
        PROG_SPEC_MATERIAL_PROPERTIES_TESTING_PARAM,  # 材料
        PROG_SPEC_MANUFACTURING_PARAM,  # 製造
        PROG_SPEC_MEASUREMENT_PARAM,  # 測量
        PROG_SPEC_THERMODYNAMICS_PARAM,  # 熱力
        PROG_SPEC_COMPUTER_SCIENCE_PARAM,  # 計算機概論
        PROG_SPEC_MECHANICAL_DESIGN_PARAM,  # 機械設計
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機電子
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_CALCULUS_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MECHANICAL_DESIGN_PARAM,  # 機械設計
        PROG_SPEC_MECHANICAL_DESIGN_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMICS_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_MATERIAL_PROPERTIES_TESTING_PARAM,  # 材料
        PROG_SPEC_CONTROL_THEORY_PARAM,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_MECHANICS_PARAM,  # 力學,機械
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機電子
        PROG_SPEC_MANUFACTURING_PARAM,  # 製造
        PROG_SPEC_COMPUTER_SCIENCE_PARAM,  # 計算機概論
        PROG_SPEC_MECHATRONICS_PARAM,  # 機電
        PROG_SPEC_MEASUREMENT_PARAM,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


def UNI_HANNOVER_INTER_MECHATRONICS(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'UNI_HANNOVER_INTER_MECHATRONICS'
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
        'Program_Category': 'Mathematics', 'Required_CP': 15}
    PROG_SPEC_ELECTRICAL_ENG_PARAM = {
        'Program_Category': 'Electricl Engineering', 'Required_CP': 20}
    PROG_SPEC_MECHANICS_PARAM = {
        'Program_Category': 'Technical Mechanics', 'Required_CP': 15}
    PROG_SPEC_CONTROL_MEASUREMENT_PARAM = {
        'Program_Category': 'Control Engineeringv/ Measuring Technology', 'Required_CP': 10}

    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機
        PROG_SPEC_MECHANICS_PARAM,  # 力學
        PROG_SPEC_CONTROL_MEASUREMENT_PARAM,  # 控制理論
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MECHANICS_PARAM,  # 機械設計
        PROG_SPEC_MECHANICS_PARAM,  # 機構
        PROG_SPEC_MECHANICS_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_OTHERS,  # 材料
        PROG_SPEC_CONTROL_MEASUREMENT_PARAM,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_MECHANICS_PARAM,  # 力學,機械
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


def TU_DORTMUND_MANUFAC_TECH(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TU_DORTMUND_MANUFAC_TECH'
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

    PROG_SPEC_ELECTRICAL_ENG_PARAM = {
        'Program_Category': 'Electricl Engineering', 'Required_CP': 12}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_CP': 18}
    PROG_SPEC_MANUFACTURE_PARAM = {
        'Program_Category': 'Manufacturing Subjects', 'Required_CP': 20}  # materials engineering, production engineering, theory of de­sign, and/or metallurgy and feedback control
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機
        PROG_SPEC_MANUFACTURE_PARAM,  # 製造工程
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MANUFACTURE_PARAM,  # 機械設計
        PROG_SPEC_MANUFACTURE_PARAM,  # 機構
        PROG_SPEC_OTHERS,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_MANUFACTURE_PARAM,  # 材料
        PROG_SPEC_MANUFACTURE_PARAM,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_OTHERS,  # 力學,機械
        PROG_SPEC_ELECTRICAL_ENG_PARAM,  # 基礎電機電子
        PROG_SPEC_MANUFACTURE_PARAM,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


def TU_CHEMNITZ_AD_MANUFAC(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TU_CHEMNITZ_AD_MANUFAC'
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
    # 1. special mathematical methods of engineering, totalling at least 18 ECTS and including the topics of Fourier transforms, regression calculation, probability and mathematical statistics,
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_CP': 18}
    # 2. scientific and engineering data processing, totalling at least 12 ECTS and including the topics CAD, CAS, numerical simulation and data acquisition as well as multiphysics simulation and practical experience,
    PROG_SPEC_CAD_PARAM = {
        'Program_Category': 'CAD, CAS, numerical simulation', 'Required_CP': 12}
    # #3. metrology and control engineering, totalling at least 8 ECTS and including the topics of sensors, actuators and digital methods of manufacturing,
    PROG_SPEC_METROL_CONTROL_PARAM = {
        'Program_Category': 'Metrology and Control Engineering', 'Required_CP': 8}
    # 4. new materials for engineering, totalling at least 8 LP and including the topics of polymers, metals, composites, matrix systems and functional properties,
    PROG_SPEC_MATERIALS_PARAM = {
        'Program_Category': 'Materials for Engineering', 'Required_CP': 8}
    # #5. in-depth theoretical basics of engineering, in a total of at least 12 ECTS and including the subjects of engineering mechanics, design, manufacturing and fluid dynamics,
    PROG_SPEC_MECH_MANU_PARAM = {
        'Program_Category': 'Mechanics, Design, Manufacturing and Fluid dynamics', 'Required_CP': 12}
    # #6. resource-efficient manufacturing concepts, totalling at least 8 ECTS and including the topics of technical and natural cycles and networks, system optimization and energy concepts,
    PROG_SPEC_MANU_CONCE_PARAM = {
        'Program_Category': 'Manufacturing Concepts', 'Required_CP': 8}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}
    # TODO:
    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_CAD_PARAM,  # 設計
        PROG_SPEC_METROL_CONTROL_PARAM,  # 控制 測量
        PROG_SPEC_MATERIALS_PARAM,  # 材料
        PROG_SPEC_MECH_MANU_PARAM,  # 力學
        PROG_SPEC_MANU_CONCE_PARAM,  # 製造工程
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_CAD_PARAM,  # 機械設計
        PROG_SPEC_CAD_PARAM,  # 機構
        PROG_SPEC_OTHERS,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_MATERIALS_PARAM,  # 材料
        PROG_SPEC_METROL_CONTROL_PARAM,  # 控制工程
        PROG_SPEC_MECH_MANU_PARAM,  # 流體
        PROG_SPEC_MECH_MANU_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_MANU_CONCE_PARAM,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_METROL_CONTROL_PARAM,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


def TUM_COMP_MECH(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUM_COMP_MECH'
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

    PROG_SPEC_MECHANICS_PARAM = {
        'Program_Category': 'Mechanics', 'Required_CP': 40}
    PROG_SPEC_INFO_PARAM = {
        'Program_Category': 'Computer Science and Programming', 'Required_CP': 10}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Mathematics', 'Required_CP': 10}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MECHANICS_PARAM,  # 力學
        PROG_SPEC_INFO_PARAM,  # 基礎資工
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_MECHANICS_PARAM,  # 物理
        PROG_SPEC_MECHANICS_PARAM,  # 物理實驗
        PROG_SPEC_OTHERS,  # 機械設計
        PROG_SPEC_OTHERS,  # 機構
        PROG_SPEC_MECHANICS_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_OTHERS,  # 材料
        PROG_SPEC_OTHERS,  # 控制工程
        PROG_SPEC_MECHANICS_PARAM,  # 流體
        PROG_SPEC_MECHANICS_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_INFO_PARAM,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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

# FPSO: https://www.vm.tu-berlin.de/fileadmin/f5/FAKV_Dateien/StuBe_Maschinenbau/Master/Maschinenbau/AMBl._Nr._10_vom_29.03.2019.pdf


def TUBerlin_ME(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUBerlin_ME'
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
        'Program_Category': 'Mathematics', 'Required_CP': 21}
    PROG_SPEC_MECHANICS_PARAM = {
        'Program_Category': 'Mechanik', 'Required_CP': 18}
    PROG_SPEC_FLUID_THERMODYNAMIK_PARAM = {
        'Program_Category': 'Strömungslehre oder Thermodynamik', 'Required_CP': 6}
    PROG_SPEC_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Elektrotechnik', 'Required_CP': 6}
    PROG_SPEC_MESS_SENSORIK_REGELUNG_PARAM = {
        'Program_Category': 'Mess-/Sensor/Regelungstechnik', 'Required_CP': 6}
    PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM = {
        'Program_Category': 'Konstruktionslehre', 'Required_CP': 16}
    PROG_SPEC_INFORMATIONSTECHNIK_PARAM = {
        'Program_Category': 'Informationstechnik', 'Required_CP': 6}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde', 'Required_CP': 6}
    PROG_SPEC_PRODUKTIONSTECHNIK_PARAM = {
        'Program_Category': 'Produktionstechnik', 'Required_CP': 6}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_MECHANICS_PARAM,  # 力學
        PROG_SPEC_FLUID_THERMODYNAMIK_PARAM,  # 流體或熱力學
        PROG_SPEC_ELEKTROTECHNIK_PARAM,  # 基礎電機
        PROG_SPEC_MESS_SENSORIK_REGELUNG_PARAM,  # 測量 感測 控制
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,     # 機械構造 繪圖
        PROG_SPEC_INFORMATIONSTECHNIK_PARAM,    # 基礎資工
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,     # 材料
        PROG_SPEC_PRODUKTIONSTECHNIK_PARAM,  # 製造
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,  # 機械設計
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,  # 機構
        PROG_SPEC_FLUID_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_MESS_SENSORIK_REGELUNG_PARAM,  # 控制工程
        PROG_SPEC_FLUID_THERMODYNAMIK_PARAM,  # 流體
        PROG_SPEC_MECHANICS_PARAM,  # 力學,機械
        PROG_SPEC_ELEKTROTECHNIK_PARAM,  # 基礎電機電子
        PROG_SPEC_PRODUKTIONSTECHNIK_PARAM,  # 製造
        PROG_SPEC_INFORMATIONSTECHNIK_PARAM,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_MESS_SENSORIK_REGELUNG_PARAM,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


# https://www.maschinenbau.rwth-aachen.de/cms/Maschinenbau/Studium/Im-Studium/~mjdb/Bewerbung-fuer-einen-Masterstudiengang-d/
# https://www.rwth-aachen.de/global/show_document.asp?id=aaaaaaaaaitfehe
def RWTH_ME(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'RWTH_ME'
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

    PROG_SPEC_MECHANIK_PARAM = {
        'Program_Category': 'Mechanik I/II/III', 'Required_CP': 18}
    PROG_SPEC_MASCHINENGESTALTUNG_PARAM = {
        'Program_Category': 'CAD + Maschinengestaltung I/II/III', 'Required_CP': 13}
    PROG_SPEC_THERMODYNAMIK_PARAM = {
        'Program_Category': 'Thermodynamik I + II', 'Required_CP': 7}
    PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM = {
        'Program_Category': 'Wärm_und_Stoffübertragung', 'Required_CP': 6}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde I/II', 'Required_CP': 8}
    PROG_SPEC_CONTROL_TECHNIQUE_PARAM = {
        'Program_Category': 'Regelungstechnik', 'Required_CP': 6}
    PROG_SPEC_STROEMUNGSMECHANIK_PARAM = {
        'Program_Category': 'Strömungsmechanik I', 'Required_CP': 6}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Höhere Mathematik I/II/III', 'Required_CP': 17}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MECHANIK_PARAM,  # 力學
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機械繪圖
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機械設計
        PROG_SPEC_MASCHINENGESTALTUNG_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_WARMSTOFFUBERTRAGUNG_PARAM,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_MECHANIK_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_OTHERS,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


# https: // www.tu-braunschweig.de/fileadmin/Redaktionsgruppen/Fakultaeten/FK4/studierende/Master_MB_ab_WS1415/ZO/BZO.pdf
def TUBraunschweig_ME(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TUBraunschweig_ME'
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

    PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM = {
        'Program_Category': 'CAD Konstruktionslehre', 'Required_CP': 6}
    PROG_SPEC_INFORMATIONSTECHNIK_PARAM = {
        'Program_Category': 'Informationstechnik', 'Required_CP': 4}
    PROG_SPEC_MATH_PARAM = {
        'Program_Category': 'Höhere Mathematik', 'Required_CP': 12}
    PROG_SPEC_CONTROL_TECHNIQUE_PARAM = {
        'Program_Category': 'Regelungstechnik', 'Required_CP': 4}
    PROG_SPEC_TECHNISCHE_MECHANIK_PARAM = {
        'Program_Category': 'Technische Mechanik', 'Required_CP': 12}
    PROG_SPEC_THERMODYNAMIK_PARAM = {
        'Program_Category': 'Thermodynamik I + II', 'Required_CP': 4}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde I/II', 'Required_CP': 4}

    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,  # 機構 繪圖
        PROG_SPEC_INFORMATIONSTECHNIK_PARAM,  # 基礎資工
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_TECHNISCHE_MECHANIK_PARAM,  # 力學
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,  # 機械設計
        PROG_SPEC_KONSTRUKTIONSLEHRE_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMIK_PARAM,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_CONTROL_TECHNIQUE_PARAM,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_TECHNISCHE_MECHANIK_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_INFORMATIONSTECHNIK_PARAM,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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

# https://www.mach.kit.edu/download/MACH-BSc-de_82-604-H-20165_2020-09-15.pdf   page 32
# https://www.sle.kit.edu/downloads/studiengaenge/KIT_Maschinenbau_BA_MA.pdf  page 17


def KIT_ME(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'KIT_ME'
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
        'Program_Category': 'Höhere Mathematik I/II/III', 'Required_CP': 21}
    PROG_SPEC_MECHANIK_PARAM = {
        'Program_Category': 'Technische Mechanik', 'Required_CP': 23}
    PROG_SPEC_MASCHINENKONSTRUKTIONSLEHRE_PARAM = {
        'Program_Category': 'Maschinenkonstruktionslehre', 'Required_CP': 20}
    PROG_SPEC_WERKSTOFFKUNDE_PARAM = {
        'Program_Category': 'Werkstoffkunde I/II + Internship/Project', 'Required_CP': 14}
    PROG_SPEC_THERMODYNAMIK_UBERTRAGUNG_PARAM = {
        'Program_Category': 'Thermodynamik I + II + Wärm_und_Stoffübertragung', 'Required_CP': 15}
    PROG_SPEC_STROEMUNGSMECHANIK_PARAM = {
        'Program_Category': 'Strömungsmechanik I', 'Required_CP': 8}
    PROG_SPEC_ELEKTROTECHNIK_PARAM = {
        'Program_Category': 'Elektrotechnik', 'Required_CP': 8}
    PROG_SPEC_MESS_REGELUNG_PARAM = {
        'Program_Category': 'Mess-/Regelungstechnik', 'Required_CP': 7}
    PROG_SPEC_INFORMATIK_PARAM = {
        'Program_Category': 'Informatik', 'Required_CP': 6}

    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_MECHANIK_PARAM,  # 力學
        PROG_SPEC_MASCHINENKONSTRUKTIONSLEHRE_PARAM,    # 機械構造/ 繪圖
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_THERMODYNAMIK_UBERTRAGUNG_PARAM,  # 熱力學 傳導
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_ELEKTROTECHNIK_PARAM,  # 基礎電機電子
        PROG_SPEC_MESS_REGELUNG_PARAM,  # 測量 控制
        PROG_SPEC_INFORMATIK_PARAM,     # 基礎資工
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_MASCHINENKONSTRUKTIONSLEHRE_PARAM,  # 機械設計
        PROG_SPEC_MASCHINENKONSTRUKTIONSLEHRE_PARAM,  # 機構
        PROG_SPEC_THERMODYNAMIK_UBERTRAGUNG_PARAM,  # 熱力學
        PROG_SPEC_THERMODYNAMIK_UBERTRAGUNG_PARAM,  # 熱 物質傳導
        PROG_SPEC_WERKSTOFFKUNDE_PARAM,  # 材料
        PROG_SPEC_MESS_REGELUNG_PARAM,  # 控制工程
        PROG_SPEC_STROEMUNGSMECHANIK_PARAM,  # 流體
        PROG_SPEC_MECHANIK_PARAM,  # 力學,機械
        PROG_SPEC_ELEKTROTECHNIK_PARAM,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_INFORMATIK_PARAM,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_MESS_REGELUNG_PARAM,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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

# https://www.e-technik.tu-dortmund.de/cms1/Medienpool/Lehre_Studium/Entwuerfe_PO_2019/MPO-A_R.pdf


def TU_DORTMUND_ROBOTICS(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'TU_DORTMUND_ROBOTICS'
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
        'Program_Category': 'Mathematik', 'Required_CP': 18}
    PROG_SPEC_INFORMATIK_PARAM = {
        'Program_Category': 'Computer Science', 'Required_CP': 12}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_INFORMATIK_PARAM,     # 基礎資工
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PARAM,  # 微積分
        PROG_SPEC_MATH_PARAM,  # 數學
        PROG_SPEC_OTHERS,  # 物理
        PROG_SPEC_OTHERS,  # 物理實驗
        PROG_SPEC_OTHERS,  # 機械設計
        PROG_SPEC_OTHERS,  # 機構
        PROG_SPEC_OTHERS,  # 熱力學
        PROG_SPEC_OTHERS,  # 熱 物質傳導
        PROG_SPEC_OTHERS,  # 材料
        PROG_SPEC_OTHERS,  # 控制工程
        PROG_SPEC_OTHERS,  # 流體
        PROG_SPEC_OTHERS,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_OTHERS,  # 製造
        PROG_SPEC_INFORMATIK_PARAM,  # 計算機概論
        PROG_SPEC_OTHERS,  # 機電
        PROG_SPEC_OTHERS,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


# https://www.academy.rwth-aachen.de/en/programs/masters-degree-programs/detail/msc-robotic-systems-engineering?file=files/downloads/masters-degree-programs/full-time-masters-degree-programs/academic-documents/robosys-reading-aid-examination-regulations-rwth-international-academy.pdf&cid=4291
# Fundamentals of Engineering Sciences:
# -Mechanics I, Mechanics II, Mechanics III
# -Machine Design
# -Mechatronics
# -Material Science
# -Thermodynamics
# -Design Theory
# -Manufacturing Technology
# -Introduction to CAD
# 60 CP
# Fundamentals of Mathematics and Natural Sciences:
# -Physics
# -Mathematics, Mathematics II, Mathematics III
# -Numerical Mathematics
# 25 CP
# Fundamentals of System Sciences:
# -Computer Science / Software Techniques / Programming
# -Sensor Techniques / Metrology
# -Simulation Technology / FEM / Multi-Body Simulation-Control Engineering
# 25 CP
def RWTH_ROBOTICS(transcript_sorted_group_map, df_transcript_array, df_category_courses_sugesstion_data, writer):
    program_name = 'RWTH_ROBOTICS'
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
    PROG_SPEC_MATH_PHY_PARAM = {
        'Program_Category': 'Mathematics and Natural Science', 'Required_CP': 25}
    PROG_SPEC_ENG_SCIENCE_PARAM = {
        'Program_Category': 'Fundamentals of Engineering Sciences', 'Required_CP': 60}
    PROG_SPEC_SYSTEM_SCIENCE_PARAM = {
        'Program_Category': 'Fundamentals of System Sciences', 'Required_CP': 25}
    PROG_SPEC_OTHERS = {
        'Program_Category': 'Others', 'Required_CP': 0}

    # This fixed to program course category.
    program_category = [
        PROG_SPEC_MATH_PHY_PARAM,  # 數學
        PROG_SPEC_ENG_SCIENCE_PARAM,     # 機械主修
        PROG_SPEC_SYSTEM_SCIENCE_PARAM,  # 計概 測量 FEM 模擬 控制
        PROG_SPEC_OTHERS  # 其他
    ]

    # Mapping table: same dimension as transcript_sorted_group/ The length depends on how fine the transcript is classified
    program_category_map = [
        PROG_SPEC_MATH_PHY_PARAM,  # 微積分
        PROG_SPEC_MATH_PHY_PARAM,  # 數學
        PROG_SPEC_MATH_PHY_PARAM,  # 物理
        PROG_SPEC_MATH_PHY_PARAM,  # 物理實驗
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 機械設計
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 機構
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 熱力學
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 熱 物質傳導
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 材料
        PROG_SPEC_SYSTEM_SCIENCE_PARAM,  # 控制工程
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 流體
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 力學,機械
        PROG_SPEC_OTHERS,  # 基礎電機電子
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 製造
        PROG_SPEC_SYSTEM_SCIENCE_PARAM,  # 計算機概論
        PROG_SPEC_ENG_SCIENCE_PARAM,  # 機電
        PROG_SPEC_SYSTEM_SCIENCE_PARAM,  # 測量
        PROG_SPEC_OTHERS,  # 車輛
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


program_sort_function = [RWTH_AUTO, TUM_MW, TUHH_MECHATRONICS,
                         UNI_HANNOVER_INTER_MECHATRONICS,
                         TU_DORTMUND_MANUFAC_TECH,
                         TU_CHEMNITZ_AD_MANUFAC,
                         TUM_COMP_MECH,
                         TUBerlin_ME, RWTH_ME,
                         TUBraunschweig_ME,
                         KIT_ME,
                         TU_DORTMUND_ROBOTICS,
                         RWTH_ROBOTICS,
                         ]


def ME_sorter(program_idx, file_path, abbrev):

    basic_classification_en = {
            '微積分': [ME_CALCULUS_KEY_WORDS, ME_CALCULUS_ANTI_KEY_WORDS, ['一', '二']],
            '數學': [ME_MATH_KEY_WORDS, ME_MATH_ANTI_KEY_WORDS],
            '物理': [ME_PHYSICS_KEY_WORDS, ME_PHYSICS_ANTI_KEY_WORDS, ['一', '二']],
            '物理實驗': [ME_PHYSICS_EXP_KEY_WORDS, ME_PHYSICS_EXP_ANTI_KEY_WORDS, ['一', '二']],
            '機械設計': [ME_MASCHINENGESTALTUNG_KEY_WORDS, ME_MASCHINENGESTALTUNG_ANTI_KEY_WORDS, ['一', '二']],
            '機構': [ME_MASCHINEN_ELEMENTE_KEY_WORDS, ME_MASCHINEN_ELEMENTE_ANTI_KEY_WORDS, ['一', '二']],
            '熱力學': [ME_THERMODYN_KEY_WORDS, ME_THERMODYN_ANTI_KEY_WORDS, ['一', '二']],
            '熱傳導': [ME_WARMTRANSPORT_KEY_WORDS, ME_WARMTRANSPORT_ANTI_KEY_WORDS, ['一', '二']],
            '材料': [ME_WERKSTOFFKUNDE_KEY_WORDS, ME_WERKSTOFFKUNDE_ANTI_KEY_WORDS, ['一', '二']],
            '控制系統': [ME_CONTROL_THEORY_KEY_WORDS, ME_CONTROL_THEORY_ANTI_KEY_WORDS, ['一', '二']],
            '流體': [ME_FLUIDDYN_KEY_WORDS, ME_FLUIDDYN_ANTI_KEY_WORDS, ['一', '二']],
            '力學': [ME_MECHANIK_KEY_WORDS, ME_MECHANIK_ANTI_KEY_WORDS],
            '基礎電機電子': [ME_ELECTRICAL_ENG_KEY_WORDS, ME_ELECTRICAL_ENG_ANTI_KEY_WORDS],
            '製造工程': [ME_MANUFACTURE_ENG_KEY_WORDS, ME_MANUFACTURE_ENG_ANTI_KEY_WORDS],
            '計算機概論': [ME_COMPUTER_SCIENCE_KEY_WORDS, ME_COMPUTER_SCIENCE_ANTI_KEY_WORDS],
            '機電': [ME_MECHATRONICS_KEY_WORDS, ME_MECHATRONICS_ANTI_KEY_WORDS],
            '測量': [ME_MEASUREMENT_KEY_WORDS, ME_MEASUREMENT_ANTI_KEY_WORDS],
            '車輛': [ME_VEHICLE_KEY_WORDS, ME_VEHICLE_ANTI_KEY_WORDS],
            '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    basic_classification_zh = {
            '微積分': [ME_CALCULUS_KEY_WORDS, ME_CALCULUS_ANTI_KEY_WORDS, ['一', '二']],
            '數學': [ME_MATH_KEY_WORDS, ME_MATH_ANTI_KEY_WORDS],
            '物理': [ME_PHYSICS_KEY_WORDS, ME_PHYSICS_ANTI_KEY_WORDS, ['一', '二']],
            '物理實驗': [ME_PHYSICS_EXP_KEY_WORDS, ME_PHYSICS_EXP_ANTI_KEY_WORDS, ['一', '二']],
            '機械設計': [ME_MASCHINENGESTALTUNG_KEY_WORDS, ME_MASCHINENGESTALTUNG_ANTI_KEY_WORDS, ['一', '二']],
            '機構': [ME_MASCHINEN_ELEMENTE_KEY_WORDS, ME_MASCHINEN_ELEMENTE_ANTI_KEY_WORDS, ['一', '二']],
            '熱力學': [ME_THERMODYN_KEY_WORDS, ME_THERMODYN_ANTI_KEY_WORDS, ['一', '二']],
            '熱傳導': [ME_WARMTRANSPORT_KEY_WORDS, ME_WARMTRANSPORT_ANTI_KEY_WORDS, ['一', '二']],
            '材料': [ME_WERKSTOFFKUNDE_KEY_WORDS, ME_WERKSTOFFKUNDE_ANTI_KEY_WORDS, ['一', '二']],
            '控制系統': [ME_CONTROL_THEORY_KEY_WORDS, ME_CONTROL_THEORY_ANTI_KEY_WORDS, ['一', '二']],
            '流體': [ME_FLUIDDYN_KEY_WORDS, ME_FLUIDDYN_ANTI_KEY_WORDS, ['一', '二']],
            '力學': [ME_MECHANIK_KEY_WORDS, ME_MECHANIK_ANTI_KEY_WORDS],
            '基礎電機電子': [ME_ELECTRICAL_ENG_KEY_WORDS, ME_ELECTRICAL_ENG_ANTI_KEY_WORDS],
            '製造工程': [ME_MANUFACTURE_ENG_KEY_WORDS, ME_MANUFACTURE_ENG_ANTI_KEY_WORDS],
            '計算機概論': [ME_COMPUTER_SCIENCE_KEY_WORDS, ME_COMPUTER_SCIENCE_ANTI_KEY_WORDS],
            '機電': [ME_MECHATRONICS_KEY_WORDS, ME_MECHATRONICS_ANTI_KEY_WORDS],
            '測量': [ME_MEASUREMENT_KEY_WORDS, ME_MEASUREMENT_ANTI_KEY_WORDS],
            '車輛': [ME_VEHICLE_KEY_WORDS, ME_VEHICLE_ANTI_KEY_WORDS],
            '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    Classifier(program_idx, file_path, abbrev, env_file_path,
               basic_classification_en, basic_classification_zh, column_len_array, program_sort_function)

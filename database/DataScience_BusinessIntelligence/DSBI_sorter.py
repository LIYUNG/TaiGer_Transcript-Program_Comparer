import xlsxwriter
from CourseSuggestionAlgorithms import *
from util import *
from database.DataScience_BusinessIntelligence.DSBI_KEYWORDS import *
from database.DataScience_BusinessIntelligence.DSBI_Programs import program_sort_function, column_len_array
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import pandas as pd
import sys
import os
env_file_path = os.path.realpath(__file__)
env_file_path = os.path.dirname(env_file_path)


def DSBI_sorter(program_idx, file_path, abbrev, Generated_Version):

    basic_classification_en = {
        '微積分': [DSBI_CALCULUS_KEY_WORDS_EN, DSBI_CALCULUS_ANTI_KEY_WORDS_EN],
        '數學': [DSBI_MATH_KEY_WORDS_EN, DSBI_MATH_ANTI_KEY_WORDS_EN],
        '經濟': [DSBI_ECONOMICS_KEY_WORDS_EN, DSBI_ECONOMICS_ANTI_KEY_WORDS_EN],
        '企業': [DSBI_BUSINESS_KEY_WORDS_EN, DSBI_BUSINESS_ANTI_KEY_WORDS_EN],
        '管理': [DSBI_MANAGEMENT_KEY_WORDS_EN, DSBI_MANAGEMENT_ANTI_KEY_WORDS_EN],
        '會計': [DSBI_ACCOUNTING_KEY_WORDS_EN, DSBI_ACCOUNTING_ANTI_KEY_WORDS_EN],
        '統計': [DSBI_STATISTICS_KEY_WORDS_EN, DSBI_STATISTICS_ANTI_KEY_WORDS_EN],
        '金融': [DSBI_FINANCE_KEY_WORDS_EN, DSBI_FINANCE_ANTI_KEY_WORDS_EN],
        '行銷': [DSBI_MARKETING_KEY_WORDS_EN, DSBI_MARKETING_ANTI_KEY_WORDS_EN],
        '作業研究': [DSBI_OP_RESEARCH_KEY_WORDS_EN, DSBI_OP_RESEARCH_ANTI_KEY_WORDS_EN],
        '觀察研究': [DSBI_EP_RESEARCH_KEY_WORDS_EN, DSBI_EP_RESEARCH_ANTI_KEY_WORDS_EN],
        '資工': [DSBI_BASIC_CS_KEY_WORDS_EN, DSBI_BASIC_CS_ANTI_KEY_WORDS_EN],
        '程式': [DSBI_PROGRAMMING_KEY_WORDS_EN, DSBI_PROGRAMMING_ANTI_KEY_WORDS_EN],
        '資料科學': [DSBI_DATA_SCIENCE_KEY_WORDS_EN, DSBI_DATA_SCIENCE_ANTI_KEY_WORDS_EN],
        '論文': [DSBI_BACHELOR_THESIS_KEY_WORDS_EN, DSBI_BACHELOR_THESIS_ANTI_KEY_WORDS_EN],
        '其他': [USELESS_COURSES_KEY_WORDS_EN, USELESS_COURSES_ANTI_KEY_WORDS_EN], }

    basic_classification_zh = {
        '微積分': [DSBI_CALCULUS_KEY_WORDS, DSBI_CALCULUS_ANTI_KEY_WORDS],
        '數學': [DSBI_MATH_KEY_WORDS, DSBI_MATH_ANTI_KEY_WORDS],
        '經濟': [DSBI_ECONOMICS_KEY_WORDS, DSBI_ECONOMICS_ANTI_KEY_WORDS],
        '企業': [DSBI_BUSINESS_KEY_WORDS, DSBI_BUSINESS_ANTI_KEY_WORDS],
        '管理': [DSBI_MANAGEMENT_KEY_WORDS, DSBI_MANAGEMENT_ANTI_KEY_WORDS],
        '會計': [DSBI_ACCOUNTING_KEY_WORDS, DSBI_ACCOUNTING_ANTI_KEY_WORDS],
        '統計': [DSBI_STATISTICS_KEY_WORDS, DSBI_STATISTICS_ANTI_KEY_WORDS],
        '金融': [DSBI_FINANCE_KEY_WORDS, DSBI_FINANCE_ANTI_KEY_WORDS],
        '行銷': [DSBI_MARKETING_KEY_WORDS, DSBI_MARKETING_ANTI_KEY_WORDS],
        '作業研究': [DSBI_OP_RESEARCH_KEY_WORDS, DSBI_OP_RESEARCH_ANTI_KEY_WORDS],
        '觀察研究': [DSBI_EP_RESEARCH_KEY_WORDS, DSBI_EP_RESEARCH_ANTI_KEY_WORDS],
        '資工': [DSBI_BASIC_CS_KEY_WORDS, DSBI_BASIC_CS_ANTI_KEY_WORDS],
        '程式': [DSBI_PROGRAMMING_KEY_WORDS, DSBI_PROGRAMMING_ANTI_KEY_WORDS],
        '資料科學': [DSBI_DATA_SCIENCE_KEY_WORDS, DSBI_DATA_SCIENCE_ANTI_KEY_WORDS],
        '論文': [DSBI_BACHELOR_THESIS_KEY_WORDS, DSBI_BACHELOR_THESIS_ANTI_KEY_WORDS],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    Classifier(program_idx, file_path, abbrev, env_file_path,
               basic_classification_en, basic_classification_zh, column_len_array, program_sort_function, Generated_Version)

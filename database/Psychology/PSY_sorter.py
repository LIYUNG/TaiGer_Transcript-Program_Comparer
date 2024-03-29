import xlsxwriter
from CourseSuggestionAlgorithms import *
from util import *
from database.Psychology.PSY_KEYWORDS import *
from database.Psychology.PSY_Programs import program_sort_function, column_len_array
from cell_formatter import red_out_failed_subject, red_out_insufficient_credit
import pandas as pd
import sys
import os
env_file_path = os.path.realpath(__file__)
env_file_path = os.path.dirname(env_file_path)


def PSY_sorter(program_idx, file_path, abbrev, Generated_Version):
    basic_classification_en = {
        '微積分': [PSY_CALCULUS_KEY_WORDS_EN, PSY_CALCULUS_ANTI_KEY_WORDS_EN],
        '線性代數': [PSY_MATH_LIN_KEY_WORDS_EN, PSY_MATH_LIN_ANTI_KEY_WORDS_EN],
        '機率': [PSY_MATH_PROB_KEY_WORDS_EN, PSY_MATH_PROB_ANTI_KEY_WORDS_EN],
        '數學': [PSY_MATH_KEY_WORDS_EN, PSY_MATH_ANTI_KEY_WORDS_EN],
        '經濟': [PSY_ECONOMICS_KEY_WORDS_EN, PSY_ECONOMICS_ANTI_KEY_WORDS_EN],
        '計量經濟': [PSY_ECONOMETRICS_KEY_WORDS_EN, PSY_ECONOMETRICS_ANTI_KEY_WORDS_EN],
        '企業': [PSY_BUSINESS_KEY_WORDS_EN, PSY_BUSINESS_ANTI_KEY_WORDS_EN],
        '管理': [PSY_MANAGEMENT_KEY_WORDS_EN, PSY_MANAGEMENT_ANTI_KEY_WORDS_EN],
        '統計': [PSY_STATISTICS_KEY_WORDS_EN, PSY_STATISTICS_ANTI_KEY_WORDS_EN],
        '金融': [PSY_FINANCE_KEY_WORDS_EN, PSY_FINANCE_ANTI_KEY_WORDS_EN],
        '行銷': [PSY_MARKETING_KEY_WORDS_EN, PSY_MARKETING_ANTI_KEY_WORDS_EN],
        '作業研究': [PSY_OP_RESEARCH_KEY_WORDS_EN, PSY_OP_RESEARCH_ANTI_KEY_WORDS_EN],
        '觀察研究': [PSY_EP_RESEARCH_KEY_WORDS_EN, PSY_EP_RESEARCH_ANTI_KEY_WORDS_EN],
        '程式': [PSY_PROGRAMMING_KEY_WORDS_EN, PSY_PROGRAMMING_ANTI_KEY_WORDS_EN],
        '資料科學': [PSY_DATA_SCIENCE_KEY_WORDS_EN, PSY_DATA_SCIENCE_ANTI_KEY_WORDS_EN],
        '論文': [PSY_BACHELOR_THESIS_KEY_WORDS_EN, PSY_BACHELOR_THESIS_ANTI_KEY_WORDS_EN],
        '心理學': [PSY_PSYCHOLOGY_KEY_WORDS_EN, PSY_PSYCHOLOGY_ANTI_KEY_WORDS_EN],
        '心理學實驗': [PSY_PSYCHOLOGY_EXP_KEY_WORDS_EN, PSY_PSYCHOLOGY_EXP_ANTI_KEY_WORDS_EN],
        '認知': [PSY_COGNITIVE_SCIENCE_KEY_WORDS_EN, PSY_COGNITIVE_SCIENCE_ANTI_KEY_WORDS_EN],
        '行為': [PSY_BEHAVIOR_KEY_WORDS_EN, PSY_BEHAVIOR_ANTI_KEY_WORDS_EN],
        '神經科學': [PSY_NEURO_SCIENCE_KEY_WORDS_EN, PSY_NEURO_SCIENCE_ANTI_KEY_WORDS_EN],
        '其他': [USELESS_COURSES_KEY_WORDS_EN, USELESS_COURSES_ANTI_KEY_WORDS_EN], }

    basic_classification_zh = {
        '微積分': [PSY_CALCULUS_KEY_WORDS, PSY_CALCULUS_ANTI_KEY_WORDS, ['一', '二']],
        '線性代數': [PSY_MATH_LIN_KEY_WORDS, PSY_MATH_LIN_ANTI_KEY_WORDS],
        '機率': [PSY_MATH_PROB_KEY_WORDS, PSY_MATH_PROB_ANTI_KEY_WORDS],
        '數學': [PSY_MATH_KEY_WORDS, PSY_MATH_ANTI_KEY_WORDS, ['一', '二']],
        '經濟': [PSY_ECONOMICS_KEY_WORDS, PSY_ECONOMICS_ANTI_KEY_WORDS, ['一', '二']],
        '計量經濟': [PSY_ECONOMETRICS_KEY_WORDS, PSY_ECONOMETRICS_ANTI_KEY_WORDS],
        '企業': [PSY_BUSINESS_KEY_WORDS, PSY_BUSINESS_ANTI_KEY_WORDS],
        '管理': [PSY_MANAGEMENT_KEY_WORDS, PSY_MANAGEMENT_ANTI_KEY_WORDS, ['一', '二']],
        '統計': [PSY_STATISTICS_KEY_WORDS, PSY_STATISTICS_ANTI_KEY_WORDS, ['一', '二']],
        '金融': [PSY_FINANCE_KEY_WORDS, PSY_FINANCE_ANTI_KEY_WORDS],
        '行銷': [PSY_MARKETING_KEY_WORDS, PSY_MARKETING_ANTI_KEY_WORDS],
        '作業研究': [PSY_OP_RESEARCH_KEY_WORDS, PSY_OP_RESEARCH_ANTI_KEY_WORDS, ['一', '二']],
        '觀察研究': [PSY_EP_RESEARCH_KEY_WORDS, PSY_EP_RESEARCH_ANTI_KEY_WORDS, ['一', '二']],
        '程式': [PSY_PROGRAMMING_KEY_WORDS, PSY_PROGRAMMING_ANTI_KEY_WORDS, ['一', '二']],
        '資料科學': [PSY_DATA_SCIENCE_KEY_WORDS, PSY_DATA_SCIENCE_ANTI_KEY_WORDS],
        '論文': [PSY_BACHELOR_THESIS_KEY_WORDS, PSY_BACHELOR_THESIS_ANTI_KEY_WORDS, ['一', '二']],
        '心理學': [PSY_PSYCHOLOGY_KEY_WORDS, PSY_PSYCHOLOGY_ANTI_KEY_WORDS, ['一', '二']],
        '心理學實驗': [PSY_PSYCHOLOGY_EXP_KEY_WORDS, PSY_PSYCHOLOGY_EXP_ANTI_KEY_WORDS, ['一', '二']],
        '認知': [PSY_COGNITIVE_SCIENCE_KEY_WORDS, PSY_COGNITIVE_SCIENCE_ANTI_KEY_WORDS, ['一', '二']],
        '行為': [PSY_BEHAVIOR_KEY_WORDS, PSY_BEHAVIOR_ANTI_KEY_WORDS, ['一', '二']],
        '神經科學': [PSY_NEURO_SCIENCE_KEY_WORDS, PSY_NEURO_SCIENCE_ANTI_KEY_WORDS, ['一', '二']],
        '其他': [USELESS_COURSES_KEY_WORDS, USELESS_COURSES_ANTI_KEY_WORDS], }

    Classifier(program_idx, file_path, abbrev, env_file_path,
               basic_classification_en, basic_classification_zh, column_len_array, program_sort_function, Generated_Version)

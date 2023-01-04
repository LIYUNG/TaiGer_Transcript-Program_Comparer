KEY_WORDS = 0
ANTI_KEY_WORDS = 1
DIFFERENTIATE_KEY_WORDS = 2

#################################################################################################################
########################################## Electrical Engineering ###############################################
#################################################################################################################

# TODO: defining the keywords in proper way (iterative steps: generating, and see result, if mssing, then add keywords/anti_keywords)
TE_CALCULUS_KEY_WORDS = ['微積分']
TE_CALCULUS_ANTI_KEY_WORDS = ['asdgladfj;l']
TE_MATH_KEY_WORDS = ['數學', '代數', '微分', '函數', '機率', '離散', '複變', '數值', '向量']
TE_MATH_ANTI_KEY_WORDS = ['asdgladfj;l']
TE_INTRO_COMPUTER_SCIENCE_KEY_WORDS = ['計算機', '演算', '資料', '物件', '運算',
                                       '資電', '作業系統', '資料結構', '編譯器']
TE_INTRO_COMPUTER_SCIENCE_ANTI_KEY_WORDS = ['倫理', 'Python', '機器學習', '傳輸']
TE_PROGRAMMING_KEY_WORDS = ['程式設計', '程式語言', 'Python', 'C++', 'C語言']
TE_PROGRAMMING_ANTI_KEY_WORDS = ['倫理', 'Python', '機器學習']
TE_SOFTWARE_SYSTEM_KEY_WORDS = ['軟體工程', ]
TE_SOFTWARE_SYSTEM_ANTI_KEY_WORDS = ['asdgladfj;l']
TE_CONTROL_THEORY_KEY_WORDS = ['控制', '線性系統', '非線性系統']
TE_CONTROL_THEORY_ANTI_KEY_WORDS = ['asdgladfj;l']
TE_PHYSICS_KEY_WORDS = ['物理']
TE_PHYSICS_ANTI_KEY_WORDS = ['半導體', '元件', '實驗']
TE_PHYSICS_EXP_KEY_WORDS = ['物理']
TE_PHYSICS_EXP_ANTI_KEY_WORDS = ['半導體', '元件']
TE_ELECTRONICS_KEY_WORDS = ['電子']
TE_ELECTRONICS_ANTI_KEY_WORDS = [
    '實驗', '專題', '電力', '固態', '自動化', '倫理', '素養', '進階', '電路', '實驗']
TE_ELECTRONICS_EXP_KEY_WORDS = ['電子實驗', '電子']
TE_ELECTRONICS_EXP_ANTI_KEY_WORDS = ['專題', '電力', '固態']
TE_ELECTRO_CIRCUIT_KEY_WORDS = ['電路', '數位邏輯', '邏輯設計']
TE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS = ['超大型', '專題', '倫理', '素養', '進階', '高頻']
TE_SIGNAL_SYSTEM_KEY_WORDS = ['訊號與系統', '信號與系統', '訊號', '信號']
TE_SIGNAL_SYSTEM_ANTI_KEY_WORDS = ['asdgladfj;l''超大型', '專題']
TE_ELECTRO_MAGNET_KEY_WORDS = ['電磁']
TE_ELECTRO_MAGNET_ANTI_KEY_WORDS = ['asdgladfj;l', '專題', '進階']
TE_POWER_ELECTRO_KEY_WORDS = ['電力', '能源', '電動機', '電機', '高壓電', '電傳輸', '配電']
TE_POWER_ELECTRO_ANTI_KEY_WORDS = ['asdgladfj;l', '專題', '進階']
TE_COMMUNICATION_KEY_WORDS = ['高頻', '天線', '微波', '密碼學', '安全', '傳輸', '射頻',
                              '網路', '無線', '通信', '通訊', '電波', '無線網路', '消息']
TE_COMMUNICATION_ANTI_KEY_WORDS = ['asdgladfj;l', '專題', '進階']
TE_SEMICONDUCTOR_KEY_WORDS = ['半導體', '元件', '固態']
TE_SEMICONDUCTOR_ANTI_KEY_WORDS = ['專題', '倫理', '素養']
TE_ADVANCED_ELECTRO_KEY_WORDS = ['積體電路', '自動化',  '藍芽', '晶片', '類比', '數位訊號', '數位信號',
                                 '微算機', '微處理', 'VLSI', '嵌入式', '人工智慧', '機器學習']
TE_ADVANCED_ELECTRO_ANTI_KEY_WORDS = ['倫理', '素養']
TE_APPLICATION_ORIENTED_KEY_WORDS = ['生醫', '光機電', '電腦', '微系統', '物聯網', '聲學',
                                     '影像', '深度學習', '光電', '應用', '綠能', '雲端運算', '醫學工程', '再生能源']
TE_APPLICATION_ORIENTED_ANTI_KEY_WORDS = ['倫理', '素養']
TE_MACHINE_RELATED_KEY_WORDS = ['力學', '流體',
                                '熱力', '傳導', '熱傳', '機械', '動力', '熱流', '機動']
TE_MACHINE_RELATED_ANTI_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_ANTI_KEY_WORDS = ['asdgladfj;l']

#################################################################################################################
################################### Electrical Engineering English ##############################################
#################################################################################################################

# TODO: defining the keywords in proper way (iterative steps: generating, and see result, if mssing, then add keywords/anti_keywords)
TE_CALCULUS_KEY_WORDS_EN = ['calculus']
TE_CALCULUS_ANTI_KEY_WORDS_EN = ['asdgladfj;l']
TE_MATH_KEY_WORDS_EN = ['mathe', 'algebra', 'differential', '函數',
                        'probability', 'discrete', 'complex', 'numerical', 'vector analy']
TE_MATH_ANTI_KEY_WORDS_EN = ['asdgladfj;l']

TE_INTRO_COMPUTER_SCIENCE_KEY_WORDS_EN = ['computer', 'algorithm', 'object', 'computing',
                                          '資電', 'operating system', 'data structure', 'software', 'compiler']
TE_INTRO_COMPUTER_SCIENCE_ANTI_KEY_WORDS_EN = [
    '倫理', 'Python', 'machine learning']
TE_PROGRAMMING_KEY_WORDS_EN = ['programming', 'program',
                               'language', 'Python', 'C++', 'c language']
TE_PROGRAMMING_ANTI_KEY_WORDS_EN = [
    'ethnic', 'Python', 'machine learning', 'vision']
TE_SOFTWARE_SYSTEM_KEY_WORDS_EN = ['software engineering', ]
TE_SOFTWARE_SYSTEM_ANTI_KEY_WORDS_EN = ['asdgladfj;l']
TE_CONTROL_THEORY_KEY_WORDS_EN = ['control', 'linear system', '非線性系統']
TE_CONTROL_THEORY_ANTI_KEY_WORDS_EN = ['asdgladfj;l']
TE_PHYSICS_KEY_WORDS_EN = ['physics']
TE_PHYSICS_ANTI_KEY_WORDS_EN = ['semicondu', 'device', 'experiment']
TE_PHYSICS_EXP_KEY_WORDS_EN = ['physics']
TE_PHYSICS_EXP_ANTI_KEY_WORDS_EN = ['semicondu', '元件']
TE_ELECTRONICS_KEY_WORDS_EN = ['electronic', 'electrical']
TE_ELECTRONICS_ANTI_KEY_WORDS_EN = [
    '專題', 'power', 'solid', 'automation', 'ethnic', '素養', 'advanced', 'lab']
TE_ELECTRONICS_EXP_KEY_WORDS_EN = ['lab', '電子']
TE_ELECTRONICS_EXP_ANTI_KEY_WORDS_EN = ['physic', 'chemistry', 'general']
TE_ELECTRO_CIRCUIT_KEY_WORDS_EN = [
    'circuit', 'signal', '數位邏輯', 'logic']
TE_ELECTRO_CIRCUIT_ANTI_KEY_WORDS_EN = ['超大型', '專題']
TE_SIGNAL_SYSTEM_KEY_WORDS_EN = ['signals and systems']
TE_SIGNAL_SYSTEM_ANTI_KEY_WORDS_EN = ['asdgladfj;l', '超大型', '專題']
TE_ELECTRO_MAGNET_KEY_WORDS_EN = ['electromagne', 'magne']
TE_ELECTRO_MAGNET_ANTI_KEY_WORDS_EN = ['asdgladfj;l', '專題', '進階']
TE_POWER_ELECTRO_KEY_WORDS_EN = [
    'power', 'energy', '電動機', 'electrical machine', 'high voltage', 'transmission']
TE_POWER_ELECTRO_ANTI_KEY_WORDS_EN = ['asdgladfj;l', '專題', '進階']
TE_COMMUNICATION_KEY_WORDS_EN = ['high frequen', 'antenna', 'microwave', 'crypto', 'security', 'radio frequ',
                                 'network', 'wireless', 'communication', '通訊', '電波', 'information theory']
TE_COMMUNICATION_ANTI_KEY_WORDS_EN = ['asdgladfj;l', '專題', '進階']
TE_SEMICONDUCTOR_KEY_WORDS_EN = ['semicondu', '元件', '固態']
TE_SEMICONDUCTOR_ANTI_KEY_WORDS_EN = ['專題', 'ethnic', '素養']
TE_ADVANCED_ELECTRO_KEY_WORDS_EN = ['integrated circuit', 'automation',  'bluetooth', 'chip', 'analog', 'digital', 'digital signal',
                                    'microprocessor', 'microcontroller', 'VLSI', 'embedded', 'artificial', 'machine learning']
TE_ADVANCED_ELECTRO_ANTI_KEY_WORDS_EN = ['ethnic', '素養']
TE_APPLICATION_ORIENTED_KEY_WORDS_EN = ['生醫', 'neuro', '光機電', 'mems', 'iot', 'accoustics', 'solar',
                                        'image', 'deep learning', 'optoelectronics', '應用', 'green', 'cloud', 'medical', 'renewable']
TE_APPLICATION_ORIENTED_ANTI_KEY_WORDS_EN = ['ethnic', '素養']
TE_MACHINE_RELATED_KEY_WORDS_EN = ['mechanics', 'fluid', "statics", "dynamics",
                                   'thermodyna', '傳導', 'conduction', 'machine', 'heat flux', '機動']
TE_MACHINE_RELATED_ANTI_KEY_WORDS_EN = ['asdgladfj;l']
USELESS_COURSES_KEY_WORDS_EN = ['asdgladfj;l']
USELESS_COURSES_ANTI_KEY_WORDS_EN = ['asdgladfj;l']

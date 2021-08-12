KEY_WORDS = 0
ANTI_KEY_WORDS = 1
DIFFERENTIATE_KEY_WORDS = 2

# TODO: defining the keywords in proper way (iterative steps: generating, and see result, if mssing, then add keywords/anti_keywords)
CALCULUS_KEY_WORDS = ['微積分']
CALCULUS_ANTI_KEY_WORDS = ['asdgladfj;l']
MATH_KEY_WORDS = ['數學', '代數', '微分', '函數', '機率', '離散', '複變', '數值', '向量']
MATH_ANTI_KEY_WORDS = ['asdgladfj;l']
PROGRAMMING_KEY_WORDS = ['計算機', '演算', '資料', '物件', '運算',
                         '資電', '作業系統', '資料結構', '軟體', '編譯器', '程式設計', '程式語言', 'Python', 'C++', 'C語言']
PROGRAMMING_ANTI_KEY_WORDS = ['倫理']
PHYSICS_KEY_WORDS = ['物理']
PHYSICS_ANTI_KEY_WORDS = ['半導體', '元件']
ELECTRONICS_KEY_WORDS = ['電子']
ELECTRONICS_ANTI_KEY_WORDS = ['專題', '電力', '固態', '自動化', '倫理', '素養','進階']
ELECTRO_CIRCUIT_KEY_WORDS = ['電路', '訊號', '數位邏輯', '邏輯設計', '信號與系統']
ELECTRO_CIRCUIT_ANTI_KEY_WORDS = ['超大型', '專題', '倫理', '素養', '進階']
ELECTRO_MAGNET_KEY_WORDS = ['電磁']
ELECTRO_MAGNET_ANTI_KEY_WORDS = ['asdgladfj;l', '專題', '進階']
ADVANCED_ELECTRO_KEY_WORDS = ['微波', '積體電路', '自動化', '天線', '網路', '高頻', '無線', '藍芽', '晶片',
                              '類比', '數位訊號', '通信',  '通訊', '微算機', '微處理', '電波', 'VLSI', '固態', '嵌入式', '人工智慧', '無線網路', '機器學習', '消息']
ADVANCED_ELECTRO_ANTI_KEY_WORDS = ['倫理', '素養']
SEMICONDUCTOR_KEY_WORDS = ['半導體', '元件']
SEMICONDUCTOR_ANTI_KEY_WORDS = ['專題', '倫理', '素養']
APPLICATION_ORIENTED_KEY_WORDS = ['電力', '生醫', '能源', '光機電', '電動機', '電腦', '微系統', '物聯網','密碼學','聲學',
                                  '電機', '影像', '深度學習', '光電', '應用', '綠能', '雲端運算', '醫學工程', '再生能源']
APPLICATION_ORIENTED_ANTI_KEY_WORDS = ['倫理','素養']

CONTROL_THEORY_KEY_WORDS = ['控制']
CONTROL_THEORY_ANTI_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_KEY_WORDS = ['asdgladfj;l']
USELESS_COURSES_ANTI_KEY_WORDS = ['asdgladfj;l']

# Other keywords not relevant to EE major
CHEMISTRY_KEY_WORDS = ['化學']
CHEMISTRY_ANTI_KEY_WORDS = ['asdgladfj;l']
MECHANICS_KEY_WORDS = ['熱力學', '動力', '靜力', '材料力', '摩擦', '流體']
MECHANICS_ANTI_KEY_WORDS = ['asdgladfj;l']

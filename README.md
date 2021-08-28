# TaiGer_Transcript-Program_Comparer

├───database
│   ├───ComputerScience
│   │   ├─── CS_Course_database.xlsx  # All course that students can take in their faculty. (here: Computer Science) 
│   │   ├─── CS_KEYWORDS.py           # Keywords collection for the relevant course category
│   │   ├─── CS_Programs.xlsx         # Program to be analyzed selector
│   │   └─── CS_sorter.py             # main function of the program analyzer for the program (here Computer Science for example)
│   ├───ElectricalEngineering
│   │   ├─── ......                   # (same as above)
│   ├───Management
│   └───MechanicalEngineering
├───CourseSuggestionAlgorithms.py   # Course Suggestion Algorithms
└───main.py                         # Software entry point

## Enviroment :
```
Python 3.8
```

## Requirement Python package (using: pip install <module_name>):
```
xlsxwriter
gc
pandas
```

## How to use?
```
python main.py <transcript_course.xlsx> <ee/cs>
```

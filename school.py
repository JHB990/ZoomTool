import webbrowser
import datetime 
from pandas import *

schedule_file = ExcelFile("horario.xlsx")
schedule_dataframe = schedule_file.parse(schedule_file.sheet_names[0])
schedule_dict = schedule_dataframe.to_dict()
print(schedule_dict)
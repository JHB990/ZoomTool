import webbrowser
import datetime 
from pandas import *

schedule_file = ExcelFile("horario.xlsx")
schedule_dataframe = schedule_file.parse(schedule_file.sheet_names[0])
schedule_dict = schedule_dataframe.to_dict()

days = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"]

def open_class():
    started = False
    i = 0

    while i < len(schedule_dict["Hora"]):
        index_day = datetime.datetime.now().weekday()
        week_day = days[index_day]
        if schedule_dict[week_day][i] == "nan":
            pass
        else:
            for day in schedule_dict.keys():
                if day == week_day:
                    hour = schedule_dict["Hora"][i].strftime("%H:%M:%S")
                    link_key = "Enlace"+week_day
                    link = schedule_dict[link_key][i]
                    
                    while started != True:
                        current_time = take_current_time()
                        if current_time == hour:
                            print("Abriendo clases...")
                            webbrowser.open(link)
                            started = True
                        else:
                            continue
            i += 1
            started = False
def take_current_time():
    return datetime.datetime.now().strftime('%H:%M:%S')

if __name__ == '__main__':
    open_class()
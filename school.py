import webbrowser
import datetime 
#import pyautogui as pytho
import time
from pandas import *

schedule_file = ExcelFile("horario.xlsx")
schedule_dataframe = schedule_file.parse(schedule_file.sheet_names[0]).fillna(0)
schedule_dict = schedule_dataframe.to_dict()


#print(schedule_dict)

days = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"]

def open_class():
    started = False
    i = 0

    while i < len(schedule_dict["Hora"]):
        index_day = datetime.datetime.now().weekday()
        week_day = days[index_day]
        if schedule_dict[week_day][i] == 0:
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
                            brave_path = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe %s --incognito'
                            webbrowser.get(brave_path).open_new(link)
                            started = True
                            #time.sleep(8)
                            #pytho.press('c')
                            #time.sleep(1)
                            #pytho.write("Listo la sesion esta iniciada")
                            #time.sleep(2)
                            #pytho.press("enter")
                        else:
                            continue
            i += 1
            started = False
def take_current_time():
    return datetime.datetime.now().strftime('%H:%M:%S')

if __name__ == '__main__':
    open_class()
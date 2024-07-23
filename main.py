#Устанавливаем pyinstaller(2 строка), компилим приложение через консоль(3 строка)
#pip install pyinstaller
#PyInstaller --onefile --name "UltraPenetrator" ".\main.py"

import serial, datetime, time, re, warnings, logging, openpyxl
import pandas as pd
import flet as ft
from threading import Event, Thread


warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

logging.basicConfig(level=logging.WARNING, filename="./UltraPenetrator_LOG.log",filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

#logging.error("polling failed",exc_info=True)
#logging.warning('File replace - ./log/EVfront.xlsx')

global Cycle_break
Cycle_break = False

def main(page: ft.Page):
    
    GRBL_port_path = 'COM10'
    gcode_path = './source/TEST15.gcode'
    homing_gcode_path = './source/home.gcode'
    Change_List_Gcode_path = './source/change_list.gcode'
    DF = pd.DataFrame()
    

    chart = ft.LineChart(
        data=DF,
        animate=True,
        
        )


    page.title = "UltraPenetrator_HMI_V0.0.7"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = 'dark'
   
    
    
    
    CycleTxt = ft.Text("Цикл выполняется!",visible=False)
    FileSaveTxt = ft.Text("Файл сохранен!",visible=False)
    tenz = ft.Text("Тензодатчик подключен!",visible=False)
    Max_Value_Txt = ft.Text("",visible=False)
    Min_Value_Txt = ft.Text("",visible=False)
    timetostart = ft.Text("",visible=False)
    
    
    def PollingComPortTD(): #Получение данных с тензодатчика
        tenz.visible = True
        global Cycle_break
        page.update()
        s = serial.Serial(      #конфиг порта
                        port='COM4',
                        baudrate = 115200,
                        parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE,
                        bytesize=serial.EIGHTBITS,
                        timeout=1
                )
        nowDateTime = datetime.datetime.now()
        nowTime = nowDateTime.time().strftime('%H-%M') #Вывод часов и минут реального времени
        nowDate = nowDateTime.date() 
        FileName = './test_'+str(nowDate)+'_'+nowTime+'.xlsx' #Формирование имени файла
        StopFileName = './stop_'+str(nowDate)+'_'+nowTime+'.xlsx'
        EmergencyStopName = './Alarm_stop_'+str(nowDate)+'_'+nowTime+'.xlsx'
        timetostartmsg = 'Время начала теста: '+str(nowTime)
        timetostart.value = timetostartmsg
        timetostart.visible = True
        writeCouter = 1 #Каунтер который осуществляет переход на следующий столбец
        ROWCOUNTER = 1 #Счетчик строк
        measure = 0 #Счетчик выполнения измерений
        zero = False #Переменная для корректного обнуления  writeCounter
        res=0
        onecycle = False
        Max_value = 0
        Min_value = 0
        print('Pollingtd started')
        tic=0 #Время начала цикла
        toc=0 #Время окончания
        while True:    
            #try:
            if Cycle_break:  # Проверяем флаг
                    DF.to_excel(StopFileName,sheet_name='Sheet1')
                    break
            grab = s.readline() #Чтение порта
            grab=str(grab)  
            grab=grab.replace(' ','') #Удаление пробелов
            nums = re.findall(r"-?\d+\.\d+", grab)#Удаление символов кроме - и цифр
            print(nums)
            if nums != []:
                
                try:
                    res = nums[0]
                except:
                    logging.error("Nums failed",exc_info=True)
                    print('nums = ',nums)
                result = float(res)
                #startcylcetime = datetime.datetime.now().strftime('%H:%M:%S')#[:-3]
                if float(result) > 0.250: 
                    onecycle = True              #Отсечение наводок и непонятных данных
                    zero = True
                    tic = time.perf_counter()
                    writeCouter +=1
                    DF.at[ROWCOUNTER,writeCouter] = float(result)
                else: 
                    toc = time.perf_counter()
                    cycletime = toc - tic
                    if cycletime >= (0.5) and zero==True: #Если данные из компорта не идут более 0.2 сек и измерения выполнялись хотя бы один раз, то
                        ROWCOUNTER +=1 #Переход на следующую строку в таблице
                        writeCouter = 1 #Возврат к второй колонке таблицы
                        measure +=1 #Счетчик измерений +1
                        zero = False #Выполнение измерений 0
                        
                        Max_value = DF.max(axis=1).max()
                        Min_value = DF.max(axis=1).min()
                        if Max_value >= 7.500: #Усилие для аварийной остановки
                            Cycle_break = True
                            DF.to_excel(EmergencyStopName,sheet_name='Sheet1')
                            break
                        msg_max_value = 'Максимальное усилие прокола: '+ str(Max_value)+' Кг!'
                        msg_min_value = 'Минимальное усилие прокола: '+ str(Min_value)+' Кг!'

                        Max_Value_Txt.value = str(msg_max_value)
                        Max_Value_Txt.visible = True

                        Min_Value_Txt.value = str(msg_min_value)
                        Min_Value_Txt.visible = True

                        page.update()
                    if cycletime >= 60 and onecycle==True:
                        DF.to_excel(FileName,sheet_name='Sheet1')
                        print('complete')
                        onecycle=False
                        FileSaveName = ('File save - ' + str(FileName))
                        logging.warning(FileSaveName)
                        FileSaveTxt.visible = True
                        page.update() 
                        time.sleep(20)
                        FileSaveTxt.visible = False
                        page.update()   
                        return
                            
            #except:
            #    print("алярм нет работы")
            #   traceback.print_tb(limit=None, file=None)
            #   time.sleep(5)



    def stream_grbl_gcode(GRBL_port_path,gcode_path):
        RX_BUFFER_SIZE = 48
        with open(gcode_path, "r") as f, serial.Serial(GRBL_port_path, 115200) as s:            
            l_count = 0
            verbose = True

            # Wake up grbl
            print ("Initializing grbl...")
            s.write(str.encode("\r\n\r\n"))

            # Wait for grbl to initialize and flush startup text in serial input
            time.sleep(4)
            s.flush()

            g_count = 0
            c_line = []
            # periodic() # Start status report periodic timer
            for line in f:
                if Cycle_break:  # Проверяем флаг
                    s.write(str.encode("!"))
                    break
                l_count += 1 # Iterate line counter
                l_block = line.strip()
                c_line.append(len(l_block)+1) # Track number of characters in grbl serial read buffer
                grbl_out = '' 
                while sum(c_line) >= RX_BUFFER_SIZE-1 | s.inWaiting() :
                    out_temp = s.readline().strip().decode('utf-8') # Wait for grbl response
                    if out_temp.find('ok') < 0 and out_temp.find('error') < 0 :
                        print ("  Debug: ",out_temp) # Debug response 
                        if "ALARM" in out_temp:
                            Pstartbtn.disabled = False
                            PChL.disabled = False
                            RefXYZbtn.disabled = False
                            page.update()
                            return
                        #Зона для описания ошибки во время выполнения программы
                    else :
                        grbl_out += out_temp
                        g_count += 1 # Iterate g-code counter
                        grbl_out += str(g_count); # Add line finished indicator
                        del c_line[0] # Delete the block character count corresponding to the last 'ok'
                if verbose: print ("SND: " + str(l_count) + " : " + l_block,)
                s.write(str.encode(l_block + '\n')) # Send g-code block to grbl
                if verbose : print( "BUF:",str(sum(c_line)),"REC:",grbl_out)

        # Wait for user input after streaming is completed
        print ("G-code streaming finished!\n")
        print ("WARNING: Wait until grbl completes buffered g-code blocks before exiting.")
        logging.warning("G-code streaming finished!")

        # Close file and serial port
        f.close()
        s.close()
        print('End of gcode')
        return
    
                        

    def remove_comment(string):
        if (string.find(';') == -1):
            return string
        else:
            return string[:string.index(';')]

    def remove_eol_chars(string):
        # removed \n or traling spaces
        return string.strip()

    def send_wake_up(ser):
        # Wake up
        # Hit enter a few times to wake the Printrbot
        ser.write(str.encode("\r\n\r\n"))
        time.sleep(2)   # Wait for Printrbot to initialize
        ser.flushInput()  # Flush startup text in serial input

    def wait_for_movement_completion(ser,cleaned_line):

        Event().wait(1)

        if cleaned_line != '$X' or '$$':

            idle_counter = 0

            while True:

                # Event().wait(0.01)
                ser.reset_input_buffer()
                command = str.encode('?' + '\n')
                ser.write(command)
                grbl_out = ser.readline() 
                grbl_response = grbl_out.strip().decode('utf-8')

                if grbl_response != 'ok':

                    if grbl_response.find('Idle') > 0:
                        idle_counter += 1

                if idle_counter > 10:
                    break
        return

    def stream_gcode(GRBL_port_path,gcode_path):
        # with contect opens file/connection and closes it if function(with) scope is left
        with open(gcode_path, "r") as file, serial.Serial(GRBL_port_path, 115200) as ser:
            send_wake_up(ser)
            for line in file:
                # cleaning up gcode from file
                cleaned_line = remove_eol_chars(remove_comment(line))
                if cleaned_line:  # checks if string is empty
                    print("Sending gcode:" + str(cleaned_line))
                    # converts string to byte encoded string and append newline
                    command = str.encode(line + '\n')
                    ser.write(command)  # Send g-code

                    wait_for_movement_completion(ser,cleaned_line)

                    grbl_out = ser.readline()  # Wait for response with carriage return
                    print(" : " , grbl_out.strip().decode('utf-8'))
        ser.close()
        file.close()

        
    #Закрытие диалоговых окон.
    def close_dlg(e):
        dlg_modal.open = False
        RefHomeXYZ.open = False
        page.update()


    
    #Запуск Gcode и записи данных с тензодатчика
    def start(e):    
        global Cycle_break 
        Cycle_break=False  
        print('Нажата кнопка СТАРТ')
        close_dlg(e)
        Max_Value_Txt.visible = False
        Max_Value_Txt.value = ''
        Min_Value_Txt.visible = False
        Min_Value_Txt.value = ''
        timetostart.value = ''
        Pstartbtn.disabled = True
        PChL.disabled = True
        RefXYZbtn.disabled = True
        Pstopbtn.disabled = False
        page.update()
        #multiprocessing.set_start_method('spawn',force=True)
        #p1 = multiprocessing.Process(target=stream_grbl_gcode, args = (GRBL_port_path,gcode_path),name='Stream_GCode',daemon=True)
        #p2 = multiprocessing.Process(target=PollingComPortTD,name='PollingTD',daemon=True)
        p1 = Thread(target=stream_grbl_gcode, args = (GRBL_port_path,gcode_path),name='Stream_GCode',daemon=True)
        p2 = Thread(target=PollingComPortTD,name='PollingTD',daemon=True)
        try:
            p2.start()
            print('запущена функция опроса ТД')
            CycleTxt.visible=True
            page.update()
            p1.start() 
            print("Запущена функция передачи Gcode")
            
            '''if Cycle_break:
                p1.join()  
                p2.join() 
                '''
            p1.join()
            print('Функция передачи завершена')
            p2.join() 
            print('Функция опроса завершена')
            CycleTxt.visible=False
            tenz.visible = False
            Pstopbtn.disabled = True
            page.update()
            time.sleep(5)
             
        except:
            logging.error("Start_thread delaet mozg",exc_info=True)
            print('Ошибка цикла старт')
            CycleTxt.visible=False
            page.update()
            
        time.sleep(3)
        Pstartbtn.disabled = False
        PChL.disabled = False
        RefXYZbtn.disabled = False
        page.update()
    def Startbtn(e):
        page.dialog = dlg_modal
        dlg_modal.open = True
        page.update()

    
    def Stopbtn(e):
        global Cycle_break 
        if Cycle_break == False:
            Cycle_break = True
        else:
            Cycle_break=False
            
        print( Cycle_break)
        return(Cycle_break)


 
    def changelist(e):
        try:
            print('Запущен цикл смены листа')
            stream_gcode(GRBL_port_path,Change_List_Gcode_path)
            print('Цикл смены листа завершен')
        except:
            logging.error("list change imposible",exc_info=True)
            print('Ошибка - смена листа')
            pass
        page.update()
    


    def home(e):
        close_dlg(e)
        page.update()
        try:
            stream_gcode(GRBL_port_path,homing_gcode_path)
        except:
            logging.error("Dont have connection or pizda",exc_info=True)
            pass
        Pstartbtn.visible=True
        Pstopbtn.visible=True
        PChL.visible=True
        page.update()
    def ask_home(e):
        page.dialog = RefHomeXYZ
        RefHomeXYZ.open = True
        page.update()

    Pstartbtn = ft.ElevatedButton(text="Старт", on_click=Startbtn,visible=False)
    Pstopbtn = ft.ElevatedButton(text="Стоп", on_click=Stopbtn,visible=False,disabled=True)
    PChL = ft.ElevatedButton(text="Замена листа", on_click=changelist,visible=False)
    RefXYZbtn = ft.ElevatedButton(text="Выполнить поиск ИТ для XYZ", on_click=ask_home)

    page.add(
        ft.Row(
            [   
                Pstartbtn,
                Pstopbtn,
                PChL,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=3,
            adaptive=True,
        ),
        ft.Row(
            [
                ft.Text(' \n \n ')
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=6,
            adaptive=True,
        ),
        ft.Row(
            [
                RefXYZbtn,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=3,
            adaptive=True,
        ),
        ft.Row(
            [
                ft.Text(' \n \n \n')
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=6,
            adaptive=True,
        ),
        ft.Row(
            [   
                CycleTxt,
                FileSaveTxt,
                tenz,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            scale=2,
        ),
        ft.Row(
            [
                Max_Value_Txt,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=2,
            adaptive=True,
        ),
        ft.Row(
            [
                Min_Value_Txt,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=2,
            adaptive=True,
        ),       
        ft.Row(
            [
                timetostart,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            #spacing= 120,
            scale=2,
            adaptive=True,
        )
    )
    
    #Диалоговое окно на старт программы
    dlg_modal = ft.AlertDialog(
        modal=True,
        title=ft.Text("Пожалуйста подтвердите!"),
        content=ft.Text("Установлен новый лист?"),
        actions=[
            ft.TextButton("Да", on_click=start),
            ft.TextButton("Нет", on_click=close_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        on_dismiss=lambda e: print("Лист установлен"),
    )
    #Диалоговое окно при запуске программы при нажатии на поиск ИТ
    RefHomeXYZ = ft.AlertDialog(
        modal=False,
        title=ft.Text("Пожалуйста будьте осторожны! Уберите руки от подвижных частей!"),
        content=ft.Text("Начать поиск исходной точки?"),
        actions=[
            ft.TextButton("Да", on_click=home),
            ft.TextButton("Нет", on_click=close_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        on_dismiss=lambda e: print("Поиск ИТ"),

    )
    
ft.app(target=main)#, view=ft.WEB_BROWSER, port= 8510)
# блок логирования
import logging
logging.basicConfig(level=logging.INFO, filename="//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/py_log.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")
# https://habr.com/ru/companies/wunderfund/articles/683880/   - ссылка на статью логирования
# filemode="a" дозапись "w" - перезапись
logging.info("Запуск скрипта starter_temp.py")

# блок обратки данных
import pandas as pd
import os
from datetime import datetime, date, timedelta
pd.options.display.max_colwidth = 100 # увеличить максимальную ширину столбца
pd.set_option('display.max_columns', None) # макс кол-во отображ столбц

# блок импортов для обновления сводных
import pythoncom
pythoncom.CoInitializeEx(0)
import win32com.client
import time

# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

def read_email_adress(mail = fr'\\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\temp_\Список_адресатов.xlsx'):
    """Функция считывания адресатов для рассылки

    Args:
        mail (_type_, optional): _description_. Defaults to fr'\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\temp_\Список_адресатов.xlsx'.

    Returns:
        _type_: возфращает строку со списком email
    """
    logging.info(f"запуск функции {read_email_adress.__name__}")
    
    try:
        em_list = pd.read_excel(mail)
        return list(em_list['email'])
    except:
        logging.error(f"{read_email_adress.__name__} - ОШИБКА", exc_info=True)

def testing_links(links:list):
    """Проверка ссылкок на файлы  
    
    Результатом является вывод текстового сообщения с результатом проверки ссылки 

    Args:
        links (_type_): list _description_ - подается список ссылок
    """
    logging.info(f"запуск функции {testing_links.__name__}")

    for i in links:
        if os.path.exists(f"{i}"):
            #print(f'OK - ', i)
            logging.info(f"{testing_links.__name__} ссылка рабочая {i}")
        else:
            print(f"ОШИБКА - ", i)
            logging.error(f"{testing_links.__name__} ссылка не рабочая {i}", exc_info=True)
            
def reg_test(rg, podr):
    """функция находит YAR и проверяет есть ли там RYB

    Args:
        rg (_type_): столбец регион
        podr (_type_): столбец подразделение

    Returns:
        _type_: _description_
    """
    
    if rg == 'YAR':
        if 'яр' in str(podr).lower():
            return 'YAR'
        elif 'рыб' in str(podr).lower():
            return 'RYB'
    else:
        return rg
            
# функция проверки шапки по первой строке если шапка не в первой строке и такм нен упоминания слова vin значит ищем по всем столцам и находим первое вхождение возвращаем строку и обрезаем df
def header_df(df):
    """Преобразование шапки df  
    
    если названия заголовков в таблице не в первой строке, скрипт ищет шапку по ключевому значению vin, 
    удаляет лишние строки и  переопределяет строку в заголовок

    Args:
        df (_type_): df - принимает

    Returns:
        _type_: df - возвращает 
    """
    
    logging.info(f"{header_df.__name__} - ЗАПУСК")
    
    try:
        
        count_col = 0
        for i in df.columns:
            if str(i).lower() == 'vin':
                count_col +=1

            counter_vin = df[i].apply(lambda x: str(x).lower()).str.contains('^vin').sum() # ^ - в регулярке используется для поиска когда слово начинается с 
            name_column = i
            row_number = None
            if counter_vin >0:
                row_number = df[df[name_column].apply(lambda x: str(x).lower())=='vin'].index[0]
                break
                
        if count_col != 0:
            return df # если шапка в первой строке, ничего не изменяем
        else:
            new_header = df.iloc[row_number] # берем первую строку как заголовок
            df = df[row_number+1:]  # отбрасываем исходный заголовок
            df.rename(columns=new_header, inplace=True) # переименовываем столбцы
            #df = df[(df['VIN'].notna()) & (df['VIN'].apply(lambda x: not str(x).isdigit())) & (df['VIN'].apply(lambda x: len(str(x))>1)) ] # vin не пусто и vin не число и длина больше 1 символа
            #df = df[(df['VIN'].apply(lambda x: not str(x).isdigit()))] 
            return df
    except:
        logging.error(f"{header_df.__name__} - ОШИБКА", exc_info=True)
    

# считываем актуальный пароль 
def my_pass():
    """функция считывания пароля

    Returns:
        _type_: _description_
    """
    logging.info(f"{my_pass.__name__} - ЗАПУСК")
    
    try:
        with open(f'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/password_email.txt', 'r') as actual_pass:
            return actual_pass.read()
        
    except:
        logging.error(f"{my_pass.__name__} - ОШИБКА", exc_info=True)
    
# письмо если нет ошибок
def send_mail(send_to:list):
    """рассылка почты

    Args:
        send_to (list): _description_
    """
    logging.info(f"{send_mail.__name__} - ЗАПУСК")
    
    try:
        send_from = 'skrutko@sim-auto.ru'                                                                
        subject = f"Темпы на {(datetime.now()-timedelta(1)).strftime('%d-%m-%Y')}"                                                                  
        text = f"Здравствуйте\nВо вложении темпы на {(datetime.now()- timedelta(1)).strftime('%d-%m-%Y')}"                                                                      
        files = "//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/ТЕМП_СВОД1.xlsx"  
        server = "server-vm36.SIM.LOCAL"
        port = 587
        username='skrutko'
        password=my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="temp.xlsx"') # имя файла должно быть на латинице иначе придет в кодировке bin
        msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        logging.info(f"{send_mail.__name__} - ВЫПОЛНЕНО")
        logging.info(f"Адреса рассылки {send_to}")
    except:
        logging.error(f"{send_mail.__name__} - ОШИБКА", exc_info=True)
    

# письмо если есть ошибки
def send_mail_danger(send_to:list):
    """расслыка почты если ошибка

    Args:
        send_to (_type_): _description_
    """
    logging.info(f"{send_mail_danger.__name__} - ЗАПУСК")
    
    try:                                                                                       
        send_from = 'skrutko@sim-auto.ru'                                                                
        subject =  f"проверьте исходники {'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_'}"                                                                  
        text = f"проверьте исходники {'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_'}"                                                                      
        files = '//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/py_log.log'  
        server = "server-vm36.SIM.LOCAL"
        port = 587
        username='skrutko'
        password=my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="log.txt"') # имя файла должно быть на латинице иначе придет в кодировке bin
        msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        logging.info(f"{send_mail_danger.__name__} - ВЫПОЛНЕНО")
        logging.info(f"Адреса рассылки {send_to}")
    except:
        logging.error(f"{send_mail_danger.__name__} - ОШИБКА", exc_info=True)
    

def detected_danger(filename_log = "//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/py_log.log"):
    """обнаружение ошибок в логах   
    ищет 'warning'

    Returns:
        _type_: bool
    """
    logging.info(f"{detected_danger.__name__} - ЗАПУСК")
    
    try:
        with open(filename_log, '+r') as file:
            return 'warning' in file.read().lower()
    except:
        logging.error(f"{detected_danger.__name__} - ОШИБКА", exc_info=True)
        
        
def sending_mail(lst_email, lst_email_error):
    """рассылка почты - если нет ошибок вызываем send_mail(),   
    если есть ошибки send_mail_error()   
    """
    logging.info(f"{sending_mail.__name__} - ЗАПУСК")
    
    try:
        if detected_danger()==False:
            send_mail(lst_email)
        else:
            send_mail_danger(lst_email_error)
            
        logging.info(f"{sending_mail.__name__} - ВЫПОЛНЕНО")
    except:
        logging.error(f"{sending_mail.__name__} - ОШИБКА", exc_info=True)
        
        


lst_df = {'df_chr_msc' : {'link': '//sim.local/data/Varsh/DPA/CHERY-ЮЗ/Payment/NP-Chery.xlsm',
                          'reg': 'MSK',
                          'brand' : 'CHERY',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'кре_нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'кре_нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_kia_msc' : {'link': '//sim.local/data/Varsh/DPA/КИА-ЮЗ/NP-ЮЗ-Киа.xlsm',
                          'reg': 'MSK',
                          'brand' : 'KIA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_kia_msc_korp' : {'link': '//sim.local/data/Varsh/DPA/КИА-ЮЗ/NP-ЮЗ-Киа-Корп.xlsm',
                          'reg': 'MSK',
                          'brand' : 'KIA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          
          'df_kia_msc_arhiv' : {'link': '//sim.local/data/Varsh/OFFICE/CAGROUP/АВТО/АВТО продажи/Архив/КИА/NP_kia.xlsx',
                          'reg': 'MSK',
                          'brand' : 'KIA',
                          'lst_sheet_name' : 'NP',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_suz_msc_vved' : {'link': '//sim.local/data/Varsh/DPA/Юго-Запад/Payment/NP-ЮЗ.xlsm',
                          'reg': 'MSK',
                          'brand' : 'SUZUKI',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_suz_msc_madi' : {'link': '//Sim.local/data/Madi/FIN/Payment/NP-M.xlsm',
                          'reg': 'MSK',
                          'brand' : 'SUZUKI',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_ovp_msc_vved' : {'link': '//sim.local/data/Varsh/DPA/Юго-Запад/Payment/ОВП ЮЗ.xlsx',
                          'reg': 'MSK',
                          'brand' : 'OVP',
                          'lst_sheet_name' : 'Склад',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа__контракта', 'дата_выдачи', 'дата_полной_оплаты', 'цена_продажи,_руб.', 'примечание', 'источник_а_м'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа__контракта' : 'дата_заказа', 
                                          'дата_выдачи' : 'дата_выдачи', 
                                          'дата_полной_оплаты' : 'дата_оплаты', 
                                          'цена_продажи,_руб.' : 'сум_спр_сч', 
                                          'примечание': 'кре_нал',
                                          'источник_а_м': 'подразделение'}},
          
          'df_ovp_msc_madi' : {'link': '//sim.local/data/Madi/FIN/Payment/ОВП МАДИ.xlsx',
                          'reg': 'MSK',
                          'brand' : 'OVP',
                          'lst_sheet_name' : 'Склад',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа__контракта', 'дата_выдачи', 'дата_полной_оплаты', 'цена_продажи,_руб.', 'примечание', 'источник_а_м'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа__контракта' : 'дата_заказа', 
                                          'дата_выдачи' : 'дата_выдачи', 
                                          'дата_полной_оплаты' : 'дата_оплаты', 
                                          'цена_продажи,_руб.' : 'сум_спр_сч', 
                                          'примечание': 'кре_нал',
                                          'источник_а_м': 'подразделение'}},
          
          'df_mzd_msc_vved' : {'link': '//sim.local/data/Varsh/DPA/MAZDA-ЮЗ/Payment/NP-МаздаМосква.xlsm',
                          'reg': 'MSK',
                          'brand' : 'MAZDA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_jtr_msc_vved' : {'link': '//Sim.local/data/Varsh/DPA/MAZDA-ЮЗ/JETOUR ЮЗ/Диспонент/NP-Jetour.xlsm',
                          'reg': 'MSK',
                          'brand' : 'JETOUR',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
          
          'df_bai_msc_varsh' : {'link': '//sim.local/data/Varsh/SalonHyundai/Секретари-оформители/внутренние отчеты/Склад BAIC Москва.xlsx',
                          'reg': 'MSK',
                          'brand' : 'BAIC',
                          'lst_sheet_name' : 'СКЛАД',
                          'white_list_columns' : ['модель', 'vin', 'дата_заключения_договора', 'дата_выдачи_клиенту', 'дата_полной_оплаты_ам_клиентом', 'сумма_оплаченная_клиентом', 'нал_кредит', 'регион'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заключения_договора' : 'дата_заказа', 
                                          'дата_выдачи_клиенту' : 'дата_выдачи', 
                                          'дата_полной_оплаты_ам_клиентом' : 'дата_оплаты', 
                                          'сумма_оплаченная_клиентом' : 'сум_спр_сч', 
                                          'нал_кредит': 'кре_нал',
                                          'регион': 'подразделение'}},
          
           'df_hyu_msc_varsh' : {'link': '//sim.local/data/Varsh/SalonHyundai/Секретари-оформители/внутренние отчеты/Склад HYUNDAI Москва.xlsx',
                          'reg': 'MSK',
                          'brand' : 'HYUNDAI',
                          'lst_sheet_name' : 'СКЛАД',
                          'white_list_columns' : ['модель', 'vin', 'дата_заключения_договора', 'дата_выдачи_клиенту', 'дата_полной_оплаты_ам_клиентом', 'сумма_оплаченная_клиентом', 'нал_кредит', 'регион'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заключения_договора' : 'дата_заказа', 
                                          'дата_выдачи_клиенту' : 'дата_выдачи', 
                                          'дата_полной_оплаты_ам_клиентом' : 'дата_оплаты', 
                                          'сумма_оплаченная_клиентом' : 'сум_спр_сч', 
                                          'нал_кредит': 'кре_нал',
                                          'регион': 'подразделение'}},
           
           'df_hyu_uka_msc_varsh' : {'link': '//sim.local/data/Varsh/SalonHyundai/Секретари-оформители/online продажи/UKA/Склад UKA - Аукцион.xlsx',
                          'reg': 'MSK',
                          'brand' : 'UKA',
                          'lst_sheet_name' : 'Склад UKA-Аукцион',
                          'white_list_columns' : ['модель', 'vin', 'дата_заключения_договора', 'дата_выдачи_клиенту', 'дата_полной_оплаты_ам_клиентом', 'сумма_оплаченная_клиентом', 'нал_кредит', 'пустой1'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заключения_договора' : 'дата_заказа', 
                                          'дата_выдачи_клиенту' : 'дата_выдачи', 
                                          'дата_полной_оплаты_ам_клиентом' : 'дата_оплаты', 
                                          'сумма_оплаченная_клиентом' : 'сум_спр_сч', 
                                          'нал_кредит': 'кре_нал',
                                          'пустой1': 'подразделение'}},
           
           'df_suz_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Mazda!/Сетевая информация/Отчеты NP/NP-SuzYarNEW 2020.xlsm',
                          'reg': 'YAR',
                          'brand' : 'SUZUKI',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
           
           'df_mzd_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Mazda!/Сетевая информация/Отчеты NP/NP-Mazda 2020.xlsm',
                          'reg': 'YAR',
                          'brand' : 'MAZDA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
           
            'df_ren_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Отдел логистики/NP. PLAN-REAL/New_Pay$ (RUB) Ярославль.xlsm',
                          'reg': 'YAR',
                          'brand' : 'RENAULT',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
            
             'df_hyu_yar' : {'link': '//sim.local/data/Yar/YAR_Hyundai/Отдел продаж/ОТЧЕТЫ/New_Pay$ (RUB)Hyundai.xlsm',
                          'reg': 'YAR',
                          'brand' : 'HYUNDAI',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
             
             'df_vw_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Volkswagen/!Отдел продаж/ЛОГИСТ/NP/NP-Volkswagen.xlsm',
                          'reg': 'YAR',
                          'brand' : 'VOLKSWAGEN',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
             
              'df_chv_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Chevrolet/Логистика/New_Pay$ (RUB) Chevrolet.xlsm',
                          'reg': 'YAR',
                          'brand' : 'CHEVROLET',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
              
              'df_ovp_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/TRADE-IN/Отчеты для Москвы/Новый ОВП/ОВП-Яр..xlsx',
                          'reg': 'YAR',
                          'brand' : 'OVP',
                          'lst_sheet_name' : 'Склад',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа__контракта', 'дата_выдачи', 'дата_полной_оплаты', 'цена_продажи,_руб.', 'примечание', 'регион'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа__контракта' : 'дата_заказа', 
                                          'дата_выдачи' : 'дата_выдачи', 
                                          'дата_полной_оплаты' : 'дата_оплаты', 
                                          'цена_продажи,_руб.' : 'сум_спр_сч', 
                                          'примечание': 'кре_нал',
                                          'регион': 'подразделение'}},
              
               'df_jac_yar' : {'link': '//sim.local/data/Yar/YAR_JAC/Отдел продаж/Логистика/New_Pay$ (RUB) Jac.xlsm',
                          'reg': 'YAR',
                          'brand' : 'JAC',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
               
                'df_mch_yar' : {'link': '//sim.local/data/Yar/YAR_Moskvich/Отдел продаж/ЛОГИСТИКА/New_Pay$ (RUB) Москвич.xlsm',
                          'reg': 'YAR',
                          'brand' : 'MOSKVICH',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                
                'df_faw_yar' : {'link': '//sim.local/data/Yar/YAR_FAW/Отдел продаж/ЛОГИСТ FAW+JETTA/NP/NP-FAW.xlsm',
                          'reg': 'YAR',
                          'brand' : 'FAW',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                
                'df_jet_yar' : {'link': '//sim.local/data/Yar/YAR_FAW/Отдел продаж/ЛОГИСТ FAW+JETTA/NP/NP-JETTA.xlsm',
                          'reg': 'YAR',
                          'brand' : 'JETTA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                
                 'df_chr_yar' : {'link': '//sim.local/data/Yar/Старая папка Общая/Mazda!/Сетевая информация/Отчеты NP/NP-Chery 2023.xlsm',
                          'reg': 'YAR',
                          'brand' : 'CHERY',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд','бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                 
                  'df_bai_yar' : {'link': '//sim.local/data/Yar/YAR_BAIC/Отдел продаж/ЛОГИСТИКА/New_Pay$ (RUB)Baic.xlsm',
                          'reg': 'YAR',
                          'brand' : 'BAIC',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                  
                   'df_ren_ryb' : {'link': '//sim.local/data/Yar/Старая папка Общая/Отдел логистики/NP. PLAN-REAL/New_Pay$ (RUB) Рыбинск.xlsm',
                          'reg': 'RYB',
                          'brand' : 'RENAULT',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                   
                    'df_prch_ryb' : {'link': '//sim.local/data/Yar/YAR_Rybinsk/Отдел продаж/!ОТЧЕТЫ/New_Pay$ (RUB)Rybinsk.xlsm',
                          'reg': 'RYB',
                          'brand' : 'PRCH',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                    
                    'df_mzd_sar' : {'link': '//Sim.local/data/Varsh/DPA/САРАТОВ/NP-MazdaСаратов.xlsm',
                          'reg': 'SAR',
                          'brand' : 'MAZDA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                    
                    
                     'df_omd_sar' : {'link': '//Sim.local/data/Varsh/DPA/САРАТОВ/NP-OmodaСаратов.xlsm',
                          'reg': 'SAR',
                          'brand' : 'OMODA',
                          'lst_sheet_name' : 'Авто',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа', 'дата_ф._выдачи', 'ф.дата_полн._опл.', 'сумма_спр.сч._(руб)', 'б_н___нал', 'подразделение'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа' : 'дата_заказа', 
                                          'дата_ф._выдачи' : 'дата_выдачи', 
                                          'ф.дата_полн._опл.' : 'дата_оплаты', 
                                          'сумма_спр.сч._(руб)' : 'сум_спр_сч', 
                                          'б_н___нал': 'кре_нал',
                                          'подразделение': 'подразделение'}},
                     
                     
                     'df_ovp_sar' : {'link': '//Sim.local/data/Varsh/DPA/САРАТОВ/ОВП-Саратов.xlsx',
                          'reg': 'SAR',
                          'brand' : 'OVP',
                          'lst_sheet_name' : 'Склад',
                          'white_list_columns' : ['модель', 'vin', 'дата_заказа__контракта', 'дата_выдачи', 'дата_полной_оплаты', 'цена_продажи,_руб.', 'примечание', 'площадка'],
                          'order_columns': ['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения', 'подразделение'],
                          'rename_col' : {'модель' : 'модель', 
                                          'vin' : 'vin', 
                                          'дата_заказа__контракта' : 'дата_заказа', 
                                          'дата_выдачи' : 'дата_выдачи', 
                                          'дата_полной_оплаты' : 'дата_оплаты', 
                                          'цена_продажи,_руб.' : 'сум_спр_сч', 
                                          'примечание': 'кре_нал',
                                          'площадка': 'подразделение'}},
                     
          
          
          }

# \\sim.local\data\Varsh\OFFICE\CAGROUP\АВТО\АВТО продажи\Архив\КИА   - здесь кияшные NP_kia.xlsx в этом файле все архивы всех годов и текущего

# тестирование ссылок
testing_links([lst_df[i]['link'] for i in lst_df.keys()])


class Manufactory_df:
    
    def __init__(self, 
                 name_df, 
                 lst_diction_df, 
                 flg = False):
        """ ИНИЦИАЛИЗАЦИЯ

        Аргументы на вход:
            name_df (_type_): имя датафрейма
            lst_diction_df (_type_): словарь с фходными данными
            flg (str) - если флаг False То функции предобратки ниже не выполняются авто по умолчанию
            
        Автоаргументы:
            white_list (str) - ключ словаря с белым списком столбцов по умолчанию white_list_columns
            rename_col - имя ключа словаря со словарем для переименования
            reg - параметр по умолчанию (регион) ищет раздел  по словарю 
            brand - параметр по умолчанию (бренд) ищет раздел  по словарю 
            order_columns - порядок столбцов в конечном df
            filename - собирается общая ссылка/путь на файл который будем считывать
            lst_sheet_name - имя листа с которого будем считывать даные 
            df - считанный df
            mtime - последнее дата время изменения файла
            mtime_readable - изменения БД с разбивкой г/м/д/ч/м/с
            
        """
        
        self.name_df = name_df    # иницализируем имя, которое будет подаваться из словаря
        self.lst_diction_df = lst_diction_df
        self.flg = flg
       
        self.white_list = 'white_list_columns'
        self.rename_col = 'rename_col'
        self.reg = 'reg' 
        self.brand = 'brand'
        self.order_columns = 'order_columns'
        self.kre_nal = 'кре_нал'
        self.filename = lst_diction_df[self.name_df]['link'] # сылка к файлу / путь к каталогу
        self.lst_sheet_name = lst_diction_df[self.name_df]['lst_sheet_name']  # тащим имя листа с которого будем считывать даные 
        self.df = pd.read_excel(self.filename, dtype='str', sheet_name=self.lst_sheet_name)  # читам данные / передали имя листа
        self.mtime = os.path.getmtime(self.filename)
        self.mtime_readable = datetime.fromtimestamp(self.mtime)
        logging.info(f"создание объекта класса {__class__.__name__} имя {self.name_df}")
        self.fnc_auto()
        
        
    def header_df_act(self):
        """применяем глобальную функцию header_df для поиска шапки в df 
        """
        print(f"{self.header_df_act.__name__} - ЗАПУСК")
        logging.info(f"{self.header_df_act.__name__} - ЗАПУСК")

        try:
            self.df = header_df(self.df)
            
            print(f"{self.header_df_act.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.header_df_act.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.header_df_act.__name__} - ОШИБКА")
            logging.error(f"{self.header_df_act.__name__} - ОШИБКА", exc_info=True)

        
    def registr_df(self):
        """приводит колонки df в нижний регистр
        """
        print(f"{self.registr_df.__name__} - ЗАПУСК")
        logging.info(f"{self.registr_df.__name__} - ЗАПУСК")
        
        try:
            self.df = self.df.rename(columns={str(i):str(i).lower().
                                            replace(' ','_').replace('/','_') for i in self.df.columns})
            
            print(f"{self.registr_df.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.registr_df.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.registr_df.__name__} - ОШИБКА")
            logging.error(f"{self.registr_df.__name__} - ОШИБКА", exc_info=True)
        
        
    def cleaner_columns(self):
        """очищает df от лишних столбцов - ориентируясь на белый список столбцов из словаря
        """
        print(f"{self.cleaner_columns.__name__} - ЗАПУСК")
        logging.info(f"{self.cleaner_columns.__name__} - ЗАПУСК")
        
        try:
            self.df = self.df[[i for i in self.df.columns if i in self.lst_diction_df[self.name_df][self.white_list]]]
            
            print(f"{self.cleaner_columns.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.cleaner_columns.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.cleaner_columns.__name__} - ОШИБКА")
            logging.error(f"{self.cleaner_columns.__name__} - ОШИБКА", exc_info=True)
        
    
    def rename_columns(self):
        """переименовываем название колонк под единый шаблон
        """
        print(f"{self.rename_columns.__name__} - ЗАПУСК")
        logging.info(f"{self.rename_columns.__name__} - ЗАПУСК")
        
        try:
            for i in self.df.columns:
                try:
                    self.df = self.df.rename(columns={i:self.lst_diction_df[self.name_df][self.rename_col][i]})
                    #print(f"Переименование столбца {i}")
                except (KeyError) as err:
                    print(f'Ошибка KeyError в словаре не удалось найти ключи по столбцам {err}')
                    logging.warning(f'Ошибка KeyError в словаре не удалось найти ключи по столбцам {err}')
                    
            print(f"{self.rename_columns.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.rename_columns.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.rename_columns.__name__} - ОШИБКА")
            logging.error(f"{self.rename_columns.__name__} - ОШИБКА", exc_info=True)
                
    
    def add_columns(self):
        """добавление столбцов в df
        """
        
        print(f"{self.add_columns.__name__} - ЗАПУСК")
        logging.info(f"{self.add_columns.__name__} - ЗАПУСК")
        
        try:
            self.df['регион'] = self.lst_diction_df[self.name_df][self.reg] 
            self.df['бренд'] = self.lst_diction_df[self.name_df][self.brand] 
            self.df['дата_изменения'] = self.mtime_readable
            
            print(f"{self.add_columns.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.add_columns.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.add_columns.__name__} - ОШИБКА")
            logging.error(f"{self.add_columns.__name__} - ОШИБКА", exc_info=True)
            
            
            
    def isalph_kre(self):
        ''' проверяем есть ли в столбце кре_нал цифры, каких быть не должно  
        обрезаем датафрейм по условияю : в столбце кре_нал только буквы
        '''
        
        print(f"{self.isalph_kre.__name__} - ЗАПУСК")
        logging.info(f"{self.isalph_kre.__name__} - ЗАПУСК")
        try:
            self.df = self.df[self.df[self.kre_nal].apply(lambda x: str(x).isalpha() or len(str(x))>2 )]
            print(f"{self.isalph_kre.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.isalph_kre.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.isalph_kre.__name__} - ОШИБКА")
            logging.error(f"{self.isalph_kre.__name__} - ОШИБКА", exc_info=True)
        

        
    def date_ISO(self):
        """ищет в df столбцы с именем "дата" и правит формат даты по стандарту format='ISO8601'
        """
        print(f"{self.date_ISO.__name__} - ЗАПУСК")
        logging.info(f"{self.date_ISO.__name__} - ЗАПУСК")
        
        try:
            for i in self.df.columns:
                if 'дата' in i:
                    self.df[i] = pd.to_datetime(self.df[i], format='mixed') # mixed format='ISO8601'
                    
            print(f"{self.date_ISO.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.date_ISO.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.date_ISO.__name__} - ОШИБКА")
            logging.error(f"{self.date_ISO.__name__} - ОШИБКА", exc_info=True)
                
                
    def del_NAN(self, sub:list = ['дата_заказа']):            # 'vin', 'модель'  'дата_заказа'
        """убираем NaN по умолчанию столбцы 'vin', 'модель'
        
        Args:   
            sub (list, optional): столбцы по умолчанию ['vin', 'модель'].
        """
        print(f"{self.del_NAN.__name__} - ЗАПУСК")
        logging.info(f"{self.del_NAN.__name__} - ЗАПУСК")
        try:
            self.df = self.df.dropna(subset=sub) # столбцы по которым убираем NaN
            
            print(f"{self.del_NAN.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.del_NAN.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.del_NAN.__name__} - ОШИБКА")
            logging.error(f"{self.del_NAN.__name__} - ОШИБКА", exc_info=True)
            
            
    def bd_name_columns(self):
        """добавляет столбец с именем датафрейма, так будет легче искать ошибки
        """
        print(f"{self.bd_name_columns.__name__} - ЗАПУСК")
        logging.info(f"{self.bd_name_columns.__name__} - ЗАПУСК")
        try:
            self.df['бд'] = self.name_df
            print(f"{self.bd_name_columns.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.bd_name_columns.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.bd_name_columns.__name__} - ОШИБКА")
            logging.error(f"{self.bd_name_columns.__name__} - ОШИБКА", exc_info=True)
            
            
    def order_columns_fn(self):
        """упорядочивает последовательность столбцов
        """
        print(f"{self.order_columns_fn.__name__} - ЗАПУСК")
        logging.info(f"{self.order_columns_fn.__name__} - ЗАПУСК")
        try:
            
            self.df = self.df[self.lst_diction_df[self.name_df][self.order_columns]]
            print(f"{self.order_columns_fn.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.order_columns_fn.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.order_columns_fn.__name__} - ОШИБКА")
            logging.error(f"{self.order_columns_fn.__name__} - ОШИБКА", exc_info=True)
            
            
    def kredit_nal(self):
        """функция опеделяет кредит в единый вид кре
        """
        
        print(f"{self.kredit_nal.__name__} - ЗАПУСК")
        logging.info(f"{self.kredit_nal.__name__} - ЗАПУСК")
        try:
            self.df['кре_нал'] = self.df['кре_нал'].apply(lambda x: 'кре' if 'кре' in str(x).lower() or 'лиз' in str(x).lower() or 'лизинг' in str(x).lower() or 'лиизинг' in str(x).lower() or 'банк' in str(x).lower() or 'finance' in str(x).lower() or 'финанс' in str(x).lower() else str(x).lower())
            self.df['кре_нал'] = self.df['кре_нал'].apply(lambda x: 'нал' if 'бн' in str(x).lower() or 'б/н' in str(x).lower() or 'безнал' in str(x).lower() or 'нал' in str(x).lower() or 'nan' in str(x).lower() or 'корп' in str(x).lower() else str(x).lower())
            
            print(f"{self.kredit_nal.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.kredit_nal.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.kredit_nal.__name__} - ОШИБКА")
            logging.error(f"{self.kredit_nal.__name__} - ОШИБКА", exc_info=True)
            
            
    def OVP_individ_kredit(self):
        """работает только для ОВП    
        находит столбец примечание и по нему определяет кре / нал   
        по новым ам этот столбец определен 
        """
        print(f"{self.OVP_individ_kredit.__name__} - ЗАПУСК")
        logging.info(f"{self.OVP_individ_kredit.__name__} - ЗАПУСК")
        
        try:
            if self.lst_diction_df[self.name_df][self.brand] == 'OVP':
                self.df['кре_нал'] = self.df['кре_нал'].apply(lambda x: 'кре' if 'кре' in " ".join(list(str(x).lower().split())) else 'нал')
            print(f"{self.OVP_individ_kredit.__name__} - ВЫПОЛНЕНО")
            logging.info(f"{self.OVP_individ_kredit.__name__} - ВЫПОЛНЕНО")

        except:
            print(f"{self.OVP_individ_kredit.__name__} - ОШИБКА")
            logging.error(f"{self.OVP_individ_kredit.__name__} - ОШИБКА", exc_info=True)
            
    def SAR_OMD_split(self):
        ''' применяется только к 'OMODA' 'SAR'  
        по столбцу модель если находит JAECOO - изменяет значение в столбце бренд  
        тем самым разбивая авто на JAECOO и OMODA  
        '''
        print(f"{self.SAR_OMD_split.__name__} - ЗАПУСК")
        logging.info(f"{self.SAR_OMD_split.__name__} - ЗАПУСК")
    
        try:
            if self.lst_diction_df[self.name_df][self.brand] == 'OMODA' and self.lst_diction_df[self.name_df][self.reg] == 'SAR':
                self.df['бренд'] = self.df['модель'].apply(lambda x: 'JAECOO' if 'jaecoo' in str(x).lower() else 'OMODA')
                print(f"{self.SAR_OMD_split.__name__} - ВЫПОЛНЕНО")
                logging.info(f"{self.SAR_OMD_split.__name__} - ВЫПОЛНЕНО")
        except:
            print(f"{self.SAR_OMD_split.__name__} - ОШИБКА")
            logging.error(f"{self.SAR_OMD_split.__name__} - ОШИБКА", exc_info=True)
        
        
    def fnc_auto(self):
        """запускаем все функции предобработки
        Если нужно проверить или донастроить каждую комментируем, нужно настоить срабатывание на переменную 
        flg - при создание экхемпляра класса, если True - выполняются функции по умолчанию, если False - нет
        """
        print(f"{self.fnc_auto.__name__} - ЗАПУСК")
        logging.info(f"{self.fnc_auto.__name__} - ЗАПУСК")
        
        try:
            
            if self.flg == True:
                print('Автоприменение функций предобработки - ВКЛ')
                logging.info(f"{self.fnc_auto.__name__} - Автоприменение функций предобработки - ВКЛ")
                self.header_df_act()
                self.registr_df()
                self.cleaner_columns()
                self.rename_columns()
                self.add_columns()
                self.isalph_kre()
                self.date_ISO()
                self.del_NAN()
                self.bd_name_columns()
                self.order_columns_fn()
                self.kredit_nal()
                self.OVP_individ_kredit() # работает только для ОВП
                self.SAR_OMD_split() # работает только для SAR OMD
                
            else:
                print('Автоприменение функций предобработки - ОТКЛ')
                logging.info(f"{self.fnc_auto.__name__} - Автоприменение функций предобработки - ОТКЛ")
        except:
            logging.error(f"{self.fnc_auto.__name__} - ОШИБКА", exc_info=True)
            
        
# наполняем словарь базами данных создавая экземпляры класса
catalog_df = {} # словарь со всеми базами

for i in lst_df.keys():
    catalog_df[i] = Manufactory_df(i, lst_df, flg=True)
    
logging.info(f"словарь с датафреймами заполнен")


# сборка всех df в 1
logging.info(f"конкатинация всех датафремов")
frames = [catalog_df[i].df for i in catalog_df.keys()]
result = pd.concat(frames)
logging.info(f"конкатинация выполнена")


# применяем функцию для вычленения Рыбинска из Ярославля
logging.info(f"запуск функции reg_test")
try:
    result['регион'] = result.apply(lambda x: reg_test(x.регион, x.подразделение), axis=1)
    logging.info(f"{reg_test.__name__} - ВЫПОЛНЕНО")
except:
    logging.error(f"{reg_test.__name__} - ОШИБКА", exc_info=True)


# оставляем только нужные столбцы
logging.info(f"обрезаем столбцы")
result = result[['модель', 'vin', 'дата_заказа', 'дата_выдачи', 'дата_оплаты', 'сум_спр_сч', 'кре_нал', 'регион', 'бренд', 'бд', 'дата_изменения']]
logging.info(f"обрезали")


logging.info(f"сохранение датафремов")
try:
    result.to_excel('//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/temp.xlsx')
    logging.info(f"сохранены")
except:
    logging.error(f"ОШИБКА сохранения", exc_info=True)
        
        
logging.info(f"разделение копированием датафремов на выдачи и заказы ")
try:
    df_yudach = result.copy()
    df_zakaz = result.copy()
    logging.info(f"разделены")
except:
    logging.error(f"ОШИБКА копирования/разделения", exc_info=True)

          
logging.info(f"Формируем выдачи ")
try:
    df_yudach['выдача'] = df_yudach['дата_выдачи'].apply(lambda x: 1 if len(str(x))>4 else 0)
    df_yudach = df_yudach.rename(columns={'дата_выдачи':'дата'})
    del df_yudach['дата_заказа']
    del df_yudach['дата_оплаты']
    logging.info(f"сформированы")
    #df_yudach
except:
    logging.error(f"ОШИБКА формирования выдач", exc_info=True)
    
    
logging.info(f"Формируем заказы ")
try:
    df_zakaz['заказ'] = df_zakaz['дата_заказа'].apply(lambda x: 1 if len(str(x))>3 else 0)
    df_zakaz = df_zakaz.rename(columns={'дата_заказа':'дата'})
    del df_zakaz['дата_выдачи']
    del df_zakaz['дата_оплаты']
    logging.info(f"сформированы")
    #df_zakaz
except:
    logging.error(f"ОШИБКА формирования заказов", exc_info=True)
    

# конкатинируем выдачи и заказы
logging.info(f"Конкатинируем выдачи с заказами")
try:
    result1 = pd.concat([df_yudach, df_zakaz])
    logging.info(f"сконкатинированы")
    #result1
except:
    logging.error(f"ОШИБКА конкатинации", exc_info=True)
    
    
logging.info(f"Добавляем столбец - день")
try:
    result1['день'] = result1['дата'].apply(lambda x:  date.fromisoformat(str(x).split()[0]).day if len(str(x))>3 else x)
    logging.info(f"добавлен")
    #result1
except:
    logging.error(f"ОШИБКА добавления столбца", exc_info=True)
    

# исключаем сегодняшнюю дату

logging.info(f"Исключаем сегодняшний день")
try:
    result1 = result1[(result1['дата'] < datetime.today().date().isoformat()) | result1['дата'].isna()]
    logging.info(f"добавлен")
except:
    logging.error(f"ОШИБКА", exc_info=True)
    
    
# сохраняем

logging.info(f"сохраняем df //sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/temp_result.xlsx")
try:
    result1.to_excel('//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/temp_result.xlsx')
    logging.info(f"добавлен")
except:
    logging.error(f"ОШИБКА", exc_info=True)
    
    
# Обновляем сводные таблицы
logging.info(f"Обновляем сводные таблицы")
try:
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open("//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/ТЕМП_СВОД1.xlsx")
    wb.Application.AskToUpdateLinks = False   # разрешает автоматическое  обновление связей (файл - парметры - дополнительно - общие - убирает галку запрашивать об обновлениях связей)
    wb.Application.DisplayAlerts = True  # отображает панель обновления иногда из-за перекрестного открытия предлагает ручной выбор обновления True - показать панель
    wb.RefreshAll()
    #xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
    time.sleep(40) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
    wb.Application.AskToUpdateLinks = True   # запрещает автоматическое  обновление связей / то есть в настройках экселя (ставим галку обратно)
    wb.Save()
    wb.Close()
    xlapp.Quit()
    wb = None # обнуляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    xlapp = None # обнуляем сслыки переменных иначе процесс эксел ь не завершается и висит в дистпетчере
    del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    logging.info(f"сводные таблицы - обновлены")
except:
    logging.error(f"ОШИБКА", exc_info=True)
    
    
# список с адресами рассылки
lst_email = read_email_adress()
lst_email_error = ['skrutko@sim-auto.ru'] # есть ошибки
# запуск функции рассылки почты
logging.info(f"детектим ошибки, проверяем почту")
sending_mail(lst_email, lst_email_error)
logging.info(f"почта отправлена")

from pywinauto.application import Application
import schedule
import os
import datetime
import time


def create_dir(filename):
    # Функция создает директорию C:\LOG\год\месяц
    # % B - месяц,% Y - год
    path = r'C:\LOG\{year}\{month}'.format(
            year=filename.strftime("%Y"),
            month=filename.strftime("%B")
        )
    os.makedirs(path)                                             # создаем директорию
    print('directory was created')


def change_log():
    filename = datetime.date.today()  # получаем текущую дату
    '''Создаем экземляр класса приложения, подключаемся к уже запущенному FarScan через PID процесса'''
    app = Application().connect(process=344)
    '''Альтернаивное имя окна программы "FS4W_Manager_Class"'''
    win = app.FS4W_Manager_Class                                    # определяем окно программы.
    win.menu_select("Конфигруация->Режим файла регистрации...")     # выбирает нужный элемент меню программы
    win_journal = app["Режим дискового\принтерного журнала"]        # обьект окна "Режим дискового\принтерного журнала" присваиваем переменной
    win_journal.Отключить.click()                                   # кликает Отключить в окне выбора пути лога
    win_journal.Да.click()                                          # кликает Да в окне выбора пути лога
    win.menu_select("Конфигруация->Режим файла регистрации...")
    win_journal.Edit.click()                                                # кликнуть на поле пути лога
    win_journal.Edit.select()                                               # выбделить текст пути лога
    win_journal.Edit.type_keys(r'C:\LOG\{year}\{month}\{day}.log'.format(   # ввод пути для лога в соответсвии с текущей датой
                year=filename.strftime("%Y"),
                month=filename.strftime("%B"),
                day=filename.strftime("%d")
        )
    )
    win_journal.Включить.click()                                    # кликает Включить в окне выбора пути лога
    win_journal.Да.click()
    print('log was changed')


schedule.every().day.at('00:01').do(change_log)                     # создание нового лог-файла каждый день в 00:01

# запуск программы
while True:
    log_name = datetime.date.today()
    directory = r'C:\LOG\{year}\{month}'.format(
        year=log_name.strftime("%Y"),
        month=log_name.strftime("%B"),
    )
    # если наступило 1 число месяца и директория не была создана
    if log_name.strftime("%d") == '01' and not os.path.exists(directory):
        create_dir(log_name)  # создаем директорию
    schedule.run_pending()  # запуск функции создания лога
    time.sleep(1)  # пауза 1 сек

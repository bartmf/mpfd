import time
import threading
import psutil
from datetime import datetime
import xlsxwriter


loop = True
row = 2


def new_xlsx(name_file=str):
    workbook = xlsxwriter.Workbook(name_file.split('.')[0] + datetime.now().strftime('_%H-%M__%m_%d_%Y') + '.xlsx')
    worksheet = workbook.add_worksheet('Показатели')
    bold = workbook.add_format({'bold': 1})
    head = ['Time', 'Cpu %', 'Memory %']
    worksheet.write_row('A1', head, bold)
    return [workbook, worksheet, name_file]


def write_value_to_xlsx(date, cpu_value, mem_value,
                        worksheet=xlsxwriter.Workbook.worksheet_class):
    global row
    worksheet.write_row("A%d" % row, [date, cpu_value, mem_value])
    row += 1


def creat_graf(workbook=xlsxwriter.Workbook, name_proc = str):
    worksheet = workbook.add_worksheet('График')
    chart1 = workbook.add_chart({'type': 'line'})
    chart1.add_series({
        'name': ['Показатели', 0, 1],
        'categories': ['Показатели', 1, 0, row - 1, 0],
        'values': ['Показатели', 1, 1, row - 1, 1],
    })
    chart1.add_series({
        'name': ['Показатели', 0, 2],
        'categories': ['Показатели', 1, 0, row - 1, 0],
        'values': ['Показатели', 1, 2, row - 1, 2],
    })
    chart1.set_size({'width' : 800, 'height' : 500})
    chart1.set_title({'name': "График потребления %s"%name_proc})
    chart1.set_x_axis({'name': 'дата/время'})
    chart1.set_y_axis({'name': 'Загрузка в %'})
    chart1.set_style(11)
    worksheet.insert_chart('A1', chart1, {'x_offset': 50, 'y_offset': 30})


def find_process(name=str):
    proc_name = name
    for iterProc in psutil.process_iter():
        if iterProc.name() == proc_name:
            print("Процесс найде - Process ID = %d" % iterProc.pid)
            return iterProc
        else:
            continue
    print('Процесс не найден, попробуйте снова')
    return False


def info_process(proc=psutil.Process, time_sleep = int):
    # f = open('/home/bart/Документы/test.txt', 'w')
    # while (loop):
    #     f.write('date: %s\ttime: %s\tprocessor_name: %s\tcpu = %d\tmemmory = %d\n' % (
    #         datetime.now().date(), datetime.now().time(), proc.name(), proc.cpu_percent(), proc.memory_percent()))
    #     time.sleep(10)

    xls_blocks = new_xlsx(proc.name())
    while (loop):
        write_value_to_xlsx(datetime.now().strftime('%H:%M:%S'), proc.cpu_percent(), proc.memory_percent(), xls_blocks[1])
        time.sleep(time_sleep)
    creat_graf(xls_blocks[0], xls_blocks[2])
    xls_blocks[0].close()
    print('Отчёт сформирован')


if __name__ == "__main__":
    arg = False
    inp = str
    while (1):
        inp = input('Введите имя процесса\t(для выхода из программы - введите \"exit\")\n')
        time_slepp = int(input('Введите выборку в секундах\n'))
        if (inp == 'exit'):
            exit(0)
        else:
            arg = find_process(inp)
            if (arg == False):
                continue
            break
    thread = threading.Thread(target=info_process, args=(arg, time_slepp, ))
    thread.start()
    date_start = datetime.now()
    while (loop):
        inp = input("Введите номер пункта для управления.\n"
                    "Меню программы:\n"
                    "1. Узнать время работы программы\n"
                    "2. Выход из программы\n"
                    )
        if (inp == '1'):
            print((datetime.now() - date_start))
        else:
            loop = False
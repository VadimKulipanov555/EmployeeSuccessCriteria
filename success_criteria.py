import os
import pandas as pd
import shutil


class ProgramError(Exception):
    def __init__(self, text):
        self.text = text

# Отсеиваются все файлы не формата xlsx и сохраняет имена файлов xlsx, а так-же делая их копии


def ReadingXlsxFiles():
    file_list = []

    # Добавляем все файлы с расширением xlsx
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".xlsx"):
                file_list.append(os.path.join(file))

    # Делаем копии файлов, для дальнейшей работы с ними
    for i in file_list:
        open('DataExcelCopy/%s' % i, 'w').close()
        shutil.copyfile('DataExcel/%s' % i, 'DataExcelCopy/%s' % i)

    DeletingProjectsOfTheSameName(file_list)
    DataFromExcel(file_list)


# Проходим по всем файлам и удаляет проект в файле, если в следующих файлах есть проект с таким-же названием.
# Так-же удаляет и проекты с одинаковыми названиями в данном файле, оставляя последний найденный.


def DeletingProjectsOfTheSameName(file_list):
    # Проходим по всем файлам
    for name_1 in file_list:
        Book_1 = pd.read_excel('DataExcelCopy/%s' % name_1)
        for name_2 in file_list:
            # Сравнение в самом файле или сравнение с другим файлом
            if name_1 != name_2:
                # Сравниваем разные файлы
                Book_2 = pd.read_excel('DataExcelCopy/%s' % name_2)
                project_book_1, project_book_2 = list(Book_1['Название проекта']), list(Book_2['Название проекта'])

                # Проходим по всем строка в двух файлах сравнивая названия проектов
                for i in project_book_1:
                    for j in project_book_2:
                        if i == j:
                            # Удаляем все проекты с именем i в столбце 'Название проекта'
                            Book_1 = Book_1[Book_1['Название проекта'].ne(i)]

                            # Сохраняем файл
                            Book_1.to_excel('DataExcelCopy/%s' % name_1, index=False)
            else:
                # Сравниваем один файл
                project_3 = list(Book_1['Название проекта'])
                for i in set(project_3):
                    if project_3.count(i) > 1:
                        # Сохраняем последнее вхождение строки с именем i в столбце 'Название проекта'
                        present = Book_1[Book_1['Название проекта'] == i].iloc[-1:]

                        # Удаляем все проекты с именем i в столбце 'Название проекта'
                        Book_1 = Book_1[Book_1['Название проекта'].ne(i)]

                        # Добавляем ранее сохраненную строку в файл
                        Book_1 = Book_1.append(present, ignore_index=False)

                        # Сохраняем файл
                        Book_1.to_excel('DataExcelCopy/%s' % name_1, index=False)


# Сохраняем все данные в файл rate.xlsx по убыванию успешности сотрудника


def SavingDataToXlsx(info_all_all):
    names = []
    year = []
    total = []
    employee_success_rate = pd.DataFrame()

    # Проходим по ключам словаря в словаре и сохраняем их в списки
    for entry in info_all_all.keys():
        for i in info_all_all[entry].keys():
            names.append(i)
            year.append(entry)
            total.append(info_all_all[entry][i] / 200)

    # Одинаково сортируем списки по уменьшению total
    total, names, year = zip(*[(c, b, a) for c, b, a in sorted(zip(total, names, year), reverse=True)])

    # Записываем все данные
    employee_success_rate['name'] = names
    employee_success_rate['year'] = year
    employee_success_rate['total'] = total

    # Сохраняем файл
    employee_success_rate.to_excel('DataExcel/rate.xlsx')
    print('You will find the data file on the path DataExcel/rate.xlsx')


# Оставляем данные подходящие под критерий успешности и считает работу сотрудника по оставшимся проектам
# Критерий успешности - сколько трудодней в каждом году в непросроченных проектах есть у сотрудника фактически/ 20


def DataFromExcel(file_list):
    # Будет хранить данные {'человек': количество часов }
    info_book = {}

    # Будет хранит все имена сотрудников(.., факт)
    list_keys = []

    # Будет хранит года без повторов
    list_data = []

    # Будет хранить всю информацию из всех файлов ({'год': {'человек': успешность }})
    info_all = {}

    # Проходимся по всем файлам
    for name in file_list:
        # Считываем данные из файла
        Book = pd.read_excel('DataExcelCopy/%s' % name)

        # Убираем все проекты не подходящие условию
        Book = Book[(Book['Дата сдачи, план.']) >= Book['Дата сдачи, факт.']]

        # Сохраняем все года проектов в список
        for data in list(Book['Дата сдачи, план.']):
            list_data.append((str(data).split('-'))[0])

        # Убираем копии годов
        list_data = sorted(list(set(list_data)))

        # Запоминаем все имена сотрудников(.., факт)
        list_keys = list(set(list_keys + list(Book.keys()[6::2])))
        info_file = {}

        # Делаем словарь в словаре
        for i in list_data:
            info_file[i] = info_book

        # Проходим по годам
        for entry in list_data:
            # Берем только те проекты, которые были выполнены в определенном году
            Books = Book[(Book['Дата сдачи, план.'] >= '%s-01-01' % entry) & (
                    Book['Дата сдачи, план.'] <= '%s-12-31' % entry)]

            if 'Unnamed' in str(list_keys):
                raise ProgramError('Columns with empty names of people were found in the file, '
                                   'in fact, which have values in projects.')

            # Записываем данные {'человек': количество часов }
            for i in list_keys:
                if i in Books.keys():
                    info_book[i] = Books[i].sum()

            # Записываем данные ({'год': {'человек': количество часов }})
            info_file[entry] = info_book

            # Складываем количество часов человека за целый год
            for i in info_file.keys():
                if i in info_book.keys():
                    info_file[entry][i] += info_book[i]

            # Освобождаем для новых данных следующего года
            info_book = {}

        # Складываем данные из всех файлов в ({'год': {'человек': успешность }}
        for i in info_file.keys():
            for j in info_file[i].keys():
                if i in info_all.keys():
                    if j in info_all[i].keys():
                        info_all[i][j] += info_file[i][j]
                    else:
                        info_all[i][j] = info_file[i][j]
                else:
                    info_all[i] = {}
                    if j in info_all[i].keys():
                        info_all[i][j] += info_file[i][j]
                    else:
                        info_all[i][j] = info_file[i][j]

    SavingDataToXlsx(info_all)


try:
    # В данном файле будут содержаться все файлы xlsx
    if os.path.isdir('DataExcel'):
        path = 'DataExcel'

        if os.path.isfile('DataExcel/rate.xlsx'):
            os.remove('DataExcel/rate.xlsx')

        # В данной папке будут храниться копии файлов xlsx с ними мы и будем работать в программе
        if not os.path.isdir('DataExcelCopy'):
            os.mkdir('DataExcelCopy')

        ReadingXlsxFiles()
    else:
        os.mkdir('DataExcel')
        print('Put all the files in the DataExcel folder .xlsx, '
              'the folder has already been created. And start the program again!')

except IOError:
    print("Could not open file! Please close the excel file in the folder DataExcel!")
except ValueError:
    print('The program has nothing to process, perhaps you have not placed the files in the DataExcel folder, '
          'perhaps there is no data in the Excel files or there are no projects submitted in time, '
          'so there is nothing to process.')
except KeyError:
    print('The file does not comply with the norms of the table, namely the name of significant columns.')
except ProgramError as pe:
    print(pe)
except Exception:
    print('Report the error found. ' + str(Exception))

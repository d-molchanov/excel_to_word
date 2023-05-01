from os import walk
from os.path import abspath, relpath, splitext, join, basename

from time import perf_counter

import openpyxl
from docx import Document
from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL

from directive import Directive

def timeit(operation, _end):
    def decorator(method):
        def wrapped(*args, **kwargs):
            base_name = basename(kwargs['_path'])
            print(f'{operation} <{base_name}>...', end=_end)
            time_start = perf_counter()
            result = method(*args, **kwargs)
            time_finish = (perf_counter() - time_start)*1e3
            print(f'{operation} <{base_name}> done in {time_finish:.3f} ms.')
            return result
        return wrapped
    return decorator

def create_xlsx_file_list(target_dir):
    result = []
    for root, dirs, files in walk(target_dir):
        for f in files:
            ap_file = abspath(join(root, f))
            filename, ext = splitext(ap_file)
            if ext == '.xlsx':
                result.append(ap_file)
    return result

@timeit('Reading', '\r')
def read_xlsx(_path):
    wb = openpyxl.load_workbook(_path)
    sheet = wb.active
    rows = sheet.rows
    data = [[cell.value for cell in row] for row in sheet.rows]
    return [row for row in data if row != [None for el in row]]

def extract_columns(data, columns):
    return [[row[el] for el in columns] for row in data] 

def change_ext(filename, new_ext):
    name, ext = splitext(filename)
    return f'{name}.{new_ext}'


def convert_data_to_str(data, formatting):
    return [
        [f.format(r).replace('.', ',') if type(r) != str else r for r, f in 
        zip(row, formatting)] for row in data
    ]


def scan_directory(target_dir, filenames):
    result = dict()
    for root, dirs, files in walk(target_dir):
        found_files = [f for f in files if f in filenames]
        found_files.sort()
        if found_files == filenames:
            result[f'{abspath(root)}'] = found_files
    print(*list(result.keys()), sep='\n')
    return result

@timeit('Writing', '\r')
def write_txtfile(data, sep, _path):
    try:
        with open(_path, 'w') as f:
            for row in data:
                f.write(f"{sep.join(row)}\n")
    except IOError:
        print(f'I/O error with <{_path}>.')

@timeit('Reading', '\r')
def read_textfile(_path):
    try:
        with open(_path, 'r', encoding='utf-8') as f:
            return [line.rstrip() for line in f.readlines()]
    except IOError:
        print(f'I/O error with <{_path}>.')
        return None 


def enumerate_part_of_directive(part_number, part_of_directive):
    result = [f'{part_number}. {part_of_directive[0]}']
    if len(part_of_directive) > 1:
        for i, p in enumerate(part_of_directive[1:], 1):
            result.append(f'{part_number}.{i}. {p}')
    return result

def create_directive_text(directive_template, appendixes_2_and_3_are_equal):
    result = directive_template[:2]
    part_1 = directive_template[2:3]
    part_2 = directive_template[3:4]
    if appendixes_2_and_3_are_equal:
        part_2.append(directive_template[6])
    else:
        part_2 += directive_template[4:6]
    part_2 += directive_template[7:8]
    part_3 = directive_template[8:9]
    part_4 = directive_template[9:17]
    part_5 = directive_template[17:18]
    parts = [part_1, part_2, part_3, part_4, part_5]
    for i, p in enumerate(parts, 1):
        result += enumerate_part_of_directive(i, p)
    result += directive_template[-4:]
    return result

#!Rewrite with list comprehension
def create_document_framework(template_data, indices, separators):
    result = []
    for i, j, s in zip(indices[:-1], indices[1:], separators):
        result.append(s.join(template_data[i:j]))
    return result

def create_appendix_content(framework, appendix_number):
    indices = []
    if appendix_number == '1':
        indices += [2, 6]
    elif appendix_number == '2':
        indices += [3, 7]
    elif appendix_number == '3':
        indices += [4, 8]
    else:
        indices += [5, 9]
        appendix_number = '2'
    result = [
        framework[0].format(AN=appendix_number),
        framework[1]
    ]
    for i in indices:
        result.append(framework[i])
    return result

def write_docx_file(document, output_file):
    try:
        document.save(output_file)
    except PermissionError:
        print(f'<{output_file}> is busy - permission denied.')

@timeit('Processing', '\n')
def process_waterbody(filenames, directive_template, appendix_framework, _path):
    # print(f'Start processing <.{relpath(target_dir, ".")}>:')

    content_txt = read_textfile(_path=join(_path, 'content.txt'))
    subst_keys = ['WBN', 'DN', 'WBL', 'WPZ', 'PSB']
    substitution = {k:f'{{{k}}}' for k in subst_keys}
    if content_txt:
        # for k, c in zip(subst_keys[1:], content_txt[1:]):
        for k, c in zip(subst_keys, content_txt):
            substitution[k] = c

    waterbody_name = substitution['WBN']

    xlsx_files = [f for f in filenames if splitext(f)[1] == '.xlsx']
    xlsx_files.sort()
    appendix_document = Document()
    xlsx_data = []
    coordinates_title = []
    appendix_numbers = []
    for i, f in enumerate(xlsx_files, 1):
        appendix_numbers.append(str(i))
        data = read_xlsx(_path=join(_path, f))
        xlsx_data.append(data)
        coordinates_title.append(data[2][1])
    coordinates = [extract_columns(el[5:], [0, 1, 2]) for el in xlsx_data]
    formatting = ['{:.0f}', '{:.2f}', '{:.2f}']
    str_coordinates = [convert_data_to_str(el, formatting) for el in coordinates]

    for f, d in zip(xlsx_files, str_coordinates):
        new_filename = change_ext(f, 'txt')
        write_txtfile(extract_columns(d, [1, 2]), _path=join(_path, new_filename), sep='\t')


    appendixes_2_and_3_are_equal = False
    if coordinates[1] == coordinates[2]:
        appendix_numbers = ['1', '23']
        appendixes_2_and_3_are_equal = True
    
    directive_text = create_directive_text(directive_template, appendixes_2_and_3_are_equal)
    directive = Directive()
    doc = directive._create_directive(directive_text, substitution)
    directive._set_page_properties(doc)
  
    for c, c_t, n in zip(str_coordinates, coordinates_title, appendix_numbers):
        a_c = create_appendix_content(appendix_framework, n)
        doc = directive._add_appendix(doc, a_c, substitution)
        doc = directive._add_table(doc, c, c_t)

    directive._create_right_numeration(doc)

    doc.save(join(_path, 'Directive.docx'))

@timeit('Processing', '\n')
def process_directory(filenames, _path):
    dir_content = scan_directory(_path, filenames)

    directive_template = read_textfile(_path='directive_template.txt')
    appendix_content = read_textfile(_path='appendix_template.txt')
    
    appendix_framework = create_document_framework(
        appendix_content,
        [0, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
        ['\n', '', '', '', '', '', '', '', '', '']
    )

    for directory, files in dir_content.items():
        process_waterbody(files, directive_template, appendix_framework, _path=directory)

if __name__ == '__main__':
    # target_dir = './data/26_река_Нахавня_(Одинцовские г.о.)'
    # target_dir = './data/10_ручей_без_названия_(г.о. Егорьевск)'
    # target_dir = './data/12_река_Плесенка_(Наро-Фоминский г.о.)'
    # target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
    # target_dir = './data/(2023_04_02)/13_река_Шатуха_(Наро-Фоминский г.о., Рузский г.о.)'
    # target_dir = './data/(2023_04_02)/16_река_Сумерь_(г.о. Пушкинский, Сергиево-Посадский г.о.)'
    # target_dir = './data/(2023_04_02)/26_река_Нахавня_(Одинцовские г.о.)'
    # target_dir = './data/(2023_04_02)/33_река_Малые_Вяземы_(Одинцовский г.о.)'
    # target_dir = './data/(2023_04_02)/34_ручей_без_названия_(г.о. Домодедово)'
    # target_dir = './data/(2023_04_02)/37_ручей_без_названия_(г.о. Домодедово)'
    # target_dir = './data/(2023_04_02)/42_река_Беляна_(Одинцовский г.о., г.о. Истра)'
    # target_dir = './data/(2023_04_02)/44_река_Жданка_(Раменский г.о., г.о. Домодедово)'
    # target_dir = './data/(2023_04_02)/46_река_Лубянка_(г.о. Ступино)'
    # target_dir = './data/(2023_04_02)/49_река_Камариха_(г.о. Пушкинский, Дмитровский г.о.)'
    # target_dir = './data/(2023_04_02)/50_река_Вырка_(Орехово-Зуевский г.о.)'

    # target_dir ='./data/Проекты распоряжений'
    # target_dir ='./data/Проекты распоряжений/Лотошинский район/299 река Черная test'
    target_dir ='./data/Проекты распоряжений/Рузский район'
    # target_dir ='./data/Проекты распоряжений/Лотошинский район/299 река Черная'
    # target_dir ='./data/Проекты распоряжений/Лотошинский район/865 река Безымянная'

    # filenames = [
    #     'Приложение 1.xlsx',
    #     'Приложение 2.xlsx',
    #     'Приложение 3.xlsx',
    #     'content.txt'
    # ]

    filenames = [
        'каталог координат БЛ.xlsx',
        'каталог координат ВОЗ.xlsx',
        'каталог координат ПЗП.xlsx',
        'content.txt'
    ]
    filenames.sort()

    process_directory(filenames, _path=target_dir)
    
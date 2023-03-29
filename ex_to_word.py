from os import walk
from os.path import abspath, splitext, join

from time import perf_counter

import openpyxl
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_xlsx_file_list(target_dir):
	result = []
	for root, dirs, files in walk(target_dir):
		for f in files:
			ap_file = abspath(join(root, f))
			filename, ext = splitext(ap_file)
			if ext == '.xlsx':
				result.append(ap_file)
	return result


def read_xlsx(filename):
	wb = openpyxl.load_workbook(filename)
	sheet = wb.active
	data = [[cell.value for cell in row] for row in sheet.rows]
	return [row for row in data if row != [None for el in row]]

def extract_columns(data, columns):
	return [[row[el] for el in columns] for row in data] 

def write_data_new(data, columns, filename):
	data_to_write = []
	for row in data:
		data_to_write.append([row[i] for i in columns])

	document = Document()
	table = document.add_table(rows=len(data), cols=len(columns))
	table.style = 'Table Grid'
	table.allow_autofit = False
	for col in table.columns:
		col.width = Cm(2.5)
	for i, (row_read, row_write) in enumerate(zip(data_to_write, table.rows)):
		time_start = perf_counter()
		for el, cell in zip(row_read, row_write.cells):
			cell.text = str(el)
		print(f'Row {i+1}:\t{round((perf_counter() - time_start)*1000, 3)}\tms')
	document.add_paragraph()
	document.save(filename)

def write_data(data, columns, filename):
	document = Document()
	indent_style = document.styles.add_style('Indent', WD_STYLE_TYPE.PARAGRAPH)
	indent_style.paragraph_format.left_indent = Cm(10)
	par1 = document.add_paragraph('Приложение 2\nк распоряжению\nМинистерства экологии\nи природопользования\nМосковской области\n№______ от _____________', style=indent_style)
	par1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
	# document.sections[0].left_margin = Cm(10)

	par2 = document.add_paragraph()
	par2.alignment = WD_ALIGN_PARAGRAPH.CENTER
	par2.add_run('Границы водоохранной зоны, прибрежной защитной полосы\nручья без названия в Сергиево-Посадском городском округе Московской области').bold = True
	par3 = document.add_paragraph()
	par3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
	par3.add_run('Координаты границ водоохранной зоны, прибрежной защитной полосы ручья без названия в Сергиево-Посадском городском округе\nМосковской области.')
	table_section = document.add_section(WD_SECTION.CONTINUOUS)
	sectPr = table_section._sectPr
	cols = sectPr.xpath('./w:cols')[0]
	cols.set(qn('w:num'),'2')
	table = document.add_table(rows=0, cols=len(columns), style='Table Grid')
	table.allow_autofit = False
	table.columns[0].width = Cm(1.5)
	for col in list(table.columns)[1:]:
		col.width = Cm(2.5)
	for i, row in enumerate(data):
		time_start = perf_counter()
		row_cells = table.add_row().cells
		input_cells = [row[i] for i in columns]
		# print(input_cells)
		for c_in, c_out in zip(input_cells, row_cells):
			c_out.text = str(c_in)
			# c_out.style = my_style
		# print(f'Row {i+1}:\t{round((perf_counter() - time_start)*1000, 3)}\tms')
	last_section = document.add_section(WD_SECTION.CONTINUOUS)
	sectPr = last_section._sectPr
	cols = sectPr.xpath('./w:cols')[0]
	cols.set(qn('w:num'),'1')
	document.save(filename)

def write_txt_file(data, filename, sep):
	try:
		with open(filename, 'w') as f:
			for row in data:
				str_data = ['{:.2f}'.format(el).replace('.', ',') for el in row if type(el) != str]
				f.write(f"{sep.join(str_data)}\n")
	except IOError:
		print(f'I/O error with <{filename}>.')

def change_ext(filename, new_ext):
	name, ext = splitext(filename)
	return f'{name}.{new_ext}'
# print('Start reading data.')
# data = read_xlsx('data/26_река_Нахавня_(Одинцовские г.о.)/Приложение 1.xlsx')
# print('Reading data finished. Start writing data')
# # write_data_new(data[4:200], [0, 4, 5], 'data/26_река_Нахавня_(Одинцовские г.о.)/Приложение 1.docx')
# write_data(data[4:2000], [0, 4, 5], 'data/26_река_Нахавня_(Одинцовские г.о.)/Приложение 1.docx')
# print('Writing data finished')

if __name__ == '__main__':
	# target_dir = './data/26_река_Нахавня_(Одинцовские г.о.)'
	target_dir = './data/10_ручей_без_названия_(г.о. Егорьевск)'
	files = create_xlsx_file_list(target_dir)
	print(f'List of xlsx files in <{abspath(target_dir)}>:')
	for f in files[:1]:
		print(f)
		data = read_xlsx(f)
		new_filename = change_ext(f, 'txt')
		write_txt_file(extract_columns(data[4:], [4, 5]), new_filename, '\t')
		new_filename = change_ext(f, 'docx')
		write_data(data[4:], [0, 4, 5], new_filename)
	# filename = files[0]
	# data = read_xlsx(filename)
	# print(len(data))
	# new_filename = change_ext(filename, 'txt')
	# write_txt_file(extract_columns(data[4:], [4, 5]), new_filename, '\t')

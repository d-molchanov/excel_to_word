from os import walk
from os.path import abspath, splitext, join

from time import perf_counter

import openpyxl
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

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
	rows = sheet.rows
	temp = [cell.value for cell in next(rows)]
	pattern = [None for el in temp]
	data = []
	while temp != pattern:
		data.append(temp)
		temp = [cell.value for cell in next(rows)]
	# data = [[cell.value for cell in row] for row in sheet.rows]
	return data
	# return [row for row in data if row != [None for el in row]]

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

def convert_data_to_str(data, formatting):
	return [
		[f.format(r).replace('.', ',') for r, f in 
		zip(row, formatting)] for row in data
	]

def create_docx_document(content):
	document = Document()

	header = document.sections[0]
	header.page_width = Cm(21)
	header.page_height = Cm(29.7)
	header.left_margin = Cm(3)
	header.right_margin = Cm(1.5)
	header.top_margin = Cm(2)
	header.bottom_margin = Cm(2)
	appendix_style = document.styles.add_style('Appendix Title', WD_STYLE_TYPE.PARAGRAPH)
	appendix_style.paragraph_format.left_indent = Cm(11)
	number_style = document.styles.add_style('Document Number', WD_STYLE_TYPE.PARAGRAPH)
	number_style.paragraph_format.left_indent = Cm(10.5)
	title_style = document.styles.add_style('Document Title', WD_STYLE_TYPE.PARAGRAPH)
	subtitle_style = document.styles.add_style('Document Subtitle', WD_STYLE_TYPE.PARAGRAPH)
	subtitle_style.paragraph_format.first_line_indent = Cm(1.25)

	styles = [
		appendix_style,
		number_style,
		title_style,
		subtitle_style
	]
	alignments = [
		WD_ALIGN_PARAGRAPH.LEFT,
		WD_ALIGN_PARAGRAPH.LEFT,
		WD_ALIGN_PARAGRAPH.CENTER,
		WD_ALIGN_PARAGRAPH.JUSTIFY
	]

	for s, a, c in zip(styles, alignments, content):
		s.font.name = 'Times New Roman'
		s.font.size = Pt(13)
		s.paragraph_format.line_spacing = 1.06
		# s.paragraph_format.space_after = s.font.size
		s.paragraph_format.space_after = Pt(0)
		p = document.add_paragraph(style=s)
		p.alignment = a
		p.add_run(c)

	title_style.paragraph_format.space_before = Pt(14)
	title_style.font.size = Pt(14)
	title_style.paragraph_format.space_after = Pt(14)
	subtitle_style.paragraph_format.space_after = Pt(14)
	subtitle_style.font.size = Pt(14)


	# appendix_style.font.size = Pt(13)
	# appendix_style.paragraph_format.line_spacing = 1.08
	document.paragraphs[2].runs[0].bold = True

	return document

def add_table_title(table, koord_zone):
	first_row_cells = table.add_row().cells
	for c in first_row_cells:
		c.paragraphs[0].style.font.name = 'Times New Roman'
		c.paragraphs[0].style.font.size = Pt(12)

	for i in range(2):
		table.add_row()


	cells = ((0, 0), (0, 1), (1, 1), (1, 2), (2, 0), (2, 1), (2, 2))

	text = [
		'№\nп/п',
		koord_zone,
		'X',
		'Y',
		'(1)',
		'(2)',
		'(3)'
	]

	table.cell(0,0).merge(table.cell(1,0))
	table.cell(0,1).merge(table.cell(0,2))
	


	# table.cell(0,0).paragraphs[0].style.paragraph_format.space_before = Pt(12)
	# table.cell(0,1).paragraphs[0].style.paragraph_format.space_before = Pt(6)
	# table.cell(0,1).paragraphs[0].style.paragraph_format.space_after = Pt(6)
	# table.cell(1,1).paragraphs[0].style.paragraph_format.space_before = Pt(0)
	# table.cell(1,1).paragraphs[0].style.paragraph_format.space_after = Pt(0)
	# table.cell(1,2).paragraphs[0].style.paragraph_format.space_before = Pt(0)
	# table.cell(1,2).paragraphs[0].style.paragraph_format.space_after = Pt(0)
	# table.cell(2,0).paragraphs[0].style.paragraph_format.space_before = Pt(0)
	# table.cell(2,0).paragraphs[0].style.paragraph_format.space_after = Pt(0)
	# table.cell(2,1).paragraphs[0].style.paragraph_format.space_before = Pt(0)
	# table.cell(2,1).paragraphs[0].style.paragraph_format.space_after = Pt(0)
	# table.cell(2,2).paragraphs[0].style.paragraph_format.space_before = Pt(0)
	# table.cell(2,2).paragraphs[0].style.paragraph_format.space_after = Pt(0)
	
	table.cell(0,0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
	table.cell(0,1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
	table.rows[0].height = Cm(1.96)

	for c, t in zip(cells, text):
		table.cell(*c).text = t
		table.cell(*c).paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
		table.cell(*c).paragraphs[0].runs[0].bold = True
	
	# print('cell(0,0)', table.cell(0,0).paragraphs[0].style.paragraph_format.space_before)
	# print('cell(1,0)', table.cell(1,0).paragraphs[0].style.paragraph_format.space_before)
	# print('cell(0,1)', table.cell(0,1).paragraphs[0].style.paragraph_format.space_before)
	# print('cell(0,2)', table.cell(0,2).paragraphs[0].style.paragraph_format.space_before)

	return table

def add_table(document, data, koord_zone):
	table_section = document.add_section(WD_SECTION.CONTINUOUS)
	sectPr = table_section._sectPr #table_section._sectPr
	cols = sectPr.xpath('./w:cols')[0]
	cols.set(qn('w:num'),'2')

	table = document.add_table(rows=0, cols=len(data[0]), style='Table Grid')
	table = add_table_title(table, koord_zone)
	table.allow_autofit = False
	table.columns[0].width = Cm(2)
	for row in data:
		cells = table.add_row().cells
		for c, r in zip(cells, row):
			c.text = r
			c.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

	footer_section = document.add_section(WD_SECTION.CONTINUOUS)
	sectPr = footer_section._sectPr
	cols = sectPr.xpath('./w:cols')[0]
	cols.set(qn('w:num'),'1')		
	return document

def scan_directory(target_dir, filenames):
	result = dict()
	for root, dirs, files in walk(target_dir):
		found_files = [f for f in files if f in filenames]
		if found_files:
			result[f'{abspath(root)}'] = found_files
	return result

def write_txtfile(data, filename, sep):
	try:
		with open(filename, 'w') as f:
			for row in data:
				f.write(f"{sep.join(row)}\n")
	except IOError:
		print(f'I/O error with <{filename}>.')

def read_txtfile(filename):
	try:
		with open(filename, 'r') as f:
			return [line.rstrip() for line in f.readlines()]
	except IOError:
		print(f'I/O error with <{filename}>.')
		return None	

def process_directory(dir_dict):
	target_dir = list(dir_dict.keys())[0]
	print(f'List of files in <{target_dir}> :')
	list_of_files = list(dir_dict.values())[0]
	for i, f in enumerate(list_of_files, 1):
		print(f'{i}\t{f}')
	if 'content.txt' in list_of_files:
		content_txt = read_txtfile(join(target_dir, 'content.txt'))
	xlsx_files = [f for f in list_of_files if splitext(f)[1] == '.xlsx']
	xlsx_files.sort()
	xlsx_data = []
	koord_zone = []
	for f in xlsx_files:
		time_start = perf_counter()
		print(f'Start reading <{f}>.')
		data = read_xlsx(join(target_dir, f))
		xlsx_data.append(data)
		koord_zone.append(data[1][4])
		print(f'{len(data)} rows has been read in {round((perf_counter() - time_start)*1e3, 3)} ms.')
	partial_data = [extract_columns(el[4:], [0, 4, 5]) for el in xlsx_data]
	if partial_data[1] == partial_data[2]:
		print('\nWARNING: appendix 2 equals to appendix 3!\n')
	formatting = ['{:.0f}', '{:.2f}', '{:.2f}']
	str_data = [convert_data_to_str(el, formatting) for el in partial_data]
	print('Start creating txt files.')
	for f, d in zip(xlsx_files, str_data):
		time_start = perf_counter()
		new_filename = change_ext(f, 'txt')
		write_txtfile(extract_columns(d, [1, 2]), join(target_dir, new_filename), '\t')
		print(f'<{new_filename}> has been created in {round((perf_counter() - time_start)*1e3, 3)} ms.')
	print('Start creating docx files.')
	content = [
		'\n'.join([
			'Приложение {}',
			'к распоряжению',
			'Министерства экологии',
			'и природопользования',
			'Московской области'
		]),
		'№______ от _____________',
		'\n'.join([
			'Местоположение береговой линии (границы водного объекта)',
			'{}'
		]),
		'\n'.join([
			'Границы водоохранной зоны',
			'{}'
		]),
		'\n'.join([
			'Границы прибрежной защитной полосы',
			'{}'
		]),
		'Координаты местоположения береговой линии (границы водного объекта) {}.',
		'Координаты границ водоохранной зоны {}.',
		'Координаты прибрежной защитной полосы {}.'
	]

	for f, d, k_z in zip(xlsx_files, str_data, koord_zone):
		time_start = perf_counter()
		appendix_number = f[-6]
		insert_content = None
		if appendix_number == '1':
			insert_content = [content[el] for el in [0, 1, 2, 5]]
		elif appendix_number == '2':
			insert_content = [content[el] for el in [0, 1, 3, 6]]
		elif appendix_number == '3':
			insert_content = [content[el] for el in [0, 1, 4, 7]]
		
		insert_content[0] = insert_content[0].format(appendix_number)
		insert_content[2] = insert_content[2].format('\n'.join(content_txt))
		insert_content[3] = insert_content[3].format(' '.join(content_txt))

		new_filename = change_ext(f, 'docx')
		document = create_docx_document(insert_content)
		document = add_table(document, d, k_z)
		try:
			document.save(join(target_dir, new_filename))
			print(f'<{new_filename}> has been created in {round((perf_counter() - time_start)*1e3, 3)} ms.')
		except PermissionError:
			print(f'<{new_filename}> is busy - permission denied.')


if __name__ == '__main__':
	# target_dir = './data/26_река_Нахавня_(Одинцовские г.о.)'
	# target_dir = './data/10_ручей_без_названия_(г.о. Егорьевск)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	target_dir = './data/(2023_04_02)/13_река_Шатуха_(Наро-Фоминский г.о., Рузский г.о.)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	# target_dir = './data/(2023_04_02)/6_река_Вьюница_(г.о. Шатура)'
	# target_dir = './data/(2023_04_02)/37_ручей_без_названия_(г.о. Домодедово)'

	
	filenames = [
		'Приложение 1.xlsx',
		'Приложение 2.xlsx',
		'Приложение 3.xlsx',
		'content.txt'
	]
	test = scan_directory(target_dir, filenames)
	
	process_directory(test)

	# files = create_xlsx_file_list(target_dir)
	# print(f'List of xlsx files in <{abspath(target_dir)}>:')
	# for f in files[:1]:
	# 	print(f)
	# 	data = read_xlsx(f)
		

	# 	content = ['\n'.join([
	# 			'Приложение 2',
	# 			'к распоряжению',
	# 			'Министерства экологии',
	# 			'и природопользования',
	# 			'Московской области'
	# 		]),
	# 		'№______ от _____________',
	# 		'\n'.join([
	# 			'Границы водоохранной зоны, прибрежной защитной полосы',
	# 			'ручья без названия в Сергиево-Посадском городском округе Московской области'
	# 		]),
	# 		'\n'.join([
	# 			'Координаты границ водоохранной зоны, прибрежной защитной полосы ручья без названия в Сергиево-Посадском городском округе',
	# 			'Московской области.'
	# 		])

	# 	]

	# 	formatting = ['{:.0f}', '{:.2f}', '{:.2f}']

	# 	partial_data = extract_columns(data[4:], [0, 4, 5])
	# 	str_data = convert_data_to_str(partial_data, formatting)

	# 	filename = 'demo.docx'
	# 	document = create_docx_document(content)
	# 	document = add_table(document, str_data)
	# 	try:
	# 		document.save(filename)
	# 	except PermissionError:
	# 		print(f'<{filename}> is busy - permission denied.')


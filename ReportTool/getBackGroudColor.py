#encoding=utf-8
import xlrd

book = xlrd.open_workbook("CSC63_RATS_Issue_List1.xls",formatting_info=True)
# sheets = book.sheet_names()
# print "sheets are:", sheets
# for index, sh in enumerate(sheets):
sheet = book.sheet_by_index(1)
print "Sheet:", sheet.name
rows, cols = sheet.nrows, sheet.ncols
print "Number of rows: %s Number of cols: %s" % (rows, cols)
# for row in range(rows):
	# for col in range(cols):
	# print "row, col is:", row+1, col+1,
	# thecell = sheet.cell(row, col)  # could get 'dump','value', 'xf_index'
	# print thecell.value,
	# xfx = sheet.cell_xf_index(row, col)
	# xf = book.xf_list[xfx]
	# bgx = xf.background.pattern_colour_index
	# print bgx
thecell = sheet.cell(1, 1)# could get 'dump','value', 'xf_index'
xfx = sheet.cell_xf_index(1, 1)
xf = book.xf_list[xfx]
print xf.background
bgx = xf.background.pattern_colour_index
print bgx
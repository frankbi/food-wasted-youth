# Parse the Loss-Adjusted Food Availability files at
# http://ers.usda.gov/data-products/food-availability-(per-capita)-data-system.aspx#.U2VX0HWx20j
# for the total consumer loss for each food type.

import xlrd, csv, sys

w = csv.writer(sys.stdout)

for foodgroup in ('Dairy', 'Fruit', 'grain', 'meat', 'veg'):
	book = xlrd.open_workbook(foodgroup+'.xls')
	for sheeti in range(book.nsheets):
		sheet = book.sheet_by_index(sheeti)
		foodname = sheet.name

		# find the row and columns we are interested in

		current_year_row = None
		for row in range(sheet.nrows):
			if sheet.cell(row, 0).value == 2010.0: # most recent years are sometimes missing stuff like rice
				current_year_row = row

		total_loss_col = None
		consumer_loss_col = None # not including "Nonedible share"
		for col in range(sheet.ncols):
			if sheet.cell(1,col).value == "Total loss, all levels":
				total_loss_col = col
			if sheet.cell(2,col).value == "Other (cooking loss and uneaten food)":
				consumer_loss_col = col

		if current_year_row is None: continue
		if consumer_loss_col is None: continue

		w.writerow((foodgroup, foodname, str(sheet.cell(current_year_row, consumer_loss_col).value)))



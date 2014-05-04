# Parse the Loss-Adjusted Food Availability files at
# http://ers.usda.gov/data-products/food-availability-(per-capita)-data-system.aspx#.U2VX0HWx20j
# for the total consumer loss for each food type.
#
# Usage:
# python3 parse_usda_excel_files.py > consumer_loss.csv

import xlrd, csv, sys

def str2(v):
	if v is None: return "NULL"
	return str(v)

w = csv.writer(sys.stdout)
w.writerow(("foodgroup", "foodcategory", "fooditem", "total_loss", "consumer_loss", "caloriesperday"))

for foodgroup in ('Dairy', 'Fruit', 'grain', 'meat', 'veg'):
	book = xlrd.open_workbook(foodgroup+'.xls')

	# The first worksheet is a table of contents that provides groupings
	# for the food items on the subsequent worksheets. Make a mapping
	# from food items to the name of their parent group.
	food_item_category = { }
	tocsheet = book.sheet_by_index(0)
	for col in range(1, tocsheet.ncols):
		category = None
		for row in range(tocsheet.nrows):
			v = tocsheet.cell(row, col).value
			if v.strip() == "": v = None
			if category is None or v is None:
				category = v
			elif category is not None and v is not None:
				food_item_category[v] = category

	# The following sheets are statistics by food item.
	for sheeti in range(1, book.nsheets):
		sheet = book.sheet_by_index(sheeti)
		foodname = sheet.name

		# find the row and columns we are interested in

		current_year_row = None
		for row in range(sheet.nrows):
			if sheet.cell(row, 0).value == 2010.0: # most recent years are sometimes missing stuff like rice
				current_year_row = row
		if current_year_row is None: continue

		total_loss_col = None
		consumer_loss_col = None # not including "Nonedible share"
		calories_daily_col = None

		for col in range(sheet.ncols):
			if sheet.cell(1,col).value == "Total loss, all levels":
				total_loss_col = col
			if sheet.cell(2,col).value == "Other (cooking loss and uneaten food)":
				consumer_loss_col = col
			if sheet.cell(1,col).value.startswith("Calories available daily"):
				calories_daily_col = col


		total_loss = sheet.cell(current_year_row, total_loss_col).value if total_loss_col is not None else None
		consumer_loss = sheet.cell(current_year_row, consumer_loss_col).value if consumer_loss_col is not None else None
		calories_daily = sheet.cell(current_year_row, calories_daily_col).value if calories_daily_col is not None else None

		w.writerow((foodgroup, str2(food_item_category.get(foodname)), foodname, str2(total_loss), str2(consumer_loss), str2(calories_daily)))



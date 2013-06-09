import sys
import glob
import xlrd
import xlwt
import os
import time
import logging

# Setup log
logging.basicConfig(filename='output.log', level=logging.DEBUG)
# Passed 3 arguments, excel.py, folder to open
try:
	folder = sys.argv[1]
except IndexError:
	logging.warning("No folder provided")
	folder = os.getcwd()
logging.info("Using folder " + folder)
# Get all excel files
files = glob.glob(folder + '/*.xls*')
sheet_name = u'For Summary File'
row_value = 2
# List to hold all the rows
total_values = []
# Output file name
output_file = 'compiled_info.xls'
read_col_vals = False
for file_name in files:
	logging.debug("Attempting to open file: " + file_name)
	try:
		# Open the book
		book = xlrd.open_workbook(file_name)
	except IOError:
		logging.warning("File: " + file_name + " does not exist")
	else:
		# Try to open the specific sheet
		try:
			summary_sheet = book.sheet_by_name('For Summary File')
		except:
			logging.warning('Could not open sheet ' + sheet_name + ' in file ' + file_name)
		else:
			logging.debug('Copying row.')
			if not read_col_vals:
				# Append col vals to list
				col_vals = [summary_sheet.cell(0, col).value for col in range(summary_sheet.ncols)]
				total_values.append(col_vals)
				# Flip flag
				read_col_vals = True
			row = [summary_sheet.cell(1, col).value for col in range(summary_sheet.ncols)]
			total_values.append(row)
			logging.debug('Row successfully copied.')
# As long as there is data to write, write something
if total_values:
	logging.info("Writing data to: " + output_file)
	# Create a new workbook
	new_workbook = xlwt.Workbook()
	# Create a new sheet
	new_sheet = new_workbook.add_sheet('Sheet1')
	# For each row
	for row in range(0, len(total_values)):
		# Copy the row
		current_row = total_values[row]
		logging.info("Writting row " + str(row))
		# For each col
		for col in range(0, len(current_row)):
			# Write the data to the sheet
			new_sheet.write(row, col, label=current_row[col])
	logging.info("Writing complete, saving data..")
	new_workbook.save(output_file)
	logging.info("File saved")
else:
	logging.debug("No data was found")
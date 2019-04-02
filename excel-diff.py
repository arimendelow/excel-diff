import pandas as pd
from pathlib import Path
import sys # for argv


def excel_diff(path_OLD, path_NEW, index_col_OLD, index_col_NEW):

	# use pandas DataFrames for the comparison - read the files
	df_OLD = pd.read_excel(path_OLD, index_col=index_col_OLD).fillna(0)
	df_NEW = pd.read_excel(path_NEW, index_col=index_col_NEW).fillna(0)

	# create a new DataFrame for the diff and loop through the originals to identify changes
	df_diff = df_NEW.copy()
	dropped_rows = []
	new_rows = []
	dropped_cols = []
	new_cols = []
	mod_vals = 0

	cols_OLD = df_OLD.columns
	cols_NEW = df_NEW.columns
	sharedCols = list(set(cols_OLD).intersection(cols_NEW))
	
	for row in df_diff.index:
		# if the row is in both tables
		if (row in df_OLD.index) and (row in df_NEW.index):
			# go through the stuff that's in both sheets, checking if it's been changed
			for col in sharedCols:
				value_OLD = df_OLD.loc[row,col]
				value_NEW = df_NEW.loc[row,col]
				# if the value is unchanged, then:
				if value_OLD==value_NEW:
					df_diff.loc[row,col] = df_NEW.loc[row,col]
				# otherwise, if the value has been changed:
				else:
					mod_vals += 1
					df_diff.loc[row,col] = ('{} → {}').format(value_OLD,value_NEW)
		else:
			new_rows.append(row)

	for row in df_diff.index:
		# for values in columns that have been deleted:
		# (this is needed because otherwise these values would never be added to the output sheet)
		for col in df_OLD.columns:
			# if the ROW is in the new sheet
			if row in df_OLD.index:
				# but the COLUMN isnt
				if col not in df_NEW.columns:
					df_diff.loc[row,col] = df_OLD.loc[row,col]

	for row in df_OLD.index:
		# if the row is NOT in the new table
		if row not in df_NEW.index:
			dropped_rows.append(row)
			df_diff = df_diff.append(df_OLD.loc[row,:])

	for col in df_OLD.columns:
		# if the col is NOT in the new table
		if col not in df_NEW.columns:
			dropped_cols.append(col)
	for col in df_NEW.columns:
		if col not in df_OLD.columns:
			new_cols.append(col)

	df_diff = df_diff.sort_index().fillna('')

	# add in the information about new/dropped rows/cols
	# creating a new dataframe that will be added as a third sheet, called "results"
	# DataFrame.append(other, ignore_index=False, verify_integrity=False, sort=None)
	df_results = pd.DataFrame({"RESULTS (see worksheet DIFF for more information)":[
											'{} modified values; {} new rows; {} dropped rows; {} new columns; {} dropped columns:'
											.format(mod_vals, len(new_rows), len(dropped_rows), len(new_cols), len(dropped_cols)),
											"New rows:                {}".format(new_rows),
											"Dropped rows:        {}".format(dropped_rows),
											"New columns:         {}".format(new_cols),
											"Dropped columns: {}".format(dropped_cols)
										]})

	print(df_diff)
	print('\n{} modified values; {} new rows; {} dropped rows; {} new columns; {} dropped columns:'
			.format(mod_vals, len(new_rows), len(dropped_rows), len(new_cols), len(dropped_cols)))
	print('\nNew Rows:     	 {}'.format(new_rows))
	print('Dropped Rows: 	 {}'.format(dropped_rows))
	print('New Columns:     {}'.format(new_cols))
	print('Dropped Columns: {}'.format(dropped_cols))

	# Save output and format
	fname = '{} vs {}.xlsx'.format(path_OLD.stem,path_NEW.stem)
	writer = pd.ExcelWriter(fname, engine='xlsxwriter')

	df_results.to_excel(writer, sheet_name='SUMMARY', index=True)
	df_diff.to_excel(writer, sheet_name='DIFF', index=True)
	df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
	df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)

	# get xlsxwriter objects
	workbook  = writer.book
	worksheet = writer.sheets['DIFF']
	# worksheet.hide_gridlines(2)
	worksheet.set_default_row(15)

	# define formats
	date_fmt = workbook.add_format({'align': 'center', 'num_format': 'yyyy-mm-dd'})
	center_fmt = workbook.add_format({'align': 'center'})
	number_fmt = workbook.add_format({'align': 'center', 'num_format': '#,##0.00'})
	cur_fmt = workbook.add_format({'align': 'center', 'num_format': '$#,##0.00'})
	perc_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
	rm_fmt = workbook.add_format({'font_color': '#FF0000', 'bold':True, 'align': 'center'})
	rm_fmt_header = workbook.add_format({'font_color': '#FF0000', 'bold':True, 'align': 'center', 'border':1})
	changed_fmt = workbook.add_format({'font_color': '#FFFF00', 'bold':True, 'bg_color':'#B1B3B3', 'align': 'center'})
	new_fmt = workbook.add_format({'font_color': '#14a810','bold':True, 'align': 'center'})
	new_fmt_header = workbook.add_format({'font_color': '#14a810','bold':True, 'align': 'center', 'border':1})
	orig_fmt_header = workbook.add_format({'font_color': '#000000', 'bold':True, 'align': 'center', 'border':1})

	# Indicate added/removed columns:
	# Write the column headers, highlighting changed cells
	# (needed to do column[1] to actually get the column name)
	# write(row, col, *args (string or cell_format))
	col_num = 1
	for column in enumerate(df_diff.columns):
		if column[1] not in df_OLD.columns:
			worksheet.write(0, col_num, column[1], new_fmt_header)
		elif column[1] not in df_NEW.columns:
			worksheet.write(0, col_num, column[1], rm_fmt_header)
		else:
			worksheet.write(0, col_num, column[1], orig_fmt_header)
		col_num += 1

	# set format over range
	## highlight changed cells
	worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
											'criteria': 'containing',
											'value':'→',
											'format': changed_fmt})
	
	# highlight new/changed rows
	# set_row(row, height, cell_format)
	row_index = 1
	for row in df_diff.index:
		if row in new_rows:
			# format row content
			worksheet.set_row(row_index, 15, new_fmt)
			# format row header
			worksheet.write(row_index, 0, row, new_fmt_header)
		if row in dropped_rows:
			worksheet.set_row(row_index, 15, rm_fmt)
			worksheet.write(row_index, 0, row, rm_fmt_header)
		row_index += 1

	# highlight new/changed cols
	# set_column(first_col, last_col, width, cell_format)
	# need to start col_index at 1 becuase 0 is the index
	col_index = 1
	for col in df_diff.columns:
		if col not in df_OLD.columns:
			worksheet.set_column(col_index, col_index, 15, new_fmt)
		if col not in df_NEW.columns:
			worksheet.set_column(col_index, col_index, 15, rm_fmt)
		col_index += 1

	# set approx column widths:
	for i, width in enumerate(get_col_widths(df_diff)):
		worksheet.set_column(i, i, width)

	# finally, insert in the bottom of the table a report of how many changes exist between the two:

	
	# save
	writer.save()
	print('\nDone! Check "{} vs {}.xlsx" for the result.\n'.format(path_OLD.stem,path_NEW.stem))

def get_col_widths(dataframe):
	# First we find the maximum length of the index column   
	idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
	# Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
	return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def print_cols(df, path):
	# prints out the column names of an excel document and asks which one you want to use for indexing

	# print out the column names with numbers for user selection
	i = 0
	print(f"\nColumns in {path}:")
	for column in df.columns:
		print(f"col {i}: ", end = "")
		print(column)
		i += 1
		
	return int(input(f"\nWhich column do you want to use for indexing {path}? "))
	

def main():
	print("\nWelcome to excel-diff V1.0.0!")
	print("Written by Ari Mendelow, Copyright © 2019")
	print("See https://github.com/arimendelow/excel-diff for more information.\n")
	print("Note that in this version, mapping is done using column names.")
	print("Therefore, column names with the same data must be identical.")
	print("For now, you can manually ensure that column names are the same.")
	print("The ability to map one column name to another is currently in development.")
	# make sure that the command was written along with the two file names
	if not (len(sys.argv) == 3):
		print("Usage: excel-diff.exe old_file.xlsx new_file.xlsx")
		exit(1)
	else:
		path_OLD = Path(sys.argv[1])
		path_NEW = Path(sys.argv[2])

		# get index col from data
		df_OLD = pd.read_excel(path_OLD)
		df_NEW = pd.read_excel(path_NEW)
		
		print("\nBy default, I use the first column in both spreadsheets for indexing.")
		print("Note that this means that I'll pull the information in this column to")
		print("match the rows, so the content must be the same,")
		print("though the column titles can be different.")

		print("Do you want to continue with this default behavior?")
		opt = input("Choose NO if you'd rather pick the index column in each spreadsheet.\n(YES or NO): ")
		while opt not in ["Y", "y", "YES", "yes", "N", "n", "NO", "no"]:
			opt = input("You must choose YES or NO\n")
		if opt in ["N", "n", "NO", "no"]:
			index_col_OLD = print_cols(df_OLD, path_OLD)
			index_col_NEW = print_cols(df_NEW, path_NEW)
		else:
			index_col_OLD = 0
			index_col_NEW = 0
		print('\nIndex column in OLD spreadsheet: {}'.format(df_OLD.columns[index_col_OLD]))
		print('\nIndex column in NEW spreadsheet: {}\n'.format(df_NEW.columns[index_col_NEW]))

		excel_diff(path_OLD, path_NEW, index_col_OLD, index_col_NEW)


if __name__ == '__main__':
	main()
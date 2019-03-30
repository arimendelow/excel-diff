import pandas as pd
from pathlib import Path
import sys # for argv


def excel_diff(path_OLD, path_NEW, index_col_OLD, index_col_NEW):

    # use pandas DataFrames for the comparison - read the files
    df_OLD = pd.read_excel(path_OLD, index_col=index_col_OLD).fillna(0)
    df_NEW = pd.read_excel(path_NEW, index_col=index_col_NEW).fillna(0)

    # create a new DataFrame for the diff and loop through the originals to identify changes
    dfDiff = df_NEW.copy()
    droppedRows = []
    newRows = []

    cols_OLD = df_OLD.columns
    cols_NEW = df_NEW.columns
    sharedCols = list(set(cols_OLD).intersection(cols_NEW))
    
    for row in dfDiff.index:
        if (row in df_OLD.index) and (row in df_NEW.index):
            for col in sharedCols:
                value_OLD = df_OLD.loc[row,col]
                value_NEW = df_NEW.loc[row,col]
                if value_OLD==value_NEW:
                    dfDiff.loc[row,col] = df_NEW.loc[row,col]
                else:
                    dfDiff.loc[row,col] = ('{}→{}').format(value_OLD,value_NEW)
        else:
            newRows.append(row)

    for row in df_OLD.index:
        if row not in df_NEW.index:
            droppedRows.append(row)
            dfDiff = dfDiff.append(df_OLD.loc[row,:])

    dfDiff = dfDiff.sort_index().fillna('')
    print(dfDiff)
    print('\nNew Rows:     {}'.format(newRows))
    print('Dropped Rows: {}'.format(droppedRows))

    # Save output and format
    fname = '{} vs {}.xlsx'.format(path_OLD.stem,path_NEW.stem)
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='DIFF', index=True)
    df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
    df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)

    # get xlsxwriter objects
    workbook  = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)
    worksheet.set_default_row(15)

    # define formats
    date_fmt = workbook.add_format({'align': 'center', 'num_format': 'yyyy-mm-dd'})
    center_fmt = workbook.add_format({'align': 'center'})
    number_fmt = workbook.add_format({'align': 'center', 'num_format': '#,##0.00'})
    cur_fmt = workbook.add_format({'align': 'center', 'num_format': '$#,##0.00'})
    perc_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
    grey_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    new_fmt = workbook.add_format({'font_color': '#32CD32','bold':True})

    # set format over range
    ## highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'→',
                                            'format': highlight_fmt})

    # highlight new/changed rows
    for row in range(dfDiff.shape[0]):
        if row+1 in newRows:
            worksheet.set_row(row+1, 15, new_fmt)
        if row+1 in droppedRows:
            worksheet.set_row(row+1, 15, grey_fmt)

    # save
    writer.save()
    print('\nDone.\n')

def print_cols(df, path):
    # prints out the column names of an excel document and asks which one you want to use for indexing

    # print out the column names with numbers for user selection
    i = 0
    print(f"Columns in {path}:")
    for column in df.columns:
        print(f"col {i}: ", end = "")
        print(column)
        i += 1
        
    return int(input(f"\nWhich column do you want to use for indexing {path}? "))
    

def main():

    # make sure that the command was wrong along with the two file names
    if not (len(sys.argv) == 3):
        print("Usage: python excel-diff.py old_file.xlsx new_file.xlsx")
        exit(1)
    else:
        path_OLD = Path(sys.argv[1])
        path_NEW = Path(sys.argv[2])

        # get index col from data
        df_OLD = pd.read_excel(path_OLD)
        df_NEW = pd.read_excel(path_NEW)

        index_col_OLD = print_cols(df_OLD, path_OLD)
        index_col_NEW = print_cols(df_NEW, path_NEW)


        # # print out the column names with numbers for user selection
        # i = 0
        # print(f"Columns in {path_NEW}:")
        # for column in dfNew.columns:
        #     print(f"col {i}: ", end = "")
        #     print(column)
        #     i += 1
        
        # index_col = int(input("\nWhich column do you want to use for indexing? "))
        print('\nIndex column in OLD spreadsheet: {}\n'.format(df_OLD.columns[index_col_OLD]))
        print('\nIndex column in NEW spreadsheet: {}\n'.format(df_NEW.columns[index_col_NEW]))

        excel_diff(path_OLD, path_NEW, index_col_OLD, index_col_NEW)


if __name__ == '__main__':
    main()
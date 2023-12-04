# Excel utilities - read me

## Objective

The objective of this module is to create a number of utilities that make it easy to write excel files

## Initializing

This code initializes a number of styles that are then available as global parameters. In order to usethis, the main code must import the module:   import excel_utilities as exc
or put this code in sub-folder utils, then
sys.path.append('..')                              this enables the code to go and fetch something in another branch
from utils import excel_utilities as exc           this imports

First thing to do is to initialize global variables which may be used in calling routines
Call this up in the main routine with
if __name__ == "__main__":
exc.set_excel_styles()

## Writing cells

Purpose

* The write_cells function writes a contiguous block of cells in a spreadsheet

Coding

* write_cells(wb_sheet,first_col,first_row, last_col,last_row,font_style,border_style,alignment_style,lov)

input

* wb_sheet: the name of the sheet that is currently open in the Excel document
* first_col: the first column from which you write the data: use column 1 to signal Excel column A(and not column 0 in Python)
* first_row: the first row from which you write the data: use row 1 to signal Excel row 1 (and not row 0 in Python)
* last_col: the first column to which you write the data: starting to count from column 1
* last_row: the first row to which you write the data: starting to count from row 1
  the entire matrix size is therefore (last_col - first_col + 1) *  (last_row - first_row+ 1)
* font_style: global parameter for appearance of text in cell
* border_style:  global parameter for appearance of border
* alignment_style:  global parameter for alignment of text in cell
* lov: a list of values to be written, which must have a length equal to the matrix size
  data must be presented as row by row  (row 1, col 1 col 2 col 3, ...., row 2, col 1 col 2 col  3)

action:

output:

* wb_sheet: the name of the sheet that is currently open in the Excel document

usage example

* wb_sheet : ex.write_cells(wb_sheet, 1, 2, 2, 4, cell_bold_style, cell_border_style, cell_align_left_style,'Paul',25,'Jim', 12, 'Mary', 41, 'Anne', 22])

## Writing a new Excel or CSV file with 1 sheet

Purpose

* The write_excel_file function writes 1 sheet in a new xslx or csv file

Coding

* write_excel_file (df_in, file_prod = False,  path_in = "", file_name_in = "", file_ext_in = ["xlsx"] ,
  width_adjust = False, sheet_name_in = 'Sheet1', overwrite = True)

input

* df_in: the data_frame to write
* file_prod: do you want to create a file or not ?  False by default, which means the function is not doing anything unless set to True
* path_in: the location you which to write the file to (where the root is the location where you are calling this function from)
* file_name_in : the name of the file - not including the extension
* file_ext_in: either xslx or csv
* width_adjust: True if you want column widths to be adjusted to the length of the data in them, False otherwise. Default is False. This parameter only works for xlsx files
* sheet_name_in: what you want the excel sheet to be called. This parameter only works for xlsx files
* overwrite:  whether you want any existing file with same location/name to be overwritten. Default is True (note: be careful as this is not well implemented... overwrite should be kept True at all times)

action:

* The xslx or csv file is saved

output:

* none

## Opening a new or existing Excel (xslx) file

Purpose

* The open_excel_wb function creates a new excel file, or else it opens  an existing excel file with path/name provided, and returns workbook handle and file_name

Coding

* open_excel_wb(df_in, path_in = "", file_root_name_in = "")

input

* df_in: the data_frame to write
* path_in: the location you which to write the file to (where the root is the location where you are calling this function from)
* file_root_name_in : the name of the file, not including its extension. xlsx is assumed

output:

* wb: workbook handle
* file_name: the full file_name including its location path and extension

## Obtain the workbook handle of an existing Excel (xslx) file

Purpose

* The return_excel_wb function returns the handle of an existing excel file.

Coding

* return_excel_wb(path_in, file_root_name_in)

input

* path_in: the location you which to write the file to (where the root is the location where you are calling this function from)
* file_root_name_in : the name of the file, not including its extension. xlsx is assumed

output:

* wb: workbook handle
* the function fails if the file does not exist

## Write 1 excel sheet in an existing Excel (xslx) file, if you have its workbook handle

Purpose

* The write_excel_sheet_in_wb function writes 1 sheet in an existing xslx file, if you provide it with its workbook handle

Coding

* write_excel_sheet_in_wb(df_in, wb_in, sheet_index_in = 0, sheet_name_in = "no_sheet", width_adjust = False, overwrite = False)

input

* df_in: the data_frame to write
* wb_in: the workbook handle (if not valid the function fails)
* sheet_index_in: the position of the sheet in the Excel file. Default is 0 (first sheet)
* sheet_name_in: what you want the excel sheet to be called
* width_adjust: True if you want column widths to be adjusted to the length of the data in them, False otherwise. Default is False
* overwrite:  True if you want the sheet overwritten if it already exists, False otherwise. Default is False

Action:

* the xlsx file is NOT saved

output:

* none

## Write 1 excel sheet in an existing or new Excel (xslx) file

Purpose

* The write_excel_sheet function writes 1 sheet in an existing xslx file. This is the most used function as it combines several actions and is fairly failsafe, as it ensures the file gets created if it does not yet exist 

Coding

* write_excel_sheet (df_in, file_prod = False,  path_in = "", file_root_name_in = "", sheet_index_in = 0, sheet_name_in = "no_sheet", width_adjust = False, overwrite = False)

input

* df_in: the data_frame to write
* file_prod: do you want to create a file or not ?  False by default, which means the function is not doing anything unless set to True
* path_in: the location you which to write the file to (where the root is the location where you are calling this function from)
* file_root_name_in : the name of the file - not including the extension, as xlsx is assumed
* sheet_index_in: the position of the sheet in the Excel file. Default is 0 (first sheet)
* sheet_name_in: what you want the excel sheet to be called. Default is "no sheet"
* width_adjust: True if you want column widths to be adjusted to the length of the data in them, False otherwise. Default is False.
* overwrite:  True if you want the sheet overwritten if it already exists, False otherwise. Default is False

Action:

* this function combines the following 3 actions:
  * Open a new or existing Excel (xslx) file 
        open_excel_wb(df_in, path_in = "", file_root_name_in = "")
  * Write 1 excel sheet in an existing Excel (xslx) file
        write_excel_sheet (df_in, file_prod = False,  path_in = "", file_root_name_in = "", sheet_index_in = 0, sheet_name_in = "no_sheet", width_adjust = False, overwrite = False)
  * Save the excel file 
        wb.save(file_name)


output:

* none

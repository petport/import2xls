# import2xls
Easy to use command line based script to import and export data - to and from - .xlsx files. The script is implemented to fullfil some specific use cases.
- To extract specific data from an xlsx file (specified by the --master_file option) and put it in a new file under csv format. New file's name is specified 
by the --output_file option.
- To select specific data from a file (--input_file) and place them in the desired position in another file (--master_file)
- Because the files needed for our use cases are .xls and .xlsx, a --convert_xlx2xlsx option is provided in order to first convert a .xls file to .xlsx format, and 
then use it as needed as use cases 1 and 2 suggest. 

Implemented and tested only in Python 3.8.8

Help page: 

└──╼ $python xls_handler.py --help
Usage: main.py [OPTIONS]

Options:
  --master_file TEXT          Master file. File to either select data to write
                              to file, or import data from another file.

  --output_file TEXT          Where to save the selected data.
  --input_file TEXT           File from which to import data.
  --master_ranges TEXT        One or more ranges. Use _ in between. Example:
                              "A1:B6_D1:E6".

  --input_file_ranges TEXT    Data to import from existing file. Use _ in
                              between. Example: "A1:B6_D1:E6".

  --convert_xls2xlsx TEXT...  Convert .xlsx file in .xls format.
  --help                      Show this message and exit.
  
Basic Usage:
Options --master_file, --input_file and --convert_xls2xlsx expect path to the desired file, or just the file name if you are working at the desired directory.

Common use case examples:

python xls_handler.py --master_file dummy_xlsx.xlsx --output_file output.csv --master_ranges A1:B15
The example above, selects the data located in the range A1 to B15 of file dummy_xlsx.xlsx and exports it in an .csv file.

python main.py --master_file dummy_xlsx.xlsx --master_ranges A17:B19 --input_file test_input_file.xlsx --input_file_ranges F12:G14
The example above specifies the data that needs to be imported from an input file (from cell F12 to G14), and copies it to the specified cells of the master_file.

python main.py --convert_xls2xlsx dummy_xls.xls new_xlsx.xlsx
The example above converts an .xls file to a .xlsx file. 

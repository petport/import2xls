# import2xls
Easy to use command line based script to import and export data - to and from - xlsx files. The script is implemented to fullfil some specific use cases.
- To extract specific data from an xlsx file, specified by the --master_file option, and put them in a new file under csv format. New file's name is specified 
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
Options --master_file, --input_file and --convert_xls2xlsx expect path to the desired file, or just the file name of you are at the desired directory. Same thing
applies for xls_handler.py


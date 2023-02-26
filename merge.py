
import csv

# the openpyxl library is used to read excel files
from openpyxl import load_workbook, Workbook

def merge_files_into_csv(file_list, output_file, lines_to_skip_in_each_file):

    with open(output_file, 'w', newline='') as outfile:                       
        writer = csv.writer(outfile)

        res = merge_files(file_list, writer.writerow, lines_to_skip_in_each_file)

    return res

def merge_files_into_xlsx(file_list, output_file, lines_to_skip_in_each_file):

    workbook  = Workbook()
    sheet = workbook.active

    res = merge_files(file_list, sheet.append, lines_to_skip_in_each_file)
    
    workbook.save(output_file)
    
    return res
                                          

def merge_files(file_list, write_function, lines_to_skip):
    row_counter = 0
    for file in file_list:                                    
            workbook = load_workbook(filename=file)               
            sheet = workbook.worksheets[0]                        
            row_counter += \
                process_sheet(sheet, write_function, lines_to_skip)
            workbook.close()   
    
    return row_counter

def process_sheet(sheet, write_function, lines_to_skip):
    row_counter = 0
    for row_values in sheet.iter_rows(min_row = lines_to_skip + 1, values_only=True):      
        if(row_values[0] == None):
            continue
        write_function(row_values)
        row_counter += 1
        
    return row_counter


def usage():
     print("Usage: python merge_files.py <lines_to_skip> <output_file> <input_file_1> <input_file_2> ...")

if __name__ == "__main__":
    import sys
    
    if(len(sys.argv) < 4):
        usage()
        exit(1)
       
    try:
        lines = int(sys.argv[1])
    except:
        usage()
        exit(1)
        
    out = sys.argv[2]
    in_ = sys.argv[3:]
    
    try:
        if out.endswith(".csv"):
            res = merge_files_into_csv(in_, out, lines)
        elif out.endswith(".xlsx"):
            res = merge_files_into_xlsx(in_, out, lines)
        else:
            res =merge_files_into_csv(in_, out + ".csv", lines)
            
        print(f"Successfully saved {res} lines" )
    except Exception as e:
        print(f"Failed: " + str(e) )
    
    
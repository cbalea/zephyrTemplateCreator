import re
import sys

import xlrd
import xlwt



def write_one_test(test, sheet, start_row):
    write_data_to_excel_sheet(sheet, start_row, 0, test["test_name"])
    write_data_to_excel_sheet(sheet, start_row, 1, test["description"])
    
    for step in xrange(len(test["steps"])):
        write_data_to_excel_sheet(sheet, start_row + step, 2, test["steps"][step])
        write_data_to_excel_sheet(sheet, start_row + step, 3, test["results"][step])
        write_data_to_excel_sheet(sheet, start_row + step, 4, test["test_data"][step])
    end_row = start_row + step
    
    write_data_to_excel_sheet(sheet, start_row, 5, test["priority"])
    write_data_to_excel_sheet(sheet, start_row, 6, test["components"])
    write_data_to_excel_sheet(sheet, start_row, 7, test["story_id"])
    
#     print empty row
#     end_row += 1
#     write_data_to_excel_sheet(sheet, end_row, 0, "")
    
    return end_row+1


def write_sheet_header(sheet):
    write_data_to_excel_sheet(sheet, 0, 0, "Test case name")
    write_data_to_excel_sheet(sheet, 0, 1, "Description")
    write_data_to_excel_sheet(sheet, 0, 2, "Steps")
    write_data_to_excel_sheet(sheet, 0, 3, "Results")
    write_data_to_excel_sheet(sheet, 0, 4, "Test data")
    write_data_to_excel_sheet(sheet, 0, 5, "Priority")
    write_data_to_excel_sheet(sheet, 0, 6, "Component")
    write_data_to_excel_sheet(sheet, 0, 7, "Story ID")


def write_data_to_excel_sheet(sheet, row, column, data):
    sheet.write(row, column, data)

def is_empty_row(data):
    is_empty_row = True
    for index, value in enumerate(data):
        if value != "":
            is_empty_row = False
    return is_empty_row

def strip_list(orig_list=[]):
    new_list = []
    for element in orig_list:
        if element != "":
            new_list.append(element.strip())
    return new_list

def convert_to_import_template(row_content, story_id, row_nb, component=None):
    test_name_column = 5
    description_column = 5
    steps_column = 7
    result_column = 6
    priority_column = 4
    test_data_column = 0
    
    steps = strip_list(re.split("\d+\.", row_content[steps_column]))
    
    if len(steps) == 0:
        raise Exception("Test case <%s> contains no EXPECTED RESULT." %str(row_content[test_name_column]).strip())
    elif row_content[4] == "":
        raise Exception("Test case <%s> has no defined PRIORITY." %str(row_content[test_name_column]).strip())
    
    results = []
    test_data = []
    for i in xrange(len(steps)):
        results.append("")
        test_data.append("")
    test_data[0] = row_content[test_data_column]
    results[-1] = row_content[result_column]
    
    try:
        return {"test_name":str(row_content[test_name_column]).strip(), 
          "description":row_content[description_column].strip(), 
          "steps":steps, 
          "results":results,
          "test_data":test_data, 
          "priority":row_content[priority_column].strip(), 
          "components":component.strip(), 
          "story_id":story_id.strip()}
    except Exception as e:
        exception_title = "Row <%d> rasied exception: \n" %(row_nb+1) 
        raise Exception(exception_title + str(e))



def read_input_file(input_file):
    input_workbook = xlrd.open_workbook(input_file)
    sheet = input_workbook.sheet_by_index(0)

    start_row = 1
    story_id_column = 18
    
    rows_for_import_template = []
    for row_nb in xrange(start_row, sheet.nrows):
        row_content = [sheet.cell_value(row_nb, col) for col in range(sheet.ncols)]
        if row_content[1]:
            story_id = row_content[story_id_column]
        if not is_empty_row(row_content):
            converted_data = convert_to_import_template(row_content, story_id, row_nb, "OTSN")
            rows_for_import_template.append(converted_data)
    return rows_for_import_template


def write_destination_file(destination_file, all_rows):
    dest_file = xlwt.Workbook(encoding="utf-8")
    sheet = dest_file.add_sheet("Sheet1")
    write_sheet_header(sheet)
    start_row = 1
    for row in all_rows:
        end_row = write_one_test(row, sheet, start_row)
        start_row = end_row
    dest_file.save(destination_file)




input_file = sys.argv[1].lower()
destination_file = sys.argv[2].lower()


converted_rows = read_input_file(input_file)
print "Conversion completed successfully."
write_destination_file(destination_file, converted_rows)
print "Destination file written at: " + destination_file 
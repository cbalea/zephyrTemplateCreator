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

def strip_list(orig_list, element_to_remove=None):
    new_list = []
    for element in orig_list:
        if element != "":
            if element_to_remove:
                if element_to_remove not in element:
                    new_list.append(element.strip())
            else:
                new_list.append(element.strip())
    return new_list

def convert_to_import_template(row_content, story_id, row_nb, labels=None):
    test_name_column = 4
    description_column = 4
    steps_column = 6
    result_column = 7
    priority_column = 3
    test_data_column = 5
    deprecated_column = None
    
    res = re.split("\d+\.", row_content[result_column])
    row_content[result_column] = strip_list(res, "URL should be launched successfully")
    
    stps = re.split("\d+\.", row_content[steps_column])
    steps = strip_list(stps)
    
    if len(steps) == 0:
        raise Exception("Row <%d> contains no STEPS." %(row_nb+1))
    elif row_content[result_column] == "":
        raise Exception("Row <%d> contains no EXPECTED RESULT." %(row_nb+1))
    elif row_content[priority_column] == "":
        raise Exception("Row <%d> has no defined PRIORITY." %(row_nb+1))
    
    results = []
    test_data = []
    for i in xrange(len(steps)):
        results.append("")
        test_data.append("")
    if test_data_column:
        test_data[0] = row_content[test_data_column]
    results[-1] = row_content[result_column]
    if deprecated_column and row_content[deprecated_column]:
        labels += ", deprecated"
    
    try:
        return {"test_name":u''.join(row_content[test_name_column]).encode('utf-8').strip(), 
          "description":row_content[description_column].strip(), 
          "steps":steps, 
          "results":results,
          "test_data":test_data, 
          "priority":row_content[priority_column].strip(), 
          "components":labels, 
          "story_id":story_id.strip()}
    except Exception as e:
        exception_title = "Row <%d> rasied exception: \n" %(row_nb+1) 
        raise Exception(exception_title + str(e))



def read_input_file(input_file):
    test_cases_sheet = 0
    start_row = 1
    story_id_column = None
    test_name_column = 4
    labels_column = 2
    general_label = "cts_desktop"

    input_workbook = xlrd.open_workbook(input_file)
    sheet = input_workbook.sheet_by_index(test_cases_sheet)
    
    rows_for_import_template = []
    story_id = ""
    for row_nb in xrange(start_row, sheet.nrows):
        row_content = [sheet.cell_value(row_nb, col) for col in range(sheet.ncols)]
        if story_id_column and row_content[story_id_column] != "":
            story_id = row_content[story_id_column]
        else:
            pass # when we have multiple tests/story and the story id is only printed once for the entire batch
#             story_id = "" # when we have tests without stories
        if labels_column:
            label = row_content[labels_column].lower().replace(" ", "_") + ", " + general_label
        else:
            label = general_label
        if not is_empty_row(row_content):
            converted_data = convert_to_import_template(row_content, story_id, row_nb, label)
            rows_for_import_template.append(converted_data)
        if '\n' in row_content[test_name_column]:
            row_content[test_name_column] = row_content[test_name_column].replace('\n', " ")
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
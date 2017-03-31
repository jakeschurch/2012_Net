"""Create .NET files from Book of States data.

Program utilies Book of States excel files to iteratively create relationships
between nodes, edges, and weights. NET files
for every State
d
Last updated: 3/24/2017
"""
__author__ = 'Jake Schurch'
__version__ = '1.1'

import os
import xlrd


def append_headers_2_list(excel_sheet):
    """Append header values from excel spreadsheet to a list.

    If the column in the excel spreadsheet is not 0, the
    values in the column are appended to a list. If the column is 0,
    column is passed.

    Args:   excel_sheet: Excel spreadsheet that is iterated over

    Returns:    header_list
    """
    header_list = []
    for col in range(excel_sheet.ncols):
        if col != 0:
            header_list.append(excel_sheet.cell_value(0, col))
        else:
            pass
    return header_list


def append_vals_2_list(excel_sheet):
    """Function that appends values from an excel spreadsheet to a list.

       If the column in the excel spreadsheet is not 0, the values in
       the column are appended to a list. If the column is 0,
       column is passed.

    Args:
        excel_sheet: Excel spreadsheet that is iterated over

    Returns:    List containing values from excel_sheet where column not 0
    """
    val_list = []
    for col in range(excel_sheet.ncols):
        if col != 0:
            val_list.append(excel_sheet.cell_value(row_count, col))
        else:
            pass
    val_list = encode_vals_2_utf8(val_list)
    return val_list


def append_vals_2_list_comp(excel_sheet):
    """Function that appends values from an excel spreadsheet to a list.

       If the column in the excel spreadsheet is not 0, the values in
       the column are appended to a list. If the column is 0,
       column is passed.

    Args:   excel_sheet: Excel spreadsheet that is iterated over

    Returns:    List of values that are not in column 0
    """
    val_list = [excel_sheet.cell_value(row_count, col)
                for col in range(excel_sheet.ncols) if col != 0]
    val_list = encode_vals_2_utf8(val_list)
    return val_list


def encode_vals_2_utf8(input_list):
    """Encode values from a list to utf-8 encoding.

    Args:   input_list: list containing values

    Returns:    output_list
    """
    assert (isinstance(input_list, list)), 'arg: input_list is not type: list!'
    # output_list = list(map(lambda x: x.encode('utf-8'), input_list))
    output_list = []
    for node in input_list:
        if isinstance(node, float):
            output_list.append(node)
        else:
            output_list.append(node.encode('utf-8'))
    return output_list


wb = xlrd.open_workbook('2012 Appointment Power.xlsx')   # Create link to WB

appt_sheet1 = wb.sheet_by_name('2012 Appt Party (1)')
appt_sheet2 = wb.sheet_by_name('2012 Appt Party (2)')
appt_sheet3 = wb.sheet_by_name('2012 Appt Party (3)')
appt_sheet4 = wb.sheet_by_name('2012 Appt Party (4)')
aprvl_sheet1 = wb.sheet_by_name('2012 Aprvl Party (1)')
aprvl_sheet2 = wb.sheet_by_name('2012 Aprvl Party (2)')

appt_weight1 = wb.sheet_by_name('2012 Appt Party Weight (1)')
appt_weight2 = wb.sheet_by_name('2012 Appt Party Weight (2)')
appt_weight3 = wb.sheet_by_name('2012 Appt Party Weight (3)')
appt_weight4 = wb.sheet_by_name('2012 Appt Party Weight (4)')
aprvl_weight1 = wb.sheet_by_name('2012 Aprvl Party Weight (1)')
aprvl_weight2 = wb.sheet_by_name('2012 Aprvl Party Weight (2)')

row_count = 1
while row_count < 51:

    encoded_headernodes = append_headers_2_list(appt_sheet1)
    encoded_appt1_rownodes = append_vals_2_list(appt_sheet1)
    encoded_appt2_rownodes = append_vals_2_list(appt_sheet2)
    encoded_appt3_rownodes = append_vals_2_list(appt_sheet3)
    encoded_appt4_rownodes = append_vals_2_list(appt_sheet4)
    encoded_aprvl1_rownodes = append_vals_2_list(aprvl_sheet1)
    encoded_aprvl2_rownodes = append_vals_2_list(aprvl_sheet2)

    nodelist = (encoded_appt1_rownodes + encoded_appt2_rownodes +
                encoded_appt3_rownodes + encoded_aprvl1_rownodes +
                encoded_aprvl2_rownodes + encoded_headernodes)

    nodelist = list(set(nodelist))  # Remove all duplicates from list
    nodelist.sort()  # sorts list alphabhetically

    # CREATE NUMERIC INDICATOR FOR EACH VERTICE
    numberlist = [x for x in range(len(nodelist) + 1) if x != 0]

    sorted_vertices = dict(zip(numberlist, nodelist))

    # CREATE APPOINTMENT NODE RELATIONSHIPS
    alt_vertices = {v: k for k, v in sorted_vertices.items()}
    # using alt_vertices as a reference to each node

    # CREATE LIST OF SOURCENODES
    source_appt1_nodes = [alt_vertices[x] for x in encoded_appt1_rownodes]
    source_appt2_nodes = [alt_vertices[x] for x in encoded_appt2_rownodes]
    source_appt3_nodes = [alt_vertices[x] for x in encoded_appt3_rownodes]
    source_appt4_nodes = [alt_vertices[x] for x in encoded_appt3_rownodes]
    source_aprvl1_nodes = [alt_vertices[x] for x in encoded_aprvl1_rownodes]
    source_aprvl2_nodes = [alt_vertices[x] for x in encoded_aprvl2_rownodes]

    # CREATE LIST OF TARGETNODES
    target_nodes = [alt_vertices[x] for x in encoded_headernodes]

    # create list of encoded weights from weight spreadsheets
    new_appt1_weights = append_vals_2_list_comp(appt_weight1)
    new_appt2_weights = append_vals_2_list_comp(appt_weight2)
    new_appt3_weights = append_vals_2_list_comp(appt_weight3)
    new_appt4_weights = append_vals_2_list_comp(appt_weight4)
    new_aprvl1_weights = append_vals_2_list_comp(aprvl_weight1)
    new_aprvl2_weights = append_vals_2_list_comp(aprvl_weight2)

    node_not_wanted = int((numberlist[-1] - 1) + 1)  # Gets rid of None values
    number_of_vertices = str(numberlist[-1] - 1)

    state_name = appt_sheet1.cell_value(row_count, 0)
    # through every iteration, value will be the State we are appending data

    temp_node_file = 'Temporary {} 2012 Map.net'.format(state_name)
    node_file = '{} 2012 Map.net'.format(state_name)

    # MERGE LISTS TO CREATE APPOINTMENT TUPLE RELATIONSHIPS #
    appt1_edges = zip(source_appt1_nodes, target_nodes, new_appt1_weights)
    appt2_edges = zip(source_appt2_nodes, target_nodes, new_appt2_weights)
    appt3_edges = zip(source_appt3_nodes, target_nodes, new_appt3_weights)
    appt4_edges = zip(source_appt4_nodes, target_nodes, new_appt4_weights)
    aprvl1_edges = zip(source_aprvl1_nodes, target_nodes, new_aprvl1_weights)
    aprvl2_edges = zip(source_aprvl2_nodes, target_nodes, new_aprvl2_weights)

    appt2_edges = [(x, y, z) for (x, y, z) in appt2_edges
                   if y != (node_not_wanted) and b'none' and z != b'none']
    appt1_edges = [(x, y, z) for (x, y, z) in appt1_edges
                   if x != (node_not_wanted) and y != b'none' and z != b'none']
    appt3_edges = [(x, y, z) for (x, y, z) in appt3_edges
                   if x != (node_not_wanted) and y != b'none' and z != b'none']
    appt4_edges = [(x, y, z) for (x, y, z) in appt3_edges
                   if x != (node_not_wanted) and y != b'none' and z != b'none']
    aprvl1_edges = [(x, y, z) for (x, y, z) in aprvl1_edges
                    if x != (node_not_wanted) and y != b'none' and z != b'none']
    aprvl2_edges = [(x, y, z) for (x, y, z) in aprvl2_edges
                    if x != (node_not_wanted) and y != b'none' and z != b'none']

    sorted_vertices = {key: value for key, value
                       in sorted_vertices.items()
                       if key != 'none' and value != 'none'}

    with open(temp_node_file, 'a') as f:
        f.write('*Vertices {} \n'.format(number_of_vertices))
        f.writelines('{0} "{1}" \n'.format(k, v)
                     for k, v in sorted_vertices.items())
        f.write('*Arcs :1 Appointment \n')
        f.writelines('{} \n'.format(x) for x in appt1_edges)
        f.writelines('{} \n'.format(x) for x in appt2_edges)
        f.writelines('{} \n'.format(x) for x in appt3_edges)
        f.writelines('{} \n'.format(x) for x in appt4_edges)
        f.writelines('{} \n'.format(x) for x in aprvl1_edges)
        f.writelines('{} \n'.format(x) for x in aprvl2_edges)
    f.close()

    input_file = open(temp_node_file, 'r')
    output_file = open(node_file, 'w')
    for line in input_file:
        line = line.replace('(', '')
        line = line.replace(')', '')
        line = line.replace(',', '')
        output_file.write(line)

    output_file.close()
    input_file.close()
    os.remove(temp_node_file)
    row_count += 1

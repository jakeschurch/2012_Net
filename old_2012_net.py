'''
Script created by Jacob Schurch
Last updated: August 18, 2016
Purpose of this program is to relationships between
source and target nodes iteratively
'''
import os
import xlrd
import re

wb = xlrd.open_workbook('2012 Appointment Power.xlsx')   #Create link to workbook

# For following five lines - create variables that reference party values
appt_sheet1 = wb.sheet_by_name("2012 Appt Party (1)")
appt_sheet2 = wb.sheet_by_name("2012 Appt Party (2)")
appt_sheet3 = wb.sheet_by_name("2012 Appt Party (3)")
appt_sheet4 = wb.sheet_by_name("2012 Appt Party (4)")
aprvl_sheet1 = wb.sheet_by_name("2012 Aprvl Party (1)")
aprvl_sheet2 = wb.sheet_by_name("2012 Aprvl Party (2)")

row_count = 1
while row_count < 51:

    ### CREATE VERTICES LIST ###
    headernodes = [] #Create list where header values will be appended to
    appt1_rownodes = [] #Create list where row values will be appended to
    appt2_rownodes = []
    appt3_rownodes = []
    appt4_rownodes = []
    aprvl1_rownodes = []
    aprvl2_rownodes = []

    #For following six lines - append values to corresponding list

    for col in range(appt_sheet1.ncols):
        if col == 0:
            pass
        else:
            headernodes.append(appt_sheet1.cell_value(0,col))

    encoded_headernodes = []

    for node in headernodes:
        encoded_headernodes.append(node.encode("utf8")) # encodes nodes in utf-8



    for col in range(appt_sheet1.ncols):
        if col == 0:
            pass
        else:
            appt1_rownodes.append(appt_sheet1.cell_value(row_count, col))

    for col in range(appt_sheet2.ncols):
        if col == 0:
            pass
        else:
            appt2_rownodes.append(appt_sheet2.cell_value(row_count, col))

    for col in range(appt_sheet3.ncols):
        if col == 0:
            pass
        else:
            appt3_rownodes.append(appt_sheet3.cell_value(row_count, col))

    for col in range(appt_sheet4.ncols):
        if col == 0:
            pass
        else:
            appt4_rownodes.append(appt_sheet4.cell_value(row_count, col))

    for col in range(aprvl_sheet1.ncols):
        if col == 0:
            pass
        else:
            aprvl1_rownodes.append(aprvl_sheet1.cell_value(row_count, col))

    for col in range(aprvl_sheet2.ncols):
        if col == 0:
            pass
        else:
            aprvl2_rownodes.append(aprvl_sheet2.cell_value(row_count,col))

    encoded_appt1_rownodes = []
    encoded_appt2_rownodes = []
    encoded_appt3_rownodes = []
    encoded_appt4_rownodes = []
    encoded_aprvl1_rownodes = []
    encoded_aprvl2_rownodes = []

    for node in appt1_rownodes:
        encoded_appt1_rownodes.append(node.encode("utf-8")) # changes encoding to utf-8

    for node in appt2_rownodes:
        encoded_appt2_rownodes.append(node.encode("utf-8"))

    for node in appt3_rownodes:
        encoded_appt3_rownodes.append(node.encode("utf-8"))

    for node in appt4_rownodes:
        encoded_appt4_rownodes.append(node.encode("utf-8"))

    for node in aprvl1_rownodes:
        encoded_aprvl1_rownodes.append(node.encode("utf-8"))

    for node in aprvl2_rownodes:
        encoded_aprvl2_rownodes.append(node.encode("utf-8"))

    nodelist = encoded_appt1_rownodes + encoded_appt2_rownodes + encoded_appt3_rownodes + encoded_aprvl1_rownodes + encoded_aprvl2_rownodes + encoded_headernodes #Create reference list of all possible nodes

    #CREATE NEW LIST TO STORE ENCODED NODES IN

    nodelist = list(set(nodelist)) #Remove all duplicates from list

    nodelist.sort() #sorts list alphabhetically

    ### CREATE NUMERIC INDICATOR FOR EACH VERTICE ###

    numberlist = []

    i = 1
    while i <= len(nodelist):
        numberlist.append(i)
        i = i ++ 1

    sorted_vertices = dict(zip(numberlist, nodelist))




    ### CREATE APPOINTMENT NODE RELATIONSHIPS ###

    alt_vertices = dict(zip(nodelist,numberlist)) #using alt_vertices as a reference to each node

    # CREATE LIST OF SOURCENODES #

    source_appt1_nodes = [alt_vertices[x] for x in encoded_appt1_rownodes]
    source_appt2_nodes = [alt_vertices[x] for x in encoded_appt2_rownodes]
    source_appt3_nodes = [alt_vertices[x] for x in encoded_appt3_rownodes]
    source_appt4_nodes = [alt_vertices[x] for x in encoded_appt3_rownodes]
    source_aprvl1_nodes = [alt_vertices[x] for x in encoded_aprvl1_rownodes]
    source_aprvl2_nodes = [alt_vertices[x] for x in encoded_aprvl2_rownodes]

    # CREATE LIST OF TARGETNODES #

    target_nodes = [alt_vertices[x] for x in encoded_headernodes]

    # CREATE LIST OF WEIGHTED NODES #

    appt_weight1 = wb.sheet_by_name('2012 Appt Party Weight (1)')
    appt_weight2 = wb.sheet_by_name('2012 Appt Party Weight (2)')
    appt_weight3 = wb.sheet_by_name('2012 Appt Party Weight (3)')
    appt_weight4 = wb.sheet_by_name('2012 Appt Party Weight (4)')
    aprvl_weight1 = wb.sheet_by_name('2012 Aprvl Party Weight (1)')
    aprvl_weight2 = wb.sheet_by_name('2012 Aprvl Party Weight (2)')

    appt1_node_weights = []
    appt2_node_weights = []
    appt3_node_weights = []
    appt4_node_weights = []
    aprvl1_node_weights = []
    aprvl2_node_weights = []

    for col in range(appt_weight1.ncols):
        if col == 0:
            pass
        else:
            appt1_node_weights.append(appt_weight1.cell_value(row_count,col))

    for col in range(appt_weight2.ncols):
        if col == 0:
            pass
        else:
            appt2_node_weights.append(appt_weight2.cell_value(row_count,col))

    for col in range(appt_weight3.ncols):
        if col == 0:
            pass
        else:
            appt3_node_weights.append(appt_weight3.cell_value(row_count,col))

    for col in range(appt_weight4.ncols):
        if col == 0:
            pass
        else:
            appt4_node_weights.append(appt_weight4.cell_value(row_count,col))

    for col in range(aprvl_weight1.ncols):
        if col == 0:
            pass
        else:
            aprvl1_node_weights.append(aprvl_weight1.cell_value(row_count,col))

    for col in range(aprvl_weight2.ncols):
        if col == 0:
            pass
        else:
            aprvl2_node_weights.append(aprvl_weight2.cell_value(row_count,col))

    new_appt1_weights = []
    new_appt2_weights = []
    new_appt3_weights = []
    new_appt4_weights = []
    new_aprvl1_weights = []
    new_aprvl2_weights = []

    for node in appt1_node_weights:
        if isinstance(node, float):
            new_appt1_weights.append(node)
        else:
            new_appt1_weights.append(node.encode("utf-8"))

    for node in appt2_node_weights:
        if isinstance(node, float):
            new_appt2_weights.append(node)
        else:
            new_appt2_weights.append(node.encode("utf-8"))

    for node in appt3_node_weights:
        if isinstance(node, float):
            new_appt3_weights.append(node)
        else:
            new_appt3_weights.append(node.encode("utf-8"))

    for node in appt4_node_weights:
        if isinstance(node, float):
            new_appt4_weights.append(node)
        else:
            new_appt4_weights.append(node.encode("utf-8"))

    for node in aprvl1_node_weights:
        if isinstance(node, float):
            new_aprvl1_weights.append(node)
        else:
            new_aprvl1_weights.append(node.encode("utf-8"))

    for node in aprvl2_node_weights:
        if isinstance(node, float):
            new_aprvl2_weights.append(node)
        else:
            new_aprvl2_weights.append(node.encode("utf-8"))

     ### WRITE NUMBERS OF VERTICES AND NODE VALUES TO FILE ###
    node_not_wanted = int((numberlist[-1]-1) + 1)
    number = str(numberlist[-1]-1) #using string instead of integer becasue cannot concatenate str and int objects
    state_name = appt_sheet1.cell_value(row_count,0) # through every iteration, value will be the State we are appending data
    temp_node_file = 'Temporary ' + state_name + ' 2012 Map.net'
    node_file = state_name + ' 2012 Map.net'

    # MERGE LISTS TO CREATE APPOINTMENT TUPLE RELATIONSHIPS #

    appt1_edges = zip(source_appt1_nodes,target_nodes,new_appt1_weights)
    appt2_edges = zip(source_appt2_nodes,target_nodes,new_appt2_weights)
    appt3_edges = zip(source_appt3_nodes,target_nodes,new_appt3_weights)
    appt4_edges = zip(source_appt4_nodes,target_nodes,new_appt4_weights)
    aprvl1_edges = zip(source_aprvl1_nodes,target_nodes,new_aprvl1_weights)
    aprvl2_edges = zip(source_aprvl2_nodes,target_nodes,new_aprvl2_weights)

    appt1_edges = [(x,y,z) for (x,y,z) in appt1_edges if x != (node_not_wanted) and y != "none" and z != "none"]
    appt2_edges = [(x,y,z) for (x,y,z) in appt2_edges if y != (node_not_wanted) and "none" and z != "none"]
    appt3_edges = [(x,y,z) for (x,y,z) in appt3_edges if x != (node_not_wanted) and y != "none" and z != "none"]
    appt4_edges = [(x,y,z) for (x,y,z) in appt3_edges if x != (node_not_wanted) and y != "none" and z != "none"]
    aprvl1_edges = [(x,y,z) for (x,y,z) in aprvl1_edges if x != (node_not_wanted) and y != "none" and z != "none"]
    aprvl2_edges = [(x,y,z) for (x,y,z) in aprvl2_edges if x != (node_not_wanted) and y != "none" and z != "none"]

    sorted_vertices = {key:value for key, value in sorted_vertices.items() if value != "none"}

    with open(temp_node_file, 'a') as f:
        f.write('*Vertices' + ' %s \n' % (number))
        f.writelines('{} "{}" \n'.format(k, v) for k, v in sorted_vertices.items())
    f.close()
    # WRITE APPOINTMENT TUPLE RELATIONSHIPS TO TXT FILE #

    with open(temp_node_file, 'a') as f:
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
        line = line.replace('(','')
        line = line.replace(')','')
        line = line.replace(',','')
        output_file.write(line)
    output_file.close()
    input_file.close()

    os.remove(temp_node_file)

    row_count += 1

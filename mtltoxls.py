import re
import glob
import sys
import os
from  openpyxl import Workbook
from pprint import pprint

re_key = re.compile(r'(\S*)\s*=\s*\{\s*Name\s*=\s*(\S*)\s*\n')
re_parameter = re.compile(r'(\{(?:\s*.* = .*\s*\n)*\})')
re_key_value_parameter = re.compile(r'\s*(?P<key>\S*) = (?P<value>\S*)( .*)?\s*\n')


def parse_mtl(fname):
    with open(fname, 'r', encoding='utf8') as fd:
        mtl = fd.read()

    # Find all parameter sections in curly brackets like for example:
    #    {
    #        Name = PTC_INITIAL_BEND_Y_FACTOR
    #        Type = Real
    #        Default = 5.000000e-01
    #        Access = Full
    #    },
    m = re.findall(re_parameter, mtl)

    # For each parameter set, extract the list of Key-Value pairs like "Name = PTC_INITIAL..."
    parameters = []
    for parameter in m:
        matches = re.findall(re_key_value_parameter, parameter)
        if matches:
            parameters.append(matches)

    # Extract all parameter sets into a list of dictionaries
    # Adding a new dict key: 'Unit' which comes from the 'Default' value field.
    parameter_dicts = []
    for parameter in parameters:
        pd = {'Unit': None}
        for key, value, unit in parameter:
            pd.update({key: value})
            if unit is not '':
                pd.update({'Unit': unit.strip()})
        parameter_dicts.append(pd)

    # Convert the string values of all of the parameters 'Default' field into a proper python
    # data type - so that it can later be stored appropriately in the spreadsheet
    for pd in parameter_dicts:
        value = pd['Default']
        pdtype = pd['Type']
        if pdtype == 'Real':
            value = float(value)
        elif pdtype == 'Integer':
            value = int(value)
        elif pdtype == 'String':
            value = str(value.strip("\'"))
        pd['Default'] = value

    # make a dictionary of parameters where the parameter Name is the key
    material_params = {}
    for pd in parameter_dicts:
        material_params.update({pd['Name']: pd})

    # TODO: remote debug print
    pprint(material_params)

    # Scan file content again to extract the material ID and Name.
    material_key = re.findall(re_key, mtl[1:])
    if material_key:
        key, material = material_key[0]
    else:
        key = 'unknown_key'
        material = 'unknown_material'

    # Generate final top-level dictionary view of the entire file including material Name and ID.
    result = {key: {material: material_params}}
    return result


def store_in_spreadsheet(fname, material):
    material_params = {}

    # Create a index dictionary to look up parameter name -> spreadsheet row
    parameter_row_start_index = 5
    param_index_row = parameter_row_start_index
    parameters_index_rows = {}
    for param in material_params.keys():
        parameters_index_rows.update({param: param_index_row})
        param_index_row += 1

    # Create a Data point to spreadsheet column lookup
    data_columns = {
        'Name': 'A',
        'Default': 'B',
        'Unit': 'C',
        'Type': 'D',
        'Access': 'E'
    }

    wb = Workbook()
    ws = wb.active

    # Create headers
    ws['A1'] = 'ID:'
    #ws['B1'] = key
    ws['A2'] = 'Material'
    ws['B2'] = material
    ws.merge_cells('B1:E1')
    ws.merge_cells('B2:E2')

    for col in data_columns:
        ws["{}{}".format(data_columns[col], parameter_row_start_index-1)] = col
    ws["A{}".format(parameter_row_start_index-1)] = 'Parameter'

    # Fill in parameters
    for pd_key in material_params:
        pd = material_params[pd_key]
        row = parameters_index_rows[pd['Name']]
        for field in pd:
            cell_name = "{}{}".format(data_columns[field], row)
            ws[cell_name] = pd[field]

    wb.save('materials.xlsx')


def main():
    path = sys.argv[1]
    full_path = os.path.abspath(path)
    print(F"Searching dir for mtl files: \"{full_path}\"")
    fnames = glob.glob(F"{full_path}/*.mtl")
    pprint(fnames)
    materials = []
    for fname in fnames:
        materials.append(parse_mtl(fname))

    pprint(materials)


if __name__=="__main__":
    main()

import re
from  openpyxl import Workbook
from pprint import pprint
# {
#  Name = PTC_MATERIAL_DESCRIPTION
#  Type = String
#  Default = 'BRONZE'
#  Access = Full
#},
#{
#  Name = TEMPERATURE
#  Type = Real
#  Default = 0.000000e+00 F
#  Access = Full
#},

re_key = re.compile(r'(\S*)\s*=\s*\{\s*Name\s*=\s*(\S*)\s*\n')
re_parameter = re.compile(r'(\{(?:\s*.* = .*\s*\n)*\})')
re_key_value_parameter = re.compile(r'\s*(?P<key>\S*) = (?P<value>\S*)( .*)?\s*\n')

with open('data/bronze.mtl', 'r', encoding='utf8') as fd:
    mtl = fd.read()
    
m = re.findall(re_parameter, mtl)

parameters = []
for parameter in m:
    matches = re.findall(re_key_value_parameter, parameter)
    if matches:
        parameters.append(matches)

parameter_dicts = []
for parameter in parameters:
    pd = {'Unit': None}
    for key, value, unit in parameter:
        pd.update({key: value})
        if unit is not '':
            pd.update({'Unit': unit.strip()})
    parameter_dicts.append(pd)
    
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
    
material_params = {}
for pd in parameter_dicts:
    material_params.update({pd['Name']: pd})     
pprint(material_params)

material_key = re.findall(re_key, mtl[1:])
if material_key:
    key, material = material_key[0]
else:
    key = 'unknown_key'
    material = 'unknown_material'
result = {key: {material: material_params}}

#pprint(result)

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
ws['B1'] = key
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

from openpyxl import load_workbook
import tempfile
from openpyxl.formula.tokenizer import Token
from datetime import datetime
from time import sleep
import sys

redemption_model = load_workbook('ARM Template_Simplified.xlsx', data_only=True)
inputs_sheet = redemption_model['Inputs']
calculations_sheet = redemption_model['Calculations']
cell_value = calculations_sheet['M9'].value
print(cell_value)





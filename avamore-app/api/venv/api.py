from flask import Flask, jsonify, request
import os
import xlwings
from datetime import datetime
from openpyxl import load_workbook

app = Flask(__name__)

file_path = os.path.join(os.path.dirname(__file__), 'ARM Template_Simplified.xlsx')

def load_excel():
    if os.path.exists(file_path):
        workbook = load_workbook(file_path)
        return workbook
    else:
        return None

def save_and_close(workbook):
    if workbook:
        try:
            workbook.save(file_path)
            workbook.close()
            excel_app = xlwings.App(visible=False)
            excel_book = excel_app.books.open(file_path)
            excel_book.save()
            excel_book.close()
            excel_app.quit()
        except Exception as e:
            print("Error saving and closing workbook:", e)
        finally:
            # Delete temporary files
            temp_files = [f for f in os.listdir(os.path.dirname(file_path)) if f[-2:]=="00"]
            for temp_file in temp_files:
                temp_file_path = os.path.join(os.path.dirname(file_path), temp_file)
                try:
                    os.remove(temp_file_path)
                except Exception as e:
                    print("Error deleting temporary file:", e)


@app.route('/values', methods=['GET', 'POST'])
def get_values():
    if request.method == 'GET':
        workbook = load_workbook(file_path, data_only=True)
        if workbook:
            inputs_sheet = workbook['Inputs']
            facility_A = inputs_sheet['C15'].value
            interest_rate = inputs_sheet['C22'].value * 100
            default_start = inputs_sheet['C28'].value
            default_end = inputs_sheet['C29'].value
            default_start_formatted = default_start.strftime('%Y-%m-%d')
            default_end_formatted = default_end.strftime('%Y-%m-%d')
            
            data_calculations_sheet = workbook['Calculations']
            interest_due = round(data_calculations_sheet['B3'].value,2)
            
            workbook.close()
            
            return jsonify({'facility_A': facility_A,
                            'interest_rate': interest_rate,
                            'default_start': default_start_formatted,
                            'default_end': default_end_formatted,
                            'interest_due': interest_due})
        else:
            return jsonify({'error': 'The Excel file was not found'}), 404
    elif request.method == 'POST':
        data = request.json
        workbook = load_excel()
        if workbook:
            calculations_sheet = workbook['Calculations']      
            inputs_sheet = workbook['Inputs']
            inputs_sheet['C15'] = float(data['facility_A'])
            inputs_sheet['C22'] = float(data['interest_rate']) / 100
            inputs_sheet['C28'] = datetime.strptime(data['default_start'], '%Y-%m-%d')
            inputs_sheet['C29'] = datetime.strptime(data['default_end'], '%Y-%m-%d')
            save_and_close(workbook)
            
            # Reopen the workbook in data-only mode to retrieve the calculated interest_due
            workbook = load_workbook(file_path, data_only=True)
            calculations_sheet = workbook['Calculations']
            interest_due = round(calculations_sheet['B3'].value,2)
            
            workbook.close()
            
            return jsonify({'facility_A': data['facility_A'],
                            'interest_rate': data['interest_rate'],
                            'default_start': data['default_start'],
                            'default_end': data['default_end'],
                            'interest_due': interest_due})
        else:
            return jsonify({'error': 'Excel file not found'}), 404

if __name__ == '__main__':
    app.run()

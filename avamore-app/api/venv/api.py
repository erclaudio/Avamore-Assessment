from flask import Flask, jsonify, request
import time
from openpyxl import load_workbook
import os

app = Flask(__name__)

def load_excel():
    file_path = os.path.join(os.path.dirname(__file__), 'ARM Template_Simplified.xlsx')
    if os.path.exists(file_path):
        workbook = load_workbook(file_path)
        return workbook
    else:
        return None

def save_and_close(workbook):
    
    workbook.save('ARM Template_Simplified.xlsx')
    workbook.close()


@app.route('/default-period', methods = ['GET'])
def default_period_value():
    workbook = load_excel()
    if workbook: 
        inputs_worksheet = workbook['Inputs']
        default_start = inputs_worksheet['C28'].value
        default_end = inputs_worksheet['C29'].value
        default_start_formatted = default_start.strftime('%Y-%m-%d')
        default_end_formatted = default_end.strftime('%Y-%m-%d')
        return jsonify({'default_start':default_start_formatted,'default_end':default_end_formatted})
    else: 
        return jsonify({'error':'Excel file not found'}), 404



@app.route('/interest-rate', methods = ['GET', 'POST'])
def interest_rate_value():
    if request.method == 'GET':
        workbook = load_excel()
        if workbook:
            inputs_worksheet = workbook['Inputs']
            interest_rate = inputs_worksheet['C22'].value * 100
            save_and_close(workbook)
            return jsonify({'interest_rate': interest_rate})
        else:
            return jsonify({'error': 'Excel file not found'}), 404
    elif request.method == 'POST':
        data = request.json
        new_interest_rate = data.get('interest_rate')
        if new_interest_rate is not None:
            workbook = load_excel()
            if workbook:
                update_interest_rate(workbook, new_interest_rate)
                print("newrate", new_interest_rate)
                return jsonify({'message': 'Interest rate updated successfully'})
            else:
                return jsonify({'error': 'Excel file not found'}), 404
        else:
            return jsonify({'error': 'Invalid data provided'}), 400

def update_interest_rate(workbook, new_rate):
    inputs_worksheet = workbook['Inputs']
    inputs_worksheet['C22'] = new_rate / 100    
    save_and_close(workbook)

@app.route('/interest-due', methods=['GET'])
def interest_due_value():
    file_path = os.path.join(os.path.dirname(__file__), 'ARM Template_Simplified.xlsx')
    if os.path.exists(file_path):
        workbook = load_workbook(file_path, data_only=True)
    if workbook:
        data_calculations_sheet = workbook['Calculations']
        total_interest = round(data_calculations_sheet['B3'].value,2)
        save_and_close(workbook)
        return jsonify({'interest_due': total_interest})  # Return JSON response
    else:
        return jsonify({'error': 'Excel file not found'}), 404


@app.route('/facility', methods=['GET', 'POST'])
def handle_facility():
    if request.method == 'GET':
        workbook = load_excel()
        if workbook:
            inputs_sheet = workbook['Inputs']
            facility_value = inputs_sheet['C15'].value
            save_and_close(workbook)
            return jsonify({'facility_value': facility_value})
        else:
            return jsonify({'error': 'Excel file not found'}), 404
    elif request.method == 'POST':
        data = request.json
        workbook = load_excel()
        if workbook:
            facility_A_value(workbook, data['facility_value'])
            return jsonify({'message': 'Facility A value changed successfully'})
        else:
            return jsonify({'error': 'Excel file not found'}), 404

def facility_A_value(workbook, value):
    inputs_sheet = workbook['Inputs']
    inputs_sheet['C15'] = float(value)
    save_and_close(workbook)

@app.route('/values', methods =['GET'])
def get_values():
    workbook = load_excel()
    if workbook:
        inputs_sheet = workbook['Inputs']
        facility_A= inputs_sheet['C15'].value
        interest_rate= inputs_sheet['C22'].value
        default_start= inputs_sheet['C28'].value
        default_end= inputs_sheet['C29'].value
        save_and_close(workbook)
        return jsonify({'facility_A' : facility_A, 
                        'interest_rate' : interest_rate, 
                        'default_start' : default_start,
                        'default_end' : default_end})
    else:
        return jsonify({'error the excel file was not found'}), 404

@app.route('/time')
def get_time():
    return {'time': time.time()}

if __name__ == '__main__':
    app.run()

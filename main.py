from flask import Flask, render_template, request
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/save-to-excel', methods=['POST'])
def save_to_excel():
    name = request.form.get('name')
    height=request.form.get('height')
    weight = request.form.get('weight')
    medical_history = request.form.get('medical-history')
    stress_level = request.form.get('stress-level')

    try:
        workbook = openpyxl.load_workbook('data.xlsx')
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['Name', 'Height (in cm)','Weight (in kg)','Medical History','Stress Level'])

    sheet.append([name,height,weight,medical_history,stress_level])

    workbook.save('data.xlsx')

    return 'Data saved to Excel'

if __name__ == '__main__':
    app.run(port=5000,debug=True)

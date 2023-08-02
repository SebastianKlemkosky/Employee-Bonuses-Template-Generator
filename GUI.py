import PySimpleGUI as sg
import datetime

def validate_date_format(date_str):
    try:
        # Attempt to parse the date string in the format 'YYYY-MM-DD'
        datetime.datetime.strptime(date_str, '%Y-%m-%d')
        return True
    except ValueError:
        return False

def is_numeric(value):
    try:
        float(value)
        return True
    except ValueError:
        return False

def get_user_input():
    sg.theme('DefaultNoMoreNagging')

    aip_types = [
    'Standard - A',
    'Standard - B',
    'Standard - C',
    'Standard Plus - X',
    'Standard Plus - Y',
    'Standard Plus - Z',
    'MD - X',
    'MD - Y',
    'Global LT',
    'Standard Sales - Fizzy Drink',
    'Sales Plus - Fizzy Drink',
    'Standard Sales - Carbonated Drink',
    'Sales Plus - Carbonated Drink',
    'Standard Sales - NAFS',
    'Sales Plus - NAFS',
    'Standard Sales - LatAm',
    'Sales Plus - LatAm',
    'Standard Sales - Beverage 1',
    'Sales Plus - Beverage 1',
    'Standard Sales - Beverage 2',
    'Sales Plus - Beverage 2',
    'Standard Sales - Beverage 3',
    'Sales Plus - Beverage 3',
    'Standard Sales - Beverage 4',
    'Sales Plus - Beverage 4',
    'Standard Sales - Beverage 5',
    'Sales Plus - Beverage 5',
    'Standard Sales - Beverage 6',
    'Sales Plus - Beverage 6',
    'Standard Sales - Beverage 7',
    'Sales Plus - Beverage 7',
    'Standard Sales - Beverage 8']

    currency_codes = [
    'USD',
    'EUR',
    'GBP',
    'JPY',
    'AUD',
    'CAD',
    'CNY',
    'INR',
    'NZD',
    'MXN']

    layout = [
        [sg.Text("Employee Name:"), sg.Input(key='-EMPLOYEE_NAME-')],
        [sg.Text("Business Entity:"), sg.Input(key='-BUSINESS_ENTITY-')],
        [sg.Text("Job Title:"), sg.Input(key='-JOB_TITLE-')],
        [sg.Text("Department Name:"), sg.Input(key='-DEPARTMENT-')],
        [sg.Text("Budget Area:"), sg.Input(key='-BUDGET_AREA-')],
        [sg.Text("PIP Tier:"), sg.Input(key='-AIP_TIER-')],
        [sg.Text("PIP Type:"), sg.Combo(values=aip_types, key='-AIP_TYPE-', default_value='Standard - A')],
        [sg.Text("Currency Code:"), sg.Combo(values=currency_codes, key='-CURRENCY-', default_value='USD')],
        [sg.Text("Annual Salary:"), sg.Input(key='-ANNUAL_SALARY-')],
        [sg.Text("Hire Date (YYYY-MM-DD):"), sg.Input(key='-HIRE_DATE-')],
        [sg.Text("Minimum Bonus:"), sg.Input(key='-MINIMUM_BONUS-')],
        [sg.Text("CP Bonus:"), sg.Input(key='-CP_Bonus-')],
        [sg.Button("Submit"), sg.Button("Cancel")]
    ]

    window = sg.Window("Employee Information", layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Cancel':
            window.close()
            return None
        elif event == 'Submit':
            employee_name = values['-EMPLOYEE_NAME-']
            business_entity = values['-BUSINESS_ENTITY-']
            job_title = values['-JOB_TITLE-']
            department_name = values['-DEPARTMENT-']
            budget_area = values['-BUDGET_AREA-']
            aip_tier = values['-AIP_TIER-']
            aip_type = values['-AIP_TYPE-']
            currency_code = values['-CURRENCY-']
            annual_salary = values['-ANNUAL_SALARY-']
            hire_date = values['-HIRE_DATE-']
            minimum_bonus = values['-MINIMUM_BONUS-']
            cp_bonus = values['-CP_Bonus-']

            # Check if any field is blank
            if not all([employee_name, business_entity, job_title, department_name, budget_area, aip_tier, aip_type,
                        currency_code, annual_salary, hire_date, minimum_bonus, cp_bonus]):
                sg.popup_error("All fields must be filled. Please ensure no field is left blank.")
                continue

            # Validate the hire date format
            if not validate_date_format(hire_date):
                sg.popup_error("Invalid hire date format. Please use 'YYYY-MM-DD'.")
                continue

            # Validate 'Annual Salary', 'Minimum Bonus', 'Eligible Salary' to ensure they are numeric
            if not is_numeric(annual_salary) or not is_numeric(minimum_bonus) or not is_numeric(cp_bonus):
                sg.popup_error(
                    "Invalid input for 'Annual Salary', 'Minimum Bonus', or 'CP Bonus'. Please enter numeric values.")
                continue

            window.close()
            return {
                'Employee Name': employee_name,
                'Business Entity': business_entity,
                'Job Title': job_title,
                'Department': department_name,
                'Budget Area': budget_area,
                'AIP Tier': aip_tier,
                'AIP Type': aip_type,
                'Currency Code': currency_code,
                'Annual Salary': annual_salary,
                'Hire Date': hire_date,
                'Minimum Bonus': minimum_bonus,
                'CenterPoint (CP)': cp_bonus
            }


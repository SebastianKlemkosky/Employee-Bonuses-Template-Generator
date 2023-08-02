import csv
import os
from docx2pdf import convert
import GUI
import populate_template
import pandas as pd
from datetime import datetime
import os
import math

TEMPLATE_PATH = r'documents\templates\Template.docx'
CSV_PATH = r'documents\data\employee data.csv'
OUTPUT_FOLDER = r'documents/output'
MATRIX_IMAGES_PATH = r'documents/matrix photos'
PLAN_PATH = r'documents\data\PIP Plan Ref.xlsx'

MATRIX_MAPPING = {
    'Generic - A Matrix': '1-M',
    'Generic - B Matrix': '1',
    'Generic - C Matrix': '2',
    'Generic - D Matrix': '3',
    'Generic - E Matrix': '4',
    'Generic - F Matrix': '5',
    'Carbonated Drink Matrix': '6',
    'Generic - A DAYS Matrix': '7',
    'Generic - B DAYS Matrix': '7-M',
    'Generic - C DAYS Matrix': '8',
    'Generic - D DAYS Matrix': '9',
    'Generic - E DAYS Matrix': '10',
    'Generic - F DAYS Matrix': '11',
    'Generic - G DAYS': '12',
    'Fizzy Drink Matrix': '13',
    'Generic - G Matrix': '14',
    'Generic - H Matrix': '15',
    'Generic - I Matrix': '16',
    'Generic - J Matrix': '17',
    'Generic - K Matrix': '18',
    'Generic - L Matrix': '19',
    'Generic - M Matrix': '20',
    'Generic - N Matrix': '21',
    'Generic - O Matrix': '22'
}

CURRENCY_SYMBOLS = {
    'USD': '$',  # United States Dollar
    'EUR': '€',  # Euro
    'GBP': '£',  # British Pound Sterling
    'JPY': '¥',  # Japanese Yen
    'AUD': 'A$',  # Australian Dollar
    'CAD': 'C$',  # Canadian Dollar
    'CNY': '¥',  # Chinese Yuan
    'INR': '₹',  # Indian Rupee
    'NZD': 'NZ$',  # New Zealand Dollar
    'MXN': 'Mex$',  # Mexican Peso
}


def get_dataframe(path, sheet_name=None):
    if sheet_name:
        df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    else:
        df = pd.read_excel(path, engine='openpyxl')
    return df


def calculate_eligible_salary(user_input):
    # Prorating logic to calculate the eligible salary based on hire date
    hire_date_str = user_input.get('Hire Date')
    hire_year, hire_month, hire_day = map(int, hire_date_str.split('-'))

    # Get the current date
    current_date = datetime.now()

    # Check if the year of hire_date is not the current year
    if hire_year != current_date.year:
        # If hire_date is not in the current year, return the annual_salary as prorated salary
        prorated_salary = user_input.get('Annual Salary', 0)
    else:
        # Calculate the percentage of the year remaining
        year_start = datetime(hire_year, 1, 1)
        year_end = datetime(hire_year, 12, 31)
        days_in_year = (year_end - year_start).days + 1
        days_remaining = (year_end - current_date).days + 1
        percentage_remaining = days_remaining / days_in_year

        # Convert 'Annual Salary' to float before performing the multiplication
        annual_salary = float(user_input.get('Annual Salary', 0))

        # Calculate the prorated salary based on the percentage remaining
        prorated_salary = annual_salary * percentage_remaining

    user_input['Eligible Salary'] = prorated_salary

    return user_input


def retrieve_plan_values(user_input, plan_df):
    # Get the Plan Row that matches the AIP Type
    matching_plan = plan_df[plan_df['Plan Classification'] == user_input['AIP Type']]

    # Check if there is a matching plan
    if not matching_plan.empty:
        user_input['T1 Name'] = str(matching_plan['Target 1'].iloc[0])
        user_input['T1 Multiplier'] = str(1)
        user_input['T1 Weight'] = str(matching_plan['Target 1 Weight'].iloc[0])

        # Convert NaN to None for T2, T3, and T4 values
        user_input['T2 Name'] = str(matching_plan['Target 2'].iloc[0]) if not pd.isna(
            matching_plan['Target 2'].iloc[0]) else ""
        user_input['T2 Multiplier'] = str(1)
        user_input['T2 Weight'] = str(matching_plan['Target 2 Weight'].iloc[0]) if not pd.isna(
            matching_plan['Target 2 Weight'].iloc[0]) else ""

        user_input['T3 Name'] = str(matching_plan['Target 3'].iloc[0]) if not pd.isna(
            matching_plan['Target 3'].iloc[0]) else ""
        user_input['T3 Multiplier'] = str(1) 
        user_input['T3 Weight'] = str(matching_plan['Target 3 Weight'].iloc[0]) if not pd.isna(
            matching_plan['Target 3 Weight'].iloc[0]) else ""

        user_input['T4 Name'] = str(matching_plan['Target 4'].iloc[0]) if not pd.isna(
            matching_plan['Target 4'].iloc[0]) else ""
        user_input['T4 Multiplier'] = str(1)
        user_input['T4 Weight'] = str(matching_plan['Target 4 Weight'].iloc[0]) if not pd.isna(
            matching_plan['Target 4 Weight'].iloc[0]) else ""

    return user_input


def calculate_target_payout(center_point, prorated_salary, target_multiplier, target_weight):
    # Convert parameters to floats if possible, otherwise return None
    try:
        center_point = float(center_point)
        prorated_salary = float(prorated_salary)
        if target_multiplier is not None:
            target_multiplier = float(target_multiplier)
        if target_weight is not None:
            target_weight = float(target_weight)
    except ValueError:
        return None

    # Check if any of the parameters are NaN
    if target_multiplier is None or math.isnan(target_multiplier) or \
            target_weight is None or math.isnan(target_weight):
        return None

    # Calculate the payout and return as a string
    payout = center_point * prorated_salary * target_multiplier * target_weight
    return str(payout)


def calculate_total_payout(target1_payout, target2_payout, target3_payout, target4_payout):
    # Function to convert payout string to float or return None
    def to_float_or_none(payout):
        try:
            return float(payout)
        except (ValueError, TypeError):
            return None

    # Convert payouts to floats or None
    payouts = [to_float_or_none(payout) for payout in [target1_payout, target2_payout, target3_payout, target4_payout]]

    # Calculate the total payout by summing only the valid payouts
    total_payout = sum(payout for payout in payouts if payout is not None)

    return str(total_payout)


def print_column_headers(csv_path):
    with open(csv_path, 'r') as csv_file:
        reader = csv.reader(csv_file)
        header_row = next(reader)  # Read the first row (header row)

    print("Column Headers:")
    print(", ".join(header_row))


def get_matrix_identifier(matrix_name):
    if matrix_name is None:
        return None

    return MATRIX_MAPPING.get(matrix_name)


def get_currency_symbol(code):
    if code is None:
        return None

    return CURRENCY_SYMBOLS.get(code)


def main():
    # Check if CSV_PATH exists
    if os.path.exists(CSV_PATH):
        os.remove(CSV_PATH)

    # Create the output folder if it doesn't exist
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    user_input = GUI.get_user_input()

    plan_df = get_dataframe(PLAN_PATH)
    user_input = calculate_eligible_salary(user_input)

    user_input = retrieve_plan_values(user_input, plan_df)

    user_input['T1 Payout'] = calculate_target_payout(user_input['CenterPoint (CP)'], user_input['Eligible Salary'],
                                                      user_input['T1 Multiplier'], user_input['T1 Weight'])
    user_input['T2 Payout'] = calculate_target_payout(user_input['CenterPoint (CP)'], user_input['Eligible Salary'],
                                                      user_input['T2 Multiplier'], user_input['T2 Weight'])
    user_input['T3 Payout'] = calculate_target_payout(user_input['CenterPoint (CP)'], user_input['Eligible Salary'],
                                                      user_input['T3 Multiplier'], user_input['T3 Weight'])
    user_input['T4 Payout'] = calculate_target_payout(user_input['CenterPoint (CP)'], user_input['Eligible Salary'],
                                                      user_input['T4 Multiplier'], user_input['T4 Weight'])

    user_input['Total Payout'] = calculate_total_payout(user_input['T1 Payout'], user_input['T2 Payout'],
                                                        user_input['T3 Payout'], user_input['T4 Payout'])

    user_input['T1 Matrix Number (Bottom)'] = get_matrix_identifier(user_input['T1 Name'])
    user_input['T2 Matrix Number (Bottom)'] = get_matrix_identifier(user_input['T2 Name'])
    user_input['T3 Matrix Number (Bottom)'] = get_matrix_identifier(user_input['T3 Name'])
    user_input['T4 Matrix Number (Bottom)'] = get_matrix_identifier(user_input['T4 Name'])

    user_input['Currency Symbol'] = get_currency_symbol(user_input['Currency Code'])

    # Get the column headers (keys of user_input dictionary)
    column_headers = user_input.keys()

    # Convert user_input to a CSV row
    csv_row = list(user_input.values())

    # Write the CSV header and row to the CSV file
    with open(CSV_PATH, 'a', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)

        # Write the column headers as the first row in the CSV file
        csv_writer.writerow(column_headers)

        # Write the data row to the CSV file
        csv_writer.writerow(csv_row)


def main2():
    # Create the output folder if it doesn't exist
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Read the CSV file
    with open(CSV_PATH, 'r') as csv_file:
        reader = csv.DictReader(csv_file)

        # Process rows
        for row in reader:
            print(row)
            # Populate the template for the current row
            populated_doc = populate_template.populate_template(TEMPLATE_PATH, MATRIX_IMAGES_PATH, row)

            # Get the department from the CSV row
            department = row['Department']

            # Create a folder for the department if it doesn't exist
            department_folder = os.path.join(OUTPUT_FOLDER, department)
            os.makedirs(department_folder, exist_ok=True)

            # Generate the output file path based on the employee name and department
            employee_name = row['Employee Name']
            docx_filename = f"{employee_name}.docx"
            docx_path = os.path.join(department_folder, docx_filename)

            # Save the populated document as DOCX
            populated_doc.save(docx_path)

            # Convert the DOCX to PDF
            pdf_filename = f"{employee_name}.pdf"
            pdf_path = os.path.join(department_folder, pdf_filename)
            convert(docx_path, pdf_path)

            # Remove the DOCX file
            os.remove(docx_path)

    # Remove any remaining DOCX files in the output folder
    for root, dirs, files in os.walk(OUTPUT_FOLDER):
        for file in files:
            if file.endswith(".docx"):
                os.remove(os.path.join(root, file))


if __name__ == '__main__':
    main()
    main2()

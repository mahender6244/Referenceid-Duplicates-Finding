import openpyxl
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import io

app = Flask(__name__)
uploaded_files = []  # Initialize the list to store uploaded filenames

def find_reference_id(df):
    # Search for "Referenceid" within the DataFrame columns
    #fine PRIORITYs
    for idx, column in enumerate(df.columns):
        if "Referenceid" in str(column):
            return idx
    return None  # "Referenceid" not found in the DataFrame

def process_reference_id(file):
    result = {}
    serial_no = 2

    # Read the uploaded Excel file as bytes and create a BytesIO object
    excel_data = io.BytesIO(file.read())

    # Open the Excel file using pandas
    df = pd.read_excel(excel_data, dtype=str)  # Ensure all columns are treated as strings

    # Find the column index containing "Referenceid"
    reference_id_column = find_reference_id(df)

    if reference_id_column is not None:
        for idx, row in df.iterrows():
            reference_id = str(row[reference_id_column]).strip()
            if reference_id:
                stripped_reference_id = reference_id.lstrip('0')
                if stripped_reference_id not in result:
                    result[stripped_reference_id] = {
                        'count': 1,
                        'locations': [f"{chr(65 + reference_id_column)}{idx + 2}"],
                        'duplicates': [reference_id],
                        'most_zeros': stripped_reference_id  # Initialize with the value itself
                    }
                else:
                    result[stripped_reference_id]['count'] += 1
                    result[stripped_reference_id]['locations'].append(f"{chr(65 + reference_id_column)}{idx + 2}")
                    result[stripped_reference_id]['duplicates'].append(reference_id)
                    # Update the reference ID with the most leading zeros
                    if len(stripped_reference_id) > len(result[stripped_reference_id]['most_zeros']):
                        result[stripped_reference_id]['most_zeros'] = stripped_reference_id
            serial_no += 1

    # Sort the result by decreasing count
    sorted_result = dict(sorted(result.items(), key=lambda item: item[1]['count'], reverse=True))

    return sorted_result

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            uploaded_file = file.filename  # Get the uploaded file name
            uploaded_files.append(uploaded_file)  # Add the filename to the list
            output = process_reference_id(file)
            return render_template('upload.html', uploaded_file=uploaded_file, output=output, uploaded_files=uploaded_files)

    return render_template('upload.html')

@app.route('/view/<filename>')
def view_output(filename):
    # Check if the selected filename is in the list of uploaded files
    if filename in uploaded_files:
        # Retrieve the output for the selected file
        selected_output = process_reference_id(filename)  # Implement a function to retrieve output by filename
        return render_template('view_output.html', selected_output=selected_output, selected_filename=filename)
    else:
        return "File not found"  # Handle the case where the file is not found

if __name__ == '__main__':
    app.run(debug=True)

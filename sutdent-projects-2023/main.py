from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from google.cloud import storage

# Initialize the GCP client
storage_client = storage.Client()

# Your GCP bucket name and Excel file name
BUCKET_NAME = 'vinit_storage_1'
EXCEL_FILE_NAME = 'All_Group_Project_Data.xlsx'
TEMP_FILE_NAME = '/tmp/All_Group_Project_Data.xlsx'

# Check if the Excel file exists in GCP bucket, and if not, create a new DataFrame
bucket = storage_client.get_bucket(BUCKET_NAME)
blob = bucket.blob(EXCEL_FILE_NAME)

if not blob.exists():
    df = pd.DataFrame(columns=['Group Number', 'Roll No', 'Name', 'Group Description'])
    df.to_excel(TEMP_FILE_NAME, index=False)
    blob.upload_from_filename(TEMP_FILE_NAME)

app = Flask(__name__)

# Function to download the Excel file from GCP bucket to a temporary file
def download_excel():
    blob.download_to_filename(TEMP_FILE_NAME)

# Function to upload the Excel file from the temporary file to GCP bucket
def upload_excel():
    blob.upload_from_filename(TEMP_FILE_NAME)

# Function to read the Excel file into a DataFrame
def read_excel():
    download_excel()
    return pd.read_excel(TEMP_FILE_NAME)

# Function to save the DataFrame to the Excel file
def save_excel(df):
    df.to_excel(TEMP_FILE_NAME, index=False)
    upload_excel()

# Define the route for the form page
@app.route('/')
def form():
    return render_template('home.html')

# Define the route for form submission
@app.route('/submit', methods=['POST'])
def submit():
    # Read the existing Excel file into a DataFrame
    existing_df = read_excel()

    group_number = request.form['groupNumber']
    group_description = request.form['projectDescription']

    students_data = {
        'Group Number': [],
        'Roll No': [],
        'Name': [],
        'Group Description': []
    }

    for i in range(1, 5):  # Loop through students 1 to 4
        roll_no = request.form.get(f'student{i}RollNo', '')
        name = request.form.get(f'student{i}Name', '')

        students_data['Group Number'].append(group_number if i == 1 else '')  # Include group number only once
        students_data['Roll No'].append(roll_no)
        students_data['Name'].append(name)
        students_data['Group Description'].append(group_description if i == 1 else '')  # Include group description only once

    # Create a DataFrame to hold the current form entries
    df = pd.DataFrame(students_data)

    # Concatenate the existing DataFrame with the current form entries
    combined_df = pd.concat([existing_df, df], ignore_index=True)

    # Save the combined DataFrame back to the Excel file with empty rows between groups
    save_excel(combined_df)

    return f'Data for Group {group_number} saved successfully!'

@app.route('/download')
def download():
    download_excel()
    return send_file(TEMP_FILE_NAME, attachment_filename=EXCEL_FILE_NAME, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

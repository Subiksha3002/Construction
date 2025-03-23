from flask import Flask, request, redirect
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = 'contact_details.xlsx'

# Create Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=['Name', 'Email', 'Message'])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/submit-message', methods=['POST'])
def submit_message():
    name = request.form.get('name')
    email = request.form.get('email')
    message = request.form.get('message')

    # Read existing Excel data
    df = pd.read_excel(EXCEL_FILE)

    # Create new entry as a DataFrame
    new_entry = pd.DataFrame([{'Name': name, 'Email': email, 'Message': message}])

    # Concatenate the new entry to existing DataFrame
    df = pd.concat([df, new_entry], ignore_index=True)

    # Save updated data back to Excel
    df.to_excel(EXCEL_FILE, index=False)

    return redirect('/')  # Redirect after submission

if __name__ == '__main__':
    app.run(debug=True)

import openpyxl
import pandas as pd
import pyodbc
import decimal
import time
import re
import os
import glob

### Make sure to change the month dates

# Switch to prod or dev
connection_string = (
    r'Driver={ODBC Driver 17 for SQL Server};' +
    ('SERVER={server};' +
     'DATABASE={database};' +
     'Trusted_Connection=yes;'
    ).format(
        #server='VFSSVRSQL1',
        # server ='vfs-1012',
        database='SamuelHaleRelationshipManager')
)

conn = pyodbc.connect(connection_string)
cursor = conn.cursor()

# Replace with the path to your input Excel file
xlsx_file_path = 'C:\\Users\\AThiru\\OneDrive - Accuire\\Documents\\Trust Monthly Hours\\Trust SSN method\\Trust May SSNs.xlsx'

# Replace with the path where you want to save the output Excel file
output_file_path = 'C:\\Users\\AThiru\\OneDrive - Accuire\\Documents\\Trust Monthly Hours\\Trust SSN method\\Trust May SSNs Outputs 3.xlsx'

# Read the Excel file into a pandas DataFrame and drop duplicates based on employee_id
df = pd.read_excel(xlsx_file_path)

df['SSN'] = df['SSN TRUST'].astype(str).str.zfill(9)


# Create an empty DataFrame to store the results
result_df = pd.DataFrame()
#columns=['SSN', 'First Name', 'Last Name', 'PEO Hire Date','Address 1' , 'Address 2' , 'City' , 'State', 'Zip' , 'Home Number', 'Cell Number',  'Email Address'])

i = 1

# Iterate through each row in the DataFrame and execute the SQL query
for index, row in df.iterrows():
    print(i)
    ssn = row['SSN']
    print(ssn)
    i=i+1
    print("\n")

    #first_name = row['first_name']
    #last_name = row['last_name']

    query = f"""
        SELECT 
        e.employee_id, e.social_security_number, e.last_name,
        e.first_name, ca.client_account_name,e.termination_date,
        SUM(prd.hours) hours
        FROM PayrolLRegister pr INNER JOIN PayrollRegisterDetail prd ON pr.payroll_register_id = prd.payroll_register_id 
        INNER JOIN PayCheckCodes pcc ON pcc.pay_check_code_id = prd.pay_check_code_id 
        INNER JOIN Employee e ON e.employee_id = pr.employee_id 
        inner join ClientAccount  ca on e.client_account_id = ca.client_account_id
        WHERE  pcc.code_type= 'Earnings' 
        AND(pcc.description LIKE '%REG%' OR pcc.description LIKE '%SAL%') 
        AND pcc.description NOT LIKE '%MEM%' AND pr.check_date BETWEEN '05/01/2024' AND '05/31/2024'
        AND replace(e.social_security_number,'-','') = '{ssn}'
        and (prd.hours) <> 0
        group by  e.employee_id, e.social_security_number, e.last_name,
        e.first_name, ca.client_account_name, e.termination_date
    """

    # Pass the emp_id as a parameter when executing the query
    result_emp_df = pd.read_sql(query, conn)

    print(result_emp_df)

    result_emp_df = result_emp_df.astype(str) 
    result_emp_df['social security number'] = ssn

    # Append the result to the result DataFrame
    result_df = pd.concat([result_df, result_emp_df], ignore_index=True)

# # Write the result to a new Excel file
result_df.to_excel(output_file_path, index=False, engine='openpyxl')


# # Get the number of unique values in the 'social security number' column
# num_unique_values = result_df['social security number'].nunique()

# # Print the number of unique values
# print("Number of unique values:", num_unique_values)

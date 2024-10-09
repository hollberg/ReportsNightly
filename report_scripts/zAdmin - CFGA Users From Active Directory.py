from ldap3 import Server, Connection, ALL, SUBTREE
import csv
import os
import pandas as pd
import openpyxl

# Active Directory server details
ad_server = 'DSM-SVR1-CFGA' # or 'DSM-SVR1-CFGA' or '172.16.0.4:3389'
ad_user = 'mitchell.hollberg@cfgreateratlanta.org'
ad_password = os.environ['MITCH_PWD']
search_base = 'OU=CFGA Users,DC=cfga,DC=titan,DC=1sourcing,DC=net'

# LDAP connection
server = Server(ad_server, get_info=ALL)
conn = Connection(server, user=ad_user, password=ad_password, auto_bind=True)

# Search parameters
search_filter = '(objectClass=user)'
search_attributes = ['cn', 'sAMAccountName', 'mail', 'userAccountControl', 'distinguishedName']
conn.search(search_base, search_filter, SUBTREE, attributes=search_attributes)

# File to save user data
# Determine the path to the "OutputReports" directory relative to this script
script_dir = os.path.dirname(os.path.abspath(__file__))  # Gets the directory where the script is located
parent_dir = os.path.dirname(script_dir)  # Move up to the parent directory
output_reports_dir = os.path.join(parent_dir, 'OutputReports')  # Path to OutputReports relative to the script

# Ensure the OutputReports directory exists, create it if it doesn't
if not os.path.exists(output_reports_dir):
    os.makedirs(output_reports_dir)

# Set the output file path within the OutputReports directory
output_file_path = os.path.join(output_reports_dir, 'CFGA_Users_Data.csv')

# Write data to CSV
with open(output_file_path, 'w', newline='', encoding='utf-8') as csvfile:
    csvwriter = csv.writer(csvfile)
    # Write header
    csvwriter.writerow(['Name', 'SAMAccountName', 'EmailAddress', 'Enabled', 'OU'])

    for entry in conn.entries:
        dn = str(entry.distinguishedName)
        # Extract OU from distinguishedName
        ou = dn.split(',', 1)[1] if ',' in dn else ''
        # Check if account is enabled (assuming userAccountControl attribute is present)
        enabled = 'Yes' if (entry.userAccountControl.value & 2) == 0 else 'No'

        csvwriter.writerow([str(entry.cn), str(entry.sAMAccountName), str(entry.mail), enabled, ou])

print(f"Export completed. File saved at {output_file_path}")

# Save as XLSX
sharepoint_path = fr'C:\Users\mitchell.hollberg\Community Foundation for Greater Atlanta\IT - Documents\Public\Reports'
save_file_name = 'zAdmin - Current Staff Active Directory Export.xlsx'

departments_dict = {
    'Philanthropy': 'OU=Adv and Philanthropic',
    'Community': 'OU=Community Impact',
    'Finance': 'OU=Controllers',
    'Office of the President': 'OU=Executive',
    'Admin': 'OU=Executive',
    'Housing': 'OU=Housing Funds',
    'HR': 'OU=HR',
    'IT': 'OU=IT',
    'Marcom': 'OU=Marketing and Comms',
    'Social Impact': 'OU=Social Impact'
}

# Function to find department name from 'long text'
def find_department_name(search_text):
    for name, code in departments_dict.items():
        if code in search_text:
            return name
    return None  # Return None or a default value if no department code is found


df = pd.read_csv(output_file_path)

# Set department column
df['department'] = df['OU'].apply(find_department_name)

df.to_excel(fr'{sharepoint_path}\{save_file_name}',
            index=False, engine='openpyxl')

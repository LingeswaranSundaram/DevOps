import os
import pandas as pd
import json
import win32com.client as win32  # Outlook integration

# Function to find the latest subfolder with the specified prefix
def get_latest_subfolder_with_prefix(main_folder, prefix):
    subfolders = [
        os.path.join(main_folder, d)
        for d in os.listdir(main_folder)
        if os.path.isdir(os.path.join(main_folder, d)) and d.startswith(prefix)
    ]
    if not subfolders:
        raise FileNotFoundError(f"No subfolders found starting with '{prefix}' in the main folder.")
    latest_subfolder = max(subfolders, key=os.path.getmtime)
    return latest_subfolder

# Main folder path
main_folder_path = r'C:\Users\LQI1COB\OneDrive - Bosch Group\Macros\REN Email\Main\01.20.2025'

# Subfolder prefix
subfolder_prefix = "CX_BT_MAIN_"

# Get the latest subfolder with the specified prefix
latest_subfolder = get_latest_subfolder_with_prefix(main_folder_path, subfolder_prefix)

# Construct the full path to the .prf file
file_name = "CX_BT_MAIN.prf"
file_path = os.path.join(latest_subfolder, file_name)

# Read the content of the .prf file
try:
    with open(file_path, 'r') as file:
        data = json.load(file)  # Parse the content as JSON
except json.JSONDecodeError as e:
    print(f"Error reading JSON from file: {e}")
    data = {}
except FileNotFoundError:
    print(f"File '{file_name}' not found in the latest subfolder: {latest_subfolder}")
    exit()

# Extract relevant information (adjust this based on your needs)
project_data = data.get("project", {}).get("children", [])

# Extract only "name" and "origResult" for each child item
filtered_data = [{'name': item.get('name'), 'origResult': item.get('origResult')} for item in project_data]

# Convert the filtered data into a DataFrame for analysis
df = pd.DataFrame(filtered_data)

# Calculate cumulative counts for 'total', 'error', 'success', 'failed', 'none', and 'inconclusive'
counts = {
    'Total': df.shape[0],  # Total count of records
    'Error': df[df['origResult'] == 'ERROR'].shape[0],
    'Success': df[df['origResult'] == 'SUCCESS'].shape[0],
    'Failed': df[df['origResult'] == 'FAILED'].shape[0],
    'None': df[df['origResult'] == 'NONE'].shape[0],
    'Inconclusive': df[df['origResult'] == 'INCONCLUSIVE'].shape[0],
}

# Helper function to determine row background color and text color based on 'origResult'
def get_colors(result):
    color_map = {
        'TOTAL': ('gray', 'black'),
        'ERROR': ('red', 'white'),
        'SUCCESS': ('green', 'black'),
        'FAILED': ('yellow', 'red'),
        'NONE': ('orange', 'black'),
        'INCONCLUSIVE': ('yellow', 'black'),
    }
    return color_map.get(result, ('white', 'black'))

# Convert the DataFrame to an HTML table with updated styles
html_table = f"""
<div style="width: 50%; margin: 0 auto;">
    <table border="1" style="border-collapse: collapse; width: 31%; text-align: center;" class="data-table">
        <tr style="font-size: 18px; font-weight: bold; background-color: lightgray;">
            <th>Name</th>
            <th>OrigResult</th>
        </tr>
        {''.join(
            f"<tr>"
            f"<td>{row['name']}</td>"
            f"<td style='background-color: {get_colors(row['origResult'])[0]}; color: {get_colors(row['origResult'])[1]};'>"
            f"{row['origResult']}</td>"
            f"</tr>" for _, row in df.iterrows())}
    </table>
</div>
"""

# Create a cumulative counts table as HTML with a fixed size
count_table = f"""
<div style="width: 50%; margin: 0 auto;">
    <table border="1" style="border-collapse: collapse; width: 31%; text-align: center;" class="count-table">
        <tr style="font-size: 18px; font-weight: bold; background-color: lightgray;">
            <th>Count Type</th>
            <th>Count</th>
        </tr>
        <tr style="background-color: red; color: white;">
            <td>Error</td>
            <td>{counts['Error']}</td>
        </tr>
        <tr style="background-color: green; color: black;">
            <td>Success</td>
            <td>{counts['Success']}</td>
        </tr>
        <tr style="background-color: yellow; color: red;">
            <td>Failed</td>
            <td>{counts['Failed']}</td>
        </tr>
        <tr style="background-color: yellow; color: black;">
            <td>Inconclusive</td>
            <td>{counts['Inconclusive']}</td>
        </tr>
        <tr style="background-color: orange; color: black;">
            <td>None</td>
            <td>{counts['None']}</td>
        </tr>
        <tr style="background-color: gray; color: black;">
            <td>Total</td>
            <td>{counts['Total']}</td>
        </tr>
    </table>
</div>
"""

# Set up Outlook email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)  # Create a new email

# Configure email properties
mail.Subject = "REN CX_BT_MAIN"

# List of recipient email addresses
recipient_emails = [
    "external.lingeswaran.s@in.bosch.com",
    "external.lingeswaran.s@in.bosch.com",
    "external.lingeswaran.s@in.bosch.com"
]

# Join the list of email addresses into a single string with semicolons
mail.To = ";".join(recipient_emails)

# HTML body with centered icon, formatted counts table, and space between tables
mail.HTMLBody = f"""
<div style="text-align: left; margin: 1px;">
    <p><img src='cid:icon' style='width: 100px; height: 100px;' alt='icon' /></p>
</div>

<p style="text-align: left;">
    Splunk Dashboard: <a href="https://tamer.bosch.com/en-US/app/organisation_metrics/test_result_analyzer?form.jenkins_server=CCAS_Quoth&form.Tok_variant=ECUTEST_Execution&form.Tok_job=%2ARN%2A&form.timerange.earliest=0&form.timerange.latest=">
    Click Here</a>
</p>
<p>Please find below the test case data in table format:</p>
{count_table}
<p>&nbsp;</p> <!-- Add space between tables -->
{html_table}
"""

# Add an icon (you can replace the path to an actual image file)
icon_path = r'C:\Users\LQI1COB\OneDrive - Bosch Group\Macros\REN Email\01.21.2025\icon_2.png'  # Replace with your icon file path
attachment = mail.Attachments.Add(icon_path)
attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "icon")  # Attach the icon

# Send the email
try:
    mail.Send()
    print("Email sent successfully.")
except Exception as e:
    print(f"Error sending email: {e}")
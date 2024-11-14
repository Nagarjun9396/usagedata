import os
import pandas as pd
from glob import glob

# Define the folder path
folder_path = r"Ngage"

# Load all Excel files in the folder and concatenate them
all_data = []
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)
        workbook = pd.ExcelFile(file_path)
        
        for sheet in workbook.sheet_names:
            data = workbook.parse(sheet)
            all_data.append(data)

# Combine all files into a single DataFrame
df = pd.concat(all_data, ignore_index=True)

# Select specific columns and promote headers
df = df[["Employee ID", "Name", "Status", "Role", "Login Date", "Logout Date", "Duration"]]

# Remove header rows within data
df = df[df["Employee ID"] != "Employee ID"]


df["Login Date"] = df["Login Date"].str.replace(" AM","").str.replace(" PM","")
df["Logout Date"] = df["Logout Date"].str.replace(" AM","").str.replace(" PM","")
df["Login Date"] = pd.to_datetime(df["Login Date"], dayfirst=True)
df["Logout Date"] = pd.to_datetime(df["Logout Date"], dayfirst=True)
# Replace 'hr' and 'm' to make it compatible with pd.to_timedelta
df['Duration'] = df['Duration'].str.replace('hr', ' hours').str.replace('m', ' minutes')
df["Duration"] = df["Duration"].apply(lambda x: x + ":00")


# Remove duplicates
df = df.drop_duplicates()


# Extract date from "Login Date"
df["Date"] = df["Login Date"].dt.date

# Filter out rows where Role is "MIS"
df = df[df["Role"] != "MIS"]

# Calculate Duration based on Login and Logout dates
df["Duration"] = (df["Logout Date"] - df["Login Date"]).dt.total_seconds() / 3600  # convert to hours



# Group by Employee ID, Name, and Date
result = df.groupby(["Employee ID", "Name", "Date"]).agg({
    "Login Date": "min",
    "Logout Date": "max",
    "Duration": "sum"
}).reset_index()

# Convert decimal hours to timedelta
result['Duration'] = pd.to_timedelta(result['Duration'], unit='h')

# Format the timedelta as hh:mm:ss
result['Duration'] = result['Duration'].dt.components.apply(
    lambda x: f"{int(x['hours']):02}:{int(x['minutes']):02}:{int(x['seconds']):02}", axis=1
)

#result.to_excel("output/NData.xlsx",index=False)

ndata = result

# Directory where the Excel files are stored
folder_path = r"Tata/"
    
# Function to load and transform each Excel file
def transform_excel_file(file_path):
    # Load the Excel file
    excel_data = pd.ExcelFile(file_path)
    
    # Assuming the first sheet contains the data of interest
    df = excel_data.parse(excel_data.sheet_names[0])
    
    
    
    # Select specific columns
    selected_columns = [
        "Date", "Agent", "Intercom ID", "Group", "Department", "Login Based Calling",
        "Average Calls/Day", "Average C2C Calls/Day - Outbound Answered", "Average Inbound Calls/Day",
        "Call Handling Rate", "Total Calls", "Inbound Calls Offered", "Outbound Click to Call Attempted",
        "Inbound Calls Answered", "Outbound Click to Call Answered", "Available Duration",
        "In-Call Duration", "Break Duration", "Inbound In-Call Duration", "Outbound In-Call Duration",
        "Average Call Handling Duration", "Average Inbound Call Handling Duration", "Average Outbound Call Handling Duration"
    ]
    
    df = df[selected_columns]
    
    # Remove unwanted columns
    columns_to_remove = [
        "Intercom ID", "Group", "Login Based Calling", "Average Calls/Day",
        "Average C2C Calls/Day - Outbound Answered", "Average Inbound Calls/Day", "Department",
        "Call Handling Rate", "Inbound Calls Offered", "Inbound Calls Missed", "Break Duration",
        "Inbound In-Call Duration", "Outbound In-Call Duration", "Average Inbound Call Handling Duration",
        "Average Outbound Call Handling Duration"
    ]
    
    df = df.drop(columns=columns_to_remove, errors='ignore')
    
    # Add a new column 'Total Call Answered'
    df['Total Call Answered'] = df['Inbound Calls Answered'] + df['Outbound Click to Call Answered']
    
    # Drop 'Inbound Calls Answered' and 'Outbound Click to Call Answered'
    df = df.drop(columns=['Inbound Calls Answered', 'Outbound Click to Call Answered'])
        
    # Convert date column to datetime type
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y').dt.date
    
    return df

# Loop through all Excel files in the folder and apply the transformation
combined_df = pd.DataFrame()
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):  # Process only Excel files
        file_path = os.path.join(folder_path, filename)
        transformed_df = transform_excel_file(file_path)
        combined_df = pd.concat([combined_df, transformed_df], ignore_index=True)

# Save the final combined DataFrame to a new Excel file
#combined_df.to_excel(r"output/TData.xlsx", index=False)

tdata = combined_df

# Specify the directory where your Excel files are located
directory = 'Cmpcount'

# Initialize an empty list to store DataFrames
dataframes = []

# Loop through all Excel files in the directory
for file in glob(os.path.join(directory, '*.xlsx')):
    # Load the Excel file
    excel_file = pd.ExcelFile(file)
    
    # Extract the file name without the path
    file_name = os.path.basename(file)
    
    # Loop through each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        # Read the sheet into a DataFrame
        df = excel_file.parse(sheet_name=sheet_name)
        
        # Add a new column with the file name
        df['File Name'] = file_name.replace('.xlsx','')
        df[['Process', 'Date']] = df['File Name'].str.split(' ', expand=True)
        df['Date'] = pd.to_datetime(df["Date"], format="%d-%m-%Y").dt.date
        df.drop(columns='File Name',inplace=True)
        # Append the DataFrame to the list
        dataframes.append(df)

# Concatenate all DataFrames in the list
combined_df = pd.concat(dataframes, ignore_index=True)
combined_df = combined_df[combined_df['Completed'] != 0]
combined_df = combined_df[combined_df['Target'].notna()]
combined_df = combined_df[['Date','Name of Counsellor','User Name','Completed']]
# Display the combined DataFrame
#combined_df.to_excel("output/Cdata.xlsx",index=False)
#print(combined_df)

cdata = combined_df

hrfile= "HRMIS/HRMIS.xlsx"
hdata = pd.read_excel(hrfile)

df = pd.merge(tdata, hdata[["Tata Name", "Ngage Name"]], left_on="Agent", right_on="Tata Name", how='left')
df = pd.merge(df, ndata,left_on=["Ngage Name", "Date"], right_on=["Name", "Date"], how='left')
df = pd.merge(df, hdata, left_on="Agent", right_on="Tata Name", how='left')
df = df[df["Employee ID"].notna()]
df = df.drop(columns=["Ngage Name_y","Tata Name_y", "Yxitusername", "Employee Name", "Tata Name_x",	"Ngage Name_x",	"Employee ID"])
df = pd.merge(df, cdata,left_on=["Date", "Name"], right_on=["Date", "User Name"], how='left')
df = df.drop(columns=["Name of Counsellor",	"User Name"])

# Add date hierarchy columns
df["Date"] = pd.to_datetime(df["Date"])
df['Year'] = df['Date'].dt.year
df['Month'] = df['Date'].dt.month_name()

def week_of_month(date):
    first_day = date.replace(day=1)
    adjusted_day = date.day + first_day.weekday()
    return  f'Week {(adjusted_day - 1) // 7 + 1}'

# Apply the function to create the 'Week of Month' column
df['Week'] = df['Date'].apply(week_of_month)

df['Day'] = df['Date'].dt.day
df["Date"] = df["Date"].dt.date

#filter status = "Active"
df = df[df['Status'] == 'Active ']

#filter status = "Active"
df = df[df['Process'].isin(['Exits', 'CE', 'NHE'])]

# Reording columns
df = df[["Agent", "Date", "Day", "Week", "Month", "Year", "In-Call Duration", "Average Call Handling Duration", "Total Call Answered", "Outbound Click to Call Attempted", "Total Calls", "Completed", "Name", "Login Date", "Logout Date", "Duration", "EMP ID", "Employment Type", "DESIGNATION", "DOJ", "Direct Manager", "Process"]]

df.to_excel('output/Final.xlsx',index=False)
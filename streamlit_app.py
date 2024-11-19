import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard",layout="wide",page_icon= 'Logo.png')

@st.cache_data
def load_data(file_path):
    # Load the Excel file data
    data = pd.read_excel(file_path)
    return data

# Usage
file_path = "output/Final.xlsx"
df = load_data(file_path)


# Sidebar filters
st.image('SLogo.png',width=150,)
st.sidebar.header("Filter")


# Filter by Year with Select All
year_options = sorted(df['Year'].unique())
year_options.insert(0, "Select All")
selected_year = st.sidebar.multiselect("Select Year", year_options, default="Select All")

# Filter data by selected Year(s)
if "Select All" in selected_year:
    filtered_df = df
else:
    filtered_df = df[df['Year'].isin(selected_year)]

month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
month_options = month_order  # Use natural month order
month_options.insert(0, "Select All")
selected_month = st.sidebar.multiselect("Select Month", month_options, default="Select All")

# Filter data by selected Month(s)
if "Select All" not in selected_month:
    filtered_df = filtered_df[filtered_df['Month'].isin(selected_month)]

# Filter by Week with Select All
week_options = sorted(filtered_df['Week'].unique())
week_options.insert(0, "Select All")
selected_week = st.sidebar.multiselect("Select Week", week_options, default="Select All")

# Filter data by selected Week(s)
if "Select All" not in selected_week:
    filtered_df = filtered_df[filtered_df['Week'].isin(selected_week)]

# Filter by Day with Select All
day_options = sorted(filtered_df['Day'].unique())
day_options.insert(0, "Select All")
selected_day = st.sidebar.multiselect("Select Day", day_options, default="Select All")

# Filter data by selected Day(s)
if "Select All" not in selected_day:
    filtered_df = filtered_df[filtered_df['Day'].isin(selected_day)]
    
    
# Filter by Employment Type with Select All
EmploymentType_options = sorted(filtered_df['Employment Type'].unique())
EmploymentType_options.insert(0, "Select All")
selected_ET = st.sidebar.multiselect("Employment Type", EmploymentType_options, default="Select All")

# Filter data by selected Day(s)
if "Select All" not in selected_ET:
    filtered_df = filtered_df[filtered_df['Employment Type'].isin(selected_ET)]

# Filter by Durect Manager with Select All
dm_options = sorted(filtered_df['Direct Manager'].unique())
dm_options.insert(0, "Select All")
selected_dm = st.sidebar.multiselect("Direct Manager", dm_options, default="Select All")

# Filter data by selected Direct Manager(s)
if "Select All" not in selected_dm:
    filtered_df = filtered_df[filtered_df['Direct Manager'].isin(selected_dm)]

# Filter by Process with Select All
ds_options = sorted(filtered_df['DESIGNATION'].unique())
ds_options.insert(0, "Select All")
selected_ds = st.sidebar.multiselect("Designation", ds_options, default="Select All")

# Filter data by selected Day(s)
if "Select All" not in selected_ds:
    filtered_df = filtered_df[filtered_df['DESIGNATION'].isin(selected_ds)]


# Filter by Process with Select All
name_options = sorted(filtered_df['Agent'].unique())
name_options.insert(0, "Select All")
selected_name = st.sidebar.multiselect("Counsellor Name", name_options, default="Select All")

# Filter data by selected Day(s)
if "Select All" not in selected_name:
    filtered_df = filtered_df[filtered_df['Agent'].isin(selected_name)]




# Check if the filtered data is empty
if filtered_df.empty:
    st.warning("No data available for the selected filters.")
else:
    



    # Convert "in Call Duration" to timedelta for easy comparison
    filtered_df["In-Call Duration"] = pd.to_timedelta(filtered_df["In-Call Duration"], unit="D")
    filtered_df["Average Call Handling Duration"] = pd.to_timedelta(filtered_df["Average Call Handling Duration"], unit="D")
    
    total_calls = filtered_df['Total Call Answered'].sum()
    avg_call_duration = filtered_df["In-Call Duration"].mean()
    avg_call_duration_str = str(avg_call_duration).split(".")[0].split(" ")[2]  # Convert to "hh:mm:ss" format
    Avg_ch = filtered_df["Average Call Handling Duration"].mean()
    Avg_ch = str(Avg_ch).split(".")[0].split(" ")[2]  # Convert to "hh:mm:ss" format
    Avg_cm = filtered_df["Completed"].mean()
    d_cmp = filtered_df["Completed"].sum()
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Calls", total_calls ,delta= f"{d_cmp:.0f} | {d_cmp / total_calls * 100:.0f}%",help=f"Green color text Completed Count and Perventage%")
    col2.metric("Avg Call Duration", f"{avg_call_duration_str}")
    col3.metric("Avg Call Handling", f"{Avg_ch}")
    col4.metric("Avg Completed", f"{round(0 if pd.isna(Avg_cm) else Avg_cm)}")
    # Center-aligned metric display


    # Function to apply conditional cell formatting
    def highlight_cells(val, emp_type):
        if emp_type == "Regular" and val > pd.Timedelta("04:00:00"):
            return "background-color: green"
        elif emp_type == "PPI" and val > pd.Timedelta("03:00:00"):
            return "background-color: green"
        else:
            return "background-color: red"

    # Apply conditional formatting to specific cells
    styled_df = filtered_df.style.applymap(
        lambda val: highlight_cells(val, "Regular") if filtered_df["Employment Type"][filtered_df.index[filtered_df['In-Call Duration'] == val][0]] == "Regular" else highlight_cells(val, "PPI"),
        subset=["In-Call Duration"]
    )

    # Display the final filtered result
    with st.expander(label='Filtered Data',expanded=True,icon='🏁'):
        st.dataframe( styled_df)

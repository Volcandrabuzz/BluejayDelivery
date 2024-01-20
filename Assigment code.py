import pandas as pd

def analyze_excel(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Initialize sets to store the results
    employees_consecutive_days = set()
    employees_between_shifts = set()
    employees_more_than_14_hours = set()

    # a) Employees who have worked for 7 consecutive days
    consecutive_days = df.groupby('Employee Name')['Time'].transform(lambda x: x.diff().dt.days.ne(1).cumsum())
    df_consecutive_days = df[consecutive_days >= 7]
    employees_consecutive_days.update(df_consecutive_days['Employee Name'])

    # b) Employees who have less than 10 hours of time between shifts but greater than 1 hour
    df['Time'] = pd.to_datetime(df['Time'])
    time_diff = df.groupby('Employee Name')['Time'].diff().dt.total_seconds() / 3600
    df_time_between_shifts = df[(time_diff > 1) & (time_diff < 10)]
    employees_between_shifts.update(df_time_between_shifts['Employee Name'])

    # c) Employees who have worked for more than 14 hours in a single shift
    df_single_shift = df.dropna(subset=['Timecard Hours (as Time)'])
    
    # No need to convert to timedelta if the format is 'HH:MM'

    df_more_than_14_hours = df_single_shift[df_single_shift['Timecard Hours (as Time)'] > '14:00']
    employees_more_than_14_hours.update(df_more_than_14_hours['Employee Name'])

    # Create a DataFrame for the results
    results_df = pd.DataFrame({
        'Employee Name': list(employees_more_than_14_hours),
        'Position ID': [df.loc[df['Employee Name'] == emp, 'Position ID'].iloc[0] for emp in employees_more_than_14_hours]
    })

    # Print the tabular results in Excel format
    print(results_df.to_markdown(tablefmt="grid", headers="keys", index=False))


# Specify the path to your Excel file
excel_file_path = "Assignment_Timecard.xlsx"

# Call the function with the file path
analyze_excel(excel_file_path)

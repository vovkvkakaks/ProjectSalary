import pandas as pd
import os
import numpy as np

def create_tables():
    month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}

    data_month = {'Month': list(month_names.values()),
                  'Working Hours': [0.00] * 12,
                  'Salary': [0.00] * 12}

    d_month = pd.DataFrame(data_month)

    data_days = {'Day': [], 'Working Hours': [], 'Month': []}

    for month in range(1, 13):
        days_in_month = 31 if month in [1, 3, 5, 7, 8, 10, 12] else 30 if month in [4, 6, 9, 11] else 28

        for day in range(1, days_in_month + 1):
            date = f'{day}.{month}'
            data_days['Day'].append(date)
            data_days['Working Hours'].append(0.00)
            data_days['Month'].append(month_names[month])

    d_days = pd.DataFrame(data_days)

    cur_path = os.path.abspath(os.path.dirname(__file__))
    file_path_days = os.path.join(cur_path, 'day.xlsx')
    file_path_month = os.path.join(cur_path, 'month.xlsx')

    if not os.path.exists(file_path_days):
        d_days.to_excel(file_path_days, index=False)

    if not os.path.exists(file_path_month):
        d_month.to_excel(file_path_month, index=False)

create_tables()

month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}

file_name_days = 'day.xlsx'
file_name_month = 'month.xlsx'
cur_path = os.path.dirname(__file__)

file_path_days = os.path.join(cur_path, file_name_days)
file_path_month = os.path.join(cur_path, file_name_month)

try:
    d_days = pd.read_excel(file_path_days, dtype={'Day': str})
    d_month = pd.read_excel(file_path_month)
except Exception as e:
    print(f"Error occurred while reading Excel files: {e}")
    exit()

d_month['Working Hours'] = d_month['Working Hours'].astype(np.float64)
d_month['Salary'] = d_month['Salary'].astype(np.float64)
d_days['Working Hours'] = d_days['Working Hours'].astype(np.float64)

def format_date(date):
    day, month = map(str, date.split('.'))
    if month != '10': 
        return f'{int(day):d}.{int(month):d}'
    else:
        return f'{int(day):d}.{int(month.zfill(2))}'

def format_time(hour):
    if ':' in hour:
        hours, minutes = map(str, hour.split(':'))
        if minutes == "":
            minutes = '00'
        return f'{int(hours):d}:{int(minutes):02d}'
    elif int(hour) >= 0 and int(hour) <= 23:
        return   f'{int(hour):d}:00'

def valid_date_format(date):
    # Check date format
    parts = date.split('.')
    if len(parts) != 2:
        return False
    day, month = parts
    if not day.isdigit() or not month.isdigit():
        return False
    day = int(day)
    month = int(month)
    if not (1 <= month <= 12):
        return False
    if not (1 <= day <= 31):
        return False
    if (month == 2 and day > 29) or \
       ((month == 4 or month == 6 or month == 9 or month == 11) and day > 30):
        return False
    return True

def valid_time_format(time_str):
    # Check time format
    try:
        hours, minutes = map(int, time_str.split(':'))
        if hours < 0 or hours > 23 or minutes < 0 or minutes > 59:
            return False
        return True
    except ValueError:
        return False

def hour_change_option():
    while True:
        # Input day
        while True:
            day = input('Enter the day (e.g., 1.1, 25.10, etc.): ')
            try:
                day = format_date(day)
                if valid_date_format(day):
                    break
            except:
                print("Invalid date format. Please use the format 'DD.MM', for example '25.10'.")
        
        # Input working hours
        while True:
            h = input('Enter working hours (e.g., 3:15): ')
            try:
                h = format_time(h)
                if valid_time_format(h):
                    break
            except:
                print("Invalid time format. Please use the format 'HH:MM', for example '3:15'.")

        # Update working hours
        update_working_hours(day, h)

        # Continue or exit
        a = input('Continue adding hours? (yes/no): ').lower()
        while a not in ['yes', 'no']:
            a = input("Please enter 'yes' or 'no': ").lower()   
        
        if a == 'no':
            main()
            break

def save_connected_data():
    # Merge "days" and "months" tables on the "Month" column
    merged_df = pd.merge(d_days, d_month, on='Month', how='inner')

    # Group data by month and calculate the sum of hours for each month
    monthly_hours = merged_df.groupby('Month')['Working Hours_x'].sum().reset_index()

    # Create a dictionary where keys are months and values are sums of working hours
    monthly_hours_dict = dict(zip(monthly_hours['Month'], monthly_hours['Working Hours_x']))


    # Update the table with months, adding the sum of working hours for each month
    for index, row in d_month.iterrows():
        month = row['Month']
        d_month.at[index, 'Working Hours'] = monthly_hours_dict.get(month, 0)

    for index, row in d_month.iterrows():
        month = row['Month']
        d_month.at[index, 'Working Hours'] = monthly_hours_dict.get(month, 0)
        # Calculate the salary based on the hours of work
        d_month.at[index, 'Salary'] = round(d_month.at[index, 'Working Hours'] * 27.70, 2)

    try:
        d_month.to_excel(file_path_month, index=False)
    except Exception as e:
        print(f"Error occurred while writing to Excel file: {e}")
        exit()

def update_working_hours(date, hours):
    
    # Check if the date is in the correct format
    if not date or not date.strip() or not isinstance(hours, (int, float, str)):
        print("Invalid date or hours")
        return
        
    # Find the index of the row where the date matches
    index = d_days[d_days['Day'].astype(str) == date].index.tolist()
    # Check if the date exists in the table
    if not index:
        print(f"The date {date} was not found in the table")
        return
    if not d_days.at[index[0], 'Working Hours'] == 0:
        # Ask the user if they want to replace the existing value
        replace = input(f"Working hours for {date} already exist. Replace? (yes/no): ")
        replace = replace.lower()
        while replace not in ['yes', 'no']:
            replace = input("Please write 'yes' or 'no': ").lower()            
            
        if replace == 'no':
            print("Operation canceled.")
            a = input('Continue adding hours? (yes/no): ').lower()
            while a not in ['yes', 'no']:
                a = input("Please write 'yes' or 'no': ").lower()   
            
            # Check whether to continue
            if a == 'no':
                main()
            elif a == 'yes':
                hour_change_option()

    # Update the working hours for the corresponding date
    try:
        d_days.at[index[0], 'Working Hours'] = time_to_float(hours)
    except Exception as e:
        print(f"An error occurred while updating working hours: {e}")
        return
    
    # Save the updated DataFrame to the file 'day.xlsx'
    try:
        d_days.to_excel(file_path_days, index=False)
        print(f"Working hours for {date} successfully updated")
    except Exception as e:
        print(f"An error occurred while writing to the file 'day.xlsx': {e}")

    save_connected_data()    

def time_to_float(time_str):
    try:
        # Split the time string into hours and minutes
        hours, minutes = map(int, time_str.split(':'))
        
        # Calculate the total hours by adding hours and converted minutes to hours
        total_hours = round(hours + minutes / 60, 2)
        
        return total_hours
    except ValueError:
        print("Invalid time format. Please use 'hours:minutes' format, for example: '3:15'")
        return None    

def float_to_time(total_hours):
    try:
        # Calculate the whole hours and remaining minutes
        hours = int(total_hours)
        minutes = int((total_hours - hours) * 60)
        
        # Format the time as a string
        time_str = f"{hours} hours {minutes:02d} minutes"    
        return time_str
    except ValueError:
        print("Invalid input for total_hours. Please provide a valid floating-point number representing hours.")
        return None
    
def print_menu():
    print("Choose an option:")
    print("1. Add working hours")
    print("2. Show monthly information")
    print("3. Exit")

def convert_month_int(month):
    try:
        month = int(month)
        return month
    except:
        month_names = {'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6, 'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12}        
        month = month_names[month.capitalize()]
        month = int(month)
        return month

def option_2():
    while True:
        while True:
            m = input('Enter the month number or name (1-12): ')
            try: 
                m = convert_month_int(m)
                break
            except:
                print('Invalid month number or name. Please enter a number between 1-12 or the full month name.') 
                continue
        if m in range(1, 13):
            month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}
            month_name = month_names[m]
            salary = d_month.loc[d_month['Month'] == month_name, 'Salary'].iloc[0]
            hours_worked = d_month.loc[d_month['Month'] == month_name, 'Working Hours'].iloc[0]
            print(f'Salary for {month_name}: {salary}zł. Total hours worked: {float_to_time(hours_worked)}')

            while True:    
                continue_ = input('Interested in another month? (yes/no): ')
                 
                while continue_ not in ['yes', 'no']:
                    continue_ = input("Please enter 'yes' or 'no': ").lower()
                if continue_ == 'yes':
                    option_2()
                    break
                elif continue_ == 'no':
                    main()
                    break
                else:
                    print('Error. Please try again.')

def main():          
    while True:
        print_menu()
        choice = input("Your choice (1-3): ")

        if choice == '1':
            hour_change_option()
        elif choice == '2':
            option_2()
        elif choice == '3':
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please choose an option from 1 to 3.")
    exit() 

main()


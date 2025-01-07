## Please ensure you install the package pycalendar
## pip install pycalendar

import pandas as pd
from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta

# Load the Excel file
file_path = "Tactical Calendar (1).xlsx"  # Replace with your file path
data = pd.read_excel(file_path)

# Assuming the first column is 'Date' and the second column is 'Text'
data.columns = ['Date', 'Text']  # Adjust column names if necessary

# Initialize the calendar
cal = Calendar()

for index, row in data.iterrows():
    date = pd.to_datetime(row['Date'])
    text = row['Text']
    
    # Create an event
    event = Event()
    event.add('summary', text)
    event.add('dtstart', date)
    event.add('dtend', date + timedelta(days=1))  # Set end time to the next day
    event.add('description', f'Reminder: {text}')
    
    # Add a reminder (alarm) at 9:00 AM on the event day
    alarm = Alarm()
    alarm.add('action', 'DISPLAY')
    alarm.add('trigger', timedelta(hours=-15))  # 9:00 AM trigger
    event.add_component(alarm)
    
    # Add the event to the calendar
    cal.add_component(event)

# Save the iCal file
output_file = "Tactical Calendar 2025.ics"
with open(output_file, 'wb') as f:
    f.write(cal.to_ical())

print(f"iCal file '{output_file}' created successfully!")

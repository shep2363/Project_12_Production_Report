import openpyxl
import os
from xml.etree import ElementTree as ET

def time_to_seconds(timestr):
    """Convert time string formatted as H:M:S to seconds."""
    h, m, s = map(int, timestr.split(':'))
    return h * 3600 + m * 60 + s

def seconds_to_time(seconds):
    """Convert seconds to time string formatted as H:M:S."""
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return f"{h}:{m:02}:{s:02}"

def time_difference(time1, time2):
    """Calculate the difference between two time strings formatted as H:M:S and return the result in seconds."""
    seconds1 = time_to_seconds(time1)
    seconds2 = time_to_seconds(time2)
    difference = seconds2 - seconds1
    if difference < 0:
        difference += 24 * 60 * 60  # Adjust for time that goes over midnight
    return difference

def adjust_column_width(worksheet):
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # Adjust the width
        worksheet.column_dimensions[column].width = adjusted_width

# Get the XML file path from the user
xml_file_path = input("Enter the path to the XML file: ").strip('"')

# Parse the XML from file
tree = ET.parse(xml_file_path)
root = tree.getroot()

# Create a new Excel workbook
wb = openpyxl.Workbook()
ws = None  # Initialize worksheet variable

# Dictionary to store daily idle times
daily_idle_times = {}

# Initial totals
day_total_production_time_in_seconds = 0
day_total_pt_seconds = 0
day_total_pt_trt_seconds = 0
previous_finish_time_in_seconds = None
previous_date = None

for part_report in root.findall('.//PartReports/PartReport'):
    part_name = part_report.find('PartName').text
    creation_datetime = part_report.find('TimeWhenPartWasCreated').text
    current_date = creation_datetime.split('T')[0]

    # Check if date changed and create a new sheet if needed
    if current_date != previous_date and previous_date is not None:
        # Append totals to the previous worksheet
        ws.append(["Totals", "", "", "", seconds_to_time(day_total_production_time_in_seconds),
                   seconds_to_time(daily_idle_times.get(previous_date, 0)),
                   seconds_to_time(day_total_pt_seconds),
                   seconds_to_time(day_total_pt_trt_seconds)])

        adjust_column_width(ws)  # Auto-adjust columns' width

        # Reset totals for the new day
        day_total_production_time_in_seconds = 0
        day_total_pt_seconds = 0
        day_total_pt_trt_seconds = 0
        previous_finish_time_in_seconds = None

    if current_date != previous_date:
        ws = wb.create_sheet(title=current_date)
        ws.append(["Part Name", "Date", "Start Time", "Finish Time", "Total Run Time", "Idle Time", "Production Time", "PT-TRT"])

    start_time = creation_datetime[-8:]
    finish_datetime = part_report.find('TimeWhenPartWasFinished').text
    finish_time = finish_datetime[-8:]

    total_production_time_in_seconds = time_difference(start_time, finish_time)
    day_total_production_time_in_seconds += total_production_time_in_seconds

    production_time_seconds = time_to_seconds(part_report.find('TimeItTookToCreateThePart').text)  # Extracting Production Time
    day_total_pt_seconds += production_time_seconds

    pt_trt_in_seconds = total_production_time_in_seconds - production_time_seconds

    if pt_trt_in_seconds > 0:
        pt_trt = seconds_to_time(pt_trt_in_seconds)
        day_total_pt_trt_seconds += pt_trt_in_seconds
    else:
        pt_trt = "N/A"

    if previous_finish_time_in_seconds is not None:
        idle_time_in_seconds = time_to_seconds(start_time) - previous_finish_time_in_seconds
        if idle_time_in_seconds < 0:
            idle_time_in_seconds += 24 * 60 * 60  # Adjust for idle time that goes over midnight
    else:
        idle_time_in_seconds = 0

    daily_idle_times[current_date] = daily_idle_times.get(current_date, 0) + idle_time_in_seconds

    ws.append([part_name, current_date, start_time, finish_time, seconds_to_time(total_production_time_in_seconds),
               seconds_to_time(idle_time_in_seconds), seconds_to_time(production_time_seconds), pt_trt])

    previous_date = current_date
    previous_finish_time_in_seconds = time_to_seconds(finish_time)

# Don't forget to adjust the last sheet as well
adjust_column_width(ws)

# Append totals to the last worksheet
ws.append(["Totals", "", "", "", seconds_to_time(day_total_production_time_in_seconds),
           seconds_to_time(daily_idle_times.get(previous_date, 0)),
           seconds_to_time(day_total_pt_seconds),
           seconds_to_time(day_total_pt_trt_seconds)])

# Remove the default sheet
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# Save the workbook
output_excel_path = os.path.join(os.path.dirname(xml_file_path), "Production_Report.xlsx")
wb.save(output_excel_path)

print(f"Excel file created successfully at {output_excel_path}!")

import openpyxl
from openpyxl.utils import get_column_letter
from xml.etree import ElementTree as ET
from datetime import datetime, timedelta
import os

# Helper functions
def time_to_seconds(timestr):
    h, m, s = map(int, timestr.split(':'))
    return h * 3600 + m * 60 + s

def seconds_to_decimal_hours(seconds):
    return seconds / 3600

def time_difference(time1, time2):
    seconds1 = time_to_seconds(time1)
    seconds2 = time_to_seconds(time2)
    difference = seconds2 - seconds1
    if difference < 0:
        difference += 24 * 3600  # Adjust for time that goes over midnight
    return difference

def adjust_column_width(worksheet):
    for column in worksheet.columns:
        max_length = 0
        col = [cell for cell in column]
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)
        column_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[column_letter].width = adjusted_width

def parse_time(time_str):
    return datetime.strptime(time_str, '%Y-%m-%dT%H:%M:%S')

def get_shift_time():
    day_shift_start = input("Enter the Day Shift Start Time (HH:MM:SS): ")
    day_shift_end = input("Enter the Day Shift End Time (HH:MM:SS): ")
    night_shift_start = input("Enter the Night Shift Start Time (HH:MM:SS): ")
    night_shift_end = input("Enter the Night Shift End Time (HH:MM:SS): ")
    return day_shift_start, day_shift_end, night_shift_start, night_shift_end

def determine_shift(start_time, day_shift_start, day_shift_end, night_shift_start, night_shift_end):
    start_seconds = time_to_seconds(start_time)
    day_start_seconds = time_to_seconds(day_shift_start)
    day_end_seconds = time_to_seconds(day_shift_end)
    night_start_seconds = time_to_seconds(night_shift_start)
    night_end_seconds = time_to_seconds(night_shift_end)

    if day_start_seconds <= start_seconds < day_end_seconds:
        return 'Day'
    elif night_start_seconds <= start_seconds or start_seconds < night_end_seconds:
        return 'Night'
    else:
        return 'Undefined'

# Main function
def main():
    xml_file_path = input("Enter the path to the XML file: ").strip('"')
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    wb = openpyxl.Workbook()
    master_sheet = wb.create_sheet(title="Master Sheet")
    master_sheet.append(["Part Name", "Date", "Start Time", "Finish Time", "Total Run Time (hours)", "Idle Time (hours)", 
                         "Production Time (hours)", "PT-TRT (hours)", "Shift"])

    daily_idle_times = {}
    day_total_production_time_in_seconds = 0
    day_total_pt_seconds = 0
    day_total_pt_trt_seconds = 0
    previous_finish_time_in_seconds = None
    previous_date = None

    day_shift_start, day_shift_end, night_shift_start, night_shift_end = get_shift_time()

    for part_report in root.findall('.//PartReports/PartReport'):
        part_name = part_report.find('PartName').text
        creation_datetime = part_report.find('TimeWhenPartWasCreated').text
        creation_datetime_obj = parse_time(creation_datetime)

        current_date = creation_datetime_obj.strftime('%Y-%m-%d')

        if current_date != previous_date and previous_date is not None:
            # Add totals to previous worksheet and create a new one for the new date
            ws.append(["Totals", "", "", "",
                       seconds_to_decimal_hours(day_total_production_time_in_seconds),
                       seconds_to_decimal_hours(daily_idle_times.get(previous_date, 0)),
                       seconds_to_decimal_hours(day_total_pt_seconds),
                       seconds_to_decimal_hours(day_total_pt_trt_seconds)])
            adjust_column_width(ws)

            day_total_production_time_in_seconds = 0
            day_total_pt_seconds = 0
            day_total_pt_trt_seconds = 0
            previous_finish_time_in_seconds = None

        if current_date != previous_date:
            ws = wb.create_sheet(title=current_date)
            ws.append(["Part Name", "Date", "Start Time", "Finish Time", "Total Run Time (hours)", "Idle Time (hours)", 
                       "Production Time (hours)", "PT-TRT (hours)", "Shift"])

        start_time = creation_datetime.split('T')[1]
        finish_datetime = part_report.find('TimeWhenPartWasFinished').text
        finish_time = finish_datetime.split('T')[1]

        total_production_time_in_seconds = time_difference(start_time, finish_time)
        day_total_production_time_in_seconds += total_production_time_in_seconds

        production_time = part_report.find('TimeItTookToCreateThePart').text
        production_time_seconds = time_to_seconds(production_time)
        day_total_pt_seconds += production_time_seconds

        pt_trt_in_seconds = total_production_time_in_seconds - production_time_seconds
        day_total_pt_trt_seconds += pt_trt_in_seconds

        if previous_finish_time_in_seconds is not None:
            idle_time_in_seconds = time_difference(finish_time, previous_finish_time)
        else:
            idle_time_in_seconds = 0

        daily_idle_times[current_date] = daily_idle_times.get(current_date, 0) + idle_time_in_seconds

        shift = determine_shift(start_time, day_shift_start, day_shift_end, night_shift_start, night_shift_end)

        row_entry = [part_name, current_date, start_time, finish_time, 
                     seconds_to_decimal_hours(total_production_time_in_seconds),
                     seconds_to_decimal_hours(idle_time_in_seconds), 
                     seconds_to_decimal_hours(production_time_seconds), 
                     seconds_to_decimal_hours(pt_trt_in_seconds), shift]

        master_sheet.append(row_entry)
        ws.append(row_entry)

        previous_date = current_date
        previous_finish_time = finish_time
        previous_finish_time_in_seconds = time_to_seconds(finish_time)

    adjust_column_width(master_sheet)

    ws.append(["Totals", "", "", "",
               seconds_to_decimal_hours(day_total_production_time_in_seconds),
               seconds_to_decimal_hours(daily_idle_times.get(previous_date, 0)),
               seconds_to_decimal_hours(day_total_pt_seconds),
               seconds_to_decimal_hours(day_total_pt_trt_seconds)])

    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    output_excel_path = os.path.join(os.path.dirname(xml_file_path), "Production_Report.xlsx")
    wb.save(output_excel_path)
    print(f"Excel file created successfully at {output_excel_path}!")

# Entry point
if __name__ == "__main__":
    main()

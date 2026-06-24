import win32com.client
import time
import json
import os
import re
import math
import random
import xlwings as xw
import shutil
import otdrparser
import subprocess


class InvalidFormat(ValueError):
    """Custom exception for invalid date format."""
    pass

fiber_lengths = []
wb = None
ws = None
terminal_width = shutil.get_terminal_size().columns

def extract_json_attenuations(file_path):
    with open(file_path, 'r') as file:
        data = json.load(file)

    attenuations = {}
    
    try:
        otdr_measurements = data['Measurement'].get('OtdrMeasurements', [])
        if len(otdr_measurements) < 2:
            print(f"Error: 'OtdrMeasurements' list has less than 2 entries in file {file_path}")
            return attenuations

        events_1 = otdr_measurements[0].get('Events', [])
        events_2 = otdr_measurements[1].get('Events', [])
        if len(events_1) < 3 or len(events_2) < 3:
            raise InvalidFormat(f"'Events' list has less than 3 entries in file {file_path}")

        a = events_1[2]['PreviousFiberSection']['Attenuation']
        b = events_2[2]['PreviousFiberSection']['Attenuation']

        b_position_death = float(events_2[1]['Position'])
        b_position_link = float(events_2[2]['Position'])

        fiber_length = float(math.floor(b_position_link - b_position_death))

        fiber_lengths.append(fiber_length)
        attenuations['fl'] = fiber_length

        a = float(a)
        b = float(b)
        
        if a >= 0.3 and (b >= 0.1 and b < 0.3):
            attenuations['1310nm'] = a
            attenuations['1550nm'] = b
        elif b >= 0.3 and (a >= 0.1 and a < 0.3):
            attenuations['1310nm'] = b
            attenuations['1550nm'] = a
        
        return attenuations

    except (KeyError, IndexError) as e:
        raise InvalidFormat(f"Missing expected data in file {file_path}: {e}")
    except ValueError:
        raise InvalidFormat(f"'a' or 'b' could not be converted to a float in file {file_path}. Invalid Trace")


def to_camel_case(text):
    words = text.split()
    return words[0].lower() + ' '.join(word.capitalize() if word.isalpha() else word for word in words[1:])


def format_date(input_date):
    pattern = r"(\d{2})(st|nd|rd|th)\s([a-zA-Z]+),\s(\d{4})"
    
    match = re.match(pattern, input_date)
    if match:
        day, suffix, month, year = match.groups()
        formatted_date = f"{day}{suffix} {month.capitalize()}, {year}"
        return formatted_date
    else:
        raise InvalidFormat("Invalid date format. Expected format: '09th october, 2024'.")

def convert_to_serializable(obj):
    if isinstance(obj, bytes):
        return obj.decode(errors='replace')
    elif isinstance(obj, dict):
        return {k: convert_to_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_serializable(i) for i in obj]
    return obj

def get_events(key, trace):
    with open(trace, 'rb') as fp:
        blocks = otdrparser.parse(fp)
        serializable_blocks = convert_to_serializable(blocks)
        events = [block for block in serializable_blocks if 'KeyEvents' in block.get('name', '')]
        events = events[0].get('events', [])
        if len(events) > 2:
            return [events[1], events[-1]]
        else:
            raise Exception(f'Invalid trace {key}')
def extract_sor_attenuations(key, trace=None):
    if trace is None or not trace:
        raise Exception('Invalid or empty trace')

    attenuations = {'1310nm': None, '1550nm': None, 'fl': None}

    for wavelength, filename in trace.items():
        file_path = os.path.join(os.getcwd(), filename)

        try:
            events = get_events(key, filename)
            if events:
                death = events[0]
                fiber = events[-1]

                if wavelength == '1310':
                    attenuations['1310nm'] = fiber.get('slope', None)
                elif wavelength == '1550':
                    attenuations['1550nm'] = fiber.get('slope', None)

                    if 'distance_of_travel' in fiber and 'distance_of_travel' in death:
                        attenuations['fl'] = math.floor(fiber['distance_of_travel']) - math.floor(death['distance_of_travel'])
                        fiber_lengths.append(attenuations['fl'])

        except Exception as e:
            print(f"Error retrieving or processing file {file_path}: {e}")
            continue

    if None in attenuations.values():
        missing_keys = [key for key, value in attenuations.items() if value is None]
        print("Incomplete attenuation data:", attenuations)
        raise Exception(f"Missing attenuation data for: {', '.join(missing_keys)}")

    insert_att(ws, key, attenuations)

def process_traces(target_files):
    try:
        traces = {}

        fiber_id = None
        for filename in target_files:
            if filename.endswith('.sor'):
                match = match = re.match(r'(fiber\d+)_', filename, re.IGNORECASE)
                if match:
                    fiber_id = match.group(1)
                    if fiber_id not in traces:
                        traces[fiber_id] = {'1310': None, '1550': None}
                    if '1310' in filename:
                        traces[fiber_id]['1310'] = filename
                    elif '1550' in filename:
                        traces[fiber_id]['1550'] = filename
        
        if len(traces) < 1:
            raise Exception('No traces found.')
        
        print('-' * terminal_width)
        print(f'Found {len(traces)} trace(s)')

        for key, trace in traces.items():
            if not isinstance(trace, dict):
                raise TypeError(f"Unexpected data type for trace {key}: Expected dict, got {type(trace)}")

            extract_sor_attenuations(key, trace)


                
    except Exception as e:
        raise Exception(f"Error in group_traces: {e}")

def report_info():

    tested_by = input("Enter the name of the person who tested: ").strip()

    if tested_by:
        if len(tested_by) == 1:
            name_mapping = {
                'a': 'Amon Kibet',
                'j': 'Josedward Mukhwana',
                'v': 'Victor Kipchumba',
                's': "Sharon Chepng'eno",
                'l': 'Linet Chepkoech',
                'e': 'Emmanuel Kipkoech',
                'd': 'Denis Korir'
            }
            if tested_by.lower() in name_mapping:
                tested_by = name_mapping[tested_by.lower()]
        
        tested_by = tested_by.upper()
    else:
        tested_by = False
    
    shift = input("Enter Shift (D for Day, N for Night): ").strip()
    if shift:
        if len(shift) == 1:
            name_mapping = {
                'd': 'DAY',
                'n': 'NIGHT',
            }
            if shift.lower() in name_mapping:
                shift = name_mapping[shift.lower()]
            else:
                raise InvalidFormat("Unable to identify tester.")
        
        shift = shift.upper()
    else:
        shift = False


    test_date = input("Enter the test date (e.g., 01st January, 2024): ").strip()
    if test_date:
        test_date = format_date(test_date)
    else:
        test_date = False
    
    c_type = input("Enter the cable type: ").strip()
    if not c_type:
        c_type = False
    
    l_type = input("Enter the Type of Length: ").strip()
    if not l_type:
        l_type = False
    
    t_side = input("Enter the Test Side: ").strip()
    if not l_type:
        t_side = False
    
    cable_length = input("Enter the cable length (in meters): ").strip()
    if not cable_length:
        raise InvalidFormat("Cable length is required.")
    
    dia = input("Enter the cable diameter as specified in specification: ").strip()
    if not dia:
        raise InvalidFormat("Cable length is required.")
    
    frp = input("Enter the FRP diameter as specified in specification: ").strip()
    if not frp:
        raise InvalidFormat("FRP diameter is required.")
    
    otdr_no = input("Enter the OTDR NO.: ").strip()
    if not otdr_no:
        otdr_no = False
    
    status = input("Enter the cable status (leave blank if ok): ").strip()
    if not status:
        status = "OK"

    workbook_path = input("Enter the report path: ").strip()
    if not workbook_path:
        raise InvalidFormat("Report path is required.")
    
    return tested_by, shift, test_date, c_type, l_type, t_side, float(cable_length), float(dia), float(frp), int(otdr_no), status, workbook_path

def new_record(*args):
    tested_by, shift, test_date, c_type, l_type, t_side, cable_length, dia, frp, otdr_no, status, workbook_path  = args

    app = xw.App(visible=False)
    wb = app.books.open(workbook_path)

    if folder_name in [sheet.name for sheet in wb.sheets]:
        ws = wb.sheets[folder_name]
    else:
        last_sheet = wb.sheets[-1]
        new_sheet = last_sheet.copy(name=folder_name)

        ws = new_sheet

    if shift:
        ws.range('A9:B9').value = shift.upper()

    ws.range('C9:D9').value = folder_name.upper()

    ws.range('E9:F9').value = float(cable_length)

    if c_type:
        ws.range('K9:M11').value = c_type.upper()
    if l_type:
        ws.range('O8:P9').value = l_type.upper()
    if t_side:
        ws.range('O10: p11').value = t_side.upper()
    if otdr_no:
        ws.range('P7').value = otdr_no
    if tested_by:
        ws.range('C69:F70').value = tested_by.upper()
    if test_date:
        ws.range('I10:J11').value = test_date
    if status.lower():
        ws.range('F67:P68').value = status.upper()

    ws.range('B64').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100
    ws.range('C64').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100
    ws.range('D64').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100
    ws.range('E64').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100

    ws.range('B65').value = math.floor(random.uniform((frp - 0.01), (frp + 0.09)) * 100) / 100
    ws.range('C65').value = math.floor(random.uniform((frp - 0.01), (frp + 0.09)) * 100) / 100
    ws.range('D65').value = math.floor(random.uniform((frp - 0.01), (frp + 0.09)) * 100) / 100
    ws.range('E65').value = math.floor(random.uniform((frp - 0.01), (frp + 0.09)) * 100) / 100

    return wb, ws

def update_cable_details(ws, fiber_no):
    ws.range('E10:F11').value = float(min(fiber_lengths))
    ws.range('G9:H9').value = math.floor(fiber_no/12)
    ws.range('I9:J9').value = float(fiber_no)   

def extract_fiber_number(fiber_name):
    match = re.search(r'(\d+)', fiber_name)
    if match:
        return int(match.group(1))
    return None

def insert_att(ws, fiber_name, attenuations):
    fiber_index = extract_fiber_number(fiber_name)
    if fiber_index is None or fiber_index < 1 or fiber_index > 144:
        print(f"Invalid fiber name or out of range: {fiber_index}")
        return

    if 1 <= fiber_index <= 48:
        row = 14 + (fiber_index - 1)
        col_1310nm, col_1550nm, col_fl = "C", "D", "E"
    elif 49 <= fiber_index <= 96:
        row = 14 + (fiber_index - 49)
        col_1310nm, col_1550nm, col_fl = "H", "I", "J"
    elif 97 <= fiber_index <= 144:
        row = 14 + (fiber_index - 97)
        col_1310nm, col_1550nm, col_fl, col_remark = "M", "N", "O", "P"
    else:
        print(f"Unexpected fiber index {fiber_index} for fiber {fiber_name}")
        return

    try:
        ws.range(f"{col_1310nm}{row}").value = attenuations.get('1310nm', 0)
        ws.range(f"{col_1550nm}{row}").value = attenuations.get('1550nm', 0)
        ws.range(f"{col_fl}{row}").value = attenuations.get('fl', 0)

        if 97 <= fiber_index <= 144:
            ws.range(f"{col_remark}{row}").value = "OK"

    except Exception as e:
        raise InvalidFormat(f"Error inserting values for Fiber {fiber_index}: {e}")

def close_all_excel_apps():
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        raise Exception("Failed to close Excel processes.")
        
def pause_and_exit():
    if wb != None:
        wb.save()
        wb.close()

    close_all_excel_apps()
    input("Press any Enter to exit . . . ")
    exit()

if __name__ == "__main__":
    try:
        current_directory = os.getcwd()
        folder_name = os.path.basename(current_directory)
        traces = f'{current_directory}/json'

        tested_by, shift, test_date, c_type, l_type, t_side, cable_length, dia, frp, otdr_no, status, workbook_path = report_info()
        args = tested_by, shift, test_date, c_type, l_type, t_side, cable_length, dia, frp, otdr_no, status, workbook_path

        wb, ws = new_record(*args)

        if os.path.isdir(traces):
            json_files = [f for f in os.listdir(traces) if f.endswith('.json')]
            if json_files:
                print('-' * terminal_width)
                print(f"Found {len(json_files)} trace files in the directory.")
                for json_file in json_files:
                    file_path = os.path.join(traces, json_file)
                    attenuations = extract_json_attenuations(file_path)
                    if all(key in attenuations for key in ['1310nm', '1550nm', 'fl']):
                        insert_att(ws, key, attenuations)
                    else:
                        print(f"Incomplete attenuation data: {os.path.splitext(json_file)[0]}", attenuations)
                        raise Exception(f"{key} has incomplete attenuation data")

                    if attenuations:
                        fiber_name = os.path.splitext(json_file)[0]
                        insert_att(ws, fiber_name, attenuations)
                
                update_cable_details(ws, len(json_files))
                
        else:
            sor_files = [f for f in os.listdir(current_directory) if f.endswith('.sor')]
            if sor_files:
                process_traces(sor_files)
                update_cable_details(ws, len(sor_files) / 2)
            else:
                print('-' * terminal_width)
                print("No JSON files found in the current directory.")
                print('-' * terminal_width)
        
        wb.save()
        wb.close()
        print('-' * terminal_width)
        print(f"Attenuations saved to {workbook_path} for folder '{folder_name}'.")
        print('-' * terminal_width)
        print('Report successfully generated.')
        print('-' * terminal_width)
        
        close_all_excel_apps()
        print('Exiting in 2 seconds . . .')
        time.sleep(2)
        exit()

    except Exception as e:
        print(f"An error occurred: {e}")
        pause_and_exit()

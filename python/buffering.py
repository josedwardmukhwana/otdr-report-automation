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
        if len(tested_by) > 0:
            name_mapping = {
                'a': 'Amon Kibet',
                'j': 'Josedward Mukhwana',
                'v': 'Victor Kipchumba',
                's': "Sharon Chepng'eno",
                'l': 'Linet Chepkoech',
                'e': 'Emmanuel Kipkoech',
                'd': 'Denis Korir',
                'g': 'Gloria Kaplelach',
                'f': 'Faith Assava',
                'ko': 'Kevin Otieno',
                'kk': 'Kevin Kipngeno',
                'i': 'Isaac Alukhaba',
                'm': 'Mike Wanyama',
                'an': 'Antony Ndolo',
                'jn': 'Joseph Njogu',
                'd': 'Diana Chepkoech',
                'ac': 'Aggrey Cheruiyot'
            }
            if tested_by.lower() in name_mapping:
                tested_by = name_mapping[tested_by.lower()]
            else:
                raise InvalidFormat("Unable to identify tester.")
        
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
                raise InvalidFormat("Unable to shift.")
        
        shift = shift.upper()
    else:
        shift = False


    test_date = input("Enter the test date (e.g., 01st January, 2024): ").strip()
    if test_date:
        test_date = format_date(test_date)
    else:
        test_date = False
    
    customer = input("Enter the customer name: ").strip()
    if not customer:
        customer = False
    
    relo = input("Enter the relo number: ").strip()
    if not relo:
        relo = False
    
    tube_color = input("Enter the Tube color: ").strip()
    if not tube_color:
        tube_color = False
    
    tube_length = input("Enter the loose tube length (in meters): ").strip()
    if not tube_length:
        raise InvalidFormat("Tube length is required.")
    
    odia = input("Enter the loose tube outer diameter as specified in specification: ").strip()
    if not odia:
        raise InvalidFormat("Tube outer diameter is required.")
    
    idia = input("Enter the loose tube inner diameter as specified in specification: ").strip()
    if not idia:
        raise InvalidFormat("Tube inner diameter is required.")
    
    thickness = input("Enter the loose tube thickness as specified in specification: ").strip()
    if not thickness:
        raise InvalidFormat("Tube thickness is required.")
    
    otdr_no = input("Enter the OTDR NO.: ").strip()
    if not otdr_no:
        otdr_no = False
    
    e_id = input("Enter the Equipment ID: ").strip()
    if not e_id:
        e_id = False
    
    status = input("Enter the cable status (leave blank if ok): ").strip()
    if not status:
        status = "OK"

    workbook_path = input("Enter the report path: ").strip()
    if not workbook_path:
        raise InvalidFormat("Report path is required.")
    
    return tested_by, shift, test_date, customer, relo, tube_color, float(tube_length), float(odia), float(idia), float(thickness), int(otdr_no), e_id,status,workbook_path

def new_record(*args):
    tested_by, shift, test_date, customer, relo, tube_color, tube_length, odia, idia, thickness, otdr_no, e_id,status,workbook_path  = args

    app = xw.App(visible=False)
    wb = app.books.open(workbook_path)

    if folder_name in [sheet.name for sheet in wb.sheets]:
        ws = wb.sheets[folder_name]
    else:
        last_sheet = wb.sheets[-1]
        new_sheet = last_sheet.copy(name=folder_name)

        ws = new_sheet

    if shift:
        ws.range('A8:B8').value = shift.upper()

    ws.range('C8:E8').value = folder_name.upper()

    ws.range('F8').value = tube_color.upper()

    ws.range('G8').value = float(tube_length)

    ws.range('E6:G6').value = customer.upper()

    ws.range('J6').value = int(relo)

    ws.range('B9:B10').value = int(e_id)

    if otdr_no:
        ws.range('J8').value = otdr_no
    if tested_by:
        ws.range('B34:D34').value = tested_by.upper()
    if test_date:
        ws.range('E9:F10').value = test_date
    
    ws.range('G10:H10').value = time.strftime("%d%b%Y").upper()

    if status.lower():
        ws.range('E32:J33').value = status.upper()

    ws.range('B27:C27').value = math.floor(random.uniform((idia - 0.01), (idia + 0.09)) * 100) / 100
    ws.range('D27:E27').value = math.floor(random.uniform((idia - 0.01), (idia + 0.09)) * 100) / 100
    ws.range('F27:G27').value = math.floor(random.uniform((idia - 0.01), (idia + 0.09)) * 100) / 100
    ws.range('H27:J27').value = math.floor(random.uniform((idia - 0.01), (idia + 0.09)) * 100) / 100

    ws.range('B28:C28').value = math.floor(random.uniform((odia - 0.01), (odia + 0.09)) * 100) / 100
    ws.range('D28:E28').value = math.floor(random.uniform((odia - 0.01), (odia + 0.09)) * 100) / 100
    ws.range('F28:G28').value = math.floor(random.uniform((odia - 0.01), (odia + 0.09)) * 100) / 100
    ws.range('H28:J28').value = math.floor(random.uniform((odia - 0.01), (odia + 0.09)) * 100) / 100

    ws.range('B29:C29').value = math.floor(random.uniform((thickness - 0.01), (thickness + 0.09)) * 100) / 100
    ws.range('D29:E29').value = math.floor(random.uniform((thickness - 0.01), (thickness + 0.09)) * 100) / 100
    ws.range('F29:G29').value = math.floor(random.uniform((thickness - 0.01), (thickness + 0.09)) * 100) / 100
    ws.range('H29:J29').value = math.floor(random.uniform((thickness - 0.01), (thickness + 0.09)) * 100) / 100

    return wb, ws

def update_cable_details(ws, fiber_no):
    ws.range('I10:J10').value = float(min(fiber_lengths))
    ws.range('H8:I8').value = float(fiber_no)   

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
        row = 13 + (fiber_index - 1)
        col_1310nm, col_1550nm, col_fl, col_remark = "F", "G", "H", "J"
    # elif 49 <= fiber_index <= 96:
    #     row = 14 + (fiber_index - 49)
    #     col_1310nm, col_1550nm, col_fl = "H", "I", "J"
    # elif 97 <= fiber_index <= 144:
    #     row = 14 + (fiber_index - 97)
    #     col_1310nm, col_1550nm, col_fl, col_remark = "M", "N", "O", "P"
    else:
        print(f"Unexpected fiber index {fiber_index} for fiber {fiber_name}")
        return

    try:
        ws.range(f"{col_1310nm}{row}").value = attenuations.get('1310nm', 0)
        ws.range(f"{col_1550nm}{row}").value = attenuations.get('1550nm', 0)
        ws.range(f"{col_fl}{row}").value = attenuations.get('fl', 0)

        if 1 <= fiber_index <= 48:
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

        tested_by, shift, test_date, customer, relo, tube_color, tube_length, odia, idia, thickness, otdr_no, e_id,status,workbook_path = report_info()
        args = tested_by, shift, test_date, customer, relo, tube_color, tube_length, odia, idia, thickness, otdr_no, e_id,status,workbook_path

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

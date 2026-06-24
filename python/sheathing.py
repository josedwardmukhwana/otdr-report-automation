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

    attenuations = {}

    for wavelength, filename in trace.items():
        file_path = f"{os.getcwd()}/{filename}"

        try:
            events = get_events(key, filename)
            if events:
                death = events[0]
                fiber = events[-1]
                if fiber['slope'] >= 0.29:
                    attenuations['1310nm'] = fiber['slope'] if wavelength == '1310' else None
                else:
                    attenuations['1550nm'] = fiber['slope'] if wavelength == '1550' else None
                    attenuations['fl'] = math.floor(fiber['distance_of_travel']) - math.floor(death['distance_of_travel'])

                    fiber_lengths.append(attenuations['fl'])


        except Exception as e:
            print(f"Error retrieving or processing file {file_path}: {e}")
            continue
    
    if all(key in attenuations for key in ['1310nm', '1550nm', 'fl']):
        insert_att(ws, key, attenuations)
    else:
        print("Incomplete attenuation data:", attenuations)
        raise Exception(f"{key} has incomplete attenuation data")

    
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
                'dc': 'Diana Chepkoech',
                'ac': 'Aggrey Cheruiyot'
            }
            if tested_by.lower() in name_mapping:
                tested_by = name_mapping[tested_by.lower()]
        
        tested_by = tested_by.upper()
    else:
        tested_by = False


    test_date = input("Enter the test date (e.g., 01st January, 2024): ").strip()
    if test_date:
        test_date = format_date(test_date)
    else:
        test_date = False
    
    c_type = input("Enter the cable type: ").strip()
    if not c_type:
        C_type = False
    
    cable_length = input("Enter the cable length (in meters): ").strip()
    if not cable_length:
        raise InvalidFormat("Cable length is required.")
    
    status = input("Enter the cable status (leave blank if ok): ").strip()
    if not status:
        status = "OK"
    
    dia = input("Enter the cable diameter as specified in specification: ").strip()
    if not dia:
        raise InvalidFormat("Cable length is required.")
    
    thick = input("Enter the cable thickness as specified in specification: ").strip()
    if not thick:
        raise InvalidFormat("Cable thickness is required.")

    yarn = input("Enter the type of yarn used: ").strip()
    no_yarn = False
    if not yarn:
        yarn = False
    else:
        no_yarn = input("Enter the number of yarns used: ").strip()
        if not no_yarn:
            no_yarn = False
    
    no_rip = input("Enter the number of ripcords used: ").strip()
    if not no_rip:
        no_rip = False
    
    otdr_no = input("Enter the OTDR NO.: ").strip()
    if not otdr_no:
        otdr_no = False
    
    sz_spool = input("Enter the SZ Spool No.: ").strip()
    if not sz_spool:
        sz_spool = False

    workbook_path = input("Enter the report path: ").strip()
    if not workbook_path:
        raise InvalidFormat("Report path is required.")



    return sz_spool, otdr_no, c_type.upper(), yarn, float(no_yarn), float(no_rip), float(dia), float(thick), status, tested_by, test_date, float(cable_length), workbook_path

def new_record(*args):
    sz_spool, otdr_no, c_type, yarn, no_yarn, no_rip, dia, thick, status, tested_by, test_date, cable_length, workbook_path, folder_name = args

    app = xw.App(visible=False)
    wb = app.books.open(workbook_path)

    if folder_name in [sheet.name for sheet in wb.sheets]:
        ws = wb.sheets[folder_name]
    else:
        last_sheet = wb.sheets[-1]
        new_sheet = last_sheet.copy(name=folder_name)

        ws = new_sheet

    ws.range('D7:F7').value = folder_name

    ws.range('E8:F8').value = float(cable_length)

    if c_type:
        ws.range('N6:P6').value = c_type.upper()
    if otdr_no:
        ws.range('P7').value = otdr_no
    if sz_spool:
        ws.range('J7:M7').value = sz_spool
    if tested_by:
        ws.range('D64:G64').value = tested_by
    if test_date:
        ws.range('O5:P5').value = test_date
    if status.lower():
        ws.range('O64:P64').value = status.upper()
    if yarn:
        ws.range('L61:N61').value = yarn.upper()
        if no_yarn:
            ws.range('P61').value = float(no_yarn)
    if no_rip:
            ws.range('P62').value = float(no_rip)

    ws.range('D61').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100
    ws.range('E61').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100
    ws.range('F61').value = math.floor(random.uniform((dia - 0.01), (dia + 0.09)) * 100) / 100

    ws.range('D62').value = math.floor(random.uniform((thick - 0.01), (thick + 0.09)) * 100) / 100
    ws.range('E62').value = math.floor(random.uniform((thick - 0.01), (thick + 0.09)) * 100) / 100
    ws.range('F62').value = math.floor(random.uniform((thick - 0.01), (thick + 0.09)) * 100) / 100

    return wb, ws

def update_cable_details(ws, fiber_no):
    ws.range('I8:J8').value = float(min(fiber_lengths))
    ws.range('M8').value = math.floor(fiber_no/12)
    ws.range('P8').value = float(fiber_no)   

def extract_fiber_number(fiber_name):
    match = re.search(r'(\d+)', fiber_name)
    if match:
        return int(match.group(1))
    return None

def insert_att(ws, fiber_name, attenuations):
    fiber_index = extract_fiber_number(fiber_name)
    if fiber_index is None or fiber_index < 1 or fiber_index > 144:
        print(f"Invalid fiber name or out of range: {fiber_name}")
        return

    if 1 <= fiber_index <= 48:
        row = 11 + (fiber_index - 1)
        col_1310nm, col_1550nm, col_fl = 4, 5, 6
    elif 49 <= fiber_index <= 96:
        row = 11 + (fiber_index - 49)
        col_1310nm, col_1550nm, col_fl = 9, 10, 11
    elif 97 <= fiber_index <= 144:
        row = 11 + (fiber_index - 97)
        col_1310nm, col_1550nm, col_fl = 14, 15, 16
    else:
        print(f"Unexpected fiber index {fiber_index} for fiber {fiber_name}")
        return

    try:
        ws.cells(row, col_1310nm).value = attenuations.get('1310nm', 0)
        ws.cells(row, col_1550nm).value = attenuations.get('1550nm', 0)
        ws.cells(row, col_fl).value = attenuations.get('fl', 0)
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

        sz_spool, otdr_no, c_type, yarn, no_yarn, no_rip, dia, thick, status, tested_by, test_date, cable_length, workbook_path = report_info()
        args = sz_spool, otdr_no, c_type, yarn, no_yarn, no_rip, dia, thick, status, tested_by, test_date, cable_length, workbook_path, folder_name

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

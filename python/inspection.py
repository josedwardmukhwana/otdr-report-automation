import os
import shutil
import xlwings as xw
import subprocess

terminal_width = shutil.get_terminal_size().columns

def find_packing_file(folder_path):
    for file in os.listdir(folder_path):
        if "packing" in file.lower() and file.endswith(".xlsm"):
            return file
    return None

def create_inspection_report(folder_path, old_file_name):
    new_file_name = old_file_name.replace("PACKING", "INSPECTION REPORT").replace("packing", "INSPECTION REPORT")
    new_file_name = " ".join(new_file_name.split(" ")[1:])
    new_file_path = os.path.join(folder_path, new_file_name)
    old_file_path = os.path.join(folder_path, old_file_name)
    shutil.copy2(old_file_path, new_file_path)
    print('-' * terminal_width)
    print(f"Inspection Report workbook generated: {new_file_name}")
    print('-' * terminal_width)
    return new_file_path

def format_worksheet(ws):
    ws.api.Rows(71).Insert()
    ws.api.Rows(71).RowHeight = 22.5
    for _ in range(3):
        ws.api.Rows(72).Insert()
        ws.api.Rows(72).RowHeight = 39.75
    
    merge_ranges = [
        ("A71", "C71"), ("D71", "G71"), ("H71", "J71"), ("K71", "L71"), ("M71", "N71"),
        ("A72", "C72"), ("D72", "G72"), ("H72", "J72"), ("K72", "L72"), ("M72", "N72"),
        ("A73", "C73"), ("D73", "G73"), ("H73", "J73"), ("K73", "L73"), ("M73", "N73"),
        ("A74", "C74"), ("D74", "G74"), ("H74", "J74"), ("K74", "L74"), ("M74", "N74")
    ]
    
    for start, end in merge_ranges:
        merged_range = ws.range(f"{start}:{end}")
        merged_range.merge()
        merged_range.api.Borders.Weight = 2
        merged_range.api.Font.Bold = True

    for row in range(71, 75):
        ws.range(f"A{row}:N{row}").api.Borders.Weight = 2

    ws.range("D71").value = "NAME"
    ws.range("H71").value = "DESIGNATION"
    ws.range("K71").value = "DATE"
    ws.range("M71").value = "SIGNATURE"
    
    ws.range("A72").value = "Tested by:"
    ws.range("D72").value = ws.range("D75").value
    ws.range("H72").value = "Quality Control"
    ws.range("K72").value = ws.range("H76").value
    
    ws.range("A73").value = "Verified by:"
    ws.range("A74").value = "Inspected by:"
    
    ws.api.Rows(75).Delete()
    ws.api.Rows(75).Delete()

def close_all_excel_apps():
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        raise Exception("Failed to close Excel processes.")

def pause_and_exit():
    input("Press Enter to exit . . . ")
    exit()

if __name__ == "__main__":
    try:
        folder_path = "."
        old_file_name = find_packing_file(folder_path)
        if old_file_name:
            new_file_path = create_inspection_report(folder_path, old_file_name)
            
            app = xw.App(visible=False)
            wb = app.books.open(new_file_path)

            for ws in wb.sheets:
                format_worksheet(ws)
                print(f"Generated Inspection report for {ws.name}")

            wb.save()
            wb.close()
            app.quit()

            print('-' * terminal_width)
            print(f"Successfully generated inspection reports for {old_file_name}")
        else:
            print('-' * terminal_width)
            print("No packing file found.")

        print('-' * terminal_width)
        close_all_excel_apps()
        pause_and_exit()

    except Exception as e:
        print(f"An error occurred: {e}")
        pause_and_exit()
